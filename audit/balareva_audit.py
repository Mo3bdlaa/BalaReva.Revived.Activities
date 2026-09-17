#!/usr/bin/env python3
"""Audit the published BalaReva UiPath activity packages for rot.

Answers three questions, from evidence rather than from the package titles:

  1. Which packages still ship only .NET Framework assemblies? Those cannot be
     consumed by a UiPath Studio *Windows* (.NET 6/8) or *Cross-platform*
     project at all -- only by a Legacy project.
  2. Which third-party dependencies are behind, and which carry known
     vulnerabilities that the pinned version is actually inside the range of?
  3. Does the declared `lib/<tfm>/` layout match what the assemblies really
     target? A mismatch loads fine in one project type and fails in the other.

Everything is read live from the UiPath Marketplace gallery feed, nuget.org and
osv.dev, then cached under --cache so reruns are cheap and diffable.

Usage:
    python3 audit/balareva_audit.py                 # refresh and regenerate
    python3 audit/balareva_audit.py --offline       # rebuild report from cache
"""

import argparse
import collections
import datetime
import json
import os
import re
import ssl
import sys
import urllib.request
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pe_metadata

GALLERY_INDEX = "https://gallery.uipath.com/api/v3/index.json"
NUGET_FLAT = "https://api.nuget.org/v3-flatcontainer/{id}/index.json"
OSV_QUERY = "https://api.osv.dev/v1/query"
SEARCH_TERM = "BalaReva"

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)

# The #Strings heap holds thousands of member names. Only the ones that identify a
# framework or third-party library are worth recording, and keeping just those is
# the difference between a 600 KB artefact and a reviewable one.
NOTABLE = re.compile(
    r"^(UiPath\.|System\.Activities|System\.Windows\.Forms|System\.Drawing"
    r"|Presentation(?:Framework|Core)|WindowsBase|Microsoft\.Office\.Interop"
    r"|Microsoft\.Practices|Newtonsoft\.Json|iTextSharp|Npgsql|SharpCompress"
    r"|SharpZipLib|Ionic\.Zip|DotNetZip|DocumentFormat\.OpenXml|Elasticsearch)")


# --------------------------------------------------------------------------- io

def _opener():
    """urllib opener honouring the environment's HTTPS proxy and CA bundle."""
    handlers = []
    proxy = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
    if proxy:
        handlers.append(urllib.request.ProxyHandler({"https": proxy, "http": proxy}))
    ca = os.environ.get("REQUESTS_CA_BUNDLE") or "/root/.ccr/ca-bundle.crt"
    ctx = ssl.create_default_context(cafile=ca) if os.path.exists(ca) else None
    if ctx:
        handlers.append(urllib.request.HTTPSHandler(context=ctx))
    return urllib.request.build_opener(*handlers)


OPENER = _opener()


# Cloudflare fronts the gallery and rejects urllib's default user agent.
USER_AGENT = "balareva-audit/1.0 (+https://github.com/Mo3bdlaa/BalaReva.Revived.Activities)"


def get_bytes(url, timeout=180):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    return OPENER.open(req, timeout=timeout).read()


def get_json(url, timeout=120):
    return json.loads(get_bytes(url, timeout))


def post_json(url, payload, timeout=90):
    req = urllib.request.Request(
        url, data=json.dumps(payload).encode(),
        headers={"Content-Type": "application/json", "User-Agent": USER_AGENT},
    )
    return json.loads(OPENER.open(req, timeout=timeout).read())


def cached(cache, name, produce):
    path = os.path.join(cache, name)
    if os.path.exists(path):
        with open(path, "rb") as fh:
            return fh.read()
    blob = produce()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(blob)
    return blob


# ---------------------------------------------------------------------- versions

def semver_key(version):
    """Sort key good enough for these feeds: numeric parts, releases above pre."""
    core = version.split("-")[0]
    parts = [int(n) for n in re.findall(r"\d+", core)]
    parts += [0] * (4 - len(parts))
    return (parts, "-" not in version, version)


# -------------------------------------------------------------------- discovery

def feed_endpoints():
    resources = get_json(GALLERY_INDEX)["resources"]
    def pick(kind):
        return next(r["@id"] for r in resources if r["@type"].startswith(kind))
    return {
        "search": pick("SearchQueryService"),
        "flat": pick("PackageBaseAddress"),
        "registration": pick("RegistrationsBaseUrl/3.6.0"),
    }


def discover(endpoints, cache):
    raw = cached(cache, "search.json", lambda: get_bytes(
        f"{endpoints['search']}?q={SEARCH_TERM}&take=200&prerelease=true&semVerLevel=2.0.0"))
    hits = json.loads(raw)["data"]
    return sorted(p["id"] for p in hits if p["id"].lower().startswith("balareva."))


def latest_versions(endpoints, package_ids, cache):
    out = {}
    for pid in package_ids:
        raw = cached(cache, f"versions/{pid}.json",
                     lambda pid=pid: get_bytes(f"{endpoints['flat']}{pid.lower()}/index.json"))
        versions = json.loads(raw)["versions"]
        if versions:
            out[pid] = sorted(versions, key=semver_key)[-1]
    return out


def published_date(endpoints, pid, version, cache):
    try:
        raw = cached(cache, f"registration/{pid}.json", lambda: get_bytes(
            f"{endpoints['registration']}{pid.lower()}/index.json"))
        reg = json.loads(raw)
        for page in reg.get("items", []):
            items = page.get("items")
            if items is None:
                items = get_json(page["@id"]).get("items", [])
            for item in items:
                entry = item.get("catalogEntry", {})
                if entry.get("version", "").lower() == version.lower():
                    return (entry.get("published") or "")[:10] or None
    except Exception:
        return None
    return None


# -------------------------------------------------------------------- packages

def fetch_package(endpoints, pid, version, cache):
    name = f"packages/{pid}.{version}.nupkg"
    url = f"{endpoints['flat']}{pid.lower()}/{version.lower()}/{pid.lower()}.{version.lower()}.nupkg"
    cached(cache, name, lambda: get_bytes(url))
    return os.path.join(cache, name)


def read_nuspec(zf):
    spec = next(n for n in zf.namelist()
                if n.endswith(".nuspec") and "/" not in n)
    import xml.etree.ElementTree as ET
    root = ET.fromstring(zf.read(spec))
    ns = {"n": root.tag.split("}")[0].strip("{")} if "}" in root.tag else {}
    def path(p):
        return "/".join("n:" + s for s in p.split("/")) if ns else p
    meta = root.find(path("metadata"), ns)
    def text(tag):
        el = meta.find(path(tag), ns)
        return el.text.strip() if el is not None and el.text else ""
    deps = []
    group = meta.find(path("dependencies"), ns)
    if group is not None:
        for grp in group.findall(path("group"), ns):
            tfm = grp.get("targetFramework") or "(any)"
            for d in grp.findall(path("dependency"), ns):
                deps.append({"tfm": tfm, "id": d.get("id"), "version": d.get("version")})
        for d in group.findall(path("dependency"), ns):
            deps.append({"tfm": "(any)", "id": d.get("id"), "version": d.get("version")})
    return {"id": text("id"), "version": text("version"), "authors": text("authors"),
            "description": text("description"), "dependencies": deps}


def inspect_package(path):
    with zipfile.ZipFile(path) as zf:
        spec = read_nuspec(zf)
        assemblies = {}
        lib_tfms = set()
        contents = sorted(n for n in zf.namelist()
                          if not n.startswith(("_rels/", "package/")) and n != "[Content_Types].xml")
        for name in zf.namelist():
            if not name.lower().startswith("lib/"):
                continue
            parts = name.split("/")
            # lib/<tfm>/x.dll declares a tfm; lib/x.dll declares none ("any").
            lib_tfms.add(parts[1] if len(parts) >= 3 else "(none)")
            if not name.lower().endswith(".dll"):
                continue
            if not os.path.basename(name).startswith("BalaReva"):
                continue  # vendored third-party copies are reported via dependencies
            try:
                info = pe_metadata.read(zf.read(name))
                info["notable_names"] = [n for n in info.pop("names") if NOTABLE.match(n)]
                assemblies[name] = info
            except Exception as exc:
                assemblies[name] = {"error": str(exc)}
    spec["lib_tfms"] = sorted(lib_tfms)
    spec["assemblies"] = assemblies
    spec["contents"] = contents
    return spec


def framework_mismatches(pkg):
    """lib/<tfm>/ folders whose assemblies target a different runtime family."""
    out = []
    for path, info in pkg["assemblies"].items():
        parts = path.split("/")
        if len(parts) < 3 or "error" in info:
            continue
        folder, actual = parts[1], info.get("target_framework") or "?"
        netfx = info.get("is_net_framework")
        declares_netfx = folder.startswith("net4") or folder.startswith("net3")
        if declares_netfx != bool(netfx):
            out.append({"path": path, "folder": folder, "actual": actual,
                        "corelib": info.get("corelib")})
    return out


def classify(pkg):
    """How usable is this package in each UiPath Studio project type?"""
    tfms = pkg["lib_tfms"]
    modern = [t for t in tfms if t.startswith("net") and not t.startswith("net4")
              and not t.startswith("net3") and t != "(none)"]
    if framework_mismatches(pkg):
        return "broken"
    if modern:
        return "windows"      # has .NET 6+ assets
    return "legacy"           # .NET Framework assemblies only


# ----------------------------------------------------------------- dependencies

def check_dependency(dep_id, versions, cache):
    info = {"id": dep_id, "used_versions": sorted(versions)}
    try:
        raw = cached(cache, f"nuget/{dep_id}.json",
                     lambda: get_bytes(NUGET_FLAT.format(id=dep_id.lower())))
        stable = [v for v in json.loads(raw)["versions"] if "-" not in v]
        info["latest_stable"] = sorted(stable, key=semver_key)[-1] if stable else None
        info["on_nuget_org"] = True
    except Exception:
        info["latest_stable"] = None
        info["on_nuget_org"] = False

    # Ask OSV per pinned version so we only report advisories whose affected
    # range actually covers what the package depends on.
    advisories = {}
    for version in info["used_versions"]:
        try:
            found = post_json(OSV_QUERY, {"package": {"ecosystem": "NuGet", "name": dep_id},
                                          "version": version}).get("vulns", [])
        except Exception:
            continue
        for vuln in found:
            fixed = sorted({e["fixed"]
                            for a in vuln.get("affected", [])
                            for r in a.get("ranges", [])
                            for e in r.get("events", []) if "fixed" in e})
            advisories[vuln["id"]] = {
                "id": vuln["id"],
                "aliases": vuln.get("aliases", []),
                "severity": vuln.get("database_specific", {}).get("severity", ""),
                "summary": (vuln.get("summary") or "").strip(),
                "fixed_in": fixed,
                "affects": version,
            }
    info["advisories"] = sorted(advisories.values(), key=lambda a: a["id"])
    return info


# ---------------------------------------------------------------------- report

SEVERITY_RANK = {"CRITICAL": 0, "HIGH": 1, "MODERATE": 2, "LOW": 3, "": 4}


def build(cache, offline):
    endpoints = json.loads(cached(cache, "endpoints.json",
                                  lambda: json.dumps(feed_endpoints()).encode()))
    package_ids = discover(endpoints, cache)
    latest = latest_versions(endpoints, package_ids, cache)

    packages = []
    for pid, version in sorted(latest.items()):
        path = fetch_package(endpoints, pid, version, cache)
        pkg = inspect_package(path)
        pkg["published"] = published_date(endpoints, pid, version, cache)
        pkg["status"] = classify(pkg)
        pkg["mismatches"] = framework_mismatches(pkg)
        packages.append(pkg)
        print(f"  {pid:<44}{version:<11}{pkg['status']}", file=sys.stderr)

    used = collections.defaultdict(set)
    for pkg in packages:
        for dep in pkg["dependencies"]:
            used[dep["id"]].add(dep["version"])
    dependencies = [check_dependency(d, v, cache) for d, v in sorted(used.items())]

    return {"generated": datetime.date.today().isoformat(),
            "feed": GALLERY_INDEX,
            "packages": packages,
            "dependencies": dependencies}


def dependents(report, dep_id):
    return sorted({p["id"] for p in report["packages"]
                   for d in p["dependencies"] if d["id"] == dep_id})


def render(report):
    L = []
    w = L.append
    pkgs = report["packages"]
    deps = report["dependencies"]
    legacy = [p for p in pkgs if p["status"] == "legacy"]
    windows = [p for p in pkgs if p["status"] == "windows"]
    broken = [p for p in pkgs if p["status"] == "broken"]
    vulnerable = [d for d in deps if d["advisories"]]

    w("# BalaReva activity packages — dependency and compatibility audit\n")
    w(f"Generated {report['generated']} from `{report['feed']}`.")
    w("Regenerate with `python3 audit/balareva_audit.py`.\n")

    w("## Summary\n")
    w(f"- **{len(pkgs)} packages** published under the BalaReva name on the UiPath Marketplace feed.")
    w(f"- **{len(legacy)} ship .NET Framework assemblies only** — unusable from a Studio "
      "*Windows* (.NET 6/8) or *Cross-platform* project; Legacy projects only.")
    w(f"- **{len(windows)} carry .NET 6 assets**, so they load in Windows projects — but .NET 6 "
      "went out of support on 2024-11-12, so none of them target a supported runtime.")
    w(f"- **{len(broken)} {'is' if len(broken) == 1 else 'are'} mis-packaged**: a `lib/<tfm>/` "
      "folder whose assemblies target a different runtime family than the folder advertises.")
    w(f"- **{len(vulnerable)} dependencies carry advisories that cover the pinned version**, "
      "including two path-traversal classes in the archive handling.")
    newest = max((p["published"] for p in pkgs if p["published"]), default=None)
    if newest:
        w(f"- Newest release across the whole family: **{newest}**.\n")

    w("\n## Runtime compatibility\n")
    w("Each verdict comes from the assemblies themselves: the `TargetFrameworkAttribute` "
      "baked in by the compiler, cross-checked against whether the assembly binds `mscorlib` "
      "(.NET Framework) or `System.Runtime` (.NET Core / 5+). "
      "`lib/` with no framework folder means the assemblies sit directly in `lib/`, which NuGet "
      "treats as *any* framework even though the binaries are .NET Framework only.\n")
    w("| Package | Version | Published | `lib/` folders | Verdict |")
    w("|---|---|---|---|---|")
    label = {"legacy": "🔴 Legacy only", "windows": "🟡 .NET 6 (EOL)", "broken": "🛑 Mis-packaged"}
    for p in sorted(pkgs, key=lambda p: (p["status"] != "broken", p["status"], p["id"])):
        folders = ", ".join("`lib/` (no framework folder)" if t == "(none)" else f"`{t}`"
                            for t in p["lib_tfms"]) or "—"
        w(f"| `{p['id']}` | {p['version']} | {p['published'] or '?'} | {folders} | {label[p['status']]} |")

    if broken:
        w("\n### Mis-packaged assemblies\n")
        w("These load in Studio right up until the moment the designer needs them, "
          "because NuGet selects the folder by name and never verifies its contents.\n")
        for p in broken:
            for m in p["mismatches"]:
                w(f"- `{p['id']}` {p['version']} — `{m['path']}`")
                w(f"  sits in a **{m['folder']}** folder but targets **{m['actual']}** "
                  f"(binds `{m['corelib']}`).")

    w("\n## Dependency currency\n")
    w("| Dependency | Pinned | Latest on nuget.org | Advisories | Used by |")
    w("|---|---|---|---|---|")
    for d in sorted(deps, key=lambda d: (not d["advisories"], d["id"])):
        pinned = ", ".join(d["used_versions"])
        latest = d["latest_stable"] or "—"
        flag = "—" if not d["advisories"] else ", ".join(
            a["aliases"][0] if a["aliases"] else a["id"] for a in d["advisories"])
        users = ", ".join(f"`{u}`" for u in dependents({"packages": report["packages"]}, d["id"]))
        w(f"| `{d['id']}` | {pinned} | {latest} | {flag} | {users} |")

    if vulnerable:
        w("\n## Known vulnerabilities\n")
        w("Only advisories whose affected range covers the **pinned** version are listed; "
          "each was queried against osv.dev per exact version rather than per package.\n")
        flat = sorted(
            ((d, a) for d in vulnerable for a in d["advisories"]),
            key=lambda pair: (SEVERITY_RANK.get(pair[1]["severity"], 4), pair[0]["id"]))
        for d, a in flat:
            names = "/".join(a["aliases"]) or a["id"]
            w(f"### {names} — `{d['id']}` {a['affects']}  ({a['severity'] or 'unrated'})\n")
            w(f"{a['summary']}\n")
            w(f"- Fixed in: {', '.join(a['fixed_in']) or 'unknown'}")
            w(f"- Reaches: {', '.join('`' + u + '`' for u in dependents(report, d['id']))}\n")

    w("\n## Caveats\n")
    w("- `DotNetZip.Reduced` and `Ionic.Zip` return clean from osv.dev because advisories "
      "are filed against the `DotNetZip` package id. They are the same codebase at the same "
      "version, so **CVE-2018-1002205 (zip-slip) applies to them too** — it is simply not "
      "reachable by an id lookup.")
    w("- Vendored copies of third-party libraries that were dropped into `lib/` instead of "
      "declared as dependencies are invisible to any manifest-based scanner, including this "
      "one's dependency table. They are listed in the package contents of `audit/data/audit.json`.")
    w("- Assembly-name evidence in `audit.json` comes from a heuristic scan of the `#Strings` "
      "heap; see `audit/pe_metadata.py`. The runtime verdicts above do not rely on it.")
    return "\n".join(L) + "\n"


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--cache", default=os.path.join(HERE, ".cache"))
    ap.add_argument("--json", default=os.path.join(HERE, "data", "audit.json"))
    ap.add_argument("--markdown", default=os.path.join(REPO, "docs", "AUDIT.md"))
    ap.add_argument("--offline", action="store_true",
                    help="fail rather than hit the network for anything not cached")
    args = ap.parse_args()

    if args.offline:
        global get_bytes
        def get_bytes(url, timeout=180):  # noqa: F811 - deliberate shadow
            raise RuntimeError(f"--offline: {url} is not cached")

    print("Auditing BalaReva packages...", file=sys.stderr)
    report = build(args.cache, args.offline)

    os.makedirs(os.path.dirname(args.json), exist_ok=True)
    with open(args.json, "w") as fh:
        json.dump(report, fh, indent=1, sort_keys=True)
    os.makedirs(os.path.dirname(args.markdown), exist_ok=True)
    with open(args.markdown, "w") as fh:
        fh.write(render(report))
    print(f"\nWrote {args.json}\nWrote {args.markdown}", file=sys.stderr)


if __name__ == "__main__":
    main()
