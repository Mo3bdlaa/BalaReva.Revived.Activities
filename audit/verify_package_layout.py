#!/usr/bin/env python3
"""Fail if a built .nupkg puts assemblies in a lib/ folder they do not target.

This is the exact defect the audit found in BalaReva.EasyDataTable.Activities
5.0.0: a .NET Framework 4.6.1 design assembly shipped inside lib/net6.0-windows7.0.
NuGet picks a folder by name and never looks inside it, so the package restores
cleanly and only fails later, at the point of use. Running this over our own
output keeps us from shipping the same thing.

Usage:
    python3 audit/verify_package_layout.py artifacts/*.nupkg
"""

import glob
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pe_metadata


def framework_family(moniker):
    """'net48' and 'net461' are .NET Framework; 'net8.0', 'net6.0-windows' are not."""
    name = moniker.lower()
    return "netfx" if name.startswith("net4") or name.startswith("net3") else "netcore"


def check(path):
    problems = []
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.lower().endswith(".dll"):
                continue
            parts = name.split("/")
            if len(parts) < 3 or parts[0].lower() != "lib":
                continue
            folder = parts[1]
            try:
                info = pe_metadata.read(archive.read(name))
            except pe_metadata.NotManaged:
                continue  # native payloads have no managed metadata to disagree with

            target = info["target_framework"]
            if target is None:
                continue
            declared = framework_family(folder)
            actual = "netfx" if info["is_net_framework"] else "netcore"
            if declared != actual:
                problems.append(
                    f"{name}: folder declares {folder} ({declared}) "
                    f"but the assembly targets {target} ({actual}, binds {info['corelib']})")
    return problems


def main(argv):
    paths = [p for pattern in argv for p in sorted(glob.glob(pattern))]
    paths = [p for p in paths if not p.endswith(".snupkg")]
    if not paths:
        print("no packages matched", file=sys.stderr)
        return 1

    failed = False
    for path in paths:
        problems = check(path)
        if problems:
            failed = True
            print(f"FAIL {path}")
            for problem in problems:
                print(f"     {problem}")
        else:
            print(f"ok   {path}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
