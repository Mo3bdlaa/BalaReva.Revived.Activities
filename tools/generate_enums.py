#!/usr/bin/env python3
"""Generate C# enum declarations from the recorded surface of a published package.

Enum members are part of the binding contract: a workflow persists the member name,
and Studio persists the underlying number in some expression forms. Transcribing 27
enums with 54 members each by hand is how a digit goes wrong silently, so they are
generated straight from audit/data/api-surface.json instead.

Usage:
    python3 tools/generate_enums.py <package-id> <output-dir> [--namespace-root NS]
"""

import argparse
import json
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
SURFACE = os.path.join(os.path.dirname(HERE), "audit", "data", "api-surface.json")


def words(name):
    """Splits an identifier into readable words: 'FitFixedColumn' -> 'fit fixed column'."""
    name = re.sub(r"^(wd|xl|mso|pp)(?=[A-Z])", "", name)      # Office prefixes
    parts = re.findall(r"[A-Z]+(?![a-z])|[A-Z][a-z]*|\d+", name) or [name]
    return " ".join(parts).strip()


def summary(member):
    text = words(member)
    return text[0].upper() + text[1:].lower() if text else member


def render(type_name, namespace, members, source_id):
    lines = [
        f"namespace {namespace};",
        "",
        "/// <summary>",
        f"/// Mirrors the <c>{type_name}</c> of the published {source_id}.",
        "/// </summary>",
        "/// <remarks>",
        "/// Names and values are generated from the published assembly's metadata by",
        "/// tools/generate_enums.py. A workflow persists the member name, and the",
        "/// underlying number reaches .xaml through some expression forms, so neither",
        "/// may be renamed or renumbered.",
        "/// </remarks>",
        f"public enum {type_name.rsplit('.', 1)[-1]}",
        "{",
    ]
    for i, m in enumerate(members):
        if i:
            lines.append("")
        lines.append(f"    /// <summary>{summary(m['Name'])}.</summary>")
        lines.append(f"    {m['Name']} = {m['Value']},")
    lines.append("}")
    return "\n".join(lines) + "\n"


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("package_id")
    ap.add_argument("output_dir")
    ap.add_argument("--surface", default=SURFACE)
    args = ap.parse_args()

    with open(args.surface) as fh:
        surface = json.load(fh)
    package = next(p for p in surface["packages"] if p["Id"] == args.package_id)

    written = 0
    for entry in sorted(package["Types"], key=lambda t: t["FullName"]):
        if not entry["EnumMembers"]:
            continue
        namespace = entry["Namespace"]
        name = entry["Name"]
        folder = args.output_dir
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, f"{name}.cs")
        with open(path, "w") as fh:
            fh.write(render(entry["FullName"], namespace, entry["EnumMembers"], args.package_id))
        written += 1
        print(f"  {path}  ({len(entry['EnumMembers'])} members)")
    print(f"{written} enums generated for {args.package_id}")


if __name__ == "__main__":
    main()
