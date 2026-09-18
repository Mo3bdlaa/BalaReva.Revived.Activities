#!/usr/bin/env python3
"""Fail if a revived package no longer offers what the published one did.

A .xaml workflow binds by type full name, property name and property type. Rename
any of the three and the workflow breaks on upgrade, silently at design time and
loudly at run time. This compares the surface dumped from our own built .nupkg
against audit/data/api-surface.json, which was recorded off the packages on the
UiPath gallery, and reports every difference.

Design-time assemblies are skipped: the published packages ship activity designers
in a separate *.Design.dll, which is WPF and is not part of what a workflow binds.

Usage:
    dotnet run --project tools/ApiSurface -- artifacts ours.json
    python3 audit/verify_binding_surface.py ours.json
"""

import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
RECORDED = os.path.join(HERE, "data", "api-surface.json")

# Differences we have decided to live with, each with the reason. Anything not listed
# here is a failure. Keep this list short and keep the reasons specific: an entry is a
# workflow that will not upgrade cleanly, and it should cost something to add one.
ACCEPTED = {
    ("BalaReva.EasyPowerPoint.Scope.Main.PowerPointObject", "PptPersentation"):
        "Published as Microsoft.Office.Interop.PowerPoint.Presentation, declared here as "
        "object. Every Office interop assembly on NuGet has a hard reference on 'office' "
        "(Microsoft.Office.Core), which Microsoft publishes nowhere, so a member typed that "
        "way makes the whole scope fail to load at run time rather than only this property. "
        "A workflow reaching for the interop type has to cast.",
    ("BalaReva.EasyExcel.Main.ExcelParam", "ExcelWorkBook"):
        "Published as Microsoft.Office.Interop.Excel.Workbook, declared here as object, "
        "for the same reason as PowerPointObject.PptPersentation above.",
}


def published_id(revived_id):
    """BalaReva.Revived.Excel.Activities -> BalaReva.Excel.Activities."""
    return revived_id.replace(".Revived.", ".", 1)


def binding_types(package):
    """The types a workflow can bind to: everything outside the design assembly."""
    return {
        t["FullName"]: t
        for t in package["Types"]
        if not t["Assembly"].lower().endswith(".design.dll")
    }


def bindable(types, full_name):
    """Every property a workflow can set on this type, inherited ones included.

    ApiSurface records declared properties only, but a .xaml binds whatever the type
    exposes, so a property the published package declared on the activity and we moved
    up to a shared base is the same property as far as a workflow is concerned. The
    walk stops at the first base outside the package, which is where CodeActivity and
    NativeActivity live.
    """
    properties = {}
    seen = set()
    name = full_name
    while name in types and name not in seen:
        seen.add(name)
        for prop in types[name]["Properties"]:
            properties.setdefault(prop["Name"], prop["Type"])
        name = types[name]["BaseType"]
    return properties


def compare(published, ours):
    """Yield one line per difference."""
    theirs = binding_types(published)
    mine = binding_types(ours)

    for full_name, original in sorted(theirs.items()):
        revived = mine.get(full_name)
        if revived is None:
            yield f"{full_name}: type is gone"
            continue

        have = bindable(mine, full_name)
        for name, expected in bindable(theirs, full_name).items():
            actual = have.get(name)
            if (full_name, name) in ACCEPTED:
                continue
            if actual is None:
                yield f"{full_name}.{name}: missing, was {expected}"
            elif actual != expected:
                yield f"{full_name}.{name}: was {expected}, is now {actual}"

        for member in original["EnumMembers"]:
            if member not in revived["EnumMembers"]:
                yield f"{full_name}.{member}: enum member is gone"


def main(argv):
    if len(argv) != 2:
        print(__doc__, file=sys.stderr)
        return 2

    with open(argv[1]) as handle:
        ours = {p["Id"]: p for p in json.load(handle)["packages"]}
    with open(RECORDED) as handle:
        theirs = {p["Id"]: p for p in json.load(handle)["packages"]}

    failed = False
    for revived_id, package in sorted(ours.items()):
        original = theirs.get(published_id(revived_id))
        if original is None:
            print(f"{revived_id}: no recorded surface to compare against, skipped")
            continue

        differences = list(compare(original, package))
        if differences:
            failed = True
            print(f"\n{revived_id}: {len(differences)} difference(s) from "
                  f"{original['Id']} {original.get('Version', '')}".rstrip())
            for line in differences:
                print(f"  {line}")
        else:
            print(f"{revived_id}: binds the same as {original['Id']}")

        for (type_name, member), reason in sorted(ACCEPTED.items()):
            if type_name in binding_types(original) and type_name in binding_types(package):
                print(f"  accepted: {type_name}.{member} - {reason}")

    if failed:
        print("\nWorkflows binding to the published names would break on these.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
