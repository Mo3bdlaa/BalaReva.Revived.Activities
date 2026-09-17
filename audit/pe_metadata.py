"""Minimal, dependency-free reader for ECMA-335 (CLI) assembly metadata.

Only what the audit needs, and deliberately split by confidence level:

Authoritative
    target_framework  Value of [assembly: TargetFrameworkAttribute(..)], read as a
                      literal from the #Blob heap. This is the moniker the compiler
                      baked in, so it is what the assembly really targets.
    corelib           'mscorlib' for .NET Framework, 'System.Runtime' for .NET
                      Core / .NET 5+. An independent cross-check on the moniker:
                      the two can never legitimately disagree.

Heuristic
    names             Every identifier-looking string in the #Strings heap. That
                      heap holds assembly-reference names, but also type names,
                      method names and namespaces, so a hit means "this name
                      appears somewhere in the metadata", NOT "this assembly has
                      an AssemblyRef to it". Good enough to spot which third-party
                      libraries a package reaches for; do not use it to build a
                      precise dependency graph.

Parsing the AssemblyRef table properly would require the row layout of all 34
preceding metadata tables. That is not worth it here: every conclusion the audit
actually draws rests on the authoritative fields above.
"""

import re
import struct

_TFM_RE = re.compile(rb"\.(?:NETFramework|NETCoreApp|NETStandard),Version=v[\d.]+")
_NAME_RE = re.compile(rb"[A-Za-z][\w.\-]{2,90}\Z")


class NotManaged(Exception):
    """Raised when a file is not a PE image carrying a CLI metadata root."""


def _rva_to_offset(rva, sections):
    for virt_addr, virt_size, raw_ptr, raw_size in sections:
        if virt_addr <= rva < virt_addr + max(virt_size, raw_size):
            return raw_ptr + (rva - virt_addr)
    raise NotManaged(f"RVA 0x{rva:x} falls outside every section")


def _sections(data):
    if data[:2] != b"MZ":
        raise NotManaged("not a PE image (no MZ signature)")
    pe = struct.unpack_from("<I", data, 0x3C)[0]
    if data[pe : pe + 4] != b"PE\0\0":
        raise NotManaged("not a PE image (no PE signature)")
    coff = pe + 4
    section_count, = struct.unpack_from("<H", data, coff + 2)
    opt_size, = struct.unpack_from("<H", data, coff + 16)
    opt = coff + 20
    magic, = struct.unpack_from("<H", data, opt)
    # The data directory sits after the optional header's fixed part, whose size
    # differs between PE32 (0x10b) and PE32+ (0x20b).
    data_dirs = opt + (96 if magic == 0x10B else 112)
    table = opt + opt_size
    sections = []
    for i in range(section_count):
        entry = table + i * 40
        virt_size, virt_addr, raw_size, raw_ptr = struct.unpack_from("<IIII", data, entry + 8)
        sections.append((virt_addr, virt_size, raw_ptr, raw_size))
    return sections, data_dirs


def _streams(data):
    """Return {stream name: (absolute offset, size)} from the CLI metadata root."""
    sections, data_dirs = _sections(data)
    # Data directory 14 is the CLI header.
    cli_rva, _ = struct.unpack_from("<II", data, data_dirs + 14 * 8)
    if not cli_rva:
        raise NotManaged("no CLI header: native, not a managed assembly")
    cli = _rva_to_offset(cli_rva, sections)
    meta_rva, _ = struct.unpack_from("<II", data, cli + 8)
    root = _rva_to_offset(meta_rva, sections)
    if data[root : root + 4] != b"BSJB":
        raise NotManaged("CLI metadata root has no BSJB signature")

    version_len, = struct.unpack_from("<I", data, root + 12)
    pos = root + 16 + version_len + 2  # skip version string and flags
    stream_count, = struct.unpack_from("<H", data, pos)
    pos += 2

    out = {}
    for _ in range(stream_count):
        offset, size = struct.unpack_from("<II", data, pos)
        pos += 8
        end = data.index(b"\0", pos)
        out[data[pos:end].decode("ascii", "replace")] = (root + offset, size)
        pos = (end + 1 + 3) & ~3  # stream headers are 4-byte aligned
    return out


def read(data):
    """Inspect one assembly. Returns a dict; see the module docstring."""
    streams = _streams(data)

    target_framework = None
    if "#Blob" in streams:
        offset, size = streams["#Blob"]
        match = _TFM_RE.search(data[offset : offset + size])
        if match:
            target_framework = match.group(0).decode()

    names = set()
    if "#Strings" in streams:
        offset, size = streams["#Strings"]
        for token in data[offset : offset + size].split(b"\0"):
            if token and _NAME_RE.match(token):
                names.add(token.decode())

    corelib = next((c for c in ("mscorlib", "System.Runtime") if c in names), None)
    return {
        "target_framework": target_framework,
        "corelib": corelib,
        "is_net_framework": bool(target_framework and "NETFramework" in target_framework),
        "names": sorted(names),
    }
