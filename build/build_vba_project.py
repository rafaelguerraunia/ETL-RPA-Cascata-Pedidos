"""Builds a real vbaProject.bin (MS-OVBA compound file) from the .bas/.cls
source files, using cfb.py (CFB/OLE writer) and ovba_compress.py (MS-OVBA
'Compressed Container' codec). No Document/Sheet modules are declared (this
mirrors the well-established XlsxWriter add_vba_project() pattern: a
vbaProject.bin containing only Standard/Class modules, without ThisWorkbook /
SheetN Document modules, opens fine in Excel)."""

import struct
import uuid
import sys

sys.path.insert(0, '.')
from cfb import Storage, build_cfb
from ovba_compress import compress

CODEPAGE = 1252


def rec(id_, payload: bytes) -> bytes:
    return struct.pack("<HI", id_, len(payload)) + payload


def mbcs(s: str) -> bytes:
    return s.encode("cp1252")


def utf16(s: str) -> bytes:
    return s.encode("utf-16-le")


def name_unicode_record(id_, reserved_const, mbcs_bytes, unicode_bytes):
    """Layout used by REFERENCENAME / MODULESTREAMNAME / MODULEDOCSTRING /
    PROJECTDOCSTRING: Id(2) + SizeOfX(4) + X(mbcs) + Reserved(2) +
    SizeOfXUnicode(4) + XUnicode(utf16). NOTE: SizeOfX covers ONLY the mbcs
    field, not the whole record -- there is no separate outer 'Size' field."""
    return (struct.pack("<HI", id_, len(mbcs_bytes)) + mbcs_bytes +
            struct.pack("<H", reserved_const) +
            struct.pack("<I", len(unicode_bytes)) + unicode_bytes)


def dup_mbcs_record(id_, reserved_const, mbcs_bytes):
    """Layout used by PROJECTHELPFILEPATH: Id(2) + SizeOfHelpFile1(4) +
    HelpFile1(mbcs) + Reserved(2) + SizeOfHelpFile2(4) + HelpFile2(mbcs),
    where HelpFile2 MUST equal HelpFile1."""
    return (struct.pack("<HI", id_, len(mbcs_bytes)) + mbcs_bytes +
            struct.pack("<H", reserved_const) +
            struct.pack("<I", len(mbcs_bytes)) + mbcs_bytes)


def build_project_information():
    out = bytearray()
    out += rec(0x0001, struct.pack("<I", 1))          # PROJECTSYSKIND: Win32
    out += rec(0x0002, struct.pack("<I", 0x409))       # PROJECTLCID (MUST be 0x409)
    out += rec(0x0014, struct.pack("<I", 0x409))       # PROJECTLCIDINVOKE (MUST be 0x409)
    out += rec(0x0003, struct.pack("<H", CODEPAGE))    # PROJECTCODEPAGE
    out += rec(0x0004, mbcs("VBAProject"))              # PROJECTNAME
    out += name_unicode_record(0x0005, 0x0040, b"", b"")   # PROJECTDOCSTRING (empty)
    out += dup_mbcs_record(0x0006, 0x003D, b"")             # PROJECTHELPFILEPATH (empty)
    out += rec(0x0007, struct.pack("<I", 0))            # PROJECTHELPCONTEXT
    out += rec(0x0008, struct.pack("<I", 0))            # PROJECTLIBFLAGS
    # PROJECTVERSION: Id(2) + Reserved(4,=4) + VersionMajor(4) + VersionMinor(2)
    out += struct.pack("<HIIH", 0x0009, 4, 1, 0)
    out += name_unicode_record(0x000C, 0x003C, b"", b"")  # PROJECTCONSTANTS (empty)
    return bytes(out)


def build_reference(name: str, libid: str):
    out = bytearray()
    name_m = mbcs(name)
    name_u = utf16(name)
    out += name_unicode_record(0x0016, 0x003E, name_m, name_u)  # REFERENCENAME
    libid_m = mbcs(libid)
    out += struct.pack("<HI", 0x000D, 4 + len(libid_m) + 4 + 2)  # Size = SizeOfLibid+Libid+Reserved1+Reserved2
    out += struct.pack("<I", len(libid_m))
    out += libid_m
    out += struct.pack("<I", 0)   # Reserved1
    out += struct.pack("<H", 0)   # Reserved2
    return bytes(out)


def build_references():
    out = bytearray()
    out += build_reference(
        "VBA",
        r"*\G{000204EF-0000-0000-C000-000000000046}#4.2#9#C:\PROGRA~2\COMMON~1\MICROS~1\VBA\VBA7.1\VBE7.DLL#Visual Basic For Applications",
    )
    out += build_reference(
        "Excel",
        r"*\G{00020813-0000-0000-C000-000000000046}#1.9#0#C:\PROGRA~2\MICROS~1\Office16\EXCEL.EXE#Microsoft Excel 16.0 Object Library",
    )
    out += build_reference(
        "stdole",
        r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\SysWOW64\stdole2.tlb#OLE Automation",
    )
    return bytes(out)


def build_module_record(name: str, stream_name: str, is_class: bool, text_offset: int):
    out = bytearray()
    name_m = mbcs(name)
    out += rec(0x0019, name_m)                                   # MODULENAME
    name_u = utf16(name)
    out += rec(0x0047, name_u)                                   # MODULENAMEUNICODE (optional but included)
    sn_m = mbcs(stream_name)
    sn_u = utf16(stream_name)
    out += name_unicode_record(0x001A, 0x0032, sn_m, sn_u)   # MODULESTREAMNAME
    out += name_unicode_record(0x001C, 0x0048, b"", b"")     # MODULEDOCSTRING (empty)
    out += struct.pack("<HII", 0x0031, 4, text_offset)           # MODULEOFFSET
    out += rec(0x001E, struct.pack("<I", 0))                     # MODULEHELPCONTEXT
    out += rec(0x002C, struct.pack("<H", 0xFFFF))                # MODULECOOKIE
    out += rec(0x0022 if is_class else 0x0021, b"")               # MODULETYPE
    out += rec(0x002B, b"")                                       # terminator (Reserved=0)
    return bytes(out)


def build_dir_stream(modules):
    """modules: list of (name, stream_name, is_class, text_offset)"""
    out = bytearray()
    out += build_project_information()
    out += build_references()
    out += struct.pack("<HIH", 0x000F, 2, len(modules))  # PROJECTMODULES
    out += struct.pack("<HIH", 0x0013, 2, 0xFFFF)         # PROJECTCOOKIE
    for name, stream_name, is_class, text_offset in modules:
        out += build_module_record(name, stream_name, is_class, text_offset)
    out += rec(0x0010, b"")  # dir stream terminator (Reserved=0)
    return bytes(out)


def build_project_stream(std_modules, class_modules, project_guid):
    lines = [f'ID="{{{project_guid}}}"']
    for m in std_modules:
        lines.append(f"Module={m}")
    for c in class_modules:
        lines.append(f"Class={c}")
    lines.append('Name="VBAProject"')
    lines.append('HelpContextID="0"')
    lines.append('VersionCompatible32="393222000"')
    lines.append("")
    lines.append("[Host Extender Info]")
    text = "\r\n".join(lines) + "\r\n"
    return mbcs(text)


def build_projectwm(module_names):
    out = bytearray()
    for name in module_names:
        out += mbcs(name) + b"\x00"
        out += utf16(name) + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


def build_vba_project_stream():
    # _VBA_PROJECT header: Reserved1(0x61CC) + Version(2) + Reserved2(1)=0 + Reserved3(2)
    # No PerformanceCache -> Office recompiles from source unconditionally.
    return struct.pack("<HHBH", 0x61CC, 0xFFFF, 0x00, 0x0000)


def build(module_files, project_guid=None):
    """module_files: list of (module_name, is_class, source_text)"""
    if project_guid is None:
        project_guid = str(uuid.uuid4()).upper()

    module_infos = []
    module_streams = {}
    for name, is_class, text in module_files:
        source_bytes = mbcs(text)
        compressed = compress(source_bytes)
        module_streams[name] = compressed  # TextOffset = 0 (no perf cache prefix)
        module_infos.append((name, name, is_class, 0))

    dir_bytes = build_dir_stream(module_infos)
    dir_compressed = compress(dir_bytes)

    std_modules = [n for n, sn, is_class, off in module_infos if not is_class]
    class_modules = [n for n, sn, is_class, off in module_infos if is_class]
    project_stream = build_project_stream(std_modules, class_modules, project_guid)

    projectwm = build_projectwm([m[0] for m in module_infos])
    vba_project_stream = build_vba_project_stream()

    root = Storage("Root Entry")
    root.add_stream("PROJECT", project_stream)
    root.add_stream("PROJECTwm", projectwm)
    vba = root.add_storage("VBA")
    vba.add_stream("dir", dir_compressed)
    vba.add_stream("_VBA_PROJECT", vba_project_stream)
    for name, sn, is_class, off in module_infos:
        vba.add_stream(sn, module_streams[name])

    return build_cfb(root)


if __name__ == "__main__":
    import os

    vba_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "vba")

    std_order = [
        "modConstantes", "modLog", "modUtils", "modSAPConnection", "modConfig",
        "modSQVIEngine", "modExcelWrite", "modCascade", "modJoinBuilder", "modMain",
    ]
    class_order = ["clsSQVISpec", "clsFiltro", "clsRunResult", "clsLinhaResolvida"]

    module_files = []
    for name in std_order:
        path = os.path.join(vba_dir, f"{name}.bas")
        text = open(path, "r", encoding="utf-8").read()
        text = text.replace("\r\n", "\n").replace("\n", "\r\n")
        module_files.append((name, False, text))
    for name in class_order:
        path = os.path.join(vba_dir, f"{name}.cls")
        text = open(path, "r", encoding="utf-8").read()
        text = text.replace("\r\n", "\n").replace("\n", "\r\n")
        module_files.append((name, True, text))

    data = build(module_files)
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vbaProject.bin")
    with open(out_path, "wb") as f:
        f.write(data)
    print("Wrote", out_path, ":", len(data), "bytes")
