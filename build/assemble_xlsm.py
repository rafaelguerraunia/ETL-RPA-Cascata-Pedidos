# -*- coding: utf-8 -*-
"""Injects vbaProject.bin into the plain xlsx produced by
build_workbook_content.py, turning it into a macro-enabled .xlsm:
- adds xl/vbaProject.bin
- adds the vbaProject relationship to xl/_rels/workbook.xml.rels
- adds the Content_Types override for vbaProject.bin
- flips xl/workbook.xml's content type to the macro-enabled variant
"""

import zipfile
import re
import os

HERE = os.path.dirname(os.path.abspath(__file__))
SRC_XLSX = os.path.join(HERE, "workbook_plain.xlsx")
VBA_BIN = os.path.join(HERE, "vbaProject.bin")
OUT_XLSM = os.path.join(HERE, "..", "ETL_Cascata_Pedidos.xlsm")

VBA_RELS_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject"
VBA_CONTENT_TYPE = "application/vnd.ms-office.vbaProject"
MACRO_WORKBOOK_CONTENT_TYPE = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
PLAIN_WORKBOOK_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"


def next_rel_id(rels_xml: str) -> str:
    ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return f"rId{(max(ids) + 1) if ids else 1}"


def main():
    with open(VBA_BIN, "rb") as f:
        vba_bytes = f.read()

    zin = zipfile.ZipFile(SRC_XLSX, "r")
    names = zin.namelist()

    rels_path = "xl/_rels/workbook.xml.rels"
    rels_xml = zin.read(rels_path).decode("utf-8")
    new_id = next_rel_id(rels_xml)
    new_rel = f'<Relationship Id="{new_id}" Type="{VBA_RELS_TYPE}" Target="vbaProject.bin"/>'
    assert "</Relationships>" in rels_xml
    rels_xml = rels_xml.replace("</Relationships>", new_rel + "</Relationships>")

    ct_path = "[Content_Types].xml"
    ct_xml = zin.read(ct_path).decode("utf-8")
    # add vbaProject.bin override
    vba_override = f'<Override PartName="/xl/vbaProject.bin" ContentType="{VBA_CONTENT_TYPE}"/>'
    assert "</Types>" in ct_xml
    ct_xml = ct_xml.replace("</Types>", vba_override + "</Types>")
    # flip workbook.xml content type to macro-enabled
    before = f'<Override PartName="/xl/workbook.xml" ContentType="{PLAIN_WORKBOOK_CONTENT_TYPE}"/>'
    after = f'<Override PartName="/xl/workbook.xml" ContentType="{MACRO_WORKBOOK_CONTENT_TYPE}"/>'
    assert before in ct_xml, "workbook.xml content-type override not found as expected"
    ct_xml = ct_xml.replace(before, after)

    with zipfile.ZipFile(OUT_XLSM, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == rels_path:
                data = rels_xml.encode("utf-8")
            elif item.filename == ct_path:
                data = ct_xml.encode("utf-8")
            zout.writestr(item, data)
        zout.writestr("xl/vbaProject.bin", vba_bytes)

    zin.close()
    print("Wrote", OUT_XLSM)


if __name__ == "__main__":
    main()
