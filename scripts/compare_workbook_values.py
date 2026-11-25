#!/usr/bin/env python3
"""Compare worksheet cell values and formulas in two XLSX workbooks."""

from __future__ import annotations

import argparse
import posixpath
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"m": MAIN}


def xml_part(target: str) -> str:
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))


def read_xlsx(path: Path):
    with zipfile.ZipFile(path) as archive:
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            strings = ["".join(t.text or "" for t in si.findall(".//m:t", NS)) for si in root.findall("m:si", NS)]
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        sheets = {}
        for sheet in workbook.findall("m:sheets/m:sheet", NS):
            name = sheet.attrib["name"]
            rid = sheet.attrib[f"{{{DOC_REL}}}id"]
            part = xml_part(relmap[rid])
            root = ET.fromstring(archive.read(part))
            cells = {}
            for cell in root.findall(".//m:sheetData/m:row/m:c", NS):
                ref = cell.attrib["r"]
                formula = cell.find("m:f", NS)
                value = cell.find("m:v", NS)
                kind = cell.attrib.get("t")
                if kind == "inlineStr":
                    val = "".join(t.text or "" for t in cell.findall(".//m:t", NS))
                elif value is None:
                    val = None
                elif kind == "s":
                    val = strings[int(value.text)]
                elif kind == "b":
                    val = value.text == "1"
                else:
                    val = value.text
                if val is not None or formula is not None:
                    cells[ref] = (formula.text if formula is not None else None, val)
            sheets[name] = cells
        return sheets


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("expected", type=Path, help="Authoritative thesis workbook")
    parser.add_argument("actual", type=Path, help="Workbook to compare")
    args = parser.parse_args()
    expected, actual = read_xlsx(args.expected), read_xlsx(args.actual)
    mismatches = []
    if set(expected) != set(actual):
        mismatches.append(("<sheets>", sorted(expected), sorted(actual)))
    for sheet in sorted(set(expected) & set(actual)):
        for cell in sorted(set(expected[sheet]) | set(actual[sheet])):
            if expected[sheet].get(cell) != actual[sheet].get(cell):
                mismatches.append((f"{sheet}!{cell}", expected[sheet].get(cell), actual[sheet].get(cell)))
    if mismatches:
        for mismatch in mismatches[:100]:
            print("DIFF", mismatch)
        print(f"FAILED: {len(mismatches)} cell/sheet differences")
        raise SystemExit(1)
    print(f"PASS: {args.expected.name} and {args.actual.name} have matching cell values and formulas across {len(expected)} worksheets")


if __name__ == "__main__":
    main()
