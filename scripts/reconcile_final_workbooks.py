#!/usr/bin/env python3
"""Apply the final S1/S2 text harmonization to saved 205-cluster build outputs.

Uses only Python's standard library. It writes corrected copies to a separate
output directory; it does not overwrite the supplied workbooks.
"""

from __future__ import annotations

import argparse
import posixpath
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"m": MAIN, "r": DOC_REL, "p": PKG_REL}
ET.register_namespace("", MAIN)
ET.register_namespace("r", DOC_REL)
ET.register_namespace("p", PKG_REL)

TITLE_P105 = "“Here, Let Me Help”: An Empirical Study of User Interventions in Human–Web Agent Collaboration"
SHORT_LABEL = "Secondary synthesis"
FULL_LABEL = "Secondary synthesis with conceptual output"
CODEBOOK_D4 = (
    "Proposes or synthesizes a construct, framework, method, or measure without direct empirical evaluation of the intended Alignment claim. "
    "Empirical data used only to elicit requirements or inform design did not count as evaluation of the proposed compatibility relation or Alignment claim."
)


def q(name: str) -> str:
    return f"{{{MAIN}}}{name}"


def package_part(target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))


class Workbook:
    def __init__(self, archive: zipfile.ZipFile):
        self.archive = archive
        self.shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        self.shared = [self._string_value(si) for si in self.shared_root.findall("m:si", NS)]
        self.shared_index = {value: i for i, value in enumerate(self.shared)}
        self.sheets = {}
        book = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        for sheet in book.findall("m:sheets/m:sheet", NS):
            name = sheet.attrib["name"]
            rid = sheet.attrib[f"{{{DOC_REL}}}id"]
            part = package_part(relmap[rid])
            root = ET.fromstring(archive.read(part))
            self.sheets[name] = (part, root)

    @staticmethod
    def _string_value(si: ET.Element) -> str:
        return "".join(node.text or "" for node in si.findall(".//m:t", NS))

    def cell(self, sheet_name: str, ref: str):
        _, root = self.sheets[sheet_name]
        return root.find(f".//m:sheetData/m:row/m:c[@r='{ref}']", NS)

    def value(self, sheet_name: str, ref: str):
        cell = self.cell(sheet_name, ref)
        if cell is None:
            return None
        if cell.attrib.get("t") == "inlineStr":
            return "".join(t.text or "" for t in cell.findall(".//m:t", NS))
        value = cell.find("m:v", NS)
        if value is None:
            return None
        if cell.attrib.get("t") == "s":
            return self.shared[int(value.text)]
        return value.text

    def set_text(self, sheet_name: str, ref: str, value: str):
        cell = self.cell(sheet_name, ref)
        if cell is None:
            raise ValueError(f"Missing cell {sheet_name}!{ref}")
        if value not in self.shared_index:
            si = ET.SubElement(self.shared_root, q("si"))
            text = ET.SubElement(si, q("t"))
            if value[:1].isspace() or value[-1:].isspace():
                text.set(f"{{{XML}}}space", "preserve")
            text.text = value
            self.shared_index[value] = len(self.shared)
            self.shared.append(value)
        index = self.shared_index[value]
        cell.attrib["t"] = "s"
        for child in list(cell):
            if child.tag in (q("v"), q("is"), q("f")):
                cell.remove(child)
        node = ET.SubElement(cell, q("v"))
        node.text = str(index)

    def updates(self):
        return {
            "xl/sharedStrings.xml": ET.tostring(self.shared_root, encoding="utf-8", xml_declaration=True),
            **{
                part: ET.tostring(root, encoding="utf-8", xml_declaration=True)
                for part, root in self.sheets.values()
            },
        }


def row_by_id(book: Workbook, sheet: str, id_column: str, cluster_id: str) -> int:
    _, root = book.sheets[sheet]
    for row in root.findall(".//m:sheetData/m:row", NS):
        row_number = int(row.attrib["r"])
        if book.value(sheet, f"{id_column}{row_number}") == cluster_id:
            return row_number
    raise ValueError(f"Could not find {cluster_id} in {sheet}")


def reconcile_s1(path: Path, output: Path) -> int:
    with zipfile.ZipFile(path, "r") as source:
        book = Workbook(source)
        label_rows = []
        _, root = book.sheets["Included studies"]
        for row in root.findall(".//m:sheetData/m:row", NS):
            row_number = int(row.attrib["r"])
            if book.value("Included studies", f"J{row_number}") == SHORT_LABEL:
                label_rows.append(row_number)
        # Depending on which harmonized labels were already present in the
        # 202-cluster input, either 31 or 32 rows may need this correction.
        if len(label_rows) not in (31, 32):
            raise ValueError(f"Expected 31 or 32 short S1 labels; found {len(label_rows)}")
        for row in label_rows:
            book.set_text("Included studies", f"J{row}", FULL_LABEL)
        p105 = row_by_id(book, "Included studies", "A", "P105")
        book.set_text("Included studies", f"F{p105}", TITLE_P105)
        write_xlsx(source, output, book.updates())
        return len(label_rows)


def reconcile_s2(path: Path, output: Path):
    with zipfile.ZipFile(path, "r") as source:
        book = Workbook(source)
        updates = []
        for sheet, id_col, title_col in [
            ("Cluster codes", "A", "D"),
            ("Publication status", "A", "F"),
        ]:
            row = row_by_id(book, sheet, id_col, "P105")
            book.set_text(sheet, f"{title_col}{row}", TITLE_P105)
            updates.append(f"{sheet}!{title_col}{row}")
        book.set_text("Codebook", "F14", CODEBOOK_D4)
        updates.append("Codebook!F14")
        write_xlsx(source, output, book.updates())
        return updates


def write_xlsx(source: zipfile.ZipFile, output: Path, replacements: dict[str, bytes]):
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as tmp:
        temp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
            for info in source.infolist():
                data = replacements.get(info.filename)
                if data is None:
                    data = source.read(info.filename)
                target.writestr(info, data)
        temp_path.replace(output)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True, help="Folder containing the generated 205-cluster S1 and S2")
    parser.add_argument("--output-dir", type=Path, required=True, help="Separate folder for reconciled copies")
    args = parser.parse_args()
    s1_name = "Supplementary_File_S1_Included_Study_Register_205.xlsx"
    s2_name = "Supplementary_File_S2_Evidence_Matrix_205.xlsx"
    s1_count = reconcile_s1(args.input_dir / s1_name, args.output_dir / s1_name)
    s2_cells = reconcile_s2(args.input_dir / s2_name, args.output_dir / s2_name)
    print(f"S1: normalized {s1_count} contribution labels and corrected P105 title")
    print(f"S2: corrected P105 title in {', '.join(s2_cells)}")


if __name__ == "__main__":
    main()
