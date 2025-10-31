#!/usr/bin/env python3
"""
Create CSV exports for a mix of CSV inputs and specific worksheets from an Excel
workbook. The script copies the existing CSV files as-is and converts workbook
tabs into UTF-8 encoded CSV files so downstream tooling can consume them.
"""
from __future__ import annotations

import csv
import re
import shutil
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Sequence
from xml.etree import ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def normalize(value: str | None) -> str:
    if value is None:
        return ""
    return value.replace("\r", " ").replace("\n", " ").strip()


def excel_column_index(reference: str) -> int:
    match = re.match(r"([A-Z]+)", reference.upper())
    if not match:
        return 0
    index = 0
    for char in match.group(1):
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def load_shared_strings(archive: zipfile.ZipFile) -> List[str]:
    try:
        raw = archive.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(raw)
    ns = {"m": NS_MAIN}
    shared: List[str] = []
    for entry in root.findall("m:si", ns):
        fragments: List[str] = []
        for node in entry.findall(".//m:t", ns):
            text = node.text or ""
            if node.attrib.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve":
                fragments.append(text)
            else:
                fragments.append(text.strip())
        shared.append(normalize("".join(fragments)))
    return shared


def workbook_relationships(archive: zipfile.ZipFile) -> Dict[str, str]:
    try:
        raw = archive.read("xl/_rels/workbook.xml.rels")
    except KeyError:
        return {}
    ns = {"r": NS_PKG_REL}
    root = ET.fromstring(raw)
    mapping: Dict[str, str] = {}
    for rel in root.findall("r:Relationship", ns):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rel_id and target:
            mapping[rel_id] = target
    return mapping


def workbook_sheets(archive: zipfile.ZipFile) -> Dict[str, str]:
    raw = archive.read("xl/workbook.xml")
    ns = {"m": NS_MAIN, "r": NS_REL}
    root = ET.fromstring(raw)
    relations = workbook_relationships(archive)
    sheets: Dict[str, str] = {}
    for sheet in root.findall("m:sheets/m:sheet", ns):
        name = sheet.attrib.get("name")
        rel_id = sheet.attrib.get(f"{{{NS_REL}}}id")
        if not name or not rel_id:
            continue
        target = relations.get(rel_id)
        if not target:
            continue
        if target.startswith("/"):
            sheet_path = target.lstrip("/")
        elif target.startswith("xl/"):
            sheet_path = target
        else:
            sheet_path = f"xl/{target}"
        sheets[name] = sheet_path
    return sheets


def load_sheet_rows(archive: zipfile.ZipFile, sheet_path: str, shared: Sequence[str]) -> List[List[str]]:
    raw = archive.read(sheet_path)
    root = ET.fromstring(raw)
    ns = {"m": NS_MAIN}
    rows: List[List[str]] = []
    for row in root.findall("m:sheetData/m:row", ns):
        cells: List[str] = []
        for cell in row.findall("m:c", ns):
            ref = cell.attrib.get("r", "")
            index = excel_column_index(ref) if ref else len(cells)
            while len(cells) < index:
                cells.append("")
            cell_type = cell.attrib.get("t")
            text = ""
            if cell_type == "s":
                value = cell.find("m:v", ns)
                if value is not None and value.text is not None:
                    try:
                        text = shared[int(value.text)]
                    except (ValueError, IndexError):
                        text = value.text
            elif cell_type == "b":
                value = cell.find("m:v", ns)
                text = "TRUE" if value is not None and value.text == "1" else "FALSE"
            elif cell_type == "inlineStr":
                inline = cell.findall("m:is/m:t", ns)
                text = "".join((node.text or "").strip() for node in inline)
            else:
                value = cell.find("m:v", ns)
                if value is not None and value.text is not None:
                    text = value.text
                else:
                    inline = cell.findall("m:is/m:t", ns)
                    if inline:
                        text = "".join((node.text or "").strip() for node in inline)
            cells.append(normalize(text))
        rows.append(cells)
    return rows


def write_csv(rows: Sequence[Sequence[str]], output_path: Path) -> None:
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle, lineterminator="\r\n")
        for row in rows:
            writer.writerow(list(row))


def convert_xlsx_to_csv(input_path: Path, sheets: Iterable[str], output_dir: Path) -> None:
    with zipfile.ZipFile(input_path) as archive:
        shared = load_shared_strings(archive)
        sheet_map = workbook_sheets(archive)
        for sheet_name in sheets:
            sheet_path = sheet_map.get(sheet_name)
            if not sheet_path:
                raise ValueError(f"Sheet '{sheet_name}' not found in {input_path}")
            rows = load_sheet_rows(archive, sheet_path, shared)
            safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", sheet_name).strip("_") or "Sheet"
            output_path = output_dir / f"{input_path.stem}_{safe_name}.csv"
            write_csv(rows, output_path)


def main() -> None:
    project_root = Path(__file__).resolve().parent
    output_dir = project_root / "csv"
    ensure_dir(output_dir)

    csv_sources = [
        project_root / "build" / "ephemeris_2025-10-01T00-00_to_2025-12-01T00-00_UTC_60min.csv",
        project_root / "bdayephemeris.csv",
    ]
    for source in csv_sources:
        if not source.exists():
            raise FileNotFoundError(source)
        destination = output_dir / source.name
        shutil.copy2(source, destination)

    workbook = project_root / "build" / "Cavin_Ephemeris_Events_2025_10-11.xlsx"
    if not workbook.exists():
        raise FileNotFoundError(workbook)
    convert_xlsx_to_csv(workbook, ["Events", "Day_Metrics", "Settings"], output_dir)


if __name__ == "__main__":
    main()
