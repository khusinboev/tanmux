"""Merge research results and build the material/equipment passport table.

Usage: python scripts/build_table.py <research_dir>
Outputs: data/catalog.json,
         output/materials_equipment.xlsx (v1: normalized PDF file names),
         output/materials_equipment_links.xlsx (v2: direct PDF hyperlinks)
"""

from __future__ import annotations

import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlparse, urlsplit, urlunsplit

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
MATERIALS_FILE = ROOT / "data" / "materials.json"
CATALOG_FILE = ROOT / "data" / "catalog.json"
XLSX_FILE = ROOT / "output" / "materials_equipment.xlsx"
XLSX_LINKS_FILE = ROOT / "output" / "materials_equipment_links.xlsx"
DOWNLOAD_REPORT = ROOT / "data" / "download_report.json"

MATERIAL_DIR = "pdf/materials"
EQUIPMENT_DIR = "pdf/equipment"


PDF_URL_HINTS = ("getdocument", "/purl/", "msds.ford.com/documents")


def slug(text: str, max_len: int = 60) -> str:
    ascii_text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", ascii_text).strip("_")
    return cleaned[:max_len].rstrip("_") or "NA"


def is_pdf_url(url: str) -> bool:
    lowered = url.lower()
    return urlparse(lowered).path.endswith(".pdf") or any(h in lowered for h in PDF_URL_HINTS)


class FileNamer:
    """Assign one normalized file name per unique URL; HTML pages get no file."""

    def __init__(self) -> None:
        self._by_url: dict[str, str] = {}
        self._taken: set[str] = set()

    def assign(self, url: str, base: str, folder: str) -> str | None:
        if not is_pdf_url(url):
            return None
        if url in self._by_url:
            return self._by_url[url]
        name, seq = f"{base}.pdf", 1
        while name in self._taken:
            seq += 1
            name = f"{base}_{seq}.pdf"
        self._taken.add(name)
        self._by_url[url] = f"{folder}/{name}"
        return self._by_url[url]


def material_base(item: dict[str, Any], doc: dict[str, Any]) -> str:
    exactness = "" if doc.get("exact_or_equivalent", "exact") == "exact" else "_EQUIV"
    return f"M{item['id']:02d}_{slug(item['name'])}_{slug(doc.get('type', 'DOC'), 10)}{exactness}"


def equipment_base(doc: dict[str, Any]) -> str:
    return (
        f"EQ_{slug(doc.get('vendor', ''), 20)}_{slug(doc.get('model', ''), 40)}"
        f"_{slug(doc.get('type', 'DOC'), 12)}"
    )


def load_research(research_dir: Path) -> dict[int, dict[str, Any]]:
    merged: dict[int, dict[str, Any]] = {}
    for path in sorted(research_dir.glob("group_*.json")):
        for entry in json.loads(path.read_text(encoding="utf-8")):
            merged[int(entry["id"])] = entry
    return merged


def build_catalog(research_dir: Path) -> list[dict[str, Any]]:
    materials = json.loads(MATERIALS_FILE.read_text(encoding="utf-8"))
    research = load_research(research_dir)
    missing = [m["id"] for m in materials if m["id"] not in research]
    if missing:
        raise SystemExit(f"Research missing for material ids: {missing}")

    namer = FileNamer()
    catalog: list[dict[str, Any]] = []
    for item in materials:
        found = research[item["id"]]
        material_docs = [
            {**doc, "file": namer.assign(doc["url"], material_base(item, doc), MATERIAL_DIR)}
            for doc in found.get("material_docs", [])
            if doc.get("url")
        ]
        equipment_docs = [
            {**doc, "file": namer.assign(doc["url"], equipment_base(doc), EQUIPMENT_DIR)}
            for doc in found.get("equipment_docs", [])
            if doc.get("url")
        ]
        catalog.append({**item, "research": found, "material_docs": material_docs,
                        "equipment_docs": equipment_docs})
    return catalog


def load_download_status() -> dict[str, str]:
    if not DOWNLOAD_REPORT.exists():
        return {}
    return {row["file"]: row["status"] for row in json.loads(DOWNLOAD_REPORT.read_text())}


def file_cell(docs: list[dict[str, Any]], status: dict[str, str]) -> str:
    if not docs:
        return "Ochiq manbada topilmadi"
    lines = []
    for doc in docs:
        if doc["file"] is None:
            lines.append(f"[veb-sahifa, PDF emas] {doc['url']}")
            continue
        name = Path(doc["file"]).name
        mark = "" if status.get(doc["file"]) == "ok" else " (yuklanmagan)"
        lines.append(name + mark)
    return "\n".join(dict.fromkeys(lines))


def equipment_cell(research: dict[str, Any]) -> str:
    rows = [
        f"{e.get('vendor', '')} {e.get('model', '')} — {e.get('role', '')} [{e.get('confidence', '?')}]"
        for e in research.get("equipment", [])
    ]
    return "\n".join(rows) or "Aniqlanmadi"


def style_sheet(ws, widths: list[int]) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


HEADERS = ["№", "Material nomi", "Material spetsifikatsiyasi (pasporti)",
           "Material uchun jihoz", "Jihoz pasporti"]
MAIN_WIDTHS = [5, 40, 55, 60, 55]


def material_title(item: dict[str, Any]) -> str:
    return f"{item['name']}\n({item['sap']}; {item['spec']})"


def encode_url(url: str) -> str:
    parts = urlsplit(url)
    return urlunsplit(parts._replace(path=quote(parts.path, safe="/%:@+-._~")))


def write_file_name_sheet(ws, catalog: list[dict[str, Any]], status: dict[str, str]) -> None:
    ws.append(HEADERS)
    for item in catalog:
        ws.append([
            item["id"],
            material_title(item),
            file_cell(item["material_docs"], status),
            equipment_cell(item["research"]),
            file_cell(item["equipment_docs"], status),
        ])
    style_sheet(ws, MAIN_WIDTHS)


def pdf_docs(docs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return list({d["url"]: d for d in docs if d["file"]}.values())


def put_link(cell, doc: dict[str, Any], label: str) -> None:
    url = encode_url(doc["url"])
    cell.value = f"{label}\n{url}"
    cell.hyperlink = url
    cell.font = Font(color="0563C1", underline="single")


def write_link_sheet(ws, catalog: list[dict[str, Any]]) -> None:
    """One material spans several rows: one direct PDF link per cell."""
    ws.append(HEADERS)
    row = 2
    for item in catalog:
        mat, eq = pdf_docs(item["material_docs"]), pdf_docs(item["equipment_docs"])
        span = max(len(mat), len(eq), 1)
        ws.cell(row, 1, item["id"])
        ws.cell(row, 2, material_title(item))
        ws.cell(row, 4, equipment_cell(item["research"]))
        for offset in range(span):
            r = row + offset
            if offset < len(mat):
                d = mat[offset]
                mark = "" if d.get("exact_or_equivalent", "exact") == "exact" else ", ekvivalent"
                put_link(ws.cell(r, 3), d, f"{d.get('type', 'DOC')}{mark}")
            if offset < len(eq):
                d = eq[offset]
                put_link(ws.cell(r, 5), d,
                         f"{d.get('vendor', '')} {d.get('model', '')} ({d.get('type', '')})")
        if not mat:
            ws.cell(row, 3, "PDF topilmadi (Manbalar varag'iga qarang)")
        if not eq:
            ws.cell(row, 5, "PDF topilmadi")
        if span > 1:
            for col in (1, 2, 4):
                ws.merge_cells(start_row=row, start_column=col, end_row=row + span - 1,
                               end_column=col)
        row += span
    style_sheet(ws, MAIN_WIDTHS)
    for cells in ws.iter_rows(min_row=2):
        for cell in cells:
            if cell.hyperlink:
                cell.font = Font(color="0563C1", underline="single")


def write_xlsx(catalog: list[dict[str, Any]], target: Path, with_links: bool) -> None:
    status = load_download_status()
    wb = Workbook()
    main = wb.active
    main.title = "Jadval"
    if with_links:
        write_link_sheet(main, catalog)
    else:
        write_file_name_sheet(main, catalog, status)

    detail = wb.create_sheet("Manbalar")
    detail.append(["№", "Material", "Hujjat turi", "Fayl", "URL", "Aniq/ekvivalent",
                   "URL tekshirilgan", "Yuklash holati", "Izoh"])
    for item in catalog:
        for kind, docs in (("Material", item["material_docs"]), ("Jihoz", item["equipment_docs"])):
            for doc in docs:
                detail.append([
                    item["id"], item["name"], f"{kind}: {doc.get('type', '')}",
                    Path(doc["file"]).name if doc["file"] else "(veb-sahifa)", doc["url"],
                    doc.get("exact_or_equivalent", "-"),
                    "ha" if doc.get("verified") else "yo'q",
                    status.get(doc["file"], "yuklanmagan") if doc["file"] else "-",
                    doc.get("note", ""),
                ])
    style_sheet(detail, [5, 35, 18, 50, 60, 14, 12, 14, 50])

    ident = wb.create_sheet("Identifikatsiya")
    ident.append(["№", "SAP", "Material", "Spec (manba)", "Qadoq", "Vazifa",
                  "Ishlab chiqaruvchi", "Aniqlangan mahsulot", "Ishonch", "Jihoz asoslanishi", "Izoh"])
    for item in catalog:
        r = item["research"]
        ident.append([
            item["id"], item["sap"], item["name"], item["spec"], item["package"], item["function"],
            r.get("manufacturer", ""), r.get("product_identified", ""),
            r.get("manufacturer_confidence", ""),
            "\n".join(f"{e.get('vendor', '')} {e.get('model', '')}: {e.get('rationale', '')}"
                      for e in r.get("equipment", [])),
            r.get("notes", ""),
        ])
    style_sheet(ident, [5, 14, 35, 16, 14, 18, 25, 35, 10, 70, 60])

    target.parent.mkdir(parents=True, exist_ok=True)
    wb.save(target)


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    catalog = build_catalog(Path(sys.argv[1]))
    CATALOG_FILE.write_text(json.dumps(catalog, ensure_ascii=False, indent=1), encoding="utf-8")
    for target, with_links in ((XLSX_FILE, False), (XLSX_LINKS_FILE, True)):
        write_xlsx(catalog, target, with_links)
        print(f"{len(catalog)} materials -> {target.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
