"""Merge research results and build the material/equipment passport table.

Usage: python scripts/build_table.py <research_dir>
Outputs: data/catalog.json, output/materials_equipment.xlsx
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
MATERIALS_FILE = ROOT / "data" / "materials.json"
CATALOG_FILE = ROOT / "data" / "catalog.json"
XLSX_FILE = ROOT / "output" / "materials_equipment.xlsx"
DOWNLOAD_REPORT = ROOT / "data" / "download_report.json"

MATERIAL_DIR = "pdf/materials"
EQUIPMENT_DIR = "pdf/equipment"


def slug(text: str, max_len: int = 60) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_")
    return cleaned[:max_len].rstrip("_") or "NA"


def material_filename(item: dict[str, Any], doc: dict[str, Any], seq: int) -> str:
    suffix = f"_{seq}" if seq > 1 else ""
    exactness = "" if doc.get("exact_or_equivalent", "exact") == "exact" else "_EQUIV"
    return (
        f"M{item['id']:02d}_{slug(item['name'])}_{slug(doc.get('type', 'DOC'), 10)}"
        f"{exactness}{suffix}.pdf"
    )


def equipment_filename(doc: dict[str, Any]) -> str:
    return (
        f"EQ_{slug(doc.get('vendor', ''), 20)}_{slug(doc.get('model', ''), 40)}"
        f"_{slug(doc.get('type', 'DOC'), 12)}.pdf"
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

    catalog: list[dict[str, Any]] = []
    for item in materials:
        found = research[item["id"]]
        per_type: dict[str, int] = {}
        material_docs = []
        for doc in found.get("material_docs", []):
            if not doc.get("url"):
                continue
            key = doc.get("type", "DOC")
            per_type[key] = per_type.get(key, 0) + 1
            material_docs.append(
                {**doc, "file": f"{MATERIAL_DIR}/{material_filename(item, doc, per_type[key])}"}
            )
        equipment_docs = [
            {**doc, "file": f"{EQUIPMENT_DIR}/{equipment_filename(doc)}"}
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
        name = Path(doc["file"]).name
        state = status.get(doc["file"])
        mark = "" if state == "ok" else " (yuklanmagan)"
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


def write_xlsx(catalog: list[dict[str, Any]]) -> None:
    status = load_download_status()
    wb = Workbook()
    main = wb.active
    main.title = "Jadval"
    main.append(["№", "Material nomi", "Material spetsifikatsiyasi (pasporti)",
                 "Material uchun jihoz", "Jihoz pasporti"])
    for item in catalog:
        main.append([
            item["id"],
            f"{item['name']}\n({item['sap']}; {item['spec']})",
            file_cell(item["material_docs"], status),
            equipment_cell(item["research"]),
            file_cell(item["equipment_docs"], status),
        ])
    style_sheet(main, [5, 40, 55, 60, 55])

    detail = wb.create_sheet("Manbalar")
    detail.append(["№", "Material", "Hujjat turi", "Fayl", "URL", "Aniq/ekvivalent",
                   "URL tekshirilgan", "Yuklash holati", "Izoh"])
    for item in catalog:
        for kind, docs in (("Material", item["material_docs"]), ("Jihoz", item["equipment_docs"])):
            for doc in docs:
                detail.append([
                    item["id"], item["name"], f"{kind}: {doc.get('type', '')}",
                    Path(doc["file"]).name, doc["url"],
                    doc.get("exact_or_equivalent", "-"),
                    "ha" if doc.get("verified") else "yo'q",
                    status.get(doc["file"], "yuklanmagan"), doc.get("note", ""),
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

    XLSX_FILE.parent.mkdir(parents=True, exist_ok=True)
    wb.save(XLSX_FILE)


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    catalog = build_catalog(Path(sys.argv[1]))
    CATALOG_FILE.write_text(json.dumps(catalog, ensure_ascii=False, indent=1), encoding="utf-8")
    write_xlsx(catalog)
    print(f"{len(catalog)} materials -> {XLSX_FILE.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
