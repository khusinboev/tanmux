"""Build output/material_equipment_names.xlsx: clean material names + equipment names.

Passport columns are intentionally left empty. Second sheet ("Izoh") carries SAP,
source spec and equipment confidence so each name can be traced.
"""

from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output" / "material_equipment_names.xlsx"

ED_LINE = ["Dürr Ecopaint RoDip", "Controlled Power Company AACD Series 50 rectifier",
           "Koch Membrane Systems KPAK ultrafiltration module"]
SEALER = ["Graco E-Flo SP", "Dürr EcoShot Meter", "Dürr EcoGun2 3D"]
UNDERBODY = ["Graco King XL10K", "Graco AL Series Automatic Airless Spray Gun (288048)"]
PAINT_ROBOT = ["Dürr EcoRP E043i", "Dürr EcoBell3"]
BASECOAT = [*PAINT_ROBOT, "Dürr EcoMCC3 20"]
MIX_ROOM = ["Dürr EcoSupply2 Core", "Dürr EcoMCC3 20"]
PURGE = ["Dürr EcoBell Cleaner D2", "Dürr EcoBell PurgeBox"]
SOMAC_CUBE = ["Dürr Somac ProLine Cube", "Dürr Somac G4 Blue filling adapter"]
SOMAC_OIL = ["Dürr Somac ProLine line filling station"]
PRIMER_APPL = ["Nordson EFD 752V diaphragm valve", "Graco agitated pressure tank"]
GREASE = ["Graco Fire-Ball 300", "Lincoln PowerMaster II"]

# (id, material name, equipment names)
ROWS: list[tuple[int, str, list[str]]] = [
    (1, "CHEMFOS 740R", ["Dürr Ecopaint RoDip"]),
    (2, "ED7010 Binder Gray", ED_LINE),
    (3, "ED7010 Paste Gray", ED_LINE),
    (4, "BODY SEALER RSB-430", SEALER),
    (5, "UBC UNDER COAT RSU-230", UNDERBODY),
    (6, "UBS COAT", UNDERBODY),
    (7, "PVC ANTI CHIP", UNDERBODY),
    (8, "CR-100 PRIMER SURFACER LIGHT GRAY (N7.0)", PAINT_ROBOT),
    (9, "CR-100 PRIMER SURFACER PURE WHITE (N9.0)", PAINT_ROBOT),
    (10, "PRIMER TYPE THINNER", MIX_ROOM),
    (11, "(GAZ) BASE COAT SUMMIT WHITE (2C1B)", BASECOAT),
    (12, "(GK2) BASE COAT SON OF A GUN GRAY MET-3", BASECOAT),
    (13, "(GD8) BASE COAT BURNT COCONUT MET", BASECOAT),
    (14, "(GB0) BASE COAT BLACK MEET KETTLE MET-3", BASECOAT),
    (15, "(GYM) BASE COAT SATIN STEEL GRAY MET", BASECOAT),
    (16, "(GAN) BASE COAT SWITCHBLADE SILVER", BASECOAT),
    (17, "(GL8) BASE COAT RED-E OR NOT RED MET-3", BASECOAT),
    (18, "(GUM) BASE COAT SEEKER MET-3", BASECOAT),
    (19, "(G7G) BASE COAT ZEUS", BASECOAT),
    (20, "(GZ8) BASE COAT SOME KINDA BLUE", BASECOAT),
    (21, "(G5J) BASE COAT DARKMOON BLUE", BASECOAT),
    (22, "(G4V) BASE COAT POW ZINGA MET-3", BASECOAT),
    (23, "BASE TYPE THINNER", MIX_ROOM),
    (24, "TOP CLEAR HITEC 2000", PAINT_ROBOT),
    (25, "CLEAR THINNER", MIX_ROOM),
    (26, "CLEANING THINNER", PURGE),
    (27, "DISK BRAKE FLUID DOT-4 PLUS", SOMAC_CUBE),
    (28, "POWER STRG OIL DEXRON-VI GMN10060", SOMAC_CUBE),
    (29, "ANTIFREEZE DX-COOL CONCENTRATE RED", SOMAC_CUBE),
    (30, "T/M OIL Idemitsu MTF LSJ11EP", SOMAC_OIL),
    (31, "T/M OIL MANUAL SH XGP 75W85 GL-4", SOMAC_OIL),
    (32, "A/T OIL DEXRON-VI", SOMAC_OIL),
    (33, "TETRAFLUOROETHANE HFC-134a", ["Dürr Somac ProLine Cube", "Dürr Somac G4 Blue Twin A/C adapter"]),
    (34, "FUEL AI-92", ["Dürr Somac fuel filling station"]),
    (35, "WINDSHIELD WASHER (ETHANOL 80%)", ["Dürr Somac ProLine Cube"]),
    (36, "BETASEAL / SikaFlex-250 UZ-1", ["Graco Therm-O-Flow 200", "Atlas Copco SCA dispensing system"]),
    (37, "ADHESIVE-HDLNG T/PNL", ["Nordson ProBlue bulk hot melt melter"]),
    (38, "RKB 0101S (SIKA-209 D)", PRIMER_APPL),
    (39, "RKB 0401S (SIKA-206 G+P)", PRIMER_APPL),
    (40, "BETAPRIME (100-UZ)", ["Lint-free wipes (manual wipe-on / wipe-off)"]),
    (41, "RAD 11724 (SIKA-100-UZ)", ["Nordson EFD 752V diaphragm valve", "Graco agitated pressure tank"]),
    (42, "LUBRICANT-WHL PILOT HOLE", GREASE),
    (43, "LUBRICANT-CLU PLT HUB M/T (MICROLUBE GNY-202)", GREASE),
]

HEADERS = ["Material nomi", "Material spetsifikatsiyasi (pasporti)",
           "Material uchun jihoz", "Jihoz pasporti"]


def style(ws, widths: list[int]) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"


def main() -> None:
    catalog = {i["id"]: i for i in json.loads((ROOT / "data" / "catalog.json").read_text(encoding="utf-8"))}
    if [r[0] for r in ROWS] != sorted(catalog):
        raise SystemExit("ROWS ids do not match catalog ids")

    wb = Workbook()
    ws = wb.active
    ws.title = "Jadval"
    ws.append(HEADERS)
    for _, name, equipment in ROWS:
        ws.append([name, None, "\n".join(equipment), None])
    style(ws, [48, 30, 55, 30])

    notes = wb.create_sheet("Izoh")
    notes.append(["№", "Material nomi", "SAP", "Manbadagi spec", "Jihoz ishonchi"])
    order = ["low", "medium", "high"]
    for item_id, name, _ in ROWS:
        item = catalog[item_id]
        best = max((e.get("confidence", "low") for e in item["research"]["equipment"]), key=order.index)
        notes.append([item_id, name, item["sap"], item["spec"], best])
    style(notes, [5, 48, 40, 22, 14])

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"{len(ROWS)} rows -> {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
