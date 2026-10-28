# Material ↔ jihoz pasportlari jadvali

Avtomobil zavodi (UzAuto Motors, Asaka) materiallari ro'yxati asosida tuzilgan jadval:
material nomi → material spetsifikatsiyasi (PDF) → qo'llash jihozi → jihoz pasporti (PDF).

## Tuzilma
- `data/materials.json` — rasmlardan ko'chirilgan 43 ta material (SAP, nom, spec, qadoq, vazifa).
- `data/research/group_*.json` — har bir material bo'yicha tadqiqot: ishlab chiqaruvchi, hujjat URL'lari, jihozlar, ishonch darajasi.
- `data/catalog.json` — birlashtirilgan katalog, har bir hujjatga normallashtirilgan fayl nomi bilan.
- `output/materials_equipment.xlsx` — jadval (`Jadval`, `Manbalar`, `Identifikatsiya` varaqlari).
- `pdf/materials/`, `pdf/equipment/` — yuklab olingan PDF'lar.

## Fayl nomlash
- Material: `M{id}_{MATERIAL_NOMI}_{TDS|PDS|SDS|...}[_EQUIV].pdf` — `_EQUIV` = aynan shu mahsulot emas, ekvivalent hujjat.
- Jihoz: `EQ_{VENDOR}_{MODEL}_{datasheet|manual|brochure}.pdf`.
- Bitta URL — bitta fayl; bir nechta material bir hujjatni bo'lishishi mumkin. HTML sahifalar fayl sifatida saqlanmaydi.

## Ishga tushirish
```bash
python scripts/build_table.py data/research   # catalog.json + xlsx
python scripts/download_pdfs.py                # PDF'larni yuklash (tarmoq ruxsati kerak)
python scripts/build_table.py data/research   # yuklash holati bilan qayta yig'ish
```
