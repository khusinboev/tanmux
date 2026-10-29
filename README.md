# Material ↔ jihoz pasportlari jadvali

Avtomobil zavodi (UzAuto Motors, Asaka) materiallari ro'yxati asosida tuzilgan jadval:
material nomi → material spetsifikatsiyasi (PDF) → qo'llash jihozi → jihoz pasporti (PDF).

## Natija fayllari
- `output/materials_equipment.xlsx` — **1-versiya**: pasport ustunlarida normallashtirilgan PDF fayl nomlari.
- `output/materials_equipment_links.xlsx` — **2-versiya**: pasport ustunlarida to'g'ridan-to'g'ri PDF havolalari (har bir havola alohida katakda, bosiladigan).
- `output/material_equipment_names.xlsx` — faqat aniq material va jihoz nomlari (pasport ustunlari bo'sh), `Izoh` varag'ida zavod ro'yxatidagi nom/SAP/spec.
- 1- va 2-versiyada `Jadval`, `Manbalar` (barcha manbalar, holati bilan), `Identifikatsiya` varaqlari bor.
- Asosiy `Jadval` varag'iga **faqat yuklangan va mazmuni tekshirilgan PDF'lar** tushadi; HTML sahifalar va tekshiruvdan o'tmaganlar faqat `Manbalar`da.

## Tuzilma
- `data/materials.json` — rasmlardan ko'chirilgan 43 ta material (SAP, nom, spec, qadoq, vazifa).
- `data/research/group_*.json` — har bir material bo'yicha tadqiqot: ishlab chiqaruvchi, hujjatlar, jihozlar, ishonch darajasi.
- `data/research/rounds/` — tadqiqot bosqichlari (round2–round5) va reviewer tuzatishlari, qayta qo'llash mumkin.
- `data/equipment_selection.json` — har bir material uchun tanlangan jihozlar (rasmiy nomlar) va ularning pasport PDF'lari.
- `data/verification/` — PDF mazmuni tekshiruvi natijalari (har bir fayl: haqiqiy hujjat, mos/mos emas, harakat).
- `data/catalog.json`, `data/download_report.json`, `data/verification.json`, `data/verification_log.json` — yig'ilgan katalog, yuklash va tekshiruv holati.
- `pdf/materials/`, `pdf/equipment/` — yuklab olingan PDF'lar (**git'ga qo'shilmaydi**: repo public, hujjatlar vendorlarniki).

## Fayl nomlash
- Material: `M<id>_<TDS|PDS|SDS|...>[_EQUIV]_<haqiqiy hujjat sarlavhasi>.pdf` — `_EQUIV` = aynan shu mahsulot emas, shu sinfdagi ekvivalent hujjat.
- Jihoz: `EQ_<datasheet|manual|brochure>_<haqiqiy hujjat sarlavhasi>.pdf`.
- Bitta URL — bitta fayl; bir nechta material bir hujjatni bo'lishishi mumkin.

## Ishga tushirish
```bash
python scripts/apply_verification.py data/verification/verify_*.json   # tekshiruv xulosalarini qo'llash (bir marta)
python scripts/merge_round.py data/research/rounds/<round>.json        # yangi tadqiqot bosqichini qo'shish
python scripts/build_table.py data/research                            # catalog.json + 1/2-versiya xlsx
python scripts/download_pdfs.py                                        # PDF'larni yuklash (%PDF tekshiruvi bilan)
python scripts/verify_pdfs.py <excerpt_dir>                            # PDF matnini ajratib, kalit so'zlarni solishtirish
python scripts/build_table.py data/research                            # yuklash holati bilan qayta yig'ish
python scripts/build_names_sheet.py                                    # nomlar jadvali
```
