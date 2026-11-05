# SJR Ritual System

**Tizim · Jarayon · Natija** — Telegram Mini App va veb uchun shaxsiy rivojlanish platformasi.

- ☀️ **Bugun** — kundalik rituallar (bo'limlar, hafta kunlari jadvali, seriya, daraja, maskot)
- ✅ **Vazifalar** — kanban (papkalar, ustunlar, muddatlar, teglar, prioritet)
- 📝 **Eslatmalar** — daftarlar, avtosaqlash, qidiruv, PIN bilan himoyalangan maxfiy daftar
- 🏆 **Mukofotlar** — kunlik va uzoq muddatli (global) mukofotlar
- ⚡ **Energiya** — minnatdorlik, maqsadlar, afirmatsiyalar
- 📊 **Statistika** — issiqlik xaritasi, hafta/oy ko'rinishi, grafik, ritual reytingi
- 🤖 **AI tayyor** — `window.SJR` API, AI uchun so'rov/JSON eksport ([AI_INTEGRATION.md](AI_INTEGRATION.md))

Bitta fayl: `index.html`. Ma'lumotlar localStorage'da va (mavjud bo'lsa) Firebase Firestore'da sinxronlanadi.
PIN SHA-256 bilan xeshlanadi. Zaxira nusxa: Sozlamalar → Ma'lumotlar.

## Excel'dan import

Ilovada: **Sozlamalar → 📊 Excel'dan import** → `Hayot_Planneri.xlsx`. Batafsil: [AI_INTEGRATION.md](AI_INTEGRATION.md#4-excel-hayot-planneri-importi).

```bash
pip install openpyxl
python3 tools/import_hayot_planneri.py Hayot_Planneri.xlsx   # → data/hayot-planneri.json (gitignore)
```

## Testlar

```bash
npm i playwright          # yoki tayyor o'rnatilgan Playwright
python3 tools/import_hayot_planneri.py Hayot_Planneri.xlsx
python3 -m http.server 8765 &
XLSX=Hayot_Planneri.xlsx node tests/e2e.test.js     # 49 ta tekshiruv
```
