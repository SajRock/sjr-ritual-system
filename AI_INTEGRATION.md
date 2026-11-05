# SJR Ritual System — AI (Jarvis) integratsiyasi

Platforma sun'iy intellekt ulanishiga tayyorlangan. AI ma'lumotlarni o'qib, rejani tahlil qila oladi
va xavfsiz amallarni bajara oladi. Maxfiy eslatmalar va PIN **hech qachon** berilmaydi.
Ilova qulflangan bo'lsa, API `null` qaytaradi.

## 1. Eng oson yo'l (bugun ishlaydi)

**Sozlamalar → 🤖 Sun'iy intellekt**:
- **AI uchun so'rov nusxalash** — tahlil topshirig'i va barcha ma'lumotlaringiz (JSON) nusxalanadi.
  ChatGPT / Claude / Jarvis'ga joylashtirsangiz, u odatlar, vazifalar, maqsadlar va moliya bo'yicha maslahat beradi.
- **AI ma'lumotlari (JSON)** — xuddi shu ma'lumot fayl ko'rinishida.

## 2. JavaScript API — `window.SJR`

| Metod | Natija |
|---|---|
| `SJR.snapshot()` | To'liq holat (`schema: "sjr.snapshot/v1"`) |
| `SJR.habits({from, to})` | Odatlar: rejalashtirilgan/bajarilgan kunlar, `completionRate`, `doneDates` |
| `SJR.tasks()` | Vazifalar: `eisenhowerQuadrant` (1–4), `overdue`, `deadline`, `goal` |
| `SJR.finance()` | Oylar bo'yicha: daromad, xarajat, balans, `savingsRate`, `overLimit`, tranzaksiyalar |
| `SJR.aiPrompt()` | AI uchun tayyor matn |
| `SJR.exec(action)` | Xavfsiz amal (pastda) → `{ok, message}` |
| `SJR.on(fn)` | O'zgarishlarga obuna: `fn({keys:[...]})`. Obunani bekor qilish funksiyasini qaytaradi |

### `snapshot` tuzilishi (qisqacha)

```json
{
  "schema": "sjr.snapshot/v1",
  "today":  { "date": "2026-11-05", "completionPct": 45, "streakDays": 3, "rituals": [{"id","name","done"}] },
  "habits": [{ "id", "name", "section", "targetDaysPerMonth", "scheduledDays", "doneDays", "completionRate", "doneDates" }],
  "goals":  { "motivational": ["..."], "detailed": [{ "title", "area", "deadline", "daysLeft", "progress", "steps": [{"title","done"}] }] },
  "tasks":  { "total", "done", "overdue", "byQuadrant": [q1,q2,q3,q4], "items": [{ "title", "important", "urgent", "eisenhowerQuadrant", "deadline", "status" }] },
  "finance":{ "currency": "so'm", "months": [{ "month": "2026-07", "income", "expense", "balance", "savingsRate", "biggestExpense", "overLimit" }], "transactions": [] },
  "rewards":{ "daily": [], "global": [] },
  "notes":  { "count", "notebooks" }
}
```

### `SJR.exec` — ruxsat etilgan amallar (o'chirish amallari yo'q)

```js
SJR.exec({ type: 'ritual.done',  name: 'Meditatsiya' })            // yoki id, date
SJR.exec({ type: 'ritual.undo',  id: 'wake' })
SJR.exec({ type: 'task.add',     title: 'Hisobot', deadline: '2026-11-10', important: true, urgent: false })
SJR.exec({ type: 'task.done',    id: 'kc_...' })
SJR.exec({ type: 'goal.add',     title: 'IELTS 7.5' })
SJR.exec({ type: 'note.add',     title: 'Jarvis maslahati', text: '...' })
SJR.exec({ type: 'gratitude.add', text: 'Oilam uchun' })
SJR.exec({ type: 'finance.add',  kind: 'expense', amount: 45000, category: 'Kafe / restoran', note: 'Tushlik' })
SJR.exec({ type: 'navigate',     page: 'stats' })
SJR.exec({ type: 'notify',       text: 'Suv ichishni unutmang' })
```

Ro'yxatda yo'q har qanday amal `{ok:false}` bilan rad etiladi.

## 3. Ma'lumotlar saqlanishi (kalitlar)

Barcha ma'lumot `STORE` orqali localStorage + Firestore (`users/<id>` hujjati) da saqlanadi:

| Kalit | Mazmuni |
|---|---|
| `sjr_rituals`, `sjr_sections` | Rituallar (odatlar) va bo'limlar |
| `rd_YYYY-MM-DD` | Shu kuni bajarilgan ritual ID lari (JSON massiv) |
| `sjr_kcards`, `sjr_folders`, `sjr_cols_<folder>` | Vazifa kartochkalari, papkalar, ustunlar |
| `sjr_goals` | Maqsadlar ro'yxati (Bugun sahifasi, Energiya) |
| `sjr_finance` | `{currency, plans:{"YYYY-MM":{income,expense}}, transactions:[...]}` |
| `sjr_notes`, `sjr_notebooks` | Eslatmalar va daftarlar |
| `sjr_import_hayot` | Excel import reyestri (dublikatlarga qarshi) |

## 4. Excel (Hayot Planneri) importi

**Ilovada:** Sozlamalar → Ma'lumotlar → **📊 Excel'dan import** → `Hayot_Planneri.xlsx` faylini tanlang.
Ilova faylni o'zi o'qiydi (tashqi kutubxonasiz). Konvertor yaratgan `.json` ham qabul qilinadi.

| Excel | Platformada |
|---|---|
| Odatlar (nom, oylik maqsad, kunlik ✔) | Rituallar (mavjud o'xshashiga bog'lanadi) + ritual tarixi |
| Vazifalar (muhim/shoshilinch, muddat, holat) | "Umumiy" papkadagi kanban kartochkalari |
| Maqsadlar (soha, muddat, 5 qadam) | Maqsadlar ro'yxati + "🎯 Maqsadlar" papkasidagi kartochkalar |
| Moliya (reja, jurnal) | "💰 Moliya" daftaridagi oylik hisobot + `sjr_finance` (AI uchun) |

**Buyruq qatorida (AI/tahlil uchun JSON):**

```bash
python3 tools/import_hayot_planneri.py Hayot_Planneri.xlsx          # → data/hayot-planneri.json
python3 tools/import_hayot_planneri.py Hayot_Planneri.xlsx --embed  # faqat shaxsiy hosting!
```

Dublikatlarga qarshi qoidalar:
- har element barqaror ID oladi (`xl_<key>`; Python va brauzer bir xil kalit hisoblaydi);
- qayta import (xlsx yoki json) hech narsani ikki marta qo'shmaydi;
- import qilingan element o'chirilsa, keyingi importda qayta paydo bo'lmaydi;
- kunlik ✔ belgilari har odat uchun faqat bir marta yoziladi (keyingi o'zgarishlaringiz ustiga yozilmaydi).

## 5. Keyingi qadam: Jarvis serveri

Tavsiya etiladigan arxitektura: Jarvis (server) ilovadan `SJR.snapshot()` ni oladi, LLM ga yuboradi,
javobdagi amallarni `SJR.exec()` orqali bajaradi. API kalitlari faqat serverda saqlanishi kerak —
brauzer koduga hech qachon yozilmasin.
