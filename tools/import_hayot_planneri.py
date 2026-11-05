#!/usr/bin/env python3
"""
"Hayot Planneri" Excel faylini SJR Ritual System uchun tartibli JSON ga aylantiradi.

Foydalanish:
    python3 tools/import_hayot_planneri.py <Hayot_Planneri.xlsx> [--embed]

Natija:
  data/hayot-planneri.json — tartiblangan ma'lumot. Uni ilovada
  Sozlamalar → "Excel'dan import" orqali yuklash mumkin (ilova .xlsx ni ham o'zi o'qiydi).

  --embed  ma'lumotni index.html ichiga joylaydi va ilova ochilganda avtomatik import qiladi.
           DIQQAT: faqat shaxsiy (yopiq) hosting uchun! Ochiq repoda yoki ko'p foydalanuvchili
           platformada ishlatmang — ma'lumot hammaga ko'rinadi va har bir foydalanuvchiga import bo'ladi.

Har bir element barqaror "key" oladi (nomdan yasaladi), shuning uchun skript
yoki import qayta ishga tushirilsa ham hech narsa ikki marta qo'shilmaydi.
"""
import datetime as dt
import hashlib
import json
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
MONTHS = ['Yanvar', 'Fevral', 'Mart', 'Aprel', 'May', 'Iyun', 'Iyul', 'Avgust',
          'Sentabr', 'Oktabr', 'Noyabr', 'Dekabr']
STATUS = {'Rejada': 'todo', 'Jarayonda': 'doing', 'Bajarildi': 'done', 'Bekor': 'cancelled'}


def norm(s):
    return re.sub(r'\s+', ' ', str(s or '').replace('ʻ', "'").replace('ʼ', "'")
                  .replace('‘', "'").replace('’', "'").replace('`', "'")).strip()


def key_of(*parts):
    """Barqaror, o'qiladigan kalit: slug + qisqa xesh."""
    text = ' '.join(norm(p).lower() for p in parts)
    slug = re.sub(r'[^a-z0-9]+', '-', text).strip('-')[:40] or 'item'
    h = hashlib.sha1(text.encode('utf-8')).hexdigest()[:6]
    return f'{slug}-{h}'


def iso(d):
    if isinstance(d, (dt.datetime, dt.date)):
        return d.strftime('%Y-%m-%d')
    return None


def yes(v):
    return norm(v).lower() in ('ha', 'yes', 'true', '1')


def col_values(ws, col, start, end):
    return [norm(ws[f'{col}{r}'].value) for r in range(start, end + 1) if norm(ws[f'{col}{r}'].value)]


def convert(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    s = wb['Sozlamalar']
    lists = {
        'taskCategories': col_values(s, 'D', 4, 20),
        'statuses': col_values(s, 'E', 4, 20),
        'goalAreas': col_values(s, 'F', 4, 20),
        'incomeCategories': col_values(s, 'G', 4, 20),
        'expenseCategories': col_values(s, 'H', 4, 20),
    }

    # ── ODATLAR ──
    ws = wb['Odatlar']
    month_name, year = norm(ws['C5'].value), int(ws['C6'].value)
    month = MONTHS.index(month_name) + 1 if month_name in MONTHS else 1
    ym = f'{year}-{month:02d}'
    days_in_month = (dt.date(year + (month == 12), month % 12 + 1, 1) - dt.timedelta(days=1)).day
    habits = []
    for r in range(10, 25):
        name = norm(ws.cell(r, 3).value)
        if not name:
            continue
        done = [f'{ym}-{d:02d}' for d in range(1, days_in_month + 1)
                if norm(ws.cell(r, 4 + d).value) in ('✔', '✓', 'v', 'x', '1')]
        target = ws.cell(r, 4).value
        habits.append({'key': key_of('habit', name), 'name': name,
                       'targetDaysPerMonth': int(target) if isinstance(target, (int, float)) else None,
                       'doneDates': done})

    # ── VAZIFALAR ──
    ws = wb['Vazifalar']
    tasks = []
    for r in range(12, 72):
        title = norm(ws.cell(r, 3).value)
        if not title:
            continue
        imp, urg = yes(ws.cell(r, 5).value), yes(ws.cell(r, 6).value)
        tasks.append({
            'key': key_of('task', title), 'title': title,
            'category': norm(ws.cell(r, 4).value) or None,
            'important': imp, 'urgent': urg,
            'quadrant': 1 if imp and urg else 2 if imp else 3 if urg else 4,
            'deadline': iso(ws.cell(r, 8).value),
            'status': STATUS.get(norm(ws.cell(r, 9).value), 'todo'),
        })

    # ── MAQSADLAR ──
    ws = wb['Maqsadlar']
    goals = []
    for r in range(9, 24):
        title = norm(ws.cell(r, 3).value)
        if not title:
            continue
        steps = []
        for c in (6, 8, 10, 12, 14):  # F,H,J,L,N = qadam; keyingi ustun = ✔
            st = norm(ws.cell(r, c).value)
            if st:
                steps.append({'title': st, 'done': norm(ws.cell(r, c + 1).value) in ('✔', '✓')})
        goals.append({'key': key_of('goal', title), 'title': title,
                      'area': norm(ws.cell(r, 4).value) or None,
                      'deadline': iso(ws.cell(r, 5).value), 'steps': steps})

    # ── MOLIYA ──
    ws = wb['Moliya']
    fin_month_name, fin_year = norm(ws['C5'].value), int(ws['C6'].value)
    fin_ym = f'{fin_year}-{MONTHS.index(fin_month_name) + 1:02d}' if fin_month_name in MONTHS else ym
    # Reja satrlari Sozlamalar ro'yxati tartibida (B12:B16 = G4:G8, B20:B31 = H4:H15)
    plan_in = {cat: ws.cell(12 + i, 3).value for i, cat in enumerate(lists['incomeCategories'][:5])}
    plan_out = {cat: ws.cell(20 + i, 3).value for i, cat in enumerate(lists['expenseCategories'][:12])}
    plan_in = {k: int(v) for k, v in plan_in.items() if isinstance(v, (int, float)) and v > 0}
    plan_out = {k: int(v) for k, v in plan_out.items() if isinstance(v, (int, float)) and v > 0}
    txs = []
    for r in range(12, 212):
        date, kind, cat, amount, note = (ws.cell(r, c).value for c in range(7, 12))
        if not isinstance(amount, (int, float)) or not iso(date):
            continue
        typ = 'income' if norm(kind) == 'Daromad' else 'expense'
        txs.append({'key': key_of('tx', iso(date), typ, cat, amount, note), 'date': iso(date),
                    'type': typ, 'category': norm(cat), 'amount': int(amount), 'note': norm(note) or None})
    currency = norm(ws['C7'].value) or "so'm"

    return {
        'schema': 'sjr.import.hayot-planneri/v1',
        'source': Path(path).name,
        'convertedAt': dt.date.today().isoformat(),
        'habitsMonth': ym,
        'lists': lists,
        'habits': habits,
        'tasks': tasks,
        'goals': goals,
        'finance': {'currency': currency, 'month': fin_ym,
                    'plan': {'income': plan_in, 'expense': plan_out},
                    'transactions': sorted(txs, key=lambda t: t['date'])},
    }


def main():
    if len([a for a in sys.argv[1:] if not a.startswith('--')]) < 1:
        print(__doc__)
        sys.exit(1)
    data = convert([a for a in sys.argv[1:] if not a.startswith('--')][0])
    out = ROOT / 'data' / 'hayot-planneri.json'
    out.parent.mkdir(exist_ok=True)
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')

    if '--embed' not in sys.argv:
        print(f"OK: {len(data['habits'])} odat, {len(data['tasks'])} vazifa, {len(data['goals'])} maqsad, "
              f"{len(data['finance']['transactions'])} moliya yozuvi → {out.relative_to(ROOT)}")
        return
    html_path = ROOT / 'index.html'
    html = html_path.read_text(encoding='utf-8')
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    block = f'<script type="application/json" id="sjr-import-hayot">{payload}</script>'
    pattern = re.compile(r'<script type="application/json" id="sjr-import-hayot">.*?</script>', re.S)
    if pattern.search(html):
        html = pattern.sub(lambda _: block, html)
    else:
        html = html.replace('</body>', block + '\n</body>', 1)
    html_path.write_text(html, encoding='utf-8')
    print(f"OK: {len(data['habits'])} odat, {len(data['tasks'])} vazifa, {len(data['goals'])} maqsad, "
          f"{len(data['finance']['transactions'])} moliya yozuvi → {out.relative_to(ROOT)} + index.html")


if __name__ == '__main__':
    main()
