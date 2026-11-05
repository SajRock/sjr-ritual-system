/**
 * SJR Ritual System — end-to-end testlar (Playwright, tashqi kutubxonalarsiz).
 *
 * Ishga tushirish:
 *   python3 tools/import_hayot_planneri.py Hayot_Planneri.xlsx   # data/hayot-planneri.json
 *   python3 -m http.server 8765 &      # repo ildizida
 *   XLSX=/path/Hayot_Planneri.xlsx node tests/e2e.test.js
 *
 * Tashqi CDN so'rovlari bloklanadi (ilova lokal rejimda ishlashi kerak).
 * Agar CHART_JS muhit o'zgaruvchisi chart.umd.js fayliga ishora qilsa, grafik ham tekshiriladi.
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const BASE = process.env.BASE_URL || 'http://localhost:8765/index.html';
const SHOTS = process.env.SHOTS_DIR || null;
const CHART = process.env.CHART_JS && fs.existsSync(process.env.CHART_JS) ? fs.readFileSync(process.env.CHART_JS, 'utf8') : null;
const XLSX = process.env.XLSX;
const JSON_PATH = path.join(__dirname, '..', 'data', 'hayot-planneri.json');
if (!XLSX || !fs.existsSync(XLSX) || !fs.existsSync(JSON_PATH)) {
  console.error('XLSX muhit o\'zgaruvchisi (Excel fayl yo\'li) va data/hayot-planneri.json kerak — README ga qarang');
  process.exit(2);
}
const DATA = JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));
async function importFile(page, file) {
  await page.evaluate(() => showPage('settings'));
  const [fc] = await Promise.all([page.waitForEvent('filechooser'), page.click('text=Fayl tanlash')]);
  await fc.setFiles(file);
  await page.waitForFunction(() => !document.getElementById('notifBanner').classList.contains('hidden') && /Import|Yangi narsa/.test(document.getElementById('notifText').textContent), null, { timeout: 10000 });
  const msg = await page.textContent('#notifText');
  await page.evaluate(() => closeNotif());
  return msg;
}

let passed = 0, failed = 0;
function check(name, cond, info) {
  if (cond) { passed++; console.log('  ✓ ' + name); }
  else { failed++; console.log('  ✗ ' + name + (info !== undefined ? '  → ' + JSON.stringify(info) : '')); }
}

async function newPage(browser, opts = {}) {
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 860 }, acceptDownloads: true, ...opts });
  await ctx.route('**/*', route => {
    const u = route.request().url();
    if (u.startsWith('http://localhost')) return route.continue();
    if (CHART && /chart/i.test(u)) return route.fulfill({ status: 200, contentType: 'application/javascript', body: CHART });
    return route.abort();
  });
  const page = await ctx.newPage();
  page.errors = [];
  page.on('pageerror', e => page.errors.push(e.message));
  page.on('dialog', d => d.accept(d.type() === 'prompt' ? 'Test' : undefined));
  return page;
}
async function setupUser(page, pin = '2580', { mobile = false } = {}) {
  await page.goto(BASE);
  await page.evaluate(() => localStorage.clear());
  await page.reload();
  await page.waitForSelector('#pinScreen', { state: 'visible', timeout: 15000 });
  for (let round = 0; round < 2; round++) {
    for (const k of pin) {
      if (mobile) await page.click(`.pin-key:text-is("${k}")`); else await page.keyboard.press(k);
      await page.waitForTimeout(60);
    }
    await page.waitForTimeout(350);
  }
  await page.waitForSelector('#onboardScreen', { state: 'visible' });
  await page.click('.ob-chip >> nth=0');
  await page.click('.ob-btn');
  await page.waitForSelector('#mainApp', { state: 'visible' });
  await page.waitForTimeout(400);
}
const pageHtml = (page, name, mobile) => page.innerHTML(`#${mobile ? 'm' : 'd'}page-${name}`);

(async () => {
  const browser = await chromium.launch();

  console.log('\n1) Excel importi — desktop (ilova .xlsx ni o\'zi o\'qiydi)');
  const p = await newPage(browser);
  await setupUser(p);
  check('import oldin hech narsa avtomatik qo\'shilmagan', await p.evaluate(() => !getKanbanCards().some(c => c.source === 'excel')));
  const msg1 = await importFile(p, XLSX);
  check('xlsx import xabari', /Import tugadi/.test(msg1), msg1);
  const s1 = await p.evaluate(() => ({
    rituals: getRituals(), cards: getKanbanCards(), goals: getGoals(), folders: getFolders(),
    notes: getNotes(), notebooks: getNotebooks(), fin: JSON.parse(STORE.getItem('sjr_finance') || '{}'),
    reg: JSON.parse(STORE.getItem('sjr_import_hayot') || '{}'), first: firstUseDate(),
  }));
  const taskTitles = DATA.tasks.map(t => t.title);
  check('barcha 12 vazifa kartochka bo\'ldi', taskTitles.every(t => s1.cards.some(c => c.title === t)), taskTitles.filter(t => !s1.cards.some(c => c.title === t)));
  check('vazifalar "Umumiy" papkada', s1.cards.filter(c => c.source === 'excel' && !c.goal).every(c => c.folderId === 'folder_default'));
  const byTitle = Object.fromEntries(s1.cards.map(c => [c.title, c]));
  check('"Bajarildi" → done ustuni', byTitle['Hisobotni topshirish'].done && byTitle['Hisobotni topshirish'].colId === 'col_done');
  check('"Jarayonda" → col_doing', byTitle['Mijoz bilan uchrashuv'].colId === 'col_doing' && !byTitle['Mijoz bilan uchrashuv'].done);
  check('"Rejada" → col_todo', byTitle['IELTS uchun mock test'].colId === 'col_todo');
  check('"Bekor" → bajarilganlar + "Bekor" tegi', byTitle["Serial ko'rish"].done && byTitle["Serial ko'rish"].labels.some(l => l.name === 'Bekor'));
  check('muhim+shoshilinch → high + Shoshilinch tegi', byTitle['Hisobotni topshirish'].priority === 'high' && byTitle['Hisobotni topshirish'].labels.some(l => l.name === 'Shoshilinch'));
  check('muddat saqlandi', byTitle["Kommunal to'lovlar"].deadline === '2026-07-14');
  check('10 odat ritualga aylandi/bog\'landi', DATA.habits.every(h => s1.reg.items['xl_' + h.key]));
  check('6 tasi mavjud ritualga bog\'landi (dublikat yo\'q)', ['wake', 'sport', 'water', 'book', 'english', 'diet'].every(id => s1.rituals.find(r => r.id === id)?.source));
  check('4 ta yangi ritual qo\'shildi', ['Meditatsiya', 'Kundalik yozish', '10 000 qadam', 'Ijtimoiy tarmoq < 1 soat'].every(n => s1.rituals.filter(r => r.name === n).length === 1));
  check('ritual nomlari takrorlanmaydi', new Set(s1.rituals.map(r => r.name)).size === s1.rituals.length);
  check('oylik maqsad (25 kun) saqlandi', s1.rituals.find(r => r.name === 'Meditatsiya').target === 25);
  check('foydalanish boshlanishi 2026-07-01 ga surildi', s1.first === '2026-07-01', s1.first);
  const marks = await p.evaluate(() => ({ a: isDone('2026-07-01', 'wake'), b: isDone('2026-07-18', 'wake'), c: isDone('2026-07-06', getRituals().find(r => r.name === 'Meditatsiya').id) }));
  check('iyul belgilari ritual tarixiga yozildi', marks.a && !marks.b && marks.c, marks);
  check('tugallanmagan 3 maqsad maqsadlar ro\'yxatida', ['Ingliz tili IELTS 7.0', "Favqulodda jamg'arma 20 mln", 'Haftada 1 kitob'].every(g => s1.goals.includes(g)));
  check('100% bajarilgan maqsad ro\'yxatga qo\'shilmadi', !s1.goals.includes('-8 kg vazn'));
  check('"Maqsadlar" papkasi va 4 ta maqsad kartochkasi', s1.folders.some(f => f.id === 'xl_folder_goals') && s1.cards.filter(c => c.folderId === 'xl_folder_goals').length === 4);
  const ielts = s1.cards.find(c => c.title === 'Ingliz tili IELTS 7.0');
  check('maqsad qadamlari va progress (60%)', ielts.goal.progress === 60 && ielts.desc.includes('⬜ Speaking klub') && ielts.colId === 'col_doing', ielts.goal);
  check('17 moliya yozuvi', (s1.fin.transactions || []).length === 17);
  check('moliya rejasi 2026-07', s1.fin.plans['2026-07'].expense['Uy-joy / ijara'] === 3000000);
  const finNote = s1.notes.find(n => n.id === 'xl_fin_2026-07');
  check('"Moliya" daftarida hisobot eslatmasi', !!finNote && s1.notebooks.some(n => n.id === 'xl_nb_finance') && finNote.notebookId === 'xl_nb_finance');
  check('hisobot raqamlari to\'g\'ri (daromad 10.5 mln, xarajat 8.53 mln)', finNote.text.includes('Daromad (fakt): 10 500 000') && finNote.text.includes('Xarajat (fakt): 8 530 000'), finNote.text.slice(0, 200));

  console.log('\n2) Dublikatsizlik va o\'chirilganlarni hurmat qilish');
  const before = await p.evaluate(() => ({ c: getKanbanCards().length, r: getRituals().length, g: getGoals().length, t: JSON.parse(STORE.getItem('sjr_finance')).transactions.length, n: getNotes().length }));
  const msgJson = await importFile(p, JSON_PATH); // Python konvertori natijasi — xuddi shu kalitlar bo'lishi kerak
  check('Python JSON importi dublikat yaratmadi (kalitlar mos)', /Yangi narsa yo/.test(msgJson), msgJson);
  const msgAgain = await importFile(p, XLSX);
  const rerun = { total: /Yangi narsa yo/.test(msgAgain) ? 0 : 1, msgAgain };
  const after = await p.evaluate(() => ({ c: getKanbanCards().length, r: getRituals().length, g: getGoals().length, t: JSON.parse(STORE.getItem('sjr_finance')).transactions.length, n: getNotes().length }));
  check('qayta import hech narsa qo\'shmadi', rerun.total === 0 && JSON.stringify(before) === JSON.stringify(after), { rerun, before, after });
  await p.evaluate(() => { saveKanbanCards(getKanbanCards().filter(c => c.title !== "Serial ko'rish")); const r = getRituals().filter(x => x.name !== 'Meditatsiya'); saveRituals(r); });
  await importFile(p, XLSX);
  const del = await p.evaluate(() => ({ card: getKanbanCards().some(c => c.title === "Serial ko'rish"), rit: getRituals().some(r => r.name === 'Meditatsiya') }));
  check('foydalanuvchi o\'chirgan narsa qayta qo\'shilmaydi', !del.card && !del.rit, del);
  await p.reload();
  await p.waitForSelector('#pinScreen', { state: 'visible' });
  for (const k of '2580') { await p.keyboard.press(k); await p.waitForTimeout(50); }
  await p.waitForSelector('#mainApp', { state: 'visible' }); await p.waitForTimeout(400);
  const afterReload = await p.evaluate(() => getKanbanCards().length);
  check('sahifa qayta yuklanganda ham dublikat yo\'q', afterReload === after.c - 1, { afterReload });

  console.log('\n3) Ko\'rinish — mavjud sahifalarda');
  await p.evaluate(() => showPage('tasks'));
  let html = await pageHtml(p, 'tasks');
  check('Vazifalar sahifasida import kartochkalari', html.includes('Hisobotni topshirish') && html.includes('Shoshilinch'));
  if (SHOTS) await p.screenshot({ path: SHOTS + '/tasks.png' });
  await p.click('#dpage-tasks .folder-chip:has-text("Maqsadlar")');
  html = await pageHtml(p, 'tasks');
  check('Maqsadlar papkasida qadamlar ko\'rinadi', html.includes('Speaking klub') && html.includes('Progress: 60%'));
  if (SHOTS) await p.screenshot({ path: SHOTS + '/goals-folder.png' });
  await p.evaluate(() => { showPage('notes'); _activeNotebookId = 'xl_nb_finance'; renderNotes(); });
  html = await pageHtml(p, 'notes');
  check('Eslatmalarda "Moliya — Iyul 2026"', html.includes('Moliya — Iyul 2026'));
  await p.click('.note-card');
  const body = await p.inputValue('#noteEditorBody');
  check('moliya eslatmasi ochiladi', body.includes("Jamg'arma ulushi") && body.includes('Korzinka'));
  if (SHOTS) await p.screenshot({ path: SHOTS + '/finance-note.png' });
  await p.evaluate(() => { _statsSubTab = 'month'; currentMonth = 6; currentYear = 2026; showPage('stats'); });
  html = await pageHtml(p, 'stats');
  check('Statistika: iyul oyi foizlar bilan', /\d+%/.test(html) && html.includes('Iyul 2026'));
  if (SHOTS) await p.screenshot({ path: SHOTS + '/stats-july.png' });
  await p.evaluate(() => { _statsSubTab = 'stats'; renderStats(); });
  html = await pageHtml(p, 'stats');
  check('Statistika: oylik maqsad ko\'rsatiladi', html.includes('oylik maqsad'));
  await p.evaluate(() => showPage('today'));
  html = await pageHtml(p, 'today');
  check('Bugun sahifasida maqsad chiplari va yangi rituallar', html.includes('Ingliz tili IELTS 7.0') || html.includes('Intizomli'));
  check('Bugun: "Kundalik yozish" ritual ko\'rinadi', html.includes('Kundalik yozish'));
  await p.evaluate(() => showPage('settings'));
  html = await pageHtml(p, 'settings');
  check('Sozlamalar: import holati va AI tugmalari', html.includes('Hayot planneri') && html.includes('AI uchun so'));

  console.log('\n4) AI API (window.SJR)');
  const api = await p.evaluate(() => {
    const s = SJR.snapshot();
    return { schema: s.schema, habits: s.habits.length, tasks: s.tasks.total, q: s.tasks.byQuadrant, fin: s.finance.months[0],
      goals: s.goals.detailed.length, secret: JSON.stringify(s).includes('sjr_snotes'), prompt: SJR.aiPrompt().length,
      frozen: Object.isFrozen(SJR) };
  });
  check('snapshot sxemasi', api.schema === 'sjr.snapshot/v1' && api.frozen);
  check('snapshotda odat/vazifa/maqsad', api.habits >= 16 && api.tasks >= 15 && api.goals === 4, api);
  check('moliya tahlili (jamg\'arma ulushi 28%, ijara eng katta)', api.fin.savingsRate === 28 && api.fin.biggestExpense.category === 'Uy-joy / ijara', api.fin);
  check('maxfiy ma\'lumot yo\'q', !api.secret);
  const ex = await p.evaluate(() => [SJR.exec({ type: 'ritual.done', name: 'Kundalik yozish' }), SJR.exec({ type: 'task.add', title: 'AI test vazifa', important: true, urgent: false }), SJR.exec({ type: 'rm -rf' })]);
  check('SJR.exec whitelist ishlaydi', ex[0].ok && ex[1].ok && !ex[2].ok, ex);
  check('ritual.done haqiqatan belgiladi', await p.evaluate(() => isDone(todayStr, getRituals().find(r => r.name === 'Kundalik yozish').id)));
  await p.evaluate(() => lockApp('t'));
  check('qulflanganda API ma\'lumot bermaydi', await p.evaluate(() => SJR.snapshot() === null && SJR.exec({ type: 'notify', text: 'x' }).ok === false));
  check('desktop: JS xatolari yo\'q', p.errors.length === 0, p.errors);

  console.log('\n5) Mobil');
  const m = await newPage(browser, { viewport: { width: 390, height: 844 }, isMobile: true, hasTouch: true });
  await setupUser(m, '1357', { mobile: true });
  await importFile(m, XLSX);
  await m.click('#mtab-today'); await m.waitForTimeout(250);
  const mt = await pageHtml(m, 'today', true);
  check('mobil: Bugun sahifasi ko\'rinadi', mt.includes('ritual-row') && (await m.isVisible('#mpage-today .ritual-row >> nth=0')));
  await m.click('#mtab-tasks'); await m.waitForTimeout(200);
  check('mobil: import kartochkalari', (await pageHtml(m, 'tasks', true)).includes('Mijoz bilan uchrashuv'));
  if (SHOTS) await m.screenshot({ path: SHOTS + '/m-tasks.png' });
  await m.click('#mtab-today'); await m.waitForTimeout(200);
  if (SHOTS) await m.screenshot({ path: SHOTS + '/m-today.png', fullPage: true });
  check('mobil: JS xatolari yo\'q', m.errors.length === 0, m.errors);

  await browser.close();
  console.log(`\nNatija: ${passed} o'tdi, ${failed} xato`);
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
