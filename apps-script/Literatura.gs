/**
 * Література · 5 клас — прийом результатів уроків-квестів (сайт vadym-kulyk.github.io/5literature).
 * Код для Apps Script таблиці «5 lit» (Розширення → Apps Script → Code.gs).
 * Аркуші: «Результати» (рядок на учня й урок) і «Відповіді» (письмові відповіді).
 * Один учень + один урок = один рядок, який оновлюється (ключ: урок + клас + ім'я).
 */
const LIT_SHEET = 'Результати';
const LIT_TEXT_SHEET = 'Відповіді';
const LIT_HEAD = ['Ключ', 'Оновлено', 'Клас', 'Учень', 'Пара', 'Урок', 'Бали, %', 'Станцій виконано', 'Деталі', 'Скарби', 'Загадки', 'Таємна кімната'];
const LIT_TEXT_HEAD = ['Ключ', 'Оновлено', 'Клас', 'Учень', 'Пара', 'Урок', 'Завдання', 'Рішення', 'Відповідь', 'Знаків'];

function doGet() {
  return ContentService.createTextOutput('Таблиця «5 lit» готова приймати результати.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d || !d.lesson || !d.name) return litOut_('no name');
    const now = new Date();
    const name = litSafe_(d.name).slice(0, 60);
    const cls = litSafe_(d.cls).slice(0, 10);
    const pair = litSafe_(d.partner).slice(0, 60);
    const lesson = litSafe_(d.title || d.lesson).slice(0, 120);
    const key = [d.lesson, cls, String(d.name).trim().toLowerCase()].join('|').slice(0, 200);
    const details = (d.tasks || []).slice(0, 40).map((t) => litSafe_(t.t) + ': ' + litNum_(t.p) + '/' + litNum_(t.max)).join('\n');
    const done = (d.tasks || []).filter((t) => t.p > 0 && t.p === t.max).length + '/' + (d.tasks || []).length;
    litUpsert_(litSheet_(LIT_SHEET, LIT_HEAD, true), [key, now, cls, name, pair, lesson, litNum_(d.score), done, details,
      litNum_(d.treasures) + '/7', litNum_(d.riddles) + '/5', d.secret ? 'так' : '']);
    (d.texts || []).slice(0, 10).forEach((x) => {
      const a = litSafe_(x.a).slice(0, 5000);
      if (!a) return;
      litUpsert_(litSheet_(LIT_TEXT_SHEET, LIT_TEXT_HEAD, false), [key + '#' + x.id, now, cls, name, pair, lesson, litSafe_(x.t), litSafe_(x.choice), a, a.length]);
    });
    return litOut_('ok');
  } finally {
    lock.releaseLock();
  }
}

function litOut_(t) { return ContentService.createTextOutput(t); }
function litNum_(v) { return typeof v === 'number' && isFinite(v) ? v : 0; }
// захист від формул у тексті учня
function litSafe_(v) { const s = String(v == null ? '' : v).trim(); return /^[=+\-@]/.test(s) ? "'" + s : s; }

function litSheet_(name, head, isMain) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#99cc00').setFontColor('#1e1b26').setWrap(true).setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 38);
    sh.hideColumns(1);
    sh.getRange('B:B').setNumberFormat('dd.MM.yyyy HH:mm');
    if (isMain) {
      sh.setFrozenColumns(4);
      sh.setColumnWidth(4, 160); sh.setColumnWidth(5, 140); sh.setColumnWidth(6, 230); sh.setColumnWidth(9, 320);
      sh.getRange('I:I').setWrap(true).setFontSize(9);
      sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue('#f4c7c3', SpreadsheetApp.InterpolationType.NUMBER, '0')
        .setGradientMidpointWithValue('#fce8b2', SpreadsheetApp.InterpolationType.NUMBER, '60')
        .setGradientMaxpointWithValue('#b7e1a1', SpreadsheetApp.InterpolationType.NUMBER, '100')
        .setRanges([sh.getRange('G2:G2000')]).build()]);
    } else {
      sh.setColumnWidth(7, 180); sh.setColumnWidth(9, 480);
      sh.getRange('I:I').setWrap(true);
    }
    const first = ss.getSheetByName('Аркуш1') || ss.getSheetByName('Sheet1');
    if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  }
  return sh;
}

function litUpsert_(sh, row) {
  const n = sh.getLastRow();
  if (n > 1) {
    const keys = sh.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === String(row[0])) { sh.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
    }
  }
  sh.appendRow(row);
}

/** Запусти один раз вручну (▶ Виконати), щоб створити аркуші й дати дозволи. */
function litSetup() {
  litSheet_(LIT_SHEET, LIT_HEAD, true);
  litSheet_(LIT_TEXT_SHEET, LIT_TEXT_HEAD, false);
}
