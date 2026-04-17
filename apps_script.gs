/**
 * ══════════════════════════════════════════════
 *  ТАБЕЛЬ 2026 — Google Apps Script
 * ══════════════════════════════════════════════
 *
 *  ІНСТРУКЦІЯ:
 *  1. Вставте цей код у Apps Script
 *  2. Збережіть (Ctrl+S)
 *  3. Розгорнути → Нове розгортання
 *     • Тип: Веб-застосунок
 *     • Виконувати від імені: Я
 *     • Доступ: Всі (анонімні)
 *  4. Скопіюйте URL → вставте у форму ⚙ Налаштування
 *
 *  ДЛЯ ЩОДЕННОГО ЗВЕДЕННЯ О 20:00:
 *  Triggers → Додати тригер → sendDailySummary
 *  → Time-driven → Day timer → 8pm-9pm
 *
 *  ВАЖЛИВО: після кожної зміни коду — НОВЕ розгортання
 * ══════════════════════════════════════════════
 */

// ─── Telegram конфігурація ───────────────────────────────────────────────────
const TG_TOKEN   = '8303053237:AAHqd_SSIjM-3l08EKRbyJF7MFAHXPDDPAY';
const TG_CHAT_ID = '222538505';
const TAB_SHEET_DEFAULT = 'ТАБЕЛЬ';


// ─── doGet — завантаження списків ───────────────────────────────────────────
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const action = (e.parameter && e.parameter.action) || '';

    if (action === 'ping' || action === '') {
      output.setContent(JSON.stringify({ ok: true, message: 'Табель API працює' }));
      return output;
    }

    if (action === 'getData') {
      const empSheet = (e.parameter.empSheet || 'СПРАВОЧНИК').toString();
      const empCol   = (e.parameter.empCol   || 'A').toString().toUpperCase();
      const objSheet = (e.parameter.objSheet || 'СПРАВОЧНИК').toString();
      const objCol   = (e.parameter.objCol   || 'C').toString().toUpperCase();
      const secSheet = (e.parameter.secSheet || 'СПРАВОЧНИК').toString();
      const secCol   = (e.parameter.secCol   || 'B').toString().toUpperCase();
      const wtSheet  = (e.parameter.wtSheet  || 'СПРАВОЧНИК').toString();
      const wtCol    = (e.parameter.wtCol    || 'D').toString().toUpperCase();

      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // ── Співробітники ──
      const esht = ss.getSheetByName(empSheet);
      let employees = [];
      if (esht) {
        const lastRow = esht.getLastRow();
        if (lastRow >= 2) {
          employees = esht.getRange(empCol + '2:' + empCol + lastRow).getValues()
            .flat().map(v => v.toString().trim()).filter(v => v.length > 0);
        }
      }

      // ── Об'єкти ──
      const osht = ss.getSheetByName(objSheet);
      let objects = [];
      if (osht) {
        const lastRow = osht.getLastRow();
        if (lastRow >= 2) {
          objects = osht.getRange(objCol + '2:' + objCol + lastRow).getValues()
            .flat().map(v => v.toString().replace(/\n/g, ' ').trim()).filter(v => v.length > 0)
            .map(raw => {
              const sepIdx = raw.indexOf(' | ');
              return sepIdx > 0
                ? { code: raw.slice(0, sepIdx).trim(), name: raw.slice(sepIdx + 3).trim() }
                : { code: raw, name: '' };
            });
        }
      }

      // ── Розділи ──
      const ssht = ss.getSheetByName(secSheet);
      let sections = [];
      if (ssht) {
        const lastRow = ssht.getLastRow();
        if (lastRow >= 2) {
          sections = ssht.getRange(secCol + '2:' + secCol + lastRow).getValues()
            .flat().map(v => v.toString().trim()).filter(v => v.length > 0);
        }
      }

      // ── Види робіт ──
      const wtsht = ss.getSheetByName(wtSheet);
      let worktypes = [];
      if (wtsht) {
        const lastRow = wtsht.getLastRow();
        if (lastRow >= 2) {
          worktypes = wtsht.getRange(wtCol + '2:' + wtCol + lastRow).getValues()
            .flat().map(v => v.toString().trim()).filter(v => v.length > 0);
        }
      }

      output.setContent(JSON.stringify({
        ok: true,
        employees, objects, sections, worktypes,
        meta: {
          empSheet, empCol, empCount: employees.length,
          objSheet, objCol, objCount: objects.length,
          secSheet, secCol, secCount: sections.length,
          wtSheet, wtCol, wtCount: worktypes.length,
          timestamp: new Date().toISOString()
        }
      }));
      return output;
    }

    output.setContent(JSON.stringify({ ok: false, error: 'Невідома дія: ' + action }));
    return output;

  } catch (err) {
    output.setContent(JSON.stringify({ ok: false, error: err.message, stack: err.stack }));
    return output;
  }
}


// ─── doPost — збереження табелю ─────────────────────────────────────────────
//
// type:'arrival' → записує рядок-заглушку з приходом, шле Telegram "прийшов"
// type:'full'    → знаходить рядок приходу і оновлює його; решта проектів —
//                  нові рядки. Шле Telegram "пішов"
//
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const raw = (e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const data = JSON.parse(raw);
    const type = data.type || 'full';
    const tabSheetName = (data.tabSheet || TAB_SHEET_DEFAULT).toString();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(tabSheetName);
    if (!sheet) {
      sheet = ss.insertSheet(tabSheetName);
      _addHeaders(sheet);
    } else if (sheet.getLastRow() === 0) {
      _addHeaders(sheet);
    }

    // ── ARRIVAL ─────────────────────────────────────────────────────────────
    if (type === 'arrival') {
      sheet.appendRow([
        data.date     || '',   // A  Дата
        data.month    || '',   // B  Місяць
        data.employee || '',   // C  Співробітник
        data.ip       || '',   // D  IP-адреса
        data.location || '',   // E  Місце роботи
        data.arrival  || '',   // F  Прихід
        '', '', '', '', '', '', '', '', '', '', '', '', ''   // G–S  порожньо
      ]);

      const msg = '🟢 *' + _escMd(data.employee) + '* прийшов о *' + data.arrival + '*\n'
        + '📍 ' + _escMd(data.location) + '   🌐 `' + data.ip + '`\n'
        + '📅 ' + data.date;
      _sendTelegram(TG_TOKEN, TG_CHAT_ID, msg);

      output.setContent(JSON.stringify({ ok: true, type: 'arrival' }));
      return output;
    }

    // ── FULL ─────────────────────────────────────────────────────────────────
    if (type === 'full') {
      const rows = data.rows || [];
      if (!rows.length) {
        output.setContent(JSON.stringify({ ok: true, added: 0, message: 'Немає рядків' }));
        return output;
      }

      const r0 = rows[0];

      // Знайти рядок-заглушку (прихід без відходу) для цього співробітника і дати
      let arrivalRowIdx = -1;
      if (sheet.getLastRow() >= 2) {
        const tz = Session.getScriptTimeZone();
        const sheetData = sheet.getDataRange().getValues();
        for (let i = 1; i < sheetData.length; i++) {
          const raw = sheetData[i][0];
          const rowDate = raw instanceof Date
            ? Utilities.formatDate(raw, tz, 'yyyy-MM-dd')
            : (raw || '').toString().slice(0, 10);
          const rowEmp  = (sheetData[i][2] || '').toString().trim();
          const rowDep  = (sheetData[i][6] || '').toString().trim();
          if (rowDate === r0.date && rowEmp === r0.employee && rowDep === '') {
            arrivalRowIdx = i + 1; // 1-indexed
            break;
          }
        }
      }

      const makeRow = r => [
        r.date          || '',   // A  Дата
        r.month         || '',   // B  Місяць
        r.employee      || '',   // C  Співробітник
        r.ip            || '',   // D  IP-адреса
        r.location      || '',   // E  Місце роботи
        r.arrival       || '',   // F  Прихід
        r.departure     || '',   // G  Відхід
        r.lunchMin      || '',   // H  Обід (хв)
        r.lunchStr      || '',   // I  Обід (текст)
        r.hoursGross    || '',   // J  Загальний час
        r.hours         || '',   // K  Чистий робочий час
        r.code          || '',   // L  Шифр
        r.name          || '',   // M  Назва об'єкта
        r.workType      || '',   // N  Вид робіт
        r.timeSpent     || '',   // O  Витрачений час
        r.section       || '',   // P  Розділ
        r.desc          || '',   // Q  Опис робіт
        r.tomorrowPlan  || '',   // R  Плани на завтра
        r.tomorrowDesc  || ''    // S  Причина (завтра)
      ];

      rows.forEach((r, idx) => {
        if (idx === 0 && arrivalRowIdx > 0) {
          sheet.getRange(arrivalRowIdx, 1, 1, 19).setValues([makeRow(r)]);
        } else {
          sheet.appendRow(makeRow(r));
        }

        // ── Telegram на кожен рядок ──
        let msg = '🔴 *' + _escMd(r.employee) + '* пішов о *' + r.departure + '*\n'
          + '⏱ Відпрацював: *' + r.hours + '*'
          + (r.lunchStr ? ' (обід ' + r.lunchStr + ')' : '') + '\n'
          + '🏗 `' + _escMd(r.code) + '`'
          + (r.name ? ' — ' + _escMd(r.name) : '') + '\n';
        if (r.timeSpent) msg += '⏰ На проект: *' + r.timeSpent + '*\n';
        if (r.workType)  msg += '⚒ ' + _escMd(r.workType) + '\n';
        msg += '📅 ' + r.date;
        _sendTelegram(TG_TOKEN, TG_CHAT_ID, msg);
      });

      output.setContent(JSON.stringify({
        ok: true, type: 'full', added: rows.length, sheet: tabSheetName
      }));
      return output;
    }

    output.setContent(JSON.stringify({ ok: false, error: 'Невідомий тип: ' + type }));
    return output;

  } catch (err) {
    output.setContent(JSON.stringify({ ok: false, error: err.message }));
    return output;
  }
}


// ─── Щоденне зведення о 20:00 (тригер) ─────────────────────────────────────
function sendDailySummary() {
  const token   = TG_TOKEN;
  const chatId  = TG_CHAT_ID;
  const tabName = TAB_SHEET_DEFAULT;

  if (!token || !chatId) {
    Logger.log('sendDailySummary: TG_TOKEN або TG_CHAT_ID не заповнено');
    return;
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  if (!sheet || sheet.getLastRow() < 2) {
    _sendTelegram(token, chatId, '📊 Зведення за ' + today + ': немає записів.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const people = {};

  for (let i = 1; i < data.length; i++) {
    const row     = data[i];
    const rowDate = (row[0] || '').toString().slice(0, 10);
    if (rowDate !== today) continue;

    const emp = (row[2] || '').toString().trim();
    if (!emp) continue;

    if (!people[emp]) {
      people[emp] = { arrival: row[5] || '', departure: '', hours: '', location: row[4] || '', projects: [] };
    }
    if (row[6])  people[emp].departure = row[6].toString();
    if (row[10]) people[emp].hours     = row[10].toString();
    if (row[11]) people[emp].projects.push((row[11] || '') + (row[12] ? ' — ' + row[12] : ''));
  }

  const names = Object.keys(people);
  if (names.length === 0) {
    _sendTelegram(token, chatId, '📊 Зведення за ' + today + ': ніхто не відмітився.');
    return;
  }

  let msg = '📊 *Зведення за ' + today + '*\n';
  msg += '👥 Відмітилось: ' + names.length + ' осіб\n\n';

  names.forEach(emp => {
    const p = people[emp];
    const dep = p.departure ? p.departure : '(ще не пішов)';
    msg += '👤 *' + _escMd(emp) + '*\n';
    msg += '   🕐 ' + p.arrival + ' → ' + dep;
    if (p.hours) msg += '   ⏱ ' + p.hours;
    msg += '\n';
    if (p.projects.length > 0) {
      msg += '   🏗 ' + p.projects.slice(0, 3).map(_escMd).join('; ') + '\n';
    }
    msg += '\n';
  });

  _sendTelegram(token, chatId, msg.trim());
}


// ─── Тест Telegram (запускати вручну з редактора) ────────────────────────────
function testTelegram() {
  _sendTelegram(TG_TOKEN, TG_CHAT_ID, '✅ Тест з Apps Script — працює!');
}


// ─── Приватні функції ────────────────────────────────────────────────────────

function _sendTelegram(token, chatId, text) {
  try {
    const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: 'Markdown' }),
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());
    if (!result.ok) {
      console.error('Telegram send failed [' + result.error_code + ']: ' + result.description
        + ' | text preview: ' + text.slice(0, 120));
    }
  } catch (e) {
    console.error('Telegram exception: ' + e.message);
  }
}

function _escMd(s) {
  return (s || '').replace(/([_*`\[])/g, '\\$1');
}

function _addHeaders(sheet) {
  const HEADERS = [
    'Дата', 'Місяць', 'Співробітник', 'IP-адреса', 'Місце роботи', 'Прихід', 'Відхід',
    'Обід (хв)', 'Обід (текст)', 'Загальний час', 'Чистий роб. час',
    'Шифр', 'Назва об\'єкта', 'Вид робіт', 'Витрачений час', 'Розділ', 'Опис робіт',
    'Плани на завтра', 'Причина (завтра)'
  ];
  sheet.appendRow(HEADERS);

  const hdr = sheet.getRange(1, 1, 1, HEADERS.length);
  hdr.setBackground('#1a237e');
  hdr.setFontColor('#ffffff');
  hdr.setFontWeight('bold');
  hdr.setFontSize(10);
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1,  95);   // A  Дата
  sheet.setColumnWidth(2,  90);   // B  Місяць
  sheet.setColumnWidth(3, 130);   // C  Співробітник
  sheet.setColumnWidth(4, 110);   // D  IP-адреса
  sheet.setColumnWidth(5, 100);   // E  Місце роботи
  sheet.setColumnWidth(6,  70);   // F  Прихід
  sheet.setColumnWidth(7,  70);   // G  Відхід
  sheet.setColumnWidth(8,  70);   // H  Обід (хв)
  sheet.setColumnWidth(9,  85);   // I  Обід (текст)
  sheet.setColumnWidth(10, 95);   // J  Загальний час
  sheet.setColumnWidth(11, 95);   // K  Чистий час
  sheet.setColumnWidth(12, 130);  // L  Шифр
  sheet.setColumnWidth(13, 270);  // M  Назва об'єкта
  sheet.setColumnWidth(14, 160);  // N  Вид робіт
  sheet.setColumnWidth(15,  80);  // O  Витрачений час
  sheet.setColumnWidth(16, 200);  // P  Розділ
  sheet.setColumnWidth(17, 260);  // Q  Опис робіт
  sheet.setColumnWidth(18, 110);  // R  Плани на завтра
  sheet.setColumnWidth(19, 220);  // S  Причина (завтра)
}
