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

// ─── Конфігурація ────────────────────────────────────────────────────────────
const TG_TOKEN          = '8303053237:AAHqd_SSIjM-3l08EKRbyJF7MFAHXPDDPAY';
const TG_CHAT_ID        = '222538505';
const TAB_SHEET_DEFAULT = 'ТАБЕЛЬ';
const IP_WHITELIST_SHEET = 'IP_WHITELIST';


// ─── doGet — читання даних ───────────────────────────────────────────────────
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const action = (e.parameter && e.parameter.action) || '';

    // ── Ping ──────────────────────────────────────────────────────────────────
    if (action === 'ping' || action === '') {
      output.setContent(JSON.stringify({ ok: true, message: 'Табель API працює' }));
      return output;
    }

    // ── Списки співробітників / об'єктів / розділів / видів робіт ─────────────
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

      const employees = _readCol(ss, empSheet, empCol)
        .map(v => v.toString().trim()).filter(v => v.length > 0);

      const objects = _readCol(ss, objSheet, objCol)
        .map(v => v.toString().replace(/\n/g, ' ').trim()).filter(v => v.length > 0)
        .map(raw => {
          const sepIdx = raw.indexOf(' | ');
          return sepIdx > 0
            ? { code: raw.slice(0, sepIdx).trim(), name: raw.slice(sepIdx + 3).trim() }
            : { code: raw, name: '' };
        });

      const sections  = _readCol(ss, secSheet, secCol)
        .map(v => v.toString().trim()).filter(v => v.length > 0);

      const worktypes = _readCol(ss, wtSheet, wtCol)
        .map(v => v.toString().trim()).filter(v => v.length > 0);

      output.setContent(JSON.stringify({
        ok: true, employees, objects, sections, worktypes,
        meta: {
          empSheet, empCol, empCount: employees.length,
          objSheet, objCol, objCount: objects.length,
          secSheet, secCol, secCount: sections.length,
          wtSheet,  wtCol,  wtCount:  worktypes.length,
          timestamp: new Date().toISOString()
        }
      }));
      return output;
    }

    // ── Серверний час ──────────────────────────────────────────────────────────
    if (action === 'getServerTime') {
      const now = new Date();
      const tz  = Session.getScriptTimeZone();
      output.setContent(JSON.stringify({
        ok:    true,
        iso:   now.toISOString(),
        date:  Utilities.formatDate(now, tz, 'yyyy-MM-dd'),
        time:  Utilities.formatDate(now, tz, 'HH:mm'),
        epoch: now.getTime()
      }));
      return output;
    }

    // ── Перевірка незакритої сесії (прихід без відходу) ───────────────────────
    if (action === 'checkSession') {
      const employee     = (e.parameter.employee  || '').toString().trim();
      const date         = (e.parameter.date      || '').toString().trim();
      const tabSheetName = (e.parameter.tabSheet  || TAB_SHEET_DEFAULT).toString();

      if (!employee || !date) {
        output.setContent(JSON.stringify({ ok: true, hasOpenSession: false }));
        return output;
      }

      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(tabSheetName);

      if (!sheet || sheet.getLastRow() < 2) {
        output.setContent(JSON.stringify({ ok: true, hasOpenSession: false }));
        return output;
      }

      const tz        = Session.getScriptTimeZone();
      const sheetData = sheet.getDataRange().getValues();

      for (let i = 1; i < sheetData.length; i++) {
        const row     = sheetData[i];
        const rawDate = row[0];
        const rowDate = rawDate instanceof Date
          ? Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd')
          : (rawDate || '').toString().slice(0, 10);
        const rowEmp = (row[2] || '').toString().trim();
        const rowDep = (row[6] || '').toString().trim();

        if (rowDate === date && rowEmp === employee && rowDep === '') {
          const arrivalIP = (row[3] || '').toString().split(' / ')[0].trim();
          const rawArr = row[5];
          const arrivalStr = rawArr instanceof Date
            ? Utilities.formatDate(rawArr, tz, 'HH:mm')
            : (rawArr || '').toString().trim();
          output.setContent(JSON.stringify({
            ok: true, hasOpenSession: true,
            date:     date,
            month:    (row[1] || '').toString(),
            arrival:  arrivalStr,
            location: (row[4] || '').toString(),
            ip:       arrivalIP
          }));
          return output;
        }
      }

      output.setContent(JSON.stringify({ ok: true, hasOpenSession: false }));
      return output;
    }

    // ── Довідник білих IP ──────────────────────────────────────────────────────
    if (action === 'getWhiteIPs') {
      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(IP_WHITELIST_SHEET);

      if (!sheet || sheet.getLastRow() < 2) {
        output.setContent(JSON.stringify({ ok: true, ips: [] }));
        return output;
      }

      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
      const ips  = rows
        .filter(r => r[0])
        .map(r => ({
          ip:       r[0].toString().trim(),
          location: r[1].toString().trim(),
          employee: r[2].toString().trim() || 'Всі',
          comment:  r[3].toString().trim()
        }));

      output.setContent(JSON.stringify({ ok: true, ips }));
      return output;
    }

    // ── Статистика користувача ─────────────────────────────────────────────────
    if (action === 'myStats') {
      const employee     = (e.parameter.employee || '').toString().trim();
      const month        = (e.parameter.month    || '').toString().trim(); // YYYY-MM
      const tabSheetName = (e.parameter.tabSheet || TAB_SHEET_DEFAULT).toString();

      if (!employee || !month) {
        output.setContent(JSON.stringify({ ok: false, error: 'Потрібні employee та month' }));
        return output;
      }

      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(tabSheetName);

      if (!sheet || sheet.getLastRow() < 2) {
        output.setContent(JSON.stringify({ ok: true, records: [] }));
        return output;
      }

      const tz      = Session.getScriptTimeZone();
      const allData = sheet.getDataRange().getValues();
      const records = [];

      for (let i = 1; i < allData.length; i++) {
        const row     = allData[i];
        const rawDate = row[0];
        const rowDate = rawDate instanceof Date
          ? Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd')
          : (rawDate || '').toString().slice(0, 10);
        const rowEmp = (row[2] || '').toString().trim();

        if (!rowDate.startsWith(month) || rowEmp !== employee) continue;

        records.push({
          date:       rowDate,
          location:   (row[4]  || '').toString(),
          arrival:    row[5] instanceof Date ? Utilities.formatDate(row[5], tz, 'HH:mm') : (row[5]||'').toString().trim(),
          departure:  row[6] instanceof Date ? Utilities.formatDate(row[6], tz, 'HH:mm') : (row[6]||'').toString().trim(),
          lunchMin:   Number(row[7]) || 0,
          hoursGross: (row[9]  || '').toString(),
          hours:      (row[10] || '').toString(),
          code:       (row[11] || '').toString(),
          name:       (row[12] || '').toString(),
          workType:   (row[13] || '').toString(),
          timeSpent:  (row[14] || '').toString(),
          section:    (row[15] || '').toString()
        });
      }

      output.setContent(JSON.stringify({ ok: true, records }));
      return output;
    }

    // ── Адмін-моніторинг ──────────────────────────────────────────────────────
    if (action === 'adminStats') {
      const tabSheetName = (e.parameter.tabSheet || TAB_SHEET_DEFAULT).toString();
      const filterMonth  = (e.parameter.month    || '').toString().trim();

      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(tabSheetName);
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

      if (!sheet || sheet.getLastRow() < 2) {
        output.setContent(JSON.stringify({ ok: true, today, onlineNow: [], projectStats: [], suspiciousRecords: [] }));
        return output;
      }

      const tz         = Session.getScriptTimeZone();
      const allData    = sheet.getDataRange().getValues();
      const onlineNow  = [];
      const projectMap = {};
      const suspicious = [];

      for (let i = 1; i < allData.length; i++) {
        const row     = allData[i];
        const rawDate = row[0];
        const rowDate = rawDate instanceof Date
          ? Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd')
          : (rawDate || '').toString().slice(0, 10);

        const emp       = (row[2]  || '').toString().trim();
        const dep       = row[6] instanceof Date ? Utilities.formatDate(row[6], tz, 'HH:mm') : (row[6]||'').toString().trim();
        const arrival   = row[5] instanceof Date ? Utilities.formatDate(row[5], tz, 'HH:mm') : (row[5]||'').toString().trim();
        const location  = (row[4]  || '').toString();
        const ip        = (row[3]  || '').toString();
        const code      = (row[11] || '').toString().trim();
        const name      = (row[12] || '').toString().trim();
        const timeSpent = (row[14] || '').toString().trim();
        const flag      = (row[19] || '').toString().trim();

        if (rowDate === today && emp && !dep) {
          if (!onlineNow.some(x => x.employee === emp)) {
            onlineNow.push({ employee: emp, arrival, location, ip });
          }
        }

        if (code && timeSpent && (!filterMonth || rowDate.startsWith(filterMonth))) {
          if (!projectMap[code]) {
            projectMap[code] = { code, name, totalMins: 0, employees: [] };
          }
          const parts = timeSpent.split(':');
          projectMap[code].totalMins += (parseInt(parts[0]) || 0) * 60 + (parseInt(parts[1]) || 0);
          if (emp && !projectMap[code].employees.includes(emp)) {
            projectMap[code].employees.push(emp);
          }
        }

        if (flag && flag.includes('⚠')) {
          suspicious.push({ date: rowDate, employee: emp, ip, location, flag, arrival, departure: dep });
        }
      }

      const projectStats = Object.values(projectMap)
        .sort((a, b) => b.totalMins - a.totalMins)
        .slice(0, 40);

      output.setContent(JSON.stringify({
        ok: true, today,
        onlineNow,
        projectStats,
        suspiciousRecords: suspicious.slice(0, 60)
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


// ─── doPost — запис табелю ───────────────────────────────────────────────────
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const raw  = (e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const data = JSON.parse(raw);
    const type = data.type || 'full';
    const tabSheetName = (data.tabSheet || TAB_SHEET_DEFAULT).toString();

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Збереження довідника білих IP ──────────────────────────────────────────
    if (type === 'saveWhiteIPs') {
      const ips = data.ips || [];
      let ipSheet = ss.getSheetByName(IP_WHITELIST_SHEET);

      if (!ipSheet) {
        ipSheet = ss.insertSheet(IP_WHITELIST_SHEET);
        ipSheet.appendRow(['IP-адреса', 'Локація', 'Співробітник', 'Коментар']);
        ipSheet.getRange(1, 1, 1, 4).setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold');
        ipSheet.setFrozenRows(1);
        ipSheet.setColumnWidth(1, 140);
        ipSheet.setColumnWidth(2, 110);
        ipSheet.setColumnWidth(3, 150);
        ipSheet.setColumnWidth(4, 220);
      } else if (ipSheet.getLastRow() > 1) {
        ipSheet.getRange(2, 1, ipSheet.getLastRow() - 1, 4).clearContent();
      }

      if (ips.length > 0) {
        const rows = ips.map(entry => [
          entry.ip       || '',
          entry.location || '',
          entry.employee || 'Всі',
          entry.comment  || ''
        ]);
        ipSheet.getRange(2, 1, rows.length, 4).setValues(rows);
      }

      output.setContent(JSON.stringify({ ok: true, saved: ips.length }));
      return output;
    }

    // ── Підготовка листа табелю ────────────────────────────────────────────────
    let sheet = ss.getSheetByName(tabSheetName);
    if (!sheet) {
      sheet = ss.insertSheet(tabSheetName);
      _addHeaders(sheet);
    } else if (sheet.getLastRow() === 0) {
      _addHeaders(sheet);
    }

    // ── ARRIVAL — фіксація приходу ─────────────────────────────────────────────
    if (type === 'arrival') {
      sheet.appendRow([
        data.date     || '',   // A  Дата
        data.month    || '',   // B  Місяць
        data.employee || '',   // C  Співробітник
        data.ip       || '',   // D  IP-адреса (прихід)
        data.location || '',   // E  Місце роботи
        data.arrival  || '',   // F  Прихід
        '', '', '', '', '', '', '', '', '', '', '', '', '', ''  // G–T  порожньо
      ]);

      const msg = '🟢 *' + _escMd(data.employee) + '* прийшов о *' + data.arrival + '*\n'
        + '📍 ' + _escMd(data.location) + '   🌐 *' + data.ip + '*\n'
        + '📅 ' + data.date;
      _sendTelegram(TG_TOKEN, TG_CHAT_ID, msg);

      output.setContent(JSON.stringify({ ok: true, type: 'arrival' }));
      return output;
    }

    // ── FULL — повний запис (відхід + проекти) ─────────────────────────────────
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
        const tz        = Session.getScriptTimeZone();
        const sheetData = sheet.getDataRange().getValues();
        for (let i = 1; i < sheetData.length; i++) {
          const rawDate = sheetData[i][0];
          const rowDate = rawDate instanceof Date
            ? Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd')
            : (rawDate || '').toString().slice(0, 10);
          const rowEmp = (sheetData[i][2] || '').toString().trim();
          const rowDep = (sheetData[i][6] || '').toString().trim();
          if (rowDate === r0.date && rowEmp === r0.employee && rowDep === '') {
            arrivalRowIdx = i + 1;
            break;
          }
        }
      }

      const ipWhiteList = _getWhiteIPs(ss);

      const makeRow = r => [
        r.date         || '',        // A  Дата
        r.month        || '',        // B  Місяць
        r.employee     || '',        // C  Співробітник
        _mergeIP(r.ip, r.ipDep),     // D  IP (прихід / відхід)
        r.location     || '',        // E  Місце роботи
        r.arrival      || '',        // F  Прихід
        r.departure    || '',        // G  Відхід
        r.lunchMin     || '',        // H  Обід (хв)
        r.lunchStr     || '',        // I  Обід (текст)
        r.hoursGross   || '',        // J  Загальний час
        r.hours        || '',        // K  Чистий роб. час
        r.code         || '',        // L  Шифр
        r.name         || '',        // M  Назва об'єкта
        r.workType     || '',        // N  Вид робіт
        r.timeSpent    || '',        // O  Витрачений час
        r.section      || '',        // P  Розділ
        r.desc         || '',        // Q  Опис робіт
        r.tomorrowPlan || '',        // R  Плани на завтра
        r.tomorrowDesc || '',        // S  Причина (завтра)
        ''                           // T  Перевірка IP (заповнюється нижче)
      ];

      rows.forEach((r, idx) => {
        let rowIndex;
        if (idx === 0 && arrivalRowIdx > 0) {
          sheet.getRange(arrivalRowIdx, 1, 1, 20).setValues([makeRow(r)]);
          rowIndex = arrivalRowIdx;
        } else {
          sheet.appendRow(makeRow(r));
          rowIndex = sheet.getLastRow();
        }

        // Перевірка IP — якщо IP відходу відомий використовуємо його (більш актуальний)
        const checkIP = (r.ipDep && r.ipDep !== 'невідомо') ? r.ipDep : r.ip;
        if (_isSuspiciousIP(checkIP, r.location, r.employee, ipWhiteList)) {
          sheet.getRange(rowIndex, 1, 1, 20).setBackground('#fff59d');
          sheet.getRange(rowIndex, 20).setValue('⚠ IP не збігається');
        }

        // Telegram-повідомлення
        let msg = '🔴 *' + _escMd(r.employee) + '* пішов о *' + r.departure + '*\n'
          + '⏱ Відпрацював: *' + r.hours + '*'
          + (r.lunchStr ? ' (обід ' + r.lunchStr + ')' : '') + '\n'
          + '🏗 *' + _escMd(r.code) + '*'
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


// ─── Щоденне зведення о 20:00 (встановити як тригер Time-driven → 8pm-9pm) ──
function sendDailySummary() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_SHEET_DEFAULT);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  if (!sheet || sheet.getLastRow() < 2) {
    _sendTelegram(TG_TOKEN, TG_CHAT_ID, '📊 Зведення за ' + today + ': немає записів.');
    return;
  }

  const data   = sheet.getDataRange().getValues();
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
    _sendTelegram(TG_TOKEN, TG_CHAT_ID, '📊 Зведення за ' + today + ': ніхто не відмітився.');
    return;
  }

  let msg = '📊 *Зведення за ' + today + '*\n';
  msg += '👥 Відмітилось: ' + names.length + ' осіб\n\n';
  names.forEach(emp => {
    const p   = people[emp];
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
  _sendTelegram(TG_TOKEN, TG_CHAT_ID, msg.trim());
}


// ─── Тест Telegram ───────────────────────────────────────────────────────────
function testTelegram() {
  _sendTelegram(TG_TOKEN, TG_CHAT_ID, '✅ Тест з Apps Script — працює!');
}


// ─── Приватні функції ────────────────────────────────────────────────────────

function _readCol(ss, sheetName, col) {
  const sht = ss.getSheetByName(sheetName);
  if (!sht || sht.getLastRow() < 2) return [];
  return sht.getRange(col + '2:' + col + sht.getLastRow()).getValues().flat();
}

function _mergeIP(ipArr, ipDep) {
  const a = (ipArr || '').toString().trim();
  const d = (ipDep || '').toString().trim();
  if (!d || d === 'невідомо') return a;
  return a + ' / ' + d;
}

function _getWhiteIPs(ss) {
  const sheet = ss.getSheetByName(IP_WHITELIST_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(r => r[0])
    .map(r => ({
      ip:       r[0].toString().trim(),
      location: r[1].toString().trim(),
      employee: r[2].toString().trim() || 'Всі'
    }));
}

function _isSuspiciousIP(ip, location, employee, whitelist) {
  if (!whitelist || whitelist.length === 0) return false;
  const matches = whitelist.filter(e => e.ip === ip);
  if (matches.length === 0) return true;
  return !matches.some(e =>
    e.location === location &&
    (e.employee === 'Всі' || e.employee === employee)
  );
}

function _sendTelegram(token, chatId, text) {
  try {
    const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: 'Markdown' }),
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());
    if (!result.ok) {
      console.error('Telegram error [' + result.error_code + ']: ' + result.description);
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
    'Шифр', "Назва об'єкта", 'Вид робіт', 'Витрачений час', 'Розділ', 'Опис робіт',
    'Плани на завтра', 'Причина (завтра)', 'Перевірка IP'
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
  sheet.setColumnWidth(4, 200);   // D  IP-адреса (прихід / відхід)
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
  sheet.setColumnWidth(20, 150);  // T  Перевірка IP
}
