/**
 * ══════════════════════════════════════════════
 *  ТАБЕЛЬ 2026 — Google Apps Script  v26.05.10
 * ══════════════════════════════════════════════
 *
 *  ІНСТРУКЦІЯ:
 *  1. Вставте цей код у Apps Script
 *  2. Збережіть (Ctrl+S)
 *  3. Розгорнути → Нове розгортання
 *     • Тип: Веб-застосунок
 *     • Виконувати від імені: Я
 *     • Доступ: Всі (анонімні)
 *  4. Скопіюйте URL → вставте у форму ⚙ Меню → Налаштування
 *
 *  ДЛЯ ЩОДЕННОГО ЗВЕДЕННЯ О 20:00:
 *  Triggers → Додати тригер → sendDailySummary
 *  → Time-driven → Day timer → 8pm-9pm
 *
 *  ВАЖЛИВО: після кожної зміни коду — НОВЕ розгортання
 * ══════════════════════════════════════════════
 */

// ─── Конфігурація ────────────────────────────────────────────────────────────
const TG_TOKEN           = '8303053237:AAHqd_SSIjM-3l08EKRbyJF7MFAHXPDDPAY';
const TG_CHAT_ID         = '222538505';
const TAB_SHEET_DEFAULT  = 'ТАБЕЛЬ';
const IP_WHITELIST_SHEET = 'IP_WHITELIST';
const OTP_SHEET          = 'КОДИ';
const REQ_SHEET          = 'ЗАПИТИ';
const OTP_TTL_MIN        = 30; // хвилин до закінчення дії коду


// ═══════════════════════════════════════════════════════════════════════════════
//  doGet — читання даних
// ═══════════════════════════════════════════════════════════════════════════════
function doGet(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);

  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const tz     = Session.getScriptTimeZone();

    // ── ping ─────────────────────────────────────────────────────────────────
    if (!action || action === 'ping') {
      out.setContent(JSON.stringify({ ok: true, message: 'Табель API v26.05.10 працює' }));
      return out;
    }

    // ── getData ───────────────────────────────────────────────────────────────
    if (action === 'getData') {
      const p = e.parameter;
      const employees = _readCol(ss, p.empSheet||'СПРАВОЧНИК', (p.empCol||'A').toUpperCase())
        .map(v=>v.toString().trim()).filter(Boolean);
      const objects = _readCol(ss, p.objSheet||'СПРАВОЧНИК', (p.objCol||'C').toUpperCase())
        .map(v=>v.toString().replace(/\n/g,' ').trim()).filter(Boolean)
        .map(raw=>{const i=raw.indexOf(' | ');return i>0?{code:raw.slice(0,i).trim(),name:raw.slice(i+3).trim()}:{code:raw,name:''};});
      const sections  = _readCol(ss, p.secSheet||'СПРАВОЧНИК', (p.secCol||'B').toUpperCase()).map(v=>v.toString().trim()).filter(Boolean);
      const worktypes = _readCol(ss, p.wtSheet||'СПРАВОЧНИК',  (p.wtCol||'D').toUpperCase()).map(v=>v.toString().trim()).filter(Boolean);
      out.setContent(JSON.stringify({
        ok:true, employees, objects, sections, worktypes,
        meta:{timestamp:new Date().toISOString(), empCount:employees.length, objCount:objects.length}
      }));
      return out;
    }

    // ── getServerTime ─────────────────────────────────────────────────────────
    if (action === 'getServerTime') {
      const now = new Date();
      out.setContent(JSON.stringify({
        ok:true, iso:now.toISOString(),
        date:Utilities.formatDate(now,tz,'yyyy-MM-dd'),
        time:Utilities.formatDate(now,tz,'HH:mm'),
        epoch:now.getTime()
      }));
      return out;
    }

    // ── checkSession — незакрита сесія сьогодні ───────────────────────────────
    if (action === 'checkSession') {
      const employee = (e.parameter.employee||'').trim();
      const date     = (e.parameter.date||'').trim();
      const tabName  = e.parameter.tabSheet||TAB_SHEET_DEFAULT;
      if (!employee||!date) { out.setContent(JSON.stringify({ok:true,hasOpenSession:false})); return out; }
      const sheet = ss.getSheetByName(tabName);
      if (!sheet||sheet.getLastRow()<2) { out.setContent(JSON.stringify({ok:true,hasOpenSession:false})); return out; }
      const data = sheet.getDataRange().getValues();
      for (let i=1;i<data.length;i++) {
        const row    = data[i];
        const rDate  = _fmtDate(row[0],tz);
        const rEmp   = (row[2]||'').toString().trim();
        const rDep   = (row[6]||'').toString().trim();
        if (rDate===date && rEmp===employee && rDep==='') {
          out.setContent(JSON.stringify({
            ok:true, hasOpenSession:true,
            date, month:(row[1]||'').toString(),
            arrival:_fmtTime(row[5],tz),
            location:(row[4]||'').toString(),
            ip:(row[3]||'').toString().split(' / ')[0].trim()
          }));
          return out;
        }
      }
      out.setContent(JSON.stringify({ok:true,hasOpenSession:false}));
      return out;
    }

    // ── checkPrevSession — незакриті попередні дні ────────────────────────────
    if (action === 'checkPrevSession') {
      const employee = (e.parameter.employee||'').trim();
      const today    = (e.parameter.today||'').trim();
      const tabName  = e.parameter.tabSheet||TAB_SHEET_DEFAULT;
      if (!employee) { out.setContent(JSON.stringify({ok:true,hasOpenSession:false})); return out; }
      const sheet = ss.getSheetByName(tabName);
      if (!sheet||sheet.getLastRow()<2) { out.setContent(JSON.stringify({ok:true,hasOpenSession:false})); return out; }
      const data = sheet.getDataRange().getValues();
      // шукаємо з кінця — найновіший незакритий
      for (let i=data.length-1;i>=1;i--) {
        const row   = data[i];
        const rDate = _fmtDate(row[0],tz);
        const rEmp  = (row[2]||'').toString().trim();
        const rDep  = (row[6]||'').toString().trim();
        const rArr  = _fmtTime(row[5],tz);
        if (rDate===today||rEmp!==employee||rDep!==''||!rArr||rArr==='—') continue;
        out.setContent(JSON.stringify({
          ok:true, hasOpenSession:true,
          date:rDate, month:(row[1]||'').toString(),
          arrival:rArr, location:(row[4]||'').toString(),
          ip:(row[3]||'').toString().split(' / ')[0].trim()
        }));
        return out;
      }
      out.setContent(JSON.stringify({ok:true,hasOpenSession:false}));
      return out;
    }

    // ── verifyOTP — перевірка коду підтвердження ──────────────────────────────
    if (action === 'verifyOTP') {
      const employee = (e.parameter.employee||'').trim();
      const code     = (e.parameter.code||'').trim();
      const date     = (e.parameter.date||'').trim();
      const sheet    = ss.getSheetByName(OTP_SHEET);
      if (!sheet||sheet.getLastRow()<2||!code||!employee) {
        out.setContent(JSON.stringify({ok:true,valid:false,error:'Код не знайдено'}));
        return out;
      }
      const now  = new Date();
      const rows = sheet.getDataRange().getValues();
      for (let i=1;i<rows.length;i++) {
        const r      = rows[i];
        const rEmp   = (r[0]||'').toString().trim();
        const rDate  = r[1] instanceof Date ? Utilities.formatDate(r[1],tz,'yyyy-MM-dd') : (r[1]||'').toString().trim();
        const rCode  = (r[2]||'').toString().trim();
        const rTs    = r[3];  // timestamp Date
        const rStat  = (r[4]||'').toString().trim();
        if (rEmp!==employee||rDate!==date||rCode!==code) continue;
        if (rStat==='used') {
          out.setContent(JSON.stringify({ok:true,valid:false,error:'Код вже використано'}));
          return out;
        }
        const ageMins = (now - (rTs instanceof Date ? rTs : new Date(rTs))) / 60000;
        if (ageMins > OTP_TTL_MIN) {
          sheet.getRange(i+1,5).setValue('expired');
          out.setContent(JSON.stringify({ok:true,valid:false,error:'Термін дії коду вийшов ('+OTP_TTL_MIN+' хв)'}));
          return out;
        }
        sheet.getRange(i+1,5).setValue('used');
        out.setContent(JSON.stringify({ok:true,valid:true}));
        return out;
      }
      out.setContent(JSON.stringify({ok:true,valid:false,error:'Невірний код'}));
      return out;
    }

    // ── getRequests — список запитів на пропущені дні ─────────────────────────
    if (action === 'getRequests') {
      const filterEmp = (e.parameter.employee||'').trim();
      const sheet     = ss.getSheetByName(REQ_SHEET);
      if (!sheet||sheet.getLastRow()<2) { out.setContent(JSON.stringify({ok:true,requests:[]})); return out; }
      const rows = sheet.getRange(2,1,sheet.getLastRow()-1,7).getValues();
      const requests = rows
        .filter(r=>r[0]&&(!filterEmp||r[2].toString().trim()===filterEmp))
        .map(r=>({
          id:         r[0].toString(),
          requestDate:r[1] instanceof Date?Utilities.formatDate(r[1],tz,'yyyy-MM-dd HH:mm'):r[1].toString(),
          employee:   r[2].toString(),
          missedDate: r[3] instanceof Date?Utilities.formatDate(r[3],tz,'yyyy-MM-dd'):r[3].toString(),
          reason:     r[4].toString(),
          status:     r[5].toString()
        }));
      out.setContent(JSON.stringify({ok:true,requests}));
      return out;
    }

    // ── getWhiteIPs ───────────────────────────────────────────────────────────
    if (action === 'getWhiteIPs') {
      const sheet = ss.getSheetByName(IP_WHITELIST_SHEET);
      if (!sheet||sheet.getLastRow()<2) { out.setContent(JSON.stringify({ok:true,ips:[]})); return out; }
      const ips = sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues()
        .filter(r=>r[0])
        .map(r=>({ip:r[0].toString().trim(),location:r[1].toString().trim(),employee:r[2].toString().trim()||'Всі',comment:r[3].toString().trim()}));
      out.setContent(JSON.stringify({ok:true,ips}));
      return out;
    }

    // ── myStats — статистика співробітника ────────────────────────────────────
    if (action === 'myStats') {
      const employee = (e.parameter.employee||'').trim();
      const month    = (e.parameter.month||'').trim(); // YYYY-MM
      const tabName  = e.parameter.tabSheet||TAB_SHEET_DEFAULT;
      if (!employee||!month) { out.setContent(JSON.stringify({ok:false,error:'Потрібні employee та month'})); return out; }
      const sheet = ss.getSheetByName(tabName);
      if (!sheet||sheet.getLastRow()<2) { out.setContent(JSON.stringify({ok:true,records:[]})); return out; }
      const allData = sheet.getDataRange().getValues();
      const records = [];
      for (let i=1;i<allData.length;i++) {
        const row  = allData[i];
        const rDate= _fmtDate(row[0],tz);
        const rEmp = (row[2]||'').toString().trim();
        if (!rDate.startsWith(month)||rEmp!==employee) continue;
        records.push({
          date:rDate, location:(row[4]||'').toString(),
          arrival:_fmtTime(row[5],tz), departure:_fmtTime(row[6],tz),
          lunchMin:Number(row[7])||0,
          hoursGross:(row[9]||'').toString(), hours:(row[10]||'').toString(),
          code:(row[11]||'').toString(), name:(row[12]||'').toString(),
          workType:(row[13]||'').toString(), timeSpent:(row[14]||'').toString(),
          section:(row[15]||'').toString(), violation:(row[20]||'').toString()
        });
      }
      out.setContent(JSON.stringify({ok:true,records}));
      return out;
    }

    // ── adminStats — моніторинг для адміна ───────────────────────────────────
    if (action === 'adminStats') {
      const tabName     = e.parameter.tabSheet||TAB_SHEET_DEFAULT;
      const filterMonth = (e.parameter.month||'').trim();
      const today       = Utilities.formatDate(new Date(),tz,'yyyy-MM-dd');
      const sheet       = ss.getSheetByName(tabName);
      if (!sheet||sheet.getLastRow()<2) {
        out.setContent(JSON.stringify({ok:true,today,onlineNow:[],projectStats:[],suspiciousRecords:[]}));
        return out;
      }
      const allData=sheet.getDataRange().getValues();
      const onlineNow=[]; const projectMap={}; const suspicious=[];
      for (let i=1;i<allData.length;i++) {
        const row     = allData[i];
        const rDate   = _fmtDate(row[0],tz);
        const emp     = (row[2]||'').toString().trim();
        const dep     = _fmtTime(row[6],tz);
        const arrival = _fmtTime(row[5],tz);
        const loc     = (row[4]||'').toString();
        const ip      = (row[3]||'').toString();
        const code    = (row[11]||'').toString().trim();
        const name    = (row[12]||'').toString().trim();
        const ts      = (row[14]||'').toString().trim();
        const flag    = (row[19]||'').toString().trim();
        if (rDate===today&&emp&&!dep&&!onlineNow.some(x=>x.employee===emp)) {
          onlineNow.push({employee:emp,arrival,location:loc,ip});
        }
        if (code&&ts&&(!filterMonth||rDate.startsWith(filterMonth))) {
          if (!projectMap[code]) projectMap[code]={code,name,totalMins:0,employees:[]};
          const p=ts.split(':');
          projectMap[code].totalMins+=(parseInt(p[0])||0)*60+(parseInt(p[1])||0);
          if (emp&&!projectMap[code].employees.includes(emp)) projectMap[code].employees.push(emp);
        }
        if (flag&&flag.includes('⚠')) suspicious.push({date:rDate,employee:emp,ip,location:loc,flag,arrival,departure:dep});
      }
      out.setContent(JSON.stringify({
        ok:true,today,onlineNow,
        projectStats:Object.values(projectMap).sort((a,b)=>b.totalMins-a.totalMins).slice(0,40),
        suspiciousRecords:suspicious.slice(0,60)
      }));
      return out;
    }

    out.setContent(JSON.stringify({ok:false,error:'Невідома дія: '+action}));
    return out;

  } catch(err) {
    out.setContent(JSON.stringify({ok:false,error:err.message,stack:err.stack}));
    return out;
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
//  doPost — запис даних
// ═══════════════════════════════════════════════════════════════════════════════
function doPost(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);

  try {
    const raw  = (e.postData&&e.postData.contents)?e.postData.contents:'{}';
    const data = JSON.parse(raw);
    const type = data.type||'full';
    const tabName = data.tabSheet||TAB_SHEET_DEFAULT;
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const tz   = Session.getScriptTimeZone();

    // ── saveWhiteIPs ──────────────────────────────────────────────────────────
    if (type==='saveWhiteIPs') {
      const ips=data.ips||[];
      let sh=ss.getSheetByName(IP_WHITELIST_SHEET);
      if (!sh) {
        sh=ss.insertSheet(IP_WHITELIST_SHEET);
        sh.appendRow(['IP-адреса','Локація','Співробітник','Коментар']);
        sh.getRange(1,1,1,4).setBackground('#1a237e').setFontColor('#fff').setFontWeight('bold');
        sh.setFrozenRows(1);
        [140,110,150,220].forEach((w,i)=>sh.setColumnWidth(i+1,w));
      } else if (sh.getLastRow()>1) {
        sh.getRange(2,1,sh.getLastRow()-1,4).clearContent();
      }
      if (ips.length>0) sh.getRange(2,1,ips.length,4).setValues(ips.map(r=>[r.ip||'',r.location||'',r.employee||'Всі',r.comment||'']));
      out.setContent(JSON.stringify({ok:true,saved:ips.length}));
      return out;
    }

    // ── requestOTP — запит OTP-коду для закриття незакритого дня ─────────────
    if (type==='requestOTP') {
      const employee = (data.employee||'').toString().trim();
      const date     = (data.date||'').toString().trim();
      const arrival  = (data.arrival||'').toString().trim();
      const location = (data.location||'').toString().trim();
      if (!employee||!date) {
        out.setContent(JSON.stringify({ok:false,error:'Не вказано employee або date'}));
        return out;
      }
      // Генеруємо 6-значний код
      const code = String(Math.floor(100000+Math.random()*900000));
      // Зберігаємо в лист КОДИ
      let otpSheet = ss.getSheetByName(OTP_SHEET);
      if (!otpSheet) {
        otpSheet = ss.insertSheet(OTP_SHEET);
        otpSheet.appendRow(['Співробітник','Дата дня','Код','Час створення','Статус']);
        otpSheet.getRange(1,1,1,5).setBackground('#1a237e').setFontColor('#fff').setFontWeight('bold');
        otpSheet.setFrozenRows(1);
        [130,100,60,150,80].forEach((w,i)=>otpSheet.setColumnWidth(i+1,w));
      }
      // Помічаємо старі коди цього співробітника як expired
      if (otpSheet.getLastRow()>1) {
        const rows=otpSheet.getRange(2,1,otpSheet.getLastRow()-1,5).getValues();
        rows.forEach((r,i)=>{
          if (r[0].toString().trim()===employee&&r[4].toString()==='pending') {
            otpSheet.getRange(i+2,5).setValue('expired');
          }
        });
      }
      otpSheet.appendRow([employee,date,code,new Date(),'pending']);
      // Форматуємо дату для повідомлення
      const [y,m,d2]=date.split('-');
      const dateFmt=(d2||date)+'.'+( m||'')+'.'+( y||'');
      const msg='⚠️ *Запит на закриття дня*\n'
        +'👤 '+_escMd(employee)+'\n'
        +'📅 '+dateFmt+'   🕐 Прихід: '+arrival+'   📍 '+_escMd(location)+'\n'
        +'🔑 Код підтвердження: *'+code+'*\n'
        +'_(дійсний '+OTP_TTL_MIN+' хв)_';
      _sendTelegram(TG_TOKEN,TG_CHAT_ID,msg);
      out.setContent(JSON.stringify({ok:true,message:'Код надіслано адміністратору'}));
      return out;
    }

    // ── requestMissedDay — запит на пропущений день ───────────────────────────
    if (type==='requestMissedDay') {
      let reqSheet=ss.getSheetByName(REQ_SHEET);
      if (!reqSheet) {
        reqSheet=ss.insertSheet(REQ_SHEET);
        reqSheet.appendRow(['ID','Дата запиту','Співробітник','Дата пропуску','Причина','Статус','Дата рішення']);
        reqSheet.getRange(1,1,1,7).setBackground('#1a237e').setFontColor('#fff').setFontWeight('bold');
        reqSheet.setFrozenRows(1);
        [80,140,130,100,300,80,140].forEach((w,i)=>reqSheet.setColumnWidth(i+1,w));
      }
      const id=Date.now().toString();
      reqSheet.appendRow([id,Utilities.formatDate(new Date(),tz,'yyyy-MM-dd HH:mm'),data.employee||'',data.missedDate||'',data.reason||'','pending','']);
      const msg='📩 *Запит на внесення пропущеного дня*\n'
        +'👤 '+_escMd(data.employee||'')+'\n'
        +'📅 Дата пропуску: *'+(data.missedDate||'')+'*\n'
        +'📝 Причина: '+_escMd(data.reason||'')+'\n'
        +'Очікує підтвердження в адмін-панелі';
      _sendTelegram(TG_TOKEN,TG_CHAT_ID,msg);
      out.setContent(JSON.stringify({ok:true,id}));
      return out;
    }

    // ── approveRequest / rejectRequest ────────────────────────────────────────
    if (type==='approveRequest'||type==='rejectRequest') {
      const reqId     =(data.id||'').toString();
      const newStatus = type==='approveRequest'?'approved':'rejected';
      const reqSheet  = ss.getSheetByName(REQ_SHEET);
      if (!reqSheet) { out.setContent(JSON.stringify({ok:false,error:'Лист ЗАПИТИ не знайдено'})); return out; }
      const rows=reqSheet.getDataRange().getValues();
      for (let i=1;i<rows.length;i++) {
        if (rows[i][0].toString()!==reqId) continue;
        reqSheet.getRange(i+1,6).setValue(newStatus);
        reqSheet.getRange(i+1,7).setValue(Utilities.formatDate(new Date(),tz,'yyyy-MM-dd HH:mm'));
        const emp=rows[i][2].toString();
        const missed=rows[i][3].toString();
        const icon=newStatus==='approved'?'✅':'❌';
        const label=newStatus==='approved'?'схвалено':'відхилено';
        _sendTelegram(TG_TOKEN,TG_CHAT_ID,icon+' Запит *'+label+'*\n👤 '+_escMd(emp)+'\n📅 '+missed);
        out.setContent(JSON.stringify({ok:true,status:newStatus}));
        return out;
      }
      out.setContent(JSON.stringify({ok:false,error:'Запит не знайдено'}));
      return out;
    }

    // ── missedDayFull — заповнення схваленого пропущеного дня (повний wizard) ─
    if (type==='missedDayFull') {
      let tabSheet=ss.getSheetByName(tabName);
      if (!tabSheet){tabSheet=ss.insertSheet(tabName);_addHeaders(tabSheet);}
      else if (tabSheet.getLastRow()===0){_addHeaders(tabSheet);}
      const rows=data.rows||[];
      if (!rows.length){out.setContent(JSON.stringify({ok:false,error:'Немає рядків'}));return out;}
      const r0=rows[0];
      rows.forEach(r=>{
        tabSheet.appendRow([
          r.date||'',r.month||'',r.employee||'',_mergeIP(r.ip,r.ipDep),r.location||'',
          r.arrival||'',r.departure||'',r.lunchMin||'',r.lunchStr||'',
          r.hoursGross||'',r.hours||'',r.code||'',r.name||'',
          r.workType||'',r.timeSpent||'',r.section||'',r.desc||'',
          r.tomorrowPlan||'',r.tomorrowDesc||'','','пропущений день'
        ]);
      });
      if (data.requestId) {
        const reqSheet=ss.getSheetByName(REQ_SHEET);
        if (reqSheet&&reqSheet.getLastRow()>1) {
          const reqRows=reqSheet.getDataRange().getValues();
          for (let i=1;i<reqRows.length;i++) {
            if (reqRows[i][0].toString()===data.requestId.toString()) {
              reqSheet.getRange(i+1,6).setValue('completed');
              reqSheet.getRange(i+1,7).setValue(Utilities.formatDate(new Date(),tz,'yyyy-MM-dd HH:mm'));
              break;
            }
          }
        }
      }
      _sendTelegram(TG_TOKEN,TG_CHAT_ID,
        '📋 *Пропущений день заповнено*\n'
        +'👤 '+_escMd(r0.employee||'')+'\n'
        +'📅 '+(r0.date||'')+'   🕐 '+(r0.arrival||'—')+' - '+(r0.departure||'—')+'\n'
        +'⏱ Відпрацював: '+(r0.hours||'')+'\n'
        +'Порушення: пропущений день');
      out.setContent(JSON.stringify({ok:true,type:'missedDayFull',added:rows.length}));
      return out;
    }

    // ── Підготовка листа табелю ───────────────────────────────────────────────
    let sheet=ss.getSheetByName(tabName);
    if (!sheet){sheet=ss.insertSheet(tabName);_addHeaders(sheet);}
    else if (sheet.getLastRow()===0){_addHeaders(sheet);}

    // ── arrival — фіксація приходу ────────────────────────────────────────────
    if (type==='arrival') {
      sheet.appendRow([
        data.date||'',data.month||'',data.employee||'',data.ip||'',data.location||'',data.arrival||'',
        '','','','','','','','','','','','','','',''   // G–U порожньо
      ]);
      _sendTelegram(TG_TOKEN,TG_CHAT_ID,
        '🟢 *'+_escMd(data.employee||'')+'* прийшов о *'+(data.arrival||'')+'*\n'
        +'📍 '+_escMd(data.location||'')+'   🌐 *'+(data.ip||'')+'*\n'
        +'📅 '+(data.date||''));
      out.setContent(JSON.stringify({ok:true,type:'arrival'}));
      return out;
    }

    // ── full — повний запис (відхід + проекти) ────────────────────────────────
    if (type==='full') {
      const rows=data.rows||[];
      if (!rows.length){out.setContent(JSON.stringify({ok:true,added:0}));return out;}
      const r0=rows[0];
      const isLate=!!data.isLate;

      // Шукаємо рядок-заглушку (прихід без відходу)
      let arrIdx=-1;
      if (sheet.getLastRow()>=2) {
        const sd=sheet.getDataRange().getValues();
        for (let i=1;i<sd.length;i++) {
          if (_fmtDate(sd[i][0],tz)===r0.date&&(sd[i][2]||'').toString().trim()===r0.employee&&(sd[i][6]||'').toString().trim()==='') {
            arrIdx=i+1; break;
          }
        }
      }

      const ipWL=_getWhiteIPs(ss);
      const makeRow=r=>[
        r.date||'',r.month||'',r.employee||'',_mergeIP(r.ip,r.ipDep),r.location||'',
        r.arrival||'',r.departure||'',r.lunchMin||'',r.lunchStr||'',
        r.hoursGross||'',r.hours||'',r.code||'',r.name||'',r.workType||'',
        r.timeSpent||'',r.section||'',r.desc||'',r.tomorrowPlan||'',r.tomorrowDesc||'',
        '',                     // T  Перевірка IP (заповнюється нижче)
        isLate?'пізнє внесення':''  // U  Порушення
      ];

      rows.forEach((r,idx)=>{
        let rowIndex;
        if (idx===0&&arrIdx>0) {
          sheet.getRange(arrIdx,1,1,21).setValues([makeRow(r)]);
          rowIndex=arrIdx;
        } else {
          sheet.appendRow(makeRow(r));
          rowIndex=sheet.getLastRow();
        }
        // IP-перевірка
        const checkIP=(r.ipDep&&r.ipDep!=='невідомо')?r.ipDep:r.ip;
        if (_isSuspiciousIP(checkIP,r.location,r.employee,ipWL)) {
          sheet.getRange(rowIndex,1,1,21).setBackground('#fff59d');
          sheet.getRange(rowIndex,20).setValue('⚠ IP не збігається');
        }
        // Telegram
        let msg=(isLate?'⚠️ *ПІЗНЄ ВНЕСЕННЯ*\n':'')
          +'🔴 *'+_escMd(r.employee||'')+'* пішов о *'+(r.departure||'')+'*\n'
          +'⏱ Відпрацював: *'+(r.hours||'')+'*'+(r.lunchStr?' (обід '+r.lunchStr+')':'')+'\n'
          +'🏗 *'+_escMd(r.code||'')+'*'+(r.name?' - '+_escMd(r.name):'')+'\n';
        if (r.timeSpent) msg+='⏰ На проект: *'+r.timeSpent+'*\n';
        if (r.workType)  msg+='⚒ '+_escMd(r.workType)+'\n';
        msg+='📅 '+(r.date||'');
        _sendTelegram(TG_TOKEN,TG_CHAT_ID,msg);
      });

      out.setContent(JSON.stringify({ok:true,type:'full',added:rows.length}));
      return out;
    }

    out.setContent(JSON.stringify({ok:false,error:'Невідомий тип: '+type}));
    return out;

  } catch(err) {
    out.setContent(JSON.stringify({ok:false,error:err.message}));
    return out;
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
//  Тригер — щоденне зведення о 20:00
// ═══════════════════════════════════════════════════════════════════════════════
function sendDailySummary() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_SHEET_DEFAULT);
  const tz    = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(),tz,'yyyy-MM-dd');
  if (!sheet||sheet.getLastRow()<2) {
    _sendTelegram(TG_TOKEN,TG_CHAT_ID,'📊 Зведення за '+today+': немає записів.');
    return;
  }
  const allData=sheet.getDataRange().getValues();
  const people={};
  for (let i=1;i<allData.length;i++) {
    const row=allData[i];
    if (_fmtDate(row[0],tz)!==today) continue;
    const emp=(row[2]||'').toString().trim();
    if (!emp) continue;
    if (!people[emp]) people[emp]={arrival:_fmtTime(row[5],tz),departure:'',hours:'',location:(row[4]||'').toString(),projects:[]};
    if (row[6]) people[emp].departure=_fmtTime(row[6],tz);
    if (row[10]) people[emp].hours=row[10].toString();
    if (row[11]) people[emp].projects.push((row[11]||'')+(row[12]?' - '+row[12]:''));
  }
  const names=Object.keys(people);
  if (!names.length){_sendTelegram(TG_TOKEN,TG_CHAT_ID,'📊 Зведення за '+today+': ніхто не відмітився.');return;}
  let msg='📊 *Зведення за '+today+'*\n👥 Відмітилось: '+names.length+' осіб\n\n';
  names.forEach(emp=>{
    const p=people[emp];
    msg+='👤 *'+_escMd(emp)+'*\n'
      +'   🕐 '+p.arrival+' → '+(p.departure||'(ще не пішов)'+(p.hours?'   ⏱ '+p.hours:''))+'\n';
    if (p.projects.length) msg+='   🏗 '+p.projects.slice(0,3).map(_escMd).join('; ')+'\n';
    msg+='\n';
  });
  _sendTelegram(TG_TOKEN,TG_CHAT_ID,msg.trim());
}


// ─── Тести та діагностика ─────────────────────────────────────────────────────
function testTelegram() {
  const res = _sendTelegramDebug(TG_TOKEN, TG_CHAT_ID, '✅ Тест з Apps Script v26.05.10 — працює!');
  console.log('Telegram result:', JSON.stringify(res));
}

function testOTP() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const code = String(Math.floor(100000+Math.random()*900000));
  let sh = ss.getSheetByName(OTP_SHEET);
  if (!sh) {
    sh = ss.insertSheet(OTP_SHEET);
    sh.appendRow(['Співробітник','Дата дня','Код','Час створення','Статус']);
    sh.getRange(1,1,1,5).setBackground('#1a237e').setFontColor('#fff').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.appendRow(['TEST_USER','2026-01-01',code,new Date(),'pending']);
  const msg = '🔑 Тест OTP\nКод: *'+code+'*\n_(дійсний '+OTP_TTL_MIN+' хв)_';
  const res = _sendTelegramDebug(TG_TOKEN, TG_CHAT_ID, msg);
  console.log('OTP test:', code, '| Telegram:', JSON.stringify(res));
}

function testAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log('=== testAll() ===');
  console.log('Spreadsheet:', ss.getName());
  const sheets = ss.getSheets().map(s=>s.getName());
  console.log('Sheets:', sheets.join(', '));
  const tg = _sendTelegramDebug(TG_TOKEN, TG_CHAT_ID,
    '🔧 testAll() пройшов\n📊 Таблиця: '+ss.getName()+'\n📋 Листи: '+sheets.join(', '));
  console.log('Telegram ok:', tg.ok, tg.ok?'':'| error: '+tg.description);
  if (!tg.ok) throw new Error('Telegram: '+tg.error_code+' '+tg.description);
}

function _sendTelegramDebug(token, chatId, text) {
  try {
    const res = UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/sendMessage', {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({chat_id:chatId, text, parse_mode:'Markdown'}),
      muteHttpExceptions:true
    });
    return JSON.parse(res.getContentText());
  } catch(e) {
    return {ok:false, description:e.message};
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
//  Приватні допоміжні функції
// ═══════════════════════════════════════════════════════════════════════════════
function _fmtDate(raw,tz){
  if (!raw) return '';
  if (raw instanceof Date) return Utilities.formatDate(raw,tz,'yyyy-MM-dd');
  return raw.toString().slice(0,10);
}
function _fmtTime(raw,tz){
  if (!raw) return '';
  if (raw instanceof Date) return Utilities.formatDate(raw,tz,'HH:mm');
  return raw.toString().trim();
}
function _readCol(ss,sheetName,col){
  const sht=ss.getSheetByName(sheetName);
  if (!sht||sht.getLastRow()<2) return [];
  return sht.getRange(col+'2:'+col+sht.getLastRow()).getValues().flat();
}
function _mergeIP(ipArr,ipDep){
  const a=(ipArr||'').toString().trim();
  const d=(ipDep||'').toString().trim();
  if (!d||d==='невідомо') return a;
  return a+' / '+d;
}
function _getWhiteIPs(ss){
  const sheet=ss.getSheetByName(IP_WHITELIST_SHEET);
  if (!sheet||sheet.getLastRow()<2) return [];
  return sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues()
    .filter(r=>r[0])
    .map(r=>({ip:r[0].toString().trim(),location:r[1].toString().trim(),employee:r[2].toString().trim()||'Всі'}));
}
function _isSuspiciousIP(ip,location,employee,whitelist){
  if (!whitelist||!whitelist.length) return false;
  const matches=whitelist.filter(e=>e.ip===ip);
  if (!matches.length) return true;
  return !matches.some(e=>e.location===location&&(e.employee==='Всі'||e.employee===employee));
}
function _sendTelegram(token,chatId,text){
  try {
    const res=UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/sendMessage',{
      method:'post',contentType:'application/json',
      payload:JSON.stringify({chat_id:chatId,text,parse_mode:'Markdown'}),
      muteHttpExceptions:true
    });
    const r=JSON.parse(res.getContentText());
    if (!r.ok) console.error('Telegram ['+r.error_code+']: '+r.description);
  } catch(e){ console.error('Telegram exception: '+e.message); }
}
function _escMd(s){
  return (s||'').replace(/([_*`\[])/g,'\\$1');
}
function _addHeaders(sheet){
  const H=[
    'Дата','Місяць','Співробітник','IP-адреса','Місце роботи','Прихід','Відхід',
    'Обід (хв)','Обід (текст)','Загальний час','Чистий роб. час',
    'Шифр',"Назва об'єкта",'Вид робіт','Витрачений час','Розділ','Опис робіт',
    'Плани на завтра','Причина (завтра)','Перевірка IP','Порушення'
  ];
  sheet.appendRow(H);
  const hdr=sheet.getRange(1,1,1,H.length);
  hdr.setBackground('#1a237e').setFontColor('#fff').setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  [95,90,130,200,100,70,70,70,85,95,95,130,270,160,80,200,260,110,220,150,140]
    .forEach((w,i)=>sheet.setColumnWidth(i+1,w));
}
