/**
 * 南瓜節慶帳本 · Apps Script 後端 v1.2（2026-09-25）
 * v1.2：紀錄加 photoUrl；uploadPhoto 把禮盒照片存到 Drive「節慶帳本_照片」資料夾（連結可看），回傳縮圖網址
 * 功能：① 首次 setup() 自動建 5 個分頁＋每日 09:00 觸發
 *       ② doGet / doPost 給 index.html 讀寫資料（JSON）
 *       ③ dailyCheck() 每天檢查 14 天內的節慶／生日／週年 → Email、Telegram、Google Calendar
 * 秘密（Telegram Token／Chat ID）放「專案設定 → 指令碼屬性」：TG_TOKEN、TG_CHAT。不要寫進 Sheet。
 */

var SHEETS = {
  records:   { name: '紀錄',     headers: ['id','year','festival','date','targetType','targetName','itemType','itemName','unitPrice','qty','subtotal','vendor','status','note','photoUrl','updatedAt','updatedBy'] },
  festivals: { name: '節慶',     headers: ['name','type','d2026','d2027','d2028','d2029','d2030','leadDays','enabled','note'] },
  people:    { name: '人員',     headers: ['name','type','org','title','role','birthday','joinDate','email','receivesGift','note'] },
  settings:  { name: '設定',     headers: ['key','value','說明'] },
  log:       { name: '提醒紀錄', headers: ['key','sentAt','channels','result'] }
};

var DEFAULT_FESTIVALS = [
  ['春節（年終／紅包／尾牙春酒）','農曆','2026-02-17','2027-02-06','2028-01-26','2029-02-13','2030-02-03',14,true,'尾牙通常在春節前 2～4 週，另建活動紀錄'],
  ['端午','農曆','2026-06-19','2027-06-09','2028-05-28','2029-06-16','2030-06-05',14,true,''],
  ['中秋','農曆','2026-09-25','2027-09-15','2028-10-03','2029-09-22','2030-09-12',14,true,''],
  ['聖誕／跨年','國曆','2026-12-25','2027-12-25','2028-12-25','2029-12-25','2030-12-25',14,true,'交換禮物、年末活動'],
  ['員工生日','個人','','','','','',14,true,'由「人員」分頁的生日展開'],
  ['到職週年','個人','','','','','',14,true,'由「人員」分頁的到職日展開']
];

var DEFAULT_SETTINGS = [
  ['recipients','you@example.com','提醒收件人 Email，多個用逗號分隔'],
  ['calendarId','primary','要建立事件的 Google Calendar ID；primary＝自己的主日曆'],
  ['leadDays','14','預設提前幾天提醒'],
  ['yearlyBudget','0','年度節慶總預算上限（0＝不限制）'],
  ['budgetEmployee','800-1500','員工每人禮盒預算帶（元）'],
  ['budgetPartner','1500-3000','合作夥伴每份預算帶（元）'],
  ['budgetVendor','1000-2000','廠商每份預算帶（元）'],
  ['companyName','南瓜虛擬科技','寄信署名'],
  ['toolUrl','','index.html 放在 Drive 的連結（寫在提醒信裡）']
];

/* ───────── 初始化 ───────── */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (k) {
    var def = SHEETS[k];
    var sh = ss.getSheetByName(def.name);
    if (!sh) sh = ss.insertSheet(def.name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
      if (k === 'festivals') sh.getRange(2, 1, DEFAULT_FESTIVALS.length, def.headers.length).setValues(DEFAULT_FESTIVALS);
      if (k === 'settings') sh.getRange(2, 1, DEFAULT_SETTINGS.length, def.headers.length).setValues(DEFAULT_SETTINGS);
    }
  });
  var first = ss.getSheets()[0];
  if (first.getName() === '工作表1' && first.getLastRow() === 0) ss.deleteSheet(first);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyCheck') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyCheck').timeBased().everyDays(1).atHour(9).create();
  Logger.log('✅ setup 完成：5 個分頁＋每日 09:00 觸發已建立');
}

/* ───────── 讀寫工具 ───────── */
function sheet_(k) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEETS[k].name) || (setup(), ss.getSheetByName(SHEETS[k].name));
}
function readAll_(k) {
  var sh = sheet_(k), h = SHEETS[k].headers, last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, h.length).getValues();
  return rows.filter(function (r) { return r.join('') !== ''; }).map(function (r) {
    var o = {}; h.forEach(function (key, i) { o[key] = fmt_(r[i]); }); return o;
  });
}
function fmt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v;
}
function writeAll_(k, objs) {
  var sh = sheet_(k), h = SHEETS[k].headers;
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, h.length).clearContent();
  if (!objs.length) return;
  sh.getRange(2, 1, objs.length, h.length).setValues(objs.map(function (o) { return h.map(function (key) { return o[key] == null ? '' : o[key]; }); }));
}
function settings_() {
  var o = {}; readAll_('settings').forEach(function (r) { o[r.key] = String(r.value); }); return o;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function payloadAll_() {
  return { ok: true, records: readAll_('records'), festivals: readAll_('festivals'), people: readAll_('people'), settings: settings_(), log: readAll_('log').slice(-50), serverTime: new Date().toISOString() };
}

/* ───────── Web API ───────── */
function doGet(e) {
  var a = (e && e.parameter && e.parameter.action) || 'all';
  if (a === 'ping') return json_({ ok: true, name: '南瓜節慶帳本', version: '1.0', time: new Date().toISOString() });
  return json_(payloadAll_());
}
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse(e.postData.contents || '{}');
    var a = body.action, p = body.payload || {};
    var who = body.user || '';
    if (a === 'upsertRecords') {            // p = [record,...]
      var recs = readAll_('records'), byId = {};
      recs.forEach(function (r, i) { byId[r.id] = i; });
      p.forEach(function (r) {
        r.subtotal = (Number(r.unitPrice) || 0) * (Number(r.qty) || 0);
        r.updatedAt = new Date().toISOString(); r.updatedBy = who;
        if (byId[r.id] != null) recs[byId[r.id]] = r; else { recs.push(r); byId[r.id] = recs.length - 1; }
      });
      writeAll_('records', recs);
    } else if (a === 'deleteRecords') {     // p = [id,...]
      writeAll_('records', readAll_('records').filter(function (r) { return p.indexOf(r.id) < 0; }));
    } else if (a === 'replacePeople') {     // p = [person,...]
      writeAll_('people', p);
    } else if (a === 'replaceFestivals') {  // p = [festival,...]
      writeAll_('festivals', p);
    } else if (a === 'saveSettings') {      // p = {key:value}
      var cur = readAll_('settings');
      Object.keys(p).forEach(function (k) {
        var hit = cur.filter(function (r) { return r.key === k; })[0];
        if (hit) hit.value = p[k]; else cur.push({ key: k, value: p[k], '說明': '' });
      });
      writeAll_('settings', cur);
    } else if (a === 'uploadPhoto') {       // p = {id, name, dataUrl}
      var m = String(p.dataUrl).match(/^data:(image\/[a-z]+);base64,(.+)$/); if (!m) return json_({ ok: false, error: '不是圖片' });
      var folder = photoFolder_();
      var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (p.id || 'photo') + '_' + (p.name || 'gift') ));
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
      var recs2 = readAll_('records'); recs2.forEach(function (r) { if (r.id === p.id) { r.photoUrl = url; r.updatedAt = new Date().toISOString(); r.updatedBy = who; } }); writeAll_('records', recs2);
      return json_({ ok: true, url: url, fileUrl: file.getUrl() });
    } else if (a === 'testReminder') {      // p = {name, date}
      var res = sendReminder_(buildContext_(p.name, p.date, p.kind || '節慶', p.who || ''), true);
      return json_({ ok: true, result: res });
    } else {
      return json_({ ok: false, error: '未知的 action：' + a });
    }
    return json_(payloadAll_());
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function photoFolder_() {
  var it = DriveApp.getFoldersByName('節慶帳本_照片');
  return it.hasNext() ? it.next() : DriveApp.createFolder('節慶帳本_照片');
}

/* ───────── 提醒 ───────── */
function todayStr_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function daysBetween_(a, b) { return Math.round((new Date(b) - new Date(a)) / 86400000); }
function thisYearDate_(mmdd, year) { return year + '-' + mmdd.slice(5); }

/** 找出未來 leadDays 內要提醒的事件清單 */
function upcoming_(leadDefault) {
  var today = todayStr_(), y = Number(today.slice(0, 4)), out = [];
  readAll_('festivals').forEach(function (f) {
    if (String(f.enabled) === 'false' || f.enabled === '' ) return;
    var lead = Number(f.leadDays) || leadDefault;
    if (f.type === '個人') {
      readAll_('people').forEach(function (p) {
        if (String(p.receivesGift) === 'false') return;
        var src = f.name.indexOf('生日') >= 0 ? p.birthday : p.joinDate;
        if (!src) return;
        [y, y + 1].forEach(function (yy) {
          var d = thisYearDate_(String(src), yy), diff = daysBetween_(today, d);
          if (diff >= 0 && diff <= lead) out.push({ name: f.name, date: d, days: diff, kind: '個人', who: p.name, key: f.name + '|' + p.name + '|' + yy });
        });
      });
      return;
    }
    [y, y + 1].forEach(function (yy) {
      var d = f['d' + yy]; if (!d) return;
      var diff = daysBetween_(today, d);
      if (diff >= 0 && diff <= lead) out.push({ name: f.name, date: d, days: diff, kind: '節慶', key: f.name + '|' + yy });
    });
  });
  return out;
}

function buildContext_(name, date, kind, who) {
  var s = settings_(), recs = readAll_('records');
  var y = Number(String(date).slice(0, 4));
  var sum = function (arr) { return arr.reduce(function (t, r) { return t + (Number(r.subtotal) || 0); }, 0); };
  var same = function (r, yy) { return Number(r.year) === yy && String(r.festival).slice(0, 2) === String(name).slice(0, 2); };
  var last = recs.filter(function (r) { return same(r, y - 1); });
  var cur = recs.filter(function (r) { return same(r, y); });
  var todo = ['選定禮盒與預算（看「選禮顧問」）', '確認名單與數量（員工／夥伴／廠商）', '下單並確認到貨日', '賀卡／明信片與發放安排'];
  return { name: name, date: date, days: daysBetween_(todayStr_(), date), kind: kind, who: who, settings: s,
    lastYearTotal: sum(last), lastYearItems: last.map(function (r) { return r.itemName + '×' + r.qty + '（' + r.targetType + '）'; }).slice(0, 8),
    thisYearTotal: sum(cur), thisYearCount: cur.length, thisYearDone: cur.filter(function (r) { return r.status === '已發放'; }).length, todo: todo };
}

function messageText_(c) {
  var title = c.kind === '個人' ? (c.who + ' 的' + c.name) : c.name;
  var lines = [
    '🎁 【' + c.settings.companyName + '】' + title + ' 還有 ' + c.days + ' 天（' + c.date + '）',
    '',
    '去年：' + (c.lastYearTotal ? 'NT$ ' + c.lastYearTotal.toLocaleString() + '｜' + c.lastYearItems.join('、') : '沒有紀錄'),
    '今年：' + (c.thisYearCount ? 'NT$ ' + c.thisYearTotal.toLocaleString() + '｜' + c.thisYearCount + ' 筆，' + c.thisYearDone + ' 筆已發放' : '尚未規劃'),
    '',
    '建議現在開始：'
  ].concat(c.todo.map(function (t, i) { return (i + 1) + '. ' + t; }));
  if (c.settings.toolUrl) lines.push('', '打開帳本：' + c.settings.toolUrl);
  return lines.join('\n');
}

function sendReminder_(c, force) {
  var s = c.settings, results = [];
  var text = messageText_(c), subject = '【節慶提醒】' + (c.kind === '個人' ? c.who + ' 的' + c.name : c.name) + ' 還有 ' + c.days + ' 天';
  // Email
  try {
    var to = (s.recipients || '').split(',').map(function (x) { return x.trim(); }).filter(String);
    if (to.length) { GmailApp.sendEmail(to.join(','), subject, text, { name: s.companyName + ' 節慶帳本' }); results.push('email:ok'); }
    else results.push('email:無收件人');
  } catch (e) { results.push('email:' + e); }
  // Telegram
  try {
    var props = PropertiesService.getScriptProperties(), tk = props.getProperty('TG_TOKEN'), chat = props.getProperty('TG_CHAT');
    if (tk && chat) {
      UrlFetchApp.fetch('https://api.telegram.org/bot' + tk + '/sendMessage', { method: 'post', payload: { chat_id: chat, text: text }, muteHttpExceptions: true });
      results.push('telegram:ok');
    } else results.push('telegram:未設定');
  } catch (e) { results.push('telegram:' + e); }
  // Calendar（節慶當天＋準備日，全天事件）
  try {
    var cal = (s.calendarId && s.calendarId !== 'primary') ? CalendarApp.getCalendarById(s.calendarId) : CalendarApp.getDefaultCalendar();
    if (cal) {
      var d = new Date(c.date + 'T00:00:00'), prep = new Date(d); prep.setDate(prep.getDate() - (Number(s.leadDays) || 14));
      var t1 = '🎁 ' + (c.kind === '個人' ? c.who + ' ' + c.name : c.name);
      if (!cal.getEventsForDay(d, { search: t1 }).length) cal.createAllDayEvent(t1, d, { description: text }).addPopupReminder(9 * 60);
      var t2 = '🛒 準備：' + t1.slice(2);
      if (prep >= new Date() && !cal.getEventsForDay(prep, { search: t2 }).length) cal.createAllDayEvent(t2, prep, { description: text }).addPopupReminder(9 * 60);
      results.push('calendar:ok');
    }
  } catch (e) { results.push('calendar:' + e); }
  if (!force) sheet_('log').appendRow([c.key || (c.name + '|' + c.date), new Date().toISOString(), 'email,telegram,calendar', results.join(' ')]);
  return results;
}

/** 每日 09:00 觸發 */
function dailyCheck() {
  var s = settings_(), lead = Number(s.leadDays) || 14;
  var sent = {}; readAll_('log').forEach(function (r) { sent[r.key] = true; });
  var list = upcoming_(lead).filter(function (u) { return !sent[u.key]; });
  list.forEach(function (u) {
    var c = buildContext_(u.name, u.date, u.kind, u.who); c.key = u.key;
    Logger.log(u.key + ' → ' + sendReminder_(c, false).join(' '));
  });
  Logger.log('dailyCheck：檢查 ' + list.length + ' 件');
  return list.length;
}

/** 在編輯器手動跑：強制寄一封中秋測試信（不寫入提醒紀錄） */
function testReminderNow() {
  var c = buildContext_('中秋', todayStr_(), '節慶', '');
  Logger.log(sendReminder_(c, true));
}
