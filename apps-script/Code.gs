/**
 * 南瓜節慶帳本 · Apps Script 後端 v2.0（2026-09-30）
 * 這一版把整個工具搬進 Apps Script 網頁應用：
 *   ① Google 帳號登入＋只有白名單信箱能開（ALLOWED_EMAILS）
 *   ② 資料存本 Google Sheet；照片、歸檔 Sheet／PDF 存到 Drive「南瓜節慶帳本」資料夾（年份／節慶 分層）
 *   ③ Claude 金鑰放「指令碼屬性」CLAUDE_API_KEY，網頁上不出現
 *   ④ 每日 09:00：提醒（Email／Telegram／Calendar）＋自動歸檔已結束的節慶
 * 指令碼屬性（專案設定 → 指令碼屬性）：
 *   ALLOWED_EMAILS  逗號分隔白名單（不填就用下面 DEFAULT_ALLOWED）
 *   DRIVE_FOLDER_ID Drive 根資料夾 ID（不填就用 DEFAULT_FOLDER）
 *   CLAUDE_API_KEY  sk-ant-…（賀圖 AI 生成用）
 *   CLAUDE_MODEL    預設 claude-opus-5
 *   TG_TOKEN / TG_CHAT  Telegram（選填）
 */

var DEFAULT_ALLOWED = ['you@example.com'];
var DEFAULT_FOLDER  = '';   // 我的雲端硬碟／財務資料／南瓜節慶帳本（Festival Gift Ledger）

var SHEETS = {
  records:   { name: '紀錄',     headers: ['id','year','festival','date','targetType','targetName','itemType','itemName','unitPrice','qty','subtotal','vendor','status','note','photoUrl','updatedAt','updatedBy'] },
  festivals: { name: '節慶',     headers: ['name','type','d2026','d2027','d2028','d2029','d2030','leadDays','enabled','note'] },
  people:    { name: '人員',     headers: ['name','type','org','title','role','birthday','joinDate','email','receivesGift','note'] },
  settings:  { name: '設定',     headers: ['key','value','說明'] },
  log:       { name: '提醒紀錄', headers: ['key','sentAt','channels','result'] },
  archives:  { name: '歸檔紀錄', headers: ['key','year','festival','archivedAt','by','count','total','folderUrl','sheetUrl','pdfUrl'] }
};
var DEFAULT_FESTIVALS = [
  ['春節（年終／紅包／尾牙春酒）','農曆','2026-02-17','2027-02-06','2028-01-26','2029-02-13','2030-02-03',14,true,'尾牙通常在春節前 2～4 週'],
  ['端午','農曆','2026-06-19','2027-06-09','2028-05-28','2029-06-16','2030-06-05',14,true,''],
  ['中秋','農曆','2026-09-25','2027-09-15','2028-10-03','2029-09-22','2030-09-12',14,true,''],
  ['聖誕／跨年','國曆','2026-12-25','2027-12-25','2028-12-25','2029-12-25','2030-12-25',14,true,'交換禮物、年末活動'],
  ['員工生日','個人','','','','','',14,true,'由「人員」的生日展開'],
  ['到職週年','個人','','','','','',14,true,'由「人員」的到職日展開']
];
var DEFAULT_SETTINGS = [
  ['recipients','you@example.com','提醒收件人 Email，多個用逗號分隔'],
  ['calendarId','primary','建立事件的 Google Calendar ID；primary＝自己的主日曆'],
  ['leadDays','14','預設提前幾天提醒'],
  ['yearlyBudget','0','年度節慶總預算上限（0＝不限制）'],
  ['budgetEmployee','800-1500','員工每人禮盒預算帶（元）'],
  ['budgetPartner','1500-3000','合作夥伴每份預算帶（元）'],
  ['budgetVendor','1000-2000','廠商每份預算帶（元）'],
  ['companyName','南瓜虛擬科技','寄信署名'],
  ['toolUrl','','網頁應用程式網址（寫在提醒信裡）']
];

/* ───────── 屬性與權限 ───────── */
function prop_(k, def) { var v = PropertiesService.getScriptProperties().getProperty(k); return (v == null || v === '') ? def : v; }
function allowed_() { return prop_('ALLOWED_EMAILS', DEFAULT_ALLOWED.join(',')).split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(String); }
function me_() { return (Session.getActiveUser().getEmail() || '').toLowerCase(); }
function assertAllowed_() { var e = me_(); if (allowed_().indexOf(e) < 0) throw new Error('這個帳號（' + (e || '未登入') + '）沒有權限使用節慶帳本'); return e; }

/* ───────── 初始化 ───────── */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (k) {
    var def = SHEETS[k], sh = ss.getSheetByName(def.name);
    if (!sh) sh = ss.insertSheet(def.name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold'); sh.setFrozenRows(1);
      if (k === 'festivals') sh.getRange(2, 1, DEFAULT_FESTIVALS.length, def.headers.length).setValues(DEFAULT_FESTIVALS);
      if (k === 'settings') sh.getRange(2, 1, DEFAULT_SETTINGS.length, def.headers.length).setValues(DEFAULT_SETTINGS);
    } else {
      // 補新欄位（升級用）
      var cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
      def.headers.forEach(function (h, i) { if (cur.indexOf(h) < 0) sh.getRange(1, cur.length + 1).setValue(h).setFontWeight('bold'), cur.push(h); });
    }
  });
  var first = ss.getSheets()[0]; if (first.getName() === '工作表1' && first.getLastRow() === 0) ss.deleteSheet(first);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (['dailyCheck', 'daily'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('daily').timeBased().everyDays(1).atHour(9).create();
  // 把這份 Sheet 搬進 Drive 根資料夾（已在裡面就略過）
  try { var f = DriveApp.getFileById(ss.getId()); var root = rootFolder_(); var inside = false; var ps = f.getParents(); while (ps.hasNext()) if (ps.next().getId() === root.getId()) inside = true; if (!inside) f.moveTo(root); } catch (e) { Logger.log('搬移 Sheet 略過：' + e); }
  rootSub_('匯出'); 
  Logger.log('✅ setup 完成：6 個分頁、每日 09:00 觸發、Drive 資料夾已就緒。白名單：' + allowed_().join(', '));
}

/* ───────── 讀寫工具 ───────── */
function sheet_(k) { var ss = SpreadsheetApp.getActiveSpreadsheet(); return ss.getSheetByName(SHEETS[k].name) || (setup(), ss.getSheetByName(SHEETS[k].name)); }
function headerIndex_(sh, headers) { var cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]; return headers.map(function (h) { return cur.indexOf(h); }); }
function readAll_(k) {
  var sh = sheet_(k), h = SHEETS[k].headers, last = sh.getLastRow(); if (last < 2) return [];
  var idx = headerIndex_(sh, h), rows = sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), 1)).getValues();
  return rows.filter(function (r) { return r.join('') !== ''; }).map(function (r) { var o = {}; h.forEach(function (key, i) { o[key] = idx[i] < 0 ? '' : fmt_(r[idx[i]]); }); return o; });
}
function fmt_(v) { if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd'); return v; }
function writeAll_(k, objs) {
  var sh = sheet_(k), h = SHEETS[k].headers, idx = headerIndex_(sh, h), width = Math.max(sh.getLastColumn(), h.length);
  var last = sh.getLastRow(); if (last > 1) sh.getRange(2, 1, last - 1, width).clearContent();
  if (!objs.length) return;
  sh.getRange(2, 1, objs.length, width).setValues(objs.map(function (o) { var row = []; for (var c = 0; c < width; c++) row.push(''); h.forEach(function (key, i) { if (idx[i] >= 0) row[idx[i]] = o[key] == null ? '' : o[key]; }); return row; }));
}
function settings_() { var o = {}; readAll_('settings').forEach(function (r) { o[r.key] = String(r.value); }); return o; }
function payloadAll_() {
  return { ok: true, user: me_(), records: readAll_('records'), festivals: readAll_('festivals'), people: readAll_('people'), settings: settings_(), archives: readAll_('archives'), log: readAll_('log').slice(-50), serverTime: new Date().toISOString(), hasClaude: !!prop_('CLAUDE_API_KEY', ''), rootUrl: rootFolder_().getUrl() };
}
function festShort_(n) { return String(n || '').split('（')[0]; }

/* ───────── 網頁入口 ───────── */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {                 // 舊版 JSON API（本機模式用）
    if (e.parameter.action === 'ping') return json_({ ok: true, name: '南瓜節慶帳本', version: '2.0' });
    return json_(payloadAll_());
  }
  var email = me_();
  if (allowed_().indexOf(email) < 0) {
    return HtmlService.createHtmlOutput('<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>南瓜節慶帳本</title></head><body style="font-family:-apple-system,\'Noto Sans TC\',sans-serif;background:#F5F5F5;color:#161415;display:flex;min-height:100vh;align-items:center;justify-content:center;margin:0"><div style="background:#fff;border:1px solid rgba(22,20,21,.08);border-radius:18px;padding:32px;max-width:420px;line-height:1.7"><div style="font-size:22px;font-weight:600;margin-bottom:8px">🎁 南瓜節慶帳本</div><p>目前登入：<b>' + (email || '（未登入公司帳號）') + '</b></p><p>這個帳號不在名單裡。請用公司信箱登入，或請南瓜把你加進名單。</p><p><a href="https://accounts.google.com/AccountChooser" style="color:#8A6400">切換 Google 帳號</a></p></div></body></html>').setTitle('南瓜節慶帳本');
  }
  var t = HtmlService.createTemplateFromFile('index'); t.userEmail = email;
  return t.evaluate().setTitle('南瓜節慶帳本').setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function doPost(e) {                                            // 舊版 JSON API（本機模式用）
  try { var body = JSON.parse(e.postData.contents || '{}'); return json_(api(body.action, body.payload, body.user || '')); }
  catch (err) { return json_({ ok: false, error: String(err) }); }
}

/* ───────── 統一 API（網頁用 google.script.run.api(action, payload)） ───────── */
function api(action, payload, legacyUser) {
  var who = legacyUser || me_();
  if (!legacyUser) assertAllowed_();
  var p = payload || {};
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    switch (action) {
      case 'all': return payloadAll_();
      case 'upsertRecords': {
        var recs = readAll_('records'), byId = {}; recs.forEach(function (r, i) { byId[r.id] = i; });
        p.forEach(function (r) { r.subtotal = (Number(r.unitPrice) || 0) * (Number(r.qty) || 0); r.updatedAt = new Date().toISOString(); r.updatedBy = who; delete r.photo; if (byId[r.id] != null) { r.photoUrl = r.photoUrl || recs[byId[r.id]].photoUrl; recs[byId[r.id]] = r; } else { recs.push(r); byId[r.id] = recs.length - 1; } });
        writeAll_('records', recs); return payloadAll_();
      }
      case 'deleteRecords': writeAll_('records', readAll_('records').filter(function (r) { return p.indexOf(r.id) < 0; })); return payloadAll_();
      case 'replacePeople': writeAll_('people', p); return payloadAll_();
      case 'replaceFestivals': writeAll_('festivals', p); return payloadAll_();
      case 'saveSettings': { var cur = readAll_('settings'); Object.keys(p).forEach(function (k) { var hit = cur.filter(function (r) { return r.key === k; })[0]; if (hit) hit.value = p[k]; else cur.push({ key: k, value: p[k], '說明': '' }); }); writeAll_('settings', cur); return payloadAll_(); }
      case 'uploadPhoto': return uploadPhoto_(p, who);
      case 'claude': return claude_(p);
      case 'exportFile': { var f = rootSub_('匯出').createFile(Utilities.newBlob(p.text, p.mime || 'text/plain', p.name)); return { ok: true, url: f.getUrl(), name: p.name }; }
      case 'archiveFestival': return archiveFestival_(Number(p.year), p.festival, who);
      case 'testReminder': return { ok: true, result: sendReminder_(buildContext_(p.name || '中秋', p.date || todayStr_(), '節慶', ''), true) };
      default: return { ok: false, error: '未知的 action：' + action };
    }
  } finally { try { lock.releaseLock(); } catch (_) {} }
}

/* ───────── Drive 資料夾結構：根／年份／節慶／照片 ───────── */
function rootFolder_() { return DriveApp.getFolderById(prop_('DRIVE_FOLDER_ID', DEFAULT_FOLDER)); }
function sub_(parent, name) { var it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function rootSub_(name) { return sub_(rootFolder_(), name); }
function festFolder_(year, fest) { return sub_(sub_(rootFolder_(), String(year)), festShort_(fest)); }
function replaceFile_(folder, name, blob) { var it = folder.getFilesByName(name); while (it.hasNext()) it.next().setTrashed(true); return folder.createFile(blob.setName(name)); }

function uploadPhoto_(p, who) {
  var m = String(p.dataUrl).match(/^data:(image\/[a-z]+);base64,(.+)$/); if (!m) return { ok: false, error: '不是圖片' };
  var folder = sub_(festFolder_(p.year || new Date().getFullYear(), p.festival || '未分類'), '照片');
  var ext = m[1].split('/')[1] === 'png' ? 'png' : 'jpg';
  var name = [p.targetName, p.itemName].filter(String).join('_').replace(/[\\/:*?"<>|]/g, '') || 'photo';
  var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name + '_' + (p.id || '') + '.' + ext));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
  var recs = readAll_('records'); recs.forEach(function (r) { if (r.id === p.id) { r.photoUrl = url; r.updatedAt = new Date().toISOString(); r.updatedBy = who; } }); writeAll_('records', recs);
  return { ok: true, url: url, fileUrl: file.getUrl(), folderUrl: folder.getUrl() };
}

/* ───────── 歸檔：一鍵把某年某節慶做成 Sheet＋PDF＋照片 ───────── */
function archiveFestival_(year, fest, who) {
  fest = festShort_(fest);
  var recs = readAll_('records').filter(function (r) { return Number(r.year) === year && festShort_(r.festival) === fest; });
  if (!recs.length) return { ok: false, error: year + ' ' + fest + ' 沒有紀錄可歸檔' };
  var folder = festFolder_(year, fest), s = settings_();
  var name = year + '_' + fest + '_紀錄';
  // 1) 快照 Sheet
  var it = folder.getFilesByName(name); while (it.hasNext()) it.next().setTrashed(true);
  var ss = SpreadsheetApp.create(name); DriveApp.getFileById(ss.getId()).moveTo(folder);
  var sh = ss.getSheets()[0]; sh.setName('紀錄');
  var total = recs.reduce(function (t, r) { return t + (Number(r.subtotal) || 0); }, 0);
  var head = [[s.companyName + '　' + year + ' 年 ' + fest + '　節慶紀錄'], ['歸檔時間：' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + '　歸檔者：' + who + '　共 ' + recs.length + ' 筆　合計 NT$ ' + total.toLocaleString()], ['']];
  var cols = ['日期', '對象類型', '對象名稱', '項目', '品名／內容', '單價', '數量', '小計', '供應商／連結', '狀態', '備註', '照片'];
  var rows = recs.sort(function (a, b) { return String(a.date).localeCompare(String(b.date)) || String(a.targetType).localeCompare(String(b.targetType)); }).map(function (r) { return [r.date, r.targetType, r.targetName, r.itemType, r.itemName, Number(r.unitPrice) || 0, Number(r.qty) || 0, Number(r.subtotal) || 0, r.vendor, r.status, r.note, r.photoUrl ? r.photoUrl.replace('/thumbnail?id=', '/file/d/').replace(/&sz=.*/, '/view') : '']; });
  var byType = {}, byTarget = {}; recs.forEach(function (r) { byType[r.itemType] = (byType[r.itemType] || 0) + (Number(r.subtotal) || 0); byTarget[r.targetType] = (byTarget[r.targetType] || 0) + (Number(r.subtotal) || 0); });
  var summary = [[''], ['合計', '', '', '', '', '', recs.reduce(function (t, r) { return t + (Number(r.qty) || 0); }, 0), total], ['']].concat(Object.keys(byType).map(function (k) { return ['依項目', k, '', '', '', '', '', byType[k]]; }), [['']], Object.keys(byTarget).map(function (k) { return ['依對象', k, '', '', '', '', '', byTarget[k]]; }));
  var all = head.concat([cols], rows, summary);
  var width = cols.length; all = all.map(function (r) { while (r.length < width) r.push(''); return r; });
  sh.getRange(1, 1, all.length, width).setValues(all);
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold'); sh.getRange(2, 1).setFontColor('#666666');
  sh.getRange(4, 1, 1, width).setFontWeight('bold').setBackground('#F5F5F5'); sh.setFrozenRows(4);
  sh.getRange(5, 6, rows.length + summary.length, 3).setNumberFormat('#,##0');
  sh.getRange(5 + rows.length + 1, 1, 1, width).setFontWeight('bold');
  [11, 10, 14, 8, 30, 9, 7, 11, 16, 8, 24, 40].forEach(function (w, i) { sh.setColumnWidth(i + 1, w * 8); });
  // 2) PDF（用 Sheet 匯出，中文字型正常）
  var pdfUrl = '';
  try {
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf&portrait=false&size=A4&fitw=true&gridlines=false&sheetnames=false&printtitle=false&pagenumbers=true&gid=' + sh.getSheetId();
    var blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
    pdfUrl = replaceFile_(folder, name + '.pdf', blob).getUrl();
  } catch (e) { pdfUrl = '（PDF 失敗：' + e + '）'; }
  // 3) 照片整理到 照片/（上傳時已在裡面；這裡補命名對照表）
  var photos = recs.filter(function (r) { return r.photoUrl; });
  if (photos.length) sub_(folder, '照片');
  // 4) 記錄
  var key = year + '|' + fest, arch = readAll_('archives').filter(function (a) { return a.key !== key; });
  var row = { key: key, year: year, festival: fest, archivedAt: new Date().toISOString(), by: who, count: recs.length, total: total, folderUrl: folder.getUrl(), sheetUrl: ss.getUrl(), pdfUrl: pdfUrl };
  arch.push(row); writeAll_('archives', arch);
  return { ok: true, archive: row };
}

/** 每日：節慶已過、該節慶全部「已發放」、且未歸檔或歸檔後又有修改 → 自動歸檔 */
function autoArchive() {
  var today = todayStr_(), y = Number(today.slice(0, 4)), recs = readAll_('records'), arch = {}; readAll_('archives').forEach(function (a) { arch[a.key] = a; });
  var done = [];
  readAll_('festivals').forEach(function (f) {
    if (f.type === '個人') return;
    [y - 1, y].forEach(function (yy) {
      var d = f['d' + yy]; if (!d || d >= today) return;
      var fest = festShort_(f.name), rs = recs.filter(function (r) { return Number(r.year) === yy && festShort_(r.festival) === fest; });
      if (!rs.length || rs.some(function (r) { return r.status !== '已發放'; })) return;
      var a = arch[yy + '|' + fest], lastEdit = rs.map(function (r) { return r.updatedAt || ''; }).sort().pop();
      if (a && a.archivedAt >= lastEdit) return;
      archiveFestival_(yy, fest, 'auto'); done.push(yy + ' ' + fest);
    });
  });
  Logger.log('autoArchive：' + (done.join('、') || '沒有需要歸檔的'));
  return done;
}

/* ───────── Claude（金鑰在指令碼屬性，網頁看不到） ───────── */
function claude_(p) {
  var key = prop_('CLAUDE_API_KEY', ''); if (!key) return { ok: false, error: '後端尚未設定 CLAUDE_API_KEY' };
  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: prop_('CLAUDE_MODEL', 'claude-opus-5'), max_tokens: 6000, output_config: { effort: 'medium' }, system: p.system, messages: [{ role: 'user', content: p.user }] })
  });
  var d = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() >= 300) return { ok: false, error: (d.error && d.error.message) || ('HTTP ' + res.getResponseCode()) };
  if (d.stop_reason === 'refusal') return { ok: false, error: '模型拒絕了這個請求' };
  return { ok: true, text: (d.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join(''), model: d.model };
}

/* ───────── 提醒 ───────── */
function todayStr_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function daysBetween_(a, b) { return Math.round((new Date(b) - new Date(a)) / 86400000); }
function upcoming_(leadDefault) {
  var today = todayStr_(), y = Number(today.slice(0, 4)), out = [];
  readAll_('festivals').forEach(function (f) {
    if (String(f.enabled) === 'false' || f.enabled === '') return;
    var lead = Number(f.leadDays) || leadDefault;
    if (f.type === '個人') {
      readAll_('people').forEach(function (p) {
        if (String(p.receivesGift) === 'false') return;
        var src = f.name.indexOf('生日') >= 0 ? p.birthday : p.joinDate; if (!src) return;
        [y, y + 1].forEach(function (yy) { var d = yy + '-' + String(src).slice(5), diff = daysBetween_(today, d); if (diff >= 0 && diff <= lead) out.push({ name: f.name, date: d, days: diff, kind: '個人', who: p.name, key: f.name + '|' + p.name + '|' + yy }); });
      }); return;
    }
    [y, y + 1].forEach(function (yy) { var d = f['d' + yy]; if (!d) return; var diff = daysBetween_(today, d); if (diff >= 0 && diff <= lead) out.push({ name: f.name, date: d, days: diff, kind: '節慶', key: f.name + '|' + yy }); });
  });
  return out;
}
function buildContext_(name, date, kind, who) {
  var s = settings_(), recs = readAll_('records'), y = Number(String(date).slice(0, 4));
  var sum = function (arr) { return arr.reduce(function (t, r) { return t + (Number(r.subtotal) || 0); }, 0); };
  var same = function (r, yy) { return Number(r.year) === yy && festShort_(r.festival) === festShort_(name); };
  var last = recs.filter(function (r) { return same(r, y - 1); }), cur = recs.filter(function (r) { return same(r, y); });
  return { name: name, date: date, days: daysBetween_(todayStr_(), date), kind: kind, who: who, settings: s,
    lastYearTotal: sum(last), lastYearItems: last.map(function (r) { return r.itemName + '×' + r.qty + '（' + r.targetType + '）'; }).slice(0, 8),
    thisYearTotal: sum(cur), thisYearCount: cur.length, thisYearDone: cur.filter(function (r) { return r.status === '已發放'; }).length,
    todo: ['選定禮盒與預算（看「選禮顧問」）', '確認名單與數量', '下單並確認到貨日', '賀卡／明信片與發放安排'] };
}
function messageText_(c) {
  var title = c.kind === '個人' ? (c.who + ' 的' + c.name) : c.name;
  var lines = ['🎁 【' + c.settings.companyName + '】' + title + ' 還有 ' + c.days + ' 天（' + c.date + '）', '',
    '去年：' + (c.lastYearTotal ? 'NT$ ' + c.lastYearTotal.toLocaleString() + '｜' + c.lastYearItems.join('、') : '沒有紀錄'),
    '今年：' + (c.thisYearCount ? 'NT$ ' + c.thisYearTotal.toLocaleString() + '｜' + c.thisYearCount + ' 筆，' + c.thisYearDone + ' 筆已發放' : '尚未規劃'), '', '建議現在開始：'].concat(c.todo.map(function (t, i) { return (i + 1) + '. ' + t; }));
  if (c.settings.toolUrl) lines.push('', '打開帳本：' + c.settings.toolUrl);
  return lines.join('\n');
}
function sendReminder_(c, force) {
  var s = c.settings, results = [], text = messageText_(c), subject = '【節慶提醒】' + (c.kind === '個人' ? c.who + ' 的' + c.name : c.name) + ' 還有 ' + c.days + ' 天';
  try { var to = (s.recipients || '').split(',').map(function (x) { return x.trim(); }).filter(String); if (to.length) { GmailApp.sendEmail(to.join(','), subject, text, { name: s.companyName + ' 節慶帳本' }); results.push('email:ok'); } else results.push('email:無收件人'); } catch (e) { results.push('email:' + e); }
  try { var tk = prop_('TG_TOKEN', ''), chat = prop_('TG_CHAT', ''); if (tk && chat) { UrlFetchApp.fetch('https://api.telegram.org/bot' + tk + '/sendMessage', { method: 'post', payload: { chat_id: chat, text: text }, muteHttpExceptions: true }); results.push('telegram:ok'); } else results.push('telegram:未設定'); } catch (e) { results.push('telegram:' + e); }
  try {
    var cal = (s.calendarId && s.calendarId !== 'primary') ? CalendarApp.getCalendarById(s.calendarId) : CalendarApp.getDefaultCalendar();
    if (cal) { var d = new Date(c.date + 'T00:00:00'), prep = new Date(d); prep.setDate(prep.getDate() - (Number(s.leadDays) || 14));
      var t1 = '🎁 ' + (c.kind === '個人' ? c.who + ' ' + c.name : c.name); if (!cal.getEventsForDay(d, { search: t1 }).length) cal.createAllDayEvent(t1, d, { description: text }).addPopupReminder(9 * 60);
      var t2 = '🛒 準備：' + t1.slice(2); if (prep >= new Date() && !cal.getEventsForDay(prep, { search: t2 }).length) cal.createAllDayEvent(t2, prep, { description: text }).addPopupReminder(9 * 60);
      results.push('calendar:ok'); }
  } catch (e) { results.push('calendar:' + e); }
  if (!force) sheet_('log').appendRow([c.key || (c.name + '|' + c.date), new Date().toISOString(), 'email,telegram,calendar', results.join(' ')]);
  return results;
}
function dailyCheck() {
  var s = settings_(), lead = Number(s.leadDays) || 14, sent = {}; readAll_('log').forEach(function (r) { sent[r.key] = true; });
  var list = upcoming_(lead).filter(function (u) { return !sent[u.key]; });
  list.forEach(function (u) { var c = buildContext_(u.name, u.date, u.kind, u.who); c.key = u.key; Logger.log(u.key + ' → ' + sendReminder_(c, false).join(' ')); });
  return list.length;
}
/** 每日 09:00 觸發：提醒＋自動歸檔 */
function daily() { dailyCheck(); autoArchive(); }
/** 手動測試：寄一封中秋測試信 */
function testReminderNow() { Logger.log(sendReminder_(buildContext_('中秋', todayStr_(), '節慶', ''), true)); }
/** 手動測試：歸檔今年中秋 */
function testArchiveNow() { Logger.log(JSON.stringify(archiveFestival_(new Date().getFullYear(), '中秋', 'manual'))); }
