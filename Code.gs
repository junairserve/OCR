/**
 * SN紐づけツール（完全版 v3）
 * - v2 + 管理者画面「在庫へ戻す」ワンクリック化（修理戻り運用）
 * - ステータス変更は EVENT_LOG に履歴を残す（いつ・誰が・何を・どう変えたか）
 */

const SHEET_PCB   = 'PCB_MASTER';
const SHEET_BODY  = 'BODY_MASTER';
const SHEET_LOG   = 'LINK_LOG';
const SHEET_EVENT = 'EVENT_LOG';

const BODY_STATUS_IN_STOCK = '在庫';
const BODY_STATUS_IN_USE   = '使用中';
const BODY_STATUS_REPAIR   = '修理中';
const BODY_STATUS_DISPOSED = '廃棄';

const PCB_STATUS_UNUSED    = '未使用';
const PCB_STATUS_IN_USE    = '使用中';
const PCB_STATUS_DISPOSED  = '廃棄';

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';
  if (e && e.parameter && e.parameter.api === '1') return handleApiGet_(e.parameter);

  const t = HtmlService.createTemplateFromFile(page === 'admin' ? 'Admin' : 'Index');
  t.appTitle = 'SN紐づけツール';
  return t.evaluate()
    .setTitle('SN紐づけツール')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function handleApiGet_(p) {
  const action = (p.action || '').toString();

  if (action === 'list_pcb') {
    const status = (p.status ?? PCB_STATUS_UNUSED).toString();
    const q = (p.q || '').toString().trim();
    const limit = Math.min(Number(p.limit || 200), 2000);
    return jsonOut_({ ok:true, rows:listPcb_(status, q, limit) });
  }
  if (action === 'list_body') {
    // status = '' なら全件（管理者検索用）
    const status = (p.status ?? BODY_STATUS_IN_STOCK).toString();
    const q = (p.q || '').toString().trim();
    const limit = Math.min(Number(p.limit || 200), 2000);
    return jsonOut_({ ok:true, rows:listBody_(status, q, limit) });
  }
  if (action === 'list_log') {
    const from = (p.from || '').toString().trim();
    const to   = (p.to   || '').toString().trim();
    const limit = Math.min(Number(p.limit || 500), 5000);
    return jsonOut_({ ok:true, rows:listLog_(from, to, limit) });
  }
  if (action === 'export_log_csv') {
    const from = (p.from || '').toString().trim();
    const to   = (p.to   || '').toString().trim();
    return exportLogCsv_(from, to);
  }
  if (action === 'list_event') {
    const from = (p.from || '').toString().trim();
    const to   = (p.to   || '').toString().trim();
    const q    = (p.q || '').toString().trim();
    const limit = Math.min(Number(p.limit || 500), 5000);
    return jsonOut_({ ok:true, rows:listEvent_(from, to, q, limit) });
  }
  if (action === 'export_event_csv') {
    const from = (p.from || '').toString().trim();
    const to   = (p.to   || '').toString().trim();
    const q    = (p.q || '').toString().trim();
    return exportEventCsv_(from, to, q);
  }

  return jsonOut_({ ok:false, error:'unknown action' });
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? e.postData.contents : '';
    const data = JSON.parse(body || '{}');
    const action = (data.action || '').toString();

    if (action === 'import_pcb')  return jsonOut_(Object.assign({ ok:true }, importPcbItems_(Array.isArray(data.items)?data.items:[])));
    if (action === 'import_body') return jsonOut_(Object.assign({ ok:true }, importBodyItems_(Array.isArray(data.items)?data.items:[])));
    if (action === 'link')        return jsonOut_(Object.assign({ ok:true }, saveLink_(data.payload || {})));

    if (action === 'update_status') {
      return jsonOut_(Object.assign({ ok:true }, updateStatus_(
        (data.kind||'').toString(),
        data.sn,
        (data.status||'').toString(),
        (data.operator||'').toString(),
        (data.note||'').toString()
      )));
    }

    return jsonOut_({ ok:false, error:'unknown action' });
  } catch (err) {
    return jsonOut_({ ok:false, error:String(err) });
  }
}

// ===== Sheets / Utils =====
function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}
function ensureHeaders_() {
  const pcb = getSheet_(SHEET_PCB);
  if (pcb.getLastRow() === 0) pcb.appendRow(['pcb_sn','received_date','status','note','imported_at']);
  const body = getSheet_(SHEET_BODY);
  if (body.getLastRow() === 0) body.appendRow(['body_sn','model','status','note','updated_at']);
  const log = getSheet_(SHEET_LOG);
  if (log.getLastRow() === 0) log.appendRow(['timestamp_jst','body_sn','pcb_sn','work_type','operator','note']);
  const ev = getSheet_(SHEET_EVENT);
  if (ev.getLastRow() === 0) ev.appendRow(['timestamp_jst','kind','sn','from_status','to_status','operator','note']);
}
function toJst_(d) { return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); }
function normalizeSn8_(sn) {
  let s = (sn === null || sn === undefined) ? '' : String(sn).trim();
  s = s.replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
  s = s.replace(/[^0-9]/g, '');
  if (!s) return '';
  if (s.length > 8) s = s.slice(-8);
  return s.padStart(8, '0');
}
function normalizeDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  const s = (v === null || v === undefined) ? '' : String(v).trim();
  if (!s) return '';
  const m = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (m) {
    const y = m[1], mo = String(Number(m[2])).padStart(2,'0'), da = String(Number(m[3])).padStart(2,'0');
    return `${y}-${mo}-${da}`;
  }
  return s;
}

// ===== List =====
function listPcb_(status, q, limit) {
  ensureHeaders_();
  const sh = getSheet_(SHEET_PCB);
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i=1;i<values.length;i++){
    const pcbSn = String(values[i][0]||'').trim();
    const recDate = String(values[i][1]||'').trim();
    const st = String(values[i][2]||'').trim();
    const note = String(values[i][3]||'').trim();
    if (status && st !== status) continue;
    if (q && pcbSn.indexOf(q) === -1) continue;
    out.push({ pcb_sn: pcbSn, received_date: recDate, status: st, note });
    if (out.length >= limit) break;
  }
  return out;
}
function listBody_(status, q, limit) {
  ensureHeaders_();
  const sh = getSheet_(SHEET_BODY);
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i=1;i<values.length;i++){
    const bodySn = String(values[i][0]||'').trim();
    const model = String(values[i][1]||'').trim();
    const st = String(values[i][2]||'').trim();
    const note = String(values[i][3]||'').trim();
    if (status && st !== status) continue;
    if (q && bodySn.indexOf(q) === -1) continue;
    out.push({ body_sn: bodySn, model, status: st, note });
    if (out.length >= limit) break;
  }
  return out;
}
function listLog_(from, to, limit) {
  ensureHeaders_();
  const sh = getSheet_(SHEET_LOG);
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i=1;i<values.length;i++){
    const ts = String(values[i][0]||'').trim();
    const d = ts.slice(0,10);
    if (from && d < from) continue;
    if (to && d > to) continue;
    out.push({
      timestamp_jst: ts,
      body_sn: String(values[i][1]||'').trim(),
      pcb_sn:  String(values[i][2]||'').trim(),
      work_type: String(values[i][3]||'').trim(),
      operator: String(values[i][4]||'').trim(),
      note: String(values[i][5]||'').trim(),
    });
    if (out.length >= limit) break;
  }
  return out;
}
function listEvent_(from, to, q, limit) {
  ensureHeaders_();
  const sh = getSheet_(SHEET_EVENT);
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i=1;i<values.length;i++){
    const ts = String(values[i][0]||'').trim();
    const d = ts.slice(0,10);
    const kind = String(values[i][1]||'').trim();
    const sn = String(values[i][2]||'').trim();
    if (from && d < from) continue;
    if (to && d > to) continue;
    if (q && sn.indexOf(q) === -1) continue;
    out.push({
      timestamp_jst: ts,
      kind, sn,
      from_status: String(values[i][3]||'').trim(),
      to_status: String(values[i][4]||'').trim(),
      operator: String(values[i][5]||'').trim(),
      note: String(values[i][6]||'').trim(),
    });
    if (out.length >= limit) break;
  }
  return out;
}

// ===== Import =====
function importPcbItems_(items) {
  ensureHeaders_();
  const sh = getSheet_(SHEET_PCB);
  const values = sh.getDataRange().getValues();
  const exists = new Set();
  for (let i=1;i<values.length;i++){
    const sn = String(values[i][0]||'').trim();
    if (sn) exists.add(sn);
  }
  let added=0, skipped=0, invalid=0;
  const nowJst = toJst_(new Date());
  items.forEach(it => {
    const sn = normalizeSn8_(it.sn);
    const date = normalizeDate_(it.date);
    if (!sn) { invalid++; return; }
    if (exists.has(sn)) { skipped++; return; }
    sh.appendRow([sn, date, PCB_STATUS_UNUSED, '', nowJst]);
    exists.add(sn); added++;
  });
  return { added, skipped, invalid };
}

function importBodyItems_(items) {
  ensureHeaders_();
  const sh = getSheet_(SHEET_BODY);
  const values = sh.getDataRange().getValues();
  const index = new Map();
  for (let i=1;i<values.length;i++){
    const sn = String(values[i][0]||'').trim();
    if (sn) index.set(sn, i+1);
  }
  let added=0, updated=0, invalid=0;
  const nowJst = toJst_(new Date());
  items.forEach(it => {
    const sn = normalizeSn8_(it.sn);
    const model = (it.model || '').toString().trim();
    if (!sn) { invalid++; return; }
    if (index.has(sn)) {
      const r = index.get(sn);
      if (model) sh.getRange(r, 2).setValue(model);
      sh.getRange(r, 3).setValue(BODY_STATUS_IN_STOCK);
      sh.getRange(r, 5).setValue(nowJst);
      updated++;
    } else {
      sh.appendRow([sn, model, BODY_STATUS_IN_STOCK, '', nowJst]);
      added++;
    }
  });
  return { added, updated, invalid };
}

// ===== Link Save =====
function saveLink_(payload) {
  ensureHeaders_();
  const bodySn = normalizeSn8_(payload.body_sn);
  const pcbSn  = normalizeSn8_(payload.pcb_sn);
  if (!bodySn || !pcbSn) throw new Error('body_sn / pcb_sn are required');

  const workType = (payload.work_type || '組立').toString();
  const operator = (payload.operator || '').toString();
  const note = (payload.note || '').toString();

  // body check
  const bodySh = getSheet_(SHEET_BODY);
  const bodyVals = bodySh.getDataRange().getValues();
  let bodyRow = -1;
  let bodyStatus = '';
  for (let i=1;i<bodyVals.length;i++){
    const sn = String(bodyVals[i][0]||'').trim();
    if (sn === bodySn) { bodyRow = i+1; bodyStatus = String(bodyVals[i][2]||'').trim(); break; }
  }
  if (bodyRow === -1) throw new Error('本体SNがマスタに存在しません');
  if (bodyStatus !== BODY_STATUS_IN_STOCK) throw new Error('この本体SNは在庫ではありません（重複使用防止）');

  // pcb check
  const pcbSh = getSheet_(SHEET_PCB);
  const pcbVals = pcbSh.getDataRange().getValues();
  let pcbRow = -1;
  let pcbStatus = '';
  for (let i=1;i<pcbVals.length;i++){
    const sn = String(pcbVals[i][0]||'').trim();
    if (sn === pcbSn) { pcbRow = i+1; pcbStatus = String(pcbVals[i][2]||'').trim(); break; }
  }
  if (pcbRow === -1) throw new Error('基板SNがマスタに存在しません');
  if (pcbStatus !== PCB_STATUS_UNUSED) throw new Error('この基板SNは未使用ではありません（重複使用防止）');

  const nowJst = toJst_(new Date());

  bodySh.getRange(bodyRow, 3).setValue(BODY_STATUS_IN_USE);
  bodySh.getRange(bodyRow, 5).setValue(nowJst);

  pcbSh.getRange(pcbRow, 3).setValue(PCB_STATUS_IN_USE);

  // link log
  getSheet_(SHEET_LOG).appendRow([nowJst, bodySn, pcbSn, workType, operator, note]);

  // event logs (state transitions)
  appendEvent_('body', bodySn, bodyStatus, BODY_STATUS_IN_USE, operator, `紐づけ保存：${note||''}`.trim());
  appendEvent_('pcb',  pcbSn,  pcbStatus,  PCB_STATUS_IN_USE,  operator, `紐づけ保存：${note||''}`.trim());

  return { linked:true };
}

// ===== Status update with event log =====
function updateStatus_(kind, sn, status, operator, note) {
  ensureHeaders_();
  const k = (kind || '').toLowerCase();
  if (k !== 'body' && k !== 'pcb') throw new Error('kind must be body or pcb');
  const s = normalizeSn8_(sn);
  if (!s) throw new Error('sn is required');

  const nowJst = toJst_(new Date());
  const op = (operator || '').toString();
  const nt = (note || '').toString();

  if (k === 'body') {
    const allowed = new Set([BODY_STATUS_IN_STOCK, BODY_STATUS_IN_USE, BODY_STATUS_REPAIR, BODY_STATUS_DISPOSED]);
    if (!allowed.has(status)) throw new Error('invalid body status');
    const sh = getSheet_(SHEET_BODY);
    const vals = sh.getDataRange().getValues();
    for (let i=1;i<vals.length;i++){
      if (String(vals[i][0]||'').trim() === s) {
        const from = String(vals[i][2]||'').trim();
        sh.getRange(i+1, 3).setValue(status);
        sh.getRange(i+1, 5).setValue(nowJst);
        appendEvent_('body', s, from, status, op, nt);
        return { updated:true, from_status: from, to_status: status };
      }
    }
    throw new Error('body sn not found');
  }

  const allowed = new Set([PCB_STATUS_UNUSED, PCB_STATUS_IN_USE, PCB_STATUS_DISPOSED]);
  if (!allowed.has(status)) throw new Error('invalid pcb status');
  const sh = getSheet_(SHEET_PCB);
  const vals = sh.getDataRange().getValues();
  for (let i=1;i<vals.length;i++){
    if (String(vals[i][0]||'').trim() === s) {
      const from = String(vals[i][2]||'').trim();
      sh.getRange(i+1, 3).setValue(status);
      appendEvent_('pcb', s, from, status, op, nt);
      return { updated:true, from_status: from, to_status: status };
    }
  }
  throw new Error('pcb sn not found');
}

function appendEvent_(kind, sn, fromStatus, toStatus, operator, note) {
  ensureHeaders_();
  getSheet_(SHEET_EVENT).appendRow([toJst_(new Date()), kind, sn, fromStatus, toStatus, operator, note]);
}

// ===== CSV Exports =====
function exportLogCsv_(from, to) {
  const rows = listLog_(from, to, 50000);
  const header = ['timestamp_jst','body_sn','pcb_sn','work_type','operator','note'];
  const lines = [header.join(',')];
  rows.forEach(r => lines.push([r.timestamp_jst,r.body_sn,r.pcb_sn,r.work_type,r.operator,r.note].map(csvEsc_).join(',')));
  const out = ContentService.createTextOutput(lines.join('\n'));
  out.setMimeType(ContentService.MimeType.CSV);
  out.setHeader('Content-Disposition', 'attachment; filename="link_log.csv"');
  return out;
}

function exportEventCsv_(from, to, q) {
  const rows = listEvent_(from, to, q, 50000);
  const header = ['timestamp_jst','kind','sn','from_status','to_status','operator','note'];
  const lines = [header.join(',')];
  rows.forEach(r => lines.push([r.timestamp_jst,r.kind,r.sn,r.from_status,r.to_status,r.operator,r.note].map(csvEsc_).join(',')));
  const out = ContentService.createTextOutput(lines.join('\n'));
  out.setMimeType(ContentService.MimeType.CSV);
  out.setHeader('Content-Disposition', 'attachment; filename="event_log.csv"');
  return out;
}

function csvEsc_(v) {
  const s = String(v ?? '');
  if (s.includes('"') || s.includes(',') || s.includes('\n')) return '"' + s.replace(/"/g,'""') + '"';
  return s;
}

function jsonOut_(obj) {
  const out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}
