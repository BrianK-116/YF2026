const CONFIG = {
  TOKEN: 'CEER',
  TZ: 'Asia/Ho_Chi_Minh',
  SHEET: 'ChamCong',
  SHEET_LOG: 'Log',
  FIRST_ROW: 4,
  COL_NAME: 2,       // B
  COL_MSSV: 5,       // E
  COL_ROLE: 6,       // F
  FIRST_TIME_COL: 8, // H
  CACHE_TTL: 21600,  // 6 giờ
  MEDIA_IDS: ['25128013', '25124045', '24156133'],  // team Media (cả 3/10 và 4/10)
  SHEET_TONG: 'DANH SÁCH TỔNG',  // đọc: B tên, E MSSV, F vai trò, G SĐT, H form đăng ký, I..M đăng ký 5 ca, N team 3/10, O team 4/10, Q ghi chú

  COLOR_IN: '#bbf7d0',
  COLOR_OUT: '#fecaca',
  COLOR_SPAN: '#fef08a',

  SESSIONS: [
    { name: 'SÁNG 3/10',  date: '2026-10-03', from: 7.0,  to: 12.5 },
    { name: 'CHIỀU 3/10', date: '2026-10-03', from: 12.5, to: 20.0 },
    { name: 'SÁNG 4/10',  date: '2026-10-04', from: 3.0,  to: 11.5 },
    { name: 'CHIỀU 4/10', date: '2026-10-04', from: 11.5, to: 15.5 },
    { name: 'TỐI 4/10',   date: '2026-10-04', from: 15.5, to: 23.0 }
  ]
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.token !== CONFIG.TOKEN) return json_({ ok: false, code: 'AUTH', message: 'Sai token' });
  if (p.action === 'status') {
    try { return json_(getStatus_()); }
    catch (err) { return json_({ ok: false, code: 'ERROR', message: String(err) }); }
  }
  return json_({ ok: true, message: 'API hoạt động' });
}

// Tìm sheet theo tên, bỏ qua khoảng trắng đầu/cuối (vd ' GÂY QUỸ')
function findSheet_(ss, name) {
  const t = name.trim().toLowerCase();
  return ss.getSheets().find(s => s.getName().trim().toLowerCase() === t) || null;
}
// Lấy các MSSV (8 chữ số) trong 1 cột của 1 sheet
function colIds_(ss, sheetName, col) {
  const sh = findSheet_(ss, sheetName);
  if (!sh) throw new Error('Không thấy sheet: ' + sheetName);
  const last = sh.getLastRow();
  if (last < 1) return [];
  return sh.getRange(1, col, last, 1).getDisplayValues()
    .map(r => { const m = String(r[0]).match(/\d{8}/); return m ? m[0] : ''; })
    .filter(Boolean);
}
// Chia team theo các tab: gặp trước thì ưu tiên trước; Media ghi đè tất cả
function teamMaps_(ss) {
  const t3 = {}, t4 = {};
  const put = (map, ids, label) => ids.forEach(m => { if (!map[m]) map[m] = label; });
  put(t3, colIds_(ss, '3.10', 4), 'Gian hàng 3.10');                     // 3.10 cột D
  put(t3, colIds_(ss, '3.10', 8), 'Set up gian hàng Ban');               // 3.10 cột H
  put(t4, colIds_(ss, 'GIAN HÀNG NHÀ TÀI TRỢ', 3), 'Gian hàng nhà tài trợ');   // cột C
  put(t4, colIds_(ss, 'GIAN HÀNG NHÀ TÀI TRỢ', 10), 'Gian hàng nhà tài trợ');  // cột J
  put(t4, colIds_(ss, 'GÂY QUỸ', 4), 'Gây quỹ');                         // GÂY QUỸ cột D
  CONFIG.MEDIA_IDS.forEach(m => { t3[m] = 'Media'; t4[m] = 'Media'; });
  return { t3: t3, t4: t4 };
}

// Danh sách người + đăng ký + team + giờ check in/out (join theo MSSV). Cache 4 giây.
function getStatus_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('status');
  if (hit) return JSON.parse(hit);

  const ss = SpreadsheetApp.getActive();
  const n = CONFIG.SESSIONS.length;

  // Giờ chấm công theo MSSV
  const cc = ss.getSheetByName(CONFIG.SHEET);
  const att = {};
  const ccLast = cc.getLastRow();
  if (ccLast >= CONFIG.FIRST_ROW) {
    const lastCol = CONFIG.FIRST_TIME_COL + n * 2 - 1;
    cc.getRange(CONFIG.FIRST_ROW, 1, ccLast - CONFIG.FIRST_ROW + 1, lastCol).getDisplayValues().forEach(r => {
      const m = String(r[CONFIG.COL_MSSV - 1]).trim();
      if (m) att[m] = r.slice(CONFIG.FIRST_TIME_COL - 1).map(x => String(x).trim());
    });
  }

  const tm = teamMaps_(ss);

  // Danh sách tổng
  const tg = ss.getSheetByName(CONFIG.SHEET_TONG);
  const tLast = tg.getLastRow();
  const people = [];
  if (tLast >= CONFIG.FIRST_ROW) {
    tg.getRange(CONFIG.FIRST_ROW, 1, tLast - CONFIG.FIRST_ROW + 1, 17).getDisplayValues().forEach(r => {
      const m = String(r[4]).trim();
      if (!m) return;
      let name = String(r[1]).trim();
      if (!name || name === 'nan') name = (String(r[2]).trim() + ' ' + String(r[3]).trim()).trim();
      const reg = [];
      for (let i = 0; i < n; i++) reg.push(String(r[8 + i]).toUpperCase() === 'TRUE' ? 1 : 0);
      people.push({
        m: m, n: name, r: String(r[5]).trim(), p: String(r[6]).trim(), h: String(r[7]).trim(),
        reg: reg, tN: tm.t3[m] || '', tO: tm.t4[m] || '', note: String(r[16]).trim(),
        a: att[m] || new Array(n * 2).fill('')
      });
    });
  }
  const out = { ok: true, ts: new Date().toISOString(), sessions: CONFIG.SESSIONS.map(s => s.name), people: people };
  try { cache.put('status', JSON.stringify(out), 4); } catch (e) {}
  return out;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
  } catch (err) {
    return json_({ ok: false, code: 'BUSY', message: 'Máy chủ bận' });
  }
  let out, logArgs = null;
  try {
    const p = JSON.parse(e.postData.contents);
    const r = handleScan_(p);
    out = r.result;
    logArgs = r.log;
  } catch (err) {
    out = { ok: false, code: 'ERROR', message: String(err) };
  } finally {
    lock.releaseLock();
  }
  // Ghi Log SAU khi đã nhả khoá để người sau không phải chờ
  if (logArgs) { try { log_.apply(null, logArgs); } catch (e2) {} }
  return json_(out);
}

function handleScan_(p) {
  if (p.token !== CONFIG.TOKEN) return { result: { ok: false, code: 'AUTH', message: 'Sai token' } };

  // 1) Chống trùng theo id: gửi lại cùng id => trả đúng kết quả cũ
  const cache = CacheService.getScriptCache();
  if (p.id) {
    const prev = cache.get('id_' + p.id);
    if (prev) return { result: JSON.parse(prev) };
  }

  const mssv = String(p.mssv || '').trim();
  const type = p.type === 'out' ? 'out' : 'in';

  let ts = new Date(p.ts);
  if (isNaN(ts.getTime())) ts = new Date();
  const dayStr = Utilities.formatDate(ts, CONFIG.TZ, 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(ts, CONFIG.TZ, 'HH:mm:ss');
  const timeDec = Number(Utilities.formatDate(ts, CONFIG.TZ, 'H')) + Number(Utilities.formatDate(ts, CONFIG.TZ, 'm')) / 60;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SHEET);
  const n = CONFIG.SESSIONS.length;
  const lastCol = CONFIG.FIRST_TIME_COL + n * 2 - 1;

  // 2) Tìm hàng (cache) + đọc CẢ HÀNG trong 1 lệnh
  const found = readRow_(sh, mssv, cache, lastCol);
  if (!found) {
    const res = { ok: false, code: 'NOT_FOUND', message: 'Không có MSSV trong Sheet', mssv };
    return { result: res, log: [ss, ts, mssv, '', type, '', res, p.device] };
  }
  const row = found.row, vals = found.vals;
  const name = String(vals[CONFIG.COL_NAME - 1]).trim();
  const role = String(vals[CONFIG.COL_ROLE - 1]).trim();
  const cells = vals.slice(CONFIG.FIRST_TIME_COL - 1);
  const IN = i => String(cells[i * 2]).trim();
  const OUT = i => String(cells[i * 2 + 1]).trim();

  let open = -1;
  for (let i = 0; i < n; i++) if (IN(i) && !OUT(i)) { open = i; break; }

  let curSess = -1;
  if (p.session !== undefined && p.session !== '' && CONFIG.SESSIONS[Number(p.session)]) {
    curSess = Number(p.session);
  } else {
    curSess = CONFIG.SESSIONS.findIndex(x => x.date === dayStr && timeDec >= x.from && timeDec < x.to);
  }

  let result;
  if (type === 'in') {
    if (curSess < 0) {
      result = { ok: false, code: 'NO_SESSION', message: 'Ngoài khung giờ sự kiện' };
    } else if (open >= 0 && open >= curSess) {
      result = { ok: false, code: 'DUP_IN', message: 'Đã check in ' + CONFIG.SESSIONS[open].name + ', chưa check out' };
    } else if (IN(curSess)) {
      result = { ok: false, code: 'SESSION_DONE', message: CONFIG.SESSIONS[curSess].name + ' đã check in trước đó' };
    } else {
      // Quên check out ca cũ => tự đóng ca cũ rồi cho check in ca mới
      if (open >= 0 && open < curSess) {
        sh.getRange(row, CONFIG.FIRST_TIME_COL + open * 2 + 1)
          .setValue('Quên out').setBackground(CONFIG.COLOR_SPAN);
      }
      sh.getRange(row, CONFIG.FIRST_TIME_COL + curSess * 2).setValue(timeStr).setBackground(CONFIG.COLOR_IN);
      result = { ok: true, mssv, name, cls: role, type, session: CONFIG.SESSIONS[curSess].name, time: timeStr };
    }
  } else {
    if (open < 0) {
      result = { ok: false, code: 'NO_IN', message: 'Chưa check in ca nào' };
    } else {
      const targetOut = (curSess >= open) ? curSess : open;
      const inCol = CONFIG.FIRST_TIME_COL + open * 2;
      const outCol = CONFIG.FIRST_TIME_COL + targetOut * 2 + 1;
      const numCols = outCol - inCol + 1;
      const rowVals = [IN(open)], rowBgs = [CONFIG.COLOR_IN];
      for (let c = 1; c < numCols - 1; c++) { rowVals.push('—'); rowBgs.push(CONFIG.COLOR_SPAN); }
      rowVals.push(timeStr); rowBgs.push(CONFIG.COLOR_OUT);
      const range = sh.getRange(row, inCol, 1, numCols);
      range.setValues([rowVals]);
      range.setBackgrounds([rowBgs]);
      const sName = (targetOut === open) ? CONFIG.SESSIONS[open].name
        : CONFIG.SESSIONS[open].name + ' → ' + CONFIG.SESSIONS[targetOut].name;
      result = { ok: true, mssv, name, cls: role, type, session: sName, time: timeStr };
    }
  }
  if (!result.ok) { result.mssv = mssv; result.name = name; }

  cache.remove('status');
  if (p.id) cache.put('id_' + p.id, JSON.stringify(result), CONFIG.CACHE_TTL);
  return { result, log: [ss, ts, mssv, name, type, result.session || '', result, p.device] };
}

// Cache map MSSV -> số hàng; luôn kiểm tra lại MSSV ở hàng đọc được (an toàn khi chèn/xoá hàng)
function readRow_(sh, mssv, cache, lastCol) {
  const read = r => sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
  let map = null;
  const raw = cache.get('rowmap');
  if (raw) map = JSON.parse(raw);

  if (map && map[mssv]) {
    const vals = read(map[mssv]);
    if (String(vals[CONFIG.COL_MSSV - 1]).trim() === mssv) return { row: map[mssv], vals };
  }
  // Dựng lại map
  const last = sh.getLastRow();
  if (last < CONFIG.FIRST_ROW) return null;
  const ids = sh.getRange(CONFIG.FIRST_ROW, CONFIG.COL_MSSV, last - CONFIG.FIRST_ROW + 1, 1).getDisplayValues();
  map = {};
  ids.forEach((r, i) => { const k = String(r[0]).trim(); if (k) map[k] = CONFIG.FIRST_ROW + i; });
  try { cache.put('rowmap', JSON.stringify(map), CONFIG.CACHE_TTL); } catch (e) {}
  if (!map[mssv]) return null;
  return { row: map[mssv], vals: read(map[mssv]) };
}

function getLog_(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_LOG);
    sh.appendRow(['Thời gian', 'MSSV', 'Họ tên', 'Loại', 'Buổi', 'Kết quả', 'Ghi chú', 'Thiết bị']);
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');
  }
  return sh;
}
function log_(ss, ts, mssv, name, type, session, result, device) {
  getLog_(ss).appendRow([ts, "'" + mssv, name, type === 'in' ? 'Check in' : 'Check out', session,
    result.ok ? 'OK' : result.code, result.ok ? '' : result.message, device || '']);
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Chạy tay khi bạn sửa/chèn/xoá hàng trong Sheet
function clearCache() { CacheService.getScriptCache().remove('rowmap'); }
