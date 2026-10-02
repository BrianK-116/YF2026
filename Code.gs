/**
 * ĐIỂM DANH BARCODE - NGÀY HỘI CHÀO ĐÓN TÂN SINH VIÊN 2026 (Google Apps Script)
 * Gắn vào CHÍNH Google Sheet của danh sách TV/CTV: Extensions > Apps Script > dán file này.
 *
 * Bước 1: chọn hàm setupChamCong > Run (chỉ chạy 1 lần, cấp quyền khi được hỏi).
 * Bước 2: Deploy > New deployment > Web app
 *         Execute as: Me | Who has access: Anyone  -> copy URL dán vào index.html
 *
 * Sheet "ChamCong": hàng 3 là tiêu đề, dữ liệu từ hàng 4, cột E là MSSV (8 số),
 *   B = Họ và tên, F = Hiện là (Thành viên / Cộng tác viên).
 *   Từ cột H: mỗi buổi 2 cột IN / OUT (5 buổi = H..Q).
 */
const CONFIG = {
  TOKEN: 'DOI-MA-BI-MAT-NAY',        // phải trùng TOKEN trong index.html
  TZ: 'Asia/Ho_Chi_Minh',
  SHEET: 'ChamCong',
  SHEET_LOG: 'Log',
  HEADER_ROW: 3,
  FIRST_ROW: 4,
  COL_NAME: 2,       // B
  COL_MSSV: 5,       // E
  COL_ROLE: 6,       // F
  FIRST_TIME_COL: 8, // H
  // Buổi xác định theo ngày + giờ lúc CHECK IN (from <= giờ < to)
  SESSIONS: [
    { name: 'SÁNG 3/10',  date: '2026-10-03', from: 0,  to: 12 },
    { name: 'CHIỀU 3/10', date: '2026-10-03', from: 12, to: 24 },
    { name: 'SÁNG 4/10',  date: '2026-10-04', from: 0,  to: 12 },
    { name: 'CHIỀU 4/10', date: '2026-10-04', from: 12, to: 17 },
    { name: 'TỐI 4/10',   date: '2026-10-04', from: 17, to: 24 }
  ]
};

/** Chạy 1 lần: ghi lại tiêu đề cột giờ (H3:Q3) và đặt định dạng văn bản. */
function setupChamCong() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET);
  if (!sh) throw new Error('Không thấy sheet "' + CONFIG.SHEET + '"');
  const n = CONFIG.SESSIONS.length * 2;
  sh.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_TIME_COL, 1, 12).clearContent(); // xóa tiêu đề cũ H..S
  const hdr = [];
  CONFIG.SESSIONS.forEach(s => hdr.push('IN ' + s.name, 'OUT ' + s.name));
  sh.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_TIME_COL, 1, n)
    .setValues([hdr]).setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  const rows = Math.max(sh.getMaxRows() - CONFIG.FIRST_ROW + 1, 1);
  sh.getRange(CONFIG.FIRST_ROW, CONFIG.FIRST_TIME_COL, rows, n).setNumberFormat('@');
  getLog_(SpreadsheetApp.getActive());
}

function doGet() { return json_({ ok: true, message: 'Attendance API is running' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);                       // HÀNG ĐỢI: nhiều thiết bị xếp hàng lấy khóa
  } catch (err) {
    return json_({ ok: false, code: 'BUSY', message: 'Máy chủ đang bận, thử lại.' });
  }
  try {
    return json_(handleScan_(JSON.parse(e.postData.contents)));
  } catch (err) {
    return json_({ ok: false, code: 'ERROR', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function handleScan_(p) {
  if (p.token !== CONFIG.TOKEN) return { ok: false, code: 'AUTH', message: 'Sai token.' };

  const mssv = String(p.mssv || '').trim();
  const type = p.type === 'out' ? 'out' : 'in';
  if (!/^\d{8}$/.test(mssv)) return { ok: false, code: 'BAD_CODE', message: 'MSSV không hợp lệ: ' + mssv };

  // Nếu thiết bị gửi lại (mạng chập chờn) cùng 1 lượt quét -> trả kết quả cũ, không ghi lần 2
  const cache = CacheService.getScriptCache();
  const idKey = 'scan_' + p.id;
  if (p.id) { const done = cache.get(idKey); if (done) return JSON.parse(done); }

  let ts = new Date(p.ts);
  if (isNaN(ts.getTime())) ts = new Date();
  const dayStr = Utilities.formatDate(ts, CONFIG.TZ, 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(ts, CONFIG.TZ, 'HH:mm:ss');
  const hour = Number(Utilities.formatDate(ts, CONFIG.TZ, 'H'));

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SHEET);
  if (!sh) throw new Error('Không thấy sheet "' + CONFIG.SHEET + '"');
  const n = CONFIG.SESSIONS.length;

  // Tìm hàng sinh viên theo MSSV (cột E)
  const last = sh.getLastRow();
  let row = -1;
  if (last >= CONFIG.FIRST_ROW) {
    const ids = sh.getRange(CONFIG.FIRST_ROW, CONFIG.COL_MSSV, last - CONFIG.FIRST_ROW + 1, 1).getDisplayValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === mssv) { row = CONFIG.FIRST_ROW + i; break; }
  }
  if (row < 0) {
    const r = { ok: false, code: 'NOT_FOUND', mssv, message: 'Không có MSSV ' + mssv + ' trong danh sách.' };
    log_(ss, ts, mssv, '', type, '', r, p.device);
    return remember_(cache, idKey, p.id, r);
  }

  const name = String(sh.getRange(row, CONFIG.COL_NAME).getDisplayValue()).trim();
  const role = String(sh.getRange(row, CONFIG.COL_ROLE).getDisplayValue()).trim();
  const cells = sh.getRange(row, CONFIG.FIRST_TIME_COL, 1, n * 2).getDisplayValues()[0];
  const IN = i => String(cells[i * 2]).trim();
  const OUT = i => String(cells[i * 2 + 1]).trim();

  // Buổi đang mở = đã check in nhưng chưa check out
  let open = -1;
  for (let i = 0; i < n; i++) if (IN(i) && !OUT(i)) { open = i; break; }

  let result;
  if (type === 'in') {
    if (open >= 0) {
      result = fail_('DUP_IN', 'Đã check in ' + CONFIG.SESSIONS[open].name + ' lúc ' + IN(open) + ', chưa check out.', mssv, name, role);
    } else {
      let s = -1;
      if (p.session !== undefined && p.session !== null && p.session !== '' && CONFIG.SESSIONS[Number(p.session)]) {
        s = Number(p.session);                               // chọn buổi thủ công
      } else {
        s = CONFIG.SESSIONS.findIndex(x => x.date === dayStr && hour >= x.from && hour < x.to);
      }
      if (s < 0) {
        result = fail_('NO_SESSION', 'Ngoài thời gian sự kiện — hãy chọn buổi thủ công.', mssv, name, role);
      } else if (IN(s)) {
        result = fail_('SESSION_DONE', CONFIG.SESSIONS[s].name + ' đã hoàn thành (vào ' + IN(s) + ', ra ' + OUT(s) + ').', mssv, name, role);
      } else {
        sh.getRange(row, CONFIG.FIRST_TIME_COL + s * 2).setNumberFormat('@').setValue(timeStr);
        result = { ok: true, mssv, name, cls: role, type, session: CONFIG.SESSIONS[s].name, time: timeStr };
      }
    }
  } else {
    if (open < 0) {
      result = fail_('NO_IN', 'Chưa check in nên không thể check out.', mssv, name, role);
    } else {
      sh.getRange(row, CONFIG.FIRST_TIME_COL + open * 2 + 1).setNumberFormat('@').setValue(timeStr);
      result = { ok: true, mssv, name, cls: role, type, session: CONFIG.SESSIONS[open].name, time: timeStr };
    }
  }

  log_(ss, ts, mssv, name, type, result.session || '', result, p.device);
  SpreadsheetApp.flush();
  return remember_(cache, idKey, p.id, result);
}

function fail_(code, message, mssv, name, role) {
  return { ok: false, code, message, mssv, name, cls: role };
}
function remember_(cache, key, id, result) {
  if (id) cache.put(key, JSON.stringify(result), 21600);
  return result;
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
  getLog_(ss).appendRow([ts, mssv, name, type === 'in' ? 'Check in' : 'Check out', session,
    result.ok ? 'OK' : result.code, result.ok ? '' : result.message, device || '']);
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
