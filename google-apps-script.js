// ================================================================
//  IT Asset Manager — Google Apps Script Backend
//  Phiên bản: 2.0
// ================================================================
//
//  HƯỚNG DẪN CÀI ĐẶT (làm 1 lần duy nhất):
//
//  1. Tạo Google Sheet mới tại sheets.google.com
//  2. Mở menu: Extensions (Tiện ích mở rộng) > Apps Script
//  3. Xóa toàn bộ code mặc định, dán toàn bộ file này vào
//  4. Nhấn Save (Ctrl+S)
//  5. Nhấn Deploy > New deployment
//     - Type: Web app
//     - Execute as: Me
//     - Who has access: Anyone   ← quan trọng
//  6. Nhấn Deploy, copy URL hiện ra
//  7. Mở file it-asset-manager.html bằng Notepad/VS Code
//     Tìm dòng:  const SCRIPT_URL = 'PASTE_YOUR_URL_HERE';
//     Thay bằng URL vừa copy
//  8. Lưu HTML, mở bằng trình duyệt — xong!
//
//  LƯU Ý: Mỗi khi sửa code Apps Script, phải Deploy lại
//         (Deploy > Manage deployments > Edit > New version)
// ================================================================

const ASSET_COLS = [
  'id', 'type', 'itemId', 'createdAt',
  'userName', 'pcName', 'tenThietBi', 'description',
  'boPhan', 'loai', 'ngayMua', 'baoHanh',
  'congTy', 'trangThai', 'ghiChu'
];

const HIST_COLS = [
  'id', 'assetId', 'itemId', 'assetName', 'assetType',
  'fromUser', 'toUser', 'fromDept', 'toDept',
  'action', 'date', 'notes', 'createdAt'
];

// ── Entry point ──────────────────────────────────────────────
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Nếu có tham số ?d=... thì đây là thao tác ghi
    if (e.parameter && e.parameter.d) {
      const req = JSON.parse(decodeURIComponent(e.parameter.d));
      handleWrite(ss, req);
      return respond({ ok: true });
    }

    // Không có tham số → trả về toàn bộ dữ liệu
    return respond({
      ok:      true,
      assets:  readSheet(ss, 'assets',  ASSET_COLS),
      history: readSheet(ss, 'history', HIST_COLS),
    });

  } catch (err) {
    return respond({ ok: false, error: err.toString() });
  }
}

// ── Xử lý ghi (add / update / delete) ───────────────────────
function handleWrite(ss, req) {
  const { action, sheet, payload } = req;
  const cols = sheet === 'assets' ? ASSET_COLS : HIST_COLS;

  if (action === 'add')    addRow(ss, sheet, payload, cols);
  if (action === 'update') updateRow(ss, sheet, payload, cols);
  if (action === 'delete') deleteRow(ss, sheet, payload.id);
}

// ── Đọc dữ liệu từ sheet ─────────────────────────────────────
function readSheet(ss, name, cols) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() <= 1) return [];

  const data = sh.getDataRange().getValues();
  const hdrs = data[0].map(String);

  return data.slice(1)
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => {
      const obj = {};
      hdrs.forEach((h, i) => {
        obj[h] = (row[i] !== '' && row[i] !== null && row[i] !== undefined)
          ? String(row[i])
          : null;
      });
      return obj;
    });
}

// ── Tạo sheet nếu chưa có, thêm header ──────────────────────
function getOrCreate(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (sh.getLastRow() === 0) {
    sh.appendRow(cols);
    const hdrRange = sh.getRange(1, 1, 1, cols.length);
    hdrRange.setFontWeight('bold')
            .setBackground('#c9daf8')
            .setHorizontalAlignment('center');
    sh.setFrozenRows(1);

    // Tự động điều chỉnh độ rộng cột
    for (let i = 1; i <= cols.length; i++) {
      sh.setColumnWidth(i, 140);
    }
  }
  return sh;
}

// ── CRUD ─────────────────────────────────────────────────────
function addRow(ss, name, data, cols) {
  const sh = getOrCreate(ss, name, cols);
  sh.appendRow(cols.map(c => (data[c] != null ? data[c] : '')));
}

function updateRow(ss, name, data, cols) {
  const sh = getOrCreate(ss, name, cols);
  if (sh.getLastRow() <= 1) return;

  const vals  = sh.getDataRange().getValues();
  const idIdx = vals[0].map(String).indexOf('id');

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][idIdx]) === String(data.id)) {
      sh.getRange(i + 1, 1, 1, cols.length)
        .setValues([cols.map(c => (data[c] != null ? data[c] : ''))]);
      return;
    }
  }
}

function deleteRow(ss, name, id) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() <= 1) return;

  const vals  = sh.getDataRange().getValues();
  const idIdx = vals[0].map(String).indexOf('id');

  // Xóa từ dưới lên để không bị lệch index
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][idIdx]) === String(id)) {
      sh.deleteRow(i + 1);
      return;
    }
  }
}

// ── Trả về JSON response ─────────────────────────────────────
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
