/**
 * Apps Script STANDALONE "bayich2_dathang" — backend CHỈ ĐỌC của trang Đặt Hàng Vân Bao Bì (bayich2-dathang.vercel.app).
 *
 *  ĐỌC : doPost {hanhDong:'retail', pin} → các mặt hàng Retail để đặt hàng — cần Mã PIN chung (tab CauHinh).
 *        {hanhDong:'kiemPin', pin} → trang kiểm PIN ngay lúc nhập. GET chỉ có ping. KHÔNG có lệnh ghi.
 *
 * Chỉ trả đúng các cột việc đặt hàng cần: tên, Đơn Vị Sỉ (Thùng / Bành… — trống thì đặt theo Đơn Vị Lẻ: Kg, Gói…),
 * Số Lượng (đơn vị lẻ trong 1 đơn vị sỉ) và Giá Nhập Sỉ (giá lần nhập trước). KHÔNG trả % lãi, giá bán.
 * Cột tìm theo CHỮ TIÊU ĐỀ, không theo vị trí: chèn / đổi thứ tự cột Retail không làm đọc sai.
 *
 * CÀI / CẤP QUYỀN: chọn hàm caiDat → Run (chỉ đọc thử, không ghi gì).
 * CẬP NHẬT CODE: clasp push → clasp deploy -i <deploymentId> (link /exec giữ nguyên).
 */

var ID_BAYICH2 = '1Wd4Zvq2xiIiEzou_dvE2YtOk-bJJhAD7se0yREYe9c8';   // sheet có tab Retail + CauHinh

var TAB_RETAIL = 'Retail';
var TAB_CAU_HINH = 'CauHinh';
var TD_TRUONG = 'Trường', TD_GIA_TRI = 'Giá trị';

var COT_BAT_BUOC = { ten: 'Mặt Hàng', giaNhapSi: 'Giá Nhập Sỉ', soLuong: 'Số Lượng' };
var COT_PHU = { donViSi: 'Đơn Vị Sỉ', donViLe: 'Đơn Vị Lẻ' };   // có thì dùng

/* ══════════════════ ĐỌC (cần Mã PIN chung) ══════════════════ */
function doGet(e) {
  try {
    return traLoi({ ok: true, ten: 'bayich2_dathang', thoiGian: new Date().toISOString() });
  } catch (err) {
    return traLoi({ ok: false, loi: String(err) });
  }
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var p = kiemPin(SpreadsheetApp.openById(ID_BAYICH2), d.pin);
    if (!p.ok) return traLoi(p);
    if (d.hanhDong === 'kiemPin') return traLoi({ ok: true });
    if (d.hanhDong === 'retail') return traLoi(docRetailDatHang());
    return traLoi({ ok: false, loi: 'Hành động không hợp lệ' });
  } catch (err) {
    return traLoi({ ok: false, loi: String(err) });
  }
}

/* { ok, mon:[{ ten, donViSi, donViLe, soLuong, giaNhapSi }], thoiGian } */
function docRetailDatHang() {
  var ss = SpreadsheetApp.openById(ID_BAYICH2);
  var sh = ss.getSheetByName(TAB_RETAIL);
  if (!sh) return { ok: false, loi: 'Không tìm thấy tab ' + TAB_RETAIL };
  var hang = sh.getDataRange().getValues();
  if (!hang.length) return { ok: false, loi: 'Tab ' + TAB_RETAIL + ' đang trống' };

  var cot = {}, thieu = [];
  Object.keys(COT_BAT_BUOC).forEach(function (k) {
    cot[k] = timCot(hang[0], COT_BAT_BUOC[k]);
    if (cot[k] < 0) thieu.push(COT_BAT_BUOC[k]);
  });
  if (thieu.length) return { ok: false, loi: 'Tab ' + TAB_RETAIL + ' thiếu cột: ' + thieu.join(', ') + ' — kiểm tra lại tên cột ở dòng tiêu đề' };
  Object.keys(COT_PHU).forEach(function (k) { cot[k] = timCot(hang[0], COT_PHU[k]); });

  var ds = [];
  for (var i = 1; i < hang.length; i++) {
    var d = hang[i], ten = String(d[cot.ten] || '').trim();
    if (!ten) continue;
    var gia = soHoa(d[cot.giaNhapSi]);
    if (gia == null) continue;
    ds.push({
      ten: ten,
      donViSi: cot.donViSi < 0 ? '' : String(d[cot.donViSi] || '').trim(),
      donViLe: cot.donViLe < 0 ? '' : String(d[cot.donViLe] || '').trim(),
      soLuong: soThuc(d[cot.soLuong]) || 1,
      giaNhapSi: gia
    });
  }
  return { ok: true, mon: ds, thoiGian: new Date().toISOString() };
}

/* ══════════════════ CÀI / CẤP QUYỀN ══════════════════ */
function caiDat() {
  var kq = docRetailDatHang();
  Logger.log(kq.ok ? ('Đọc thử: ' + kq.mon.length + ' mặt hàng Retail để đặt hàng') : ('Lỗi: ' + kq.loi));
}

/* ══════════════════ phụ trợ (chép từ bayich2_tinhgia — chỉ phần đọc) ══════════════════ */
/* { chuanHoa(tên trường): giá trị } của tab CauHinh — chưa có tab thì rỗng */
function docCauHinhChung(ss) {
  var sh = ss.getSheetByName(TAB_CAU_HINH), gt = {};
  if (!sh) return gt;
  var hang = sh.getDataRange().getValues(), td = hang[0] || [];
  var cTr = timCot(td, TD_TRUONG), cGt = timCot(td, TD_GIA_TRI);
  if (cTr < 0 || cGt < 0) return gt;
  for (var i = 1; i < hang.length; i++) { var t = chuanHoa(hang[i][cTr]); if (t && !(t in gt)) gt[t] = hang[i][cGt]; }
  return gt;
}

function layTheoTen(gt, ds) {
  for (var i = 0; i < ds.length; i++) { var k = chuanHoa(ds[i]); if (k in gt) return gt[k]; }
  return undefined;
}

/* Mã PIN chung (tab CauHinh) — sai / thiếu thì chờ 2 giây như sổ bán hàng (chống dò PIN). Chép từ bayich2_doichieutoa. */
var TRUONG_PIN = 'Mã PIN chung';
function kiemPin(ss, pin) {
  var dung = layTheoTen(docCauHinhChung(ss), [TRUONG_PIN]);
  if (dung === undefined || String(dung).trim() === '')
    return { ok: false, maLoi: 'THIEU_PIN', loi: 'Chưa có "' + TRUONG_PIN + '" trong tab ' + TAB_CAU_HINH + ' — chưa đọc được' };
  if (String(pin == null ? '' : pin).trim() !== String(dung).trim()) {
    Utilities.sleep(2000);
    return { ok: false, maLoi: 'PIN', loi: 'Sai mã PIN — xem ô "' + TRUONG_PIN + '" ở tab ' + TAB_CAU_HINH };
  }
  return { ok: true };
}

/* Số từ ô Sheet: số giữ nguyên, chữ kiểu "17.000" / "0,5" / "10%" thì bóc ra. Trống → null */
function soThuc(v) {
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function soHoa(v) {
  var n = soThuc(v);
  return n == null ? null : Math.round(n);
}

function chuanHoa(s) {
  return String(s == null ? '' : s)
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function timCot(tieuDe, ten) {
  var can = chuanHoa(ten);
  for (var i = 0; i < tieuDe.length; i++) if (chuanHoa(tieuDe[i]) === can) return i;
  return -1;
}

function traLoi(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
