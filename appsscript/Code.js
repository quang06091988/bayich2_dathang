/**
 * Apps Script STANDALONE "bayich2_dathang" — backend CHỈ ĐỌC của trang Đặt Hàng Vân Bao Bì (bayich2-dathang.vercel.app).
 *
 *  ĐỌC : doPost {hanhDong:'retail', pin} → các mặt hàng Retail để đặt hàng — cần Mã PIN chung (tab CauHinh).
 *        {hanhDong:'kiemPin', pin} → trang kiểm PIN ngay lúc nhập. GET chỉ có ping. KHÔNG có lệnh ghi.
 *
 * Chỉ trả đúng các cột việc đặt hàng cần: tên, Đơn Vị Sỉ (Thùng / Bành… — trống thì đặt theo Đơn Vị Lẻ: Kg, Gói…),
 * Số Lượng (đơn vị lẻ trong 1 đơn vị sỉ) và Giá Nhập Sỉ (giá lần nhập trước). KHÔNG trả % lãi, giá bán.
 * Kèm số Zalo Vân Bao Bì (dòng "Zalo Vân Bao Bì" ở tab CauHinh).
 * Cột tìm theo CHỮ TIÊU ĐỀ, không theo vị trí: chèn / đổi thứ tự cột Retail không làm đọc sai.
 *
 * CÀI / CẤP QUYỀN: chọn hàm caiDat → Run (thêm dòng "Zalo Vân Bao Bì" vào CauHinh nếu chưa có, rồi đọc thử).
 * CẬP NHẬT CODE: clasp push → clasp deploy -i <deploymentId> (link /exec giữ nguyên).
 */

var ID_BAYICH2 = '1Wd4Zvq2xiIiEzou_dvE2YtOk-bJJhAD7se0yREYe9c8';   // sheet có tab Retail + CauHinh

var TAB_RETAIL = 'Retail';
var TAB_CAU_HINH = 'CauHinh';
var TD_TRUONG = 'Trường', TD_GIA_TRI = 'Giá trị';

/* Số Zalo nhà phân phối: dòng "Zalo Vân Bao Bì" trong tab CauHinh (sheet bayich2) — sửa số ở đó. */
var TRUONG_ZALO = 'Zalo Vân Bao Bì';

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

/* { ok, mon:[{ ten, donViSi, donViLe, soLuong, giaNhapSi }], zaloVanBaoBi, thoiGian } */
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
  var zalo = layTheoTen(docCauHinhChung(ss), [TRUONG_ZALO]);
  return { ok: true, mon: ds, zaloVanBaoBi: zalo == null ? '' : String(zalo).trim(), thoiGian: new Date().toISOString() };
}

/* ══════════════════ CÀI / CẤP QUYỀN ══════════════════ */
function caiDat() {
  themDongZalo(SpreadsheetApp.openById(ID_BAYICH2));
  var kq = docRetailDatHang();
  Logger.log(kq.ok ? ('Đọc thử: ' + kq.mon.length + ' mặt hàng Retail để đặt hàng · Zalo Vân Bao Bì: ' + (kq.zaloVanBaoBi || 'CHƯA CÓ — điền ở tab ' + TAB_CAU_HINH))
    : ('Lỗi: ' + kq.loi));
}

/* Thêm dòng "Zalo Vân Bao Bì" (để trống) vào cuối tab CauHinh nếu chưa có. Đã có dòng thì giữ nguyên, không ghi gì. */
function themDongZalo(ss) {
  var sh = ss.getSheetByName(TAB_CAU_HINH);
  if (!sh) { Logger.log('Chưa có tab ' + TAB_CAU_HINH + ' — chưa thêm được dòng "' + TRUONG_ZALO + '"'); return false; }
  if (layTheoTen(docCauHinhChung(ss), [TRUONG_ZALO]) !== undefined) { Logger.log('Tab ' + TAB_CAU_HINH + ' đã có dòng "' + TRUONG_ZALO + '" — giữ nguyên'); return false; }
  var hang = sh.getDataRange().getValues(), td = hang[0] || [];
  var cTr = timCot(td, TD_TRUONG), cGt = timCot(td, TD_GIA_TRI), cGc = timCot(td, 'Ghi chú'), cDc = timCot(td, 'Dùng cho');
  if (cTr < 0 || cGt < 0) { Logger.log('Tab ' + TAB_CAU_HINH + ' thiếu tiêu đề Trường / Giá trị — chưa thêm được dòng "' + TRUONG_ZALO + '"'); return false; }
  var cuoi = 1;
  for (var i = 1; i < hang.length; i++) if (String(hang[i][cTr] || '').trim()) cuoi = i + 1;
  var r = cuoi + 1;
  sh.getRange(r, cTr + 1).setValue(TRUONG_ZALO);
  sh.getRange(r, cGt + 1).setNumberFormat('@').setValue('');   // dạng chữ để giữ số 0 đầu khi điền số
  if (cGc >= 0) sh.getRange(r, cGc + 1).setValue('Số Zalo của nhà phân phối — trang Đặt hàng mở Zalo tới số này. Ghi như số điện thoại (vd 0909 123 456)');
  if (cDc >= 0) sh.getRange(r, cDc + 1).setValue('Đặt hàng');
  Logger.log('  + ' + TRUONG_ZALO + ' — để trống, điền số Zalo vào tab ' + TAB_CAU_HINH);
  return true;
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
  if (dung === undefined || chuanPin(dung) === '')
    return { ok: false, maLoi: 'THIEU_PIN', loi: 'Chưa có "' + TRUONG_PIN + '" trong tab ' + TAB_CAU_HINH + ' — chưa đọc được' };
  if (chuanPin(pin) !== chuanPin(dung)) {
    Utilities.sleep(2000);
    return { ok: false, maLoi: 'PIN', loi: 'Sai mã PIN — xem ô "' + TRUONG_PIN + '" ở tab ' + TAB_CAU_HINH };
  }
  return { ok: true };
}

/* So PIN giống sổ bán hàng (bayich2_pos/appsscript): bỏ mọi khoảng trắng và số 0 đầu — ô PIN bị Sheet đổi thành số vẫn khớp */
function chuanPin(s) { return String(s == null ? '' : s).replace(/\s+/g, '').replace(/^0+(?=\d)/, ''); }

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
