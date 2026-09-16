# bayich2_dathang — Đặt Hàng Vân Bao Bì

Trang nội bộ để Tạp Hóa Bảy Ích 2 đặt hàng với nhà phân phối **Vân Bao Bì**: chọn số thùng / bành (ly, nắp, ống hút)
hoặc số ký / gói (bị, muỗng), xem tạm tính theo giá lần nhập trước, rồi **gửi đơn qua Zalo**.
Trang chỉ đọc, **không bao giờ ghi vào Sheet**. Mở từ mục "Công cụ" trong POS.

## Dữ liệu

- Mặt hàng: tab **Retail** (sheet bayich2) qua Apps Script riêng, **chỉ đọc, cần Mã PIN chung** (tab `CauHinh`).
  Máy chủ tìm cột theo tiêu đề và chỉ trả `Mặt Hàng`, `Đơn Vị Sỉ`, `Đơn Vị Lẻ`, `Số Lượng`, `Giá Nhập Sỉ` —
  không trả % lãi hay giá bán.
- Đơn vị đặt: `Đơn Vị Sỉ` nếu có (Thùng, Bành), không thì `Đơn Vị Lẻ` (Kg cho gõ số lẻ như 0,5; Gói…).
- Số Zalo của Vân Bao Bì: ô "Số điện thoại" trong tab **Info** của sheet vanbaobi (đã xuất bản công khai).
- Trong máy chỉ lưu: Mã PIN chung, giỏ đang chọn và 5 đơn gần nhất (tên + số lượng) — không lưu bảng giá nhập.

## Cấu trúc

```
index.html        trang web (Vercel: bayich2-dathang.vercel.app)
og.png            ảnh xem trước khi chia sẻ (1200×630, không có tên miền, số minh hoạ)
.clasp.json       Script ID của backend, rootDir = appsscript
appsscript/       backend — Apps Script STANDALONE "bayich2_dathang" (không deploy lên Vercel)
  Code.js         doPost {hanhDong:'retail'|'kiemPin', pin} — không có lệnh ghi
  appsscript.json
```

- Script: https://script.google.com/d/12i4PUpHiyQGVE6HbGZ5RasVvvV0xmpEWydssAdObz-PjQ_0lE4m9jVEg/edit
- Deployment (`API` trong index.html): `AKfycbxN7DJTtoGzdk1jVHO3xR_Lz9USEm-veq9rzrhRoPyX_mbtfRn2LiigGJVpIaXPnRU36A`
- Cập nhật backend: `clasp push` → `clasp deploy -i <deploymentId>` (link /exec giữ nguyên).
- Lần đầu: mở script → chọn hàm `caiDat` → Run → cấp quyền đọc Sheet (hàm chỉ đọc thử).
- Các hàm phụ trợ (`kiemPin`, `docCauHinhChung`…) chép từ `bayich2_tinhgia` — sửa thì sửa cả hai.
