# Sổ CSKH SmartHome

Website thống kê và quản lý danh sách khách hàng CSKH, có đăng nhập và phân quyền. Dữ liệu được **mã hoá AES-256** trước khi đưa lên GitHub. Không có tài khoản thì không đọc được, kể cả khi tải thẳng file trong repo.

## Bảo mật hoạt động thế nào

| File | Lên GitHub? | Nội dung |
|---|---|---|
| `data/Report_CSKH.xlsx` | **Không** (đã chặn bằng `.gitignore`) | File Excel gốc, chỉ nằm trên máy bạn |
| `secrets.local.json` | **Không** (đã chặn bằng `.gitignore`) | Username, mật khẩu và khoá dữ liệu. **Giữ kỹ, không gửi cho ai** |
| `data/Report_CSKH.enc` | Có | File Excel đã mã hoá (AES-256-GCM) |
| `data/keys.json` | Có | Khoá dữ liệu được bọc riêng bằng mật khẩu từng tài khoản (PBKDF2-SHA256, 600.000 vòng). Username chỉ lưu dạng mã băm |

- Khi đăng nhập, trình duyệt dùng mật khẩu để mở khoá dữ liệu. Sai mật khẩu thì không giải mã được.
- Chỉnh sửa lưu trong trình duyệt cũng được mã hoá.
- Web tự đăng xuất sau 60 phút không thao tác, hoặc khi đóng tab.

## Tài khoản và quyền

| Quyền | Xem | Thêm / sửa KH | Xuất Excel | Xoá KH | Bỏ chỉnh sửa |
|---|---|---|---|---|---|
| **admin** (Quản trị) | ✓ | ✓ | ✓ | ✓ | ✓ |
| **editor** (Biên tập) | ✓ | ✓ | ✓ | – | – |
| **viewer** (Chỉ xem) | ✓ | – | – | – | – |

Xem lại tài khoản:

```bash
node tools/encrypt.js --show
```

Đổi toàn bộ mật khẩu. Lệnh này tạo tài khoản mới và mã hoá lại; mật khẩu cũ hết hiệu lực sau khi đẩy lên GitHub:

```bash
node tools/encrypt.js --new-passwords
```

> Phân quyền được áp dụng trên giao diện web. Dữ liệu gốc trên GitHub chỉ thay đổi được bởi người có quyền push vào repo, nên dù ai chỉnh sửa trên web cũng không ảnh hưởng tới file trên GitHub.

## Cập nhật dữ liệu (thay file Excel)

1. Chép file Excel mới đè lên `data/Report_CSKH.xlsx`, giữ nguyên tên.
2. Mã hoá lại:

   ```bash
   node tools/encrypt.js
   ```

3. Đẩy lên GitHub:

   ```bash
   git add data/Report_CSKH.enc data/keys.json
   git commit -m "Cập nhật dữ liệu CSKH"
   git push
   ```

4. Tải lại trang. GitHub Pages có thể mất 1–2 phút mới cập nhật.

Về file Excel:
- File cần có một sheet chứa cột "Điện thoại" và "Trạng thái". Web nhận cột theo **tên tiêu đề**, không theo vị trí.
- **Chỉnh sửa trên web** chỉ lưu trong trình duyệt của máy đó. Muốn giữ lại thì bấm **Xuất Excel**, rồi dùng file xuất ra thay `data/Report_CSKH.xlsx` và mã hoá lại.
- Nếu file dữ liệu bị thay trong khi trên web còn chỉnh sửa chưa xuất, web hiện nút **Tải bản chỉnh sửa cũ**.

## Cách chạy

- **GitHub Pages**: vào repo → *Settings → Pages → Branch: `main` / `(root)` → Save*. Web chạy tại `https://<tài-khoản>.github.io/<tên-repo>/`.
- **Trên máy** (cần Node.js): chạy lệnh dưới đây rồi mở `http://localhost:8080`.

  ```bash
  node server.js
  ```

  Chức năng mã hoá của trình duyệt chỉ chạy trên `https://` hoặc `localhost`. Truy cập từ máy khác qua `http://<IP>` sẽ không đăng nhập được, hãy dùng GitHub Pages.

## Chức năng

1. **Tổng quan**: KPI so với cùng kỳ năm trước; biểu đồ và bảng theo tháng, trạng thái, nguồn, loại KH, lý do kết thúc, phân khúc, chiến dịch QC, khu vực, người phụ trách, hợp đồng.
   - Mỗi bảng có khung **Nhận định** tự phát hiện số liệu bất thường.
   - Mọi bảng **sắp xếp được** khi bấm tiêu đề cột.
2. **Danh sách KH**: tìm kiếm, lọc (kể cả theo vấn đề dữ liệu), sửa, thêm ghi chú, xoá.
3. **Kiểm tra dữ liệu**: các dòng lỗi, cảnh báo hoặc lưu ý, kèm đối chiếu với sheet Thống kê của Excel.
4. **+ Thêm khách hàng**.
5. **Xuất Excel**: file gồm các phần sau:
   - Giữ nguyên các cột và sheet gốc.
   - Thêm sheet "Thống kê (web)" và "Kiểm tra dữ liệu (web)".

## Cấu trúc

| File | Vai trò |
|---|---|
| `index.html`, `style.css` | Giao diện |
| `app.js` | Đăng nhập, giải mã, thống kê, nhận định, kiểm tra dữ liệu, form, xuất Excel |
| `tools/encrypt.js` | Tạo tài khoản và mã hoá dữ liệu |
| `data/Report_CSKH.enc`, `data/keys.json` | Dữ liệu đã mã hoá |
| `vendor/xlsx.full.min.js` | Thư viện SheetJS 0.18.5 |
| `server.js` | Web server chạy trên máy; tự chặn không phục vụ file `.xlsx` và `secrets.local.json` |
