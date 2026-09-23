# Sổ CSKH SmartHome

Web quản lý và thống kê khách hàng CSKH:

- **Giao diện**: trang tĩnh trong `public/`.
- **Server**: Vercel Functions trong `api/`, chạy trên gói Vercel Hobby.
- **Database**: Postgres (Neon, gói miễn phí 0,5 GB).

Dữ liệu nằm trong database. Mọi người cùng xem và sửa trên một bản duy nhất, mọi thay đổi đều được lưu lịch sử, và có thể **xuất Excel** bất cứ lúc nào.

## Kiến trúc

```
Trình duyệt ──► Vercel ── public/     (giao diện: index.html, assets/, vendor/)
                       └─ api/*.js    (đăng nhập, dữ liệu, nhập Excel, tài khoản)
                              │
                              ▼
                       Neon Postgres  (khách hàng, lịch sử chỉnh sửa, tài khoản, file mẫu Excel)
```

**Phân quyền kiểm tra ở server**:

| Quyền | Xem | Thêm / sửa KH | Xuất Excel | Xoá KH | Nhập Excel, quản lý tài khoản |
|---|---|---|---|---|---|
| Quản trị (admin) | ✓ | ✓ | ✓ | ✓ | ✓ |
| Biên tập (editor) | ✓ | ✓ | ✓ | – | – |
| Chỉ xem (viewer) | ✓ | – | – | – | – |

**Bảo mật**:
- Mật khẩu được băm bằng scrypt.
- Phiên đăng nhập lưu trong cookie HttpOnly, hết hạn sau 12 giờ; web tự đăng xuất sau 60 phút không thao tác.
- Sai mật khẩu quá 8 lần trong 15 phút thì tài khoản bị tạm chặn.
- Có chống CSRF.
- Khoá tài khoản hoặc đặt lại mật khẩu thì phiên cũ của người đó mất hiệu lực ngay.

## Triển khai lần đầu

### 1. Đẩy code lên GitHub
Chạy trong thư mục `web`:

```bash
git add -A
git status
git commit -m "Chuyển sang database + server"
git push
```

Sau khi chạy `git status`, danh sách **không được có** `data/`, file `.xlsx` hay `secrets.local.json`. Các file này đã bị chặn bằng `.gitignore`.

### 2. Cấu hình project trên Vercel
Vào **Settings → Build and Deployment**:
- **Framework Preset**: `Other`
- **Root Directory**: để trống
- **Build Command / Output Directory / Install Command**: để trống, không bật Override (`vercel.json` đã cấu hình sẵn)

### 3. Tạo database Neon miễn phí
1. Vào tab **Storage** → **Create Database**.
2. Chọn **Neon** (Serverless Postgres) → **Continue**.
3. Region: **Singapore (ap-southeast-1)**. Plan: **Free**. Đặt tên, ví dụ `cskh-db` → **Create**.
4. Ở bước **Connect to Project**, chọn project này và tích cả **Production, Preview, Development** → **Connect**.

Vercel tự thêm biến `DATABASE_URL` vào project. Bảng dữ liệu được tạo tự động ở lần chạy đầu.

### 4. Thêm biến môi trường
Vào **Settings → Environment Variables** → **Add**, áp dụng cho cả Production và Preview:

| Tên | Giá trị |
|---|---|
| `AUTH_SECRET` | Chuỗi ngẫu nhiên ≥ 32 ký tự (dùng để ký phiên đăng nhập) |
| `ADMIN_USERNAME` | Tên đăng nhập admin đầu tiên, ví dụ `admin.t7s62tcvqe` |
| `ADMIN_PASSWORD` | Mật khẩu admin đầu tiên (≥ 12 ký tự) |

Tạo `AUTH_SECRET` bằng lệnh:

```bash
node -e "console.log(require('crypto').randomBytes(32).toString('base64url'))"
```

`ADMIN_USERNAME` và `ADMIN_PASSWORD` **chỉ dùng cho lần đăng nhập đầu tiên**, khi database chưa có tài khoản nào. Sau đó đổi mật khẩu trong web: tab *Tài khoản → Đặt lại mật khẩu*.

### 5. Deploy lại
Vào **Deployments** → bấm **⋯** ở lần deploy mới nhất → **Redeploy**. Việc này cần làm vì biến môi trường chỉ có hiệu lực từ lần deploy sau khi thêm.

### 6. Nhập dữ liệu
1. Mở web → đăng nhập bằng `ADMIN_USERNAME` / `ADMIN_PASSWORD`.
2. Bấm **Nhập Excel** → chọn file `Report_CSKH….xlsx`. Khoảng 10–20 giây là xong.
3. Vào tab **Tài khoản & hệ thống** → tạo tài khoản cho từng người, chọn quyền phù hợp.
   - Mật khẩu ngẫu nhiên chỉ hiện **một lần**. Bấm *Sao chép* rồi gửi riêng cho người đó.

## Sử dụng hằng ngày

- **Sửa hoặc thêm khách hàng**: bấm **Lưu** là ghi thẳng vào database.
  - Người khác thấy thay đổi trong vòng khoảng 45 giây, hoặc ngay khi quay lại tab web.
  - Nếu 2 người sửa cùng lúc một khách hàng, người lưu sau được báo và được tải bản mới nhất, không ai ghi đè mất dữ liệu của ai.
- **Lịch sử**: mở một khách hàng → **Xem lịch sử chỉnh sửa** để biết ai sửa gì, lúc nào, giá trị trước và sau.
- **Xuất Excel**: file có đủ cột như file gốc, các sheet gốc (kèm công thức), nhật ký tuần, cùng sheet "Thống kê (web)" và "Kiểm tra dữ liệu (web)".
- **Nhập Excel lại** (admin): **thay thế toàn bộ** dữ liệu hiện tại. Nên bấm **Xuất Excel** trước để giữ bản đang có.
- **Dung lượng**: tab *Tài khoản & hệ thống* hiển thị dung lượng đã dùng và nút dọn lịch sử chỉnh sửa cũ.

## Chạy thử trên máy
Cần Node.js 20 trở lên:

```bash
npm install
node tools/server.js
```

Mở http://localhost:8080 và đăng nhập `admin` / `admin-local-123`. Khi chạy trên máy, web dùng database cục bộ (PGlite) lưu ở `~/.cskh-localdb`, không ảnh hưởng dữ liệu thật. Muốn chạy trên máy nhưng nối vào database thật trên Neon thì đặt biến `DATABASE_URL` trước khi chạy.

## Giới hạn gói miễn phí

- **Neon Free**:
  - Dung lượng 0,5 GB. 2.500 khách hàng chiếm khoảng 7 MB, nên còn chứa được khoảng 180.000 khách hàng nữa.
  - Database tự "ngủ" khi không có ai dùng, nên lần mở đầu tiên sau đó có thể chậm thêm 1–2 giây.
- **Vercel Hobby**: chỉ dành cho mục đích **phi thương mại** (theo điều khoản của Vercel). Dùng chính thức cho công ty thì nên nâng lên gói Pro.

## Cấu trúc

| Đường dẫn | Vai trò |
|---|---|
| `public/index.html`, `public/assets/` | Giao diện, thống kê, nhận định, kiểm tra dữ liệu, form, xuất Excel |
| `public/vendor/xlsx.full.min.js` | Thư viện SheetJS đọc/ghi Excel |
| `api/login.js`, `logout.js`, `me.js` | Đăng nhập / đăng xuất / người dùng hiện tại |
| `api/data.js` | Đọc dữ liệu (phân trang) và đồng bộ thay đổi |
| `api/record.js` | Thêm / sửa / xoá khách hàng, xem lịch sử |
| `api/import.js`, `api/template.js` | Nhập Excel vào database; file mẫu để xuất Excel |
| `api/users.js`, `api/stats.js` | Quản lý tài khoản; dung lượng và dọn lịch sử |
| `api/_lib/` | Kết nối database, xác thực, tiện ích HTTP |
| `tools/server.js` | Server chạy thử trên máy |
| `vercel.json` | Cấu hình Vercel (thư mục `public/`, API, header bảo mật) |
