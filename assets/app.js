"use strict";
/* Sổ CSKH SmartHome — chạy hoàn toàn trên trình duyệt.
   Import Excel → thống kê + nhận định + kiểm tra dữ liệu → sửa/thêm KH → xuất lại Excel
   (giữ nguyên các cột và sheet khác của file gốc). */

// ================= Nguồn dữ liệu =================
// Web đọc file dữ liệu ĐÃ MÃ HOÁ (tạo bằng: node tools/encrypt.js từ data/Report_CSKH.xlsx).
// Muốn cập nhật dữ liệu: chép file Excel mới đè lên data/Report_CSKH.xlsx, chạy lại tools/encrypt.js rồi đẩy lên GitHub.
const DATA_FILE = "data/Report_CSKH.enc";
const KEYS_FILE = "data/keys.json";
const IDLE_MIN = 60;
const STATE_VERSION = 3; // tăng khi đổi cách đọc file để trình duyệt đọc lại từ đầu

// ================= Danh mục =================
// Cột lõi của sheet danh sách KH: [key, tiêu đề cột trong Excel]
const CORE = [
  ["checkTrung", "Check trùng theo SĐT"], ["nam", "Năm"], ["acc", "ACC PT"],
  ["dlChuyen", "Tên Đại lý được chuyển KH"], ["doiTac", "Tên Đối tác giới thiệu công trình"],
  ["ngay", "Ngày nhận"], ["tuan", "Tuần"], ["thang", "Tháng"], ["nguon", "Nguồn"],
  ["nguonQC", "Nguồn QC"], ["chienDich", "Chiến dịch QC"], ["link", "Link bài viết"],
  ["bietBkav", "KH biết đến Bkav chưa"], ["ten", "Tên khách hàng"], ["sdt", "Điện thoại"],
  ["email", "Email"], ["diaChi", "Địa chỉ"], ["tinh", "Tỉnh"], ["kv", "Phân Loại KV"],
  ["kvCode", "Phân loại KV"], ["thongTin", "Thông tin tiếp nhận"], ["loai", "Loại KH"],
  ["nganh", "Lĩnh vực, ngành nghề"], ["congTrinh", "Loại công trình"], ["phanKhuc", "Đánh giá phân khúc"],
  ["trangThai", "Trạng thái"], ["lyDo", "Lý do kết thúc"], ["phanLoaiDL", "Phân loại ĐL"],
  ["phanLoaiTTK", "Phân loại KH tự tìm kiếm"], ["thangKy", "Tháng dự kiến ký HĐg"],
  ["giaTri", "Giá trị báo giá/HĐg"], ["giaiPhap", "Giải pháp"], ["nhatKy", "Nhật ký"],
];
// Các cột lưu dạng số trong Excel (để công thức COUNTIFS ở sheet Thống kê vẫn đếm đúng)
const NUMERIC_KEYS = new Set(["checkTrung", "nam", "tuan", "thang", "nguonQC", "chienDich", "kvCode", "lyDo", "phanKhuc", "phanLoaiDL", "phanLoaiTTK", "thangKy", "bietBkav"]);

const KV_MAP = {"An Giang":"Nam","Bắc Cạn":"Bắc","Bắc Giang":"Bắc","Bạc Liêu":"Nam","Bắc Ninh":"Bắc","Bến Tre":"Nam","Bình Định":"Nam","Bình Dương":"Nam","Bình Phước":"Nam","Bình Thuận":"Nam","Cà Mau":"Nam","Cần Thơ":"Nam","Cao Bằng":"Bắc","Đà Nẵng":"Bắc","Đắk Lắk":"Nam","Đắk Nông":"Nam","Điện Biên":"Bắc","Đồng Nai":"Nam","Đồng Tháp":"Nam","Gia Lai":"Nam","Hà Giang":"Bắc","Hà Nam":"Bắc","Hà Tĩnh":"Bắc","Hải Dương":"Bắc","Hải Phòng":"Bắc","Hậu Giang":"Nam","HCM":"HCM","HN":"HN","Hòa Bình":"Bắc","Huế":"Bắc","Hưng Yên":"Bắc","Khánh Hòa":"Nam","Kiên Giang":"Nam","Kon Tum":"Nam","Lai Châu":"Bắc","Lâm Đồng":"Nam","Lạng Sơn":"Bắc","Lào Cai":"Bắc","Long An":"Nam","Nam Định":"Bắc","Nghệ An":"Bắc","Ninh Bình":"Bắc","Ninh Thuận":"Nam","Phú Thọ":"Bắc","Phú Yên":"Nam","Quảng Bình":"Bắc","Quảng Nam":"Bắc","Quảng Ngãi":"Bắc","Quảng Ninh":"Bắc","Quảng Trị":"Bắc","Sóc Trăng":"Nam","Sơn La":"Bắc","Tây Ninh":"Nam","Thái Bình":"Bắc","Thái Nguyên":"Bắc","Thanh Hóa":"Bắc","Tiền Giang":"Nam","Trà Vinh":"Nam","Tuyên Quang":"Bắc","Vĩnh Long":"Nam","Vĩnh Phúc":"Bắc","Vũng Tàu":"Nam","Yên Bái":"Bắc"};
const KV_KEY = Object.fromEntries(Object.entries(KV_MAP).map(([k, v]) => [k.toLowerCase(), v]));
const kvOf = t => KV_KEY[String(t || "").trim().toLowerCase()];
const LOAI = [["ĐL", "Đại lý (ĐL)"], ["KHC", "Khách hàng cuối (KHC)"], ["KTS", "Kiến trúc sư (KTS)"], ["DA", "Dự án (DA)"]];
const LOAI_COLOR = {"ĐL": "var(--s1)", "KHC": "var(--s2)", "KTS": "var(--s3)", "DA": "var(--s4)"};
const NGUON = ["FB", "HL", "Web", "TTK", "Zalo", "Mail", "Khác"];
const NGUON_LABEL = {FB: "Facebook", HL: "Hotline", Web: "Website", TTK: "Tự tìm kiếm", Zalo: "Zalo", Mail: "Email", "Khác": "Khác"};
const TT = ["Chăm sóc", "Tiềm năng", "Tham khảo", "KH đồng ý ký HĐ", "Đã ký HĐ", "Kết thúc"];
const ACTIVE = new Set(["Chăm sóc", "Tiềm năng", "KH đồng ý ký HĐ"]);
const QC = {"1": "Chiến dịch ĐL", "2": "Chiến dịch KHC", "3": "Chiến dịch module", "4": "Chiến dịch ĐL 2026"};
const LYDO = {"1": "Không nghe máy", "2": "Giá cao", "3": "Chọn hãng khác", "4": "Chỉ tham khảo", "5": "Khác"};
const PK = {"1": "Cao", "2": "Trung", "3": "Giá rẻ", "4": "Trung lập"};
const CT = ["Nhà phố", "Chung cư", "Biệt thự", "Liền kề", "Nhà ống", "Nhà cấp 4", "Nhà vườn", "Khác"];
const GP = ["Luxury", "Lite", "S8", "Module", "Khác"];
const STALE_DAYS = 180;

// Các loại vấn đề dữ liệu mà web tự phát hiện
const SEV_LABEL = {crit: "Lỗi", warn: "Cảnh báo", info: "Lưu ý", good: "Tốt"};
const SEV_ICON = {crit: "!", warn: "!", info: "i", good: "✓"};
const SEV_RANK = {crit: 0, warn: 1, info: 2};
const ISSUES = {
  NO_YEAR:        {sev: "crit", label: "Thiếu hoặc sai Năm", desc: "Cột Năm trống / không phải 20xx (VD 1900) và không suy ra được từ Ngày nhận", affects: "Mọi thống kê lọc theo năm"},
  YEAR_MISMATCH:  {sev: "crit", label: "Năm lệch với Ngày nhận", desc: "Cột Năm khác năm của Ngày nhận", affects: "Thống kê theo năm"},
  NO_MONTH:       {sev: "crit", label: "Thiếu Tháng", desc: "Có Năm nhưng không có Tháng → không hiện trên biểu đồ theo tháng", affects: "Biểu đồ theo tháng"},
  MONTH_MISMATCH: {sev: "crit", label: "Tháng lệch với Ngày nhận", desc: "Cột Tháng khác tháng của Ngày nhận", affects: "Biểu đồ theo tháng"},
  NO_LOAI:        {sev: "crit", label: "Thiếu Loại KH", desc: "Không có Loại KH (ĐL/KHC/KTS/DA)", affects: "Thống kê theo loại KH"},
  NO_STATUS:      {sev: "crit", label: "Thiếu Trạng thái", desc: "Không có trạng thái chăm sóc", affects: "Trạng thái, tỷ lệ ký"},
  KV_MISMATCH:    {sev: "crit", label: "Khu vực không khớp Tỉnh", desc: "Khu vực khác với bảng Phân loại KV theo tỉnh", affects: "Thống kê khu vực"},
  BAD_CODE:       {sev: "crit", label: "Mã ngoài danh mục", desc: "Nguồn / Nguồn QC / Lý do / Phân khúc có giá trị không thuộc danh mục", affects: "Nguồn, chiến dịch, lý do, phân khúc"},
  STATUS_SPELL:   {sev: "warn", label: "Trạng thái viết không chuẩn trong Excel", desc: "Thừa dấu cách hoặc sai hoa/thường (VD “ Chăm sóc ”) → sheet Thống kê của Excel đếm sót; web đã tự chuẩn hoá", affects: "Sheet Thống kê trong Excel"},
  YEAR_DERIVED:   {sev: "info", label: "Năm được suy ra từ Ngày nhận", desc: "Cột Năm trong Excel trống; web lấy năm từ Ngày nhận — sheet Thống kê của Excel không đếm các dòng này", affects: "Chênh lệch web ↔ Excel"},
  NO_DATE:        {sev: "warn", label: "Thiếu Ngày nhận", desc: "Không có ngày tiếp nhận KH", affects: "Xu hướng theo thời gian"},
  FUTURE_DATE:    {sev: "warn", label: "Ngày nhận ở tương lai", desc: "Ngày nhận sau ngày hôm nay — có thể gõ nhầm năm/tháng", affects: "Thống kê theo tháng"},
  PLACEHOLDER_DATE:{sev: "warn", label: "Ngày nhận nghi là ngày mặc định", desc: "Rất nhiều KH trùng một ngày nhận (VD 01/01/2025) — thường do nhập bù dữ liệu cũ, làm tháng đó cao đột biến", affects: "Biểu đồ theo tháng"},
  NO_PHONE:       {sev: "warn", label: "Thiếu số điện thoại", desc: "Không có SĐT để liên hệ và kiểm tra trùng", affects: "Chăm sóc, kiểm tra trùng"},
  BAD_PHONE:      {sev: "warn", label: "SĐT sai định dạng", desc: "Có số không đúng 10 chữ số (di động) hoặc 11 chữ số (số bàn 02x). Nhiều số thì ghi “số1;số2”", affects: "Chăm sóc, kiểm tra trùng"},
  DUP_PHONE:      {sev: "warn", label: "Trùng số điện thoại", desc: "Một trong các SĐT của KH này xuất hiện ở dòng khác → có thể đếm 1 KH nhiều lần", affects: "Tổng số KH"},
  END_NO_REASON:  {sev: "warn", label: "Kết thúc nhưng không ghi lý do", desc: "Trạng thái Kết thúc mà trống Lý do kết thúc", affects: "Bảng Lý do kết thúc"},
  REASON_NOT_END: {sev: "warn", label: "Có lý do kết thúc nhưng chưa Kết thúc", desc: "Có Lý do kết thúc nhưng trạng thái khác “Kết thúc”", affects: "Bảng Lý do kết thúc"},
  SIGNED_NO_VALUE:{sev: "warn", label: "Đã ký HĐ nhưng thiếu giá trị", desc: "Không có Giá trị báo giá/HĐ", affects: "Giá trị HĐ, doanh thu"},
  VALUE_ODD:      {sev: "warn", label: "Giá trị HĐ bất thường", desc: "Ghi dạng chữ (VD “138tr”) hoặc nhỏ hơn 1 triệu (VD “561”) — có thể thiếu số 0 / sai đơn vị", affects: "Giá trị HĐ"},
  ACC_CASE:       {sev: "warn", label: "Tên ACC viết không thống nhất", desc: "Cùng một tên nhưng khác hoa/thường (VD “HieuNTB” và “Hieungtb”) — nếu là 1 người thì thống kê đang bị tách đôi", affects: "Bảng theo người phụ trách"},
  UNKNOWN_TINH:   {sev: "warn", label: "Tỉnh không có trong danh mục", desc: "Tên tỉnh không khớp bảng Phân loại KV (viết tắt, sai chính tả…)", affects: "Thống kê khu vực"},
  NO_TINH:        {sev: "warn", label: "Thiếu Tỉnh", desc: "Không có tỉnh/thành", affects: "Thống kê khu vực"},
  STALE:          {sev: "info", label: "Chăm sóc quá 6 tháng chưa chốt", desc: "Đang Chăm sóc/Tiềm năng nhưng đã nhận hơn 180 ngày — nên đóng hoặc cập nhật trạng thái", affects: "Phễu chăm sóc"},
  FB_NO_QC:       {sev: "info", label: "KH từ Facebook chưa gắn chiến dịch QC", desc: "Nguồn FB (từ 2025) nhưng trống Nguồn QC", affects: "Bảng chiến dịch quảng cáo"},
  DL_NO_SEG:      {sev: "info", label: "Đại lý chưa đánh giá phân khúc", desc: "KH loại ĐL chưa có Đánh giá phân khúc", affects: "Bảng phân khúc đại lý"},
  NO_NAME:        {sev: "info", label: "Thiếu tên khách hàng", desc: "Chỉ có SĐT, không có tên", affects: "Tra cứu"},
};

// ================= Tiện ích =================
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}[c]));
const fmt = n => (Math.round(n * 10) / 10).toLocaleString("vi-VN");
const pctN = (a, b) => b ? a / b * 100 : null;
const pctS = v => v == null ? "—" : v.toLocaleString("vi-VN", {maximumFractionDigits: 1}) + "%";
const pct = (a, b) => pctS(pctN(a, b));
const pp = d => (d > 0 ? "+" : "") + d.toLocaleString("vi-VN", {maximumFractionDigits: 1}) + " điểm %";
const phoneKey = p => String(p || "").replace(/\D/g, "");
// Một ô SĐT có thể ghi nhiều số: "0912345678;0987654321" (dấu ; — cũng chấp nhận , / xuống dòng)
function phonesOf(v) {
  const out = [];
  for (const part of String(v || "").split(/[;,\/|\n]+/)) {
    const t = part.trim(); if (!t) continue;
    let d = t.replace(/\D/g, "");
    if (!d) { out.push(""); continue; }                 // có chữ nhưng không có số (VD "Anh Thịnh")
    if (d.length >= 19 && /\s/.test(t)) { t.split(/\s+/).map(x => x.replace(/\D/g, "")).filter(Boolean).forEach(x => out.push(x)); continue; }
    if (/^84\d{9,10}$/.test(d)) d = "0" + d.slice(2);    // +84 → 0
    out.push(d);
  }
  return out;
}
const phoneOk = p => /^0[35789]\d{8}$/.test(p) || /^02\d{9}$/.test(p) || /^0\d{9}$/.test(p); // di động 10 số, số bàn 11 số (02x)
const validPhones = v => phonesOf(v).filter(p => p.length >= 9);
const norm = s => String(s ?? "").replace(/\s+/g, " ").trim().toLowerCase();
const pad2 = n => String(n).padStart(2, "0");
const isoOf = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const dmy = iso => iso ? iso.slice(8, 10) + "/" + iso.slice(5, 7) + "/" + iso.slice(0, 4) : "";
const dm = iso => iso ? iso.slice(8, 10) + "/" + iso.slice(5, 7) : "";
const TODAY = isoOf(new Date());
const daysBetween = (a, b) => Math.round((new Date(b + "T00:00:00") - new Date(a + "T00:00:00")) / 864e5);
function weekNum(d) { // giống WEEKNUM(date) của Excel: tuần bắt đầu Chủ nhật, 1/1 là tuần 1
  const j = new Date(d.getFullYear(), 0, 1);
  const doy = Math.round((new Date(d.getFullYear(), d.getMonth(), d.getDate()) - j) / 864e5) + 1;
  return Math.floor((doy + j.getDay() - 1) / 7) + 1;
}
function parseMoney(v) { // hiểu được: 57417150 · 57,417,150 · 57.417.150 · 43286345.00000001 · 138tr · 1,2 tỷ
  if (typeof v === "number") return v;
  const m = String(v ?? "").toLowerCase().match(/(\d[\d.,]*)\s*(tỷ|ty|triệu|trieu|tr|nghìn|ngàn|k)?/); if (!m) return 0;
  const u = m[2] || "", unit = /^t(ỷ|y)$/.test(u) ? 1e9 : /^tr/.test(u) ? 1e6 : u ? 1e3 : 1;
  let n = m[1].replace(/[.,]$/, "");
  if (/^\d+(\.\d+)?$/.test(n) && !(unit === 1 && /^\d{1,3}\.\d{3}$/.test(n))) n = parseFloat(n);          // 43286345.5
  else if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(n)) n = parseFloat(n.replace(/,/g, ""));                     // 57,417,150
  else if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(n)) n = parseFloat(n.replace(/\./g, "").replace(",", "."));  // 57.417.150
  else if (/^\d+,\d+$/.test(n)) n = parseFloat(n.replace(",", "."));                                         // 1,5 tỷ
  else n = parseFloat(n.replace(/[^\d]/g, ""));
  return isNaN(n) ? 0 : n * unit;
}
const vnd = n => n >= 1e9 ? (n / 1e9).toLocaleString("vi-VN", {maximumFractionDigits: 2}) + " tỷ" : n >= 1e6 ? (n / 1e6).toLocaleString("vi-VN", {maximumFractionDigits: 1}) + " tr" : Math.round(n).toLocaleString("vi-VN") + " đ";
const newId = () => "N" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const median = a => { if (!a.length) return 0; const s = [...a].sort((x, y) => x - y), m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
const quantile = (a, q) => { const s = [...a].sort((x, y) => x - y); if (!s.length) return 0; const p = (s.length - 1) * q, lo = Math.floor(p); return s[lo] + (s[Math.min(lo + 1, s.length - 1)] - s[lo]) * (p - lo); };
const sum = a => a.reduce((x, y) => x + y, 0);
function count(arr, f) { const o = {}; arr.forEach(r => { const k = f(r) || "(trống)"; o[k] = (o[k] || 0) + 1; }); return o; }
function delta(cur, prev, {unit = "%", inverse = false} = {}) { // hiển thị ▲▼ so với kỳ trước
  if (prev == null || cur == null) return "";
  if (unit === "%") { if (!prev) return ""; const d = (cur - prev) / prev * 100; const cls = Math.abs(d) < 3 ? "flat" : (d > 0) !== inverse ? "up" : "down"; return `<span class="delta ${cls}">${d > 0 ? "▲" : d < 0 ? "▼" : "•"} ${Math.abs(d).toLocaleString("vi-VN", {maximumFractionDigits: 0})}%</span>`; }
  const d = cur - prev; const cls = Math.abs(d) < 1 ? "flat" : (d > 0) !== inverse ? "up" : "down";
  return `<span class="delta ${cls}">${d > 0 ? "▲" : d < 0 ? "▼" : "•"} ${Math.abs(d).toLocaleString("vi-VN", {maximumFractionDigits: 1})} điểm</span>`;
}

// ================= Lưu trữ (IndexedDB) =================
const idb = {
  open() { return this._p ||= new Promise((res, rej) => { const r = indexedDB.open("cskh-smarthome", 1); r.onupgradeneeded = () => r.result.createObjectStore("kv"); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); },
  async get(k) { const db = await this.open(); return new Promise((res, rej) => { const q = db.transaction("kv").objectStore("kv").get(k); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); }); },
  async set(k, v) { const db = await this.open(); return new Promise((res, rej) => { const t = db.transaction("kv", "readwrite"); t.objectStore("kv").put(v, k); t.oncomplete = () => res(); t.onerror = () => rej(t.error); }); },
  async del(k) { const db = await this.open(); return new Promise(res => { const t = db.transaction("kv", "readwrite"); t.objectStore("kv").delete(k); t.oncomplete = () => res(); }); },
};

// ================= Đăng nhập, phân quyền & mã hoá =================
/* Dữ liệu trên GitHub chỉ có bản đã mã hoá (AES-256-GCM). Mỗi tài khoản có một bản khoá dữ liệu
   được bọc bằng mật khẩu (PBKDF2-SHA256) trong data/keys.json — sai mật khẩu thì không giải mã được.
   Phân quyền admin/editor/viewer được áp dụng trên giao diện; chỉnh sửa chỉ lưu trong trình duyệt,
   file trên GitHub chỉ thay đổi được bởi người có quyền push vào repo. */
const ROLE_LABEL = {admin: "Quản trị", editor: "Biên tập", viewer: "Chỉ xem"};
let SESSION = null; // {username, role, key: CryptoKey}
const can = {
  edit: () => !!SESSION && (SESSION.role === "admin" || SESSION.role === "editor"),   // thêm / sửa KH
  del: () => !!SESSION && SESSION.role === "admin",                                     // xoá KH
  export: () => !!SESSION && (SESSION.role === "admin" || SESSION.role === "editor"), // xuất Excel
  reset: () => !!SESSION && SESSION.role === "admin",                                   // bỏ toàn bộ chỉnh sửa
};
const te = new TextEncoder(), td = new TextDecoder();
const unb64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
async function hex256(s) { return [...new Uint8Array(await crypto.subtle.digest("SHA-256", te.encode(s)))].map(b => b.toString(16).padStart(2, "0")).join(""); }
const aesKey = raw => crypto.subtle.importKey("raw", raw, "AES-GCM", false, ["encrypt", "decrypt"]);
const decryptBuf = (key, iv, ct) => crypto.subtle.decrypt({name: "AES-GCM", iv}, key, ct);
async function encryptBuf(key, data) { const iv = crypto.getRandomValues(new Uint8Array(12)); return {iv, ct: await crypto.subtle.encrypt({name: "AES-GCM", iv}, key, data)}; }
async function decryptDataFile(buf) {
  const u = new Uint8Array(buf);
  if (td.decode(u.slice(0, 5)) !== "CSKH1") throw new Error("File dữ liệu không đúng định dạng đã mã hoá (hãy tạo bằng node tools/encrypt.js)");
  return decryptBuf(SESSION.key, u.slice(5, 17), u.slice(17));
}
// Trạng thái lưu trong trình duyệt (IndexedDB) cũng được mã hoá bằng khoá dữ liệu
const ser = o => JSON.stringify(o, function (k, v) { const raw = this[k]; return raw instanceof Date ? {__d: raw.getTime()} : v; });
const deser = s => JSON.parse(s, (k, v) => v && typeof v === "object" && v.__d !== undefined && Object.keys(v).length === 1 ? new Date(v.__d) : v);
async function saveEnc(name, obj) { const {iv, ct} = await encryptBuf(SESSION.key, te.encode(ser(obj))); await idb.set(name, {iv, ct}); }
async function loadEnc(name) { const e = await idb.get(name); if (!e || !e.iv) return null; try { return deser(td.decode(await decryptBuf(SESSION.key, e.iv, e.ct))); } catch (err) { return null; } }

async function login(username, password) {
  const res = await fetch(KEYS_FILE + "?t=" + Date.now(), {cache: "no-store"});
  if (!res.ok) throw new Error(`Không tải được ${KEYS_FILE} (lỗi ${res.status}). Hãy chạy node tools/encrypt.js`);
  const K = await res.json();
  const id = await hex256("cskh-user:" + username.trim().toLowerCase());
  const hit = K.users.find(x => x.id === id);
  const base = await crypto.subtle.importKey("raw", te.encode(password), "PBKDF2", false, ["deriveKey"]);
  const salt = hit ? unb64(hit.salt) : crypto.getRandomValues(new Uint8Array(16)); // vẫn tính khi sai username để không lộ username nào tồn tại
  const wrap = await crypto.subtle.deriveKey({name: "PBKDF2", salt, iterations: K.iter, hash: "SHA-256"}, base, {name: "AES-GCM", length: 256}, false, ["decrypt"]);
  const fail = new Error("Sai tên đăng nhập hoặc mật khẩu");
  if (!hit) throw fail;
  let info; try { info = JSON.parse(td.decode(await decryptBuf(wrap, unb64(hit.iv), unb64(hit.ct)))); } catch (e) { throw fail; }
  SESSION = {username: info.username, role: info.role, key: await aesKey(unb64(info.dataKey))};
  try { sessionStorage.setItem("cskh.session", JSON.stringify({u: info.username, r: info.role, k: info.dataKey, t: Date.now()})); } catch (e) {}
}
async function restoreSession() {
  try {
    const s = JSON.parse(sessionStorage.getItem("cskh.session") || "null"); if (!s) return false;
    if (Date.now() - (s.t || 0) > IDLE_MIN * 60e3) { sessionStorage.removeItem("cskh.session"); sessionStorage.setItem("cskh.msg", "Phiên làm việc đã hết hạn, vui lòng đăng nhập lại."); return false; }
    SESSION = {username: s.u, role: s.r, key: await aesKey(unb64(s.k))}; return true;
  } catch (e) { return false; }
}
async function logout(msg) {
  if (S && SESSION) { clearTimeout(saveT); try { await saveEnc("state", S); } catch (e) {} }
  try { sessionStorage.removeItem("cskh.session"); if (msg) sessionStorage.setItem("cskh.msg", msg); } catch (e) {}
  location.reload();
}
// Tự đăng xuất khi không thao tác quá IDLE_MIN phút
let lastAct = Date.now(), lastTouch = 0;
["mousemove", "keydown", "click", "touchstart", "scroll"].forEach(ev => document.addEventListener(ev, () => {
  lastAct = Date.now();
  if (SESSION && lastAct - lastTouch > 60e3) { lastTouch = lastAct; try { const s = JSON.parse(sessionStorage.getItem("cskh.session")); if (s) { s.t = lastAct; sessionStorage.setItem("cskh.session", JSON.stringify(s)); } } catch (e) {} }
}, {passive: true}));
setInterval(() => { if (SESSION && Date.now() - lastAct > IDLE_MIN * 60e3) logout("Đã tự đăng xuất sau " + IDLE_MIN + " phút không thao tác."); }, 30e3);

function showLogin() {
  $("#statusBox").hidden = true; $("#loginBox").hidden = false;
  const m = sessionStorage.getItem("cskh.msg"); if (m) { $("#lgMsg").textContent = m; sessionStorage.removeItem("cskh.msg"); }
  $("#lgUser").focus();
}
$("#loginForm").onsubmit = async ev => {
  ev.preventDefault();
  const btn = $("#lgBtn"), msg = $("#lgMsg"); msg.textContent = "";
  const u = $("#lgUser").value, p = $("#lgPass").value; if (!u.trim() || !p) { msg.textContent = "Nhập tên đăng nhập và mật khẩu"; return; }
  btn.disabled = true; btn.textContent = "Đang kiểm tra…";
  try { await login(u, p); $("#lgPass").value = ""; afterLogin(); }
  catch (e) { msg.textContent = e.message; btn.disabled = false; btn.textContent = "Đăng nhập"; $("#lgPass").select(); }
};
function afterLogin() {
  $("#loginBox").hidden = true;
  document.body.dataset.role = SESSION.role;
  $("#userChip").innerHTML = `<span class="who"><b>${esc(SESSION.username)}</b><span class="role role-${SESSION.role}">${ROLE_LABEL[SESSION.role] || SESSION.role}</span></span><button class="btn" id="btnLogout">Đăng xuất</button>`;
  $("#btnLogout").onclick = () => logout();
  loadDataFile();
}

// ================= Trạng thái ứng dụng =================
let S = null;          // {fileName, sheetName, headerRow, importedAt, header[], colOf{}, records[]}
let origFile = null;   // ArrayBuffer của file Excel gốc (để giữ các sheet khác khi xuất)
let all = [];          // records chưa bị xoá
let issueMap = new Map(); // id -> [mã vấn đề]
let saveT;
function persist() { clearTimeout(saveT); saveT = setTimeout(() => saveEnc("state", S).catch(e => console.error(e)), 300); if (S) showData(true); }
function refreshAll() { all = S ? S.records.filter(r => !r._deleted) : []; computeIssues(); refreshOptions(); renderAll(); }
const issuesOf = r => issueMap.get(r.id) || [];
const hasIssue = (r, code) => issuesOf(r).includes(code);

// ================= Đọc file Excel =================
const ST_CANON = s => { const t = String(s ?? "").trim(); return TT.find(x => x.toLowerCase() === t.toLowerCase()) || t; };
function cellStr(v, key) {
  if (v == null) return "";
  if (v instanceof Date) return isoOf(v);
  if (typeof v === "number") {
    if (key === "sdt") { const s = String(Math.round(v)); return s.length === 9 ? "0" + s : s; }
    return String(Math.abs(v - Math.round(v)) < 1e-6 ? Math.round(v) : v);
  }
  return String(v).trim();
}
function parseDateStr(s) { // chuỗi ngày dạng m/d/yy -> ISO
  const m = String(s).match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/); if (!m) return "";
  let y = +m[3]; if (y < 100) y += 2000; return `${y}-${pad2(m[1])}-${pad2(m[2])}`;
}
// Đọc sheet thành mảng dòng; ô ngày được đổi từ mã ngày Excel sang Date đúng nửa đêm giờ địa phương.
// (Không dùng cellDates của SheetJS: ở múi giờ Việt Nam nó lùi mọi ngày 30 giây → thành ngày hôm trước.)
function serialToDate(v) { const p = XLSX.SSF.parse_date_code(v); return p ? new Date(p.y, p.m - 1, p.d, p.H, p.M, Math.round(p.S)) : null; }
function dateToSerial(d) { return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - Date.UTC(1899, 11, 30)) / 864e5; }
function sheetRows(ws) {
  const rg = XLSX.utils.decode_range(ws["!ref"]), rows = [];
  for (let R = rg.s.r; R <= rg.e.r; R++) {
    const row = [];
    for (let C = rg.s.c; C <= rg.e.c; C++) {
      const c = ws[XLSX.utils.encode_cell({r: R, c: C})];
      let v = c ? c.v : null;
      if (c && c.t === "n" && c.z && XLSX.SSF.is_date(c.z)) v = serialToDate(c.v) || v;
      else if (c && c.t === "d") v = c.v;
      else if (c && c.t === "e") v = null;
      row[C - rg.s.c] = v === undefined ? null : v;
    }
    rows.push(row);
  }
  return rows;
}
function findListSheet(wb) {
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name]; if (!ws["!ref"]) continue;
    const start = XLSX.utils.decode_range(ws["!ref"]).s.r;
    const rows = sheetRows(ws);
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const h = (rows[i] || []).map(norm);
      if (h.includes("điện thoại") && h.includes("trạng thái")) return {name, rows, headerRow: i, start};
    }
  }
  return null;
}
function mapHeader(header) {
  const colOf = {}, used = new Set();
  header.forEach((h, i) => {
    const n = norm(h); if (!n) return;
    const hit = CORE.find(([k, label]) => !used.has(k) && norm(label) === n);
    if (hit) { colOf[hit[0]] = i; used.add(hit[0]); }
  });
  return colOf;
}
function recordFromRow(row, colOf, i, excelRow) {
  const r = {id: "R" + String(i + 1).padStart(5, "0"), _row: excelRow};
  for (const [k] of CORE) {
    if (colOf[k] == null) continue;
    let v = cellStr(row[colOf[k]], k);
    if (k === "ngay" && v && !/^\d{4}-\d{2}-\d{2}$/.test(v)) v = parseDateStr(v) || v;
    if (k === "trangThai") v = ST_CANON(v);
    if (k === "nguon" && v.toLowerCase() === "khác") v = "Khác";
    if (v) r[k] = v;
  }
  if (!r.nam && /^\d{4}-/.test(r.ngay || "")) { r.nam = r.ngay.slice(0, 4); r._namDerived = true; }
  if (!r.thang && /^\d{4}-/.test(r.ngay || "")) r.thang = String(+r.ngay.slice(5, 7));
  r._orig = Object.fromEntries(CORE.map(([k]) => [k, r[k] ?? ""]));
  let end = row.length; while (end > 0 && (row[end - 1] == null || row[end - 1] === "")) end--;
  r._raw = row.slice(0, end);
  return r;
}
// Đọc file Excel (ArrayBuffer) thành trạng thái ứng dụng
function parseWorkbook(buf, fileName, sig, modified) {
  const wb = XLSX.read(buf, {type: "array", cellNF: true});
  const found = findListSheet(wb);
  if (!found) throw new Error("Không tìm thấy sheet có cột “Điện thoại” và “Trạng thái” trong " + fileName);
  const header = (found.rows[found.headerRow] || []).map(h => h == null ? "" : String(h));
  const colOf = mapHeader(header);
  const records = [];
  found.rows.forEach((row, idx) => {
    if (idx <= found.headerRow || !row || !row.some(v => v != null && String(v).trim() !== "")) return;
    records.push(recordFromRow(row, colOf, records.length, idx + 1 + found.start));
  });
  const st = {v: STATE_VERSION, fileName, fileSig: sig, fileModified: modified, sheetName: found.name, headerRow: found.headerRow, importedAt: new Date().toISOString(), header, colOf, records};
  buildWeekMap(st);
  return st;
}
async function sha1(buf) {
  try { return [...new Uint8Array(await crypto.subtle.digest("SHA-1", buf))].map(b => b.toString(16).padStart(2, "0")).join(""); }
  catch (e) { return "len" + buf.byteLength; } // trình duyệt không hỗ trợ crypto.subtle (http không phải localhost)
}
const editCount = st => st ? st.records.filter(r => r._new || r._deleted || r._wkEdited || isEdited(r)).length : 0;

// Nạp file dữ liệu cố định trong thư mục data/ (xem DATA_FILE ở đầu file)
async function loadDataFile() {
  setStatus("Đang tải dữ liệu từ " + DATA_FILE + "…");
  let res;
  try { res = await fetch(DATA_FILE + "?t=" + Date.now(), {cache: "no-store"}); }
  catch (e) {
    setStatus(location.protocol === "file:"
      ? `Trình duyệt không cho đọc file khi mở trực tiếp index.html. Hãy chạy <code>node tools/server.js</code> trong thư mục web rồi mở <b>http://localhost:8080</b>, hoặc mở bản đã deploy (Vercel / GitHub Pages).`
      : `Không tải được <code>${esc(DATA_FILE)}</code>: ${esc(e.message)}`, true);
    return;
  }
  if (!res.ok) { setStatus(`Không tìm thấy file dữ liệu <code>${esc(DATA_FILE)}</code> (lỗi ${res.status}). Hãy chép file Excel vào <code>data/Report_CSKH.xlsx</code> rồi chạy <code>node tools/encrypt.js</code>.`, true); return; }
  const enc = await res.arrayBuffer();
  let buf; try { buf = await decryptDataFile(enc); }
  catch (e) { setStatus(esc(e.message && e.message.includes("định dạng") ? e.message : "Không giải mã được dữ liệu — file đã được mã hoá lại bằng khoá khác. Hãy đăng xuất rồi đăng nhập lại."), true); return; }
  const sig = await sha1(enc), modified = res.headers.get("Last-Modified") || "";
  let saved = null; try { saved = await loadEnc("state"); } catch (e) { console.warn(e); }
  origFile = buf;
  try {
    if (saved && saved.v === STATE_VERSION && saved.fileSig === sig && saved.dataFile === DATA_FILE) { S = saved; S.fileModified = modified || S.fileModified; }
    else {
      S = parseWorkbook(buf, "Report_CSKH.xlsx", sig, modified); S.dataFile = DATA_FILE;
      const n = editCount(saved);
      if (n) { await saveEnc("prevState", saved); const old = await idb.get("file"); if (old) await idb.set("prevFile", old); }
      await saveEnc("state", S); await idb.set("file", enc);
    }
  } catch (e) { console.error(e); setStatus(esc(e.message || "Không đọc được file Excel."), true); return; }
  try { const ps = await loadEnc("prevState"); prevNotice = ps ? editCount(ps) : 0; } catch (e) {}
  yearsInit = false; refreshAll(); showData(true);
}
let prevNotice = 0;
function setStatus(html, isErr) { const el = $("#statusBox"); el.hidden = false; el.classList.toggle("err", !!isErr); el.innerHTML = `<h2>${isErr ? "Chưa tải được dữ liệu" : "Đang tải…"}</h2><p>${html}</p>`; }

// ================= Nhật ký theo tuần =================
/* Trong Excel, nhật ký tuần nằm ở các cột "Tuần N", xếp theo từng năm từ mới đến cũ:
   [Tuần 27 … Tuần 2] của 2026, [Tuần 53 … Tuần 1] của 2025, …  Tên cột lặp lại giữa các năm,
   nên web xác định năm của mỗi cột theo thứ tự: số tuần tăng lên = bắt đầu nhóm của năm trước. */
const WEEK_RE = /^tuần\s*(\d+)$/i;
function buildWeekMap(st) {
  const core = new Set(Object.values(st.colOf)), years = st.records.map(r => +r.nam).filter(y => y >= 2000 && y < 2100);
  const y0 = years.length ? Math.max(...years) : new Date().getFullYear();
  const map = {}; let prev = Infinity, blk = 0;
  st.header.forEach((h, i) => { const m = String(h || "").trim().match(WEEK_RE); if (!m || core.has(i)) return; const w = +m[1]; if (w > prev) blk++; prev = w; map[i] = {y: y0 - blk, w}; });
  st.weekCols = map;
}
// Tìm (hoặc tạo) cột "Tuần w" của năm y; cột mới được chèn đúng vị trí theo thứ tự năm/tuần
function weekCol(y, w, create) {
  const hit = Object.entries(S.weekCols || {}).find(([, v]) => v.y === y && v.w === w); if (hit) return +hit[0];
  if (!create) return -1;
  const cols = Object.entries(S.weekCols || {}).map(([c, v]) => [+c, v]).sort((a, b) => a[0] - b[0]);
  let pos = null;
  for (const [c, v] of cols) if (v.y < y || (v.y === y && v.w < w)) { pos = c; break; }
  if (pos == null) pos = cols.length ? cols[cols.length - 1][0] + 1 : (S.colOf.nhatKy != null ? S.colOf.nhatKy + 1 : S.header.length);
  S.header.splice(pos, 0, "Tuần " + w);
  for (const r of S.records) if (r._raw && r._raw.length > pos) r._raw.splice(pos, 0, null);
  for (const k in S.colOf) if (S.colOf[k] >= pos) S.colOf[k]++;
  const nm = {}; for (const [c, v] of Object.entries(S.weekCols || {})) nm[+c >= pos ? +c + 1 : +c] = v; nm[pos] = {y, w}; S.weekCols = nm;
  return pos;
}
function weeklyEntries(r) { // [{col, y, w, text}] mới nhất trước
  if (!r._raw || !S || !S.weekCols) return [];
  return Object.entries(S.weekCols).map(([c, v]) => ({col: +c, ...v, raw: r._raw[+c]}))
    .filter(e => e.raw != null && String(e.raw).trim() !== "")
    .map(e => ({...e, text: e.raw instanceof Date ? dmy(isoOf(e.raw)) : String(e.raw)}))
    .sort((a, b) => b.y - a.y || b.w - a.w);
}
const latestNote = r => { const e = weeklyEntries(r)[0]; return e ? `T${e.w}/${e.y}: ${e.text}` : String(r.nhatKy || r.thongTin || ""); };

// ================= Phát hiện vấn đề dữ liệu =================
let placeholderDates = new Set(), accVariants = {};
function computeIssues() {
  issueMap = new Map();
  const phoneN = {}, dateN = {}, accSpell = {};
  for (const r of all) {
    for (const p of new Set(validPhones(r.sdt))) phoneN[p] = (phoneN[p] || 0) + 1;
    if (r.ngay) dateN[r.ngay] = (dateN[r.ngay] || 0) + 1;
    if (r.acc) (accSpell[r.acc.toLowerCase()] ||= new Set()).add(r.acc);
  }
  const med = median(Object.values(dateN));
  placeholderDates = new Set(Object.entries(dateN).filter(([, n]) => n >= 30 && n >= 8 * Math.max(1, med)).map(([d]) => d));
  accVariants = Object.fromEntries(Object.entries(accSpell).filter(([, s]) => s.size > 1).map(([k, s]) => [k, [...s]]));
  const tIdx = S && S.colOf ? S.colOf.trangThai : null;
  for (const r of all) {
    const out = [];
    const y = r.nam || "", dOk = /^\d{4}-\d{2}-\d{2}$/.test(r.ngay || "");
    if (!/^20\d\d$/.test(y)) out.push("NO_YEAR");
    else if (dOk && r.ngay.slice(0, 4) !== y) out.push("YEAR_MISMATCH");
    if (r._namDerived) out.push("YEAR_DERIVED");
    if (/^20\d\d$/.test(y)) { if (!r.thang) out.push("NO_MONTH"); else if (dOk && +r.ngay.slice(5, 7) !== +r.thang && r.ngay.slice(0, 4) === y) out.push("MONTH_MISMATCH"); }
    if (!r.ngay) out.push("NO_DATE"); else if (dOk && r.ngay > TODAY) out.push("FUTURE_DATE");
    if (placeholderDates.has(r.ngay)) out.push("PLACEHOLDER_DATE");
    const ps = phonesOf(r.sdt);
    if (!ps.some(Boolean)) out.push(r.sdt ? "BAD_PHONE" : "NO_PHONE"); else if (ps.some(p => !phoneOk(p))) out.push("BAD_PHONE");
    if (validPhones(r.sdt).some(p => phoneN[p] > 1)) out.push("DUP_PHONE");
    if (!r.ten) out.push("NO_NAME");
    if (!r.loai) out.push("NO_LOAI");
    if (!r.trangThai) out.push("NO_STATUS");
    if (tIdx != null && r._raw && r._orig && r.trangThai === r._orig.trangThai) {
      const raw = r._raw[tIdx]; if (raw != null && String(raw) !== r.trangThai && TT.includes(r.trangThai)) out.push("STATUS_SPELL");
    }
    if (r.trangThai === "Kết thúc" && !r.lyDo) out.push("END_NO_REASON");
    if (r.lyDo && r.trangThai && r.trangThai !== "Kết thúc") out.push("REASON_NOT_END");
    if (r.trangThai === "Đã ký HĐ" && !parseMoney(r.giaTri)) out.push("SIGNED_NO_VALUE");
    if (r.giaTri && (/[a-zà-ỹ]/i.test(r.giaTri) || (parseMoney(r.giaTri) > 0 && parseMoney(r.giaTri) < 1e6))) out.push("VALUE_ODD");
    if (!r.tinh) out.push("NO_TINH");
    else if (!kvOf(r.tinh)) out.push("UNKNOWN_TINH");
    else if (r.kv && kvOf(r.tinh) !== r.kv) out.push("KV_MISMATCH");
    if ((r.nguonQC && !QC[r.nguonQC]) || (r.lyDo && !LYDO[r.lyDo]) || (r.phanKhuc && !PK[r.phanKhuc]) || (r.nguon && !NGUON.includes(r.nguon)) || (r.loai && !LOAI.some(l => l[0] === r.loai))) out.push("BAD_CODE");
    if (r.acc && accVariants[r.acc.toLowerCase()]) out.push("ACC_CASE");
    if (ACTIVE.has(r.trangThai) && dOk && daysBetween(r.ngay, TODAY) > STALE_DAYS) out.push("STALE");
    if (r.nguon === "FB" && !r.nguonQC && +y >= 2025) out.push("FB_NO_QC");
    if (r.loai === "ĐL" && !r.phanKhuc) out.push("DL_NO_SEG");
    issueMap.set(r.id, out);
  }
  const nCrit = all.filter(r => issuesOf(r).some(c => ISSUES[c].sev === "crit")).length;
  const b = $("#checkBadge"); b.hidden = !nCrit; b.textContent = nCrit > 999 ? "999+" : nCrit;
}
// Mô tả cụ thể lỗi của một dòng (giá trị đang có trong Excel)
function issueDetail(r, c) {
  const o = r._orig || {}, raw = k => (S && S.colOf[k] != null && r._raw) ? r._raw[S.colOf[k]] : undefined;
  const show = v => v == null || v === "" ? "(trống)" : v instanceof Date ? dmy(isoOf(v)) : String(v);
  switch (c) {
    case "NO_YEAR": return `Năm trong Excel: ${show(raw("nam") ?? r.nam)}`;
    case "YEAR_MISMATCH": return `Năm = ${r.nam} nhưng Ngày nhận = ${dmy(r.ngay)} → bấm Lưu để tự sửa Năm theo Ngày nhận`;
    case "MONTH_MISMATCH": return `Tháng = ${r.thang} nhưng Ngày nhận = ${dmy(r.ngay)} → bấm Lưu để tự sửa Tháng theo Ngày nhận`;
    case "NO_MONTH": return `Năm = ${r.nam}, Tháng trống`;
    case "YEAR_DERIVED": return `Cột Năm trong Excel trống, web lấy ${r.nam} từ Ngày nhận ${dmy(r.ngay)}`;
    case "FUTURE_DATE": return `Ngày nhận = ${dmy(r.ngay)}`;
    case "PLACEHOLDER_DATE": return `Ngày nhận = ${dmy(r.ngay)} (trùng với rất nhiều KH khác)`;
    case "BAD_PHONE": return `SĐT = “${r.sdt}” — ${phonesOf(r.sdt).filter(p => !phoneOk(p)).map(p => p ? p + " (" + p.length + " chữ số)" : "không có chữ số").join(", ")}. Nhiều số thì cách nhau bằng dấu ;`;
    case "DUP_PHONE": { const mine = new Set(validPhones(r.sdt)); const d = all.filter(x => x.id !== r.id && validPhones(x.sdt).some(p => mine.has(p))); return `Trùng với: ${d.slice(0, 3).map(x => (x.ten || "KH") + " – " + (x.acc || "?") + " – " + (dmy(x.ngay) || x.nam || "")).join("; ")}`; }
    case "KV_MISMATCH": return `Tỉnh = ${r.tinh}, Khu vực = ${r.kv} (đúng phải là ${kvOf(r.tinh)})`;
    case "UNKNOWN_TINH": return `Tỉnh = “${r.tinh}”`;
    case "BAD_CODE": return [r.nguon && !NGUON.includes(r.nguon) && `Nguồn = ${r.nguon}`, r.nguonQC && !QC[r.nguonQC] && `Nguồn QC = ${r.nguonQC}`, r.lyDo && !LYDO[r.lyDo] && `Lý do = ${r.lyDo}`, r.phanKhuc && !PK[r.phanKhuc] && `Phân khúc = ${r.phanKhuc}`, r.loai && !LOAI.some(l => l[0] === r.loai) && `Loại KH = ${r.loai}`].filter(Boolean).join(", ");
    case "STATUS_SPELL": return `Trong Excel ghi “${show(raw("trangThai"))}”`;
    case "REASON_NOT_END": return `Trạng thái = ${r.trangThai}, Lý do kết thúc = ${LYDO[r.lyDo] || r.lyDo}`;
    case "VALUE_ODD": return `Giá trị ghi: “${r.giaTri}”`;
    case "ACC_CASE": return `ACC = ${r.acc}; các cách viết khác: ${(accVariants[r.acc.toLowerCase()] || []).filter(a => a !== r.acc).join(", ")}`;
    case "STALE": return `Nhận ngày ${dmy(r.ngay)} (${daysBetween(r.ngay, TODAY)} ngày trước), vẫn ${r.trangThai}`;
    default: return "";
  }
}
function issueCount(rs, code) { return rs.filter(r => hasIssue(r, code)).length; }

// ================= Bảng sắp xếp được =================
const sortState = {};
function cmp(a, b) {
  const ea = a == null || a === "", eb = b == null || b === "";
  if (ea || eb) return ea && eb ? 0 : ea ? 1 : -1;
  if (typeof a === "number" && typeof b === "number") return a - b;
  return String(a).localeCompare(String(b), "vi", {numeric: true});
}
function sortRows(host, cols, rows, def) {
  const st = sortState[host] ||= {k: def?.k ?? null, dir: def?.dir ?? "desc"};
  if (!st.k) return rows.slice();
  const c = cols.find(c => c.k === st.k); if (!c) return rows.slice();
  const v = c.val || (r => r[c.k]);
  return rows.slice().sort((a, b) => {
    const x = v(a), y = v(b), ea = x == null || x === "", eb = y == null || y === "";
    if (ea || eb) return ea && eb ? 0 : ea ? 1 : -1; // ô trống luôn ở cuối
    return cmp(x, y) * (st.dir === "asc" ? 1 : -1);
  });
}
/* cols: [{k, label, num, html(row), val(row), cls(row), title}] ; opts: {def, onRow, foot[], presorted, rerender, empty} */
function table(host, cols, rows, opts = {}) {
  const el = $(host);
  const st = sortState[host] ||= {k: opts.def?.k ?? null, dir: opts.def?.dir ?? "desc"};
  const data = opts.presorted ? rows : sortRows(host, cols, rows, opts.def);
  if (!data.length && opts.empty !== false) { el.innerHTML = `<p class="muted" style="padding:${opts.pad || 0}">${opts.empty || "Chưa có dữ liệu"}</p>`; return; }
  const th = c => `<th class="sortable ${c.num ? "num" : ""}" data-k="${c.k}" ${c.title ? `title="${esc(c.title)}"` : ""} aria-sort="${st.k === c.k ? (st.dir === "asc" ? "ascending" : "descending") : "none"}" tabindex="0">${c.label}</th>`;
  const td = (c, r) => { const cl = [c.num ? "num" : "", c.cls ? c.cls(r) || "" : ""].join(" ").trim(); return `<td${cl ? ` class="${cl}"` : ""}>${c.html ? c.html(r) : esc(r[c.k] ?? "")}</td>`; };
  el.innerHTML = `<table><thead><tr>${cols.map(th).join("")}</tr></thead><tbody>${data.map((r, i) => `<tr${opts.onRow ? ` class="click" data-i="${i}"` : ""}>${cols.map(c => td(c, r)).join("")}</tr>`).join("")}</tbody>${opts.foot ? `<tfoot><tr>${opts.foot.map((f, i) => `<td${cols[i] && cols[i].num ? ' class="num"' : ""}>${f ?? ""}</td>`).join("")}</tr></tfoot>` : ""}</table>`;
  el.querySelectorAll("th.sortable").forEach(h => {
    const go = () => { const k = h.dataset.k, c = cols.find(c => c.k === k); if (st.k === k) st.dir = st.dir === "asc" ? "desc" : "asc"; else { st.k = k; st.dir = c.num ? "desc" : "asc"; } opts.rerender ? opts.rerender() : table(host, cols, rows, opts); };
    h.onclick = go; h.onkeydown = e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go(); } };
  });
  if (opts.onRow) el.querySelectorAll("tbody tr[data-i]").forEach(tr => tr.onclick = () => opts.onRow(data[+tr.dataset.i]));
}

// ================= Nhận định =================
const I = (sev, html, go) => ({sev, html, go});
function renderIns(host, items, title = "Nhận định") {
  const el = $(host); items = items.filter(Boolean);
  if (!items.length) { el.innerHTML = ""; return; }
  items.sort((a, b) => (SEV_RANK[a.sev] ?? 3) - (SEV_RANK[b.sev] ?? 3));
  el.innerHTML = `<h3>${title}</h3>` + items.map((it, i) => `<div class="ins ${it.go ? "go" : ""}" data-i="${i}" ${it.go ? 'role="button" tabindex="0" title="Xem danh sách KH liên quan"' : ""}><span class="sev sev-${it.sev}" aria-label="${SEV_LABEL[it.sev]}">${SEV_ICON[it.sev]}</span><span>${it.html}</span></div>`).join("");
  el.querySelectorAll(".ins.go").forEach(d => { const go = () => applyGo(items[+d.dataset.i].go); d.onclick = go; d.onkeydown = e => { if (e.key === "Enter") go(); }; });
}
function applyGo(go) {
  resetListFilters();
  const map = {year: "#lYear", loai: "#lLoai", st: "#lSt", acc: "#lAcc", nguon: "#lNguon", issue: "#lIssue"};
  for (const [k, sel] of Object.entries(map)) if (go[k] != null) { const el = $(sel); if (![...el.options].some(o => o.value === go[k])) el.insertAdjacentHTML("beforeend", `<option value="${esc(go[k])}">${esc(go[k])}</option>`); el.value = go[k]; }
  if (go.q) $("#q").value = go.q;
  page = 0; renderList(); show("list"); window.scrollTo({top: 0});
}
function issueIns(rs, codes, ctx) { // dòng nhận định chung cho các vấn đề dữ liệu ảnh hưởng tới bảng
  return codes.map(c => { const n = issueCount(rs, c); if (!n) return null; const d = ISSUES[c];
    return I(d.sev, `<b>${fmt(n)}</b> dòng: ${esc(d.label.toLowerCase())}`, {issue: c, year: ctx.Y, loai: ctx.L, acc: ctx.A}); });
}

// ================= Điều hướng =================
let view = "overview";
const VIEWS = ["overview", "list", "check", "add"];
document.querySelectorAll("nav.tabs button").forEach(b => b.onclick = () => show(b.dataset.view));
function show(v) {
  if (v === "add" && !can.edit()) v = "overview";
  view = v;
  document.querySelectorAll("nav.tabs button").forEach(b => b.setAttribute("aria-selected", b.dataset.view === v));
  if (!S) return;
  VIEWS.forEach(k => $("#view-" + k).hidden = k !== v);
  if (v === "add") mountAddForm();
  if (v === "check") renderCheck();
}
function showData(has) {
  $("#statusBox").hidden = has; $("nav.tabs").hidden = !has;
  $("#btnExport").disabled = !has; $("#btnClear").disabled = !has || !editCount(S);
  $("#btnExport").hidden = !can.export(); $("#clearHost").hidden = !can.reset();
  document.querySelector('nav.tabs [data-view="add"]').hidden = !can.edit();
  const pb = $("#prevBox"); pb.hidden = !(has && prevNotice);
  if (has && prevNotice) pb.innerHTML = `<span class="sev sev-warn">!</span><span>File dữ liệu đã được thay bằng bản mới. Có <b>${fmt(prevNotice)}</b> chỉnh sửa trên web của file cũ chưa xuất Excel — đã được giữ lại để bạn tải về.</span>${can.export() ? '<button class="btn" id="prevExport">Tải bản chỉnh sửa cũ</button>' : ""}<button class="btn" id="prevDrop">Bỏ qua</button>`;
  if (has && prevNotice) { if ($("#prevExport")) $("#prevExport").onclick = exportPrev; $("#prevDrop").onclick = async () => { prevNotice = 0; await idb.del("prevState"); await idb.del("prevFile"); showData(true); }; }
  if (!has) VIEWS.forEach(k => $("#view-" + k).hidden = true); else show(view);
  $("#srcInfo").textContent = has ? `Dữ liệu: Report_CSKH (đã mã hoá) · sheet “${S.sheetName}”${S.fileModified ? " · file cập nhật " + new Date(S.fileModified).toLocaleString("vi-VN") : ""}${editCount(S) ? " · " + fmt(editCount(S)) + " chỉnh sửa trên web chưa xuất" : ""}` : "Chưa có dữ liệu";
}

// ================= Bộ lọc =================
function uniq(f) { return [...new Set(all.map(f).filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), "vi")); }
function fillSel(sel, vals, labels) {
  const el = $(sel), cur = el.value, first = el.options[0] ? el.options[0].outerHTML : "";
  el.innerHTML = first + vals.map(v => `<option value="${esc(v)}">${esc(labels ? labels(v) : v)}</option>`).join("");
  if ([...el.options].some(o => o.value === cur)) el.value = cur;
}
let yearsInit = false;
const yearsList = () => uniq(r => /^20\d\d$/.test(r.nam) ? r.nam : null).sort().reverse();
function refreshOptions() {
  const years = yearsList();
  const yo = `<option value="">Tất cả năm</option>` + years.map(y => `<option>${y}</option>`).join("");
  for (const s of ["#fYear", "#lYear", "#cYear"]) { const el = $(s), cur = el.value; el.innerHTML = yo; if ([...el.options].some(o => o.value === cur)) el.value = cur; }
  if (!yearsInit && years.length) { $("#fYear").value = years[0]; $("#cYear").value = years[0]; yearsInit = true; }
  const accs = uniq(r => r.acc); fillSel("#fAcc", accs); fillSel("#lAcc", accs);
  fillSel("#fLoai", LOAI.map(l => l[0])); fillSel("#lLoai", LOAI.map(l => l[0]));
  fillSel("#lSt", TT); fillSel("#lNguon", NGUON, v => NGUON_LABEL[v] || v);
  const ic = {}; all.forEach(r => issuesOf(r).forEach(c => ic[c] = (ic[c] || 0) + 1));
  fillSel("#lIssue", Object.keys(ISSUES).filter(c => ic[c]), c => `${SEV_LABEL[ISSUES[c].sev]}: ${ISSUES[c].label} (${fmt(ic[c])})`);
  $("#dlAcc").innerHTML = accs.map(a => `<option value="${esc(a)}">`).join("");
  $("#dlTinh").innerHTML = Object.keys(KV_MAP).map(t => `<option value="${esc(t)}">`).join("");
  $("#dlDL").innerHTML = uniq(r => r.dlChuyen).map(a => `<option value="${esc(a)}">`).join("");
}

// ================= Tổng quan =================
["#fYear", "#fLoai", "#fAcc"].forEach(s => $(s).onchange = renderOverview);
function hbars(host, entries, total, color) {
  const max = Math.max(1, ...entries.map(e => e[1]));
  $(host).innerHTML = entries.length ? entries.map(([k, v, extra]) =>
    `<div class="hb" data-tip="${esc(k)}\n${fmt(v)} KH · ${pct(v, total)}${extra ? "\n" + esc(extra) : ""}"><span class="n">${esc(k)}</span><div class="track"><div class="fill" style="width:${v / max * 100}%;${color ? `background:${color}` : ""}"></div></div><span class="val">${fmt(v)}<small>${pct(v, total)}</small></span></div>`
  ).join("") : `<p class="muted">Chưa có dữ liệu</p>`;
}
function niceStep(max) { const raw = max / 4; const p = Math.pow(10, Math.floor(Math.log10(raw || 1))); const f = raw / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * p; }
const signedOf = g => g.filter(r => r.trangThai === "Đã ký HĐ").length;
const endedOf = g => g.filter(r => r.trangThai === "Kết thúc").length;
const activeOf = g => g.filter(r => ACTIVE.has(r.trangThai)).length;

function buildCtx() {
  const Y = $("#fYear").value, L = $("#fLoai").value, A = $("#fAcc").value;
  const match = r => (!L || r.loai === L) && (!A || r.acc === A);
  const rs = all.filter(r => (!Y || r.nam === Y) && match(r));
  let prev = [], prevSP = [], cutoff = "", PY = "";
  if (Y) {
    PY = String(+Y - 1);
    prev = all.filter(r => r.nam === PY && match(r));
    const dates = rs.map(r => r.ngay).filter(d => /^\d{4}-/.test(d || "") && d.slice(0, 4) === Y).sort();
    cutoff = dates.length ? dates[dates.length - 1] : "";
    prevSP = cutoff ? prev.filter(r => /^\d{4}-/.test(r.ngay || "") && r.ngay.slice(5) <= cutoff.slice(5)) : prev;
  }
  return {Y, L, A, rs, prev, prevSP, cutoff, PY, hasPrev: prev.length > 0};
}

function renderOverview() {
  const ctx = buildCtx(), {Y, rs, prevSP, cutoff, PY} = ctx;
  renderBanner(ctx);
  // ---- KPI
  const tot = rs.length, signed = rs.filter(r => r.trangThai === "Đã ký HĐ");
  const val = sum(signed.map(r => parseMoney(r.giaTri)));
  const act = activeOf(rs), end = endedOf(rs);
  const sp = ctx.hasPrev && cutoff;
  const vs = sp ? `so với cùng kỳ ${PY} (đến ${dm(cutoff)})` : "";
  $("#kpis").innerHTML = [
    ["Tổng khách hàng", fmt(tot), sp ? `${delta(tot, prevSP.length)} ${vs}` : (Y ? `Năm ${Y}` : "Tất cả các năm"), ""],
    ["Đang chăm sóc", fmt(act), `${pct(act, tot)} tổng KH` + (issueCount(rs, "STALE") ? ` · ${fmt(issueCount(rs, "STALE"))} quá 6 tháng` : ""), ""],
    ["Kết thúc", fmt(end), `${pct(end, tot)} ` + (sp ? delta(pctN(end, tot), pctN(endedOf(prevSP), prevSP.length), {unit: "pp", inverse: true}) : "tổng KH"), ""],
    ["Đã ký hợp đồng", fmt(signed.length), `Tỷ lệ ký ${pct(signed.length, tot)} ` + (sp ? delta(pctN(signed.length, tot), pctN(signedOf(prevSP), prevSP.length), {unit: "pp"}) : ""), "hi"],
    ["Giá trị HĐ ghi nhận", val ? vnd(val) : "—", `${signed.filter(r => parseMoney(r.giaTri)).length}/${signed.length} HĐ có giá trị`, ""],
  ].map(([l, v, s, c]) => `<div class="kpi ${c}"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`).join("");

  renderMonthly(ctx); renderStatus(ctx); renderSource(ctx); renderLoai(ctx); renderReason(ctx);
  renderSeg(ctx); renderQC(ctx); renderKV(ctx); renderAcc(ctx); renderContracts(ctx);
}

function renderBanner(ctx) {
  const {rs, Y} = ctx;
  const bad = rs.filter(r => issuesOf(r).some(c => ISSUES[c].sev === "crit"));
  const warn = rs.filter(r => issuesOf(r).some(c => ISSUES[c].sev === "warn"));
  const noYear = Y ? all.filter(r => hasIssue(r, "NO_YEAR")).length : 0;
  const clean = rs.filter(r => !issuesOf(r).some(c => ISSUES[c].sev !== "info")).length;
  const el = $("#healthBanner");
  if (!rs.length) { el.innerHTML = ""; return; }
  el.innerHTML = `<div class="banner ${bad.length || noYear ? "" : "ok"}"><span class="sev ${bad.length || noYear ? "sev-warn" : "sev-good"}">${bad.length || noYear ? "!" : "✓"}</span>
    <span>Chất lượng dữ liệu${Y ? " năm " + Y : ""}: <b>${pct(clean, rs.length)}</b> dòng không có lỗi/cảnh báo · <b>${fmt(bad.length)}</b> dòng có lỗi làm sai thống kê · <b>${fmt(warn.length)}</b> dòng cần xem lại${noYear ? ` · <b>${fmt(noYear)}</b> dòng không xác định được năm (không nằm trong số liệu này)` : ""}</span>
    <button class="btn" id="bannerGo">Xem chi tiết</button></div>`;
  $("#bannerGo").onclick = () => { $("#cYear").value = Y; show("check"); };
}

// ---- Biểu đồ theo tháng / năm
function renderMonthly(ctx) {
  const {Y, rs, prev, prevSP, cutoff, PY} = ctx;
  const byYear = !Y;
  $("#tMonthly").textContent = byYear ? "Khách hàng mới theo năm" : "Khách hàng mới theo tháng · " + Y;
  const keys = byYear ? yearsList().slice().reverse().filter(y => rs.some(r => r.nam === y)) : Array.from({length: 12}, (_, i) => String(i + 1));
  const kf = byYear ? r => r.nam : r => String(+r.thang || "");
  const types = LOAI.map(l => l[0]).filter(t => rs.some(r => r.loai === t));
  const buckets = keys.map(k => { const g = rs.filter(r => kf(r) === k); const o = {k, total: g.length, g}; types.forEach(t => o[t] = g.filter(r => r.loai === t).length); return o; });
  const prevBy = byYear ? {} : count(prev, r => String(+r.thang || ""));
  const max = Math.max(1, ...buckets.map(b => b.total)); const step = niceStep(max); const top = Math.ceil(max / step) * step;
  $("#lgMonthly").innerHTML = types.map(t => `<span><i style="background:${LOAI_COLOR[t]}"></i>${t}</span>`).join("");
  let grid = ""; for (let v = step; v <= top; v += step) grid += `<div class="gl" style="bottom:${v / top * 100}%"><span>${fmt(v)}</span></div>`;
  $("#chMonthly").innerHTML = buckets.length ? `<div class="colchart">${grid}${buckets.map(b => {
    const tip = (byYear ? "Năm " : "Tháng ") + b.k + "\nTổng: " + fmt(b.total) + types.map(t => b[t] ? `\n${t}: ${fmt(b[t])}` : "").join("") +
      (!byYear && ctx.hasPrev ? `\nCùng tháng ${PY}: ${fmt(prevBy[b.k] || 0)}` : "") + (b.total ? `\nĐã ký: ${signedOf(b.g)} · Kết thúc: ${endedOf(b.g)}` : "");
    return `<div class="col" data-tip="${esc(tip)}">${types.map(t => b[t] ? `<div class="seg" style="height:max(0px,calc(${b[t] / top * 100}% - 2px));background:${LOAI_COLOR[t]}"></div>` : "").join("")}${b.total ? `<span class="tot" style="bottom:calc(${b.total / top * 100}% + 2px)">${fmt(b.total)}</span>` : ""}</div>`;
  }).join("")}</div><div class="xl">${buckets.map(b => `<span>${byYear ? b.k : "T" + b.k}</span>`).join("")}</div>` : `<p class="muted">Chưa có dữ liệu</p>`;

  const ins = [];
  if (byYear) {
    for (let i = 1; i < buckets.length; i++) {
      const a = buckets[i - 1].total, b = buckets[i].total; if (!a) continue; const d = (b - a) / a * 100;
      if (Math.abs(d) >= 50) ins.push(I(d > 0 ? "info" : "warn", `Năm <b>${buckets[i].k}</b>: ${fmt(b)} KH, ${d > 0 ? "tăng" : "giảm"} <b>${fmt(Math.abs(d))}%</b> so với ${buckets[i - 1].k}`));
    }
    const cy = String(new Date().getFullYear()); if (buckets.some(b => b.k === cy)) ins.push(I("info", `Năm ${cy} chưa hết năm — so sánh chính xác hơn khi chọn một năm cụ thể (web tự so với cùng kỳ)`));
  } else {
    const counts = buckets.map(b => b.total);
    const withData = counts.map((n, i) => n ? i : -1).filter(i => i >= 0);
    if (withData.length) {
      const first = withData[0], last = withData[withData.length - 1];
      const inRange = counts.slice(first, last + 1), med = median(inRange.filter(n => n));
      for (let i = first; i <= last; i++) if (!counts[i]) ins.push(I("crit", `Tháng <b>T${i + 1}</b> = 0 KH trong khi các tháng xung quanh đều có dữ liệu → có thể thiếu dữ liệu hoặc nhập sai cột Tháng`));
      for (let i = first; i <= last; i++) {
        const n = counts[i]; if (!n || withData.length < 3) continue;
        if (n >= 2 * med && n - med >= 10) {
          const ph = buckets[i].g.filter(r => placeholderDates.has(r.ngay));
          ins.push(I(ph.length ? "warn" : "info", `Tháng <b>T${i + 1}</b> cao đột biến: ${fmt(n)} KH, gấp ${fmt(n / med)} lần trung vị (${fmt(med)}/tháng)` + (ph.length ? ` — trong đó <b>${fmt(ph.length)}</b> KH cùng ngày nhận ${dmy(ph[0].ngay)} (nghi là ngày mặc định khi nhập bù)` : ""), ph.length ? {issue: "PLACEHOLDER_DATE", year: Y} : null));
        } else if (n <= 0.4 * med && med - n >= 10) ins.push(I("warn", `Tháng <b>T${i + 1}</b> thấp bất thường: chỉ ${fmt(n)} KH (trung vị ${fmt(med)}/tháng)`));
      }
      const curY = String(new Date().getFullYear()), curM = new Date().getMonth() + 1;
      if (Y === curY && cutoff && last + 1 < curM) ins.push(I("warn", `Dữ liệu mới nhất: <b>${dmy(cutoff)}</b>. Tháng T${last + 2}${last + 2 < curM ? "–T" + curM : ""} chưa có KH nào → file có thể chưa được cập nhật`));
      else if (Y < curY && last < 11) ins.push(I("warn", `Năm ${Y} chỉ có dữ liệu đến tháng T${last + 1}`));
      if (withData.length >= 6) {
        const l3 = inRange.slice(-3), p3 = inRange.slice(-6, -3), a = sum(l3) / 3, b = sum(p3) / 3;
        if (b) { const d = (a - b) / b * 100; if (Math.abs(d) >= 15) ins.push(I(d > 0 ? "good" : "warn", `Xu hướng: 3 tháng gần nhất trung bình <b>${fmt(a)}</b> KH/tháng, ${d > 0 ? "tăng" : "giảm"} ${fmt(Math.abs(d))}% so với 3 tháng trước đó`)); }
      }
      ins.push(I("info", `Trung bình <b>${fmt(sum(inRange) / inRange.length)}</b> KH/tháng · cao nhất T${counts.indexOf(Math.max(...counts)) + 1} (${fmt(Math.max(...counts))}) · thấp nhất T${first + inRange.indexOf(Math.min(...inRange)) + 1} (${fmt(Math.min(...inRange))})`));
    }
    if (ctx.hasPrev && cutoff) {
      const a = rs.length, b = prevSP.length, d = b ? (a - b) / b * 100 : 0;
      ins.push(I(d >= 0 ? "good" : d < -20 ? "warn" : "info", `Cùng kỳ (01/01–${dm(cutoff)}): <b>${fmt(a)}</b> KH năm ${Y} so với ${fmt(b)} KH năm ${PY} (${d >= 0 ? "+" : ""}${fmt(d)}%)`));
      for (const [t] of LOAI) {
        const s1 = pctN(rs.filter(r => r.loai === t).length, rs.length), s0 = pctN(prevSP.filter(r => r.loai === t).length, prevSP.length);
        if (s1 != null && s0 != null && Math.abs(s1 - s0) >= 10) ins.push(I("info", `Cơ cấu thay đổi: ${t} chiếm <b>${pctS(s1)}</b> (cùng kỳ ${PY}: ${pctS(s0)})`));
      }
    }
  }
  ins.push(...issueIns(rs, ["NO_MONTH", "MONTH_MISMATCH", "FUTURE_DATE"], ctx));
  renderIns("#inMonthly", ins);
}

// ---- Trạng thái
function renderStatus(ctx) {
  const {rs, prevSP, PY, Y} = ctx, tot = rs.length, st = count(rs, r => r.trangThai);
  hbars("#chStatus", TT.map(t => [t, st[t] || 0]).filter(e => e[1]), tot);
  const ins = [];
  if (ctx.hasPrev && prevSP.length) {
    const e1 = pctN(endedOf(rs), tot), e0 = pctN(endedOf(prevSP), prevSP.length);
    if (e1 != null && Math.abs(e1 - e0) >= 5) ins.push(I(e1 > e0 ? "warn" : "good", `Tỷ lệ kết thúc <b>${pctS(e1)}</b>, ${e1 > e0 ? "cao hơn" : "thấp hơn"} cùng kỳ ${PY} (${pctS(e0)})`));
  }
  const stale = issueCount(rs, "STALE");
  if (stale) ins.push(I(stale / Math.max(1, activeOf(rs)) > .3 ? "warn" : "info", `<b>${fmt(stale)}</b> KH đang chăm sóc đã quá 6 tháng (${pct(stale, activeOf(rs))} số đang chăm sóc) → nên chốt hoặc chuyển Kết thúc`, {issue: "STALE", year: Y, loai: ctx.L, acc: ctx.A}));
  if (!st["Tiềm năng"] && !st["KH đồng ý ký HĐ"] && (st["Chăm sóc"] || 0) >= 30) ins.push(I("info", `Không có KH nào ở trạng thái <b>Tiềm năng</b> / <b>Đồng ý ký HĐ</b> → phễu chưa được phân loại, khó dự báo số HĐ sắp ký`));
  if (tot && (st["Tham khảo"] || 0) / tot > .15) ins.push(I("info", `Nhóm “Tham khảo” chiếm ${pct(st["Tham khảo"], tot)} — nhóm này khác với “Kết thúc – chỉ tham khảo”, nên thống nhất cách dùng`));
  ins.push(...issueIns(rs, ["NO_STATUS", "STATUS_SPELL"], ctx));
  renderIns("#inStatus", ins);
}

// ---- Nguồn
function renderSource(ctx) {
  const {rs, Y} = ctx, tot = rs.length, allRate = pctN(signedOf(rs), tot);
  const rows = [...new Set([...NGUON, ...rs.map(r => r.nguon || "(trống)")])].map(n => {
    const g = rs.filter(r => (r.nguon || "(trống)") === n); if (!g.length) return null;
    return {n, label: NGUON_LABEL[n] || n, t: g.length, share: pctN(g.length, tot), ky: signedOf(g), rate: pctN(signedOf(g), g.length), end: pctN(endedOf(g), g.length), act: activeOf(g)};
  }).filter(Boolean);
  const big = rows.filter(r => r.t >= 20), best = big.length ? Math.max(...big.map(r => r.rate)) : null;
  table("#tbSource", [
    {k: "label", label: "Nguồn"},
    {k: "t", label: "KH", num: true, html: r => fmt(r.t)},
    {k: "share", label: "Tỷ trọng", num: true, html: r => pctS(r.share)},
    {k: "act", label: "Đang CS", num: true, html: r => fmt(r.act)},
    {k: "ky", label: "Đã ký", num: true, html: r => fmt(r.ky)},
    {k: "rate", label: "Tỷ lệ ký", num: true, html: r => pctS(r.rate) + (r.t < 20 && r.ky ? '<span class="warnmark" title="Mẫu nhỏ, tỷ lệ chưa đáng tin">!</span>' : ""), cls: r => r.t >= 20 && r.ky && r.rate === best && big.length > 1 ? "good" : ""},
    {k: "end", label: "Kết thúc", num: true, html: r => pctS(r.end), cls: r => r.t >= 30 && r.end - pctN(endedOf(rs), tot) >= 15 ? "bad" : ""},
  ], rows, {def: {k: "t", dir: "desc"}, foot: ["Tổng", fmt(tot), "100%", fmt(activeOf(rs)), fmt(signedOf(rs)), pctS(allRate), pct(endedOf(rs), tot)],
    onRow: r => applyGo({nguon: r.n === "(trống)" ? null : r.n, year: Y, loai: ctx.L, acc: ctx.A})});
  const ins = [];
  const topS = rows.slice().sort((a, b) => b.t - a.t)[0];
  if (topS && topS.share >= 80 && rows.length > 1) ins.push(I("warn", `Phụ thuộc một nguồn: <b>${esc(topS.label)}</b> chiếm ${pctS(topS.share)} KH — rủi ro nếu kênh này giảm hiệu quả`));
  if (best != null && big.length > 1) { const b = big.find(r => r.rate === best); if (b.ky >= 3) ins.push(I("good", `Nguồn chuyển đổi tốt nhất: <b>${esc(b.label)}</b> ${pctS(b.rate)} (${b.ky}/${fmt(b.t)} KH), gấp ${fmt(b.rate / Math.max(allRate, .01))} lần trung bình`)); }
  big.filter(r => r.share >= 30 && signedOf(rs) >= 5 && r.rate < allRate / 2).forEach(r => ins.push(I("warn", `<b>${esc(r.label)}</b> mang ${pctS(r.share)} KH nhưng tỷ lệ ký chỉ ${pctS(r.rate)} — dưới một nửa trung bình (${pctS(allRate)})`)));
  rows.filter(r => r.t < 20 && r.ky).forEach(r => ins.push(I("info", `Tỷ lệ ký của ${esc(r.label)} (${pctS(r.rate)}) chỉ dựa trên ${r.t} KH — chưa đủ để kết luận`)));
  ins.push(...issueIns(rs, ["FB_NO_QC"], ctx));
  if (rows.some(r => r.n === "(trống)")) ins.push(I("crit", `${fmt(rows.find(r => r.n === "(trống)").t)} KH không ghi Nguồn`));
  renderIns("#inSource", ins);
}

// ---- Loại KH
function renderLoai(ctx) {
  const {rs, prevSP, PY, Y} = ctx, tot = rs.length, allRate = pctN(signedOf(rs), tot);
  const rows = [...LOAI.map(l => l[0]), "(trống)"].map(k => {
    const g = rs.filter(r => (r.loai || "(trống)") === k); if (!g.length) return null;
    const p = prevSP.filter(r => (r.loai || "(trống)") === k);
    const vals = g.filter(r => r.trangThai === "Đã ký HĐ").map(r => parseMoney(r.giaTri)).filter(Boolean);
    return {k, t: g.length, share: pctN(g.length, tot), prev: ctx.hasPrev ? p.length : null, act: activeOf(g), ky: signedOf(g), rate: pctN(signedOf(g), g.length), end: pctN(endedOf(g), g.length), avg: vals.length ? sum(vals) / vals.length : null, stale: issueCount(g, "STALE")};
  }).filter(Boolean);
  table("#tbLoai", [
    {k: "k", label: "Loại"},
    {k: "t", label: "KH", num: true, html: r => fmt(r.t) + (r.prev != null ? delta(r.t, r.prev) : ""), title: ctx.hasPrev ? `So với cùng kỳ ${PY}` : ""},
    {k: "share", label: "Tỷ trọng", num: true, html: r => pctS(r.share)},
    {k: "ky", label: "Đã ký", num: true, html: r => fmt(r.ky)},
    {k: "rate", label: "Tỷ lệ ký", num: true, html: r => pctS(r.rate), cls: r => r.t >= 30 && allRate && r.rate >= 2 * allRate && r.ky >= 2 ? "good" : ""},
    {k: "end", label: "Kết thúc", num: true, html: r => pctS(r.end)},
    {k: "stale", label: "CS > 6 tháng", num: true, html: r => fmt(r.stale), cls: r => r.stale && r.stale / r.t > .3 ? "flag" : ""},
    {k: "avg", label: "Giá trị TB/HĐ", num: true, html: r => r.avg ? vnd(r.avg) : "—"},
  ], rows, {def: {k: "t", dir: "desc"}, onRow: r => applyGo({loai: r.k === "(trống)" ? null : r.k, year: Y, acc: ctx.A})});
  const ins = [];
  const big = rows.filter(r => r.t >= 30 && r.k !== "(trống)");
  if (big.length > 1) {
    const b = big.slice().sort((x, y) => y.rate - x.rate)[0], w = big.slice().sort((x, y) => x.rate - y.rate)[0];
    if (b.ky >= 3 && b !== w) ins.push(I("good", `<b>${b.k}</b> có tỷ lệ ký cao nhất: ${pctS(b.rate)} · <b>${w.k}</b> thấp nhất: ${pctS(w.rate)} dù chiếm ${pctS(w.share)} KH`));
  }
  rows.filter(r => r.stale / r.t > .5 && r.t >= 30).forEach(r => ins.push(I("warn", `<b>${pctS(pctN(r.stale, r.t))}</b> KH loại ${r.k} đang “Chăm sóc” từ hơn 6 tháng — nhiều khả năng danh sách không được cập nhật`, {issue: "STALE", loai: r.k, year: Y})));
  ins.push(...issueIns(rs, ["NO_LOAI"], ctx));
  renderIns("#inLoai", ins);
}

// ---- Lý do kết thúc
function renderReason(ctx) {
  const {rs, prevSP, PY, Y} = ctx;
  const ended = rs.filter(r => r.trangThai === "Kết thúc"), ly = count(ended, r => LYDO[r.lyDo] || (r.lyDo ? "Mã " + r.lyDo : "Chưa ghi lý do"));
  const pEnded = prevSP.filter(r => r.trangThai === "Kết thúc"), ply = count(pEnded, r => LYDO[r.lyDo] || (r.lyDo ? "Mã " + r.lyDo : "Chưa ghi lý do"));
  const entries = Object.entries(ly).sort((a, b) => b[1] - a[1]).map(([k, v]) => [k, v, ctx.hasPrev && pEnded.length ? `Cùng kỳ ${PY}: ${pct(ply[k] || 0, pEnded.length)}` : ""]);
  hbars("#chReason", entries, ended.length, "var(--s2)");
  const ins = [];
  if (ended.length) {
    const [tk, tv] = entries[0];
    ins.push(I("info", `Lý do phổ biến nhất: <b>${esc(tk)}</b> — ${pct(tv, ended.length)} số KH kết thúc`));
    const thamKhao = ly["Chỉ tham khảo"] || 0; if (thamKhao / ended.length >= .4) ins.push(I("warn", `${pct(thamKhao, ended.length)} KH kết thúc vì <b>chỉ tham khảo</b> → lead đầu vào chưa được lọc nhu cầu; nên hỏi thời điểm xây/sửa nhà ngay từ đầu`));
    const knm = ly["Không nghe máy"] || 0; if (knm / ended.length >= .15) ins.push(I("info", `${pct(knm, ended.length)} kết thúc do <b>không nghe máy</b> — có thể liên hệ lại qua Zalo/SMS hoặc gọi khung giờ khác trước khi đóng`));
    const gia = ly["Giá cao"] || 0; if (gia / ended.length >= .15) ins.push(I("warn", `${pct(gia, ended.length)} kết thúc vì <b>giá cao</b> — cân nhắc giới thiệu gói Lite / chính sách chiết khấu`));
    if (ctx.hasPrev && pEnded.length >= 20) for (const [k, v] of Object.entries(ly)) { const a = pctN(v, ended.length), b = pctN(ply[k] || 0, pEnded.length); if (Math.abs(a - b) >= 10) ins.push(I("info", `“${esc(k)}” ${a > b ? "tăng" : "giảm"} từ ${pctS(b)} (cùng kỳ ${PY}) ${a > b ? "lên" : "xuống"} ${pctS(a)}`)); }
  }
  ins.push(...issueIns(rs, ["END_NO_REASON", "REASON_NOT_END"], ctx));
  renderIns("#inReason", ins);
}

// ---- Phân khúc đại lý
function renderSeg(ctx) {
  const {rs, Y} = ctx;
  const dlAll = rs.filter(r => r.loai === "ĐL"), dl = dlAll.filter(r => r.phanKhuc), pk = count(dl, r => PK[r.phanKhuc] || r.phanKhuc);
  const sgn = k => dl.filter(r => (PK[r.phanKhuc] || r.phanKhuc) === k);
  hbars("#chSeg", ["Cao", "Trung", "Giá rẻ", "Trung lập"].map(k => [k, pk[k] || 0, `Đã ký: ${signedOf(sgn(k))} · Kết thúc: ${pct(endedOf(sgn(k)), sgn(k).length)}`]), dl.length);
  const ins = [];
  if (dl.length) {
    const low = (pk["Giá rẻ"] || 0) + (pk["Trung lập"] || 0);
    if (low / dl.length >= .7) ins.push(I("warn", `<b>${pct(low, dl.length)}</b> đại lý thuộc phân khúc giá rẻ/trung lập, chỉ ${fmt((pk["Cao"] || 0) + (pk["Trung"] || 0))} ĐL phân khúc cao/trung → chưa khớp định vị sản phẩm cao cấp`));
    for (const k of ["Cao", "Trung", "Giá rẻ", "Trung lập"]) { const g = sgn(k); if (g.length >= 30) { const e = pctN(endedOf(g), g.length); if (e >= 60) ins.push(I("info", `ĐL phân khúc ${k}: ${pctS(e)} đã kết thúc`)); } }
  }
  if (dlAll.length) { const miss = issueCount(dlAll, "DL_NO_SEG"); if (miss / dlAll.length >= .2) ins.push(I("warn", `<b>${fmt(miss)}</b>/${fmt(dlAll.length)} đại lý (${pct(miss, dlAll.length)}) chưa được đánh giá phân khúc → biểu đồ chưa đại diện`, {issue: "DL_NO_SEG", year: Y, acc: ctx.A})); }
  renderIns("#inSeg", ins);
}

// ---- Chiến dịch quảng cáo
function renderQC(ctx) {
  const {rs, Y} = ctx, withQC = rs.filter(r => r.nguonQC), avgEnd = pctN(endedOf(withQC), withQC.length);
  const rows = [...new Set(withQC.map(r => r.nguonQC))].map(c => { const g = withQC.filter(r => r.nguonQC === c); return {c, label: QC[c] || "Mã " + c, t: g.length, share: pctN(g.length, withQC.length), act: activeOf(g), ky: signedOf(g), rate: pctN(signedOf(g), g.length), end: pctN(endedOf(g), g.length), dl: pctN(g.filter(r => r.loai === "ĐL").length, g.length)}; });
  table("#tbQC", [
    {k: "label", label: "Chiến dịch"}, {k: "t", label: "KH", num: true, html: r => fmt(r.t)},
    {k: "share", label: "Tỷ trọng", num: true, html: r => pctS(r.share)},
    {k: "dl", label: "% là ĐL", num: true, html: r => pctS(r.dl), title: "Tỷ lệ KH loại Đại lý trong chiến dịch"},
    {k: "ky", label: "Đã ký", num: true, html: r => fmt(r.ky)}, {k: "rate", label: "Tỷ lệ ký", num: true, html: r => pctS(r.rate)},
    {k: "end", label: "Kết thúc", num: true, html: r => pctS(r.end), cls: r => r.t >= 30 && r.end - avgEnd >= 15 ? "bad" : r.t >= 30 && avgEnd - r.end >= 15 ? "good" : ""},
  ], rows, {def: {k: "t", dir: "desc"}, empty: "Không có KH nào gắn Nguồn QC trong bộ lọc này",
    foot: rows.length ? ["Tổng", fmt(withQC.length), "100%", pct(withQC.filter(r => r.loai === "ĐL").length, withQC.length), fmt(signedOf(withQC)), pct(signedOf(withQC), withQC.length), pctS(avgEnd)] : null});
  const ins = [];
  rows.filter(r => r.t >= 50 && !r.ky).forEach(r => ins.push(I("warn", `<b>${esc(r.label)}</b>: ${fmt(r.t)} KH nhưng chưa ký được HĐ nào`)));
  rows.filter(r => r.t >= 30 && r.end - avgEnd >= 15).forEach(r => ins.push(I("warn", `<b>${esc(r.label)}</b> có tỷ lệ kết thúc ${pctS(r.end)}, cao hơn trung bình các chiến dịch (${pctS(avgEnd)}) → xem lại target quảng cáo`)));
  const mod = rows.find(r => r.c === "3"); if (mod && mod.dl < 20 && mod.t >= 30) ins.push(I("info", `Chiến dịch module chủ yếu mang về KH cuối (${pctS(100 - mod.dl)}) — phù hợp nếu mục tiêu là bán lẻ module`));
  if (rs.length && withQC.length / rs.length < .5 && rs.some(r => +r.nam >= 2025)) ins.push(I("info", `Chỉ ${pct(withQC.length, rs.length)} KH có gắn Nguồn QC`));
  ins.push(...issueIns(rs, ["FB_NO_QC"], ctx));
  renderIns("#inQC", ins);
}

// ---- Khu vực
function renderKV(ctx) {
  const {rs, prev, PY, Y} = ctx, tot = rs.length;
  const rows = ["HN", "Bắc", "HCM", "Nam", "(trống)"].map(k => { const g = rs.filter(r => (r.kv || "(trống)") === k); if (!g.length) return null; const p = prev.filter(r => (r.kv || "(trống)") === k).length;
    return {k, t: g.length, share: pctN(g.length, tot), pshare: ctx.hasPrev ? pctN(p, prev.length) : null, d: ctx.hasPrev ? pctN(g.length, tot) - pctN(p, prev.length) : null, ky: signedOf(g), rate: pctN(signedOf(g), g.length)}; }).filter(Boolean);
  table("#tbKV", [
    {k: "k", label: "Khu vực"}, {k: "t", label: "KH", num: true, html: r => fmt(r.t)}, {k: "share", label: "Tỷ trọng", num: true, html: r => pctS(r.share)},
    {k: "pshare", label: ctx.hasPrev ? `Năm ${PY}` : "Năm trước", num: true, html: r => pctS(r.pshare)},
    {k: "d", label: "Chênh", num: true, html: r => r.d == null ? "—" : `${r.d > 0 ? "+" : ""}${fmt(r.d)}`, cls: r => r.d == null ? "" : Math.abs(r.d) >= 10 ? "flag" : "", title: "Chênh lệch tỷ trọng (điểm %)"},
    {k: "ky", label: "Đã ký", num: true, html: r => fmt(r.ky)}, {k: "rate", label: "Tỷ lệ ký", num: true, html: r => pctS(r.rate)},
  ], rows, {def: {k: "t", dir: "desc"}, foot: ["Tổng", fmt(tot), "100%", "", "", fmt(signedOf(rs)), pct(signedOf(rs), tot)]});
  const ins = [];
  rows.filter(r => r.d != null && Math.abs(r.d) >= 10).forEach(r => ins.push(I("info", `Tỷ trọng <b>${r.k}</b> ${r.d > 0 ? "tăng" : "giảm"} ${fmt(Math.abs(r.d))} điểm % so với năm ${PY} (${pctS(r.pshare)} → ${pctS(r.share)})`)));
  const prov = Object.entries(count(rs.filter(r => r.tinh && r.tinh !== "HN" && r.tinh !== "HCM"), r => r.tinh)).sort((a, b) => b[1] - a[1]).slice(0, 5);
  if (prov.length) ins.push(I("info", `Tỉnh nhiều KH nhất (ngoài HN, HCM): ${prov.map(([t, n]) => `${esc(t)} <b>${n}</b>`).join(" · ")}`));
  ins.push(...issueIns(rs, ["KV_MISMATCH", "UNKNOWN_TINH", "NO_TINH"], ctx));
  renderIns("#inKV", ins);
}

// ---- Người phụ trách
function renderAcc(ctx) {
  const {rs, Y} = ctx, tot = rs.length;
  const avgEnd = pctN(endedOf(rs), tot), avgRate = pctN(signedOf(rs), tot);
  const rows = [...new Set(rs.map(r => r.acc || "(trống)"))].map(a => { const g = rs.filter(r => (r.acc || "(trống)") === a);
    return {a, t: g.length, share: pctN(g.length, tot), act: activeOf(g), stale: issueCount(g, "STALE"), end: pctN(endedOf(g), g.length), ky: signedOf(g), rate: pctN(signedOf(g), g.length),
      bad: g.filter(r => issuesOf(r).some(c => ISSUES[c].sev !== "info")).length, noReason: issueCount(g, "END_NO_REASON")}; });
  table("#tbAcc", [
    {k: "a", label: "ACC", html: r => `<b>${esc(r.a)}</b>${accVariants[r.a.toLowerCase()] ? '<span class="warnmark" title="Có tên viết khác hoa/thường">!</span>' : ""}`},
    {k: "t", label: "Tổng KH", num: true, html: r => fmt(r.t)}, {k: "share", label: "Tỷ trọng", num: true, html: r => pctS(r.share)},
    {k: "act", label: "Đang CS", num: true, html: r => fmt(r.act)},
    {k: "stale", label: "CS > 6 tháng", num: true, html: r => fmt(r.stale), cls: r => r.act >= 20 && r.stale / r.act > .5 ? "flag" : ""},
    {k: "end", label: "Kết thúc", num: true, html: r => pctS(r.end), cls: r => r.t >= 30 && r.end - avgEnd >= 15 ? "bad" : r.t >= 30 && avgEnd - r.end >= 15 ? "good" : ""},
    {k: "ky", label: "Đã ký", num: true, html: r => fmt(r.ky)},
    {k: "rate", label: "Tỷ lệ ký", num: true, html: r => pctS(r.rate), cls: r => r.t >= 30 && r.ky >= 2 && r.rate >= 2 * avgRate ? "good" : ""},
    {k: "bad", label: "Dòng có vấn đề", num: true, html: r => fmt(r.bad), cls: r => r.bad / r.t > .3 ? "flag" : "", title: "Số dòng có lỗi hoặc cảnh báo dữ liệu"},
  ], rows, {def: {k: "t", dir: "desc"}, foot: ["Cả nhóm", fmt(tot), "100%", fmt(activeOf(rs)), fmt(issueCount(rs, "STALE")), pctS(avgEnd), fmt(signedOf(rs)), pctS(avgRate), fmt(rs.filter(r => issuesOf(r).some(c => ISSUES[c].sev !== "info")).length)],
    onRow: r => applyGo({acc: r.a === "(trống)" ? null : r.a, year: Y, loai: ctx.L})});
  const ins = [];
  Object.values(accVariants).forEach(v => { if (rs.some(r => v.includes(r.acc))) ins.push(I("warn", `Tên ACC <b>${v.map(esc).join("</b> và <b>")}</b> chỉ khác hoa/thường — nếu là một người, số liệu đang bị tách làm ${v.length} dòng`, {issue: "ACC_CASE", year: Y})); });
  const top = rows.slice().sort((a, b) => b.t - a.t)[0];
  if (top && rows.length > 2 && top.share >= 40) ins.push(I("info", `<b>${esc(top.a)}</b> phụ trách ${pctS(top.share)} tổng KH — khối lượng tập trung vào một người`));
  rows.filter(r => r.t >= 30 && r.end - avgEnd >= 15).forEach(r => ins.push(I("warn", `<b>${esc(r.a)}</b>: tỷ lệ kết thúc ${pctS(r.end)}, cao hơn trung bình nhóm ${fmt(r.end - avgEnd)} điểm %`, {acc: r.a, st: "Kết thúc", year: Y})));
  rows.filter(r => r.act >= 20 && r.stale / r.act > .5).forEach(r => ins.push(I("info", `<b>${esc(r.a)}</b>: ${fmt(r.stale)}/${fmt(r.act)} KH đang chăm sóc đã quá 6 tháng`, {acc: r.a, issue: "STALE", year: Y})));
  rows.filter(r => r.noReason >= 5).forEach(r => ins.push(I("warn", `<b>${esc(r.a)}</b> có ${r.noReason} KH kết thúc không ghi lý do`, {acc: r.a, issue: "END_NO_REASON", year: Y})));
  renderIns("#inAcc", ins);
}

// ---- Hợp đồng đã ký
function renderContracts(ctx) {
  const {rs, Y} = ctx, sg = rs.filter(r => r.trangThai === "Đã ký HĐ");
  const vals = sg.map(r => parseMoney(r.giaTri)).filter(Boolean);
  const q1 = quantile(vals, .25), q3 = quantile(vals, .75), iqr = q3 - q1, hiT = q3 + 1.5 * iqr, loT = Math.max(0, q1 - 1.5 * iqr);
  const outlier = r => { const v = parseMoney(r.giaTri); return vals.length >= 5 && v && (v > hiT || v < loT); };
  $("#subContracts").textContent = sg.length ? `${sg.length} hợp đồng · bấm tiêu đề cột để sắp xếp, bấm dòng để xem chi tiết` : "Chưa có hợp đồng trong bộ lọc này";
  table("#tbContracts", [
    {k: "ngay", label: "Ngày nhận", html: r => dmy(r.ngay) || "—"},
    {k: "ten", label: "Khách hàng", html: r => esc(r.ten || "—")},
    {k: "sdt", label: "Điện thoại"},
    {k: "loai", label: "Loại"}, {k: "tinh", label: "Tỉnh"}, {k: "kv", label: "KV"},
    {k: "nguon", label: "Nguồn", html: r => esc(NGUON_LABEL[r.nguon] || r.nguon || "")},
    {k: "acc", label: "ACC"}, {k: "giaiPhap", label: "Giải pháp"},
    {k: "v", label: "Giá trị", num: true, val: r => parseMoney(r.giaTri) || null,
      html: r => { const v = parseMoney(r.giaTri); const flags = issuesOf(r).filter(c => c === "SIGNED_NO_VALUE" || c === "VALUE_ODD"); return (v ? vnd(v) : "—") + (flags.length ? `<span class="warnmark" title="${esc(flags.map(c => ISSUES[c].label).join(", "))}${r.giaTri ? " · gốc: " + r.giaTri : ""}">!</span>` : outlier(r) ? `<span class="warnmark" title="Lệch xa so với các HĐ khác">↕</span>` : ""); },
      cls: r => outlier(r) ? "flag" : ""},
  ], sg, {def: {k: "ngay", dir: "desc"}, empty: "",
    foot: sg.length ? [`${sg.length} HĐ`, "", "", "", "", "", "", "", "Tổng / TB / Trung vị", vals.length ? `${vnd(sum(vals))} / ${vnd(sum(vals) / vals.length)} / ${vnd(median(vals))}` : "—"] : null,
    onRow: r => openEdit(r.id)});
  const ins = [];
  if (sg.length) {
    const miss = sg.length - vals.length;
    if (miss) ins.push(I("warn", `<b>${miss}/${sg.length}</b> HĐ chưa ghi giá trị → tổng giá trị đang thấp hơn thực tế`, {issue: "SIGNED_NO_VALUE", year: Y, loai: ctx.L, acc: ctx.A}));
    const odd = issueCount(sg, "VALUE_ODD"); if (odd) ins.push(I("warn", `<b>${odd}</b> HĐ có giá trị ghi dạng chữ hoặc < 1 triệu (có thể thiếu số 0)`, {issue: "VALUE_ODD", year: Y}));
    if (vals.length >= 5) {
      const outs = sg.filter(outlier); if (outs.length) ins.push(I("info", `HĐ lệch xa mức thông thường (${loT > 0 ? vnd(loT) + "–" + vnd(hiT) : "trên " + vnd(hiT)}): ${outs.map(r => `${esc(r.ten || r.sdt)} <b>${vnd(parseMoney(r.giaTri))}</b>`).join(" · ")}`));
      ins.push(I("info", `Giá trị HĐ: trung vị <b>${vnd(median(vals))}</b>, 50% HĐ nằm trong khoảng ${vnd(q1)}–${vnd(q3)}`));
    }
    const gp = count(sg.filter(r => r.giaiPhap), r => /^lux/i.test(r.giaiPhap) ? "Luxury" : r.giaiPhap);
    if (Object.keys(gp).length) ins.push(I("info", `Theo giải pháp: ${Object.entries(gp).sort((a, b) => b[1] - a[1]).map(([k, n]) => `${esc(k)} <b>${n}</b>`).join(" · ")}` + (sg.length - sum(Object.values(gp)) ? ` · chưa ghi ${sg.length - sum(Object.values(gp))}` : "")));
    const lo = count(sg, r => r.loai); ins.push(I("info", `Theo loại KH: ${Object.entries(lo).map(([k, n]) => `${esc(k)} <b>${n}</b>`).join(" · ")}`));
    const last = sg.map(r => r.ngay).filter(Boolean).sort().pop(); if (last && daysBetween(last, TODAY) > 60) ins.push(I("warn", `KH ký HĐ gần nhất được nhận từ <b>${dmy(last)}</b> (${daysBetween(last, TODAY)} ngày trước)`));
  }
  renderIns("#inContracts", ins);
}

// ================= Danh sách =================
let page = 0; const PS = 50;
["#lYear", "#lLoai", "#lSt", "#lAcc", "#lNguon", "#lIssue"].forEach(s => $(s).onchange = () => { page = 0; renderList(); });
let qT; $("#q").oninput = () => { clearTimeout(qT); qT = setTimeout(() => { page = 0; renderList(); }, 200); };
$("#pgPrev").onclick = () => { if (page > 0) { page--; renderList(); } };
$("#pgNext").onclick = () => { page++; renderList(); };
$("#lReset").onclick = () => { resetListFilters(); page = 0; renderList(); };
function resetListFilters() { ["#lYear", "#lLoai", "#lSt", "#lAcc", "#lNguon", "#lIssue"].forEach(s => $(s).value = ""); $("#q").value = ""; }
function filtered() {
  const Y = $("#lYear").value, L = $("#lLoai").value, St = $("#lSt").value, A = $("#lAcc").value, N = $("#lNguon").value, X = $("#lIssue").value, q = $("#q").value.trim().toLowerCase(), qd = phoneKey(q);
  return all.filter(r => (!Y || r.nam === Y) && (!L || r.loai === L) && (!St || r.trangThai === St) && (!A || r.acc === A) && (!N || r.nguon === N) && (!X || hasIssue(r, X)) &&
    (!q || ([r.ten, r.tinh, r.email, r.nganh, r.thongTin, r.nhatKy, r.diaChi, r.acc, r.sdt].some(v => v && String(v).toLowerCase().includes(q)) || weeklyEntries(r).some(e => e.text.toLowerCase().includes(q))) || (qd.length >= 4 && phonesOf(r.sdt).some(p => p.includes(qd)))));
}
const LIST_COLS = [
  {k: "ngay", label: "Ngày nhận", html: r => dmy(r.ngay) || '<span class="muted">—</span>'},
  {k: "ten", label: "Khách hàng", html: r => esc(r.ten || "—") + (r._new ? '<span class="tag">mới</span>' : isEdited(r) ? '<span class="tag">đã sửa</span>' : "")},
  {k: "sdt", label: "Điện thoại"}, {k: "loai", label: "Loại"}, {k: "tinh", label: "Tỉnh"}, {k: "nguon", label: "Nguồn"}, {k: "acc", label: "ACC"},
  {k: "trangThai", label: "Trạng thái", html: r => r.trangThai ? `<span class="chip" data-s="${esc(r.trangThai)}">${esc(r.trangThai)}</span>` : ""},
  {k: "iss", label: "Vấn đề", num: true, val: r => issuesOf(r).filter(c => ISSUES[c].sev !== "info").length || null,
    html: r => { const l = issuesOf(r).filter(c => ISSUES[c].sev !== "info"); return l.length ? `<span class="warnmark" title="${esc(l.map(c => ISSUES[c].label + (issueDetail(r, c) ? ": " + issueDetail(r, c) : "")).join("\n"))}">${l.length}</span>` : ""; }},
  {k: "nhatKy", label: "Nhật ký gần nhất", val: r => { const e = weeklyEntries(r)[0]; return e ? e.y * 100 + e.w : null; }, html: r => `<span class="muted clip" style="display:inline-block">${esc(latestNote(r).split("\n")[0])}</span>`},
];
function renderList() {
  const rs = sortRows("#tbList", LIST_COLS, filtered(), {k: "ngay", dir: "desc"});
  const pages = Math.max(1, Math.ceil(rs.length / PS)); if (page >= pages) page = pages - 1;
  table("#tbList", LIST_COLS, rs.slice(page * PS, page * PS + PS), {presorted: true, rerender: () => { page = 0; renderList(); }, onRow: r => openEdit(r.id), empty: "Không có khách hàng khớp bộ lọc", pad: "24px"});
  $("#pgInfo").textContent = `${fmt(rs.length)} KH · trang ${page + 1}/${pages}`;
  $("#pgPrev").disabled = page === 0; $("#pgNext").disabled = page >= pages - 1;
}
const isEdited = r => !!r._wkEdited || (r._orig && CORE.some(([k]) => (r[k] ?? "") !== (r._orig[k] ?? "")));

// ================= Kiểm tra dữ liệu =================
["#cYear", "#cSev"].forEach(s => $(s).onchange = renderCheck);
function renderCheck() {
  if (!S) return;
  const Y = $("#cYear").value, sev = $("#cSev").value;
  const rs = Y ? all.filter(r => r.nam === Y) : all;
  const withSev = s => rs.filter(r => issuesOf(r).some(c => ISSUES[c].sev === s)).length;
  const clean = rs.filter(r => !issuesOf(r).some(c => ISSUES[c].sev !== "info")).length;
  $("#checkKpis").innerHTML = [
    ["Số dòng kiểm tra", fmt(rs.length), Y ? `Năm ${Y}` : "Tất cả các năm", ""],
    ["Dòng không có lỗi/cảnh báo", pct(clean, rs.length), `${fmt(clean)} dòng`, "hi"],
    ["Dòng có lỗi", fmt(withSev("crit")), "Làm sai số liệu thống kê", ""],
    ["Dòng có cảnh báo", fmt(withSev("warn")), "Nên kiểm tra lại", ""],
    ["Dòng có lưu ý", fmt(withSev("info")), "Thông tin tham khảo", ""],
  ].map(([l, v, s, c]) => `<div class="kpi ${c}"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`).join("");
  const rows = Object.entries(ISSUES).map(([code, d]) => ({code, ...d, n: issueCount(rs, code), share: pctN(issueCount(rs, code), rs.length), rank: SEV_RANK[d.sev]})).filter(r => r.n && (!sev || r.sev === sev));
  table("#tbCheck", [
    {k: "rank", label: "Mức độ", val: r => r.rank, html: r => `<span class="sevtag"><span class="sev sev-${r.sev}">${SEV_ICON[r.sev]}</span>${SEV_LABEL[r.sev]}</span>`},
    {k: "label", label: "Vấn đề", html: r => `<b>${esc(r.label)}</b><div class="muted" style="white-space:normal;max-width:460px;font-size:12px">${esc(r.desc)}</div>`},
    {k: "affects", label: "Ảnh hưởng tới", html: r => `<span style="white-space:normal">${esc(r.affects)}</span>`},
    {k: "n", label: "Số dòng", num: true, html: r => fmt(r.n)},
    {k: "share", label: "% dòng", num: true, html: r => pctS(r.share)},
  ], rows, {def: {k: "rank", dir: "asc"}, empty: "Không phát hiện vấn đề nào", onRow: r => applyGo({issue: r.code, year: Y || null})});
  // Đối chiếu với Excel
  const years = yearsList();
  const recon = years.map(y => {
    const web = all.filter(r => r.nam === y).length;
    const excel = all.filter(r => r.nam === y && !r._namDerived && !r._new && !(r._orig && r._orig.nam !== r.nam)).length;
    const derived = all.filter(r => r.nam === y && r._namDerived).length, added = all.filter(r => r.nam === y && r._new).length;
    const spell = all.filter(r => r.nam === y && hasIssue(r, "STATUS_SPELL")).length;
    return {y, web, excel, derived, added, spell, deleted: S.records.filter(r => r._deleted && r._orig && r._orig.nam === y).length};
  });
  table("#reconcile", [
    {k: "y", label: "Năm"}, {k: "web", label: "Số KH trên web", num: true, html: r => fmt(r.web)},
    {k: "excel", label: "Excel đếm được (theo cột Năm gốc)", num: true, html: r => fmt(r.excel)},
    {k: "derived", label: "Năm suy ra từ ngày nhận", num: true, html: r => r.derived ? fmt(r.derived) : "—", cls: r => r.derived ? "flag" : ""},
    {k: "added", label: "Thêm/sửa trên web", num: true, html: r => r.added ? fmt(r.added) : "—"},
    {k: "deleted", label: "Đã xoá trên web", num: true, html: r => r.deleted ? fmt(r.deleted) : "—"},
    {k: "spell", label: "Trạng thái viết sai (Excel đếm sót)", num: true, html: r => r.spell ? fmt(r.spell) : "—", cls: r => r.spell ? "flag" : ""},
  ], recon, {def: {k: "y", dir: "desc"}, onRow: r => applyGo({year: r.y, issue: r.derived ? "YEAR_DERIVED" : null})});
}

// ================= Xuất file Excel =================
function outValue(r, k) {
  const v = r[k] ?? "";
  if (r._orig && v === (r._orig[k] ?? "") && r._raw && S.colOf[k] != null && !(k === "nam" && r._namDerived)) return r._raw[S.colOf[k]] ?? null; // không đổi → giữ giá trị gốc
  if (v === "") return null;
  if (k === "ngay") { const d = new Date(v + "T00:00:00"); return isNaN(d) ? v : d; }
  if (NUMERIC_KEYS.has(k) && /^\d+$/.test(v)) return Number(v);
  return v;
}
function buildListSheet() {
  const header = S.header.slice();
  const colOf = {...S.colOf};
  for (const [k, label] of CORE) if (colOf[k] == null && all.some(r => r[k])) { colOf[k] = header.length; header.push(label); }
  const aoa = [];
  for (let i = 0; i < (S.headerRow || 0); i++) aoa.push([]);
  aoa.push(header);
  for (const r of all) {
    const row = (r._raw || []).slice(); while (row.length < header.length) row.push(null);
    for (const [k] of CORE) if (colOf[k] != null) row[colOf[k]] = outValue(r, k);
    aoa.push(row);
  }
  const dateCells = [];
  aoa.forEach((row, R) => row.forEach((v, C) => { if (v instanceof Date) { row[C] = dateToSerial(v); dateCells.push([R, C]); } }));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (const [R, C] of dateCells) { const c = ws[XLSX.utils.encode_cell({r: R, c: C})]; if (c) { c.t = "n"; c.z = "d/m/yyyy"; } }
  ws["!cols"] = header.map((h, i) => ({wch: i === colOf.thongTin || i === colOf.nhatKy ? 50 : Math.min(28, Math.max(8, String(h).length + 2))}));
  return ws;
}
function buildStatsSheet() {
  const years = yearsList().slice().reverse();
  const aoa = [["THỐNG KÊ TỪ WEB CSKH", "", "Xuất lúc " + new Date().toLocaleString("vi-VN")], []];
  const block = (title, rowsDef) => {
    aoa.push([title, ...years, "Tổng"]);
    for (const [label, test] of rowsDef) { const cells = years.map(y => all.filter(r => r.nam === y && test(r)).length); aoa.push([label, ...cells, sum(cells)]); }
    aoa.push([]);
  };
  block("Loại KH theo năm", LOAI.map(([k]) => [k, r => r.loai === k]));
  block("Nguồn theo năm", NGUON.map(k => [NGUON_LABEL[k], r => r.nguon === k]));
  block("Trạng thái theo năm", TT.map(k => [k, r => r.trangThai === k]));
  block("Lý do kết thúc theo năm", Object.entries(LYDO).map(([k, l]) => [l, r => r.trangThai === "Kết thúc" && r.lyDo === k]));
  block("Khu vực theo năm", ["HN", "Bắc", "HCM", "Nam"].map(k => [k, r => r.kv === k]));
  aoa.push(["Tỷ lệ ký HĐ theo năm", ...years.map(y => { const g = all.filter(r => r.nam === y); return g.length ? signedOf(g) / g.length : 0; })]);
  aoa.push(["Tỷ lệ kết thúc theo năm", ...years.map(y => { const g = all.filter(r => r.nam === y); return g.length ? endedOf(g) / g.length : 0; })]);
  aoa.push([]);
  const ys = years[years.length - 1];
  if (ys) {
    aoa.push([`Loại KH theo tháng · ${ys}`, ...Array.from({length: 12}, (_, i) => "T" + (i + 1)), "Tổng"]);
    for (const [k] of LOAI) { const c = Array.from({length: 12}, (_, i) => all.filter(r => r.nam === ys && r.loai === k && +r.thang === i + 1).length); aoa.push([k, ...c, sum(c)]); }
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa); ws["!cols"] = [{wch: 26}];
  const rg = XLSX.utils.decode_range(ws["!ref"]);
  for (let R = rg.s.r; R <= rg.e.r; R++) { const a = ws[XLSX.utils.encode_cell({r: R, c: 0})]; if (a && /^Tỷ lệ/.test(a.v)) for (let C = 1; C <= rg.e.c; C++) { const c = ws[XLSX.utils.encode_cell({r: R, c: C})]; if (c && typeof c.v === "number") c.z = "0.0%"; } }
  return ws;
}
function buildIssueSheet() {
  const aoa = [["Mức độ", "Vấn đề", "Chi tiết", "Dòng Excel gốc", "Năm", "Ngày nhận", "Tên KH", "Điện thoại", "ACC", "Loại KH", "Trạng thái", "Tỉnh", "Khu vực", "Giá trị HĐ", "Gợi ý"]];
  const tip = {NO_YEAR: "Điền cột Năm", YEAR_MISMATCH: "Sửa Năm hoặc Ngày nhận cho khớp", NO_MONTH: "Điền cột Tháng", MONTH_MISMATCH: "Sửa Tháng cho khớp Ngày nhận", KV_MISMATCH: "Sửa khu vực theo tỉnh", END_NO_REASON: "Chọn lý do kết thúc", SIGNED_NO_VALUE: "Điền giá trị HĐ", VALUE_ODD: "Ghi giá trị dạng số đầy đủ (VNĐ)", DUP_PHONE: "Kiểm tra và gộp dòng trùng", BAD_PHONE: "Sửa SĐT đủ 10 số", STATUS_SPELL: "Chọn lại trạng thái từ danh sách", ACC_CASE: "Thống nhất cách viết tên ACC", STALE: "Cập nhật trạng thái"};
  for (const r of all) for (const c of issuesOf(r)) { const d = ISSUES[c]; if (d.sev === "info" && c !== "STALE") continue;
    aoa.push([SEV_LABEL[d.sev], d.label, issueDetail(r, c), r._row || (r._new ? "mới thêm" : ""), r.nam || "", r.ngay ? dmy(r.ngay) : "", r.ten || "", r.sdt || "", r.acc || "", r.loai || "", r.trangThai || "", r.tinh || "", r.kv || "", r.giaTri || "", tip[c] || ""]); }
  const ws = XLSX.utils.aoa_to_sheet(aoa); ws["!cols"] = [{wch: 10}, {wch: 36}, {wch: 48}, {wch: 8}, {wch: 6}, {wch: 11}, {wch: 22}, {wch: 12}, {wch: 10}, {wch: 7}, {wch: 12}, {wch: 12}, {wch: 7}, {wch: 14}, {wch: 34}];
  ws["!autofilter"] = {ref: XLSX.utils.encode_range({s: {r: 0, c: 0}, e: {r: aoa.length - 1, c: aoa[0].length - 1}})};
  return ws;
}
function putSheet(wb, name, ws) { if (wb.SheetNames.includes(name)) wb.Sheets[name] = ws; else XLSX.utils.book_append_sheet(wb, ws, name); }
async function exportExcel(suffix = "") {
  if (!S) return;
  if (!can.export()) { toast("Tài khoản của bạn không có quyền xuất file"); return; }
  try {
    const wb = origFile ? XLSX.read(origFile, {type: "array", cellFormula: true, cellNF: true}) : XLSX.utils.book_new();
    putSheet(wb, S.sheetName, buildListSheet());
    putSheet(wb, "Thống kê (web)", buildStatsSheet());
    putSheet(wb, "Kiểm tra dữ liệu (web)", buildIssueSheet());
    let out = XLSX.write(wb, {bookType: "xlsx", type: "array", compression: true});
    out = forceRecalc(out);
    const base = (S.fileName || "CSKH").replace(/\.(xlsx|xlsm|xls)$/i, "");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([out], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
    a.download = `${base}_cap-nhat_${isoOf(new Date())}${suffix}.xlsx`;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    toast(`Đã xuất ${fmt(all.length)} khách hàng ra file Excel`);
  } catch (e) { console.error(e); toast("Xuất file thất bại: " + (e.message || e)); }
}
// Bắt Excel tính lại toàn bộ công thức khi mở file (để sheet “Thống kê” gốc cập nhật theo dữ liệu mới)
function forceRecalc(buf) {
  try {
    const cfb = XLSX.CFB.read(new Uint8Array(buf), {type: "array"});
    const entry = XLSX.CFB.find(cfb, "/xl/workbook.xml"); if (!entry) return buf;
    let xml = new TextDecoder().decode(entry.content);
    if (!/<calcPr/.test(xml)) xml = xml.replace("</workbook>", '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
    else xml = xml.replace(/<calcPr([^>]*?)\/?>/, (m, a) => `<calcPr${a.replace(/\s*fullCalcOnLoad="[^"]*"/, "")} fullCalcOnLoad="1"/>`);
    entry.content = new TextEncoder().encode(xml); entry.size = entry.content.length;
    return XLSX.CFB.write(cfb, {fileType: "zip", type: "array", compression: true});
  } catch (e) { console.warn("forceRecalc", e); return buf; }
}

// ================= Form thêm / sửa =================
const opt = (vals, labels, blank = true) => (blank ? `<option value=""></option>` : "") + vals.map(v => `<option value="${esc(v)}">${esc(labels ? labels(v) : v)}</option>`).join("");
function buildForm(rec, {isEdit, onDone}) {
  const form = $("#formTpl").content.cloneNode(true).querySelector("form");
  const pre = isEdit ? "E_" : "A_";
  form.querySelectorAll("[id^=F_]").forEach(el => { const nid = el.id.replace("F_", pre); form.querySelector(`label[for=${el.id}]`)?.setAttribute("for", nid); el.id = nid; });
  const E = n => form.elements[n];
  E("nguon").innerHTML = opt(NGUON, v => NGUON_LABEL[v] + " (" + v + ")");
  E("nguonQC").innerHTML = opt(Object.keys(QC), k => QC[k]);
  E("loai").innerHTML = opt(LOAI.map(l => l[0]), v => LOAI.find(l => l[0] === v)[1]);
  E("congTrinh").innerHTML = opt(CT); E("trangThai").innerHTML = opt(TT, null, false);
  E("lyDo").innerHTML = opt(Object.keys(LYDO), k => LYDO[k]); E("phanKhuc").innerHTML = opt(Object.keys(PK), k => PK[k]); E("giaiPhap").innerHTML = opt(GP);
  const r = rec || {ngay: TODAY, trangThai: "Chăm sóc", nguon: "FB"};
  for (const el of form.elements) {
    if (!el.name || el.name.startsWith("_")) continue;
    const v = r[el.name] ?? "";
    if (el.tagName === "SELECT" && v && ![...el.options].some(o => o.value === v)) el.insertAdjacentHTML("beforeend", `<option value="${esc(v)}">${esc(v)} (ngoài danh mục)</option>`);
    el.value = v;
  }
  const der = form.querySelector("[data-derived]");
  const upd = () => {
    const d = E("ngay").value ? new Date(E("ngay").value + "T00:00:00") : null;
    der.textContent = d && !isNaN(d) ? `Năm ${d.getFullYear()} · Tháng ${d.getMonth() + 1} · Tuần ${weekNum(d)}` : "";
    const ended = E("trangThai").value === "Kết thúc"; E("lyDo").disabled = !ended; if (!ended) E("lyDo").value = "";
  };
  E("ngay").oninput = upd; E("trangThai").onchange = upd; upd();
  E("tinh").onchange = () => { const k = kvOf(E("tinh").value); if (k) E("kv").value = k; };
  const pw = form.querySelector("[data-phone]");
  E("sdt").oninput = () => {
    const ps = phonesOf(E("sdt").value), msgs = [];
    const bad = ps.filter(p => !phoneOk(p));
    if (bad.length) msgs.push(`Số chưa đúng: ${bad.map(p => p || "(không có chữ số)").join(", ")} — di động 10 số, số bàn 11 số. Nhiều số cách nhau bằng dấu ;`);
    const mine = new Set(ps.filter(p => p.length >= 9));
    const dup = mine.size && all.find(x => x.id !== r.id && validPhones(x.sdt).some(p => mine.has(p)));
    if (dup) msgs.push(`Trùng SĐT với ${dup.ten || "KH"} (${dup.acc || "?"}, ${dmy(dup.ngay) || dup.nam || ""}) — vẫn có thể lưu.`);
    pw.textContent = msgs.join(" "); pw.hidden = !msgs.length;
  };
  E("sdt").oninput();
  if (isEdit) { // các vấn đề của dòng này
    const l = issuesOf(r);
    if (l.length) form.insertAdjacentHTML("afterbegin", `<div class="insights" style="grid-column:span 12;margin-top:0;padding-top:0;border-top:0"><h3>Vấn đề dữ liệu của dòng này</h3>${l.map(c => `<div class="ins"><span class="sev sev-${ISSUES[c].sev}">${SEV_ICON[ISSUES[c].sev]}</span><span><b>${esc(ISSUES[c].label)}</b> — ${esc(issueDetail(r, c) || ISSUES[c].desc)}</span></div>`).join("")}</div>`);
  }
  // ---- Nhật ký theo tuần
  // Chọn Năm + Tuần (mặc định tuần hiện tại); tuần nào của KH đã có nhật ký thì đánh dấu ✓
  const now = new Date(), curY = now.getFullYear(), curW = weekNum(now);
  const has = (y, w) => weeklyEntries(r).some(e => e.y === y && e.w === w);
  const wkYears = [...new Set([curY, ...Object.values(S.weekCols || {}).map(v => v.y)])].sort((a, b) => b - a);
  E("_wkYear").innerHTML = wkYears.map(y => `<option value="${y}">${y}</option>`).join("");
  E("_wkYear").value = curY;
  const wkLabel = form.querySelector("[data-wk-label]");
  const fillWeeks = keep => {
    const y = +E("_wkYear").value, maxW = weekNum(new Date(y, 11, 31)), lastW = y === curY ? curW : maxW;
    E("_wkWeek").innerHTML = Array.from({length: lastW}, (_, i) => lastW - i)
      .map(w => `<option value="${w}">${w}${has(y, w) ? "  ✓ đã có" : ""}</option>`).join("");
    E("_wkWeek").value = keep && +keep <= lastW ? keep : lastW;
    updWk();
  };
  const updWk = () => { const y = +E("_wkYear").value, w = +E("_wkWeek").value;
    wkLabel.textContent = has(y, w) ? "Tuần này đã có nhật ký — nội dung mới sẽ được ghi nối tiếp" : "✓ = tuần đã có nhật ký"; };
  E("_wkYear").onchange = () => fillWeeks(E("_wkWeek").value); E("_wkWeek").onchange = updWk;
  fillWeeks(curW);
  const entries = weeklyEntries(r), wh = form.querySelector("[data-weeks]");
  const byYear = []; entries.forEach(e => { const g = byYear.find(x => x.y === e.y); g ? g.list.push(e) : byYear.push({y: e.y, list: [e]}); });
  wh.innerHTML = entries.length
    ? `<div class="wk-head">Nhật ký các tuần trước <span class="muted">(${entries.length} tuần · sửa trực tiếp, xoá hết chữ để xoá)</span></div>
       <div class="wk-list"><div class="wk-row wk-cols"><span>Tuần</span><span>Nội dung</span></div>${byYear.map(g => `<div class="wk-year">Năm ${g.y}</div>${g.list.map(e => `<div class="wk-row"><label for="${pre}wk${e.col}" title="Tuần ${e.w}/${e.y}">${e.w}</label><textarea class="autosize" id="${pre}wk${e.col}" name="_wkc_${e.col}" rows="1">${esc(e.text)}</textarea></div>`).join("")}`).join("")}</div>`
    : `<div class="wk-head muted">Chưa có nhật ký tuần nào.</div>`;
  form.addEventListener("input", e => { if (e.target.classList && e.target.classList.contains("autosize")) autosize(e.target); });
  const msg = form.querySelector("[data-msg]");
  const cancel = form.querySelector("[data-cancel]");
  cancel.onclick = () => onDone(false);
  if (!isEdit) cancel.textContent = "Xoá nội dung";
  if (!can.edit()) { // tài khoản chỉ xem
    [...form.elements].forEach(el => { if (el !== cancel) el.disabled = true; });
    form.querySelector("[data-save]").hidden = true; cancel.textContent = "Đóng";
    msg.textContent = "Tài khoản chỉ xem — không thể sửa dữ liệu";
  }
  if (isEdit && can.del()) {
    const host = form.querySelector("[data-del]");
    const buildDel = () => {
      host.innerHTML = `<button type="button" class="btn danger">Xoá KH</button>`;
      host.firstChild.onclick = () => {
        host.innerHTML = `<span class="confirm">Xoá khách hàng này?<button type="button" class="btn danger" data-y>Xoá</button><button type="button" class="btn" data-n>Không</button></span>`;
        host.querySelector("[data-n]").onclick = buildDel;
        host.querySelector("[data-y]").onclick = () => { if (!can.del()) return; const t = S.records.find(x => x.id === r.id); if (t) t._deleted = true; persist(); refreshAll(); toast("Đã xoá khách hàng (sẽ không có trong file xuất)"); onDone(true); };
      };
    };
    buildDel();
  }
  form.onsubmit = ev => {
    ev.preventDefault(); msg.textContent = "";
    if (!can.edit()) { msg.textContent = "Tài khoản của bạn không có quyền sửa dữ liệu"; return; }
    const miss = [...form.elements].filter(el => el.required && !String(el.value).trim());
    if (miss.length) { msg.textContent = "Còn thiếu: " + miss.map(el => form.querySelector(`label[for=${el.id}]`).textContent.replace("*", "").trim()).join(", "); miss[0].focus(); return; }
    const out = isEdit ? S.records.find(x => x.id === r.id) : {id: newId(), _new: true};
    for (const el of form.elements) {
      if (!el.name || el.name.startsWith("_")) continue;
      const v = String(el.value).trim(); if (v) out[el.name] = v; else delete out[el.name];
    }
    const d = new Date(out.ngay + "T00:00:00");
    if (!isNaN(d)) { out.nam = String(d.getFullYear()); out.thang = String(d.getMonth() + 1); out.tuan = String(weekNum(d)); out._namDerived = false; }
    if (!out.kv && kvOf(out.tinh)) out.kv = kvOf(out.tinh);
    out.updatedAt = new Date().toISOString();
    if (!isEdit) S.records.push(out);
    // 1) sửa / xoá nhật ký các tuần đã có (làm trước khi có thể chèn cột mới)
    out._raw ||= [];
    for (const el of form.elements) {
      const m = el.name && el.name.match(/^_wkc_(\d+)$/); if (!m) continue;
      const c = +m[1], old = out._raw[c] == null ? "" : String(out._raw[c] instanceof Date ? dmy(isoOf(out._raw[c])) : out._raw[c]), v = el.value.trim();
      if (v !== old.trim()) { while (out._raw.length <= c) out._raw.push(null); out._raw[c] = v || null; out._wkEdited = true; }
    }
    // 2) thêm nhật ký cho tuần được chọn
    const wkNote = E("_wkNote").value.trim(), wy = +E("_wkYear").value, ww = +E("_wkWeek").value;
    if (wkNote && wy && ww) {
      const col = weekCol(wy, ww, true);
      while (out._raw.length <= col) out._raw.push(null);
      const line = wkNote, cur = out._raw[col];
      out._raw[col] = cur != null && String(cur).trim() ? String(cur).trim() + "\n" + line : line;
      out._wkEdited = true;
    }
    persist(); refreshAll(); if (view === "check") renderCheck();
    toast(isEdit ? "Đã lưu thay đổi" : "Đã thêm khách hàng " + (out.ten || out.sdt));
    onDone(true);
  };
  return form;
}
// Ô chữ tự giãn theo nội dung (gọi sau khi form đã gắn vào trang)
function autosize(el) { el.style.height = "auto"; el.style.height = el.scrollHeight + 2 + "px"; }
const autosizeAll = root => root.querySelectorAll("textarea.autosize").forEach(autosize);
function mountAddForm() {
  const host = $("#addFormHost"); if (host.firstChild) return;
  host.appendChild(buildForm(null, {isEdit: false, onDone: ok => { host.innerHTML = ""; mountAddForm(); if (ok) { resetListFilters(); page = 0; renderList(); show("list"); } }}));
}
function openEdit(id) {
  const r = all.find(x => x.id === id); if (!r) return;
  const host = $("#drawerHost");
  host.innerHTML = `<div class="drawer-bg"></div><aside class="drawer" role="dialog" aria-modal="true" aria-labelledby="dTitle"><div class="drawer-head"><h2 id="dTitle">${esc(r.ten || r.sdt || "Khách hàng")}</h2><span class="muted" style="font-size:12px">${r._new ? "Nhập trên web" : "Từ file Excel" + (r._row ? " · dòng " + r._row : "")}</span><button class="btn" data-x aria-label="Đóng">✕</button></div></aside>`;
  const close = () => { host.innerHTML = ""; document.removeEventListener("keydown", esck); };
  const esck = e => { if (e.key === "Escape") close(); };
  document.addEventListener("keydown", esck);
  host.querySelector(".drawer-bg").onclick = close; host.querySelector("[data-x]").onclick = close;
  host.querySelector(".drawer").appendChild(buildForm(r, {isEdit: true, onDone: close}));
  autosizeAll(host);
}

// ================= Chung =================
function toast(t) { const h = $("#toastHost"); h.innerHTML = `<div class="toast" role="status">${esc(t)}</div>`; clearTimeout(toast.t); toast.t = setTimeout(() => h.innerHTML = "", 3200); }
const tip = $("#tip");
document.addEventListener("mousemove", e => {
  const t = e.target.closest && e.target.closest("[data-tip]"); if (!t) { tip.hidden = true; return; }
  tip.textContent = t.dataset.tip; tip.hidden = false;
  const w = tip.offsetWidth, h = tip.offsetHeight; let x = e.clientX + 14, y = e.clientY - h - 10;
  if (x + w > innerWidth - 8) x = e.clientX - w - 14; if (y < 8) y = e.clientY + 16;
  tip.style.left = x + "px"; tip.style.top = y + "px";
});
function renderAll() { if (!S) return; renderOverview(); renderList(); if (view === "check") renderCheck(); }

// Xuất / bỏ chỉnh sửa
$("#btnExport").onclick = () => exportExcel();
async function exportPrev() { // xuất các chỉnh sửa của file dữ liệu cũ
  const ps = await loadEnc("prevState"), pfEnc = await idb.get("prevFile"); if (!ps) return;
  let pf = null; try { pf = pfEnc ? await decryptDataFile(pfEnc) : null; } catch (e) {}
  const keep = [S, all, origFile, issueMap]; S = ps; origFile = pf || null; all = S.records.filter(r => !r._deleted); computeIssues();
  try { await exportExcel("_ban-cu"); } finally { [S, all, origFile, issueMap] = keep; computeIssues(); }
}
function bindClear() {
  $("#btnClear").onclick = () => {
    const host = $("#clearHost");
    host.innerHTML = `<span class="inline-confirm">Bỏ ${fmt(editCount(S))} chỉnh sửa, quay về đúng file Excel?<button class="btn danger" id="cY">Bỏ chỉnh sửa</button><button class="btn" id="cN">Không</button></span>`;
    const restore = () => { host.innerHTML = `<button class="btn danger" id="btnClear">Bỏ chỉnh sửa</button>`; bindClear(); showData(!!S); };
    $("#cN").onclick = restore;
    $("#cY").onclick = async () => { if (!can.reset()) return; await idb.del("state"); restore(); await loadDataFile(); toast("Đã quay về dữ liệu gốc trong file Excel"); };
  };
}
bindClear();
window.addEventListener("beforeunload", () => { if (S && SESSION) { clearTimeout(saveT); saveEnc("state", S); } });

// Khởi động: nạp lại dữ liệu đã lưu trong trình duyệt
showData(false);
(async () => {
  if (!window.crypto || !crypto.subtle) { setStatus("Trình duyệt tắt chức năng mã hoá vì trang không chạy qua HTTPS. Hãy mở bằng <b>https://</b> (GitHub Pages) hoặc <b>http://localhost:8080</b>.", true); return; }
  if (await restoreSession()) afterLogin(); else showLogin();
})();
