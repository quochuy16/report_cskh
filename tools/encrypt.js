#!/usr/bin/env node
/* Mã hoá file dữ liệu CSKH để đưa lên GitHub an toàn.

   Cách dùng (chạy trong thư mục web):
     node tools/encrypt.js                  mã hoá data/Report_CSKH.xlsx -> data/Report_CSKH.enc + data/keys.json
     node tools/encrypt.js --new-passwords  tạo lại username/mật khẩu cho cả 3 tài khoản rồi mã hoá
     node tools/encrypt.js --show           in lại danh sách tài khoản hiện tại

   - Lần chạy đầu tiên tự tạo 3 tài khoản (admin / editor / viewer) với username và mật khẩu ngẫu nhiên,
     lưu trong secrets.local.json. FILE NÀY KHÔNG ĐƯỢC ĐƯA LÊN GITHUB (đã có trong .gitignore).
   - File Excel gốc (data/*.xlsx) cũng không đưa lên GitHub; chỉ đưa file .enc và keys.json.

   Thiết kế mã hoá:
   - Dữ liệu được mã hoá bằng AES-256-GCM với một khoá dữ liệu ngẫu nhiên (dataKey), giữ cố định giữa các lần cập nhật.
   - Với mỗi tài khoản: khoá bọc = PBKDF2-SHA256(mật khẩu, salt ngẫu nhiên, 600.000 vòng);
     {dataKey, role, username} được mã hoá bằng khoá bọc (AES-256-GCM) và lưu vào keys.json.
   - keys.json chỉ chứa mã băm của username (không lộ tên đăng nhập) và dữ liệu đã mã hoá. */
"use strict";
const fs = require("fs"), path = require("path"), crypto = require("crypto");

const ROOT = path.join(__dirname, "..");
const SRC = path.join(ROOT, "data", "Report_CSKH.xlsx");
const OUT = path.join(ROOT, "data", "Report_CSKH.enc");
const KEYS = path.join(ROOT, "data", "keys.json");
const SECRETS = path.join(ROOT, "secrets.local.json");
const ITER = 600000;
const ROLES = [["admin", "Quản trị – toàn quyền"], ["editor", "Biên tập – thêm/sửa, xuất Excel"], ["viewer", "Chỉ xem"]];

const b64 = b => Buffer.from(b).toString("base64");
const rand = n => crypto.randomBytes(n);
function randomString(len, alphabet) { // chọn ký tự ngẫu nhiên không lệch (rejection sampling)
  let out = ""; const max = 256 - (256 % alphabet.length);
  while (out.length < len) for (const b of rand(len * 2)) { if (b < max && out.length < len) out += alphabet[b % alphabet.length]; }
  return out;
}
const LOWER = "abcdefghjkmnpqrstuvwxyz", UPPER = "ABCDEFGHJKLMNPQRSTUVWXYZ", DIGIT = "23456789", SYM = "!@#$%^&*-_=+?";
function makePassword() { // 24 ký tự, chắc chắn có đủ 4 loại ký tự
  for (;;) { const p = randomString(24, LOWER + UPPER + DIGIT + SYM); if (/[a-z]/.test(p) && /[A-Z]/.test(p) && /\d/.test(p) && /[^A-Za-z0-9]/.test(p)) return p; }
}
const makeUsername = role => `${role}.${randomString(10, LOWER + DIGIT)}`;
const userId = u => crypto.createHash("sha256").update("cskh-user:" + u.trim().toLowerCase()).digest("hex");
function gcmEncrypt(key, plain) { // trả về iv + (ciphertext||tag) giống định dạng WebCrypto
  const iv = rand(12), c = crypto.createCipheriv("aes-256-gcm", key, iv);
  const ct = Buffer.concat([c.update(plain), c.final(), c.getAuthTag()]);
  return {iv, ct};
}

function loadSecrets(renew) {
  let s = null;
  if (fs.existsSync(SECRETS)) s = JSON.parse(fs.readFileSync(SECRETS, "utf8"));
  if (!s) s = {dataKey: b64(rand(32)), users: []};
  if (renew || !s.users.length) {
    s.users = ROLES.map(([role, desc]) => ({role, desc, username: makeUsername(role), password: makePassword()}));
    s.createdAt = new Date().toISOString();
    fs.writeFileSync(SECRETS, JSON.stringify(s, null, 2));
    console.log("Đã tạo tài khoản mới và lưu vào secrets.local.json\n");
    return {s, created: true};
  }
  return {s, created: false};
}
function printUsers(s) {
  console.log("Tài khoản đăng nhập:");
  for (const u of s.users) console.log(`  [${u.role.padEnd(6)}] ${u.desc}\n           username: ${u.username}\n           password: ${u.password}`);
  console.log("");
}

const args = process.argv.slice(2);
const {s, created} = loadSecrets(args.includes("--new-passwords"));
if (args.includes("--show")) { printUsers(s); process.exit(0); }
if (!fs.existsSync(SRC)) { console.error("Không tìm thấy " + path.relative(ROOT, SRC) + " — hãy chép file Excel vào đó (đúng tên)."); process.exit(1); }

const dataKey = Buffer.from(s.dataKey, "base64");
const xlsx = fs.readFileSync(SRC);
const {iv, ct} = gcmEncrypt(dataKey, xlsx);
fs.writeFileSync(OUT, Buffer.concat([Buffer.from("CSKH1"), iv, ct]));

const users = s.users.map(u => {
  const salt = rand(16);
  const wrap = crypto.pbkdf2Sync(u.password, salt, ITER, 32, "sha256");
  const e = gcmEncrypt(wrap, Buffer.from(JSON.stringify({dataKey: s.dataKey, role: u.role, username: u.username})));
  return {id: userId(u.username), salt: b64(salt), iv: b64(e.iv), ct: b64(e.ct)};
});
fs.writeFileSync(KEYS, JSON.stringify({v: 1, kdf: "PBKDF2-SHA256", iter: ITER, updatedAt: new Date().toISOString(), users}, null, 1));

console.log(`Đã mã hoá ${path.relative(ROOT, SRC)} (${(xlsx.length / 1024).toFixed(0)} KB)`);
console.log(`  -> ${path.relative(ROOT, OUT)}\n  -> ${path.relative(ROOT, KEYS)}\n`);
if (created || args.includes("--show-after")) printUsers(s);
console.log("Tiếp theo: git add data/Report_CSKH.enc data/keys.json && git commit -m \"Cập nhật dữ liệu\" && git push");
