#!/usr/bin/env node
/* Cập nhật dữ liệu lên web chỉ bằng 1 lệnh:
     node tools/update.js "D:\duong-dan\file-moi.xlsx"   chép file mới vào data/, mã hoá, commit, push
     node tools/update.js                                 mã hoá lại data/Report_CSKH.xlsx hiện có, commit, push
     thêm --no-git                                        chỉ chép + mã hoá, không commit/push
   Trên Windows: kéo thả file Excel mới vào capnhat.bat là xong. */
"use strict";
const fs = require("fs"), path = require("path"), {execFileSync, spawnSync} = require("child_process");

const ROOT = path.join(__dirname, "..");
const DEST = path.join(ROOT, "data", "Report_CSKH.xlsx");
const args = process.argv.slice(2), noGit = args.includes("--no-git");
const src = args.find(a => !a.startsWith("--"));
const die = m => { console.error("\n✖ " + m + "\n"); process.exit(1); };
const git = (...a) => execFileSync("git", a, {cwd: ROOT, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"]}).trim();

// 1. Chép file Excel mới (nếu có)
if (src) {
  const p = path.resolve(src.replace(/^"|"$/g, ""));
  if (!fs.existsSync(p)) die("Không tìm thấy file: " + p);
  if (!/\.(xlsx|xlsm)$/i.test(p)) die("File phải là Excel .xlsx: " + p);
  if (path.resolve(p) !== path.resolve(DEST)) { fs.copyFileSync(p, DEST); console.log("✔ Đã chép " + path.basename(p) + " → data/Report_CSKH.xlsx"); }
} else if (!fs.existsSync(DEST)) die("Chưa có data/Report_CSKH.xlsx. Hãy kéo thả file Excel vào capnhat.bat hoặc chạy: node tools/update.js <file.xlsx>");

// 2. Mã hoá
const r = spawnSync(process.execPath, [path.join(__dirname, "encrypt.js")], {cwd: ROOT, stdio: "inherit"});
if (r.status !== 0) die("Mã hoá thất bại");
if (noGit) { console.log("\n✔ Xong (bỏ qua bước git)."); process.exit(0); }

// 3. Kiểm tra an toàn trước khi đẩy lên GitHub
try { git("rev-parse", "--is-inside-work-tree"); } catch (e) { die("Thư mục này chưa là git repo. Làm theo phần “Đẩy lên GitHub lần đầu” trong README trước."); }
const tracked = git("ls-files").split("\n");
const leaked = tracked.filter(f => /\.xlsx?$|\.xlsm$|secrets\.local\.json$/i.test(f));
if (leaked.length) die("Repo đang chứa file không được phép đưa lên (dữ liệu gốc / mật khẩu): " + leaked.join(", ") + "\nHãy gỡ khỏi git: git rm --cached <file>");
for (const f of ["data/Report_CSKH.xlsx", "secrets.local.json"]) {
  try { git("check-ignore", "-q", f); } catch (e) { die(f + " chưa được chặn trong .gitignore — dừng lại để tránh lộ dữ liệu."); }
}

// 4. Commit + push đúng 2 file đã mã hoá
git("add", "data/Report_CSKH.enc", "data/keys.json");
const staged = git("diff", "--cached", "--name-only").split("\n").filter(Boolean);
const bad = staged.filter(f => !["data/Report_CSKH.enc", "data/keys.json"].includes(f));
if (bad.length) die("Có file khác đang chờ commit (" + bad.join(", ") + "). Hãy commit/huỷ chúng trước rồi chạy lại.");
if (!staged.length) { console.log("\n✔ Dữ liệu không thay đổi, không cần đẩy lên."); process.exit(0); }
const now = new Date(), p2 = n => String(n).padStart(2, "0");
const msg = `Cập nhật dữ liệu CSKH ${p2(now.getDate())}/${p2(now.getMonth() + 1)}/${now.getFullYear()} ${p2(now.getHours())}:${p2(now.getMinutes())}`;
git("commit", "-m", msg);
console.log("✔ Đã commit: " + msg);
let upstream = true; try { git("rev-parse", "--abbrev-ref", "@{u}"); } catch (e) { upstream = false; }
const branch = git("rev-parse", "--abbrev-ref", "HEAD");
const push = spawnSync("git", upstream ? ["push"] : ["push", "-u", "origin", branch], {cwd: ROOT, stdio: "inherit"});
if (push.status !== 0) die("Push thất bại — kiểm tra mạng / đăng nhập GitHub rồi chạy: git push");
console.log("\n✔ Đã đẩy lên GitHub. Web sẽ cập nhật sau khoảng 1–2 phút (tải lại trang để xem).");
