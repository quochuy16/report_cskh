// Chạy web + API trên máy để thử (không cần Vercel):  npm install  →  node tools/server.js  → http://localhost:8080
// - Không có DATABASE_URL: dùng database cục bộ PGlite (lưu trong thư mục ~/.cskh-localdb)
// - Muốn dùng thẳng database Neon: đặt biến môi trường DATABASE_URL trước khi chạy
// - Tài khoản admin đầu tiên lấy từ ADMIN_USERNAME / ADMIN_PASSWORD (mặc định khi chạy máy: admin / admin-local-123)
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import {fileURLToPath, pathToFileURL} from "node:url";

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const port = process.env.PORT || 8080;
process.env.ADMIN_USERNAME ||= "admin";
process.env.ADMIN_PASSWORD ||= "admin-local-123";
process.env.PGLITE_DIR ||= path.join(os.homedir(), ".cskh-localdb"); // lưu ở thư mục người dùng (một số ổ đĩa/USB không hỗ trợ)

const types = {".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".json": "application/json; charset=utf-8", ".svg": "image/svg+xml", ".png": "image/png", ".ico": "image/x-icon"};
const handlers = {};
async function apiHandler(name) {
  if (!/^[a-z]+$/.test(name)) return null;
  const f = path.join(root, "api", name + ".js"); if (!fs.existsSync(f)) return null;
  return handlers[name] ||= (await import(pathToFileURL(f).href)).default;
}

http.createServer(async (req, res) => {
  const url = new URL(req.url, "http://localhost");
  if (url.pathname.startsWith("/api/")) {
    const h = await apiHandler(url.pathname.slice(5).replace(/\/$/, ""));
    if (!h) { res.writeHead(404); return res.end("Not found"); }
    req.query = Object.fromEntries(url.searchParams);
    return h(req, res);
  }
  let p = decodeURIComponent(url.pathname); if (p === "/") p = "/index.html";
  const pub = path.join(root, "public"), f = path.join(pub, path.normalize(p)); // chỉ phục vụ file trong public/ (giống Vercel)
  if (!f.startsWith(pub)) { res.writeHead(404); return res.end("Not found"); }
  fs.stat(f, (e, st) => {
    if (e || !st.isFile()) { res.writeHead(404); return res.end("Not found"); }
    res.writeHead(200, {"Content-Type": types[path.extname(f)] || "application/octet-stream", "Cache-Control": "no-cache"});
    fs.createReadStream(f).pipe(res);
  });
}).listen(port, () => console.log(`Sổ CSKH đang chạy tại http://localhost:${port}  (database: ${process.env.DATABASE_URL ? "Neon" : "PGlite cục bộ"})`));
