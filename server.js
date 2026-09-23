// Server tĩnh đơn giản (tuỳ chọn) — chạy: node server.js  rồi mở http://localhost:8080
const http = require("http"), fs = require("fs"), path = require("path");
const root = __dirname, port = process.env.PORT || 8080;
const types = {".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".xlsx": "application/octet-stream", ".md": "text/plain; charset=utf-8", ".json": "application/json; charset=utf-8", ".enc": "application/octet-stream"};
http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split("?")[0]); if (p === "/") p = "/index.html";
  const f = path.join(root, path.normalize(p));
  if (!f.startsWith(root)) { res.writeHead(403); return res.end(); }
  // không phục vụ file bí mật / dữ liệu chưa mã hoá
  if (/secrets\.local\.json$|\.xlsx$|[\\/]\.git([\\/]|$)/i.test(f)) { res.writeHead(404); return res.end("Not found"); }
  fs.stat(f, (e, st) => {
    if (e || !st.isFile()) { res.writeHead(404); return res.end("Not found"); }
    res.writeHead(200, {"Content-Type": types[path.extname(f)] || "application/octet-stream", "Last-Modified": st.mtime.toUTCString(), "Cache-Control": "no-cache"});
    fs.createReadStream(f).pipe(res);
  });
}).listen(port, () => console.log("Sổ CSKH đang chạy tại http://localhost:" + port));
