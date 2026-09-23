// Tiện ích HTTP dùng chung cho các API (chạy được cả trên Vercel lẫn server chạy máy tools/server.js)
export function send(res, status, obj) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");
  res.end(JSON.stringify(obj));
}
export const fail = (res, status, message) => send(res, status, {error: message});

export async function body(req) {
  if (req.body !== undefined && req.body !== null && req.body !== "") return typeof req.body === "string" ? JSON.parse(req.body) : req.body;
  const chunks = []; for await (const c of req) chunks.push(c);
  const s = Buffer.concat(chunks).toString("utf8");
  return s ? JSON.parse(s) : {};
}

export function query(req) {
  if (req.query) return req.query;
  const u = new URL(req.url, "http://x"); return Object.fromEntries(u.searchParams);
}

// Chống CSRF: mọi request thay đổi dữ liệu phải có header X-CSKH (trình duyệt không tự gửi header này từ trang khác)
export function sameOrigin(req) { return req.method === "GET" || req.headers["x-cskh"] === "1"; }

/** Bọc handler: bắt lỗi, kiểm tra method + CSRF */
export function api(methods, fn) {
  return async (req, res) => {
    try {
      if (!methods.includes(req.method)) return fail(res, 405, "Method not allowed");
      if (!sameOrigin(req)) return fail(res, 403, "Thiếu header bảo mật");
      await fn(req, res);
    } catch (e) {
      console.error(e);
      if (!res.headersSent) fail(res, e.status || 500, e.status ? e.message : "Lỗi máy chủ: " + (e.message || e));
    }
  };
}
export class HttpError extends Error { constructor(status, message) { super(message); this.status = status; } }
