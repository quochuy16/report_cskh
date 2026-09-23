// Đăng nhập, phiên làm việc (cookie JWT ký HMAC) và phân quyền — kiểm tra ở phía server.
import crypto from "node:crypto";
import {q} from "./db.js";
import {HttpError} from "./http.js";

const COOKIE = "cskh_session";
const TTL_HOURS = 12;
export const ROLES = ["admin", "editor", "viewer"];
const RANK = {viewer: 1, editor: 2, admin: 3};

function secret() {
  const s = process.env.AUTH_SECRET;
  if (s && s.length >= 32) return s;
  if (!process.env.VERCEL) return "dev-only-secret-khong-dung-tren-that-0123456789";
  throw new HttpError(500, "Chưa cấu hình AUTH_SECRET (tối thiểu 32 ký tự) trong Vercel → Settings → Environment Variables");
}

// ---- Mật khẩu: scrypt + salt ngẫu nhiên
export function hashPassword(pw) {
  const salt = crypto.randomBytes(16);
  const h = crypto.scryptSync(pw, salt, 64, {N: 16384, r: 8, p: 1});
  return `scrypt$${salt.toString("base64")}$${h.toString("base64")}`;
}
export function verifyPassword(pw, stored) {
  const [alg, s, h] = String(stored).split("$"); if (alg !== "scrypt") return false;
  const got = crypto.scryptSync(pw, Buffer.from(s, "base64"), 64, {N: 16384, r: 8, p: 1});
  const want = Buffer.from(h, "base64");
  return got.length === want.length && crypto.timingSafeEqual(got, want);
}
export function randomPassword(len = 20) {
  const A = "abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%^&*-_=+?";
  for (;;) {
    let p = ""; const max = 256 - (256 % A.length);
    while (p.length < len) for (const b of crypto.randomBytes(len * 2)) if (b < max && p.length < len) p += A[b % A.length];
    if (/[a-z]/.test(p) && /[A-Z]/.test(p) && /\d/.test(p) && /[^A-Za-z0-9]/.test(p)) return p;
  }
}

// ---- Token phiên
const b64u = b => Buffer.from(b).toString("base64url");
function sign(payload) {
  const h = b64u(JSON.stringify({alg: "HS256", typ: "JWT"})), p = b64u(JSON.stringify(payload));
  const sig = crypto.createHmac("sha256", secret()).update(h + "." + p).digest("base64url");
  return `${h}.${p}.${sig}`;
}
function verify(tok) {
  const [h, p, sig] = String(tok || "").split("."); if (!sig) return null;
  const want = crypto.createHmac("sha256", secret()).update(h + "." + p).digest("base64url");
  if (want.length !== sig.length || !crypto.timingSafeEqual(Buffer.from(want), Buffer.from(sig))) return null;
  const data = JSON.parse(Buffer.from(p, "base64url").toString("utf8"));
  return data.exp > Date.now() / 1000 ? data : null;
}
function cookieOf(req) {
  const m = String(req.headers.cookie || "").match(new RegExp("(?:^|;\\s*)" + COOKIE + "=([^;]+)"));
  return m ? decodeURIComponent(m[1]) : null;
}
const secure = req => process.env.VERCEL || String(req.headers["x-forwarded-proto"] || "").includes("https");
export function setSession(req, res, user) {
  const tok = sign({uid: user.id, pv: user.pv, exp: Math.floor(Date.now() / 1000) + TTL_HOURS * 3600});
  res.setHeader("Set-Cookie", `${COOKIE}=${tok}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${TTL_HOURS * 3600}${secure(req) ? "; Secure" : ""}`);
}
export function clearSession(req, res) {
  res.setHeader("Set-Cookie", `${COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0${secure(req) ? "; Secure" : ""}`);
}

/** Người dùng hiện tại (hoặc null). Kiểm tra lại trong DB: tài khoản còn hoạt động và chưa đổi mật khẩu. */
export async function currentUser(req) {
  const t = verify(cookieOf(req)); if (!t) return null;
  const [u] = await q("SELECT id, username, role, active, pv FROM users WHERE id = $1", [t.uid]);
  return u && u.active && u.pv === t.pv ? {id: u.id, username: u.username, role: u.role, pv: u.pv} : null;
}
/** Bắt buộc đăng nhập với quyền tối thiểu minRole */
export async function requireUser(req, minRole = "viewer") {
  const u = await currentUser(req);
  if (!u) throw new HttpError(401, "Chưa đăng nhập hoặc phiên đã hết hạn");
  if (RANK[u.role] < RANK[minRole]) throw new HttpError(403, "Tài khoản của bạn không có quyền thực hiện thao tác này");
  return u;
}
