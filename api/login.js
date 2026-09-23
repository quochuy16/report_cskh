// POST /api/login  {username, password}
import {q} from "./_lib/db.js";
import {api, body, send, fail} from "./_lib/http.js";
import {hashPassword, verifyPassword, setSession} from "./_lib/auth.js";

const MAX_FAILS = 8, WINDOW_MIN = 15;

export default api(["POST"], async (req, res) => {
  const {username = "", password = ""} = await body(req);
  const u = String(username).trim().toLowerCase(), p = String(password);
  if (!u || !p) return fail(res, 400, "Nhập tên đăng nhập và mật khẩu");

  // Chặn dò mật khẩu: quá MAX_FAILS lần sai trong WINDOW_MIN phút
  const [{n}] = await q(`SELECT count(*)::int AS n FROM login_fail WHERE username = $1 AND at > now() - interval '${WINDOW_MIN} minutes'`, [u]);
  if (n >= MAX_FAILS) return fail(res, 429, `Sai quá nhiều lần. Thử lại sau ${WINDOW_MIN} phút.`);

  // Lần đầu (chưa có tài khoản nào): tạo admin từ biến môi trường ADMIN_USERNAME / ADMIN_PASSWORD
  const [{c}] = await q("SELECT count(*)::int AS c FROM users");
  if (c === 0) {
    const au = String(process.env.ADMIN_USERNAME || "").trim().toLowerCase(), ap = process.env.ADMIN_PASSWORD || "";
    if (!au || ap.length < 12) return fail(res, 503, "Hệ thống chưa có tài khoản. Hãy đặt ADMIN_USERNAME và ADMIN_PASSWORD (≥ 12 ký tự) trong Vercel → Settings → Environment Variables.");
    if (u === au && p === ap) await q("INSERT INTO users (username, pass_hash, role) VALUES ($1, $2, 'admin') ON CONFLICT (username) DO NOTHING", [au, hashPassword(ap)]);
  }

  const [user] = await q("SELECT id, username, role, active, pv, pass_hash FROM users WHERE username = $1", [u]);
  const ok = user && user.active && verifyPassword(p, user.pass_hash);
  if (!ok) {
    await q("INSERT INTO login_fail (username) VALUES ($1)", [u]);
    await q("DELETE FROM login_fail WHERE at < now() - interval '1 day'");
    return fail(res, 401, "Sai tên đăng nhập hoặc mật khẩu");
  }
  await q("DELETE FROM login_fail WHERE username = $1", [u]);
  await q("UPDATE users SET last_login = now() WHERE id = $1", [user.id]);
  setSession(req, res, user);
  send(res, 200, {user: {id: user.id, username: user.username, role: user.role}});
});
