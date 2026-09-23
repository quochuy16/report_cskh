// Quản lý tài khoản (chỉ admin)
//   GET  /api/users                              → danh sách
//   POST /api/users {username, role}             → tạo tài khoản, trả mật khẩu ngẫu nhiên (hiện 1 lần)
//   PUT  /api/users {id, role?, active?, reset?} → đổi quyền / khoá / đặt lại mật khẩu
import {q} from "./_lib/db.js";
import {api, send, fail, body} from "./_lib/http.js";
import {requireUser, hashPassword, randomPassword, ROLES} from "./_lib/auth.js";

export default api(["GET", "POST", "PUT"], async (req, res) => {
  const me = await requireUser(req, "admin");
  if (req.method === "GET") {
    const users = await q("SELECT id, username, role, active, created_at, last_login FROM users ORDER BY role, username");
    return send(res, 200, {users});
  }
  const b = await body(req);
  if (req.method === "POST") {
    const username = String(b.username || "").trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,40}$/.test(username)) return fail(res, 400, "Tên đăng nhập 3–40 ký tự: chữ thường, số, dấu . _ -");
    if (!ROLES.includes(b.role)) return fail(res, 400, "Quyền không hợp lệ");
    const password = randomPassword();
    const r = await q("INSERT INTO users (username, pass_hash, role) VALUES ($1, $2, $3) ON CONFLICT (username) DO NOTHING RETURNING id", [username, hashPassword(password), b.role]);
    if (!r.length) return fail(res, 409, "Tên đăng nhập đã tồn tại");
    return send(res, 201, {username, password});
  }
  // PUT
  const [u] = await q("SELECT id, username, role, active FROM users WHERE id = $1", [+b.id]);
  if (!u) return fail(res, 404, "Không tìm thấy tài khoản");
  const role = b.role ?? u.role, active = b.active ?? u.active;
  if (!ROLES.includes(role)) return fail(res, 400, "Quyền không hợp lệ");
  if (u.id === me.id && (role !== "admin" || !active)) return fail(res, 400, "Không thể tự hạ quyền hoặc khoá chính mình");
  if (u.role === "admin" && (role !== "admin" || !active)) {
    const [{n}] = await q("SELECT count(*)::int AS n FROM users WHERE role = 'admin' AND active AND id <> $1", [u.id]);
    if (!n) return fail(res, 400, "Phải còn ít nhất một admin đang hoạt động");
  }
  let password = null;
  if (b.reset) { password = randomPassword(); await q("UPDATE users SET pass_hash = $1, pv = pv + 1 WHERE id = $2", [hashPassword(password), u.id]); }
  // đổi quyền / khoá cũng tăng pv để phiên đang mở của người đó phải đăng nhập lại
  if (role !== u.role || active !== u.active) await q("UPDATE users SET role = $1, active = $2, pv = pv + 1 WHERE id = $3", [role, active, u.id]);
  send(res, 200, {ok: true, username: u.username, password});
});
