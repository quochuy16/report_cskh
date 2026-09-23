// GET /api/me  → người dùng đang đăng nhập
import {api, send, fail} from "./_lib/http.js";
import {currentUser} from "./_lib/auth.js";

export default api(["GET"], async (req, res) => {
  const u = await currentUser(req);
  if (!u) return fail(res, 401, "Chưa đăng nhập");
  send(res, 200, {user: {id: u.id, username: u.username, role: u.role}});
});
