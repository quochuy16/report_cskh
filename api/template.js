// GET /api/template → file Excel mẫu (các sheet gốc, sheet danh sách để trống) dùng khi xuất Excel
import {q} from "./_lib/db.js";
import {api, send, fail} from "./_lib/http.js";
import {requireUser} from "./_lib/auth.js";
import {activeMeta} from "./data.js";

export default api(["GET"], async (req, res) => {
  await requireUser(req, "editor");
  const meta = await activeMeta(); if (!meta) return fail(res, 404, "Chưa có dữ liệu");
  const [f] = await q("SELECT content FROM files WHERE name = $1", ["template_" + meta.importId]);
  send(res, 200, {template: f ? f.content : null});
});
