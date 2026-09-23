// Nhập file Excel vào database (chỉ admin). Trình duyệt đọc Excel rồi gửi lên theo từng đợt:
//   POST /api/import {action:"start",  meta, template}  → {importId}
//   POST /api/import {action:"batch",  importId, rows:[...]}
//   POST /api/import {action:"finish", importId, count}   → dữ liệu mới có hiệu lực, dữ liệu cũ bị thay thế
// Trong lúc nhập, mọi người vẫn xem dữ liệu cũ; chỉ khi "finish" thành công mới chuyển sang dữ liệu mới.
import {q} from "./_lib/db.js";
import {api, send, fail, body} from "./_lib/http.js";
import {requireUser} from "./_lib/auth.js";

export default api(["POST"], async (req, res) => {
  const u = await requireUser(req, "admin");
  const b = await body(req);

  if (b.action === "start") {
    if (!b.meta || !Array.isArray(b.meta.header)) return fail(res, 400, "Thiếu thông tin cấu trúc file");
    const [{id}] = await q("SELECT nextval('import_seq')::int AS id");
    await q("INSERT INTO meta (key, value) VALUES ($1, $2) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = now()",
      ["pending_" + id, {...b.meta, importId: id, importedBy: u.username, importedAt: new Date().toISOString()}]);
    if (b.template) await q("INSERT INTO files (name, content) VALUES ($1, $2) ON CONFLICT (name) DO UPDATE SET content = EXCLUDED.content, updated_at = now()", ["template_" + id, b.template]);
    return send(res, 200, {importId: id});
  }

  const id = +b.importId;
  const [pending] = await q("SELECT value FROM meta WHERE key = $1", ["pending_" + id]);
  if (!pending) return fail(res, 400, "Phiên nhập không tồn tại hoặc đã kết thúc");

  if (b.action === "batch") {
    if (!Array.isArray(b.rows) || b.rows.length > 1000) return fail(res, 400, "Đợt dữ liệu không hợp lệ");
    await q(`INSERT INTO customers (import_id, row_no, data, updated_by, updated_at, created_at)
             SELECT $1, (x.val->>'_row')::int, x.val, $3, now(), now() FROM jsonb_array_elements($2::jsonb) AS x(val)`, [id, JSON.stringify(b.rows), u.username]);
    return send(res, 200, {ok: true});
  }

  if (b.action === "finish") {
    const [{n}] = await q("SELECT count(*)::int AS n FROM customers WHERE import_id = $1", [id]);
    if (n !== +b.count) return fail(res, 400, `Số dòng không khớp (${n}/${b.count}). Hãy nhập lại.`);
    const old = (await q("SELECT value FROM meta WHERE key = 'active'"))[0];
    await q("INSERT INTO meta (key, value) VALUES ('active', $1) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = now()", [pending.value]);
    await q("DELETE FROM meta WHERE key = $1", ["pending_" + id]);
    // Dọn dữ liệu của lần nhập cũ và các lần nhập dở dang
    await q("DELETE FROM customers WHERE import_id <> $1", [id]);
    await q("DELETE FROM files WHERE name LIKE 'template_%' AND name <> $1", ["template_" + id]);
    await q("DELETE FROM meta WHERE key LIKE 'pending_%'");
    await q("DELETE FROM changes WHERE customer_id IS NOT NULL AND customer_id NOT IN (SELECT id FROM customers)");
    await q("INSERT INTO changes (customer_id, by_user, action, diff) VALUES (NULL, $1, 'import', $2)",
      [u.username, {file: pending.value.fileName, rows: n, replacedImport: old ? old.value.importId : null}]);
    return send(res, 200, {ok: true, rows: n});
  }
  fail(res, 400, "action không hợp lệ");
});
