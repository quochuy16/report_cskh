// POST   /api/record              {data}            → thêm KH (editor, admin)
// PUT    /api/record?id=123       {data, version}   → sửa KH (editor, admin) — chống ghi đè bằng version
// DELETE /api/record?id=123       {version}         → xoá KH (admin)
// GET    /api/record?id=123&history=1               → lịch sử chỉnh sửa của KH
import {q} from "./_lib/db.js";
import {api, send, fail, body, query} from "./_lib/http.js";
import {requireUser} from "./_lib/auth.js";
import {rowOut, activeMeta} from "./data.js";

const MAX_BYTES = 256 * 1024;
function clean(data) {
  if (!data || typeof data !== "object" || Array.isArray(data)) throw Object.assign(new Error("Dữ liệu không hợp lệ"), {status: 400});
  const d = {...data};
  for (const k of ["id", "version", "deleted", "updatedAt", "updatedBy", "createdAt"]) delete d[k];
  if (JSON.stringify(d).length > MAX_BYTES) throw Object.assign(new Error("Dữ liệu quá lớn"), {status: 413});
  return d;
}
// So sánh 2 bản ghi → {trường: [cũ, mới]} (nhật ký tuần so theo từng tuần)
function diff(a = {}, b = {}) {
  const out = {};
  const keys = new Set([...Object.keys(a), ...Object.keys(b)].filter(k => k !== "weekly" && k !== "extra"));
  for (const k of keys) if (JSON.stringify(a[k] ?? null) !== JSON.stringify(b[k] ?? null)) out[k] = [a[k] ?? null, b[k] ?? null];
  for (const grp of ["weekly", "extra"]) {
    const x = a[grp] || {}, y = b[grp] || {};
    for (const k of new Set([...Object.keys(x), ...Object.keys(y)])) if (JSON.stringify(x[k] ?? null) !== JSON.stringify(y[k] ?? null)) out[`${grp}.${k}`] = [x[k] ?? null, y[k] ?? null];
  }
  return out;
}

export default api(["GET", "POST", "PUT", "DELETE"], async (req, res) => {
  const qs = query(req), meta = await activeMeta();
  if (!meta) return fail(res, 409, "Chưa có dữ liệu. Admin cần nhập file Excel trước.");

  if (req.method === "GET") {
    await requireUser(req, "viewer");
    if (!qs.history) return fail(res, 400, "Thiếu tham số");
    const rows = await q("SELECT at, by_user, action, diff FROM changes WHERE customer_id = $1 ORDER BY at DESC LIMIT 200", [qs.id]);
    return send(res, 200, {history: rows});
  }

  if (req.method === "POST") {
    const u = await requireUser(req, "editor");
    const d = clean((await body(req)).data); d._new = true;
    const [r] = await q("INSERT INTO customers (import_id, data, updated_by) VALUES ($1, $2, $3) RETURNING *", [meta.importId, d, u.username]);
    await q("INSERT INTO changes (customer_id, by_user, action, diff) VALUES ($1, $2, 'create', $3)", [r.id, u.username, diff({}, d)]);
    return send(res, 201, {record: rowOut(r)});
  }

  const id = +qs.id; if (!id) return fail(res, 400, "Thiếu id");
  const b = await body(req);
  const [cur] = await q("SELECT * FROM customers WHERE id = $1 AND import_id = $2", [id, meta.importId]);
  if (!cur || cur.deleted) return fail(res, 404, "Không tìm thấy khách hàng (có thể đã bị xoá)");
  if (b.version != null && +b.version !== cur.version) return send(res, 409, {error: `Khách hàng này vừa được ${cur.updated_by || "người khác"} sửa. Đã tải lại bản mới nhất — hãy kiểm tra rồi sửa lại.`, record: rowOut(cur)});

  if (req.method === "PUT") {
    const u = await requireUser(req, "editor");
    const d = clean(b.data), ch = diff(cur.data, d);
    if (!Object.keys(ch).length) return send(res, 200, {record: rowOut(cur), unchanged: true});
    const [r] = await q("UPDATE customers SET data = $1, version = version + 1, updated_at = now(), updated_by = $2 WHERE id = $3 AND version = $4 RETURNING *", [d, u.username, id, cur.version]);
    if (!r) return fail(res, 409, "Khách hàng vừa được người khác sửa, hãy tải lại.");
    await q("INSERT INTO changes (customer_id, by_user, action, diff) VALUES ($1, $2, 'update', $3)", [id, u.username, ch]);
    return send(res, 200, {record: rowOut(r)});
  }

  // DELETE — xoá mềm (vẫn giữ trong DB để tra lịch sử)
  const u = await requireUser(req, "admin");
  const [r] = await q("UPDATE customers SET deleted = true, version = version + 1, updated_at = now(), updated_by = $1 WHERE id = $2 RETURNING *", [u.username, id]);
  await q("INSERT INTO changes (customer_id, by_user, action, diff) VALUES ($1, $2, 'delete', $3)", [id, u.username, {ten: [cur.data.ten ?? null, null], sdt: [cur.data.sdt ?? null, null]}]);
  send(res, 200, {record: rowOut(r)});
});
