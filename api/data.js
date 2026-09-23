// GET /api/data?offset=0&limit=500        → trang dữ liệu khách hàng (trang đầu kèm meta)
// GET /api/data?since=<ISO time>          → các dòng thay đổi từ thời điểm đó (để đồng bộ giữa nhiều người)
import {q} from "./_lib/db.js";
import {api, send, query} from "./_lib/http.js";
import {requireUser} from "./_lib/auth.js";

export const rowOut = r => ({...r.data, id: String(r.id), version: r.version, deleted: r.deleted, updatedAt: r.updated_at, updatedBy: r.updated_by, createdAt: r.created_at});

export async function activeMeta() {
  const [m] = await q("SELECT value FROM meta WHERE key = 'active'");
  return m ? m.value : null;
}

export default api(["GET"], async (req, res) => {
  await requireUser(req, "viewer");
  const qs = query(req), meta = await activeMeta();
  const now = (await q("SELECT now() AS t"))[0].t;
  if (!meta) return send(res, 200, {meta: null, records: [], total: 0, serverTime: now});

  if (qs.since) {
    const rows = await q("SELECT * FROM customers WHERE import_id = $1 AND updated_at > $2 ORDER BY updated_at LIMIT 2000", [meta.importId, qs.since]);
    return send(res, 200, {records: rows.map(rowOut), serverTime: now});
  }
  const limit = Math.min(1000, Math.max(1, +qs.limit || 500)), offset = Math.max(0, +qs.offset || 0);
  const rows = await q("SELECT * FROM customers WHERE import_id = $1 AND NOT deleted ORDER BY id LIMIT $2 OFFSET $3", [meta.importId, limit, offset]);
  const [{total}] = await q("SELECT count(*)::int AS total FROM customers WHERE import_id = $1 AND NOT deleted", [meta.importId]);
  send(res, 200, {meta: offset === 0 ? meta : undefined, records: rows.map(rowOut), total, next: offset + rows.length < total ? offset + rows.length : null, serverTime: now});
});
