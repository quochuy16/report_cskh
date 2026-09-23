// GET /api/stats → dung lượng database đang dùng + số liệu hệ thống (chỉ admin)
import {q} from "./_lib/db.js";
import {api, send, body} from "./_lib/http.js";
import {requireUser} from "./_lib/auth.js";

const LIMIT_BYTES = (+process.env.DB_LIMIT_MB || 512) * 1024 * 1024; // gói Neon Free: 0,5 GB

export default api(["GET", "POST"], async (req, res) => {
  const u = await requireUser(req, "admin");
  if (req.method === "POST") { // dọn lịch sử chỉnh sửa cũ hơn N tháng
    const {months} = await body(req);
    const m = Math.max(3, Math.min(120, +months || 24));
    const del = await q(`DELETE FROM changes WHERE at < now() - interval '${m} months' RETURNING id`);
    await q("INSERT INTO changes (customer_id, by_user, action, diff) VALUES (NULL, $1, 'prune', $2)", [u.username, {months: m, removed: del.length}]);
    return send(res, 200, {removed: del.length});
  }
  const size = async t => { try { return +(await q(`SELECT pg_total_relation_size('${t}') AS b`))[0].b; } catch (e) { return 0; } };
  let dbBytes = 0; try { dbBytes = +(await q("SELECT pg_database_size(current_database()) AS b"))[0].b; } catch (e) {}
  const [{customers}] = await q("SELECT count(*)::int AS customers FROM customers WHERE NOT deleted AND import_id = (SELECT (value->>'importId')::int FROM meta WHERE key = 'active')");
  const [{changes}] = await q("SELECT count(*)::int AS changes FROM changes");
  const [{oldest}] = await q("SELECT min(at) AS oldest FROM changes");
  send(res, 200, {
    dbBytes, limitBytes: LIMIT_BYTES, customers, changes, oldestChange: oldest,
    tables: {customers: await size("customers"), changes: await size("changes"), files: await size("files"), users: await size("users")},
  });
});
