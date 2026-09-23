// Kết nối database.
// - Trên Vercel: dùng Postgres Neon qua biến môi trường DATABASE_URL (tự có khi gắn Neon vào project).
// - Trên máy (không có DATABASE_URL): dùng PGlite — Postgres chạy ngay trong Node, lưu ở thư mục .localdb/
let _q = null, _ready = null;

async function driver() {
  if (_q) return _q;
  const url = process.env.DATABASE_URL || process.env.POSTGRES_URL;
  if (url) {
    const {neon} = await import("@neondatabase/serverless");
    const sql = neon(url);
    _q = (text, params = []) => sql.query(text, params);
  } else {
    if (process.env.VERCEL) throw Object.assign(new Error("Chưa gắn database: vào Vercel → Storage → tạo/kết nối Neon Postgres cho project này (sẽ tự có biến DATABASE_URL), rồi Redeploy."), {status: 500});
    const {PGlite} = await import("@electric-sql/pglite");
    const db = globalThis.__cskhPglite ||= new PGlite(process.env.PGLITE_DIR || ".localdb");
    _q = async (text, params = []) => (await db.query(text, params)).rows;
  }
  return _q;
}

const SCHEMA = [
  `CREATE TABLE IF NOT EXISTS users (
     id serial PRIMARY KEY,
     username text UNIQUE NOT NULL,
     pass_hash text NOT NULL,
     role text NOT NULL CHECK (role IN ('admin','editor','viewer')),
     active boolean NOT NULL DEFAULT true,
     pv int NOT NULL DEFAULT 1,
     created_at timestamptz NOT NULL DEFAULT now(),
     last_login timestamptz)`,
  `CREATE TABLE IF NOT EXISTS meta (
     key text PRIMARY KEY,
     value jsonb NOT NULL,
     updated_at timestamptz NOT NULL DEFAULT now())`,
  `CREATE TABLE IF NOT EXISTS customers (
     id bigserial PRIMARY KEY,
     import_id int NOT NULL,
     row_no int,
     data jsonb NOT NULL,
     version int NOT NULL DEFAULT 1,
     deleted boolean NOT NULL DEFAULT false,
     created_at timestamptz NOT NULL DEFAULT now(),
     updated_at timestamptz NOT NULL DEFAULT now(),
     updated_by text)`,
  `CREATE INDEX IF NOT EXISTS customers_import_idx ON customers (import_id, id)`,
  `CREATE INDEX IF NOT EXISTS customers_updated_idx ON customers (import_id, updated_at)`,
  `CREATE TABLE IF NOT EXISTS changes (
     id bigserial PRIMARY KEY,
     customer_id bigint,
     at timestamptz NOT NULL DEFAULT now(),
     by_user text,
     action text NOT NULL,
     diff jsonb)`,
  `CREATE INDEX IF NOT EXISTS changes_customer_idx ON changes (customer_id, at DESC)`,
  `CREATE TABLE IF NOT EXISTS files (
     name text PRIMARY KEY,
     content text NOT NULL,
     updated_at timestamptz NOT NULL DEFAULT now())`,
  `CREATE TABLE IF NOT EXISTS login_fail (
     username text NOT NULL,
     at timestamptz NOT NULL DEFAULT now())`,
  `CREATE SEQUENCE IF NOT EXISTS import_seq`,
];

async function migrate(f) { for (const s of SCHEMA) await f(s); }

/** Chạy một câu SQL có tham số ($1, $2…) và trả về mảng dòng. */
export async function q(text, params = []) {
  const f = await driver();
  if (!_ready) _ready = migrate(f).catch(e => { _ready = null; throw e; });
  await _ready;
  // Object/array → chuỗi JSON (cho cột jsonb), giống nhau giữa Neon và PGlite
  return f(text, params.map(p => p && typeof p === "object" && !(p instanceof Date) ? JSON.stringify(p) : p));
}
