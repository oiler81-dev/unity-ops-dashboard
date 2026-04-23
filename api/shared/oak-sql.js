// Lazy-initialized mssql connection pool for the OAK staging database
// (sql-srv-unity-prod-bi / sql-unity-prod-bi). Uses SQL auth with the
// read-only oak_reader_app login.
//
// Required Function App env vars:
//   OAK_SQL_SERVER    e.g. sql-srv-unity-prod-bi.database.windows.net
//   OAK_SQL_DATABASE  e.g. sql-unity-prod-bi
//   OAK_SQL_USER      e.g. oak_reader_app
//   OAK_SQL_PASSWORD  from Key Vault sql-unity-bi-oakreader-password
//
// Falls back to OAK_SQL_CONNECTION_STRING (ADO-style) if set.

const sql = require("mssql");

let poolPromise = null;

function getConfig() {
  if (process.env.OAK_SQL_CONNECTION_STRING) {
    return process.env.OAK_SQL_CONNECTION_STRING;
  }
  return {
    server: process.env.OAK_SQL_SERVER,
    database: process.env.OAK_SQL_DATABASE,
    user: process.env.OAK_SQL_USER,
    password: process.env.OAK_SQL_PASSWORD,
    options: {
      encrypt: true,
      trustServerCertificate: false,
      connectTimeout: 15000,
      requestTimeout: 60000
    },
    pool: {
      max: 4,
      min: 0,
      idleTimeoutMillis: 30000
    }
  };
}

async function getPool() {
  if (poolPromise) return poolPromise;
  poolPromise = (async () => {
    const cfg = getConfig();
    const pool = await sql.connect(cfg);
    pool.on("error", () => { poolPromise = null; });
    return pool;
  })().catch((err) => {
    poolPromise = null;
    throw err;
  });
  return poolPromise;
}

async function queryOak(queryText, params = {}) {
  const pool = await getPool();
  const req = pool.request();
  for (const [name, value] of Object.entries(params)) {
    req.input(name, value);
  }
  return req.query(queryText);
}

module.exports = { sql, getPool, queryOak };
