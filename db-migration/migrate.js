#!/usr/bin/env node
/**
 * migrate.js — replicate a SQL Server database into PostgreSQL and MariaDB.
 *
 *   node migrate.js              # schema + data into both targets
 *   node migrate.js --ddl-only   # write out/*.sql, touch no target
 *   node migrate.js --target pg  # one target only (pg | maria)
 *
 * Connection settings come from the environment (see README.md); the defaults
 * match the containers created for this migration.
 *
 * Not translated: triggers, stored procedures, views and functions. T-SQL
 * bodies do not machine-convert reliably, so they are dumped verbatim to
 * out/unported/ for manual porting and reported in the summary.
 */

const fs = require("fs");
const path = require("path");
const sql = require("mssql");
const { Pool } = require("pg");
const mariadb = require("mariadb");

const { introspect } = require("./lib/introspect");
const { emitPgTable, emitMariaTable, q } = require("./lib/emit");

const OUT = path.join(__dirname, "out");

const argv = process.argv.slice(2);
const DDL_ONLY = argv.includes("--ddl-only");
const TARGET = (argv[argv.indexOf("--target") + 1] || "both").toLowerCase();
const WANT_PG = !DDL_ONLY && (TARGET === "both" || TARGET === "pg");
const WANT_MARIA = !DDL_ONLY && (TARGET === "both" || TARGET === "maria");

const env = (name, fallback) => process.env[name] ?? fallback;

const SRC = {
  server: env("SRC_HOST", "localhost"),
  port: Number(env("SRC_PORT", 1433)),
  user: env("SRC_USER", "sa"),
  password: env("SRC_PASSWORD", "Password1!"),
  database: env("SRC_DB", "EMR6122"),
  options: { encrypt: false, trustServerCertificate: true },
  requestTimeout: 120000,
};

const PG = {
  host: env("PG_HOST", "localhost"),
  port: Number(env("PG_PORT", 5433)),
  user: env("PG_USER", "sa"),
  password: env("PG_PASSWORD", "Password1!"),
  database: env("PG_DB", "EMR6122"),
};

const MARIA = {
  host: env("MARIA_HOST", "127.0.0.1"),
  port: Number(env("MARIA_PORT", 3307)),
  user: env("MARIA_USER", "root"),
  password: env("MARIA_PASSWORD", "Password1!"),
  database: env("MARIA_DB", "EMR6122"),
};

const log = (...a) => console.log(...a);
const warnings = [];
const warn = (msg) => {
  warnings.push(msg);
  console.warn("  ! " + msg);
};

/**
 * PostgreSQL rejects NUL bytes inside text values; legacy fixed-width data
 * sometimes carries them. Strip and count rather than failing the table.
 */
let nulStripped = 0;
const NUL = String.fromCharCode(0);
const NUL_RE = new RegExp(NUL, "g");
function scrub(value) {
  if (typeof value === "string" && value.indexOf(NUL) !== -1) {
    nulStripped++;
    return value.replace(NUL_RE, "");
  }
  return value;
}

async function writeDdl(model) {
  fs.mkdirSync(path.join(OUT, "unported"), { recursive: true });

  const pgDdl = [];
  const mariaDdl = [];
  for (const table of model.tables) {
    pgDdl.push(...emitPgTable(table), "");
    mariaDdl.push(...emitMariaTable(table), "");
  }

  fs.writeFileSync(path.join(OUT, "schema.pg.sql"), pgDdl.join("\n"));
  fs.writeFileSync(path.join(OUT, "schema.mariadb.sql"), mariaDdl.join("\n"));

  const kinds = {
    TR: "triggers",
    P: "procedures",
    V: "views",
    FN: "functions",
    TF: "functions",
    IF: "functions",
  };
  const buckets = {};
  for (const m of model.modules) {
    const bucket = kinds[m.type.trim()] || "other";
    (buckets[bucket] ||= []).push(`-- ${m.name}\n${m.definition}\nGO\n`);
  }
  for (const [bucket, defs] of Object.entries(buckets)) {
    fs.writeFileSync(path.join(OUT, "unported", `${bucket}.tsql`), defs.join("\n"));
    log(`  unported: ${defs.length} ${bucket} -> out/unported/${bucket}.tsql`);
  }
}

/**
 * Returns the tables whose CREATE failed, so the caller can skip their indexes
 * and their data instead of drowning the log in follow-on "does not exist".
 */
async function createPgSchema(pg, model) {
  const failed = new Set();
  await pg.query("DROP SCHEMA public CASCADE; CREATE SCHEMA public;");
  for (const table of model.tables) {
    const [create, ...indexes] = emitPgTable(table);
    try {
      await pg.query(create);
    } catch (err) {
      warn(`pg CREATE TABLE failed for ${table.name}: ${err.message}`);
      failed.add(table.name);
      continue;
    }
    for (const stmt of indexes) {
      try {
        await pg.query(stmt);
      } catch (err) {
        warn(`pg index failed on ${table.name}: ${err.message}`);
      }
    }
  }
  return failed;
}

async function createMariaSchema(conn, model) {
  const failed = new Set();
  await conn.query("SET FOREIGN_KEY_CHECKS = 0");
  for (const table of model.tables) {
    await conn.query(`DROP TABLE IF EXISTS ${q(table.name)}`);
  }
  for (const table of model.tables) {
    const [create, ...indexes] = emitMariaTable(table);
    try {
      await conn.query(create);
    } catch (err) {
      warn(`mariadb CREATE TABLE failed for ${table.name}: ${err.message}`);
      failed.add(table.name);
      continue;
    }
    for (const stmt of indexes) {
      try {
        await conn.query(stmt);
      } catch (err) {
        warn(`mariadb index failed on ${table.name}: ${err.message}`);
      }
    }
  }
  return failed;
}

/** Read one table from SQL Server and fan it out to whichever targets are on. */
async function copyTable(src, table, pg, maria) {
  const cols = table.columns.filter((c) => !c.computed).sort((a, b) => a.ordinal - b.ordinal);
  const names = cols.map((c) => c.name);
  const selectList = names.map((n) => `[${n}]`).join(", ");

  const result = await src.request().query(`SELECT ${selectList} FROM [${table.name}]`);
  const rows = result.recordset;
  if (!rows.length) return 0;

  const values = rows.map((row) => names.map((n) => scrub(row[n])));

  if (pg) {
    // Keep each statement under PostgreSQL's 65535-parameter ceiling.
    const batchSize = Math.max(1, Math.floor(30000 / names.length));
    const colList = names.map(q).join(", ");
    for (let i = 0; i < values.length; i += batchSize) {
      const chunk = values.slice(i, i + batchSize);
      const params = [];
      const tuples = chunk.map((row) => {
        const placeholders = row.map((v) => {
          params.push(v);
          return `$${params.length}`;
        });
        return `(${placeholders.join(", ")})`;
      });
      await pg.query(
        `INSERT INTO ${q(table.name)} (${colList}) VALUES ${tuples.join(", ")}`,
        params,
      );
    }
  }

  if (maria) {
    const colList = names.map(q).join(", ");
    const placeholders = names.map(() => "?").join(", ");
    const stmt = `INSERT INTO ${q(table.name)} (${colList}) VALUES (${placeholders})`;
    for (let i = 0; i < values.length; i += 500) {
      await maria.batch(stmt, values.slice(i, i + 500));
    }
  }

  return rows.length;
}

/** A bulk INSERT does not advance PostgreSQL identity sequences. */
async function resyncPgIdentities(pg, model, failed) {
  for (const table of model.tables) {
    if (failed.has(table.name)) continue;
    const idCol = table.columns.find((c) => c.identity);
    if (!idCol) continue;
    try {
      // pg_get_serial_sequence folds an unquoted name to lower case, so the
      // table has to be handed to it already quoted.
      await pg.query(
        `SELECT setval(pg_get_serial_sequence($1, $2),
                       COALESCE((SELECT MAX(${q(idCol.name)}) FROM ${q(table.name)}), 0) + 1,
                       false)`,
        [q(table.name), idCol.name],
      );
    } catch (err) {
      warn(`pg identity resync failed for ${table.name}: ${err.message}`);
    }
  }
}

async function resyncMariaIdentities(conn, model, failed) {
  for (const table of model.tables) {
    if (failed.has(table.name)) continue;
    const idCol = table.columns.find((c) => c.identity);
    if (!idCol) continue;
    try {
      const [row] = await conn.query(
        `SELECT COALESCE(MAX(${q(idCol.name)}), 0) + 1 AS next FROM ${q(table.name)}`,
      );
      await conn.query(`ALTER TABLE ${q(table.name)} AUTO_INCREMENT = ${Number(row.next)}`);
    } catch (err) {
      warn(`mariadb identity resync failed for ${table.name}: ${err.message}`);
    }
  }
}

async function main() {
  log(`\nSource : mssql://${SRC.server}:${SRC.port}/${SRC.database}`);
  if (WANT_PG) log(`Target : postgres://${PG.host}:${PG.port}/${PG.database}`);
  if (WANT_MARIA) log(`Target : mariadb://${MARIA.host}:${MARIA.port}/${MARIA.database}`);

  const src = await new sql.ConnectionPool(SRC).connect();
  log("\n[1/5] Reading SQL Server catalog");
  const model = await introspect(src);
  log(`  ${model.tables.length} tables, ${model.modules.length} programmable objects`);

  log("\n[2/5] Writing DDL to out/");
  await writeDdl(model);

  let pg = null;
  let maria = null;
  if (WANT_PG) pg = new Pool(PG);
  if (WANT_MARIA) maria = await mariadb.createConnection(MARIA);

  if (pg || maria) {
    log("\n[3/5] Creating schema");
    const pgFailed = pg ? await createPgSchema(pg, model) : new Set();
    const mariaFailed = maria ? await createMariaSchema(maria, model) : new Set();

    log("\n[4/5] Copying data");
    let total = 0;
    let done = 0;
    for (const table of model.tables) {
      done++;
      try {
        const n = await copyTable(
          src,
          table,
          pgFailed.has(table.name) ? null : pg,
          mariaFailed.has(table.name) ? null : maria,
        );
        total += n;
        if (n) log(`  [${done}/${model.tables.length}] ${table.name}: ${n}`);
      } catch (err) {
        warn(`copy failed for ${table.name}: ${err.message}`);
      }
    }
    log(`  ${total} rows copied`);
    if (nulStripped) warn(`stripped NUL bytes from ${nulStripped} string values`);

    log("\n[5/5] Resyncing identity counters");
    if (pg) await resyncPgIdentities(pg, model, pgFailed);
    if (maria) await resyncMariaIdentities(maria, model, mariaFailed);
  }

  await src.close();
  if (pg) await pg.end();
  if (maria) await maria.end();

  log(`\nDone. ${warnings.length} warning(s).`);
  if (warnings.length) {
    fs.writeFileSync(path.join(OUT, "warnings.txt"), warnings.join("\n") + "\n");
    log("Full list in out/warnings.txt");
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
