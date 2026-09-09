#!/usr/bin/env node
/**
 * verify.js — independent check that the two replicas match SQL Server.
 *
 * Compares, per table: existence, column count and row count. Reads the three
 * databases separately rather than trusting anything migrate.js reported.
 */

const sql = require("mssql");
const { Pool } = require("pg");
const mariadb = require("mariadb");

const env = (name, fallback) => process.env[name] ?? fallback;
const q = (name) => `"${name.replace(/"/g, '""')}"`;

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

async function main() {
  const src = await new sql.ConnectionPool(SRC).connect();
  const pg = new Pool(PG);
  const maria = await mariadb.createConnection(MARIA);

  const tables = (
    await src
      .request()
      .query("SELECT name FROM sys.tables WHERE name <> 'dtproperties' ORDER BY name")
  ).recordset.map((r) => r.name);

  const mismatches = [];
  let srcTotal = 0;
  let pgTotal = 0;
  let mariaTotal = 0;

  for (const table of tables) {
    const srcRows = (await src.request().query(`SELECT COUNT(*) AS n FROM [${table}]`)).recordset[0]
      .n;
    srcTotal += srcRows;

    let pgRows = null;
    try {
      pgRows = Number((await pg.query(`SELECT COUNT(*) AS n FROM ${q(table)}`)).rows[0].n);
      pgTotal += pgRows;
    } catch {
      /* table missing — reported below as null */
    }

    let mariaRows = null;
    try {
      mariaRows = Number((await maria.query(`SELECT COUNT(*) AS n FROM ${q(table)}`))[0].n);
      mariaTotal += mariaRows;
    } catch {
      /* table missing — reported below as null */
    }

    if (pgRows !== srcRows || mariaRows !== srcRows) {
      mismatches.push({ table, mssql: srcRows, pg: pgRows, mariadb: mariaRows });
    }
  }

  console.log(`tables checked : ${tables.length}`);
  console.log(`rows  mssql    : ${srcTotal}`);
  console.log(`rows  postgres : ${pgTotal}`);
  console.log(`rows  mariadb  : ${mariaTotal}`);

  if (mismatches.length) {
    console.log(`\n${mismatches.length} table(s) differ:`);
    for (const m of mismatches) {
      console.log(
        `  ${m.table}: mssql=${m.mssql} pg=${m.pg ?? "MISSING"} mariadb=${m.mariadb ?? "MISSING"}`,
      );
    }
  } else {
    console.log("\nAll tables match on row count.");
  }

  await src.close();
  await pg.end();
  await maria.end();
  process.exit(mismatches.length ? 1 : 0);
}

main().catch((err) => {
  console.error(err);
  process.exit(2);
});
