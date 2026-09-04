# db-migration — replicate the SQL Server EMR database into PostgreSQL and MariaDB

Reads a SQL Server database's catalog, emits equivalent DDL for PostgreSQL and
MariaDB, creates the schema in both, and copies every row.

Built against `EMR6122` (156 tables, 70,828 rows). It is schema-driven, not
hardcoded, so it runs against `Emr` or `8848Emr` by changing `SRC_DB` — but see
[Scaling to the bigger databases](#scaling-to-the-bigger-databases) first.

## Running it

```bash
npm install
```

```bash
node migrate.js
```

```bash
node verify.js
```

`migrate.js` flags:

| Flag | Effect |
|---|---|
| _(none)_ | schema + data into both targets |
| `--ddl-only` | write `out/*.sql` only, touch no target database |
| `--target pg` | PostgreSQL only (`--target maria` for MariaDB only) |

`verify.js` re-reads all three databases independently and compares row counts
per table. It exits non-zero on any mismatch, so it is safe to use in a script.

Everything is re-runnable: each run drops and recreates the target schema.

## The containers

The source container already existed; the two targets were created for this.

| | Container | Host port | Database | User | Password |
|---|---|---|---|---|---|
| Source | `mssql-container` | 1433 | `EMR6122` | `sa` | `Password1!` |
| Target | `pg-emr6122` | **5433** | `EMR6122` | `sa` | `Password1!` |
| Target | `mariadb-emr6122` | **3307** | `EMR6122` | `root` | `Password1!` |

Non-default host ports keep these clear of anything already on 5432/3306. Data
lives in the named volumes `pg-emr6122-data` and `mariadb-emr6122-data`.

Both carry `--restart unless-stopped`, so they come back after a host reboot the
way `mssql-container` does. Recreate them with:

```bash
docker run -d --name pg-emr6122 --restart unless-stopped -e POSTGRES_PASSWORD='Password1!' -e POSTGRES_USER=sa -e POSTGRES_DB=EMR6122 -p 5433:5432 -v pg-emr6122-data:/var/lib/postgresql/data postgres:16
```

```bash
docker run -d --name mariadb-emr6122 --restart unless-stopped -e MARIADB_ROOT_PASSWORD='Password1!' -e MARIADB_DATABASE=EMR6122 -p 3307:3306 -v mariadb-emr6122-data:/var/lib/mysql mariadb:11 --sql-mode='ANSI_QUOTES,PIPES_AS_CONCAT,STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_ENGINE_SUBSTITUTION'
```

**The MariaDB `--sql-mode` is not optional.** `ANSI_QUOTES` makes `"OmCoCd"` an
identifier rather than a string literal, which is how `emr_node` writes every
query; `PIPES_AS_CONCAT` makes `||` concatenation, which the generated columns
rely on. Start the container without them and both the schema and the app break.

Override any connection setting through the environment: `SRC_*`, `PG_*`,
`MARIA_*` (`_HOST`, `_PORT`, `_USER`, `_PASSWORD`, `_DB`).

## Layout

| Path | What it is |
|---|---|
| `lib/introspect.js` | Reads the SQL Server catalog into a plain object. The only file that speaks T-SQL. |
| `lib/emit.js` | Turns that object into PostgreSQL and MariaDB DDL — type mapping, expression translation, dialect workarounds. |
| `migrate.js` | Orchestrates: introspect → write DDL → create schema → copy data → resync identity counters. |
| `verify.js` | Independent row-count comparison across all three databases. |
| `out/schema.pg.sql`, `out/schema.mariadb.sql` | The generated DDL, for reading or applying by hand. |
| `out/unported/` | T-SQL bodies this tool does not translate (see below). |
| `out/warnings.txt` | Written only when a run has warnings. |

## What actually gets replicated

Verified on the last run: **156 tables, 70,828 rows, identical in all three
databases**, plus primary keys, secondary indexes, check constraints, column
defaults, identity columns, and the one computed column.

### Not translated

`out/unported/` holds these verbatim as T-SQL. They are dumped, not converted —
T-SQL procedural bodies do not machine-convert into plpgsql or MariaDB SQL/PSM
reliably, and a silent bad conversion is worse than an obvious absence.

| Object | Count | Notes |
|---|---|---|
| Triggers | 118 | The ones that matter. `emr_node/database/triggers.sql` already has 3 ported to plpgsql — that file is the pattern to follow. |
| Stored procedures | 91 | `emr_node` calls none of them; they are VB6-era. Likely safe to drop rather than port. |
| Functions | 8 | Same. |
| Views | 5 | Small; hand-porting these is quick. |

Triggers are the real gap: they fire implicitly on INSERT/UPDATE, so any app
behaviour that depends on them is missing from both replicas until ported. Port
only the triggers on tables `emr_node` actually writes to.

## Fidelity notes

Things that are deliberately not byte-identical, and why.

**MariaDB widens some `VARCHAR` columns to `LONGTEXT`.** InnoDB caps a row at
roughly half a 16K page — far below the 65535-byte declared-width limit — and
`utf8mb4` bills every character at 4 bytes. Eight wide tables would not fit, so
the widest columns are demoted (they store off-page for the cost of a pointer):

| Table | Demoted | of |
|---|---|---|
| `UsrOpt` | 273 | 520 columns |
| `InvHd` | 91 | 286 |
| `Head` | 49 | 332 |
| `ClrSet` | 6 | 92 |
| `TmpA` | 6 | 122 |
| `TagData` | 5 | 25 |
| `CustMst` | 1 | 80 |
| `SBOutput` | 1 | 8 |

These columns lose their length constraint, and indexing them uses a 255-char
prefix. PostgreSQL has no such limit and keeps every `varchar(n)` intact.

**`DEFAULT ''` on date columns is dropped.** SQL Server silently coerces it to
1900-01-01; both targets reject it outright, and the coerced value carries no
meaning, so no default is invented in its place.

**Date defaults are rewritten to ISO.** The schema writes `'01/01/1980'`.
PostgreSQL parses that under its MDY `DateStyle`; MariaDB does not.

**`CONVERT(char(n), …)` becomes `RPAD(CAST(… AS CHAR(n)), n, ' ')`.** SQL Server
blank-pads to `n` and keeps the padding through concatenation. PostgreSQL strips
a `bpchar`'s trailing blanks as soon as it is used as text, and MariaDB never
pads. Without the `RPAD`, `RmRt.RrAllKeys` came out 31 characters against SQL
Server's 53 — on a column carrying a unique index. It now matches exactly across
all 524 rows.

**MariaDB index names are truncated to 64 characters** (SQL Server allows 128),
with a hash suffix so they stay unique. Two indexes on `SYyMm` hit this.

**NUL bytes in string data are stripped**, because PostgreSQL rejects them in
`text`. The run reports a count; the last run stripped none.

**Type mappings worth knowing:** `smalldatetime` → `timestamp(0)` / `DATETIME`;
`float` → `double precision` / `DOUBLE`; `tinyint` → `smallint` /
`TINYINT UNSIGNED`; `image` → `bytea` / `LONGBLOB`. `datetimeoffset` maps to
`timestamptz` in PostgreSQL but has no MariaDB equivalent — it becomes
`DATETIME(6)` holding UTC. `EMR6122` has no `datetimeoffset` columns; `Emr` has
14, so this matters if you retarget.

**Identity counters are resynced after load.** A bulk `INSERT` does not advance
a PostgreSQL identity sequence, so the first application insert would collide.
20 of the identity columns lead no index, which MariaDB requires for
`AUTO_INCREMENT`; the emitter adds a plain key rather than dropping identity.

## Scaling to the bigger databases

`EMR6122` is 117 MB. `Emr` is 7.9 GB across 9.85 million rows, and the copier
currently reads each table fully into memory before writing. That is fine at this
size and will not be at that one — `Emr` needs streaming (`request().stream`) or
a CSV round-trip through `bcp` + `COPY` / `LOAD DATA` before this is pointed at
it.

Neither database has foreign keys, so tables can be loaded in any order, and
parallelising the copy is straightforward when it becomes worth doing.
