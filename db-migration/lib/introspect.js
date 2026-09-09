/**
 * introspect.js — read the SQL Server catalog into a plain-JS model.
 *
 * Everything downstream (DDL emitters, data copier) works off the object this
 * returns, so the SQL Server dialect is confined to this one file.
 */

/** Columns, in ordinal order, with the bits both emitters need. */
const COLUMNS_SQL = `
SELECT t.name              AS table_name,
       c.name              AS column_name,
       c.column_id         AS ordinal,
       ty.name             AS sql_type,
       c.max_length        AS max_length,
       c.precision         AS precision,
       c.scale             AS scale,
       c.is_nullable       AS is_nullable,
       c.is_identity       AS is_identity,
       c.is_computed       AS is_computed,
       cc.definition       AS computed_definition,
       dc.definition       AS default_definition
FROM sys.columns c
JOIN sys.tables t            ON t.object_id = c.object_id
JOIN sys.types ty            ON ty.user_type_id = c.user_type_id
LEFT JOIN sys.computed_columns cc
       ON cc.object_id = c.object_id AND cc.column_id = c.column_id
LEFT JOIN sys.default_constraints dc
       ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
ORDER BY t.name, c.column_id`;

/** Primary keys, columns in key order. */
const PK_SQL = `
SELECT t.name AS table_name, c.name AS column_name, ic.key_ordinal AS key_ordinal
FROM sys.key_constraints kc
JOIN sys.tables t          ON t.object_id = kc.parent_object_id
JOIN sys.index_columns ic  ON ic.object_id = kc.parent_object_id AND ic.index_id = kc.unique_index_id
JOIN sys.columns c         ON c.object_id = ic.object_id AND c.column_id = ic.column_id
WHERE kc.type = 'PK'
ORDER BY t.name, ic.key_ordinal`;

/**
 * Secondary indexes only (index_id > 1 skips the heap and the clustered index,
 * the latter being the PK we already emit as a constraint).
 */
const INDEX_SQL = `
SELECT t.name AS table_name, i.name AS index_name, i.is_unique AS is_unique,
       c.name AS column_name, ic.key_ordinal AS key_ordinal, ic.is_descending_key AS is_desc
FROM sys.indexes i
JOIN sys.tables t          ON t.object_id = i.object_id
JOIN sys.index_columns ic  ON ic.object_id = i.object_id AND ic.index_id = i.index_id
JOIN sys.columns c         ON c.object_id = ic.object_id AND c.column_id = ic.column_id
WHERE i.index_id > 1 AND i.is_primary_key = 0 AND ic.is_included_column = 0
ORDER BY t.name, i.name, ic.key_ordinal`;

const CHECK_SQL = `
SELECT t.name AS table_name, ck.name AS constraint_name, ck.definition AS definition
FROM sys.check_constraints ck
JOIN sys.tables t ON t.object_id = ck.parent_object_id
ORDER BY t.name, ck.name`;

/** Objects we deliberately do not translate — dumped for manual porting. */
const MODULE_SQL = `
SELECT o.name AS name, o.type AS type, m.definition AS definition
FROM sys.sql_modules m
JOIN sys.objects o ON o.object_id = m.object_id
WHERE o.type IN ('TR', 'P', 'V', 'FN', 'TF', 'IF')
ORDER BY o.type, o.name`;

async function introspect(pool) {
  const run = async (sql) => (await pool.request().query(sql)).recordset;

  const [cols, pks, idx, checks, modules] = await Promise.all([
    run(COLUMNS_SQL),
    run(PK_SQL),
    run(INDEX_SQL),
    run(CHECK_SQL),
    run(MODULE_SQL),
  ]);

  const tables = new Map();
  const table = (name) => {
    if (!tables.has(name)) {
      tables.set(name, { name, columns: [], pk: [], indexes: new Map(), checks: [] });
    }
    return tables.get(name);
  };

  for (const c of cols) {
    table(c.table_name).columns.push({
      name: c.column_name,
      ordinal: c.ordinal,
      sqlType: c.sql_type,
      maxLength: c.max_length,
      precision: c.precision,
      scale: c.scale,
      nullable: !!c.is_nullable,
      identity: !!c.is_identity,
      computed: !!c.is_computed,
      computedDefinition: c.computed_definition,
      defaultDefinition: c.default_definition,
    });
  }

  for (const p of pks) table(p.table_name).pk.push(p.column_name);

  for (const i of idx) {
    const t = table(i.table_name);
    if (!t.indexes.has(i.index_name)) {
      t.indexes.set(i.index_name, { name: i.index_name, unique: !!i.is_unique, columns: [] });
    }
    t.indexes.get(i.index_name).columns.push({ name: i.column_name, desc: !!i.is_desc });
  }

  for (const ck of checks) {
    table(ck.table_name).checks.push({ name: ck.constraint_name, definition: ck.definition });
  }

  // sys.tables includes the SQL Server designer's own bookkeeping table.
  tables.delete("dtproperties");

  for (const t of tables.values()) t.indexes = [...t.indexes.values()];

  return { tables: [...tables.values()], modules };
}

module.exports = { introspect };
