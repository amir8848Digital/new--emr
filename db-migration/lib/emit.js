/**
 * emit.js — turn the introspected model into CREATE TABLE / CREATE INDEX DDL
 * for PostgreSQL and MariaDB.
 *
 * Both targets are addressed with double-quoted identifiers, which is why the
 * MariaDB container runs with ANSI_QUOTES (and PIPES_AS_CONCAT, so `||` means
 * concatenation there too). That keeps one quoting convention across all three
 * databases and matches how emr_node already writes its SQL.
 */

const q = (name) => `"${name.replace(/"/g, '""')}"`;

// InnoDB caps a row at roughly half a 16K page (8126 bytes) — a far tighter
// bound than the 65535-byte declared-width limit, and the one wide tables here
// actually hit (Head has 332 columns, UsrOpt has 520). Columns demoted to TEXT
// are stored off-page and cost only a pointer, so demoting the widest few
// brings the rest of the table back under the ceiling.
const MARIA_ROW_LIMIT = 7800;
const MARIA_BLOB_COST = 20;

/** char/varchar length in characters; null means MAX. */
function charLength(col) {
  if (col.maxLength === -1) return null;
  const isWide = col.sqlType.startsWith("n");
  return isWide ? col.maxLength / 2 : col.maxLength;
}

function pgType(col) {
  const t = col.sqlType;
  const len = charLength(col);
  switch (t) {
    case "varchar":
    case "nvarchar":
      return len === null ? "text" : `varchar(${len})`;
    case "char":
    case "nchar":
      return len === null ? "text" : `char(${len})`;
    case "text":
    case "ntext":
    case "sysname":
      return t === "sysname" ? "varchar(128)" : "text";
    case "int":
      return "integer";
    case "bigint":
      return "bigint";
    case "smallint":
      return "smallint";
    case "tinyint":
      return "smallint"; // SQL Server tinyint is 0-255, unsigned; smallint holds it
    case "bit":
      return "boolean";
    case "float":
    case "real":
      return "double precision";
    case "money":
    case "smallmoney":
      return "numeric(19,4)";
    case "numeric":
    case "decimal":
      return `numeric(${col.precision},${col.scale})`;
    case "smalldatetime":
      return "timestamp(0)";
    case "datetime":
      return "timestamp(3)";
    case "datetime2":
      return "timestamp(6)";
    case "date":
      return "date";
    case "time":
      return "time";
    case "datetimeoffset":
      return "timestamptz";
    case "uniqueidentifier":
      return "uuid";
    case "binary":
    case "varbinary":
    case "image":
      return "bytea";
    default:
      throw new Error(`No PostgreSQL mapping for SQL Server type "${t}"`);
  }
}

function mariaType(col, { forceText = false } = {}) {
  const t = col.sqlType;
  const len = charLength(col);
  switch (t) {
    case "varchar":
    case "nvarchar":
      if (len === null || forceText) return "LONGTEXT";
      return `VARCHAR(${len})`;
    case "char":
    case "nchar":
      if (len === null || forceText) return "LONGTEXT";
      return `CHAR(${len})`;
    case "text":
    case "ntext":
      return "LONGTEXT";
    case "sysname":
      return "VARCHAR(128)";
    case "int":
      return "INT";
    case "bigint":
      return "BIGINT";
    case "smallint":
      return "SMALLINT";
    case "tinyint":
      return "TINYINT UNSIGNED";
    case "bit":
      return "TINYINT(1)";
    case "float":
    case "real":
      return "DOUBLE";
    case "money":
    case "smallmoney":
      return "DECIMAL(19,4)";
    case "numeric":
    case "decimal":
      return `DECIMAL(${col.precision},${col.scale})`;
    case "smalldatetime":
      return "DATETIME";
    case "datetime":
      return "DATETIME(3)";
    case "datetime2":
      return "DATETIME(6)";
    case "date":
      return "DATE";
    case "time":
      return "TIME";
    case "datetimeoffset":
      return "DATETIME(6)"; // MariaDB has no offset-aware type; values arrive as UTC
    case "uniqueidentifier":
      return "CHAR(36)";
    case "binary":
    case "varbinary":
      return len === null ? "LONGBLOB" : `VARBINARY(${len})`;
    case "image":
      return "LONGBLOB";
    default:
      throw new Error(`No MariaDB mapping for SQL Server type "${t}"`);
  }
}

/** Declared byte cost of a column under utf8mb4, for the row-width guard. */
function mariaWidth(col, forceText) {
  const type = mariaType(col, { forceText });
  if (/TEXT|BLOB/.test(type)) return MARIA_BLOB_COST;
  const m = /^(?:VAR)?(?:CHAR|BINARY)\((\d+)\)/.exec(type);
  if (m) {
    const bytes = /BINARY/.test(type) ? Number(m[1]) : Number(m[1]) * 4;
    return bytes + 2;
  }
  return 8;
}

/**
 * Pick the columns to demote to LONGTEXT so the table fits MariaDB's row limit.
 * Widest first, so as few columns as possible lose their length constraint.
 */
function planMariaTextDemotions(table) {
  const candidates = table.columns
    .filter((c) => !c.computed && /char$/.test(c.sqlType) && charLength(c) !== null)
    .sort((a, b) => charLength(b) - charLength(a));

  const demoted = new Set();
  const width = () =>
    table.columns
      .filter((c) => !c.computed)
      .reduce((sum, c) => sum + mariaWidth(c, demoted.has(c.name)), 0);

  for (const c of candidates) {
    if (width() <= MARIA_ROW_LIMIT) break;
    demoted.add(c.name);
  }
  return demoted;
}

/**
 * Translate a SQL Server expression (default / computed / check definition).
 * Handles the constructs this schema actually uses: bracketed identifiers,
 * string literals, `+` concatenation and a couple of built-ins.
 */
function translateExpr(definition, dialect) {
  let expr = definition;

  // Defaults arrive wrapped: ('') , ((0)) — peel balanced outer parens.
  let peeled = true;
  while (peeled && expr.startsWith("(") && expr.endsWith(")")) {
    let depth = 0;
    peeled = true;
    for (let i = 0; i < expr.length; i++) {
      if (expr[i] === "(") depth++;
      else if (expr[i] === ")") depth--;
      if (depth === 0 && i < expr.length - 1) {
        peeled = false; // the leading "(" closes early, so it is not a wrapper
        break;
      }
    }
    if (peeled) expr = expr.slice(1, -1).trim();
  }

  expr = rewriteConvert(expr);

  // `+` means concatenation only when a string is involved; between numbers it
  // is addition and must survive untouched. Rewriting is done per-segment so a
  // `+` or a `[` inside a string literal is never touched.
  const isConcat = /'/.test(expr);

  return splitLiterals(expr)
    .map(({ isString, text }) => {
      if (isString) {
        // MariaDB treats backslash as an escape character inside literals.
        return dialect === "maria" ? text.replace(/\\/g, "\\\\") : text;
      }
      let code = text.replace(/\[([^\]]+)\]/g, (_, name) => q(name));
      if (isConcat) code = code.replace(/\+/g, " || ");
      return code
        .replace(/\bgetdate\s*\(\s*\)/gi, "CURRENT_TIMESTAMP")
        .replace(/\bsysdatetime\s*\(\s*\)/gi, "CURRENT_TIMESTAMP")
        .replace(/\bnewid\s*\(\s*\)/gi, dialect === "pg" ? "gen_random_uuid()" : "UUID()");
    })
    .join("");
}

/** Split into alternating code / single-quoted-literal segments. */
function splitLiterals(text) {
  const parts = [];
  let buffer = "";
  let i = 0;

  while (i < text.length) {
    if (text[i] !== "'") {
      buffer += text[i++];
      continue;
    }
    if (buffer) {
      parts.push({ isString: false, text: buffer });
      buffer = "";
    }
    let literal = "'";
    i++;
    while (i < text.length) {
      if (text[i] === "'") {
        if (text[i + 1] === "'") {
          literal += "''"; // an escaped quote, not the end
          i += 2;
          continue;
        }
        literal += "'";
        i++;
        break;
      }
      literal += text[i++];
    }
    parts.push({ isString: true, text: literal });
  }
  if (buffer) parts.push({ isString: false, text: buffer });
  return parts;
}

/**
 * CONVERT(type, expr [, style]) is T-SQL only; both targets speak CAST.
 * The style argument has no portable equivalent and is dropped — for the one
 * place this schema uses it (RmRt.RrAllKeys) the style is 0, the default.
 *
 * Converting to CHAR(n) also has to be padded explicitly. SQL Server blank-pads
 * to n and keeps the padding through concatenation; PostgreSQL strips a
 * bpchar's trailing blanks the moment it is used as text, and MariaDB never
 * pads at all. Without the RPAD, RmRt.RrAllKeys came out 31 characters against
 * SQL Server's 53 — and that column carries a unique index.
 */
function rewriteConvert(expr) {
  for (;;) {
    const match = /\bCONVERT\s*\(/i.exec(expr);
    if (!match) return expr;

    const open = match.index + match[0].length - 1;
    let depth = 0;
    let close = -1;
    for (let i = open; i < expr.length; i++) {
      if (expr[i] === "(") depth++;
      else if (expr[i] === ")" && --depth === 0) {
        close = i;
        break;
      }
    }
    if (close === -1) return expr; // unbalanced; leave it for a human

    const args = splitTopLevelArgs(expr.slice(open + 1, close));
    if (args.length < 2) return expr;
    const type = args[0].trim().replace(/\[([^\]]+)\]/g, "$1");
    const value = args[1].trim();
    const charType = /^n?char\s*\(\s*(\d+)\s*\)$/i.exec(type);
    const cast = charType
      ? `RPAD(CAST(${value} AS ${type}), ${charType[1]}, ' ')`
      : `CAST(${value} AS ${type})`;
    expr = expr.slice(0, match.index) + cast + expr.slice(close + 1);
  }
}

function splitTopLevelArgs(text) {
  const args = [];
  let depth = 0;
  let current = "";
  for (const ch of text) {
    if (ch === "(") depth++;
    else if (ch === ")") depth--;
    if (ch === "," && depth === 0) {
      args.push(current);
      current = "";
      continue;
    }
    current += ch;
  }
  args.push(current);
  return args;
}

const TEMPORAL = /^(smalldatetime|datetime|datetime2|date|time|datetimeoffset)$/;

/**
 * SQL Server tolerates DEFAULT '' on a date/time column, silently coercing it
 * to 1900-01-01. PostgreSQL and MariaDB both reject it, and the coerced value
 * carries no meaning, so the default is dropped rather than invented.
 */
function usableDefault(col) {
  if (!col.defaultDefinition) return false;
  if (TEMPORAL.test(col.sqlType) && /^\(*''\)*$/.test(col.defaultDefinition)) return false;
  return true;
}

/**
 * Date defaults in this schema are written US-style ('01/01/1980', '01/01/80').
 * PostgreSQL parses those under its default MDY DateStyle; MariaDB does not and
 * rejects the column outright, so temporal literals are normalised to ISO.
 */
function normaliseTemporalDefault(col, translated) {
  if (!TEMPORAL.test(col.sqlType)) return translated;
  const literal = /^'(.*)'$/.exec(translated);
  if (!literal) return translated;

  const us = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(literal[1].trim());
  if (!us) return translated;

  const [, month, day, rawYear] = us;
  // Two-digit years here are all 20th century (the sentinel is 1980).
  const year = rawYear.length === 4 ? rawYear : `19${rawYear}`;
  return `'${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")} 00:00:00'`;
}

/** MariaDB caps identifiers at 64 characters; SQL Server allows 128. */
function mariaIdentifier(name) {
  if (name.length <= 64) return q(name);
  let hash = 0;
  for (let i = 0; i < name.length; i++) hash = (hash * 31 + name.charCodeAt(i)) >>> 0;
  return q(`${name.slice(0, 55)}_${hash.toString(36)}`);
}

function emitPgTable(table) {
  const lines = [];

  for (const col of table.columns) {
    if (col.computed) {
      lines.push(
        `  ${q(col.name)} text GENERATED ALWAYS AS (${translateExpr(col.computedDefinition, "pg")}) STORED`,
      );
      continue;
    }
    let line = `  ${q(col.name)} ${pgType(col)}`;
    if (col.identity) line += " GENERATED BY DEFAULT AS IDENTITY";
    if (usableDefault(col)) {
      line += ` DEFAULT ${normaliseTemporalDefault(col, translateExpr(col.defaultDefinition, "pg"))}`;
    }
    if (!col.nullable) line += " NOT NULL";
    lines.push(line);
  }

  if (table.pk.length) {
    lines.push(`  CONSTRAINT ${q(`pk_${table.name}`)} PRIMARY KEY (${table.pk.map(q).join(", ")})`);
  }
  for (const ck of table.checks) {
    lines.push(`  CONSTRAINT ${q(ck.name)} CHECK (${translateExpr(ck.definition, "pg")})`);
  }

  const ddl = [`CREATE TABLE ${q(table.name)} (\n${lines.join(",\n")}\n);`];

  for (const idx of table.indexes) {
    const cols = idx.columns.map((c) => `${q(c.name)}${c.desc ? " DESC" : ""}`).join(", ");
    // Index names are schema-scoped in PostgreSQL but table-scoped in SQL
    // Server, so prefix with the table to keep them unique.
    ddl.push(
      `CREATE ${idx.unique ? "UNIQUE " : ""}INDEX ${q(`${table.name}_${idx.name}`)} ON ${q(table.name)} (${cols});`,
    );
  }

  return ddl;
}

function emitMariaTable(table) {
  const demoted = planMariaTextDemotions(table);
  const lines = [];

  // AUTO_INCREMENT requires the column to lead some key. 20 identity columns
  // here lead none, so add a plain key rather than silently dropping identity.
  const identityCol = table.columns.find((c) => c.identity);
  const leadsAKey =
    identityCol &&
    (table.pk[0] === identityCol.name ||
      table.indexes.some((i) => i.columns[0]?.name === identityCol.name));

  for (const col of table.columns) {
    if (col.computed) {
      lines.push(
        `  ${q(col.name)} LONGTEXT AS (${translateExpr(col.computedDefinition, "maria")}) STORED`,
      );
      continue;
    }
    let line = `  ${q(col.name)} ${mariaType(col, { forceText: demoted.has(col.name) })}`;
    if (!col.nullable) line += " NOT NULL";
    if (col.identity) line += " AUTO_INCREMENT";
    // MariaDB rejects DEFAULT on TEXT/BLOB columns.
    if (usableDefault(col) && !demoted.has(col.name) && !/TEXT|BLOB/.test(mariaType(col))) {
      line += ` DEFAULT ${normaliseTemporalDefault(col, translateExpr(col.defaultDefinition, "maria"))}`;
    }
    lines.push(line);
  }

  if (table.pk.length) {
    lines.push(`  PRIMARY KEY (${table.pk.map(keyPart(demoted, table)).join(", ")})`);
  }
  if (identityCol && !leadsAKey) {
    lines.push(`  KEY ${q(`ai_${identityCol.name}`)} (${q(identityCol.name)})`);
  }
  for (const ck of table.checks) {
    lines.push(`  CONSTRAINT ${q(ck.name)} CHECK (${translateExpr(ck.definition, "maria")})`);
  }

  const ddl = [
    `CREATE TABLE ${q(table.name)} (\n${lines.join(",\n")}\n) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC;`,
  ];

  for (const idx of table.indexes) {
    const cols = idx.columns
      .map((c) => `${keyPart(demoted, table)(c.name)}${c.desc ? " DESC" : ""}`)
      .join(", ");
    ddl.push(
      `CREATE ${idx.unique ? "UNIQUE " : ""}INDEX ${mariaIdentifier(idx.name)} ON ${q(table.name)} (${cols});`,
    );
  }

  return ddl;
}

/** TEXT columns need a prefix length to be indexable in MariaDB. */
function keyPart(demoted, table) {
  return (colName) => {
    const col = table.columns.find((c) => c.name === colName);
    const isText =
      col && (demoted.has(colName) || /TEXT|BLOB/.test(mariaType(col, { forceText: false })));
    return isText ? `${q(colName)}(255)` : q(colName);
  };
}

module.exports = { emitPgTable, emitMariaTable, pgType, mariaType, translateExpr, q };
