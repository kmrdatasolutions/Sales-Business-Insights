import * as XLSX from "xlsx";
import Papa from "papaparse";
import type { Dataset, Row, ColumnProfile } from "../types";

const cleanName = (value: string) =>
  value.trim().replace(/^\uFEFF/, "").replace(/\s+/g, " ").slice(0, 100);

const toRows = (headers: string[], values: unknown[][]): Row[] =>
  values
    .filter(row => row.some(v => v !== null && v !== undefined && String(v).trim() !== ""))
    .map(row => {
      const out: Row = {};
      headers.forEach((h, i) => {
        out[h] = row[i] ?? "";
      });
      return out;
    });

const parseDateValue = (value: unknown): number | null => {
  if (typeof value === "number" && value > 20000 && value < 60000) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return excelEpoch.getTime() + value * 86400000;
  }
  const s = String(value ?? "").trim();
  if (!s) return null;
  const t = Date.parse(s);
  return Number.isNaN(t) ? null : t;
};

const looksNumeric = (values: unknown[]) => {
  const nonEmpty = values.filter(v => v !== null && v !== undefined && String(v).trim() !== "");
  if (!nonEmpty.length) return false;
  const parsed = nonEmpty.map(v => Number(String(v).replace(/[$,%\s,]/g, "")));
  return parsed.filter(Number.isFinite).length / nonEmpty.length >= 0.85;
};

const looksDate = (values: unknown[]) => {
  const nonEmpty = values.filter(v => v !== null && v !== undefined && String(v).trim() !== "");
  if (!nonEmpty.length) return false;
  const parsed = nonEmpty.map(parseDateValue);
  return parsed.filter(v => v !== null).length / nonEmpty.length >= 0.75;
};

export function profileRows(rows: Row[]): ColumnProfile[] {
  if (!rows.length) return [];
  const names = Object.keys(rows[0]);
  return names.map(name => {
    const values = rows.map(r => r[name]);
    const numeric = looksNumeric(values);
    const date = !numeric && looksDate(values);
    const type: ColumnProfile["type"] =
      numeric ? "number" : date ? "date" : values.some(v => typeof v === "boolean") ? "boolean" : "string";
    const cleaned = values.filter(v => v !== null && v !== undefined && String(v).trim() !== "");
    const numbers = numeric
      ? cleaned.map(v => Number(String(v).replace(/[$,%\s,]/g, ""))).filter(Number.isFinite)
      : [];
    return {
      name,
      type,
      unique: new Set(cleaned.map(v => String(v))).size,
      nulls: values.length - cleaned.length,
      samples: cleaned.slice(0, 4).map(v => String(v)),
      ...(numeric && numbers.length ? {
        min: Math.min(...numbers),
        max: Math.max(...numbers),
        sum: numbers.reduce((a, b) => a + b, 0)
      } : {})
    };
  });
}

function asDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && value > 20000 && value < 60000) {
    const d = new Date(Date.UTC(1899, 11, 30));
    d.setUTCDate(d.getUTCDate() + value);
    return d;
  }
  const text = String(value ?? "").trim();
  if (!text) return null;
  const native = new Date(text);
  if (!Number.isNaN(native.getTime())) return native;

  const m = text.match(/^(\\d{1,2})[\\/-](\\d{1,2})[\\/-](\\d{4})$/);
  if (m) {
    const [, a, b, y] = m;
    const first = Number(a), second = Number(b);
    const month = first > 12 ? second : first;
    const day = first > 12 ? first : second;
    const d = new Date(Number(y), month - 1, day);
    return Number.isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function addDerivedFields(rows: Row[]): Row[] {
  if (!rows.length) return rows;
  const keys = Object.keys(rows[0]);
  const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, "");
  const find = (aliases: string[]) => keys.find(k => aliases.some(a => {
    const n = norm(k), x = norm(a);
    return n === x || n.includes(x);
  }));
  const orderDate = find(["Order Date", "Invoice Date", "Transaction Date", "Purchase Date"]);
  const shipDate = find(["Ship Date", "Shipment Date", "Delivery Date", "Delivered Date", "Dispatch Date"]);
  const existing = find(["Shipping Days", "Days to Ship", "Delivery Days", "Transit Days", "Ship Duration"]);
  if (existing || !orderDate || !shipDate) return rows;
  return rows.map(row => {
    const a = asDate(row[orderDate]);
    const b = asDate(row[shipDate]);
    if (!a || !b) return row;
    const days = Math.max(0, (b.getTime() - a.getTime()) / 86400000);
    return { ...row, "__Derived Shipping Days": Number(days.toFixed(2)) };
  });
}

function qualityFor(rows: Row[]) {
  const totalCells = rows.reduce((n, r) => n + Object.keys(r).length, 0);
  const missingCells = rows.reduce((n, r) => n + Object.values(r).filter(v => v === null || v === undefined || String(v).trim() === "").length, 0);
  const signatures = new Set<string>();
  let duplicateRows = 0;
  rows.forEach(r => {
    const sig = Object.values(r).map(v => String(v ?? "")).join("\u001f");
    if (signatures.has(sig)) duplicateRows++;
    signatures.add(sig);
  });
  return {
    duplicateRows,
    missingCells,
    totalCells,
    completeness: totalCells ? Math.round(((totalCells - missingCells) / totalCells) * 1000) / 10 : 100
  };
}

export function makeDataset(name: string, sourceType: Dataset["sourceType"], rows: Row[], meta?: { sheetCount?: number; sheets?: string[] }): Dataset {
  const enriched = addDerivedFields(rows);
  return {
    name,
    sourceType,
    rows: enriched,
    columns: profileRows(enriched),
    createdAt: new Date().toISOString(),
    ...meta,
    quality: qualityFor(enriched)
  };
}

export async function parseFile(file: File): Promise<Dataset> {
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    const sheetNames = workbook.SheetNames;
    const parsedSheets = sheetNames.map(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
      const headers = (matrix[0] ?? []).map(v => cleanName(String(v)));
      return { sheetName, rows: toRows(headers, matrix.slice(1)), headers };
    }).filter(s => s.rows.length);

    if (!parsedSheets.length) return makeDataset(file.name, "xlsx", []);
    // Combine sheets when their normalized header sets match; otherwise use the largest sheet.
    const base = parsedSheets[0];
    const normHeaders = (h: string[]) => h.map(x => x.toLowerCase().replace(/[^a-z0-9]/g, "")).sort().join("|");
    const compatible = parsedSheets.filter(s => normHeaders(s.headers) === normHeaders(base.headers));
    const selected = compatible.length > 1 ? compatible.flatMap(s => s.rows) :
      parsedSheets.sort((a, b) => b.rows.length - a.rows.length)[0].rows;
    return makeDataset(file.name, "xlsx", selected, {
      sheetCount: sheetNames.length,
      sheets: sheetNames
    });
  }

  const text = await file.text();
  if (lower.endsWith(".sql")) {
    const rows = parseSqlDump(text);
    return makeDataset(file.name, "sql", rows);
  }

  const delimiter = detectDelimiter(text);
  const result = Papa.parse<Record<string, unknown>>(text, {
    header: true,
    skipEmptyLines: true,
    delimiter
  });
  const rows = result.data.map(r => {
    const out: Row = {};
    Object.entries(r).forEach(([k, v]) => { out[cleanName(k)] = v; });
    return out;
  });
  return makeDataset(file.name, lower.endsWith(".txt") ? "txt" : "csv", rows);
}

function detectDelimiter(text: string): string {
  const first = text.split(/\r?\n/).find(Boolean) ?? "";
  const candidates = [",", "\t", ";", "|"];
  return candidates.sort((a, b) =>
    first.split(b).length - first.split(a).length
  )[0];
}

/**
 * Lightweight SQL dump reader.
 * Supports common INSERT INTO ... (col1, col2) VALUES (...), (...); statements.
 * It intentionally ignores DDL and other statements so users can upload exported SQL dumps.
 */
function parseSqlDump(sql: string): Row[] {
  const rows: Row[] = [];
  const insertRe = /INSERT\s+INTO\s+[`"']?([\w.\-]+)[`"']?\s*(?:\(([^)]*)\))?\s*VALUES\s*([\s\S]*?);/gi;
  let match: RegExpExecArray | null;

  while ((match = insertRe.exec(sql)) !== null) {
    const columns = match[2]
      ? splitSqlList(match[2]).map(v => v.replace(/[`"']/g, "").trim())
      : [];
    const tuples = extractTuples(match[3]);

    tuples.forEach(tuple => {
      const values = splitSqlList(tuple).map(unquoteSql);
      if (!columns.length) {
        columns.push(...values.map((_, i) => `Column ${i + 1}`));
      }
      const row: Row = {};
      columns.forEach((col, i) => { row[col] = values[i] ?? ""; });
      rows.push(row);
    });
  }
  return rows;
}

function extractTuples(block: string): string[] {
  const tuples: string[] = [];
  let depth = 0, quote = "", start = -1;
  for (let i = 0; i < block.length; i++) {
    const c = block[i];
    const prev = block[i - 1];
    if (quote) {
      if (c === quote && prev !== "\\") quote = "";
      continue;
    }
    if (c === "'" || c === '"') { quote = c; continue; }
    if (c === "(") {
      if (depth === 0) start = i + 1;
      depth++;
    } else if (c === ")") {
      depth--;
      if (depth === 0 && start >= 0) tuples.push(block.slice(start, i));
    }
  }
  return tuples;
}

function splitSqlList(input: string): string[] {
  const out: string[] = [];
  let start = 0, depth = 0, quote = "";
  for (let i = 0; i < input.length; i++) {
    const c = input[i], prev = input[i - 1];
    if (quote) {
      if (c === quote && prev !== "\\") quote = "";
      continue;
    }
    if (c === "'" || c === '"') { quote = c; continue; }
    if (c === "(") depth++;
    if (c === ")") depth--;
    if (c === "," && depth === 0) {
      out.push(input.slice(start, i).trim());
      start = i + 1;
    }
  }
  out.push(input.slice(start).trim());
  return out;
}

function unquoteSql(value: string): unknown {
  const v = value.trim();
  if (/^null$/i.test(v)) return "";
  if (/^true$/i.test(v)) return true;
  if (/^false$/i.test(v)) return false;
  if (/^'.*'$|^".*"$/s.test(v)) return v.slice(1, -1).replace(/''/g, "'");
  const n = Number(v);
  return Number.isFinite(n) ? n : v;
}
