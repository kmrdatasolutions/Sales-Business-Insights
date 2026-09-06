import type { ColumnProfile, Dataset, Filters, Row } from "../types";

const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, "");


function toDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && value > 20000 && value < 60000) {
    const d = new Date(Date.UTC(1899, 11, 30));
    d.setUTCDate(d.getUTCDate() + value);
    return d;
  }
  const text = String(value ?? "").trim();
  if (!text) return null;
  const d = new Date(text);
  return Number.isNaN(d.getTime()) ? null : d;
}

const aliasScore = (column: string, aliases: string[]) => {
  const n = normalize(column);
  let score = 0;
  for (const alias of aliases) {
    const a = normalize(alias);
    if (!a) continue;
    if (n === a) score = Math.max(score, 100);
    else if (n.startsWith(a) || n.endsWith(a)) score = Math.max(score, 80);
    else if (n.includes(a)) score = Math.max(score, 65);
    else if (a.includes(n) && n.length >= 4) score = Math.max(score, 45);
  }
  return score;
};

export function findColumn(dataset: Dataset, aliases: string[], type?: ColumnProfile["type"]) {
  const candidates = dataset.columns
    .filter(c => !type || c.type === type)
    .map(c => ({ name: c.name, score: aliasScore(c.name, aliases) }))
    .filter(c => c.score > 0)
    .sort((a, b) => b.score - a.score);
  return candidates[0]?.name;
}

function fallbackNumeric(dataset: Dataset, exclude: string[] = []) {
  return dataset.columns
    .filter(c => c.type === "number" && !exclude.some(x => normalize(x) === normalize(c.name)))
    .map(c => {
      const n = normalize(c.name);
      let score = 0;
      if (/(sales|revenue|amount|value|price|profit|income|cost|margin)/.test(n)) score += 50;
      if (/(id|code|zip|postal|year|month|day|rate|percent|discount)/.test(n)) score -= 35;
      score += Math.min(c.unique / Math.max(dataset.rows.length, 1), 1) * 10;
      return { name: c.name, score };
    })
    .sort((a, b) => b.score - a.score)[0]?.name;
}

function fallbackCategory(dataset: Dataset, exclude: string[] = []) {
  return dataset.columns
    .filter(c => c.type === "string" && c.unique > 1 && c.unique <= Math.min(50, Math.max(10, dataset.rows.length * 0.4)))
    .filter(c => !exclude.some(x => normalize(x) === normalize(c.name)))
    .map(c => {
      const n = normalize(c.name);
      let score = 0;
      if (/(region|state|country|market|territory|segment|category|department|product|customer|client|city)/.test(n)) score += 60;
      if (/(id|code|email|phone|address)/.test(n)) score -= 40;
      return { name: c.name, score };
    })
    .sort((a, b) => b.score - a.score)[0]?.name;
}

function fallbackDate(dataset: Dataset) {
  return dataset.columns.find(c => c.type === "date")?.name;
}

export function numericColumns(dataset: Dataset) {
  return dataset.columns.filter(c => c.type === "number").map(c => c.name);
}

export function categoricalColumns(dataset: Dataset) {
  return dataset.columns
    .filter(c => c.type === "string" && c.unique > 1 && c.unique <= Math.min(50, Math.max(10, dataset.rows.length * 0.4)))
    .map(c => c.name);
}

export function dateColumns(dataset: Dataset) {
  return dataset.columns.filter(c => c.type === "date").map(c => c.name);
}

export function applyFilters(dataset: Dataset, filters: Filters): Row[] {
  const dateCol = findColumn(dataset, ["Order Date", "Invoice Date", "Transaction Date", "Date", "Ship Date", "Delivery Date", "OrderDate"]);
  const regionCol = findColumn(dataset, ["Region", "State", "Province", "Country", "Territory", "Market", "Geography", "Location"]);
  const segmentCol = findColumn(dataset, ["Segment", "Customer Segment", "Client Segment", "Customer Type", "Department"]);
  const categoryCol = findColumn(dataset, ["Category", "Product Category", "Product Type", "Department", "Class"]);

  return dataset.rows.filter(row => {
    if (filters.date && dateCol) {
      const date = toDate(row[dateCol]);
      if (date && String(date.getFullYear()) !== filters.date) return false;
    }
    if (filters.region && regionCol && String(row[regionCol]) !== filters.region) return false;
    if (filters.segment && segmentCol && String(row[segmentCol]) !== filters.segment) return false;
    if (filters.category && categoryCol && String(row[categoryCol]) !== filters.category) return false;
    return true;
  });
}

export function sumBy(rows: Row[], column?: string) {
  if (!column) return 0;
  return rows.reduce((sum, row) => {
    const n = Number(String(row[column] ?? "").replace(/[$,%\s,]/g, ""));
    return sum + (Number.isFinite(n) ? n : 0);
  }, 0);
}

export function avgBy(rows: Row[], column?: string) {
  if (!column) return 0;
  const nums = rows.map(r => Number(String(r[column] ?? "").replace(/[$,%\s,]/g, ""))).filter(Number.isFinite);
  return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
}

export function groupSum(rows: Row[], groupColumn?: string, valueColumn?: string, limit = 8) {
  if (!groupColumn) return [];
  const map = new Map<string, number>();
  rows.forEach(row => {
    const key = String(row[groupColumn] ?? "Unknown");
    const value = valueColumn ? Number(String(row[valueColumn] ?? "").replace(/[$,%\s,]/g, "")) : 1;
    map.set(key, (map.get(key) ?? 0) + (Number.isFinite(value) ? value : 0));
  });
  return [...map.entries()].sort((a, b) => b[1] - a[1]).slice(0, limit).map(([name, value]) => ({ name, value }));
}

export function groupAvg(rows: Row[], groupColumn?: string, valueColumn?: string, limit = 10) {
  if (!groupColumn || !valueColumn) return [];
  const map = new Map<string, { sum: number; count: number }>();
  rows.forEach(row => {
    const key = String(row[groupColumn] ?? "Unknown");
    const value = Number(String(row[valueColumn] ?? "").replace(/[$,%\s,]/g, ""));
    if (!Number.isFinite(value)) return;
    const current = map.get(key) ?? { sum: 0, count: 0 };
    current.sum += value; current.count++;
    map.set(key, current);
  });
  return [...map.entries()].map(([name, v]) => ({ name, value: v.sum / v.count }))
    .sort((a, b) => b.value - a.value).slice(0, limit);
}

export function trend(rows: Row[], dateColumn?: string, valueColumn?: string) {
  if (!dateColumn || !valueColumn) return [];
  const map = new Map<string, number>();
  rows.forEach(row => {
    const date = toDate(row[dateColumn]);
    const value = Number(String(row[valueColumn] ?? "").replace(/[$,%\s,]/g, ""));
    if (!date || !Number.isFinite(value)) return;
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    map.set(key, (map.get(key) ?? 0) + value);
  });
  return [...map.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([name, value]) => ({ name, value }));
}

export function uniqueValues(rows: Row[], column?: string, limit = 30) {
  if (!column) return [];
  return [...new Set(rows.map(r => String(r[column] ?? "")).filter(Boolean))].sort().slice(0, limit);
}

export function currency(value: number, compact = true) {
  if (!Number.isFinite(value)) return "$0";
  if (compact) return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", notation: "compact", maximumFractionDigits: 1 }).format(value);
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(value);
}

export function number(value: number) {
  return new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: 1 }).format(value);
}

export function inferMetrics(dataset: Dataset) {
  const sales = findColumn(dataset, ["Sales", "Revenue", "Net Sales", "Total Sales", "Amount", "Turnover"], "number")
    ?? fallbackNumeric(dataset);
  const profit = findColumn(dataset, ["Profit", "Net Profit", "Gross Profit", "Operating Profit", "Earnings"], "number");
  const quantity = findColumn(dataset, ["Quantity", "Units", "Units Sold", "Qty", "Volume"], "number");
  const shipping = findColumn(dataset, ["Shipping Days", "Derived Shipping Days", "Days to Ship", "Ship Duration", "Delivery Days", "Transit Days", "Lead Time"], "number");

  const region = findColumn(dataset, ["Region", "State", "Province", "Country", "Territory", "Market", "Geography", "Location"])
    ?? fallbackCategory(dataset);
  const segment = findColumn(dataset, ["Segment", "Customer Segment", "Client Segment", "Customer Type", "Department"]);
  const category = findColumn(dataset, ["Category", "Product Category", "Product Type", "Department", "Class"]);
  const product = findColumn(dataset, ["Product Name", "Product", "SKU", "Item", "Model"]);
  const customer = findColumn(dataset, ["Customer Name", "Customer", "Client Name", "Client", "Account"]);
  const orderId = findColumn(dataset, ["Order ID", "OrderID", "Order Number", "Invoice ID", "Transaction ID"]);
  const date = findColumn(dataset, ["Order Date", "Invoice Date", "Transaction Date", "Date", "Purchase Date", "Created Date"]);
  const shipDate = findColumn(dataset, ["Ship Date", "Shipment Date", "Delivery Date", "Delivered Date", "Dispatch Date"]);

  return { sales, profit, quantity, shipping, orderId, customer, product, category, region, segment, date, shipDate };
}
