export type Row = Record<string, unknown>;

export type ColumnProfile = {
  name: string;
  type: "number" | "date" | "boolean" | "string";
  unique: number;
  nulls: number;
  samples: string[];
  min?: number;
  max?: number;
  sum?: number;
};

export type Dataset = {
  name: string;
  sourceType: "xlsx" | "csv" | "txt" | "sql" | "demo";
  rows: Row[];
  columns: ColumnProfile[];
  createdAt: string;
  sheetCount?: number;
  sheets?: string[];
  quality?: {
    duplicateRows: number;
    missingCells: number;
    totalCells: number;
    completeness: number;
  };
};

export type Filters = {
  date: string;
  category: string;
  region: string;
  segment: string;
};

export type NumericMetric = {
  column: string;
  label: string;
  sum: number;
  avg: number;
};
