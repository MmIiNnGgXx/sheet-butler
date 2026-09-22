// 统一表格模型 —— CSV 与 XLSX 用同一套结构处理,任务逻辑只写一遍
//
// 表格结构:{ headers: string[], rows: any[][], meta: { source, sheet, encoding } }
import { readFileSync, writeFileSync, existsSync, mkdirSync } from "node:fs";
import { extname, dirname, basename, join } from "node:path";
import { readXlsx, writeXlsx } from "./xlsx.mjs";
import { decodeText, parseCsv, toCsv, inferCell } from "./csv.mjs";

export const EXT_CSV = [".csv", ".txt", ".tsv"];
export const EXT_XLSX = [".xlsx", ".xlsm"];
export const isCsv = (p) => EXT_CSV.includes(extname(p).toLowerCase());
export const isXlsx = (p) => EXT_XLSX.includes(extname(p).toLowerCase());

/** 表头规范化:去首尾空格、去 BOM、连续空白压成一个空格;空表头起名 列N */
export function normalizeHeaders(raw) {
  const seen = new Map();
  return raw.map((h, i) => {
    let name = String(h ?? "").replace(/^\uFEFF/, "").trim().replace(/\s+/g, " ");
    if (!name) name = `列${i + 1}`;
    const n = (seen.get(name) ?? 0) + 1;
    seen.set(name, n);
    return n > 1 ? `${name}_${n}` : name;   // 重名自动加后缀,避免覆盖
  });
}

/**
 * 读取任意表格文件
 * @param {string} path
 * @param {{sheet?:string|number, headerRow?:number, infer?:boolean}} opts
 */
export function loadTable(path, opts = {}) {
  if (!existsSync(path)) throw new Error(`文件不存在: ${path}`);
  const headerRow = opts.headerRow ?? 0;
  // 默认丢掉空行(合并/汇总时不需要它们);clean 任务会传 false,以便如实统计"删了多少空行"
  const dropEmpty = opts.dropEmptyRows !== false;
  const isEmptyRow = (r) => !r.some((c) => c != null && String(c).trim() !== "");

  if (isXlsx(path)) {
    const wb = readXlsx(readFileSync(path));
    let sheet = wb.sheets[0];
    if (opts.sheet != null) {
      sheet = typeof opts.sheet === "number" ? wb.sheets[opts.sheet] : wb.sheets.find((s) => s.name === opts.sheet);
      if (!sheet) throw new Error(`找不到工作表「${opts.sheet}」,可用:${wb.sheets.map((s) => s.name).join(" / ")}`);
    }
    const all = sheet.rows;
    const headers = normalizeHeaders(all[headerRow] ?? []);
    const rows = all.slice(headerRow + 1).map((r) => {
      const out = new Array(headers.length).fill(null);
      for (let i = 0; i < headers.length; i++) out[i] = r[i] === undefined ? null : r[i];
      return out;
    }).filter((r) => !dropEmpty || !isEmptyRow(r));
    return { headers, rows, meta: { source: path, sheet: sheet.name, encoding: null, sheets: wb.sheets.map((s) => s.name) } };
  }

  if (isCsv(path)) {
    const { text, encoding } = decodeText(readFileSync(path));
    const delim = extname(path).toLowerCase() === ".tsv" ? "\t" : undefined;
    const all = parseCsv(text, delim ? { delimiter: delim } : {});
    const headers = normalizeHeaders(all[headerRow] ?? []);
    const infer = opts.infer !== false;
    const rows = all.slice(headerRow + 1).map((r) => {
      const out = new Array(headers.length).fill(null);
      for (let i = 0; i < headers.length; i++) out[i] = r[i] === undefined ? null : (infer ? inferCell(r[i]) : r[i]);
      return out;
    }).filter((r) => !dropEmpty || !isEmptyRow(r));
    return { headers, rows, meta: { source: path, sheet: null, encoding } };
  }

  throw new Error(`不支持的文件类型: ${extname(path)}(支持 .xlsx / .csv / .tsv)`);
}

/** 写出表格(按扩展名决定格式) */
export function saveTable(path, table) {
  const dir = dirname(path);
  if (dir && !existsSync(dir)) mkdirSync(dir, { recursive: true });
  const all = [table.headers, ...table.rows];
  if (isXlsx(path)) writeFileSync(path, writeXlsx({ sheets: [{ name: table.sheetName ?? "Sheet1", rows: all }] }));
  else writeFileSync(path, toCsv(all, { delimiter: extname(path).toLowerCase() === ".tsv" ? "\t" : "," }), "utf8");
  return path;
}

/* ---------------- 结构转换 ---------------- */
export const recordsOf = (t) => t.rows.map((r) => Object.fromEntries(t.headers.map((h, i) => [h, r[i] ?? null])));
export const tableOf = (headers, records, meta = {}) => ({
  headers,
  rows: records.map((o) => headers.map((h) => (o[h] === undefined ? null : o[h]))),
  meta,
});
export const colIndex = (t, name) => {
  const i = t.headers.indexOf(name);
  if (i < 0) throw new Error(`找不到列「${name}」。可用列:${t.headers.join(" / ")}`);
  return i;
};

/** 取两个表列的并集(保持顺序,先 A 后 B 的新增列) */
export function unionHeaders(a, b) {
  const out = [...a];
  for (const h of b) if (!out.includes(h)) out.push(h);
  return out;
}

/** 值归一化:用于对账比较(去空格、数字统一、日期统一成 yyyy-mm-dd) */
export function normValue(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) {
    const p = (x) => String(x).padStart(2, "0");
    return `${v.getFullYear()}-${p(v.getMonth() + 1)}-${p(v.getDate())}`;
  }
  if (typeof v === "number") return String(Number(v.toFixed(6)));
  const s = String(v).trim();
  const n = Number(s.replace(/[,¥$€£\s]/g, ""));
  if (s !== "" && !Number.isNaN(n) && /\d/.test(s) && !/[a-zA-Z\u4e00-\u9fa5]/.test(s.replace(/[eE]/g, ""))) {
    return String(Number(n.toFixed(6)));
  }
  return s;
}

/** 生成输出文件名:原名 + 后缀,放在 outDir */
export function outPath(outDir, srcPath, suffix, ext) {
  const base = basename(srcPath, extname(srcPath));
  return join(outDir, `${base}${suffix}${ext ?? extname(srcPath)}`);
}

/** 文件名安全化(去掉 Windows 不允许的字符) */
export const safeName = (s) => String(s).replace(/[\\/:*?"<>|\r\n\t]/g, "_").replace(/\s+/g, " ").trim().slice(0, 60) || "未命名";
