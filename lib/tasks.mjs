// 表格管家 · 核心任务
// 合并(merge)/ 拆分(split)/ 对账(reconcile)/ 清洗(clean)/ 汇总(summarize)
import { existsSync, readdirSync, mkdirSync, statSync, writeFileSync } from "node:fs";
import { join, dirname, basename, extname, resolve } from "node:path";
import { loadTable, saveTable, colIndex, unionHeaders, normValue, outPath, safeName, isXlsx } from "./table.mjs";
import { writeXlsx } from "./xlsx.mjs";

/* ============================ 输入解析 ============================ */
/** 支持:具体路径 / 目录 / 通配符(*.xlsx) */
export function resolveInputs(patterns, cwd = process.cwd()) {
  const out = [];
  for (const p of [].concat(patterns)) {
    const abs = resolve(cwd, p);
    if (existsSync(abs)) {
      const st = statSync(abs);
      if (st.isDirectory()) {
        for (const f of readdirSync(abs).sort()) if (/\.(xlsx|xlsm|csv|tsv)$/i.test(f)) out.push(join(abs, f));
      } else out.push(abs);
      continue;
    }
    // 通配符
    const dir = dirname(abs);
    const pat = basename(abs);
    if (existsSync(dir)) {
      const re = new RegExp("^" + pat.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i");
      for (const f of readdirSync(dir).sort()) if (re.test(f)) out.push(join(dir, f));
    }
  }
  if (!out.length) throw new Error(`没有找到任何输入文件:${[].concat(patterns).join(", ")}`);
  return out;
}

const ensureDir = (p) => { if (p && !existsSync(p)) mkdirSync(p, { recursive: true }); };

/* ============================ ① 合并 ============================ */
export function merge(cfg, ctx) {
  const files = resolveInputs(cfg.inputs, ctx.cwd);
  const opts = cfg.options ?? {};
  const tables = files.map((f) => ({ file: f, t: loadTable(f, { sheet: opts.sheet }) }));

  let headers = [];
  for (const { t } of tables) headers = unionHeaders(headers, t.headers);
  if (opts.sourceColumn) headers = [opts.sourceColumn, ...headers];

  const rows = [];
  const stats = [];
  for (const { file, t } of tables) {
    const idx = new Map(t.headers.map((h, i) => [h, i]));
    const added = t.rows.length;
    for (const r of t.rows) {
      const out = headers.map((h) => {
        if (h === opts.sourceColumn) return basename(file);
        const i = idx.get(h);
        return i == null ? null : (r[i] ?? null);
      });
      rows.push(out);
    }
    stats.push({ 文件: basename(file), 行数: added, 列数: t.headers.length, 编码: t.meta.encoding ?? "xlsx" });
  }

  // 去重
  let dup = 0;
  let finalRows = rows;
  if (opts.dedupeBy?.length) {
    const keys = opts.dedupeBy.map((k) => colIndex({ headers }, k));
    const seen = new Set();
    finalRows = rows.filter((r) => {
      const key = keys.map((i) => normValue(r[i])).join("\u0001");
      if (seen.has(key)) { dup++; return false; }
      seen.add(key); return true;
    });
  }

  const outFile = cfg.output ?? join(ctx.outDir, "合并结果.xlsx");
  ensureDir(dirname(outFile));
  saveTable(outFile, { headers, rows: finalRows, sheetName: opts.sheetName ?? "合并结果" });

  return {
    task: "merge",
    output: outFile,
    summary: { 输入文件数: files.length, 合并后行数: finalRows.length, 去重删除: dup, 列数: headers.length },
    details: stats,
  };
}

/* ============================ ② 拆分 ============================ */
export function split(cfg, ctx) {
  const files = resolveInputs(cfg.inputs, ctx.cwd);
  if (files.length > 1) throw new Error("拆分只接受一个输入文件(多个文件请先合并,或分别配置)");
  const opts = cfg.options ?? {};
  if (!opts.by?.length) throw new Error("split 需要 options.by(按哪一列拆分)");
  const t = loadTable(files[0], { sheet: opts.sheet });
  const keys = opts.by.map((c) => ({ name: c, i: colIndex(t, c) }));
  const outDir = cfg.outputDir ?? join(ctx.outDir, "拆分结果");
  ensureDir(outDir);
  const ext = opts.format ?? (isXlsx(files[0]) ? ".xlsx" : ".csv");

  const groups = new Map();
  for (const r of t.rows) {
    const name = keys.map((k) => String(r[k.i] ?? "空")).join("_") || "空";
    if (!groups.has(name)) groups.set(name, []);
    groups.get(name).push(r);
  }

  let entries = [...groups.entries()].sort((a, b) => b[1].length - a[1].length);
  let merged = null;
  if (opts.topN && entries.length > opts.topN) {
    merged = entries.slice(opts.topN);
    entries = entries.slice(0, opts.topN);
  }

  const details = [];
  for (const [name, rows] of entries) {
    const f = join(outDir, `${safeName(name)}${ext}`);
    saveTable(f, { headers: t.headers, rows, sheetName: safeName(name).slice(0, 31) });
    details.push({ 分组: name, 行数: rows.length, 文件: basename(f) });
  }
  if (merged) {
    const rows = merged.flatMap(([, r]) => r);
    const f = join(outDir, `其他${ext}`);
    saveTable(f, { headers: [...keys.map((k) => k.name), ...t.headers], rows: rows.map((r) => [...keys.map((k) => r[k.i]), ...r]), sheetName: "其他" });
    details.push({ 分组: `其他(${merged.length} 组)`, 行数: rows.length, 文件: basename(f) });
  }

  return {
    task: "split",
    output: outDir,
    summary: { 源文件: basename(files[0]), 总行数: t.rows.length, 拆出文件数: details.length, 分组依据: opts.by.join(" + ") },
    details,
  };
}

/* ============================ ③ 对账 ============================ */
export function reconcile(cfg, ctx) {
  const opts = cfg.options ?? {};
  const aPath = cfg.inputs?.a ?? cfg.inputs?.[0];
  const bPath = cfg.inputs?.b ?? cfg.inputs?.[1];
  if (!aPath || !bPath) throw new Error("reconcile 需要两个输入:inputs.a 与 inputs.b");
  const A = loadTable(resolve(ctx.cwd, aPath), { sheet: opts.sheetA });
  const B = loadTable(resolve(ctx.cwd, bPath), { sheet: opts.sheetB });

  const keyCols = opts.keyColumns ?? [A.headers[0]];
  const ai = keyCols.map((c) => colIndex(A, c));
  const bi = keyCols.map((c) => colIndex(B, c));

  // 比较列:默认两边同名列(排除键列)
  const compare = opts.compareColumns?.length
    ? opts.compareColumns.filter((c) => A.headers.includes(c) && B.headers.includes(c))
    : A.headers.filter((h) => B.headers.includes(h) && !keyCols.includes(h));
  if (!compare.length && !opts.compareColumns) {
    // 没有可比较的列:退化为"只比是否存在"
  }
  const acidx = compare.map((c) => A.headers.indexOf(c));
  const bcidx = compare.map((c) => B.headers.indexOf(c));

  const keyOf = (row, idx) => idx.map((i) => normValue(row[i])).join("\u0001");
  const bMap = new Map();
  B.rows.forEach((r, n) => {
    const k = keyOf(r, bi);
    if (!bMap.has(k)) bMap.set(k, []);
    bMap.get(k).push({ row: r, n });
  });
  const aMap = new Map();
  A.rows.forEach((r, n) => {
    const k = keyOf(r, ai);
    if (!aMap.has(k)) aMap.set(k, []);
    aMap.get(k).push({ row: r, n });
  });

  const onlyA = [], onlyB = [], diff = [], same = [];
  const detailHeaders = [...keyCols, ...compare.flatMap((c) => [`${c}(左)`, `${c}(右)`, `${c}·是否一致`]), "说明"];

  for (const [k, listA] of aMap) {
    const listB = bMap.get(k);
    if (!listB) {
      for (const { row } of listA) onlyA.push([...ai.map((i, n) => row[i]), ...compare.flatMap(() => [null, null, ""]), `${bPath} 中不存在`]);
      continue;
    }
    // 1 对 1 时逐字段比较;1 对多时只标注"重复"
    if (listA.length > 1 || listB.length > 1) {
      for (const { row } of listA) diff.push([...ai.map((i, n) => row[i]), ...compare.flatMap((c, n) => [row[acidx[n]], null, "键重复"]), `键在左表出现 ${listA.length} 次 / 右表 ${listB.length} 次,无法逐字段比对`]);
      continue;
    }
    const ra = listA[0].row, rb = listB[0].row;
    const cells = [], changed = [];
    compare.forEach((c, n) => {
      const va = ra[acidx[n]], vb = rb[bcidx[n]];
      const eq = normValue(va) === normValue(vb);
      cells.push(va, vb, eq ? "一致" : "不一致");
      if (!eq) changed.push(`${c}:${normValue(va) || "空"} ≠ ${normValue(vb) || "空"}`);
    });
    if (changed.length) diff.push([...ai.map((i, n) => ra[i]), ...cells, changed.join(";")]);
    else same.push([...ai.map((i, n) => ra[i]), ...cells, "完全一致"]);
  }
  for (const [k, listB] of bMap) {
    if (aMap.has(k)) continue;
    for (const { row } of listB) onlyB.push([...bi.map((i, n) => row[i]), ...compare.flatMap(() => [null, null, ""]), `${aPath} 中不存在`]);
  }

  const outFile = cfg.output ?? join(ctx.outDir, "对账结果.xlsx");
  ensureDir(dirname(outFile));
  const total = aMap.size + [...bMap.keys()].filter((k) => !aMap.has(k)).length;
  const summaryRows = [
    ["对账结果汇总", ""],
    ["生成时间", new Date().toLocaleString("zh-CN")],
    ["左表(A)", `${basename(aPath)}(${A.rows.length} 行)`],
    ["右表(B)", `${basename(bPath)}(${B.rows.length} 行)`],
    ["对账键", keyCols.join(" + ")],
    ["比较字段", compare.length ? compare.join(", ") : "(无,仅比对存在性)"],
    ["", ""],
    ["分类", "条数"],
    ["仅左表有(A 有 B 无)", onlyA.length],
    ["仅右表有(B 有 A 无)", onlyB.length],
    ["两边都有但值不一致", diff.length],
    ["完全一致", same.length],
    ["合计(去重后的键)", total],
    ["", ""],
    ["结论", (onlyA.length + onlyB.length + diff.length) === 0 ? "✅ 两边完全对得上" : `⚠ 有 ${onlyA.length + onlyB.length + diff.length} 处差异需要核对`],
  ];
  const sheets = [
    { name: "汇总", rows: summaryRows },
    { name: "仅左表有", rows: [detailHeaders, ...onlyA] },
    { name: "仅右表有", rows: [detailHeaders, ...onlyB] },
    { name: "值不一致", rows: [detailHeaders, ...diff] },
  ];
  if (opts.includeSame) sheets.push({ name: "完全一致", rows: [detailHeaders, ...same] });
  if (opts.format === ".csv" || extname(outFile).toLowerCase() === ".csv") {
    saveTable(outFile, { headers: detailHeaders, rows: [...onlyA, ...onlyB, ...diff], sheetName: "对账结果" });
  } else {
    writeFileSync(outFile, writeXlsx({ sheets }));
  }

  return {
    task: "reconcile",
    output: outFile,
    summary: {
      左表行数: A.rows.length, 右表行数: B.rows.length,
      仅左表有: onlyA.length, 仅右表有: onlyB.length,
      值不一致: diff.length, 完全一致: same.length,
      结论: (onlyA.length + onlyB.length + diff.length) === 0 ? "两边完全对得上" : `有 ${onlyA.length + onlyB.length + diff.length} 处差异`,
    },
    details: summaryRows.slice(8).map(([k, v]) => ({ 项: k, 值: v })),
  };
}

/* ============================ ④ 清洗 ============================ */
export function clean(cfg, ctx) {
  const files = resolveInputs(cfg.inputs, ctx.cwd);
  const opts = cfg.options ?? {};
  const results = [];
  const outputs = [];

  for (const f of files) {
    const t = loadTable(f, { sheet: opts.sheet, dropEmptyRows: false });   // 保留空行,好如实统计
    let rows = t.rows.slice();
    const log = { 文件: basename(f), 原始行数: rows.length };

    // 1) 去首尾空格 / 全角空格
    if (opts.trim !== false) {
      let n = 0;
      rows = rows.map((r) => r.map((v) => {
        if (typeof v !== "string") return v;
        const s = v.replace(/^[\s\u3000]+|[\s\u3000]+$/g, "");
        if (s !== v) n++;
        return s;
      }));
      log["去空格"] = n;
    }

    // 2) 删除空行
    if (opts.dropEmptyRows !== false) {
      const before = rows.length;
      rows = rows.filter((r) => r.some((c) => c != null && String(c).trim() !== ""));
      log["删空行"] = before - rows.length;
    }

    // 3) 列类型规范化
    if (opts.columns) {
      for (const [col, type] of Object.entries(opts.columns)) {
        const i = colIndex(t, col);
        let n = 0;
        rows = rows.map((r) => {
          const v = r[i];
          if (v == null || v === "") return r;
          let nv = v;
          if (type === "number") {
            const num = Number(String(v).replace(/[,¥$€£\s]/g, ""));
            if (!Number.isNaN(num)) nv = num;
          } else if (type === "date") {
            const d = new Date(String(v).replace(/\//g, "-"));
            if (!Number.isNaN(d.getTime())) nv = d;
          } else if (type === "string") nv = String(v).trim();
          if (nv !== v) n++;
          const out = r.slice(); out[i] = nv; return out;
        });
        log[`规范化·${col}=${type}`] = n;
      }
    }

    // 4) 去重
    if (opts.dedupeBy?.length || opts.dedupeWholeRow) {
      const idx = opts.dedupeWholeRow ? rows[0]?.map((_, i) => i) ?? [] : opts.dedupeBy.map((c) => colIndex(t, c));
      const seen = new Set();
      const before = rows.length;
      rows = rows.filter((r) => {
        const k = idx.map((i) => normValue(r[i])).join("\u0001");
        if (seen.has(k)) return false;
        seen.add(k); return true;
      });
      log["去重删除"] = before - rows.length;
    }

    // 5) 过滤(可选出)
    if (opts.filter) {
      const { column, op, value } = opts.filter;
      const i = colIndex(t, column);
      const before = rows.length;
      rows = rows.filter((r) => {
        const v = normValue(r[i]), target = normValue(value);
        if (op === "eq") return v === target;
        if (op === "neq") return v !== target;
        if (op === "gt") return Number(v) > Number(target);
        if (op === "lt") return Number(v) < Number(target);
        if (op === "contains") return String(r[i] ?? "").includes(String(value));
        if (op === "empty") return v === "";
        if (op === "notEmpty") return v !== "";
        return true;
      });
      log[`过滤·${column} ${op} ${value}`] = before - rows.length;
    }

    log["清洗后行数"] = rows.length;
    const outFile = cfg.output ?? outPath(ctx.outDir, f, "-清洗后", isXlsx(f) ? ".xlsx" : ".csv");
    ensureDir(dirname(outFile));
    saveTable(outFile, { headers: t.headers, rows, sheetName: "清洗后" });
    outputs.push(outFile);
    results.push(log);
  }

  return {
    task: "clean",
    output: outputs.length === 1 ? outputs[0] : outputs,
    summary: { 处理文件数: files.length },
    details: results,
  };
}

/* ============================ ⑤ 汇总 ============================ */
export function summarize(cfg, ctx) {
  const files = resolveInputs(cfg.inputs, ctx.cwd);
  const opts = cfg.options ?? {};
  if (!opts.groupBy?.length) throw new Error("summarize 需要 options.groupBy");
  if (!opts.aggregations?.length) throw new Error("summarize 需要 options.aggregations");

  const t = loadTable(files[0], { sheet: opts.sheet });
  const gi = opts.groupBy.map((c) => colIndex(t, c));
  const aggs = opts.aggregations.map((a) => ({ ...a, i: colIndex(t, a.column) }));

  const groups = new Map();
  for (const r of t.rows) {
    const key = gi.map((i) => normValue(r[i])).join("\u0001");
    if (!groups.has(key)) groups.set(key, { keyVals: gi.map((i) => r[i]), rows: [] });
    groups.get(key).rows.push(r);
  }

  const headers = [...opts.groupBy, ...aggs.map((a) => a.as ?? `${a.column}_${a.op}`)];
  const out = [];
  for (const { keyVals, rows } of groups.values()) {
    const cells = [...keyVals];
    for (const a of aggs) {
      const nums = rows.map((r) => r[a.i]).filter((v) => v != null && v !== "" && !Number.isNaN(Number(v))).map(Number);
      if (a.op === "count") cells.push(rows.length);
      else if (a.op === "sum") cells.push(Number(nums.reduce((x, y) => x + y, 0).toFixed(6)));
      else if (a.op === "avg") cells.push(nums.length ? Number((nums.reduce((x, y) => x + y, 0) / nums.length).toFixed(6)) : null);
      else if (a.op === "max") cells.push(nums.length ? Math.max(...nums) : null);
      else if (a.op === "min") cells.push(nums.length ? Math.min(...nums) : null);
      else if (a.op === "distinct") cells.push(new Set(rows.map((r) => normValue(r[a.i]))).size);
      else throw new Error(`不支持的聚合方式: ${a.op}(可用 sum/avg/count/max/min/distinct)`);
    }
    out.push(cells);
  }
  out.sort((a, b) => {
    const la = a[a.length - 1], lb = b[b.length - 1];
    return typeof lb === "number" && typeof la === "number" ? lb - la : String(a[0]).localeCompare(String(b[0]), "zh");
  });

  // 合计行
  const totalRow = ["合计", ...aggs.map((a, n) => {
    if (a.op === "count") return t.rows.length;
    const nums = t.rows.map((r) => r[a.i]).filter((v) => v != null && v !== "" && !Number.isNaN(Number(v))).map(Number);
    if (a.op === "sum") return Number(nums.reduce((x, y) => x + y, 0).toFixed(6));
    if (a.op === "avg") return nums.length ? Number((nums.reduce((x, y) => x + y, 0) / nums.length).toFixed(6)) : null;
    if (a.op === "max") return nums.length ? Math.max(...nums) : null;
    if (a.op === "min") return nums.length ? Math.min(...nums) : null;
    if (a.op === "distinct") return new Set(t.rows.map((r) => normValue(r[a.i]))).size;
    return null;
  })];
  for (let i = 2; i < headers.length; i++) totalRow[i] = "";

  const outFile = cfg.output ?? join(ctx.outDir, "汇总结果.xlsx");
  ensureDir(dirname(outFile));
  saveTable(outFile, { headers, rows: [...out, totalRow], sheetName: "汇总" });

  return {
    task: "summarize",
    output: outFile,
    summary: { 源文件: basename(files[0]), 源行数: t.rows.length, 分组数: groups.size, 分组依据: opts.groupBy.join(" + ") },
    details: out.slice(0, 20).map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i]]))),
  };
}

export const TASKS = { merge, split, reconcile, clean, summarize };
