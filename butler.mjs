#!/usr/bin/env node
// 表格管家 —— Excel / CSV 批量处理
//
// 用法:
//   node butler.mjs                      # 按 config.json 执行
//   node butler.mjs --config=my.json     # 指定配置
//   node butler.mjs --task=merge         # 临时覆盖任务类型
//   node butler.mjs --json               # 附带机器可读结果
//
// 支持的任务:merge(合并)/ split(拆分)/ reconcile(对账)/ clean(清洗)/ summarize(汇总)
// 退出码:0 成功 / 1 失败

import { readFileSync, existsSync, mkdirSync } from "node:fs";
import { join, dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { TASKS } from "./lib/tasks.mjs";

const ROOT = dirname(fileURLToPath(import.meta.url));
const argOf = (k) => process.argv.find((a) => a.startsWith(`--${k}=`))?.slice(k.length + 3);
const cfgPath = resolve(argOf("config") ?? join(ROOT, "config.json"));

if (!existsSync(cfgPath)) {
  console.error(`\n  [错误] 找不到配置文件:${cfgPath}\n`);
  process.exit(1);
}

let cfg;
try { cfg = JSON.parse(readFileSync(cfgPath, "utf8")); }
catch (e) { console.error(`\n  [错误] 配置文件不是合法 JSON:${e.message}\n`); process.exit(1); }

const task = argOf("task") ?? cfg.task;
const fn = TASKS[task];
if (!fn) {
  console.error(`\n  [错误] 不支持的任务「${task}」`);
  console.error(`         可用:${Object.keys(TASKS).join(" / ")}\n`);
  process.exit(1);
}

const outDir = resolve(ROOT, cfg.outDir ?? "out");
if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });

console.log("");
console.log("  ==================================================");
console.log(`    表格管家  ·  ${task}`);
console.log("  ==================================================");
console.log(`    配置   :  ${cfgPath}`);
console.log(`    输出到 :  ${outDir}`);
console.log("");

const t0 = Date.now();
let result;
try {
  result = await fn(cfg, { cwd: ROOT, outDir });
} catch (e) {
  console.error(`  [失败] ${e.message}\n`);
  if (process.argv.includes("--stack")) console.error(e.stack);
  process.exit(1);
}
const ms = Date.now() - t0;

/* ---------------- 输出摘要 ---------------- */
const width = (s) => [...String(s)].reduce((a, c) => a + (c.charCodeAt(0) > 255 ? 2 : 1), 0);
const padTo = (s, n) => String(s) + " ".repeat(Math.max(0, n - width(s)));
const line = (k, v) => console.log(`    ${padTo(k, 18)} ${v}`);

console.log("  ── 结果 ──────────────────────────────────────────");
for (const [k, v] of Object.entries(result.summary ?? {})) line(k, v);
console.log("");

if (result.details?.length) {
  console.log("  ── 明细 ──────────────────────────────────────────");
  const cols = [...new Set(result.details.flatMap((d) => Object.keys(d)))];
  const width = (s) => [...String(s)].reduce((a, c) => a + (c.charCodeAt(0) > 255 ? 2 : 1), 0);
  const pad = (s, n) => String(s) + " ".repeat(Math.max(0, n - width(s)));
  const w = cols.map((c) => Math.max(width(c), ...result.details.map((d) => width(d[c] ?? ""))) + 2);
  console.log("    " + cols.map((c, i) => pad(c, w[i])).join(""));
  console.log("    " + w.map((n) => "─".repeat(n - 1) + " ").join(""));
  for (const d of result.details.slice(0, 30)) {
    console.log("    " + cols.map((c, i) => pad(d[c] ?? "", w[i])).join(""));
  }
  if (result.details.length > 30) console.log(`    …(共 ${result.details.length} 条,只显示前 30 条)`);
  console.log("");
}

console.log(`  输出: ${Array.isArray(result.output) ? result.output.join("\n        ") : result.output}`);
console.log(`  耗时: ${(ms / 1000).toFixed(2)}s`);
console.log("");

if (process.argv.includes("--json")) {
  console.log(JSON.stringify({ task, ms, ...result }, null, 2));
}

process.exit(0);
