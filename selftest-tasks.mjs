// 临时验证:造真实业务数据,跑通 5 个任务
import { writeFileSync, mkdirSync, existsSync, rmSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { writeXlsx, readXlsx } from "./lib/xlsx.mjs";
import { toCsv } from "./lib/csv.mjs";
import { loadTable } from "./lib/table.mjs";
import { merge, split, reconcile, clean, summarize } from "./lib/tasks.mjs";

let pass = 0, fail = 0;
const ok = (c, n, extra = "") => { if (c) { pass++; console.log(`  ✅ ${n}`); } else { fail++; console.log(`  ❌ ${n}${extra ? "  → " + extra : ""}`); } };

const T = join(process.cwd(), "_tmp");
if (existsSync(T)) rmSync(T, { recursive: true, force: true });
const SRC = join(T, "src"), OUT = join(T, "out");
mkdirSync(SRC, { recursive: true }); mkdirSync(OUT, { recursive: true });
const ctx = { cwd: process.cwd(), outDir: OUT };

/* ============ 造数据:订单系统(xlsx) vs 财务台账(csv)============ */
const 订单 = [
  ["订单号", "客户", "金额", "日期", "状态"],
  ["A001", "张三", 100.00, new Date(2026, 0, 5), "已付"],
  ["A002", "李四", 250.50, new Date(2026, 0, 6), "已付"],
  ["A003", "王五", 80.00, new Date(2026, 0, 7), "未付"],
  ["A004", "赵六", 1200.00, new Date(2026, 0, 8), "已付"],
  ["A005", "钱七", 300.00, new Date(2026, 0, 9), "已付"],
  ["A006", "孙八", 45.50, new Date(2026, 0, 10), "已付"],
  ["A007", "周九", 620.00, new Date(2026, 0, 11), "未付"],
  ["A008", "吴十", 150.00, new Date(2026, 0, 12), "已付"],
  ["A009", "郑一", 980.00, new Date(2026, 0, 13), "已付"],
  ["A010", "王二", 75.00, new Date(2026, 0, 14), "已付"],
];
writeFileSync(join(SRC, "订单系统.xlsx"), writeXlsx({ sheets: [{ name: "订单", rows: 订单 }] }));

// 台账:少了 A007/A010,多了 A011,A009 金额不一致
const 台账 = [
  ["订单号", "金额", "状态"],
  ["A001", 100.00, "已付"],
  ["A002", 250.50, "已付"],
  ["A003", 80.00, "未付"],
  ["A004", 1200.00, "已付"],
  ["A005", 300.00, "已付"],
  ["A006", 45.50, "已付"],
  ["A008", 150.00, "已付"],
  ["A009", 990.00, "已付"],     // ← 与订单表不一致(980)
  ["A011", 500.00, "已付"],     // ← 只有台账有
];
writeFileSync(join(SRC, "财务台账.csv"), toCsv(台账), "utf8");

// 三个月度表(测合并)
for (const [m, rows] of [["1月", 5], ["2月", 4], ["3月", 3]]) {
  const body = Array.from({ length: rows }, (_, i) => [`B${m}${i + 1}`, `客户${m}${i}`, (i + 1) * 111.11, `2026-0${m[0]}-1${i % 9}`, "已付"]);
  writeFileSync(join(SRC, `月度-${m}.csv`), toCsv([["订单号", "客户", "金额", "日期", "状态"], ...body]), "utf8");
}
console.log(`造好测试数据:${readdirSync(SRC).length} 个文件\n`);

/* ============ ① 合并 ============ */
console.log("① merge 合并");
const r1 = merge({ task: "merge", inputs: [join(SRC, "月度-*.csv")], output: join(OUT, "合并.xlsx"), options: { sourceColumn: "来源文件", dedupeBy: ["订单号"] } }, ctx);
ok(r1.summary["输入文件数"] === 3, "找到 3 个月度文件", String(r1.summary["输入文件数"]));
ok(r1.summary["合并后行数"] === 12, "合并后 12 行", String(r1.summary["合并后行数"]));
ok(r1.summary["列数"] === 6, "列数 6(含来源列)", String(r1.summary["列数"]));
const m1 = loadTable(join(OUT, "合并.xlsx"));
ok(m1.headers[0] === "来源文件", "来源列在最前");
ok(m1.rows.some((r) => r[0] === "月度-1月.csv"), "来源文件被记录");
ok(new Set(m1.rows.map((r) => r[1])).size === 12, "12 个订单号不重复");

/* ============ ② 拆分 ============ */
console.log("\n② split 拆分");
const r2 = split({ task: "split", inputs: [join(SRC, "订单系统.xlsx")], outputDir: join(OUT, "按状态拆分"), options: { by: ["状态"] } }, ctx);
ok(r2.summary["拆出文件数"] === 2, "按状态拆成 2 个文件", String(r2.summary["拆出文件数"]));
ok(existsSync(join(OUT, "按状态拆分", "已付.xlsx")), "生成「已付.xlsx」");
ok(existsSync(join(OUT, "按状态拆分", "未付.xlsx")), "生成「未付.xlsx」");
const paid = loadTable(join(OUT, "按状态拆分", "已付.xlsx"));
ok(paid.rows.length === 8, "已付 8 行", String(paid.rows.length));
const unpaid = loadTable(join(OUT, "按状态拆分", "未付.xlsx"));
ok(unpaid.rows.length === 2, "未付 2 行", String(unpaid.rows.length));

/* ============ ③ 对账(核心)============ */
console.log("\n③ reconcile 对账");
const r3 = reconcile({
  task: "reconcile",
  inputs: { a: join(SRC, "订单系统.xlsx"), b: join(SRC, "财务台账.csv") },
  output: join(OUT, "对账结果.xlsx"),
  options: { keyColumns: ["订单号"], compareColumns: ["金额", "状态"] },
}, ctx);
ok(r3.summary["左表行数"] === 10, "左表 10 行", String(r3.summary["左表行数"]));
ok(r3.summary["右表行数"] === 9, "右表 9 行", String(r3.summary["右表行数"]));
ok(r3.summary["仅左表有"] === 2, "仅左表有 2 条(A007/A010)", String(r3.summary["仅左表有"]));
ok(r3.summary["仅右表有"] === 1, "仅右表有 1 条(A011)", String(r3.summary["仅右表有"]));
ok(r3.summary["值不一致"] === 1, "值不一致 1 条(A009)", String(r3.summary["值不一致"]));
ok(r3.summary["完全一致"] === 7, "完全一致 7 条", String(r3.summary["完全一致"]));
ok(/4 处差异/.test(r3.summary["结论"]), "结论正确(2+1+1=4 处差异)", r3.summary["结论"]);

const rep = readXlsx(readFileSync(join(OUT, "对账结果.xlsx")));
ok(rep.sheets.map((s) => s.name).join(",") === "汇总,仅左表有,仅右表有,值不一致", "报告含 4 个工作表", rep.sheets.map((s) => s.name).join(","));
const onlyA = rep.sheets.find((s) => s.name === "仅左表有");
ok(onlyA.rows.length === 3, "仅左表有 sheet:表头 + 2 行", String(onlyA.rows.length));
ok(onlyA.rows.slice(1).map((r) => r[0]).sort().join(",") === "A007,A010", "左边独有的是 A007/A010", onlyA.rows.slice(1).map((r) => r[0]).join(","));
const diffSheet = rep.sheets.find((s) => s.name === "值不一致");
ok(diffSheet.rows.length === 2, "值不一致 sheet:表头 + 1 行");
ok(diffSheet.rows[1][0] === "A009", "不一致的是 A009");
ok(JSON.stringify(diffSheet.rows[1]).includes("980") && JSON.stringify(diffSheet.rows[1]).includes("990"), "差异明细同时给出左右两边的值");
const sum = rep.sheets.find((s) => s.name === "汇总").rows;
ok(sum.some((r) => r[0] === "结论" && /4 处差异/.test(String(r[1]))), "汇总表带结论文本");

/* ============ ④ 清洗 ============ */
console.log("\n④ clean 清洗");
writeFileSync(join(SRC, "脏数据.csv"), toCsv([
  ["姓名", "手机号", "金额", "备注"],
  ["  张三  ", "13800000001", "1,234.50", " 有空格 "],
  ["李四", "13800000002", "¥88", ""],
  ["张三", "13800000001", "1234.5", "重复行"],
  ["", "", "", ""],
  ["王五", "138-0000-0003", "abc", "金额非数字"],
  ["  ", "  ", "  ", " "],
]), "utf8");
const r4 = clean({
  task: "clean", inputs: [join(SRC, "脏数据.csv")], output: join(OUT, "清洗后.xlsx"),
  options: { trim: true, dropEmptyRows: true, dedupeBy: ["手机号"], columns: { 金额: "number" } },
}, ctx);
const c1 = r4.details[0];
ok(c1["原始行数"] === 6, "原始 6 行数据行", String(c1["原始行数"]));
ok(c1["删空行"] === 2, "删除 2 个空行", String(c1["删空行"]));
ok(c1["去重删除"] === 1, "删除 1 个重复手机号", String(c1["去重删除"]));
ok(c1["清洗后行数"] === 3, "清洗后 3 行", String(c1["清洗后行数"]));
const cleaned = loadTable(join(OUT, "清洗后.xlsx"));
ok(cleaned.rows[0][0] === "张三", "首尾空格被去掉", JSON.stringify(cleaned.rows[0][0]));
const amt = cleaned.rows.find((r) => r[0] === "张三")[2];
ok(typeof amt === "number" && Math.abs(amt - 1234.5) < 0.001, "带千分位的金额被转成数字", String(amt));
const amt2 = cleaned.rows.find((r) => r[0] === "李四")[2];
ok(typeof amt2 === "number" && amt2 === 88, "带货币符号的金额被转成数字", String(amt2));
ok(cleaned.rows.find((r) => r[0] === "王五")[2] === "abc", "非数字保持原样(不误删)");

/* ============ ⑤ 汇总 ============ */
console.log("\n⑤ summarize 汇总");
const r5 = summarize({
  task: "summarize", inputs: [join(SRC, "订单系统.xlsx")], output: join(OUT, "汇总.xlsx"),
  options: {
    groupBy: ["状态"],
    aggregations: [
      { column: "订单号", op: "count", as: "订单数" },
      { column: "金额", op: "sum", as: "金额合计" },
      { column: "金额", op: "avg", as: "平均金额" },
      { column: "客户", op: "distinct", as: "客户数" },
    ],
  },
}, ctx);
ok(r5.summary["分组数"] === 2, "分成 2 组(已付/未付)", String(r5.summary["分组数"]));
const agg = loadTable(join(OUT, "汇总.xlsx"));
ok(agg.headers.join(",") === "状态,订单数,金额合计,平均金额,客户数", "表头正确", agg.headers.join(","));
const paidRow = agg.rows.find((r) => r[0] === "已付");
ok(paidRow[1] === 8, "已付 8 单", String(paidRow[1]));
ok(Math.abs(paidRow[2] - (100 + 250.50 + 1200 + 300 + 45.50 + 150 + 980 + 75)) < 0.001, "已付金额合计正确", String(paidRow[2]));
ok(agg.rows[agg.rows.length - 1][0] === "合计", "最后一行是合计");
ok(agg.rows[agg.rows.length - 1][1] === 10, "合计 10 单", String(agg.rows[agg.rows.length - 1][1]));

console.log(`\n${"─".repeat(46)}`);
console.log(`  任务验证:${pass} 通过${fail ? `,${fail} 失败` : ""}`);
console.log(`${"─".repeat(46)}\n`);
if (!fail) rmSync(T, { recursive: true, force: true });
process.exit(fail ? 1 : 0);
