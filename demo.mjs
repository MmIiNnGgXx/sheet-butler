#!/usr/bin/env node
// 生成示例数据,让你不用自己准备文件就能立刻试这个工具
//   node demo.mjs
// 会在 demo/ 下生成一套"订单系统 vs 财务台账"的对账场景 + 配置,然后直接告诉你下一步命令
import { mkdirSync, writeFileSync, existsSync, rmSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { writeXlsx } from "./lib/xlsx.mjs";
import { toCsv } from "./lib/csv.mjs";

const ROOT = dirname(fileURLToPath(import.meta.url));
const D = join(ROOT, "demo");
if (existsSync(D)) rmSync(D, { recursive: true, force: true });
mkdirSync(D, { recursive: true });

/* ---------- 订单系统(导出成 xlsx)---------- */
const 订单 = [
  ["订单号", "客户", "金额", "日期", "状态"],
  ["A001", "张三", 100.0, new Date(2026, 0, 5), "已付"],
  ["A002", "李四", 250.5, new Date(2026, 0, 6), "已付"],
  ["A003", "王五", 80.0, new Date(2026, 0, 7), "未付"],
  ["A004", "赵六", 1200.0, new Date(2026, 0, 8), "已付"],
  ["A005", "钱七", 300.0, new Date(2026, 0, 9), "已付"],
  ["A006", "孙八", 45.5, new Date(2026, 0, 10), "已付"],
  ["A007", "周九", 620.0, new Date(2026, 0, 11), "未付"],
  ["A008", "吴十", 150.0, new Date(2026, 0, 12), "已付"],
  ["A009", "郑一", 980.0, new Date(2026, 0, 13), "已付"],
  ["A010", "王二", 75.0, new Date(2026, 0, 14), "已付"],
  ["A011", "陈三", 430.0, new Date(2026, 0, 15), "已付"],
  ["A012", "褚四", 88.0, new Date(2026, 0, 16), "未付"],
];
writeFileSync(join(D, "订单系统导出.xlsx"), writeXlsx({ sheets: [{ name: "订单", rows: 订单 }] }));

/* ---------- 财务台账(手工维护的 csv,故意有几处对不上)---------- */
const 台账 = [
  ["订单号", "金额", "状态"],
  ["A001", 100.0, "已付"],
  ["A002", 250.5, "已付"],
  ["A003", 80.0, "未付"],
  ["A004", 1200.0, "已付"],
  ["A005", 300.0, "已付"],
  ["A006", 45.5, "已付"],
  ["A008", 150.0, "已付"],
  ["A009", 990.0, "已付"],     // ← 与订单表不一致(订单是 980)
  ["A011", 430.0, "已付"],
  ["A012", 88.0, "未付"],
  ["A999", 500.0, "已付"],     // ← 台账里多出来的
  // A007 / A010 台账里漏记
];
writeFileSync(join(D, "财务台账.csv"), toCsv(台账), "utf8");

/* ---------- 三个月度报表(测合并)---------- */
[["01", 5], ["02", 4], ["03", 3]].forEach(([m, n]) => {
  const rows = Array.from({ length: n }, (_, i) => [
    `B${m}${String(i + 1).padStart(2, "0")}`, `客户${m}${i + 1}`,
    Number(((i + 1) * 137.77).toFixed(2)), new Date(2026, Number(m) - 1, 10 + i), i % 3 === 0 ? "未付" : "已付",
  ]);
  writeFileSync(join(D, `月度报表-${m}月.csv`), toCsv([["订单号", "客户", "金额", "日期", "状态"], ...rows]), "utf8");
});

/* ---------- 脏数据(测清洗)---------- */
writeFileSync(join(D, "客户名单-脏数据.csv"), toCsv([
  ["姓名", "手机号", "消费金额", "备注"],
  ["  张三  ", "13800000001", "1,234.50", " 带空格 "],
  ["李四", "13800000002", "¥88", ""],
  ["张三", "13800000001", "1234.5", "重复记录"],
  ["", "", "", ""],
  ["王五", "138-0000-0003", "待确认", "金额非数字"],
]), "utf8");

/* ---------- 五份现成配置 ---------- */
const cfg = {
  reconcile: {
    task: "reconcile",
    inputs: { a: "demo/订单系统导出.xlsx", b: "demo/财务台账.csv" },
    output: "out/对账结果.xlsx",
    options: { keyColumns: ["订单号"], compareColumns: ["金额", "状态"] },
  },
  merge: {
    task: "merge",
    inputs: ["demo/月度报表-*.csv"],
    output: "out/合并结果.xlsx",
    options: { sourceColumn: "来源文件", dedupeBy: ["订单号"] },
  },
  split: {
    task: "split",
    inputs: ["demo/订单系统导出.xlsx"],
    outputDir: "out/按状态拆分",
    options: { by: ["状态"], format: ".xlsx" },
  },
  clean: {
    task: "clean",
    inputs: ["demo/客户名单-脏数据.csv"],
    output: "out/清洗后.xlsx",
    options: { trim: true, dropEmptyRows: true, dedupeBy: ["手机号"], columns: { 消费金额: "number" } },
  },
  summarize: {
    task: "summarize",
    inputs: ["demo/订单系统导出.xlsx"],
    output: "out/汇总结果.xlsx",
    options: {
      groupBy: ["状态"],
      aggregations: [
        { column: "订单号", op: "count", as: "订单数" },
        { column: "金额", op: "sum", as: "金额合计" },
        { column: "金额", op: "avg", as: "平均金额" },
        { column: "客户", op: "distinct", as: "客户数" },
      ],
    },
  },
};
for (const [k, v] of Object.entries(cfg)) {
  writeFileSync(join(D, `config-${k}.json`), JSON.stringify(v, null, 2) + "\n", "utf8");
}
// 默认配置 = 对账(最常用的场景)
writeFileSync(join(ROOT, "config.json"), JSON.stringify(cfg.reconcile, null, 2) + "\n", "utf8");

console.log(`
  ✅ 示例数据已生成到 demo/

     订单系统导出.xlsx      12 条订单(Excel 格式)
     财务台账.csv           手工维护的台账(故意有 4 处对不上)
     月度报表-01/02/03月.csv  用于测试合并
     客户名单-脏数据.csv      用于测试清洗

     还生成了 5 份现成配置:demo/config-{reconcile,merge,split,clean,summarize}.json
     并把「对账」写入了 config.json 作为默认任务。

  现在可以直接试(不用改任何东西):

     node butler.mjs                                    对账(默认)
     node butler.mjs --config=demo/config-merge.json     合并三个月度表
     node butler.mjs --config=demo/config-split.json     按状态拆分
     node butler.mjs --config=demo/config-clean.json     清洗脏数据
     node butler.mjs --config=demo/config-summarize.json 分组汇总

  结果会写到 out/ 目录。用 Excel 打开即可。
`);
