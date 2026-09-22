#!/usr/bin/env node
// 真实案例:月底对账
//   node 案例-月底对账.mjs
//
// 场景:一家做小生意的公司,订单系统能导出 Excel,财务用 Excel 手工记账。
//       月底要把两边对上,找出:漏记的、多记的、金额对不上的。
//       以前是财务一行行肉眼比,现在一条命令搞定。
//
// 会生成一个月的量级(约 500 单),并故意埋入几类典型差异。
import { mkdirSync, writeFileSync, existsSync, rmSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { writeXlsx } from "./lib/xlsx.mjs";
import { toCsv } from "./lib/csv.mjs";

const ROOT = dirname(fileURLToPath(import.meta.url));
const D = join(ROOT, "案例-月底对账");
if (existsSync(D)) rmSync(D, { recursive: true, force: true });
mkdirSync(D, { recursive: true });

/* ---------------- 造一个月的订单 ---------------- */
const 客户 = ["张三", "李四", "王五", "赵六", "钱七", "孙八", "周九", "吴十", "郑一", "王二", "陈三", "褚四", "卫五", "蒋六", "沈七", "韩八"];
const 地区 = ["华东", "华南", "华北", "西南", "华中"];
const 商品 = ["无线耳机", "蓝牙音箱", "充电宝", "数据线", "手机壳", "键盘", "鼠标", "支架"];
const rnd = (seed) => { let s = seed; return () => (s = (s * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff; };
const R = rnd(20260922);
const pick = (a) => a[Math.floor(R() * a.length)];

const N = 500;
const 订单 = [["订单号", "客户", "地区", "商品", "金额", "日期", "状态"]];
const 台账 = [["订单号", "金额", "状态", "记账日期"]];
const 埋的差异 = { 漏记: [], 多记: [], 金额不符: [], 状态不符: [] };

for (let i = 1; i <= N; i++) {
  const id = "SO202609" + String(i).padStart(4, "0");
  const 金额 = Number((R() * 4800 + 50).toFixed(2));
  const day = 1 + Math.floor(R() * 30);
  const 状态 = R() < 0.82 ? "已付款" : "待付款";
  const 行 = [id, pick(客户), pick(地区), pick(商品), 金额, new Date(2026, 8, day), 状态];
  订单.push(行);

  // 台账:大部分一致,埋 4 类差异
  const r = R();
  if (r < 0.02) { 埋的差异.漏记.push(id); continue; }                     // 2% 财务漏记
  let 记金额 = 金额, 记状态 = 状态;
  if (r < 0.035) { 记金额 = Number((金额 + (R() < 0.5 ? -1 : 1) * (R() * 60 + 5)).toFixed(2)); 埋的差异.金额不符.push(id); }
  else if (r < 0.045) { 记状态 = 状态 === "已付款" ? "待付款" : "已付款"; 埋的差异.状态不符.push(id); }
  台账.push([id, 记金额, 记状态, new Date(2026, 8, Math.min(30, day + 1))]);
}
// 财务多记 6 笔(可能是重复记账或串号)
for (let i = 1; i <= 6; i++) {
  const id = "SO2026099" + String(i).padStart(3, "0");
  台账.push([id, Number((R() * 3000 + 100).toFixed(2)), "已付款", new Date(2026, 8, 30)]);
  埋的差异.多记.push(id);
}

writeFileSync(join(D, "订单系统导出.xlsx"), writeXlsx({ sheets: [{ name: "订单明细", rows: 订单 }] }));
writeFileSync(join(D, "财务台账.csv"), toCsv(台账), "utf8");

/* ---------------- 配置 ---------------- */
const cfg = {
  task: "reconcile",
  inputs: { a: "案例-月底对账/订单系统导出.xlsx", b: "案例-月底对账/财务台账.csv" },
  output: "out/月底对账结果.xlsx",
  options: { keyColumns: ["订单号"], compareColumns: ["金额", "状态"] },
};
writeFileSync(join(D, "config-对账.json"), JSON.stringify(cfg, null, 2) + "\n", "utf8");
writeFileSync(join(ROOT, "config.json"), JSON.stringify(cfg, null, 2) + "\n", "utf8");

/* ---------------- 把"正确答案"单独存一份,方便核对工具有没有算错 ---------------- */
writeFileSync(join(D, "标准答案.json"), JSON.stringify({
  说明: "这份是生成时埋进去的差异清单,用来核对对账工具算得对不对",
  订单总数: N,
  台账行数: 台账.length - 1,
  漏记_订单有台账没有: 埋的差异.漏记,
  多记_台账有订单没有: 埋的差异.多记,
  金额不一致: 埋的差异.金额不符,
  状态不一致: 埋的差异.状态不符,
  预期结论: `仅左表有 ${埋的差异.漏记.length} + 仅右表有 ${埋的差异.多记.length} + 值不一致 ${埋的差异.金额不符.length + 埋的差异.状态不符.length} = ${埋的差异.漏记.length + 埋的差异.多记.length + 埋的差异.金额不符.length + 埋的差异.状态不符.length} 处差异`,
}, null, 2), "utf8");

console.log(`
  ✅ 真实案例已生成:案例-月底对账/

     订单系统导出.xlsx   ${N} 条订单(Excel,含客户/地区/商品/金额/日期/状态)
     财务台账.csv        ${台账.length - 1} 条(财务手工记的,故意有几处对不上)
     标准答案.json       埋进去的差异清单 —— 用来核对工具算得对不对

     埋入的差异:
       · 财务漏记 ${埋的差异.漏记.length} 笔
       · 财务多记 ${埋的差异.多记.length} 笔
       · 金额对不上 ${埋的差异.金额不符.length} 笔
       · 状态对不上 ${埋的差异.状态不符.length} 笔
       合计应有 ${埋的差异.漏记.length + 埋的差异.多记.length + 埋的差异.金额不符.length + 埋的差异.状态不符.length} 处差异

  下一步(两条命令):

     node butler.mjs                          命令行对账
     node webui.mjs                           网页版对账(推荐)

  对完打开 案例-月底对账/标准答案.json 核对一下,看工具算得对不对。
`);
