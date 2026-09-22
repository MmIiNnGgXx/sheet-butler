// 网页版自检:自己起一个服务(独立端口)跑完整流程,跑完自动关闭
//   node selftest-webui.mjs
import { readFileSync, writeFileSync, rmSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";
import { readXlsx } from "./lib/xlsx.mjs";

const ROOT = dirname(fileURLToPath(import.meta.url));
const PORT = 4299;                       // 独立端口,不打扰正在使用的 4220
const B = `http://127.0.0.1:${PORT}`;

rmSync(join(ROOT, "webdata"), { recursive: true, force: true });   // 干净起步
const srv = spawn(process.execPath, [join(ROOT, "webui.mjs"), `--port=${PORT}`], { cwd: ROOT, stdio: "ignore" });
const cleanup = () => { try { srv.kill(); } catch {} };
process.on("exit", cleanup);
process.on("SIGINT", () => { cleanup(); process.exit(1); });

let serverReady = false;
for (let i = 0; i < 40 && !serverReady; i++) {
  try { await fetch(`${B}/api/files`); serverReady = true; }
  catch { await new Promise((r) => setTimeout(r, 250)); }
}
if (!serverReady) { console.error("  ❌ 测试服务启动失败"); cleanup(); process.exit(1); }
console.log(`\n  表格管家 · 网页版自检(临时服务 127.0.0.1:${PORT})\n`);

let pass = 0, fail = 0;
const ok = (c, n, extra = "") => { if (c) { pass++; console.log(`  ✅ ${n}`); } else { fail++; console.log(`  ❌ ${n}${extra ? "  → " + extra : ""}`); } };

/* ---------- ① 上传(multipart)---------- */
console.log("① 上传文件");
const up = async (paths) => {
  const fd = new FormData();
  for (const p of paths) fd.append("files", new Blob([readFileSync(p)]), p.split(/[\\/]/).pop());
  const r = await fetch(`${B}/api/upload`, { method: "POST", body: fd });
  return r.json();
};
const r1 = await up(["demo/订单系统导出.xlsx", "demo/财务台账.csv"]);
ok(r1.ok, "上传成功", r1.error);
ok(r1.saved.length === 2, "保存 2 个文件", JSON.stringify(r1.saved?.map((s) => s.name)));
ok(r1.saved.some((s) => s.name === "订单系统导出.xlsx"), "xlsx 文件名保留中文");
ok(r1.saved.some((s) => s.name === "财务台账.csv"), "csv 文件名保留中文");

/* ---------- ② 拒绝不支持的类型 ---------- */
console.log("\n② 类型限制");
const fd2 = new FormData();
fd2.append("files", new Blob(["bad"]), "恶意.exe");
const r2 = await (await fetch(`${B}/api/upload`, { method: "POST", body: fd2 })).json();
ok(r2.ok && r2.saved.length === 0 && r2.rejected.length === 1, "非法扩展名被拒绝", JSON.stringify(r2.rejected));
ok(/不支持的格式/.test(r2.rejected[0].reason), "给出拒绝原因");

/* ---------- ③ 文件名安全化 ---------- */
console.log("\n③ 文件名安全化");
const fd3 = new FormData();
fd3.append("files", new Blob(["a,b\n1,2\n"]), "../../偷偷逃出去.csv");
const r3 = await (await fetch(`${B}/api/upload`, { method: "POST", body: fd3 })).json();
ok(r3.ok, "上传成功");
ok(!r3.saved[0].name.includes(".."), "路径穿越的文件名被清理", r3.saved[0].name);
ok(!r3.saved[0].name.includes("/") && !r3.saved[0].name.includes("\\"), "不含路径分隔符", r3.saved[0].name);

/* ---------- ④ 预览表头 ---------- */
console.log("\n④ 读取表头(供前端下拉)");
const pv = await (await fetch(`${B}/api/preview`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ name: "订单系统导出.xlsx" }) })).json();
ok(pv.ok, "预览成功");
ok(JSON.stringify(pv.headers) === JSON.stringify(["订单号", "客户", "金额", "日期", "状态"]), "表头正确", JSON.stringify(pv.headers));
ok(pv.total === 12, "行数正确", String(pv.total));
ok(Array.isArray(pv.rows) && pv.rows.length === 5, "返回前 5 行样例");

/* ---------- ⑤ 执行对账 ---------- */
console.log("\n⑤ 执行对账");
const run = await (await fetch(`${B}/api/run`, {
  method: "POST", headers: { "content-type": "application/json" },
  body: JSON.stringify({ task: "reconcile", files: ["订单系统导出.xlsx", "财务台账.csv"], options: { keyColumns: ["订单号"], compareColumns: ["金额", "状态"] } }),
})).json();
ok(run.ok, "执行成功", run.error);
ok(run.summary["仅左表有"] === 2 && run.summary["仅右表有"] === 1 && run.summary["值不一致"] === 1, "对账计数正确", JSON.stringify(run.summary));
ok(run.outputs?.length === 1, "产出 1 个文件", JSON.stringify(run.outputs?.map((o) => o.name)));
ok(run.outputs[0].name === "对账结果.xlsx", "文件名正确");
ok(run.ms > 0, `耗时 ${run.ms}ms`);

/* ---------- ⑥ 下载结果 ---------- */
console.log("\n⑥ 下载结果");
const dl = await fetch(`${B}/api/download?from=out&name=${encodeURIComponent("对账结果.xlsx")}`);
ok(dl.ok, "下载成功");
const buf = Buffer.from(await dl.arrayBuffer());
ok(buf.length > 1000, `文件大小 ${buf.length} 字节`);
const wb = readXlsx(buf);
ok(wb.sheets.length === 4, "下载到的确实是合法的 4 工作表 Excel", wb.sheets.map((s) => s.name).join(","));
writeFileSync("_webui-downloaded.xlsx", buf);

/* ---------- ⑦ 参数校验 ---------- */
console.log("\n⑦ 错误处理");
const bad1 = await (await fetch(`${B}/api/run`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ task: "reconcile", files: ["订单系统导出.xlsx"], options: {} }) })).json();
ok(!bad1.ok && /两个文件/.test(bad1.error), "文件不够时给出可读错误", bad1.error);
const bad2 = await (await fetch(`${B}/api/run`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ task: "nope", files: [], options: {} }) })).json();
ok(!bad2.ok && /不支持的任务/.test(bad2.error), "未知任务被拒绝", bad2.error);
const bad3 = await (await fetch(`${B}/api/run`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ task: "reconcile", files: ["订单系统导出.xlsx", "财务台账.csv"], options: { keyColumns: ["不存在的列"] } }) })).json();
ok(!bad3.ok && /找不到列/.test(bad3.error), "列名不存在时给出可用列表", bad3.error);

/* ---------- ⑧ 合并 + 拆分 + 汇总 ---------- */
console.log("\n⑧ 其他任务");
await up(["demo/月度报表-01月.csv", "demo/月度报表-02月.csv", "demo/月度报表-03月.csv"]);
const mg = await (await fetch(`${B}/api/run`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ task: "merge", files: ["月度报表-01月.csv", "月度报表-02月.csv", "月度报表-03月.csv"], options: { sourceColumn: "来源文件" } }) })).json();
ok(mg.ok && mg.summary["合并后行数"] === 12, "合并 12 行", JSON.stringify(mg.summary));

const sp = await (await fetch(`${B}/api/run`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ task: "split", files: ["订单系统导出.xlsx"], options: { by: ["状态"] } }) })).json();
ok(sp.ok && sp.outputs.length === 2, "拆分出 2 个文件", JSON.stringify(sp.outputs?.map((o) => o.name)));

const sm = await (await fetch(`${B}/api/run`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ task: "summarize", files: ["订单系统导出.xlsx"], options: { groupBy: ["状态"], aggregations: [{ column: "金额", op: "sum", as: "金额合计" }] } }) })).json();
ok(sm.ok, "汇总成功", sm.error);

/* ---------- ⑨ 输出列表 ---------- */
console.log("\n⑨ 输出文件列表");
const fl = await (await fetch(`${B}/api/files`)).json();
ok(fl.outputs.length >= 4, `输出目录有 ${fl.outputs.length} 个文件`, JSON.stringify(fl.outputs.map((o) => o.name)));

console.log(`\n${"─".repeat(46)}`);
console.log(`  网页版验证:${pass} 通过${fail ? `,${fail} 失败` : ""}`);
console.log(`${"─".repeat(46)}\n`);
process.exit(fail ? 1 : 0);
