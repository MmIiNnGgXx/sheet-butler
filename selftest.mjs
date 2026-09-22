#!/usr/bin/env node
// 表格管家 · 全量自检 —— 不联网、不碰你的数据,全部用临时目录
//   node selftest.mjs
// 退出码:0 全通过 / 1 有失败
import { spawnSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = dirname(fileURLToPath(import.meta.url));
const suites = [
  ["底层:xlsx / zip 读写", "selftest-xlsx.mjs"],
  ["任务:合并 / 拆分 / 对账 / 清洗 / 汇总", "selftest-tasks.mjs"],
  ["网页版:上传 / 执行 / 下载 / 安全防护", "selftest-webui.mjs"],
];

console.log("\n  表格管家 · 自检");
console.log("  ══════════════════════════════════════════");

let failed = 0;
const results = [];
for (const [name, file] of suites) {
  console.log(`\n  ▶ ${name}`);
  const r = spawnSync(process.execPath, [join(ROOT, file)], { stdio: "inherit", cwd: ROOT });
  const okRun = r.status === 0;
  if (!okRun) failed++;
  results.push({ name, ok: okRun });
}

console.log("\n  ══════════════════════════════════════════");
for (const r of results) console.log(`  ${r.ok ? "✅" : "❌"} ${r.name}`);
console.log(`  ══════════════════════════════════════════`);
console.log(`  ${failed ? `❌ ${failed} 个套件失败` : "✅ 全部通过"}\n`);
process.exit(failed ? 1 : 0);
