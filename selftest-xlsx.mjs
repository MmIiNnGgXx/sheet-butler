// 临时验证:zip + xlsx 读写
import { readFileSync, existsSync } from "node:fs";
import { zip, unzip, crc32 } from "./lib/zip.mjs";
import { writeXlsx, readXlsx, colName, parseRef, serialToDate } from "./lib/xlsx.mjs";

let pass = 0, fail = 0;
const ok = (c, n, extra = "") => { if (c) { pass++; console.log(`  ✅ ${n}`); } else { fail++; console.log(`  ❌ ${n}${extra ? "  → " + extra : ""}`); } };

/* ---------- ① zip 往返 ---------- */
console.log("① ZIP 读写");
const files = [
  { name: "hello.txt", data: "你好,世界\n第二行" },
  { name: "dir/data.json", data: JSON.stringify({ a: 1, b: [1, 2, 3] }) },
  { name: "bin.dat", data: Buffer.from([0, 1, 2, 255, 254, 253]) },
  { name: "big.txt", data: "重复内容".repeat(5000) },   // 触发 deflate
];
const zipped = zip(files);
ok(zipped.length > 0 && zipped.readUInt32LE(0) === 0x04034b50, "生成的 zip 有正确的本地头签名");
const un = unzip(zipped);
ok(un.size === 4, "解出 4 个条目", `实际 ${un.size}`);
ok(un.get("hello.txt").toString("utf8") === "你好,世界\n第二行", "文本内容一致");
ok(un.get("dir/data.json").toString("utf8") === JSON.stringify({ a: 1, b: [1, 2, 3] }), "子目录条目一致");
ok(Buffer.compare(un.get("bin.dat"), Buffer.from([0, 1, 2, 255, 254, 253])) === 0, "二进制内容一致");
ok(un.get("big.txt").toString("utf8") === "重复内容".repeat(5000), "压缩后的大文件还原一致");
ok(crc32(Buffer.from("123456789")) === 0xcbf43926, "CRC32 标准测试向量正确", crc32(Buffer.from("123456789")).toString(16));

/* ---------- ② 用真实 Office 文件验证解压器 ---------- */
console.log("\n② 解压真实 Office 文件(docx 也是 zip)");
const docxPath = "D:\\APP瞎忙活\\毕业论文\\(v17)BDT22031+徐铭杰+基于Hive的高考录取分数线数据分析与预测平台.docx";
if (existsSync(docxPath)) {
  try {
    const real = unzip(readFileSync(docxPath));
    const names = [...real.keys()];
    ok(names.length > 5, `解出 ${names.length} 个内部文件`);
    ok(names.includes("[Content_Types].xml"), "包含 [Content_Types].xml");
    ok(names.some((n) => n === "word/document.xml"), "包含 word/document.xml");
    const docXml = real.get("word/document.xml").toString("utf8");
    ok(docXml.length > 1000, `document.xml 有 ${Math.round(docXml.length / 1024)}KB 内容`);
    ok(/<w:document/.test(docXml), "内容是合法的 WordprocessingML");
  } catch (e) { ok(false, "解压真实 docx", e.message); }
} else {
  console.log("  (跳过:找不到测试用 docx)");
}

/* ---------- ③ xlsx 往返 ---------- */
console.log("\n③ XLSX 往返");
const now = new Date(2026, 8, 22, 14, 30, 0);
const rows1 = [
  ["姓名", "金额", "日期", "是否付款", "备注"],
  ["张三", 1234.56, now, true, "含 & < > \" ' 特殊字符"],
  ["李四", -99, new Date(2026, 0, 1), false, "中文😀emoji"],
  [null, 0, null, null, ""],
  ["王五", 1e6, new Date(2020, 11, 31), true, null],
];
const rows2 = [["汇总"], ["总计", 1999999.56]];
const xbuf = writeXlsx({ sheets: [{ name: "明细", rows: rows1 }, { name: "汇总表", rows: rows2 }] });
ok(xbuf.length > 500, `生成 xlsx ${Math.round(xbuf.length / 1024)}KB`);

const back = readXlsx(xbuf);
ok(back.sheets.length === 2, "读回 2 个工作表", `实际 ${back.sheets.length}`);
ok(back.sheets[0].name === "明细" && back.sheets[1].name === "汇总表", "工作表名一致");

const r = back.sheets[0].rows;
ok(r[0][0] === "姓名" && r[0][1] === "金额", "表头一致");
ok(r[1][0] === "张三", "中文字符串一致");
ok(r[1][1] === 1234.56, "小数一致", String(r[1][1]));
ok(r[2][1] === -99, "负数一致");
ok(r[3][1] === 0, "零值保留(不被当成空)", String(r[3][1]));
ok(r[1][3] === true && r[2][3] === false, "布尔值一致");
ok(r[1][4] === "含 & < > \" ' 特殊字符", "XML 特殊字符转义往返正确", r[1][4]);
ok(r[2][4] === "中文😀emoji", "emoji 往返正确", r[2][4]);
ok(r[4][1] === 1e6, "大数一致");

const d1 = r[1][2];
ok(d1 instanceof Date, "日期被识别为 Date", typeof d1);
ok(d1 && d1.getFullYear() === 2026 && d1.getMonth() === 8 && d1.getDate() === 22, "日期值正确", d1 ? d1.toISOString() : "");
const d2 = r[2][2];
ok(d2 instanceof Date && d2.getFullYear() === 2026 && d2.getMonth() === 0 && d2.getDate() === 1, "跨年日期正确", d2 ? d2.toISOString() : "");
const d3 = r[4][2];
ok(d3 instanceof Date && d3.getFullYear() === 2020 && d3.getMonth() === 11 && d3.getDate() === 31, "月末日期正确", d3 ? d3.toISOString() : "");

ok(r[3][0] === null || r[3][0] === undefined, "空单元格为空");
ok(back.sheets[1].rows[1][1] === 1999999.56, "第二个表的数据正确");

/* ---------- ④ 辅助函数 ---------- */
console.log("\n④ 列号与引用");
ok(colName(0) === "A" && colName(25) === "Z" && colName(26) === "AA" && colName(701) === "ZZ" && colName(702) === "AAA", "列号换算正确", `${colName(0)},${colName(25)},${colName(26)},${colName(701)},${colName(702)}`);
ok(parseRef("A1").col === 0 && parseRef("A1").row === 0, "A1 解析正确");
ok(parseRef("AB13").col === 27 && parseRef("AB13").row === 12, "AB13 解析正确");
const sd = serialToDate(45000);
ok(sd.getFullYear() === 2023, "Excel 序列号 -> 日期", sd.toISOString().slice(0, 10));

/* ---------- ⑤ 边界情况 ---------- */
console.log("\n⑤ 边界情况");
const edge = writeXlsx({ sheets: [{ name: "e", rows: [["a"], [], ["c"], [null, null, "x"]] }] });
const eb = readXlsx(edge);
ok(eb.sheets[0].rows[2][0] === "c", "空行被保留(不串行)", JSON.stringify(eb.sheets[0].rows));
ok(eb.sheets[0].rows[3][2] === "x", "稀疏单元格按列号落位正确");
const longStr = writeXlsx({ sheets: [{ name: "L", rows: [[("很长的一行" + "内容").repeat(200)]] }] });
ok(readXlsx(longStr).sheets[0].rows[0][0].length === ("很长的一行" + "内容").repeat(200).length, "超长字符串往返正确");
const badChar = writeXlsx({ sheets: [{ name: "B", rows: [["a\u0000b\u0007c"]] }] });
ok(readXlsx(badChar).sheets[0].rows[0][0] === "abc", "非法控制字符被清除(避免 Excel 报损坏)");

console.log(`\n${"─".repeat(44)}`);
console.log(`  xlsx/zip 验证:${pass} 通过${fail ? `,${fail} 失败` : ""}`);
console.log(`${"─".repeat(44)}\n`);
process.exit(fail ? 1 : 0);
