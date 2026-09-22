// 零依赖读写 .xlsx —— 只用 node:zlib + 自写 zip 层 + 正则解析 XML
//
// 支持:多工作表 / 共享字符串 / 内联字符串 / 数字 / 布尔 / 公式结果 / 日期(按样式识别)
// 写出:inline string(不依赖共享字符串表),自动为日期单元格套用日期格式
import { unzip, zip } from "./zip.mjs";

/* ============================ XML 小工具 ============================ */
const unescapeXml = (s) =>
  String(s)
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, "&");

const escapeXml = (s) =>
  String(s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    // 去掉 XML 1.0 不允许的控制字符(否则 Excel 会报文件损坏)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");

/** 列号:0 -> A, 25 -> Z, 26 -> AA */
export function colName(i) {
  let s = "";
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
/** "AB12" -> { col: 27, row: 11 }(行列都从 0 开始) */
export function parseRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col: col - 1, row: Number(m[2]) - 1 };
}

/* ============================ 日期识别 ============================ */
const BUILTIN_DATE_FMT = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57]);
function isDateFormatCode(code) {
  if (!code) return false;
  const c = code.replace(/\[[^\]]*\]/g, "").replace(/"[^"]*"/g, "");
  return /[ymdhs]/i.test(c) && !/^[^yMdhs]*(0+\.?0*|#+\.?#*)$/.test(c);
}

/** Excel 序列号 -> JS Date(1900 系统,含著名的 1900 闰年 bug) */
export function serialToDate(n) {
  const ms = Math.round((n - 25569) * 86400 * 1000);
  return new Date(ms);
}
export function dateToSerial(d) {
  return (d.getTime() / 86400000) + 25569;
}
export function fmtDate(d, withTime = false) {
  const p = (x) => String(x).padStart(2, "0");
  const s = `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
  return withTime ? `${s} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}` : s;
}

/* ============================ 读 ============================ */
function readStyles(files) {
  const xml = files.get("xl/styles.xml")?.toString("utf8");
  const dateStyle = new Set();
  if (!xml) return dateStyle;
  // 自定义 numFmt
  const custom = new Map();
  const nfRe = /<numFmt[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g;
  let m;
  while ((m = nfRe.exec(xml))) custom.set(Number(m[1]), m[2]);
  // cellXfs:每个索引对应的 numFmtId
  const cellXfs = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
  if (cellXfs) {
    const xfRe = /<xf\b[^>]*numFmtId="(\d+)"/g;
    let i = 0;
    while ((m = xfRe.exec(cellXfs[1]))) {
      const id = Number(m[1]);
      const code = custom.get(id);
      if (BUILTIN_DATE_FMT.has(id) || isDateFormatCode(code)) dateStyle.add(i);
      i++;
    }
  }
  return dateStyle;
}

function readSharedStrings(files) {
  const xml = files.get("xl/sharedStrings.xml")?.toString("utf8");
  if (!xml) return [];
  const out = [];
  const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = siRe.exec(xml))) {
    // 一个 si 里可能有多段 <t>(富文本)
    const parts = [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((x) => unescapeXml(x[1]));
    out.push(parts.join(""));
  }
  return out;
}

/** 把 workbook.xml + rels 解析成 [{name, path}] */
function readSheetList(files) {
  const wb = files.get("xl/workbook.xml")?.toString("utf8") ?? "";
  const rels = files.get("xl/_rels/workbook.xml.rels")?.toString("utf8") ?? "";
  const relMap = new Map();
  for (const m of rels.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    relMap.set(m[1], m[2].replace(/^\/?xl\//, "").replace(/^\//, ""));
  }
  const sheets = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*\/?>/g)) {
    const tag = m[0];
    const name = /name="([^"]*)"/.exec(tag)?.[1] ?? `Sheet${sheets.length + 1}`;
    const rid = /r:id="([^"]+)"/.exec(tag)?.[1];
    let target = rid ? relMap.get(rid) : null;
    if (!target) target = `worksheets/sheet${sheets.length + 1}.xml`;
    sheets.push({ name: unescapeXml(name), path: `xl/${target}` });
  }
  if (!sheets.length) {
    // 没有 workbook.xml 的极端情况:直接找 worksheets
    for (const k of files.keys()) if (/^xl\/worksheets\/sheet\d+\.xml$/.test(k)) sheets.push({ name: k, path: k });
  }
  return sheets;
}

function readSheet(xml, shared, dateStyle) {
  const rows = [];
  const rowRe = /<row\b([^>]*)>([\s\S]*?)<\/row>|<row\b([^>]*)\/>/g;
  let rm;
  while ((rm = rowRe.exec(xml))) {
    const attrs = rm[1] ?? rm[3] ?? "";
    const body = rm[2] ?? "";
    const rIdx = Number(/r="(\d+)"/.exec(attrs)?.[1] ?? rows.length + 1) - 1;
    const row = rows[rIdx] ?? (rows[rIdx] = []);

    const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let cm;
    while ((cm = cellRe.exec(body))) {
      const cAttrs = cm[1] ?? "";
      const cBody = cm[2] ?? "";
      const ref = /r="([A-Z]+\d+)"/.exec(cAttrs)?.[1];
      const t = /t="([^"]+)"/.exec(cAttrs)?.[1] ?? "n";
      const sIdx = /s="(\d+)"/.exec(cAttrs)?.[1];
      const col = ref ? parseRef(ref).col : row.length;

      let v = null;
      if (t === "inlineStr") {
        v = [...cBody.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((x) => unescapeXml(x[1])).join("");
      } else if (t === "s") {
        const i = Number(/<v>([\s\S]*?)<\/v>/.exec(cBody)?.[1] ?? -1);
        v = shared[i] ?? "";
      } else if (t === "str") {
        v = unescapeXml(/<v>([\s\S]*?)<\/v>/.exec(cBody)?.[1] ?? "");
      } else if (t === "b") {
        v = /<v>1<\/v>/.test(cBody);
      } else if (t === "e") {
        v = unescapeXml(/<v>([\s\S]*?)<\/v>/.exec(cBody)?.[1] ?? "");
      } else {
        const raw = /<v>([\s\S]*?)<\/v>/.exec(cBody)?.[1];
        if (raw != null && raw !== "") {
          const num = Number(raw);
          if (!Number.isNaN(num)) {
            // 数字:若样式是日期格式,转成 Date
            v = sIdx != null && dateStyle.has(Number(sIdx)) ? serialToDate(num) : num;
          } else v = unescapeXml(raw);
        } else {
          // 只有公式没有缓存值
          const f = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(cBody)?.[1];
          v = f ? `=${unescapeXml(f)}` : null;
        }
      }
      row[col] = v;
    }
  }
  // 补齐空洞 + 去掉尾部空行
  const out = [];
  for (const row of rows) {
    if (!row) { out.push([]); continue; }
    const arr = [];
    for (let i = 0; i < row.length; i++) arr.push(row[i] === undefined ? null : row[i]);
    out.push(arr);
  }
  while (out.length && out[out.length - 1].every((c) => c == null || c === "")) out.pop();
  return out;
}

/**
 * 读取 xlsx
 * @param {Buffer} buf
 * @returns {{sheets: Array<{name:string, rows:any[][]}>}}
 */
export function readXlsx(buf) {
  const files = unzip(buf);
  const shared = readSharedStrings(files);
  const dateStyle = readStyles(files);
  const list = readSheetList(files);
  const sheets = [];
  for (const s of list) {
    const xml = files.get(s.path)?.toString("utf8");
    if (!xml) continue;
    sheets.push({ name: s.name, rows: readSheet(xml, shared, dateStyle) });
  }
  return { sheets };
}

/* ============================ 写 ============================ */
const cellRef = (r, c) => `${colName(c)}${r + 1}`;

function sheetXml(rows) {
  const parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    "<sheetData>"];
  rows.forEach((row, r) => {
    parts.push(`<row r="${r + 1}">`);
    (row ?? []).forEach((v, c) => {
      if (v == null || v === "") return;
      const ref = cellRef(r, c);
      if (v instanceof Date) {
        parts.push(`<c r="${ref}" s="1"><v>${dateToSerial(v)}</v></c>`);
      } else if (typeof v === "number" && Number.isFinite(v)) {
        parts.push(`<c r="${ref}"><v>${v}</v></c>`);
      } else if (typeof v === "boolean") {
        parts.push(`<c r="${ref}" t="b"><v>${v ? 1 : 0}</v></c>`);
      } else {
        parts.push(`<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(v)}</t></is></c>`);
      }
    });
    parts.push("</row>");
  });
  parts.push("</sheetData></worksheet>");
  return parts.join("");
}

/**
 * 写出 xlsx
 * @param {{sheets: Array<{name?:string, rows:any[][]}>}} wb
 * @returns {Buffer}
 */
export function writeXlsx(wb) {
  const sheets = wb.sheets.map((s, i) => ({ name: s.name || `Sheet${i + 1}`, rows: s.rows ?? [] }));
  const files = [];

  files.push({
    name: "[Content_Types].xml",
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("\n")}
</Types>`,
  });

  files.push({
    name: "_rels/.rels",
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  });

  files.push({
    name: "xl/workbook.xml",
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${sheets.map((s, i) => `<sheet name="${escapeXml(s.name.slice(0, 31))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets>
</workbook>`,
  });

  files.push({
    name: "xl/_rels/workbook.xml.rels",
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("\n")}
<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
  });

  // 样式:0=常规,1=日期(yyyy-mm-dd)
  files.push({
    name: "xl/styles.xml",
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/></numFmts>
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
</styleSheet>`,
  });

  sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s.rows) }));

  return zip(files);
}
