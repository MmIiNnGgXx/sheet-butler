// CSV 读写 —— 零依赖,处理中文场景的三个真实痛点:
//   ① Excel 导出的 CSV 在中文 Windows 上是 GBK 编码(直接按 UTF-8 读会乱码)
//   ② 带逗号 / 换行 / 引号的字段
//   ③ 分隔符可能是逗号、分号或制表符

/* ---------------- 编码 ---------------- */
/** 尝试把二进制解码成文本:优先 UTF-8,出现替换字符则回退 GBK(中文 Windows 常见) */
export function decodeText(buf) {
  if (buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) {
    return { text: buf.subarray(3).toString("utf8"), encoding: "utf-8-bom" };
  }
  const utf8 = buf.toString("utf8");
  if (!utf8.includes("\uFFFD")) return { text: utf8, encoding: "utf-8" };
  try {
    const gbk = new TextDecoder("gbk", { fatal: false }).decode(buf);
    if (!gbk.includes("\uFFFD")) return { text: gbk, encoding: "gbk" };
    return { text: gbk, encoding: "gbk(有少量无法解码字符)" };
  } catch {
    return { text: utf8, encoding: "utf-8(有乱码)" };
  }
}

/* ---------------- 分隔符探测 ---------------- */
export function detectDelimiter(text) {
  const sample = text.split(/\r?\n/).slice(0, 5).join("\n");
  const counts = [[",", 0], ["\t", 0], [";", 0], ["|", 0]];
  let inQuote = false;
  for (const ch of sample) {
    if (ch === '"') inQuote = !inQuote;
    else if (!inQuote) { const c = counts.find(([d]) => d === ch); if (c) c[1]++; }
  }
  counts.sort((a, b) => b[1] - a[1]);
  return counts[0][1] > 0 ? counts[0][0] : ",";
}

/* ---------------- 解析 ---------------- */
/**
 * 解析 CSV 文本为二维数组
 * @param {string} text
 * @param {{delimiter?:string}} opts
 */
export function parseCsv(text, opts = {}) {
  const delimiter = opts.delimiter ?? detectDelimiter(text);
  const rows = [];
  let row = [];
  let field = "";
  let inQuote = false;
  let i = 0;

  const pushField = () => { row.push(field); field = ""; };
  const pushRow = () => { rows.push(row); row = []; };

  while (i < text.length) {
    const ch = text[i];

    if (inQuote) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i += 2; continue; }  // 转义的双引号
        inQuote = false; i++; continue;
      }
      field += ch; i++; continue;
    }

    if (ch === '"' && field === "") { inQuote = true; i++; continue; }
    if (ch === delimiter) { pushField(); i++; continue; }
    if (ch === "\r") { if (text[i + 1] === "\n") i++; pushField(); pushRow(); i++; continue; }
    if (ch === "\n") { pushField(); pushRow(); i++; continue; }
    field += ch; i++;
  }
  if (field !== "" || row.length) { pushField(); pushRow(); }

  // 去掉完全空的行
  return rows.filter((r) => !(r.length === 1 && r[0] === ""));
}

/* ---------------- 序列化 ---------------- */
const needsQuote = (s, d) => s.includes(d) || s.includes('"') || s.includes("\n") || s.includes("\r") || /^\s|\s$/.test(s);
const cellToText = (v) => {
  if (v == null) return "";
  if (v instanceof Date) {
    const p = (x) => String(x).padStart(2, "0");
    return `${v.getFullYear()}-${p(v.getMonth() + 1)}-${p(v.getDate())}`;
  }
  return String(v);
};

/**
 * 二维数组 -> CSV 文本
 * @param {any[][]} rows
 * @param {{delimiter?:string, bom?:boolean}} opts
 */
export function toCsv(rows, opts = {}) {
  const d = opts.delimiter ?? ",";
  const body = rows
    .map((r) => (r ?? []).map((v) => {
      const s = cellToText(v);
      return needsQuote(s, d) ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(d))
    .join("\r\n");
  return (opts.bom === false ? "" : "\uFEFF") + body + "\r\n";
}

/* ---------------- 类型推断 ---------------- */
/** CSV 里一切都是字符串,按需把"看起来像数字/日期的"转回来 */
export function inferCell(s) {
  if (s == null || s === "") return null;
  const t = String(s).trim();
  if (t === "") return null;
  if (/^-?\d+$/.test(t)) { const n = Number(t); if (Number.isSafeInteger(n)) return n; }
  if (/^-?\d*\.\d+$/.test(t) || /^-?\d+(\.\d+)?[eE][-+]?\d+$/.test(t)) {
    const n = Number(t);
    if (!Number.isNaN(n)) return n;
  }
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2})?)?$/.test(t)) {
    const d = new Date(t.replace(/\//g, "-").replace(" ", "T"));
    if (!Number.isNaN(d.getTime())) return d;
  }
  if (t === "true" || t === "TRUE") return true;
  if (t === "false" || t === "FALSE") return false;
  return String(s);
}
