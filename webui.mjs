#!/usr/bin/env node
// 表格管家 · 网页版
//
//   node webui.mjs              启动 → 浏览器打开 http://127.0.0.1:4220/
//
// 安全设计(重要):
//   ① 只绑定 127.0.0.1 —— 局域网/其他设备访问不到
//   ② 只服务 public/(前端) 与 webdata/(上传与输出) 两个目录,其他一律 403
//   ③ 下载接口做路径穿越防护(禁止 .. 与绝对路径)
//   ④ 上传有体积与类型限制,文件名做安全化处理
//
// 零依赖:连 multipart/form-data 都是自己解析的,不装任何库。

import http from "node:http";
import { readFileSync, writeFileSync, existsSync, mkdirSync, readdirSync, statSync, rmSync } from "node:fs";
import { join, dirname, extname, basename, resolve, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import { TASKS } from "./lib/tasks.mjs";

const ROOT = dirname(fileURLToPath(import.meta.url));
const PUBLIC_DIR = join(ROOT, "public");
const DATA_DIR = join(ROOT, "webdata");
const UP_DIR = join(DATA_DIR, "uploads");
const OUT_DIR = join(DATA_DIR, "out");
for (const d of [PUBLIC_DIR, DATA_DIR, UP_DIR, OUT_DIR]) mkdirSync(d, { recursive: true });

const argOf = (k) => process.argv.find((a) => a.startsWith(`--${k}=`))?.slice(k.length + 3);
const PORT = Number(argOf("port") ?? process.env.PORT ?? 4220);
const MAX_UPLOAD = 50 * 1024 * 1024;              // 单次上传上限 50MB
const ALLOWED_EXT = new Set([".xlsx", ".xlsm", ".csv", ".tsv", ".txt"]);

/* ============================ multipart 解析(零依赖) ============================ */
/**
 * 解析 multipart/form-data
 * @returns {Array<{name:string, filename:string|null, contentType:string, data:Buffer}>}
 */
function parseMultipart(buf, boundary) {
  const parts = [];
  const delim = Buffer.from(`--${boundary}`);
  let pos = buf.indexOf(delim);
  if (pos < 0) return parts;
  pos += delim.length;

  while (pos < buf.length) {
    if (buf[pos] === 0x2d && buf[pos + 1] === 0x2d) break;          // 结束标记 "--"
    if (buf[pos] === 0x0d && buf[pos + 1] === 0x0a) pos += 2;       // 跳过 CRLF

    const next = buf.indexOf(delim, pos);
    if (next < 0) break;
    let end = next;
    if (end >= 2 && buf[end - 2] === 0x0d && buf[end - 1] === 0x0a) end -= 2;

    const chunk = buf.subarray(pos, end);
    const sep = chunk.indexOf("\r\n\r\n");
    if (sep >= 0) {
      const headerText = chunk.subarray(0, sep).toString("utf8");
      const data = chunk.subarray(sep + 4);
      const headers = {};
      for (const line of headerText.split("\r\n")) {
        const i = line.indexOf(":");
        if (i > 0) headers[line.slice(0, i).trim().toLowerCase()] = line.slice(i + 1).trim();
      }
      const cd = headers["content-disposition"] ?? "";
      const name = /name="([^"]*)"/.exec(cd)?.[1] ?? "";
      let filename = /filename\*=(?:UTF-8'')?([^;]*)/i.exec(cd)?.[1] ?? /filename="([^"]*)"/i.exec(cd)?.[1] ?? null;
      if (filename) { try { filename = decodeURIComponent(filename); } catch {} }
      parts.push({ name, filename, contentType: headers["content-type"] ?? "", data });
    }
    pos = next + delim.length;
  }
  return parts;
}

const readBody = (req, limit) => new Promise((resolve, reject) => {
  const chunks = [];
  let size = 0;
  req.on("data", (c) => {
    size += c.length;
    if (size > limit) { req.destroy(); reject(new Error(`上传超过上限(${Math.round(limit / 1024 / 1024)}MB)`)); return; }
    chunks.push(c);
  });
  req.on("end", () => resolve(Buffer.concat(chunks)));
  req.on("error", reject);
});

/* ============================ 工具 ============================ */
const safeFileName = (s) => String(s || "file")
  .replace(/[\\/:*?"<>|\r\n\t]/g, "_")     // 路径分隔符与非法字符
  .replace(/\.{2,}/g, "_")                  // 连续点(路径穿越的素材)
  .replace(/^[.\s]+/, "")                   // 开头的点与空白
  .slice(0, 120) || "file";

const uniqueName = (dir, name) => {
  if (!existsSync(join(dir, name))) return name;
  const e = extname(name), b = basename(name, e);
  for (let i = 1; i < 9999; i++) if (!existsSync(join(dir, `${b}(${i})${e}`))) return `${b}(${i})${e}`;
  return `${Date.now()}-${name}`;
};

const MIME = {
  ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8", ".json": "application/json; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
  ".csv": "text/csv; charset=utf-8", ".tsv": "text/tab-separated-values; charset=utf-8",
};

const send = (res, code, body, headers = {}) => {
  res.writeHead(code, { "cache-control": "no-store", ...headers });
  res.end(body);
};
const sendJson = (res, obj, code = 200) => send(res, code, JSON.stringify(obj), { "content-type": "application/json; charset=utf-8" });

/** 只允许访问指定根目录下的文件(防路径穿越)。传入的 relPath 应当已解码。 */
function safeJoin(rootDir, relPath) {
  const root = normalize(rootDir);
  const p = normalize(join(root, String(relPath ?? "")));
  return p === root || p.startsWith(root + (process.platform === "win32" ? "\\" : "/")) ? p : null;
}

/* ============================ 任务执行 ============================ */
/** 列出文件;递归子目录(拆分任务会把结果放到子目录里),返回相对路径 */
function listFiles(dir, base = dir) {
  const out = [];
  if (!existsSync(dir)) return out;
  for (const f of readdirSync(dir)) {
    const p = join(dir, f);
    let st;
    try { st = statSync(p); } catch { continue; }
    if (st.isDirectory()) out.push(...listFiles(p, base));
    else out.push({ name: p.slice(base.length).replace(/^[\\/]/, ""), size: st.size, ext: extname(f).toLowerCase() });
  }
  return out;
}

/** 把前端的参数整理成任务配置 */
function buildConfig(task, files, opts, mode) {
  const abs = (f) => join(UP_DIR, safeFileName(f));
  const common = { task, options: { ...opts } };

  if (task === "reconcile") {
    if (files.length < 2) throw new Error("对账需要两个文件:先选「左表(A)」再选「右表(B)」");
    return { ...common, inputs: { a: abs(files[0]), b: abs(files[1]) }, output: join(OUT_DIR, "对账结果.xlsx") };
  }
  if (task === "merge") {
    if (!files.length) throw new Error("请先选择要合并的文件");
    return { ...common, inputs: files.map(abs), output: join(OUT_DIR, "合并结果.xlsx") };
  }
  if (task === "split") {
    if (files.length !== 1) throw new Error("拆分只接受一个输入文件");
    return { ...common, inputs: [abs(files[0])], outputDir: join(OUT_DIR, `拆分结果-${Date.now()}`) };
  }
  if (task === "clean") {
    if (!files.length) throw new Error("请先选择要清洗的文件");
    return { ...common, inputs: files.map(abs), output: files.length === 1 ? join(OUT_DIR, "清洗后.xlsx") : undefined };
  }
  if (task === "summarize") {
    if (files.length !== 1) throw new Error("汇总只接受一个输入文件");
    return { ...common, inputs: [abs(files[0])], output: join(OUT_DIR, "汇总结果.xlsx") };
  }
  throw new Error(`不支持的任务:${task}`);
}

/* ============================ HTTP ============================ */
const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const path = url.pathname;
  const method = req.method;

  try {
    /* ---- 上传 ---- */
    if (method === "POST" && path === "/api/upload") {
      const ct = req.headers["content-type"] ?? "";
      const m = /boundary=(?:"([^"]+)"|([^;]+))/i.exec(ct);
      if (!m) return sendJson(res, { ok: false, error: "缺少 multipart boundary" }, 400);
      const body = await readBody(req, MAX_UPLOAD);
      const parts = parseMultipart(body, (m[1] ?? m[2]).trim());
      mkdirSync(UP_DIR, { recursive: true });     // 目录被清掉也能自愈
      const saved = [], rejected = [];
      for (const p of parts) {
        if (!p.filename) continue;
        const name = safeFileName(p.filename);
        const ext = extname(name).toLowerCase();
        if (!ALLOWED_EXT.has(ext)) { rejected.push({ name, reason: `不支持的格式 ${ext || "(无扩展名)"}` }); continue; }
        const final = uniqueName(UP_DIR, name);
        writeFileSync(join(UP_DIR, final), p.data);
        saved.push({ name: final, size: p.data.length, ext });
      }
      return sendJson(res, { ok: true, saved, rejected, files: listFiles(UP_DIR) });
    }

    /* ---- 文件列表 / 清空 ---- */
    if (method === "GET" && path === "/api/files") return sendJson(res, { ok: true, files: listFiles(UP_DIR), outputs: listFiles(OUT_DIR) });
    if (method === "POST" && path === "/api/clear") {
      for (const f of readdirSync(UP_DIR)) rmSync(join(UP_DIR, f), { force: true });
      for (const f of readdirSync(OUT_DIR)) rmSync(join(OUT_DIR, f), { recursive: true, force: true });
      return sendJson(res, { ok: true });
    }

    /* ---- 预览上传文件的表头(便于用户确认选对了文件)---- */
    if (method === "POST" && path === "/api/preview") {
      const b = JSON.parse((await readBody(req, 1e6)).toString("utf8") || "{}");
      const { loadTable } = await import("./lib/table.mjs");
      const f = safeJoin(UP_DIR, b.name);
      if (!f || !existsSync(f)) return sendJson(res, { ok: false, error: "文件不存在" }, 404);
      const t = loadTable(f);
      return sendJson(res, { ok: true, headers: t.headers, rows: t.rows.slice(0, 5), total: t.rows.length, encoding: t.meta.encoding, sheet: t.meta.sheet, sheets: t.meta.sheets });
    }

    /* ---- 执行任务 ---- */
    if (method === "POST" && path === "/api/run") {
      const b = JSON.parse((await readBody(req, 1e6)).toString("utf8") || "{}");
      const fn = TASKS[b.task];
      if (!fn) return sendJson(res, { ok: false, error: `不支持的任务:${b.task}` }, 400);
      let cfg;
      try { cfg = buildConfig(b.task, b.files ?? [], b.options ?? {}); }
      catch (e) { return sendJson(res, { ok: false, error: e.message }, 400); }

      const t0 = Date.now();
      let r;
      try { r = fn(cfg, { cwd: ROOT, outDir: OUT_DIR }); }
      catch (e) { return sendJson(res, { ok: false, error: e.message }, 400); }

      // 汇总产出文件(单个文件 或 目录里的多个文件)
      const outputs = [];
      const collect = (p) => {
        if (!p) return;
        if (!existsSync(p)) return;
        const st = statSync(p);
        if (st.isDirectory()) { for (const f of readdirSync(p)) collect(join(p, f)); }
        else outputs.push({ name: basename(p), path: p, size: st.size });
      };
      collect(r.output);

      return sendJson(res, { ok: true, ms: Date.now() - t0, summary: r.summary, details: (r.details ?? []).slice(0, 100), outputs });
    }

    /* ---- 下载 ---- */
    if (method === "GET" && path === "/api/download") {
      const name = url.searchParams.get("name") ?? "";
      const from = url.searchParams.get("from") ?? "out";
      const rootDir = from === "up" ? UP_DIR : OUT_DIR;
      const f = safeJoin(rootDir, name);
      if (!f || !existsSync(f) || !statSync(f).isFile()) return send(res, 404, "404 Not Found");
      const buf = readFileSync(f);
      return send(res, 200, buf, {
        "content-type": MIME[extname(f).toLowerCase()] ?? "application/octet-stream",
        "content-disposition": `attachment; filename*=UTF-8''${encodeURIComponent(basename(f))}`,
      });
    }

    /* ---- 前端静态文件(只服务 public/) ---- */
    if (method === "GET") {
      let rel = path === "/" ? "/index.html" : path;
      try { rel = decodeURIComponent(rel); } catch { return send(res, 400, "400 Bad Request"); }
      const f = safeJoin(PUBLIC_DIR, rel);
      if (!f || !existsSync(f) || !statSync(f).isFile()) return send(res, 404, "404 Not Found");
      return send(res, 200, readFileSync(f), { "content-type": MIME[extname(f).toLowerCase()] ?? "application/octet-stream" });
    }

    return sendJson(res, { ok: false, error: `未知接口:${method} ${path}` }, 404);
  } catch (e) {
    return sendJson(res, { ok: false, error: e.message }, 500);
  }
});

server.listen(PORT, "127.0.0.1", () => {
  console.log("");
  console.log("  ==================================================");
  console.log("    表格管家 · 网页版");
  console.log("  ==================================================");
  console.log(`    打开   :  http://127.0.0.1:${PORT}/`);
  console.log(`    绑定   :  127.0.0.1  (仅本机;局域网访问不到)`);
  console.log(`    上传区 :  ./webdata/uploads/`);
  console.log(`    输出区 :  ./webdata/out/`);
  console.log("  ==================================================");
  console.log("    上传的表格只存在你自己电脑上,不会发到任何服务器。");
  console.log("    停止:按 Ctrl+C");
  console.log("");
});

server.on("error", (e) => {
  if (e.code === "EADDRINUSE") {
    console.error(`\n  [错误] 端口 ${PORT} 已被占用。可能已经在运行了 —— 直接打开 http://127.0.0.1:${PORT}/\n`);
  } else {
    console.error("\n  [错误] " + e.message + "\n");
  }
  process.exit(1);
});
