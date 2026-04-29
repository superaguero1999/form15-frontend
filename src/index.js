import * as XLSX from "xlsx";

const CODE_VERSION = "scanner-v7.2-direct-only-2026-04-29";

export default {
  async fetch(request, env) {
    const cors = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST,OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type,x-auto-scan-secret",
    };
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });

    const url = new URL(request.url);
    if (request.method !== "POST" || url.pathname !== "/scan") {
      return json({ error: "Not found", version: CODE_VERSION }, 404, cors);
    }

    const secret = s(env.SCAN_SECRET);
    if (secret) {
      const got = s(request.headers.get("x-auto-scan-secret"));
      if (got !== secret) return json({ error: "UNAUTHORIZED", version: CODE_VERSION }, 401, cors);
    }

    try {
      const result = await runScan(env);
      return json({ ok: true, version: CODE_VERSION, downloadMode: "direct_only", ...result }, 200, cors);
    } catch (e) {
      return json({ ok: false, version: CODE_VERSION, error: String(e?.message || e || "") }, 500, cors);
    }
  },
};

function s(v) {
  return String(v ?? "").trim();
}
function n(v) {
  return Number.isFinite(Number(v)) ? Number(v) : 0;
}
function norm(v) {
  return s(v)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}
function truncate(v, max = 180) {
  const x = s(v);
  return x.length <= max ? x : x.slice(0, max) + "...";
}
function isHttpUrl(v) {
  try {
    const u = new URL(s(v));
    return /^https?:$/i.test(u.protocol);
  } catch {
    return false;
  }
}
function srcStateKey(recordId) {
  return `src:${s(recordId)}`;
}

function isTestBen(v) {
  const raw = s(v);
  const x = norm(raw);
  if (!x) return false;
  if (x.includes("testben")) return true;
  if (x.includes("testbendat")) return true;
  if (x.includes("testbenkhongdat")) return true;
  if (x.includes("testbenok")) return true;
  if (x.includes("testbenfail")) return true;

  const lowerRaw = raw.toLowerCase();
  if (lowerRaw.includes("test bền")) return true;
  if (lowerRaw.includes("test-ben")) return true;
  if (lowerRaw.includes("test ben")) return true;
  if (lowerRaw.includes("test-bền")) return true;
  return false;
}

function pick(obj, keys) {
  for (const k of keys) {
    if (Object.prototype.hasOwnProperty.call(obj, k)) {
      const v = obj[k];
      if (v != null && s(v) !== "") return v;
    }
  }
  return "";
}

function pickByNorm(obj, aliases) {
  if (!obj || typeof obj !== "object") return "";
  const map = {};
  for (const k of Object.keys(obj)) map[norm(k)] = k;
  for (const a of aliases) {
    const nk = norm(a);
    if (map[nk]) {
      const val = obj[map[nk]];
      if (val != null && s(val) !== "") return val;
    }
  }
  return "";
}

function findFileUrl(sourceFields) {
  const direct = pick(sourceFields, [
    "Link BCexcel",
    "Link BC Excel",
    "Link file",
    "Link File",
    "File URL",
    "fileUrl",
    "URL",
    "Url",
    "Link",
    "Attachment URL",
    "Drive Link",
    "Google Drive Link",
  ]);
  if (isHttpUrl(direct)) return s(direct);

  const fuzzy = pickByNorm(sourceFields, [
    "link bcexcel",
    "link bc excel",
    "linkfile",
    "fileurl",
    "url",
    "link",
    "drive link",
    "google drive link",
  ]);
  if (isHttpUrl(fuzzy)) return s(fuzzy);

  for (const k of Object.keys(sourceFields || {})) {
    const v = sourceFields[k];
    if (isHttpUrl(v)) return s(v);
  }
  return "";
}

function rowKey(taskCode, sheetName, excelRowIndex, fileUrl) {
  const t = s(taskCode);
  const sh = s(sheetName);
  const idx = String(n(excelRowIndex));
  if (t) return `${t}__${sh}__${idx}`;
  return `${s(fileUrl)}__${sh}__${idx}`;
}

function unwrap(rec) {
  return rec?.fields && typeof rec.fields === "object" ? rec.fields : rec || {};
}

async function fetchJson(url, options = {}) {
  const resp = await fetch(url, options);
  const text = await resp.text();
  if (!resp.ok) throw new Error(`HTTP ${resp.status} ${url}\n${text.slice(0, 500)}`);
  try {
    return JSON.parse(text);
  } catch {
    return {};
  }
}

async function listNocoBatch({ host, token, tableId, viewId = "", offset = 0, limit = 20 }) {
  const u = new URL(`${host}/api/v2/tables/${tableId}/records`);
  u.searchParams.set("offset", String(Math.max(0, n(offset))));
  u.searchParams.set("limit", String(Math.max(1, n(limit))));
  u.searchParams.set("where", "");
  u.searchParams.set("sort", "Id");
  if (viewId) u.searchParams.set("viewId", viewId);

  const data = await fetchJson(u.toString(), {
    method: "GET",
    headers: { "xc-token": token, "xc-auth": token, Accept: "application/json" },
  });

  return Array.isArray(data?.list) ? data.list : Array.isArray(data) ? data : [];
}

async function getCursor(env) {
  if (!env.SCAN_STATE_KV) throw new Error("Missing KV binding: SCAN_STATE_KV");
  const key = s(env.SCAN_CURSOR_KEY || "form15:auto-scan:cursor:v1");
  const raw = await env.SCAN_STATE_KV.get(key);
  const cur = n(raw);
  return cur > 0 ? cur : 0;
}

async function setCursor(env, value) {
  if (!env.SCAN_STATE_KV) throw new Error("Missing KV binding: SCAN_STATE_KV");
  const key = s(env.SCAN_CURSOR_KEY || "form15:auto-scan:cursor:v1");
  await env.SCAN_STATE_KV.put(key, String(Math.max(0, n(value))));
}

function rkCacheKey(env, rowKeyValue) {
  const p = s(env.SCAN_ROWKEY_CACHE_PREFIX || "rk:");
  return `${p}${norm(rowKeyValue)}`;
}

async function getCachedTargetId(env, rowKeyValue) {
  if (!env.SCAN_STATE_KV) return "";
  const val = await env.SCAN_STATE_KV.get(rkCacheKey(env, rowKeyValue));
  return s(val);
}

async function setCachedTargetId(env, rowKeyValue, recordId) {
  if (!env.SCAN_STATE_KV) return;
  await env.SCAN_STATE_KV.put(rkCacheKey(env, rowKeyValue), s(recordId));
}

async function getSourceState(env, recordId) {
  if (!env.SCAN_STATE_KV) return "";
  return s(await env.SCAN_STATE_KV.get(srcStateKey(recordId)));
}

async function setSourceState(env, recordId, updatedAt) {
  if (!env.SCAN_STATE_KV) return;
  await env.SCAN_STATE_KV.put(srcStateKey(recordId), s(updatedAt));
}

function driveCandidates(rawUrl) {
  const out = [];
  try {
    const u = new URL(rawUrl);
    const src = u.toString();

    // Ưu tiên link gốc trước
    out.push(src);

    // drive.google.com/uc?id=...
    if (/drive\.google\.com$/i.test(u.hostname)) {
      const id = u.searchParams.get("id");
      if (id) {
        out.push(`https://drive.usercontent.google.com/download?id=${encodeURIComponent(id)}&export=download`);
      }
    }

    // docs.google.com/spreadsheets/d/<id>/edit -> export xlsx
    if (/docs\.google\.com$/i.test(u.hostname) && u.pathname.includes("/spreadsheets/d/")) {
      const m = u.pathname.match(/\/spreadsheets\/d\/([^/]+)/);
      if (m && m[1]) {
        out.push(`https://docs.google.com/spreadsheets/d/${m[1]}/export?format=xlsx`);
      }
    }
  } catch {
    // fallback giữ link thô
    out.push(String(rawUrl || ""));
  }

  // Giới hạn tối đa 2 candidates để tránh tốn tốn request
  return [...new Set(out)].filter(Boolean).slice(0, 2);
}

async function downloadExcel(env, fileUrl) {
  const MAX_MB = Math.max(1, n(env.MAX_EXCEL_MB || 5)); // tăng/giảm trong env
  const MAX_BYTES = MAX_MB * 1024 * 1024;
  const TIMEOUT_MS = Math.max(5000, n(env.DOWNLOAD_TIMEOUT_MS || 15000));

  const candidates = driveCandidates(fileUrl);
  let lastErr = "";

  async function fetchWithLimit(url) {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort("DOWNLOAD_TIMEOUT"), TIMEOUT_MS);

    try {
      const r = await fetch(url, {
        method: "GET",
        redirect: "follow",
        signal: ctrl.signal,
        headers: {
          "User-Agent": "Mozilla/5.0",
          Accept: "*/*",
        },
      });

      if (!r.ok) throw new Error(`HTTP ${r.status}`);

      // nếu không có body (hiếm), fallback arrayBuffer rồi kiểm kích thước
      if (!r.body) {
        const ab = await r.arrayBuffer();
        if (ab.byteLength > MAX_BYTES) throw new Error(`FILE_TOO_LARGE body=${ab.byteLength}`);
        return ab;
      }

      const reader = r.body.getReader();
      let received = 0;
      const chunks = [];
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        if (value) {
          received += value.byteLength;
          if (received > MAX_BYTES) {
            ctrl.abort("FILE_TOO_LARGE");
            throw new Error(`FILE_TOO_LARGE received>${MAX_BYTES}`);
          }
          chunks.push(value);
        }
      }

      const ab = new Uint8Array(received);
      let off = 0;
      for (const ch of chunks) {
        ab.set(ch, off);
        off += ch.byteLength;
      }

      // Early reject HTML (nếu Drive trả HTML lỗi)
      const headTxt = new TextDecoder("utf-8").decode(ab.slice(0, 64));
      if (headTxt.includes("<!DOCTYPE html") || headTxt.toLowerCase().includes("<html")) {
        throw new Error("NOT_EXCEL_HTML");
      }

      // Magic header kiểm tra nhanh
      const head = ab.slice(0, 4);
      const isZip = head[0] === 0x50 && head[1] === 0x4b; // PK..
      const isOle = head[0] === 0xd0 && head[1] === 0xcf; // xls cũ
      if (!isZip && !isOle) throw new Error("NOT_EXCEL_BINARY");

      return ab.buffer;
    } finally {
      clearTimeout(t);
    }
  }

  for (const c of candidates) {
    try {
      return await fetchWithLimit(c);
    } catch (e) {
      lastErr = `${String(e?.message || e || "")} @ ${c}`;
    }
  }

  throw new Error("Cannot download excel: " + lastErr);
}

function cellText(ws, r, c) {
  const cell = ws[XLSX.utils.encode_cell({ r, c })];
  if (!cell) return "";
  if (cell.w != null && s(cell.w)) return s(cell.w);
  if (cell.v == null) return "";
  return s(cell.v);
}

function headerMatches(cellValue, aliases) {
  const h = norm(cellValue);
  if (!h) return false;
  for (const a of aliases) {
    const na = norm(a);
    if (!na) continue;
    if (h === na || h.includes(na) || na.includes(h)) return true;
  }
  return false;
}

function findHeader(ws, range, aliases) {
  const maxProbe = Math.min(range.e.r, range.s.r + 40);
  for (let r = range.s.r; r <= maxProbe; r += 1) {
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c += 1) headers.push(cellText(ws, r, c));

    let idx = -1;
    for (let i = 0; i < headers.length; i += 1) {
      if (headerMatches(headers[i], aliases)) {
        idx = i;
        break;
      }
    }
    if (idx >= 0) return { headerRow: r, danhGiaCol: range.s.c + idx };
  }
  return null;
}

function findDanhGiaColumnFallback(ws, range) {
  const aliases = ["Đánh giá", "Danh gia", "Kết quả", "Ket qua", "Result", "ĐG", "DG"];
  const maxProbe = Math.min(range.e.r, range.s.r + 50);
  for (let r = range.s.r; r <= maxProbe; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const txt = cellText(ws, r, c);
      if (headerMatches(txt, aliases)) return { headerRow: r, danhGiaCol: c };
    }
  }
  return null;
}

/**
 * TỐI ƯU CPU:
 * - extractDisplayRow cũ sẽ build colMap (quét startCol..endCol) cho MỖI dòng "Test bền".
 * - code mới precompute mapping cột cần lấy cho 1 sheet (1 lần),
 *   sau đó mỗi dòng chỉ đọc đúng các cột đó.
 */
const WANTED_DISPLAY_KEYS = [
  "STT",
  "Công chuẩn",
  "Mã danh mục",
  "Hạng mục kiểm tra (Index)",
  "Tiêu chuẩn (Standard)",
  "Công cụ (Tool)",
  "Hướng dẫn / Phương pháp (Document)",
];

function buildWantedCols(ws, headerRow, startCol, endCol) {
  const headerNormToCol = {};
  for (let c = startCol; c <= endCol; c += 1) {
    const h = norm(cellText(ws, headerRow, c));
    if (h) headerNormToCol[h] = c;
  }

  const wantedCols = {};
  for (const key of WANTED_DISPLAY_KEYS) {
    const col = headerNormToCol[norm(key)];
    wantedCols[key] = Number.isInteger(col) ? col : null;
  }
  return wantedCols;
}

function buildRowDataFromWantedCols(ws, rowIdx, wantedCols) {
  const out = {};
  for (const key of WANTED_DISPLAY_KEYS) {
    const col = wantedCols[key];
    out[key] = col != null ? cellText(ws, rowIdx, col) : "";
  }
  return out;
}

function extractDisplayRow(ws, rowIdx, headerRow, startCol, endCol) {
  // giữ lại hàm cũ nếu bạn cần fallback; nhưng parse mới sẽ dùng buildWantedCols/buildRowDataFromWantedCols
  const out = {};
  const wanted = WANTED_DISPLAY_KEYS;
  const colMap = {};
  for (let c = startCol; c <= endCol; c += 1) {
    colMap[norm(cellText(ws, headerRow, c))] = c;
  }
  for (const key of wanted) {
    const col = colMap[norm(key)];
    out[key] = Number.isInteger(col) ? cellText(ws, rowIdx, col) : "";
  }
  return out;
}

function pushSample(arr, item, limit = 5) {
  if (arr.length < limit) arr.push(item);
}

async function parseOneWorkbook(env, sourceFields, dbg) {
  const fileUrl = findFileUrl(sourceFields);
  const taskCode = s(pick(sourceFields, ["Mã tác vụ", "Ma tac vu", "Task code", "Task ID"]));
  const taskName = s(pick(sourceFields, ["Tên tác vụ", "Ten tac vu", "Task name", "Task"]));
  const assignee = s(pick(sourceFields, ["Asignee", "Assignee", "Người phụ trách", "Nguoi phu trach"]));
  const completionActual = s(
    pick(sourceFields, ["Ngày trả báo cáo tức thời", "Ngay tra bao cao tuc thoi", "Ngày hoàn thành thực tế"])
  );

  if (!fileUrl) {
    dbg.fileNoUrl += 1;
    pushSample(dbg.sampleNoUrl, {
      taskCode: truncate(taskCode),
      taskName: truncate(taskName),
      keys: Object.keys(sourceFields || {}).slice(0, 12),
    });
    return [];
  }

  const ab = await downloadExcel(env, fileUrl);
  const wb = XLSX.read(ab, {
  type: "array",
  cellStyles: false,
  sheetStubs: false,
  bookVBA: false,
  bookDeps: false,
  cellDates: false,
});

  const sheetCfg = s(env.SCAN_SHEETS);
  const allowedSheets = sheetCfg ? sheetCfg.split(",").map((x) => norm(x)).filter(Boolean) : [];

  const rows = [];
  let anySheetMatched = false;
  let anyHeaderFound = false;
  let anyTestBenFound = false;

  for (const sn of wb.SheetNames || []) {
    if (allowedSheets.length > 0 && !allowedSheets.includes(norm(sn))) {
      dbg.sheetSkippedByName += 1;
      continue;
    }

    anySheetMatched = true;
    const ws = wb.Sheets[sn];
    if (!ws?.["!ref"]) {
      dbg.sheetNoRef += 1;
      continue;
    }

    const range = XLSX.utils.decode_range(ws["!ref"]);
    let hi = findHeader(ws, range, ["Đánh giá", "Danh gia", "Kết quả", "Ket qua", "Result", "ĐG", "DG"]);
    if (!hi) {
      const fb = findDanhGiaColumnFallback(ws, range);
      if (fb) hi = fb;
    }
    if (!hi) {
      dbg.sheetNoHeader += 1;
      continue;
    }

    anyHeaderFound = true;

    // PRECOMPUTE cột cho rowData: chỉ làm 1 lần/sheet
    const wantedCols = buildWantedCols(ws, hi.headerRow, range.s.c, range.e.c);

    for (let r = hi.headerRow + 1; r <= range.e.r; r += 1) {
      const danhGia = cellText(ws, r, hi.danhGiaCol);
      if (!isTestBen(danhGia)) continue;

      anyTestBenFound = true;
      const excelRowIndex = r + 1;

      rows.push({
        rowKey: rowKey(taskCode, sn, excelRowIndex, fileUrl),
        taskCode,
        taskName,
        assignee,
        completionActual,
        sheetName: sn,
        sourceSheetName: sn,
        fileUrl,
        excelRowIndex,
        danhGiaValue: danhGia,
        rowData: buildRowDataFromWantedCols(ws, r, wantedCols),
      });
    }
  }

  if (!anySheetMatched) dbg.fileNoAllowedSheet += 1;
  if (anySheetMatched && !anyHeaderFound) dbg.fileNoDanhGiaHeader += 1;
  if (anySheetMatched && anyHeaderFound && !anyTestBenFound) dbg.fileNoTestBenRows += 1;

  return rows;
}

async function pool(items, limit, worker) {
  const out = [];
  let i = 0;
  const runners = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (i < items.length) {
      const idx = i++;
      out[idx] = await worker(items[idx], idx);
    }
  });
  await Promise.all(runners);
  return out;
}

function buildTargetFields(it) {
  const rd = it.rowData || {};
  return {
    "Mã tác vụ": s(it.taskCode),
    "Tên tác vụ": s(it.taskName),
    Asignee: s(it.assignee),
    "Ngày trả báo cáo tức thời": s(it.completionActual),
    Sheet: s(it.sheetName),
    "Nguồn": s(it.sourceSheetName),
    "Link file": s(it.fileUrl),
    "Đánh giá": s(it.danhGiaValue),

    STT: s(rd["STT"]),
    "Công chuẩn": s(rd["Công chuẩn"]),
    "Mã danh mục": s(rd["Mã danh mục"]),
    "Hạng mục kiểm tra (Index)": s(rd["Hạng mục kiểm tra (Index)"]),
    "Tiêu chuẩn (Standard)": s(rd["Tiêu chuẩn (Standard)"]),
    "Công cụ (Tool)": s(rd["Công cụ (Tool)"]),
    "Hướng dẫn / Phương pháp (Document)": s(rd["Hướng dẫn / Phương pháp (Document)"]),

    rowKey: s(it.rowKey),
    excelRowIndex: n(it.excelRowIndex),
    lastScannedAt: new Date().toISOString(),
    syncSource: "cf-auto-scan",
  };
}

function targetBase(env) {
  const host = s(env.TARGET_NOCO_HOST).replace(/\/+$/, "");
  const tableId = s(env.TARGET_TABLE_ID);
  const token = s(env.TARGET_NOCO_TOKEN);
  if (!host || !tableId || !token) {
    throw new Error("Missing env: TARGET_NOCO_HOST, TARGET_TABLE_ID, TARGET_NOCO_TOKEN");
  }
  return { base: `${host}/api/v2/tables/${tableId}/records`, token };
}

async function findTargetIdByRowKey(env, rowKeyValue) {
  const { base, token } = targetBase(env);
  const rk = s(rowKeyValue).replace(/'/g, "\\'");
  const u = new URL(base);
  u.searchParams.set("limit", "1");
  u.searchParams.set("offset", "0");
  u.searchParams.set("where", `(rowKey,eq,${rk})`);

  const data = await fetchJson(u.toString(), {
    method: "GET",
    headers: { "xc-token": token, "xc-auth": token, Accept: "application/json" },
  });

  const list = Array.isArray(data?.list) ? data.list : Array.isArray(data) ? data : [];
  if (!list.length) return "";
  const rec = list[0];
  const f = unwrap(rec);
  return s(rec?.Id ?? rec?.id ?? rec?._id ?? f?.Id ?? f?.id);
}

async function patchTargetRecord(baseUrl, token, recordId, payload) {
  const rid = s(recordId);
  const ridNum = Number(rid);
  const idValue = Number.isFinite(ridNum) ? ridNum : rid;
  const patchObj = { Id: idValue, ...payload };
  const patchArr = [patchObj];

  const candidates = [
    { url: `${baseUrl}/${encodeURIComponent(rid)}`, body: JSON.stringify(payload) },
    { url: baseUrl, body: JSON.stringify(patchObj) },
    { url: baseUrl, body: JSON.stringify(patchArr) },
  ];

  let lastErr = "";
  for (const c of candidates) {
    try {
      await fetchJson(c.url, {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
          "xc-token": token,
          "xc-auth": token,
          Accept: "application/json",
        },
        body: c.body,
      });
      return;
    } catch (e) {
      lastErr = String(e?.message || e || "");
    }
  }
  throw new Error(`Target update failed recordId=${rid}: ${lastErr}`);
}

async function createTargetRecord(baseUrl, token, payload) {
  const data = await fetchJson(baseUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "xc-token": token,
      "xc-auth": token,
      Accept: "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (Array.isArray(data) && data[0]) {
    const rec = data[0];
    const f = unwrap(rec);
    return s(rec?.Id ?? rec?.id ?? rec?._id ?? f?.Id ?? f?.id);
  }
  if (data && typeof data === "object") {
    const f = unwrap(data);
    return s(data?.Id ?? data?.id ?? data?._id ?? f?.Id ?? f?.id);
  }
  return "";
}

async function syncTargetDirect(env, rows) {
  const { base, token } = targetBase(env);

  let created = 0;
  let updated = 0;
  let cacheHit = 0;
  let cacheMiss = 0;
  let lookupHit = 0;

  for (const it of rows) {
    const rk = s(it.rowKey);
    if (!rk) continue;

    const payload = buildTargetFields(it);

    let rid = await getCachedTargetId(env, rk);
    if (rid) {
      cacheHit += 1;
      try {
        await patchTargetRecord(base, token, rid, payload);
        updated += 1;
        continue;
      } catch {
        rid = "";
      }
    } else {
      cacheMiss += 1;
    }

    rid = await findTargetIdByRowKey(env, rk);
    if (rid) {
      lookupHit += 1;
      await patchTargetRecord(base, token, rid, payload);
      await setCachedTargetId(env, rk, rid);
      updated += 1;
      continue;
    }

    const newId = await createTargetRecord(base, token, payload);
    if (newId) await setCachedTargetId(env, rk, newId);
    created += 1;
  }

  return { created, updated, cacheHit, cacheMiss, lookupHit };
}

async function runScan(env) {
  const host = s(env.NOCO_HOST).replace(/\/+$/, "");
  const token = s(env.NOCO_TOKEN);
  const sourceTableId = s(env.SOURCE_TABLE_ID);
  const sourceViewId = s(env.SOURCE_VIEW_ID);

  if (!host || !token || !sourceTableId) {
    throw new Error("Missing env: NOCO_HOST, NOCO_TOKEN, SOURCE_TABLE_ID");
  }

  const batchSize = Math.max(1, n(env.BATCH_SIZE || 20));
  const conc = Math.max(1, n(env.SCAN_CONCURRENCY || 1));

  let cursor = await getCursor(env);
  let sourceRecords = await listNocoBatch({
    host,
    token,
    tableId: sourceTableId,
    viewId: sourceViewId,
    offset: cursor,
    limit: batchSize,
  });

  if (!sourceRecords.length && cursor > 0) {
    cursor = 0;
    sourceRecords = await listNocoBatch({
      host,
      token,
      tableId: sourceTableId,
      viewId: sourceViewId,
      offset: 0,
      limit: batchSize,
    });
  }

  const dbg = {
    fileNoUrl: 0,
    sheetSkippedByName: 0,
    sheetNoRef: 0,
    sheetNoHeader: 0,
    fileNoAllowedSheet: 0,
    fileNoDanhGiaHeader: 0,
    fileNoTestBenRows: 0,
    fileParseError: 0,
    sampleNoUrl: [],
    sampleParseErrors: [],
  };

  let skippedUnchanged = 0;
  let processedChanged = 0;

  const scanned = await pool(sourceRecords, conc, async (rec) => {
    const fields = unwrap(rec);

    const recId = s(rec?.Id ?? rec?.id ?? rec?._id ?? fields?.Id ?? fields?.id);
    const updatedAt = s(rec?.UpdatedAt ?? fields?.UpdatedAt ?? rec?.updatedAt ?? fields?.updatedAt);

    if (recId && updatedAt) {
      const prev = await getSourceState(env, recId);
      if (prev && prev === updatedAt) {
        skippedUnchanged += 1;
        return [];
      }
    }

    const taskCode = s(pick(fields, ["Mã tác vụ", "Ma tac vu", "Task code", "Task ID"]));
    const taskName = s(pick(fields, ["Tên tác vụ", "Ten tac vu", "Task name", "Task"]));
    const fileUrl = findFileUrl(fields);

    try {
      const rows = await parseOneWorkbook(env, fields, dbg);
      if (recId && updatedAt) await setSourceState(env, recId, updatedAt);
      processedChanged += 1;
      return rows;
    } catch (e) {
      dbg.fileParseError += 1;
      pushSample(dbg.sampleParseErrors, {
        taskCode: truncate(taskCode),
        taskName: truncate(taskName),
        fileUrl: truncate(fileUrl),
        error: truncate(String(e?.message || e || "")),
      });
      return [];
    }
  });

  const rows = scanned.flat();
  const dedup = new Map();
  for (const r of rows) dedup.set(norm(r.rowKey), r);
  const finalRows = [...dedup.values()];

const skipSync = String(env.DEBUG_SKIP_SYNC || "").toLowerCase() === "1";

let sync = { created: 0, updated: 0, cacheHit: 0, cacheMiss: 0, lookupHit: 0 };
if (!skipSync) {
  sync = await syncTargetDirect(env, finalRows);
}

return {
  batchSize,
  cursorStart: cursor,
  cursorNext: nextCursor,
  scannedSourceRecords: sourceRecords.length,
  skippedUnchanged,
  processedChanged,
  excelFilesErrorTotal: dbg.fileParseError,
  rowsMatchedTestBen: rows.length,
  rowsAfterDedup: finalRows.length,

  targetCreated: sync.created,
  targetUpdated: sync.updated,
  cacheHit: sync.cacheHit,
  cacheMiss: sync.cacheMiss,
  lookupHit: sync.lookupHit,

  skipSync,
  debug: dbg,
};

  let nextCursor = cursor + sourceRecords.length;
  if (sourceRecords.length < batchSize) nextCursor = 0;
  await setCursor(env, nextCursor);

  return {
    batchSize,
    cursorStart: cursor,
    cursorNext: nextCursor,
    scannedSourceRecords: sourceRecords.length,
    skippedUnchanged,
    processedChanged,
    excelFilesErrorTotal: dbg.fileParseError,
    rowsMatchedTestBen: rows.length,
    rowsAfterDedup: finalRows.length,
    targetCreated: sync.created,
    targetUpdated: sync.updated,
    cacheHit: sync.cacheHit,
    cacheMiss: sync.cacheMiss,
    lookupHit: sync.lookupHit,
    debug: dbg,
  };
}

function json(data, status, extra = {}) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      ...extra,
    },
  });
}
