import XLSX from "xlsx";

const NOCO_HOST = String(process.env.NOCO_HOST || "").replace(/\/+$/, "");
const NOCO_TOKEN = String(process.env.NOCO_TOKEN || "").trim();
const SOURCE_TABLE_ID = String(process.env.SOURCE_TABLE_ID || "").trim();
const SOURCE_VIEW_ID = String(process.env.SOURCE_VIEW_ID || "").trim();
const WORKER_BASE_URL = String(process.env.WORKER_BASE_URL || "").replace(/\/+$/, "");
const MAX_SOURCE_RECORDS = Math.max(1, Number(process.env.MAX_SOURCE_RECORDS || 5000));
const SCAN_CONCURRENCY = Math.max(1, Number(process.env.SCAN_CONCURRENCY || 8));

if (!NOCO_HOST || !NOCO_TOKEN || !SOURCE_TABLE_ID || !WORKER_BASE_URL) {
  throw new Error("Missing required env: NOCO_HOST, NOCO_TOKEN, SOURCE_TABLE_ID, WORKER_BASE_URL");
}

function safe(v) {
  return String(v ?? "").trim();
}
function normalizeCompact(v) {
  return String(v || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}
function isTestBenValue(v) {
  const s = normalizeCompact(v);
  return s.includes("testben") || s.includes("testbend");
}
function pickByAliases(obj, aliases) {
  if (!obj || typeof obj !== "object") return "";
  for (const key of Object.keys(obj)) {
    const nk = normalizeCompact(key);
    for (const a of aliases) {
      if (nk === normalizeCompact(a)) {
        const val = obj[key];
        if (val != null && String(val).trim() !== "") return val;
      }
    }
  }
  return "";
}
function buildRowKey(taskCode, sheetName, excelRowIndex, fileUrl) {
  const t = safe(taskCode);
  const s = safe(sheetName);
  const i = String(Number(excelRowIndex) || 0);
  if (t) return `${t}__${s}__${i}`;
  return `${safe(fileUrl)}__${s}__${i}`;
}

async function fetchJson(url, options = {}) {
  const resp = await fetch(url, options);
  const text = await resp.text();
  if (!resp.ok) throw new Error(`HTTP ${resp.status} from ${url}\n${text.slice(0, 500)}`);
  try { return JSON.parse(text); } catch { return {}; }
}

async function listSourceRecords() {
  const out = [];
  const base = `${NOCO_HOST}/api/v2/tables/${SOURCE_TABLE_ID}/records`;
  let offset = 0;
  const limit = 200;

  while (out.length < MAX_SOURCE_RECORDS) {
    const u = new URL(base);
    u.searchParams.set("offset", String(offset));
    u.searchParams.set("limit", String(limit));
    u.searchParams.set("where", "");
    if (SOURCE_VIEW_ID) u.searchParams.set("viewId", SOURCE_VIEW_ID);

    const data = await fetchJson(u.toString(), {
      method: "GET",
      headers: { "xc-token": NOCO_TOKEN, "xc-auth": NOCO_TOKEN, Accept: "application/json" },
    });

    const list = Array.isArray(data?.list) ? data.list : Array.isArray(data) ? data : [];
    if (!list.length) break;
    out.push(...list);
    if (list.length < limit) break;
    offset += limit;
  }

  return out.slice(0, MAX_SOURCE_RECORDS);
}

function unwrapFields(rec) {
  return rec?.fields && typeof rec.fields === "object" ? rec.fields : (rec || {});
}

function buildDriveCandidates(url) {
  const out = [url];
  try {
    const u = new URL(url);
    if (!/drive\.google\.com$/i.test(u.hostname)) return out;
    const id = u.searchParams.get("id");
    if (id) {
      out.push(`https://drive.usercontent.google.com/download?id=${encodeURIComponent(id)}&export=download`);
      out.push(`https://docs.google.com/uc?export=download&id=${encodeURIComponent(id)}`);
    }
  } catch {}
  return Array.from(new Set(out));
}

async function downloadExcelArrayBuffer(fileUrl) {
  const viaWorker = (url) => `${WORKER_BASE_URL}?fileUrl=${encodeURIComponent(url)}`;
  const candidates = buildDriveCandidates(fileUrl).map(viaWorker);

  let lastErr = "";
  for (const c of candidates) {
    try {
      const resp = await fetch(c);
      if (!resp.ok) {
        lastErr = `HTTP ${resp.status} from ${c}`;
        continue;
      }
      return await resp.arrayBuffer();
    } catch (e) {
      lastErr = String(e?.message || e || "");
    }
  }
  throw new Error(`Cannot download excel: ${lastErr}`);
}

function getCellText(ws, r, c) {
  const cell = ws[XLSX.utils.encode_cell({ r, c })];
  if (!cell) return "";
  if (cell.w != null && String(cell.w).trim() !== "") return String(cell.w).trim();
  if (cell.v == null) return "";
  return String(cell.v).trim();
}

function findHeaderRowAndDanhGia(ws, range) {
  const maxProbe = Math.min(range.e.r, range.s.r + 20);
  const aliases = ["Đánh giá", "Danh gia"];
  for (let r = range.s.r; r <= maxProbe; r += 1) {
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c += 1) headers.push(getCellText(ws, r, c));
    let idx = -1;
    for (let i = 0; i < headers.length; i += 1) {
      const h = normalizeCompact(headers[i]);
      if (aliases.some((a) => h === normalizeCompact(a))) { idx = i; break; }
    }
    if (idx >= 0) return { headerRow: r, headers, danhGiaCol: range.s.c + idx };
  }
  return null;
}

function rowToDisplayObj(ws, rowIndex, headerRow, startCol, endCol) {
  const out = {};
  const headerMap = {};
  for (let c = startCol; c <= endCol; c += 1) {
    headerMap[normalizeCompact(getCellText(ws, headerRow, c))] = c;
  }
  const keys = [
    "STT",
    "Công chuẩn",
    "Mã danh mục",
    "Hạng mục kiểm tra (Index)",
    "Tiêu chuẩn (Standard)",
    "Công cụ (Tool)",
    "Hướng dẫn / Phương pháp (Document)",
  ];
  for (const key of keys) {
    const col = headerMap[normalizeCompact(key)];
    out[key] = Number.isInteger(col) ? getCellText(ws, rowIndex, col) : "";
  }
  return out;
}

async function scanOneSourceRecord(rec) {
  const f = unwrapFields(rec);
  const fileUrl = safe(pickByAliases(f, ["Link BCexcel", "Link file", "File attachment"]));
  if (!fileUrl) return [];

  const taskCode = safe(pickByAliases(f, ["Mã tác vụ", "Ma tac vu", "Task code", "Task ID", "Mã"]));
  const taskName = safe(pickByAliases(f, ["Tên tác vụ", "Ten tac vu", "Task name", "Task"]));
  const assignee = safe(pickByAliases(f, ["Asignee", "Assignee", "Người phụ trách", "Nguoi phu trach"]));
  const completionActual = safe(pickByAliases(f, ["Ngày trả báo cáo tức thời", "Ngay tra bao cao tuc thoi", "Ngày hoàn thành thực tế"]));

  const ab = await downloadExcelArrayBuffer(fileUrl);
  const wb = XLSX.read(ab, { type: "array", cellStyles: false, sheetStubs: false });

  const rows = [];
  const sheetAlias = new Map([["tckt", "TCKT"], ["them", "THEM"]]);

  for (const sn of wb.SheetNames || []) {
    const canonical = sheetAlias.get(normalizeCompact(sn));
    if (!canonical) continue;

    const ws = wb.Sheets[sn];
    if (!ws || !ws["!ref"]) continue;
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const headerInfo = findHeaderRowAndDanhGia(ws, range);
    if (!headerInfo) continue;

    for (let r = headerInfo.headerRow + 1; r <= range.e.r; r += 1) {
      const danhGia = getCellText(ws, r, headerInfo.danhGiaCol);
      if (!isTestBenValue(danhGia)) continue;

      const excelRowIndex = r + 1;
      const rowData = rowToDisplayObj(ws, r, headerInfo.headerRow, range.s.c, range.e.c);

      rows.push({
        taskCode,
        taskName,
        assignee,
        completionActual,
        sheetName: canonical,
        sourceSheetName: sn,
        fileUrl,
        excelRowIndex,
        danhGiaValue: danhGia,
        rowData,
        rowKey: buildRowKey(taskCode, canonical, excelRowIndex, fileUrl),
      });
    }
  }

  return rows;
}

async function promisePool(items, worker, concurrency) {
  const out = [];
  let i = 0;
  const runners = Array.from({ length: Math.min(concurrency, items.length) }, async () => {
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
    "Mã tác vụ": safe(it.taskCode),
    "Tên tác vụ": safe(it.taskName),
    "Asignee": safe(it.assignee),
    "Ngày trả báo cáo tức thời": safe(it.completionActual),
    "Sheet": safe(it.sheetName),
    "Nguồn": safe(it.sourceSheetName),
    "Link file": safe(it.fileUrl),
    "Đánh giá": safe(it.danhGiaValue),
    "STT": safe(rd["STT"]),
    "Công chuẩn": safe(rd["Công chuẩn"]),
    "Mã danh mục": safe(rd["Mã danh mục"]),
    "Hạng mục kiểm tra (Index)": safe(rd["Hạng mục kiểm tra (Index)"]),
    "Tiêu chuẩn (Standard)": safe(rd["Tiêu chuẩn (Standard)"]),
    "Công cụ (Tool)": safe(rd["Công cụ (Tool)"]),
    "Hướng dẫn / Phương pháp (Document)": safe(rd["Hướng dẫn / Phương pháp (Document)"]),
    rowKey: safe(it.rowKey),
    excelRowIndex: Number(it.excelRowIndex || 0),
    lastScannedAt: new Date().toISOString(),
    syncSource: "gh-auto-scan",
  };
}

async function listAllTargetRowsViaWorker() {
  const out = [];
  let offset = 0;
  const limit = 200;

  for (;;) {
    const u = new URL(WORKER_BASE_URL);
    u.searchParams.set("mode", "target");
    u.searchParams.set("offset", String(offset));
    u.searchParams.set("limit", String(limit));
    u.searchParams.set("where", "");

    const data = await fetchJson(u.toString(), { method: "GET" });
    const list = Array.isArray(data?.list) ? data.list : Array.isArray(data) ? data : [];
    out.push(...list);
    if (list.length < limit) break;
    offset += limit;
  }
  return out;
}

function buildRowKeyIdMap(rows) {
  const m = {};
  for (const rec of rows) {
    const f = unwrapFields(rec);
    const id = String(rec?.Id ?? rec?.id ?? rec?._id ?? f?.Id ?? "");
    const rk = safe(f.rowKey);
    if (id && rk) m[normalizeCompact(rk)] = id;
  }
  return m;
}

async function syncToTargetViaWorker(rows) {
  const existing = await listAllTargetRowsViaWorker();
  const rowKeyMap = buildRowKeyIdMap(existing);

  const toCreate = [];
  const toUpdate = [];
  for (const it of rows) {
    const rk = normalizeCompact(it.rowKey);
    const recordId = rowKeyMap[rk];
    if (recordId) toUpdate.push({ recordId, it });
    else toCreate.push(it);
  }

  // create (batch 100)
  for (let i = 0; i < toCreate.length; i += 100) {
    const part = toCreate.slice(i, i + 100).map(buildTargetFields);
    const u = new URL(WORKER_BASE_URL);
    u.searchParams.set("mode", "target");
    u.searchParams.set("action", "create");
    await fetchJson(u.toString(), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(part),
    });
  }

  // update
  for (const uo of toUpdate) {
    const u = new URL(WORKER_BASE_URL);
    u.searchParams.set("mode", "target");
    u.searchParams.set("action", "update");
    u.searchParams.set("recordId", String(uo.recordId));
    await fetchJson(u.toString(), {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(buildTargetFields(uo.it)),
    });
  }

  return { created: toCreate.length, updated: toUpdate.length };
}

async function callPublishAuto() {
  const u = `${WORKER_BASE_URL}/snapshot/publish/auto`;
  return await fetchJson(u, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: "{}",
  });
}

async function main() {
  console.log("[scan] list source records...");
  const source = await listSourceRecords();
  console.log("[scan] source records:", source.length);

  const rowsNested = await promisePool(
    source,
    async (rec) => {
      try {
        return await scanOneSourceRecord(rec);
      } catch (e) {
        console.warn("[scan] file failed:", String(e?.message || e || ""));
        return [];
      }
    },
    SCAN_CONCURRENCY
  );

  const rows = rowsNested.flat();
  console.log("[scan] test-ben rows:", rows.length);

  const dedup = new Map();
  for (const r of rows) dedup.set(normalizeCompact(r.rowKey), r);
  const finalRows = Array.from(dedup.values());
  console.log("[scan] dedup rows:", finalRows.length);

  const syncRes = await syncToTargetViaWorker(finalRows);
  console.log("[sync] created:", syncRes.created, "updated:", syncRes.updated);

  const publishRes = await callPublishAuto();
  console.log("[publish] result:", JSON.stringify(publishRes).slice(0, 800));
}

main().catch((e) => {
  console.error("[fatal]", String(e?.message || e || ""));
  process.exit(1);
});
