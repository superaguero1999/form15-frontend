import * as XLSX from "xlsx";

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
      return json({ error: "Not found" }, 404, cors);
    }

    const secret = s(env.SCAN_SECRET);
    if (secret) {
      const got = s(request.headers.get("x-auto-scan-secret"));
      if (got !== secret) return json({ error: "UNAUTHORIZED" }, 401, cors);
    }

    try {
      const result = await runScan(env);
      return json({ ok: true, ...result }, 200, cors);
    } catch (e) {
      return json({ ok: false, error: String(e?.message || e || "") }, 500, cors);
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
  return s(v).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "");
}
function isTestBen(v) {
  const x = norm(v);
  return x.includes("testben") || x.includes("testbend");
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
function rowKey(taskCode, sheetName, excelRowIndex, fileUrl) {
  const t = s(taskCode);
  const sh = s(sheetName);
  const idx = String(n(excelRowIndex));
  if (t) return `${t}__${sh}__${idx}`;
  return `${s(fileUrl)}__${sh}__${idx}`;
}
function unwrap(rec) {
  return rec?.fields && typeof rec.fields === "object" ? rec.fields : (rec || {});
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

async function listNocoRecords({ host, token, tableId, viewId = "", limit = 200, max = 5000 }) {
  const out = [];
  let offset = 0;
  while (out.length < max) {
    const u = new URL(`${host}/api/v2/tables/${tableId}/records`);
    u.searchParams.set("offset", String(offset));
    u.searchParams.set("limit", String(limit));
    u.searchParams.set("where", "");
    if (viewId) u.searchParams.set("viewId", viewId);

    const data = await fetchJson(u.toString(), {
      method: "GET",
      headers: { "xc-token": token, "xc-auth": token, Accept: "application/json" },
    });
    const list = Array.isArray(data?.list) ? data.list : Array.isArray(data) ? data : [];
    out.push(...list);
    if (list.length < limit) break;
    offset += limit;
  }
  return out.slice(0, max);
}

function driveCandidates(rawUrl) {
  const out = [rawUrl];
  try {
    const u = new URL(rawUrl);
    if (!/drive\.google\.com$/i.test(u.hostname)) return out;
    const id = u.searchParams.get("id");
    if (id) {
      out.push(`https://drive.usercontent.google.com/download?id=${encodeURIComponent(id)}&export=download`);
      out.push(`https://docs.google.com/uc?export=download&id=${encodeURIComponent(id)}`);
    }
  } catch {}
  return [...new Set(out)];
}

async function downloadExcel(env, fileUrl) {
  const proxy = s(env.TARGET_PROXY_URL).replace(/\/+$/, "");
  if (!proxy) throw new Error("Missing env: TARGET_PROXY_URL (used to download excel)");

  const candidates = driveCandidates(fileUrl).map((u) => `${proxy}?fileUrl=${encodeURIComponent(u)}`);
  let lastErr = "";

  for (const c of candidates) {
    try {
      const r = await fetch(c);
      if (!r.ok) {
        lastErr = `HTTP ${r.status} ${c}`;
        continue;
      }
      return await r.arrayBuffer();
    } catch (e) {
      lastErr = String(e?.message || e || "");
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

function findHeader(ws, range, aliases) {
  const maxProbe = Math.min(range.e.r, range.s.r + 20);
  for (let r = range.s.r; r <= maxProbe; r += 1) {
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c += 1) headers.push(cellText(ws, r, c));

    let idx = -1;
    for (let i = 0; i < headers.length; i += 1) {
      const h = norm(headers[i]);
      if (aliases.some((a) => h === norm(a))) {
        idx = i;
        break;
      }
    }
    if (idx >= 0) return { headerRow: r, danhGiaCol: range.s.c + idx, headers };
  }
  return null;
}

function extractDisplayRow(ws, rowIdx, headerRow, startCol, endCol) {
  const out = {};
  const wanted = [
    "STT",
    "Công chuẩn",
    "Mã danh mục",
    "Hạng mục kiểm tra (Index)",
    "Tiêu chuẩn (Standard)",
    "Công cụ (Tool)",
    "Hướng dẫn / Phương pháp (Document)",
  ];
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

async function parseOneWorkbook(env, sourceFields) {
  const fileUrl = s(pick(sourceFields, ["Link BCexcel", "Link file", "fileUrl"]));
  if (!fileUrl) return [];

  const taskCode = s(pick(sourceFields, ["Mã tác vụ", "Ma tac vu", "Task code", "Task ID"]));
  const taskName = s(pick(sourceFields, ["Tên tác vụ", "Ten tac vu", "Task name", "Task"]));
  const assignee = s(pick(sourceFields, ["Asignee", "Assignee", "Người phụ trách", "Nguoi phu trach"]));
  const completionActual = s(
    pick(sourceFields, ["Ngày trả báo cáo tức thời", "Ngay tra bao cao tuc thoi", "Ngày hoàn thành thực tế"])
  );

  const ab = await downloadExcel(env, fileUrl);
  const wb = XLSX.read(ab, { type: "array", cellStyles: false, sheetStubs: false });

  const allowedSheets = s(env.SCAN_SHEETS || "TCKT,THEM").split(",").map((x) => norm(x)).filter(Boolean);
  const rows = [];

  for (const sn of wb.SheetNames || []) {
    if (!allowedSheets.includes(norm(sn))) continue;

    const ws = wb.Sheets[sn];
    if (!ws?.["!ref"]) continue;
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const hi = findHeader(ws, range, ["Đánh giá", "Danh gia"]);
    if (!hi) continue;

    for (let r = hi.headerRow + 1; r <= range.e.r; r += 1) {
      const danhGia = cellText(ws, r, hi.danhGiaCol);
      if (!isTestBen(danhGia)) continue;

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
        rowData: extractDisplayRow(ws, r, hi.headerRow, range.s.c, range.e.c),
      });
    }
  }

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
    "Asignee": s(it.assignee),
    "Ngày trả báo cáo tức thời": s(it.completionActual),
    "Sheet": s(it.sheetName),
    "Nguồn": s(it.sourceSheetName),
    "Link file": s(it.fileUrl),
    "Đánh giá": s(it.danhGiaValue),

    "STT": s(rd["STT"]),
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

async function listTargetDirect(env) {
  const host = s(env.TARGET_NOCO_HOST).replace(/\/+$/, "");
  const tableId = s(env.TARGET_TABLE_ID);
  const token = s(env.TARGET_NOCO_TOKEN);
  if (!host || !tableId || !token) {
    throw new Error("Missing env: TARGET_NOCO_HOST, TARGET_TABLE_ID, TARGET_NOCO_TOKEN");
  }

  const out = [];
  let offset = 0;
  const limit = 200;

  while (true) {
    const u = new URL(`${host}/api/v2/tables/${tableId}/records`);
    u.searchParams.set("offset", String(offset));
    u.searchParams.set("limit", String(limit));
    u.searchParams.set("where", "");

    const data = await fetchJson(u.toString(), {
      method: "GET",
      headers: { "xc-token": token, "xc-auth": token, Accept: "application/json" },
    });

    const list = Array.isArray(data?.list) ? data.list : Array.isArray(data) ? data : [];
    out.push(...list);
    if (list.length < limit) break;
    offset += limit;
  }

  return out;
}

function mapRowKeyToId(targetRows) {
  const m = {};
  for (const rec of targetRows) {
    const f = unwrap(rec);
    const id = s(rec?.Id ?? rec?.id ?? rec?._id ?? f?.Id ?? f?.id);
    const rk = s(f.rowKey);
    if (id && rk) m[norm(rk)] = id;
  }
  return m;
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

async function syncTargetDirect(env, rows) {
  const host = s(env.TARGET_NOCO_HOST).replace(/\/+$/, "");
  const tableId = s(env.TARGET_TABLE_ID);
  const token = s(env.TARGET_NOCO_TOKEN);
  if (!host || !tableId || !token) {
    throw new Error("Missing env: TARGET_NOCO_HOST, TARGET_TABLE_ID, TARGET_NOCO_TOKEN");
  }

  const base = `${host}/api/v2/tables/${tableId}/records`;
  const existing = await listTargetDirect(env);
  const rkMap = mapRowKeyToId(existing);

  const toCreate = [];
  const toUpdate = [];
  for (const it of rows) {
    const rid = rkMap[norm(it.rowKey)];
    if (rid) toUpdate.push({ recordId: rid, row: it });
    else toCreate.push(it);
  }

  for (let i = 0; i < toCreate.length; i += 100) {
    const part = toCreate.slice(i, i + 100).map(buildTargetFields);
    await fetchJson(base, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "xc-token": token,
        "xc-auth": token,
        Accept: "application/json",
      },
      body: JSON.stringify(part),
    });
  }

  for (const uo of toUpdate) {
    await patchTargetRecord(base, token, uo.recordId, buildTargetFields(uo.row));
  }

  return { created: toCreate.length, updated: toUpdate.length, existing: existing.length };
}

async function runScan(env) {
  const host = s(env.NOCO_HOST).replace(/\/+$/, "");
  const token = s(env.NOCO_TOKEN);
  const sourceTableId = s(env.SOURCE_TABLE_ID);
  const sourceViewId = s(env.SOURCE_VIEW_ID);

  if (!host || !token || !sourceTableId) {
    throw new Error("Missing env: NOCO_HOST, NOCO_TOKEN, SOURCE_TABLE_ID");
  }

  const maxRecords = Math.max(1, n(env.MAX_SOURCE_RECORDS || 3000));
  const conc = Math.max(1, n(env.SCAN_CONCURRENCY || 6));

  const sourceRecords = await listNocoRecords({
    host,
    token,
    tableId: sourceTableId,
    viewId: sourceViewId,
    limit: 200,
    max: maxRecords,
  });

  let fileErrors = 0;
  const scanned = await pool(sourceRecords, conc, async (rec) => {
    try {
      return await parseOneWorkbook(env, unwrap(rec));
    } catch {
      fileErrors += 1;
      return [];
    }
  });

  const rows = scanned.flat();
  const dedup = new Map();
  for (const r of rows) dedup.set(norm(r.rowKey), r);
  const finalRows = [...dedup.values()];

  const sync = await syncTargetDirect(env, finalRows);

  return {
    scannedSourceRecords: sourceRecords.length,
    excelFilesErrorTotal: fileErrors,
    rowsMatchedTestBen: rows.length,
    rowsAfterDedup: finalRows.length,
    targetCreated: sync.created,
    targetUpdated: sync.updated,
    targetExisting: sync.existing,
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
