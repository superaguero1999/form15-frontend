(function initForm15ManualService(global) {
  const { normalizeCompact } = global.Form15Utils;
  const MANUAL_FIELDS = [
    "manual_test_start_date",
    "manual_eta_date",
    "manual_so_luong_mau",
    "manual_test_area",
    "manual_test_area_detail",
    "manual_jig_code",
    "manual_actual_done_date",
    "manual_status",
    "manual_ket_qua",
    "manual_ghi_chu",
    "manual_updated_at",
    "manual_updated_by",
    "manual_version",
  ];

  function safeText(v) {
    return String(v ?? "").trim();
  }

  function getRowKeyFieldName(config) {
    const tech = config?.syncTarget?.technicalColumns || {};
    return safeText(tech.rowKey || "rowKey");
  }

  function buildTargetRecordsBaseUrl(config) {
    const base = safeText(config?.syncTarget?.host || config?.nocodb?.host).replace(/\/+$/, "");
    const tableId = safeText(config?.syncTarget?.tableId);
    if (!base || !tableId) return "";
    return base + "/api/v2/tables/" + tableId + "/records";
  }

  function buildTargetRecordsBaseCandidates(config) {
    const explicit = safeText(config?.syncTarget?.recordsUrl);
    if (explicit) return [explicit];
    const base = safeText(config?.syncTarget?.host || config?.nocodb?.host).replace(/\/+$/, "");
    const tableId = safeText(config?.syncTarget?.tableId);
    if (!base || !tableId) return [];
    return [
      base + "/api/v2/tables/" + tableId + "/records",
      base + "/nc/api/v2/tables/" + tableId + "/records",
    ];
  }

  function buildTargetProxyUrl(config, action) {
    try {
      const proxyBase = safeText(config?.syncTarget?.proxyUrl || config?.nocodb?.proxyUrl);
      if (!proxyBase) return "";
      const u = new URL(proxyBase);
      u.searchParams.set("mode", "target");
      if (action) u.searchParams.set("action", action);
      return u.toString();
    } catch (_) {
      return "";
    }
  }

  async function fetchJson(url, options) {
    const resp = await fetch(url, options || {});
    const text = await resp.text();
    if (!resp.ok) {
      throw new Error("HTTP " + resp.status + " from " + url + (text ? "\n" + text.slice(0, 500) : ""));
    }
    try {
      return JSON.parse(text);
    } catch (_) {
      return text;
    }
  }

  function extractFields(record) {
    if (record && record.fields && typeof record.fields === "object") return record.fields;
    return record || {};
  }

  function extractRecordId(record) {
    const fields = extractFields(record);
    return safeText(
      record?.Id ??
      record?.id ??
      record?._id ??
      fields?.Id ??
      fields?.id ??
      fields?._id
    );
  }

  /** Ten cot tren NocoDB (trong body API) co the khac key noi bo (vd: "Kết quả" thay vi manual_ket_qua). */
  function getManualFieldApiKey(config, internalKey) {
    const map = config && config.syncTarget && config.syncTarget.manualFieldKeys;
    if (map && typeof map === "object" && map[internalKey] != null && String(map[internalKey]).trim() !== "") {
      return String(map[internalKey]);
    }
    return internalKey;
  }

  function pickManualFromFields(fields, config) {
    const out = {};
    for (const key of MANUAL_FIELDS) {
      const apiKey = getManualFieldApiKey(config, key);
      let val = fields[apiKey];
      if (val === undefined || val === null) val = fields[key];
      out[key] = val ?? "";
    }
    out.manual_version = Number(out.manual_version || 0);
    return out;
  }

  function loadRowKeyIdMapFromSyncCache() {
    try {
      const raw = localStorage.getItem("form15.sync.targetRowKeyMap.v1");
      if (!raw) return {};
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object" || !parsed.map || typeof parsed.map !== "object") return {};
      return parsed.map;
    } catch (_) {
      return {};
    }
  }

  async function listAllTargetRecords(config) {
    const proxyUrlBase = buildTargetProxyUrl(config, "list");
    const directBase = buildTargetRecordsBaseUrl(config);
    const viewId = safeText(config?.syncTarget?.viewId);
    const token = safeText(config?.syncTarget?.token);
    const limit = Math.min(500, Math.max(50, Number(config?.nocodb?.limit) || 100));
    let offset = 0;
    const out = [];

    for (;;) {
      let url = "";
      let options = { method: "GET", headers: {} };
      if (proxyUrlBase) {
        const u = new URL(proxyUrlBase);
        u.searchParams.set("offset", String(offset));
        u.searchParams.set("limit", String(limit));
        if (viewId) u.searchParams.set("viewId", viewId);
        url = u.toString();
      } else if (directBase) {
        const u = new URL(directBase);
        u.searchParams.set("offset", String(offset));
        u.searchParams.set("limit", String(limit));
        u.searchParams.set("where", "");
        if (viewId) u.searchParams.set("viewId", viewId);
        url = u.toString();
        options.headers = { "xc-token": token, Accept: "application/json" };
      } else {
        throw new Error("Thiếu cấu hình syncTarget để đọc dữ liệu manual.");
      }

      const data = await fetchJson(url, options);
      const list = Array.isArray(data?.list) ? data.list : Array.isArray(data) ? data : [];
      out.push(...list);
      if (list.length < limit) break;
      offset += limit;
    }
    return out;
  }

  function buildManualMapFromRecords(config, list) {
    const rowKeyField = getRowKeyFieldName(config);
    const syncRowKeyMap = loadRowKeyIdMapFromSyncCache();
    const map = {};
    const rows = Array.isArray(list) ? list : [];
    for (const rec of rows) {
      const fields = extractFields(rec);
      const rowKey = safeText(fields[rowKeyField]);
      if (!rowKey) continue;
      const syncMappedId = safeText(syncRowKeyMap[normalizeCompact(rowKey)]);
      const recordId = syncMappedId || extractRecordId(rec);
      map[rowKey] = {
        recordId,
        fields: pickManualFromFields(fields, config),
      };
    }
    return map;
  }

  async function fetchManualMap(config) {
    const list = await listAllTargetRecords(config);
    return buildManualMapFromRecords(config, list);
  }

  function validateManualPayload(payload) {
    const start = safeText(payload?.manual_test_start_date);
    const eta = safeText(payload?.manual_eta_date);
    const done = safeText(payload?.manual_actual_done_date);
    const status = safeText(payload?.manual_status);

    if (start && eta && start > eta) {
      return '"Thời gian bắt đầu" không được lớn hơn "Thời gian dự kiến hoàn thành".';
    }
    if (status.toLowerCase() === "done" && !done) {
      return '"Trạng thái" = Done thì cần nhập "Thời gian hoàn thành thực tế".';
    }
    if (start && done && done < start) {
      return '"Thời gian hoàn thành thực tế" không được nhỏ hơn "Thời gian bắt đầu".';
    }
    return "";
  }

  async function saveManualFields(config, recordId, manualPayload) {
    const validationError = validateManualPayload(manualPayload);
    if (validationError) throw new Error(validationError);
    if (!safeText(recordId)) throw new Error("Không tìm thấy recordId để lưu dữ liệu manual.");

    const payload = {};
    for (const key of MANUAL_FIELDS) {
      if (!Object.prototype.hasOwnProperty.call(manualPayload, key)) continue;
      const apiKey = getManualFieldApiKey(config, key);
      payload[apiKey] = manualPayload[key];
    }

    const proxyUrlBase = buildTargetProxyUrl(config, "update");
    const directBases = buildTargetRecordsBaseCandidates(config);
    const token = safeText(config?.syncTarget?.token);
    const rid = safeText(recordId);
    const errors = [];

    const tryRequest = async (url, options) => {
      try {
        await fetchJson(url, options);
        return true;
      } catch (e) {
        errors.push(String(e?.message || e || ""));
        return false;
      }
    };

    // Strategy 1: existing proxy update by recordId.
    if (proxyUrlBase) {
      const u = new URL(proxyUrlBase);
      u.searchParams.set("recordId", rid);
      const ok = await tryRequest(u.toString(), {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (ok) return;
    }

    // Strategy 2+: direct API fallback (handle different NocoDB patch styles).
    if (!directBases.length) {
      throw new Error("Thiếu cấu hình syncTarget.host/tableId để lưu manual.");
    }

    const ridNum = Number(rid);
    const idValue = Number.isFinite(ridNum) && rid !== "" ? ridNum : rid;
    const rowPayloadObject = Object.assign({ Id: idValue }, payload);
    const rowPayloadArray = [rowPayloadObject];

    for (const directBase of directBases) {
      const commonHeaders = {
        "Content-Type": "application/json",
        "xc-token": token,
        "xc-auth": token,
        Accept: "application/json",
      };

      const byIdOk = await tryRequest(directBase + "/" + encodeURIComponent(rid), {
        method: "PATCH",
        headers: commonHeaders,
        body: JSON.stringify(payload),
      });
      if (byIdOk) return;

      const patchObjOk = await tryRequest(directBase, {
        method: "PATCH",
        headers: commonHeaders,
        body: JSON.stringify(rowPayloadObject),
      });
      if (patchObjOk) return;

      const patchArrayOk = await tryRequest(directBase, {
        method: "PATCH",
        headers: commonHeaders,
        body: JSON.stringify(rowPayloadArray),
      });
      if (patchArrayOk) return;
    }

    throw new Error("Không lưu được manual sau nhiều chiến lược update.\n" + errors.slice(-3).join("\n\n"));
  }

  global.Form15ManualService = {
    MANUAL_FIELDS,
    fetchManualMap,
    buildManualMapFromRecords,
    saveManualFields,
  };
})(window);

