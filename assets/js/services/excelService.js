(function initForm15ExcelService(global) {
  const { isSameLoose, normalizeCompact } = global.Form15Utils;

  function getSheetAliasMap(config) {
    const map = new Map();
    for (const sheet of config.sheetNames) map.set(normalizeCompact(sheet), sheet);
    return map;
  }

  function findDanhGiaColumnIndex(config, headers) {
    for (let i = 0; i < headers.length; i += 1) {
      for (const candidate of config.danhGiaHeaderCandidates) {
        if (isSameLoose(headers[i], candidate)) return i;
      }
    }
    return -1;
  }

  function detectHeaderRow(config, matrix) {
    const maxProbe = Math.min(matrix.length, 20);
    for (let rowIdx = 0; rowIdx < maxProbe; rowIdx += 1) {
      const row = Array.isArray(matrix[rowIdx]) ? matrix[rowIdx] : [];
      const headers = row.map((v) => String(v || "").trim());
      const danhGiaColIndex = findDanhGiaColumnIndex(config, headers);
      if (danhGiaColIndex >= 0) return { headerRowIndex: rowIdx, headers, danhGiaColIndex };
    }
    return null;
  }

  function rowToObject(headers, row) {
    const obj = {};
    for (let i = 0; i < headers.length; i += 1) {
      const key = String(headers[i] || "Cột_" + (i + 1)).trim() || "Cột_" + (i + 1);
      obj[key] = row[i] ?? "";
    }
    return obj;
  }

  function buildDriveDirectCandidates(url) {
    const out = [url];
    try {
      const u = new URL(url);
      const isGoogleDrive = /drive\.google\.com$/i.test(u.hostname);
      if (!isGoogleDrive) return out;
      const id = u.searchParams.get("id");
      if (id) {
        out.push("https://drive.usercontent.google.com/download?id=" + encodeURIComponent(id) + "&export=download");
        out.push("https://docs.google.com/uc?export=download&id=" + encodeURIComponent(id));
      }
    } catch (_) {}
    return out;
  }

  function buildProxyFileUrl(config, sourceUrl) {
    try {
      const proxyBase = String(config?.nocodb?.proxyUrl || "").trim();
      if (!proxyBase) return "";
      const u = new URL(proxyBase);
      u.searchParams.set("fileUrl", sourceUrl);
      return u.toString();
    } catch (_) {
      return "";
    }
  }

  async function downloadWorkbookArrayBuffer(config, fileUrl) {
    const directCandidates = buildDriveDirectCandidates(fileUrl);
    const allCandidates = [];
    for (const u of directCandidates) {
      allCandidates.push(u);
      const proxyU = buildProxyFileUrl(config, u);
      if (proxyU) allCandidates.push(proxyU);
    }

    let lastErr = "";
    for (const candidate of allCandidates) {
      try {
        const response = await fetch(candidate);
        if (!response.ok) {
          lastErr = "HTTP " + response.status + " from " + candidate;
          continue;
        }
        return await response.arrayBuffer();
      } catch (err) {
        lastErr = String(err?.message || err || "");
      }
    }

    throw new Error("Khong tai duoc file Excel. " + lastErr);
  }

  async function scanWorkbook(config, dataService, fileUrl, taskMeta, globalResult, progress, scanStats) {
    const arrayBuffer = await downloadWorkbookArrayBuffer(config, fileUrl);
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetAliasMap = getSheetAliasMap(config);

    for (const sheetName of workbook.SheetNames) {
      const canonicalSheet = sheetAliasMap.get(normalizeCompact(sheetName));
      if (!canonicalSheet) continue;

      const worksheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
      if (!matrix.length) continue;
      const headerInfo = detectHeaderRow(config, matrix);
      if (!headerInfo) {
        if (scanStats) scanStats.sheetsMissingDanhGia += 1;
        continue;
      }

      const headers = headerInfo.headers;
      const danhGiaColIndex = headerInfo.danhGiaColIndex;
      const rowStart = headerInfo.headerRowIndex + 1;
      if (scanStats) scanStats.sheetsMatched += 1;

      for (let rowIndex = rowStart; rowIndex < matrix.length; rowIndex += 1) {
        const row = matrix[rowIndex] || [];
        if (!dataService.isTestBenValue(config, row[danhGiaColIndex] || "")) continue;
        globalResult.push({
          taskCode: taskMeta.taskCode,
          taskName: taskMeta.taskName,
          sheetName: canonicalSheet,
          sourceSheetName: sheetName,
          fileUrl,
          rowData: rowToObject(headers, row),
        });
        if (scanStats) scanStats.matchedRows += 1;
      }
    }
    if (typeof progress === "function") progress();
  }

  global.Form15ExcelService = {
    scanWorkbook,
  };
})(window);

