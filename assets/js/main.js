(function initForm15Main(global) {
  const { CONFIG, CACHE_CONFIG } = global.Form15Config;
  const { nowMs, formatTime } = global.Form15Utils;
  const cacheService = global.Form15CacheService;
  const dataService = global.Form15DataService;
  const excelService = global.Form15ExcelService;
  const renderService = global.Form15RenderService;

  const ui = {
    demoBtn: document.getElementById("demo-btn"),
    refreshBtn: document.getElementById("refresh-btn"),
    statusBox: document.getElementById("status-box"),
    metaBox: document.getElementById("meta-box"),
    tableRoot: document.getElementById("table-root"),
  };

  let isRefreshing = false;

  async function runRefresh(options = {}) {
    const silent = !!options.silent;
    const demoMode = !!options.demoMode;
    if (isRefreshing) return;
    isRefreshing = true;
    ui.refreshBtn.disabled = true;
    if (ui.demoBtn) ui.demoBtn.disabled = true;
    const startedAt = performance.now();
    try {
      if (!silent) renderService.setStatus(ui, demoMode ? "Đang lấy dữ liệu demo..." : "Đang lấy dữ liệu từ NocoDB...", "");
      ui.metaBox.textContent = "";

      const fetchResult = await dataService.fetchAllRecords(CONFIG);
      const records = fetchResult.rows;
      const files = [];
      const demoLimit = Math.max(1, Number(CONFIG.demo?.maxFiles || 1));
      const useSampleOnly = demoMode && !!(CONFIG.demo && CONFIG.demo.sampleFileUrl);

      if (!useSampleOnly) {
        for (const record of records) {
          const fields = dataService.getRecordFields(record);
          const excelValue = dataService.getExcelFieldValue(CONFIG, fields);
          const urls = dataService.collectExcelUrls(excelValue);
          if (!urls.length) continue;
          const taskMeta = dataService.getTaskMetaFromRecord(CONFIG, record);
          for (const url of urls) {
            files.push({ url, taskMeta });
            // Demo mode: dung som ngay khi du file de scan nhanh.
            if (demoMode && files.length >= demoLimit) break;
          }
          if (demoMode && files.length >= demoLimit) break;
        }
      }

      if (!files.length) {
        renderService.renderTable(ui, []);
        renderService.setStatus(ui, "Không có file Excel hợp lệ trong field cấu hình.", "err");
        cacheService.saveCache(CACHE_CONFIG, {
          savedAt: nowMs(),
          results: [],
          metaText: "Bản ghi NocoDB: " + records.length + " | File Excel quét: 0 | Dòng Test bền tìm được: 0",
          statusText: "Không có file Excel hợp lệ trong field cấu hình.",
          statusType: "err",
        });
        return;
      }

      const filesToScan = demoMode ? (function pickDemoFiles() {
        if (CONFIG.demo && CONFIG.demo.sampleFileUrl) {
          return [{ url: CONFIG.demo.sampleFileUrl, taskMeta: { taskCode: "DEMO", taskName: "Demo scan 1 file" } }];
        }
        return files.slice(0, demoLimit);
      })() : files;

      renderService.setStatus(
        ui,
        demoMode
          ? "Đang quét DEMO... 0/" + filesToScan.length
          : "Đang quét Excel... 0/" + filesToScan.length,
        ""
      );

      const results = [];
      const scanStats = { fileErrors: 0, sheetsMatched: 0, sheetsMissingDanhGia: 0, matchedRows: 0 };
      let doneCount = 0;
      const updateProgress = () => {
        doneCount += 1;
        renderService.setStatus(
          ui,
          (demoMode ? "Đang quét DEMO... " : "Đang quét Excel... ") + doneCount + "/" + filesToScan.length,
          ""
        );
      };

      for (const file of filesToScan) {
        try {
          await excelService.scanWorkbook(CONFIG, dataService, file.url, file.taskMeta, results, updateProgress, scanStats);
        } catch (error) {
          updateProgress();
          scanStats.fileErrors += 1;
          console.warn("Scan workbook failed", file.url, error);
        }
      }

      renderService.renderTable(ui, results);
      const elapsed = Math.round(performance.now() - startedAt);
      ui.metaBox.textContent = [
        "Bản ghi NocoDB: " + records.length,
        "File Excel quét: " + filesToScan.length + (demoMode ? " (demo)" : ""),
        "Dòng Test bền tìm được: " + results.length,
        "File lỗi: " + scanStats.fileErrors,
        "Sheet có cột Đánh giá: " + scanStats.sheetsMatched,
        "Sheet thiếu cột Đánh giá: " + scanStats.sheetsMissingDanhGia,
        "Thời gian: " + elapsed + " ms",
        "NocoDB endpoint: " + fetchResult.apiPath,
        "NocoDB auth mode: " + fetchResult.authMode,
        "Sheets cấu hình: " + CONFIG.sheetNames.join(", "),
      ].join(" | ");
      renderService.setStatus(ui, demoMode ? "Hoàn tất quét DEMO 1 file." : "Hoàn tất quét dữ liệu.", "ok");
      cacheService.saveCache(CACHE_CONFIG, {
        savedAt: nowMs(),
        results,
        metaText: ui.metaBox.textContent,
        statusText: ui.statusBox.textContent,
        statusType: "ok",
      });
    } catch (error) {
      console.error(error);
      const message = String(error?.message || error || "");
      if (/Failed to fetch|Khong goi duoc NocoDB|Khong ket noi duoc NocoDB/i.test(message)) {
        renderService.setStatus(
          ui,
          "Loi ket noi NocoDB.\nKiem tra CONFIG.nocodb.host, apiPathCandidates va token.\n\n" + message,
          "err"
        );
      } else {
        renderService.setStatus(ui, "Loi: " + message, "err");
      }
    } finally {
      isRefreshing = false;
      ui.refreshBtn.disabled = false;
      if (ui.demoBtn) ui.demoBtn.disabled = false;
    }
  }

  function bootFromCacheThenRefresh() {
    const cache = cacheService.loadCache(CACHE_CONFIG);
    if (!cache) return;
    renderService.renderTable(ui, cache.results || []);
    ui.metaBox.textContent = (cache.metaText || "") + " | Cache lúc: " + formatTime(cache.savedAt);
    renderService.setStatus(ui, cache.statusText || "Đã nạp dữ liệu cache.", cache.statusType || "");
    // Khong auto full refresh luc mo trang de tranh gay nham lan voi mode DEMO.
  }

  ui.refreshBtn.addEventListener("click", () => runRefresh({ silent: false }));
  if (ui.demoBtn) {
    ui.demoBtn.addEventListener("click", () => runRefresh({ silent: false, demoMode: true }));
  }
  bootFromCacheThenRefresh();
  if (Number(CACHE_CONFIG.autoRefreshMs || 0) > 0) {
    setInterval(() => {
      if (!document.hidden) runRefresh({ silent: true });
    }, CACHE_CONFIG.autoRefreshMs);
  }
})(window);

