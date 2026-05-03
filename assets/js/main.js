(function initForm15Main(global) {
  const { CONFIG, CACHE_CONFIG } = global.Form15Config;
  const { nowMs, formatTime } = global.Form15Utils;
  const cacheService = global.Form15CacheService;
  const dataService = global.Form15DataService;
  const excelService = global.Form15ExcelService;
  const syncService = global.Form15SyncService;
  const manualService = global.Form15ManualService;
  const renderService = global.Form15RenderService;

  const LAB_MAP_ZONES_KEY = "form15.labMap.zones.v1";
  /** Phiên bản dữ liệu kèm bản lưu zones trong localStorage — khớp với CONFIG.labMap.zonesDataVersion */
  const LAB_MAP_ZONES_SNAPSHOT_VER_KEY = "form15.labMap.zones.snapshotVersion";
  const LAB_MAP_INSET_KEY = "form15.labMap.contentInset.v1";
  const SNAPSHOT_FALLBACK_COUNT_KEY = "form15.snapshot.fallbackCount.v1";
  /** Chỉ người biết mật khẩu mới bật chỉnh vùng tay (mã nguồn client — không thay thế bảo mật server). */
  const LAB_MAP_MANUAL_PASSWORD = "Linh123456@";

  function getLabMapZonesFromConfig() {
    const z = CONFIG.labMap && CONFIG.labMap.zones;
    return Array.isArray(z) ? z.map((x) => Object.assign({}, x)) : [];
  }

  /** null = không ép phiên bản (giữ hành vi cũ). Số: phải khớp snapshot trong localStorage mới dùng cache zones. */
  function getLabMapZonesDataVersionFromConfig() {
    const v = CONFIG.labMap && CONFIG.labMap.zonesDataVersion;
    if (v === undefined || v === null) return null;
    return typeof v === "number" && !Number.isNaN(v) ? v : 1;
  }

  function migrateLabMapZonesCacheIfStale() {
    const cfgVer = getLabMapZonesDataVersionFromConfig();
    if (cfgVer === null) return;
    try {
      const raw = localStorage.getItem(LAB_MAP_ZONES_KEY);
      if (!raw) return;
      const snap = localStorage.getItem(LAB_MAP_ZONES_SNAPSHOT_VER_KEY);
      const storedVer = snap == null || snap === "" ? NaN : parseInt(snap, 10);
      if (Number.isNaN(storedVer) || storedVer !== cfgVer) {
        localStorage.removeItem(LAB_MAP_ZONES_KEY);
        localStorage.removeItem(LAB_MAP_ZONES_SNAPSHOT_VER_KEY);
      }
    } catch (_) {}
  }

  migrateLabMapZonesCacheIfStale();

  function getLabMapZonesEffective() {
    migrateLabMapZonesCacheIfStale();
    try {
      const raw = localStorage.getItem(LAB_MAP_ZONES_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed) && parsed.length) return parsed;
      }
    } catch (_) {}
    return getLabMapZonesFromConfig();
  }

  function getLabMapContentInsetEffective() {
    try {
      const raw = localStorage.getItem(LAB_MAP_INSET_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === "object") return parsed;
      }
    } catch (_) {}
    const d = (CONFIG.labMap && CONFIG.labMap.contentInset) || {};
    return {
      left: Number(d.left) || 0,
      top: Number(d.top) || 0,
      right: Number(d.right) || 0,
      bottom: Number(d.bottom) || 0,
    };
  }

  function roundRect(n) {
    return Math.round(Number(n) * 1000) / 1000;
  }

  function collectLabMapZonesFromDom(m) {
    const mapArea = m.mapAreaEl;
    if (!mapArea || !m.zonesEl) return getLabMapZonesEffective();
    const mr = mapArea.getBoundingClientRect();
    if (!mr.width || !mr.height) return getLabMapZonesEffective();
    const base = getLabMapZonesEffective();
    const out = base.map((z) => Object.assign({}, z));
    const wraps = m.zonesEl.querySelectorAll(".lab-map-zone-wrap");
    wraps.forEach((wrap) => {
      const idx = parseInt(wrap.getAttribute("data-zone-index") || "-1", 10);
      if (idx < 0 || idx >= out.length) return;
      const er = wrap.getBoundingClientRect();
      let left = ((er.left - mr.left) / mr.width) * 100;
      let top = ((er.top - mr.top) / mr.height) * 100;
      let w = (er.width / mr.width) * 100;
      let h = (er.height / mr.height) * 100;
      left = Math.max(0, Math.min(100, roundRect(left)));
      top = Math.max(0, Math.min(100, roundRect(top)));
      w = Math.max(0.5, Math.min(100 - left, roundRect(w)));
      h = Math.max(0.5, Math.min(100 - top, roundRect(h)));
      out[idx].rect = { left, top, width: w, height: h };
    });
    return out;
  }

  function formatZonesForAppConfig(zones) {
    const esc = (s) => String(s ?? "").replace(/\\/g, "\\\\").replace(/"/g, '\\"');
    const lines = zones.map((z) => {
      const r = z.rect || {};
      const hint = z.hint != null && z.hint !== "" ? '"' + esc(z.hint) + '"' : '""';
      return (
        "        { value: \"" + esc(z.value) + "\", label: \"" + esc(z.label) + "\", hint: " + hint +
        ", rect: { left: " + roundRect(r.left) + ", top: " + roundRect(r.top) +
        ", width: " + roundRect(r.width) + ", height: " + roundRect(r.height) + " } }"
      );
    });
    return "zones: [\n" + lines.join(",\n") + "\n      ]";
  }

  const ui = {
    wrapEl: document.querySelector("main.wrap"),
    refreshBtn: document.getElementById("refresh-btn"),
    refreshForceBtn: document.getElementById("refresh-force-btn"),
    labPreviewBtn: document.getElementById("lab-preview-btn"),
    statusBox: document.getElementById("status-box"),
    metaBox: document.getElementById("meta-box"),
    tableRoot: document.getElementById("table-root"),
    filterField: document.getElementById("filter-field"),
    filterValue: document.getElementById("filter-value"),
    filterClear: document.getElementById("filter-clear"),
    filterCount: document.getElementById("filter-count"),
    filterDateRange: document.getElementById("filter-date-range"),
    filterDateFrom: document.getElementById("filter-date-from"),
    filterDateTo: document.getElementById("filter-date-to"),
    filterField2: document.getElementById("filter-field-2"),
    filterValue2: document.getElementById("filter-value-2"),
    filterDateRange2: document.getElementById("filter-date-range-2"),
    filterDateFrom2: document.getElementById("filter-date-from-2"),
    filterDateTo2: document.getElementById("filter-date-to-2"),
    exportExcelBtn: document.getElementById("export-excel-btn"),
    sourceBadge: document.getElementById("source-badge"),
    ttlBadge: document.getElementById("ttl-badge"),
  };

  const DATE_FILTER_FIELDS = new Set([
    "Ngày trả báo cáo tức thời",
    "Thời gian bắt đầu",
    "Thời gian dự kiến hoàn thành",
    "Thời gian hoàn thành thực tế",
  ]);

  function isDateFilterField(field) {
    return DATE_FILTER_FIELDS.has(String(field || ""));
  }

  // Chuan hoa chuoi ngay ve dang "YYYY-MM-DD" de so sanh lexicographic.
  // Ho tro: ISO "YYYY-MM-DD[T...]", "DD/MM/YYYY", "DD-MM-YYYY".
  function normalizeDateString(v) {
    const s = String(v == null ? "" : v).trim();
    if (!s) return "";
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) {
      return m[1] + "-" + String(m[2]).padStart(2, "0") + "-" + String(m[3]).padStart(2, "0");
    }
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) {
      return m[3] + "-" + String(m[2]).padStart(2, "0") + "-" + String(m[1]).padStart(2, "0");
    }
    m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
    if (m) {
      return m[3] + "-" + String(m[2]).padStart(2, "0") + "-" + String(m[1]).padStart(2, "0");
    }
    return "";
  }

  let latestResults = [];
  let latestFilteredResults = [];
  const currentFilter = { field: "", value: "", dateFrom: "", dateTo: "" };
  const currentFilter2 = { field: "", value: "", dateFrom: "", dateTo: "" };

  let hasRefreshedThisSession = false;
  /** Đang chạy Refresh (nút hoặc tự động). */
  let isRefreshing = false;
  /** Chỉ refresh thủ công mới khóa nhập liệu/lưu; refresh nền vẫn cho thao tác bình thường. */
  let isRefreshBlockingInputs = false;
  let refreshLockModalEls = null;
  /** Modal xác nhận “Quét lại từ đầu” (tránh dùng window.confirm — hay bị chặn trong trình duyệt nhúng/preview). */
  let forceScanConfirmModalEls = null;
  let publisherLockSession = null;
  let forceScanThisRefresh = false;
  /** Buộc lấy nguồn qua quét NocoDB + Excel (bỏ snapshot), kèm xóa cache file khi user chọn. */
  let forceExcelScanFromZeroThisRefresh = false;
  let serveStaleSnapshotNotice = false;
  let latestSourceBadgeText = "Nguồn: đang tải...";
  let latestTtlBadgeText = "TTL snapshot: đang tính...";

  function updateFrameState() {
    if (!ui.wrapEl || !ui.wrapEl.classList) return;
    ui.wrapEl.classList.remove("frame-ok", "frame-refreshing");
    ui.wrapEl.classList.add(isRefreshing ? "frame-refreshing" : "frame-ok");
  }

  function getSnapshotFallbackCount() {
    try {
      const raw = localStorage.getItem(SNAPSHOT_FALLBACK_COUNT_KEY);
      const n = Number(raw || 0);
      return Number.isFinite(n) && n >= 0 ? Math.floor(n) : 0;
    } catch (_) {
      return 0;
    }
  }

  function increaseSnapshotFallbackCount() {
    const next = getSnapshotFallbackCount() + 1;
    try {
      localStorage.setItem(SNAPSHOT_FALLBACK_COUNT_KEY, String(next));
    } catch (_) {}
    return next;
  }

  function setSourceBadge(source, fallbackCount) {
    const count = Number(fallbackCount || 0);
    if (ui.sourceBadge) {
      ui.sourceBadge.classList.remove("source-snapshot", "source-fallback", "source-scan");
    }
    if (source === "snapshot") {
      latestSourceBadgeText = "Nguồn: Snapshot";
      if (ui.sourceBadge) {
        ui.sourceBadge.classList.add("source-snapshot");
        ui.sourceBadge.textContent = latestSourceBadgeText;
      }
      return;
    }
    if (source === "scan-fallback") {
      latestSourceBadgeText = "Nguồn: Fallback scan (" + count + ")";
      if (ui.sourceBadge) {
        ui.sourceBadge.classList.add("source-fallback");
        ui.sourceBadge.textContent = latestSourceBadgeText;
      }
      return;
    }
    if (source === "scan") {
      latestSourceBadgeText = "Nguồn: Scan trực tiếp";
      if (ui.sourceBadge) {
        ui.sourceBadge.classList.add("source-scan");
        ui.sourceBadge.textContent = latestSourceBadgeText;
      }
      return;
    }
    latestSourceBadgeText = "Nguồn: đang tải...";
    if (ui.sourceBadge) ui.sourceBadge.textContent = latestSourceBadgeText;
  }

  function getSnapshotMaxAgeMinutes() {
    return Math.max(1, Number((getPublisherConfig().schedule || {}).snapshotMaxAgeMinutes || 60));
  }

  function setTtlBadgeText(text, state) {
    latestTtlBadgeText = String(text || "TTL snapshot: đang tính...");
    if (!ui.ttlBadge) return;
    ui.ttlBadge.classList.remove("ttl-fresh", "ttl-warn", "ttl-expired");
    if (state) ui.ttlBadge.classList.add(state);
    ui.ttlBadge.textContent = latestTtlBadgeText;
  }

  function setTtlBadgeFromBuiltAt(builtAtIso) {
    const maxAge = getSnapshotMaxAgeMinutes();
    const raw = String(builtAtIso || "").trim();
    if (!raw) {
      setTtlBadgeText("TTL snapshot: chưa có", "ttl-expired");
      return;
    }
    const t = new Date(raw).getTime();
    if (!Number.isFinite(t)) {
      setTtlBadgeText("TTL snapshot: không xác định", "ttl-expired");
      return;
    }
    const ageMs = nowMs() - t;
    if (!Number.isFinite(ageMs) || ageMs < 0) {
      setTtlBadgeText("TTL snapshot: đang tính...", "");
      return;
    }
    const remaining = Math.max(0, maxAge - Math.floor(ageMs / 60000));
    if (remaining <= 0) {
      setTtlBadgeText("TTL snapshot: hết hạn", "ttl-expired");
      return;
    }
    if (remaining <= 10) {
      setTtlBadgeText("TTL snapshot: còn " + remaining + " phút", "ttl-warn");
      return;
    }
    setTtlBadgeText("TTL snapshot: còn " + remaining + " phút", "ttl-fresh");
  }

  function ensureRefreshLockModal() {
    if (refreshLockModalEls) return refreshLockModalEls;
    const overlay = document.createElement("div");
    overlay.className = "refresh-lock-modal";
    overlay.innerHTML = [
      '<div class="refresh-lock-card">',
      '  <h3 class="refresh-lock-title">Cần tải lại dữ liệu</h3>',
      '  <p class="refresh-lock-msg">Vui lòng bấm "Refresh dữ liệu" để cập nhật dữ liệu mới nhất. Trong lúc chưa refresh, các ô nhập và nút Lưu sẽ bị khóa.</p>',
      '  <div class="refresh-lock-actions">',
      '    <button type="button" class="btn btn-primary refresh-lock-btn">Refresh dữ liệu ngay</button>',
      "  </div>",
      "</div>",
    ].join("");
    document.body.appendChild(overlay);
    const btn = overlay.querySelector(".refresh-lock-btn");
    btn.addEventListener(
      "click",
      (e) => {
        e.preventDefault();
        e.stopPropagation();
        void runRefresh({ silent: false }).catch((err) => console.error("runRefresh:", err));
      },
      true
    );
    refreshLockModalEls = { overlay, btn };
    return refreshLockModalEls;
  }

  function openRefreshLockModal() {
    const m = ensureRefreshLockModal();
    /** Luon dua len cuoi body de nam tren moi modal khac (lab-map, meta, overlay tool...) */
    document.body.appendChild(m.overlay);
    m.overlay.classList.add("open");
    m.btn.disabled = false;
  }

  function closeRefreshLockModal() {
    if (!refreshLockModalEls) return;
    refreshLockModalEls.overlay.classList.remove("open");
  }

  function ensureForceScanConfirmModal() {
    if (forceScanConfirmModalEls) return forceScanConfirmModalEls;
    const overlay = document.createElement("div");
    overlay.className = "refresh-lock-modal force-scan-confirm-modal";
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.setAttribute("aria-labelledby", "force-scan-confirm-title");
    overlay.innerHTML = [
      '<div class="refresh-lock-card">',
      '  <h3 id="force-scan-confirm-title" class="refresh-lock-title">Quét lại từ đầu?</h3>',
      "  <p class=\"refresh-lock-msg\">Thao tác này sẽ Làm mới toàn bộ dữ liệu. Vui lòng đợi <strong style=\"font-weight:700;color:#dc2626\">10-20 phút</strong>.</p>",
      '  <div class="refresh-lock-actions">',
      '    <button type="button" class="btn force-scan-confirm-cancel">Hủy</button>',
      '    <button type="button" class="btn btn-primary force-scan-confirm-ok">Tiếp tục quét</button>',
      "  </div>",
      "</div>",
    ].join("");
    document.body.appendChild(overlay);
    const btnCancel = overlay.querySelector(".force-scan-confirm-cancel");
    const btnOk = overlay.querySelector(".force-scan-confirm-ok");
    forceScanConfirmModalEls = { overlay, btnCancel, btnOk, detachListeners: null };
    return forceScanConfirmModalEls;
  }

  function closeForceScanConfirmModal() {
    if (!forceScanConfirmModalEls) return;
    forceScanConfirmModalEls.overlay.classList.remove("open");
    if (typeof forceScanConfirmModalEls.detachListeners === "function") {
      forceScanConfirmModalEls.detachListeners();
      forceScanConfirmModalEls.detachListeners = null;
    }
  }

  function openForceScanConfirmModal(onConfirm) {
    const m = ensureForceScanConfirmModal();
    closeForceScanConfirmModal();

    const stop = (e) => {
      e.preventDefault();
      e.stopPropagation();
    };

    const onCancel = (e) => {
      stop(e);
      closeForceScanConfirmModal();
    };

    const onOk = (e) => {
      stop(e);
      closeForceScanConfirmModal();
      if (typeof onConfirm === "function") onConfirm();
    };

    const onBackdrop = (e) => {
      if (e.target === m.overlay) onCancel(e);
    };

    const onKeyDown = (e) => {
      if (e.key === "Escape") onCancel(e);
    };

    m.btnCancel.addEventListener("click", onCancel, true);
    m.btnOk.addEventListener("click", onOk, true);
    m.overlay.addEventListener("click", onBackdrop);
    document.addEventListener("keydown", onKeyDown, true);

    m.detachListeners = () => {
      m.btnCancel.removeEventListener("click", onCancel, true);
      m.btnOk.removeEventListener("click", onOk, true);
      m.overlay.removeEventListener("click", onBackdrop);
      document.removeEventListener("keydown", onKeyDown, true);
    };

    document.body.appendChild(m.overlay);
    m.overlay.classList.add("open");
    try {
      m.btnOk.focus();
    } catch (_) {}
  }

  function applyManualLock() {
    /** Khóa khi chưa refresh phiên, hoặc khi đang refresh thủ công. */
    const disabled = !hasRefreshedThisSession || isRefreshBlockingInputs;
    const nodes = document.querySelectorAll(".manual-input, .manual-save-btn, .lab-map-open-btn");
    nodes.forEach((el) => { el.disabled = disabled; });
  }

  function setEditingStatusIfReady() {
    if (!hasRefreshedThisSession || isRefreshBlockingInputs) return;
    renderService.setStatus(ui, "Đang nhập dữ liệu...", "info");
  }

  const BASE_FIELD_ACCESSORS = {
    "Mã tác vụ": (it) => it.taskCode,
    "Tên tác vụ": (it) => it.taskName,
    "Asignee": (it) => it.assignee,
    "Ngày trả báo cáo tức thời": (it) => it.completionActual,
    "Sheet": (it) => it.sheetName,
    "Link file": (it) => it.fileUrl,
    "Thời gian bắt đầu": (it) => it.manual_test_start_date,
    "Thời gian dự kiến hoàn thành": (it) => it.manual_eta_date,
    "Số lượng mẫu": (it) => it.manual_so_luong_mau,
    "Khu vực test": (it) => it.manual_test_area,
    "Chi tiết vị trí test": (it) => it.manual_test_area_detail,
    "Mã / Tên Jig test": (it) => it.manual_jig_code,
    "Thời gian hoàn thành thực tế": (it) => it.manual_actual_done_date,
    "Trạng thái": (it) => it.manual_status,
    "Kết quả": (it) => it.manual_ket_qua,
    "Ghi chú": (it) => it.manual_ghi_chu,
  };

  function getFieldAccessor(field) {
    if (Object.prototype.hasOwnProperty.call(BASE_FIELD_ACCESSORS, field)) {
      return BASE_FIELD_ACCESSORS[field];
    }
    return (it) => (it.rowData || {})[field];
  }

  function computeFilterableHeaders(rows) {
    const baseHeaders = [
      "Mã tác vụ",
      "Tên tác vụ",
      "Asignee",
      "Ngày trả báo cáo tức thời",
      "Sheet",
      "Link file",
    ];
    const manualHeaders = [
      "Thời gian bắt đầu",
      "Thời gian dự kiến hoàn thành",
      "Số lượng mẫu",
      "Khu vực test",
      "Chi tiết vị trí test",
      "Mã / Tên Jig test",
      "Thời gian hoàn thành thực tế",
      "Trạng thái",
      "Kết quả",
      "Ghi chú",
    ];
    const dynamicCols = [];
    const seen = new Set();
    for (const item of rows) {
      for (const key of Object.keys(item.rowData || {})) {
        if (!seen.has(key)) {
          seen.add(key);
          dynamicCols.push(key);
        }
      }
    }
    return baseHeaders.concat(dynamicCols).concat(manualHeaders);
  }

  function computeVisibleDynamicColumns(rows) {
    const cfg = global.Form15Config && global.Form15Config.CONFIG;
    const hideExcel = new Set(Array.isArray(cfg && cfg.tableHideExcelColumns) ? cfg.tableHideExcelColumns : []);
    const out = [];
    const seen = new Set();
    const list = Array.isArray(rows) ? rows : [];
    for (const item of list) {
      for (const key of Object.keys((item && item.rowData) || {})) {
        if (hideExcel.has(key) || seen.has(key)) continue;
        seen.add(key);
        out.push(key);
      }
    }
    return out;
  }

  function buildExportRows(rows) {
    const list = Array.isArray(rows) ? rows : [];
    const baseHeaders = ["Mã tác vụ", "Tên tác vụ", "Asignee", "Ngày trả báo cáo tức thời", "Sheet", "Link file"];
    const manualHeaders = [
      "Thời gian bắt đầu",
      "Thời gian dự kiến hoàn thành",
      "Số lượng mẫu",
      "Khu vực test",
      "Chi tiết vị trí test",
      "Mã / Tên Jig test",
      "Thời gian hoàn thành thực tế",
      "Trạng thái",
      "Kết quả",
      "Ghi chú",
    ];
    const dynamicColsVisible = computeVisibleDynamicColumns(list);
    const headers = baseHeaders.concat(dynamicColsVisible).concat(manualHeaders);
    const out = list.map((it) => {
      const row = {};
      for (const h of headers) {
        const accessor = getFieldAccessor(h);
        const v = accessor(it);
        row[h] = v == null ? "" : String(v);
      }
      return row;
    });
    return { rows: out, headers };
  }

  function exportFilteredToExcel() {
    if (typeof XLSX === "undefined" || !XLSX || !XLSX.utils) {
      renderService.setStatus(ui, "Thiếu thư viện XLSX để xuất file.", "err");
      return;
    }
    const rows = Array.isArray(latestFilteredResults) ? latestFilteredResults : [];
    if (!rows.length) {
      renderService.setStatus(ui, "Không có dữ liệu theo bộ lọc hiện tại để xuất Excel.", "warn");
      return;
    }
    const exportData = buildExportRows(rows);
    const ws = XLSX.utils.json_to_sheet(exportData.rows, { header: exportData.headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Form15_Filtered");
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
    XLSX.writeFile(wb, "form15-filtered-" + stamp + ".xlsx");
    renderService.setStatus(ui, "Đã xuất Excel theo bộ lọc hiện tại: " + rows.length + " dòng.", "ok");
  }

  function populateFilterField() {
    if (!ui.filterField) return;
    const currentValue = ui.filterField.value;
    const currentValue2 = ui.filterField2 ? ui.filterField2.value : "";
    const headers = computeFilterableHeaders(latestResults);
    const optionsHtml = ['<option value="">-- Chọn trường --</option>']
      .concat(headers.map((h) => {
        const escaped = h.replace(/"/g, "&quot;");
        return '<option value="' + escaped + '">' + escaped + "</option>";
      }))
      .join("");
    ui.filterField.innerHTML = optionsHtml;
    if (ui.filterField2) ui.filterField2.innerHTML = optionsHtml;
    if (currentValue && headers.includes(currentValue)) {
      ui.filterField.value = currentValue;
    } else {
      currentFilter.field = "";
      ui.filterField.value = "";
    }
    if (ui.filterField2) {
      if (currentValue2 && headers.includes(currentValue2)) {
        ui.filterField2.value = currentValue2;
      } else {
        currentFilter2.field = "";
        ui.filterField2.value = "";
      }
    }
    updateFilterInputMode();
  }

  // Show/hide o nhap phu hop voi kieu truong dang chon.
  function updateFilterInputMode() {
    const isDate = isDateFilterField(currentFilter.field);
    if (ui.filterValue) {
      ui.filterValue.hidden = isDate;
      if (isDate) ui.filterValue.value = "";
    }
    if (ui.filterDateRange) {
      ui.filterDateRange.hidden = !isDate;
      if (!isDate) {
        if (ui.filterDateFrom) ui.filterDateFrom.value = "";
        if (ui.filterDateTo) ui.filterDateTo.value = "";
      }
    }

    const isDate2 = isDateFilterField(currentFilter2.field);
    if (ui.filterValue2) {
      ui.filterValue2.hidden = isDate2;
      if (isDate2) ui.filterValue2.value = "";
    }
    if (ui.filterDateRange2) {
      ui.filterDateRange2.hidden = !isDate2;
      if (!isDate2) {
        if (ui.filterDateFrom2) ui.filterDateFrom2.value = "";
        if (ui.filterDateTo2) ui.filterDateTo2.value = "";
      }
    }
  }

  function hasActiveFilter() {
    const has1 = (function () {
      if (!currentFilter.field) return false;
      if (isDateFilterField(currentFilter.field)) return !!(currentFilter.dateFrom || currentFilter.dateTo);
      return !!String(currentFilter.value || "").trim();
    })();
    const has2 = (function () {
      if (!currentFilter2.field) return false;
      if (isDateFilterField(currentFilter2.field)) return !!(currentFilter2.dateFrom || currentFilter2.dateTo);
      return !!String(currentFilter2.value || "").trim();
    })();
    return has1 || has2;
  }

  function applySingleFilter(rows, filterObj) {
    const field = filterObj.field;
    if (!field) return rows;
    const accessor = getFieldAccessor(field);
    if (isDateFilterField(field)) {
      const from = normalizeDateString(filterObj.dateFrom);
      const to = normalizeDateString(filterObj.dateTo);
      if (!from && !to) return rows;
      return rows.filter((it) => {
        const iso = normalizeDateString(accessor(it));
        if (!iso) return false;
        if (from && iso < from) return false;
        if (to && iso > to) return false;
        return true;
      });
    }
    const value = String(filterObj.value || "").trim().toLowerCase();
    if (!value) return rows;
    return rows.filter((it) => {
      const raw = accessor(it);
      return String(raw == null ? "" : raw).toLowerCase().includes(value);
    });
  }

  function applyFilter() {
    let out = latestResults.slice();
    out = applySingleFilter(out, currentFilter);
    out = applySingleFilter(out, currentFilter2);
    return out;
  }

  function updateFilterCount(filteredLen) {
    if (!ui.filterCount) return;
    const total = latestResults.length;
    if (!hasActiveFilter()) {
      ui.filterCount.textContent = total ? "Tổng: " + total + " dòng" : "";
    } else {
      ui.filterCount.textContent = "Hiển thị " + filteredLen + "/" + total + " dòng";
    }
  }

  const MANUAL_EDITABLE_FIELDS = [
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
  ];

  // Thu thap gia tri nguoi dung dang nhap do tren DOM va ghi nguoc vao latestResults
  // de tranh mat du lieu khi re-render do filter/sort/... (vi renderTable ghi de innerHTML).
  function harvestManualEditsFromDom() {
    if (!ui.tableRoot) return;
    const trNodes = ui.tableRoot.querySelectorAll("tr[data-row-key]");
    if (!trNodes.length) return;
    const byKey = new Map();
    for (const r of latestResults) {
      const k = String((r && r.rowKey) || "");
      if (k) byKey.set(k, r);
    }
    trNodes.forEach((tr) => {
      const key = String(tr.getAttribute("data-row-key") || "");
      const row = byKey.get(key);
      if (!row) return;
      for (const name of MANUAL_EDITABLE_FIELDS) {
        const el = tr.querySelector('[name="' + name + '"]');
        if (!el) continue;
        row[name] = String(el.value || "");
      }
    });
  }

  async function renderFiltered(options) {
    const shouldHarvest = !options || options.harvest !== false;
    if (shouldHarvest) harvestManualEditsFromDom();
    const filtered = applyFilter();
    latestFilteredResults = filtered.slice();
    /** renderTable là async — phải await để DOM xong rồi mới applyManualLock (tránh auto-refresh / refresh mở ô nhập). */
    await renderService.renderTable(ui, filtered, {
      afterPartial: () => applyManualLock(),
    });
    updateFilterCount(filtered.length);
    applyManualLock();
  }

  async function renderResults(rows) {
    latestResults = Array.isArray(rows) ? rows.map((row, idx) => {
      if (!row || typeof row !== "object") return row;
      if (!row.rowKey) {
        const computed = typeof syncService.buildRowKey === "function"
          ? syncService.buildRowKey(row)
          : [row.fileUrl || "", row.sheetName || "", row.excelRowIndex || ""].join("__");
        row.rowKey = String(computed || "").trim() || ("__row_fallback_" + String(idx));
      }
      return row;
    }) : [];
    populateFilterField();
    await renderFiltered({ harvest: false });
  }

  /** Dòng bảng đang chọn (data-row-key) — để không xóa nhãn "Đã lưu" khi click lại cùng dòng. */
  let selectedTableRowKey = "";
  const manualRecordIdByRowKey = {};
  const manualVersionByRowKey = {};
  let metaDetailText = "";
  let metaModalEls = null;
  let labPreviewModalEls = null;
  let labMapModalEls = null;
  let labMapTargetInput = null;
  let labMapModalWired = false;
  let labMapManualUnlocked = false;

  function resetLabMapManualState(m) {
    labMapManualUnlocked = false;
    if (!m || !m.overlay) return;
    const panel = m.overlay.querySelector(".lab-map-manual-panel");
    const setBtn = m.overlay.querySelector(".lab-map-manual-set-btn");
    const pw = m.overlay.querySelector(".lab-map-manual-pw");
    const err = m.overlay.querySelector(".lab-map-manual-pw-err");
    const lockBlock = m.overlay.querySelector(".lab-map-manual-lock-block");
    const controls = m.overlay.querySelector(".lab-map-manual-controls");
    const editToggle = m.overlay.querySelector("#lab-map-edit-toggle");
    if (panel) {
      panel.hidden = true;
      panel.setAttribute("aria-hidden", "true");
    }
    if (setBtn) setBtn.setAttribute("aria-expanded", "false");
    if (pw) pw.value = "";
    if (err) {
      err.textContent = "";
      err.hidden = true;
    }
    if (lockBlock) lockBlock.hidden = false;
    if (controls) {
      controls.hidden = true;
      controls.setAttribute("aria-hidden", "true");
    }
    if (editToggle) {
      editToggle.checked = false;
      editToggle.disabled = true;
    }
    m.overlay.classList.remove("lab-map-edit");
    const saveBtns = m.overlay.querySelectorAll(
      ".lab-map-btn-save, .lab-map-btn-json, .lab-map-btn-copy, .lab-map-btn-reset"
    );
    saveBtns.forEach((b) => {
      b.disabled = true;
    });
  }

  function setLabMapManualUnlocked(m, unlocked) {
    labMapManualUnlocked = !!unlocked;
    if (!m || !m.overlay) return;
    const lockBlock = m.overlay.querySelector(".lab-map-manual-lock-block");
    const controls = m.overlay.querySelector(".lab-map-manual-controls");
    const editToggle = m.overlay.querySelector("#lab-map-edit-toggle");
    const saveBtns = m.overlay.querySelectorAll(
      ".lab-map-btn-save, .lab-map-btn-json, .lab-map-btn-copy, .lab-map-btn-reset"
    );
    if (lockBlock) lockBlock.hidden = labMapManualUnlocked;
    if (controls) {
      controls.hidden = !labMapManualUnlocked;
      controls.setAttribute("aria-hidden", labMapManualUnlocked ? "false" : "true");
    }
    if (editToggle) editToggle.disabled = !labMapManualUnlocked;
    saveBtns.forEach((b) => {
      b.disabled = !labMapManualUnlocked;
    });
  }

  function ensureLabMapModal() {
    if (labMapModalEls) return labMapModalEls;
    const overlay = document.createElement("div");
    overlay.className = "lab-map-modal";
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.innerHTML = [
      '<div class="lab-map-card">',
      '  <div class="lab-map-head">',
      '    <h3 class="lab-map-title">Chọn khu vực test</h3>',
      '    <button type="button" class="lab-map-close">Đóng</button>',
      "  </div>",
      '  <p class="lab-map-hint">Bấm vào vùng trên sơ đồ để điền mã khu vực vào ô đang chọn.</p>',
      '  <div class="lab-map-body">',
      '    <div class="lab-map-stage">',
      '      <div class="lab-map-inner">',
      '        <img class="lab-map-img" alt="Sơ đồ phòng lab" />',
      '        <div class="lab-map-map-area">',
      '          <div class="lab-map-zones"></div>',
      "        </div>",
      "      </div>",
      "    </div>",
      "  </div>",
      '  <div id="lab-map-manual-panel" class="lab-map-manual-panel" hidden aria-hidden="true">',
      '    <div class="lab-map-manual-panel-card">',
      '      <p class="lab-map-manual-lead">Thiết lập thủ công vùng bấm trên sơ đồ: chỉ sau khi nhập đúng mật khẩu mới được <strong>chỉnh vùng</strong> (kéo / đổi kích thước). Các bước chi tiết xem mục bên dưới.</p>',
      '      <div class="lab-map-manual-lock-block">',
      '        <div class="lab-map-manual-pw-row">',
      '          <label class="lab-map-manual-pw-label">Mật khẩu <input type="password" class="lab-map-manual-pw" autocomplete="off" spellcheck="false" /></label>',
      '          <button type="button" class="btn lab-map-manual-unlock">Mở khóa</button>',
      "        </div>",
      '        <p class="lab-map-manual-pw-err" role="alert" hidden></p>',
      "      </div>",
      '      <div class="lab-map-manual-controls" hidden aria-hidden="true">',
      '        <div class="lab-map-manual-controls-row">',
      '          <label class="lab-map-edit-label"><input type="checkbox" id="lab-map-edit-toggle" disabled /> Chỉnh vùng (kéo &mdash; góc phải-dưới để đổi kích thước)</label>',
      '          <button type="button" class="btn lab-map-btn-save" disabled title="Lưu vị trí đã điều chỉnh vào trình duyệt (localStorage)">Lưu vào trình duyệt</button>',
      '          <button type="button" class="btn lab-map-btn-json" disabled title="Tải file JSON">Tải JSON</button>',
      '          <button type="button" class="btn lab-map-btn-copy" disabled title="Sao chép đoạn zones dán vào appConfig.js">Sao chép code</button>',
      '          <button type="button" class="btn lab-map-btn-reset" disabled title="Xóa bản chỉnh trong trình duyệt, dùng lại config trong file">Xóa bản chỉnh</button>',
      "        </div>",
      "      </div>",
      '      <details class="lab-map-manual-guide" open>',
      '        <summary class="lab-map-manual-guide-sum">Hướng dẫn các bước (chi tiết)</summary>',
      '        <ol class="lab-map-manual-guide-list">',
      "          <li><strong>Mở panel:</strong> bấm <em>Manual Set</em> để xem mật khẩu và các thao tác.</li>",
      "          <li><strong>Mở khóa:</strong> nhập mật khẩu đúng, bấm <em>Mở khóa</em> — lúc này mới bật được <em>Chỉnh vùng</em> và các nút lưu / xuất.</li>",
      "          <li><strong>Chọn vùng (bình thường):</strong> không tick <em>Chỉnh vùng</em>; bấm trực tiếp lên khung trên sơ đồ để điền mã vào ô <em>Khu vực test</em>.</li>",
      "          <li><strong>Chỉnh vùng:</strong> tick <em>Chỉnh vùng</em> — kéo khung để di chuyển, kéo ô nhỏ góc phải-dưới để đổi kích thước.</li>",
      "          <li><strong>Lưu tạm trên máy:</strong> <em>Lưu vào trình duyệt</em> (localStorage). Đóng tab hoặc đổi trình duyệt thì cần lưu lại hoặc dán vào file.</li>",
      "          <li><strong>Đưa vào mã nguồn:</strong> <em>Sao chép code</em> rồi dán thay mảng <code>zones</code> trong <code>form15/assets/js/config/appConfig.js</code>; hoặc <em>Tải JSON</em> để giữ bản dự phòng.</li>",
      "          <li><strong>Hoàn tác bản lưu tạm:</strong> <em>Xóa bản chỉnh</em> để dùng lại <code>zones</code> trong file config (nếu đã lưu localStorage trước đó).</li>",
      "          <li><strong>Đóng modal:</strong> mật khẩu và trạng thái mở khóa được reset — cần nhập lại nếu mở <em>Manual Set</em> sau đó.</li>",
      "        </ol>",
      "      </details>",
      "    </div>",
      "  </div>",
      '  <div class="lab-map-footer lab-map-toolbar-compact">',
      '    <button type="button" class="btn lab-map-manual-set-btn" aria-expanded="false" aria-controls="lab-map-manual-panel">Manual Set</button>',
      "  </div>",
      "</div>",
    ].join("");
    document.body.appendChild(overlay);
    const closeBtn = overlay.querySelector(".lab-map-close");
    const img = overlay.querySelector(".lab-map-img");
    const mapAreaEl = overlay.querySelector(".lab-map-map-area");
    const zonesEl = overlay.querySelector(".lab-map-zones");
    const close = () => {
      overlay.classList.remove("open");
      labMapTargetInput = null;
      resetLabMapManualState(labMapModalEls);
    };
    closeBtn.addEventListener("click", close);
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) close();
    });
    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      if (!overlay.classList.contains("open")) return;
      close();
    });
    img.addEventListener("load", () => {
      img.alt = "Sơ đồ phòng lab";
    });
    img.addEventListener("error", () => {
      img.alt = "Không tải được ảnh. Kiểm tra CONFIG.labMap.imageUrl và tên file trong thư mục assets/lab-map/.";
    });
    labMapModalEls = { overlay, img, mapAreaEl, zonesEl, close };
    return labMapModalEls;
  }

  function applyLabMapContentInset(m) {
    const inset = getLabMapContentInsetEffective();
    const l = Math.max(0, Number(inset.left) || 0);
    const t = Math.max(0, Number(inset.top) || 0);
    const r = Math.max(0, Number(inset.right) || 0);
    const b = Math.max(0, Number(inset.bottom) || 0);
    const ma = m.mapAreaEl;
    if (!ma) return;
    const w = Math.max(0, 100 - l - r);
    const h = Math.max(0, 100 - t - b);
    ma.style.left = l + "%";
    ma.style.top = t + "%";
    ma.style.width = w + "%";
    ma.style.height = h + "%";
  }

  function buildLabMapZones(m) {
    const zones = getLabMapZonesEffective();
    m.zonesEl.innerHTML = "";
    const editMode = m.overlay.classList.contains("lab-map-edit");
    zones.forEach((z, i) => {
      if (!z || !z.rect) return;
      const wrap = document.createElement("div");
      wrap.className = "lab-map-zone-wrap";
      wrap.setAttribute("data-zone-index", String(i));
      wrap.style.left = Number(z.rect.left) + "%";
      wrap.style.top = Number(z.rect.top) + "%";
      wrap.style.width = Number(z.rect.width) + "%";
      wrap.style.height = Number(z.rect.height) + "%";
      const face = document.createElement("div");
      face.className = "lab-map-zone";
      face.setAttribute("role", "button");
      face.setAttribute("tabindex", "0");
      const val = String(z.value || "").trim();
      const label = String(z.label != null ? z.label : z.value || "").trim();
      const hint = String(z.hint || "").trim();
      face.title = hint || label || val;
      face.setAttribute("data-value", val);
      face.textContent = label || val;
      const handle = document.createElement("div");
      handle.className = "lab-map-zone-handle";
      handle.title = "Kéo góc để đổi kích thước";
      handle.setAttribute("aria-hidden", "true");
      wrap.appendChild(face);
      wrap.appendChild(handle);
      m.zonesEl.appendChild(wrap);
    });
    if (!editMode) {
      m.overlay.classList.remove("lab-map-edit");
    }
  }

  function wireLabMapToolbarOnce() {
    if (labMapModalWired) return;
    const m = labMapModalEls;
    if (!m || !m.overlay) return;
    labMapModalWired = true;
    const btnManualSet = m.overlay.querySelector(".lab-map-manual-set-btn");
    const manualPanel = m.overlay.querySelector(".lab-map-manual-panel");
    const btnUnlock = m.overlay.querySelector(".lab-map-manual-unlock");
    const pwInput = m.overlay.querySelector(".lab-map-manual-pw");
    const pwErr = m.overlay.querySelector(".lab-map-manual-pw-err");
    const editToggle = m.overlay.querySelector("#lab-map-edit-toggle");
    const btnSave = m.overlay.querySelector(".lab-map-btn-save");
    const btnJson = m.overlay.querySelector(".lab-map-btn-json");
    const btnCopy = m.overlay.querySelector(".lab-map-btn-copy");
    const btnReset = m.overlay.querySelector(".lab-map-btn-reset");

    if (btnManualSet && manualPanel) {
      btnManualSet.addEventListener("click", () => {
        const show = manualPanel.hidden;
        manualPanel.hidden = !show;
        manualPanel.setAttribute("aria-hidden", show ? "false" : "true");
        btnManualSet.setAttribute("aria-expanded", show ? "true" : "false");
      });
    }

    function tryUnlockManual() {
      const v = pwInput && pwInput.value != null ? String(pwInput.value) : "";
      if (v === LAB_MAP_MANUAL_PASSWORD) {
        setLabMapManualUnlocked(m, true);
        if (pwErr) {
          pwErr.textContent = "";
          pwErr.hidden = true;
        }
        renderService.setStatus(ui, "Đã mở khóa Manual Set.", "ok");
      } else {
        if (pwErr) {
          pwErr.textContent = "Mật khẩu không đúng.";
          pwErr.hidden = false;
        }
      }
    }
    if (btnUnlock) btnUnlock.addEventListener("click", tryUnlockManual);
    if (pwInput) {
      pwInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") tryUnlockManual();
      });
    }

    if (editToggle) {
      editToggle.addEventListener("change", () => {
        if (!labMapManualUnlocked) {
          editToggle.checked = false;
          return;
        }
        if (editToggle.checked) {
          m.overlay.classList.add("lab-map-edit");
        } else {
          m.overlay.classList.remove("lab-map-edit");
        }
        buildLabMapZones(m);
      });
    }

    if (btnSave) {
      btnSave.addEventListener("click", () => {
        if (!labMapManualUnlocked) return;
        const zones = collectLabMapZonesFromDom(m);
        try {
          localStorage.setItem(LAB_MAP_ZONES_KEY, JSON.stringify(zones));
          const ver = getLabMapZonesDataVersionFromConfig();
          if (ver !== null) {
            localStorage.setItem(LAB_MAP_ZONES_SNAPSHOT_VER_KEY, String(ver));
          }
          renderService.setStatus(ui, "Đã lưu vị trí vùng vào trình duyệt (localStorage).", "ok");
        } catch (e) {
          renderService.setStatus(ui, "Không lưu được: " + String(e?.message || e || ""), "err");
        }
      });
    }

    if (btnJson) {
      btnJson.addEventListener("click", () => {
        if (!labMapManualUnlocked) return;
        const zones = collectLabMapZonesFromDom(m);
        const blob = new Blob([JSON.stringify(zones, null, 2)], { type: "application/json;charset=utf-8" });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "lab-map-zones.json";
        a.click();
        URL.revokeObjectURL(a.href);
        renderService.setStatus(ui, "Đã tải lab-map-zones.json", "ok");
      });
    }

    if (btnCopy) {
      btnCopy.addEventListener("click", () => {
        if (!labMapManualUnlocked) return;
        const zones = collectLabMapZonesFromDom(m);
        const text = formatZonesForAppConfig(zones);
        navigator.clipboard.writeText(text).then(() => {
          renderService.setStatus(ui, "Đã sao chép đoạn zones — dán vào appConfig.js (thay mảng zones cũ).", "ok");
        }).catch(() => {
          renderService.setStatus(ui, "Không sao chép được (clipboard).", "err");
        });
      });
    }

    if (btnReset) {
      btnReset.addEventListener("click", () => {
        if (!labMapManualUnlocked) return;
        try {
          localStorage.removeItem(LAB_MAP_ZONES_KEY);
          localStorage.removeItem(LAB_MAP_ZONES_SNAPSHOT_VER_KEY);
          localStorage.removeItem(LAB_MAP_INSET_KEY);
        } catch (_) {}
        if (editToggle) editToggle.checked = false;
        m.overlay.classList.remove("lab-map-edit");
        applyLabMapContentInset(m);
        buildLabMapZones(m);
        renderService.setStatus(ui, "Đã xóa bản chỉnh — dùng dữ liệu trong appConfig.js.", "ok");
      });
    }

    m.zonesEl.addEventListener("click", (e) => {
      if (m.overlay.classList.contains("lab-map-edit")) return;
      const face = e.target.closest(".lab-map-zone");
      if (!face) return;
      const val = String(face.getAttribute("data-value") || "").trim();
      if (labMapTargetInput) {
        labMapTargetInput.value = val;
        labMapTargetInput.dispatchEvent(new Event("input", { bubbles: true }));
      }
      m.close();
    });

    m.zonesEl.addEventListener("mousedown", (e) => {
      if (!labMapManualUnlocked) return;
      if (!m.overlay.classList.contains("lab-map-edit")) return;
      const mapArea = m.mapAreaEl;
      if (!mapArea) return;
      const handle = e.target.closest(".lab-map-zone-handle");
      const wrap = e.target.closest(".lab-map-zone-wrap");
      if (!wrap) return;
      e.preventDefault();
      const mr = mapArea.getBoundingClientRect();
      const wr = wrap.getBoundingClientRect();
      if (handle) {
        const startX = e.clientX;
        const startY = e.clientY;
        const startW = wr.width;
        const startH = wr.height;
        const move = (ev) => {
          const dw = ev.clientX - startX;
          const dh = ev.clientY - startY;
          const wPct = ((startW + dw) / mr.width) * 100;
          const hPct = ((startH + dh) / mr.height) * 100;
          const minW = 0.8;
          const minH = 0.8;
          wrap.style.width = Math.max(minW, wPct) + "%";
          wrap.style.height = Math.max(minH, hPct) + "%";
        };
        const up = () => {
          document.removeEventListener("mousemove", move);
          document.removeEventListener("mouseup", up);
        };
        document.addEventListener("mousemove", move);
        document.addEventListener("mouseup", up);
        return;
      }
      if (e.target.closest(".lab-map-zone")) {
        const offsetX = e.clientX - wr.left;
        const offsetY = e.clientY - wr.top;
        const move = (ev) => {
          let nx = ev.clientX - mr.left - offsetX;
          let ny = ev.clientY - mr.top - offsetY;
          const wPct = (wr.width / mr.width) * 100;
          const hPct = (wr.height / mr.height) * 100;
          let leftPct = (nx / mr.width) * 100;
          let topPct = (ny / mr.height) * 100;
          leftPct = Math.max(0, Math.min(100 - wPct, leftPct));
          topPct = Math.max(0, Math.min(100 - hPct, topPct));
          wrap.style.left = leftPct + "%";
          wrap.style.top = topPct + "%";
        };
        const up = () => {
          document.removeEventListener("mousemove", move);
          document.removeEventListener("mouseup", up);
        };
        document.addEventListener("mousemove", move);
        document.addEventListener("mouseup", up);
      }
    });
  }

  function resolveLabMapImageUrl(relativePath) {
    const p = String(relativePath || "").trim();
    if (!p) return "";
    try {
      return new URL(p, window.location.href).href;
    } catch (_) {
      return p;
    }
  }

  function openLabMapModal(targetInput) {
    const cfg = CONFIG.labMap;
    const zonesEff = getLabMapZonesEffective();
    if (!cfg || !cfg.enabled || !cfg.imageUrl || !zonesEff.length) return;
    labMapTargetInput = targetInput;
    const m = ensureLabMapModal();
    wireLabMapToolbarOnce();

    resetLabMapManualState(m);

    const url = resolveLabMapImageUrl(cfg.imageUrl);

    function afterImageReady() {
      requestAnimationFrame(() => {
        applyLabMapContentInset(m);
        buildLabMapZones(m);
      });
    }

    m.img.src = url;
    if (m.img.complete && m.img.naturalWidth > 0) {
      afterImageReady();
    } else {
      m.img.addEventListener("load", afterImageReady, { once: true });
    }
    m.overlay.classList.add("open");
  }

  function ensureMetaModal() {
    if (metaModalEls) return metaModalEls;
    const overlay = document.createElement("div");
    overlay.className = "meta-modal";
    overlay.innerHTML = [
      '<div class="meta-modal-card">',
      '  <div class="meta-modal-head">',
      '    <h3 class="meta-modal-title">Bảng chi tiết</h3>',
      '    <button type="button" class="meta-modal-close">Đóng</button>',
      "  </div>",
      '  <pre class="meta-modal-content"></pre>',
      "</div>",
    ].join("");
    document.body.appendChild(overlay);
    const closeBtn = overlay.querySelector(".meta-modal-close");
    const close = () => overlay.classList.remove("open");
    closeBtn.addEventListener("click", close);
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) close();
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") close();
    });
    metaModalEls = {
      overlay,
      content: overlay.querySelector(".meta-modal-content"),
    };
    return metaModalEls;
  }

  function ensureLabPreviewModal() {
    if (labPreviewModalEls) return labPreviewModalEls;
    const overlay = document.createElement("div");
    overlay.className = "lab-preview-modal";
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.innerHTML = [
      '<div class="lab-preview-card">',
      '  <div class="lab-preview-head">',
      '    <h3 class="lab-preview-title">Khu vực test - phòng Lab_DQA 2026</h3>',
      '    <button type="button" class="lab-preview-close">Đóng</button>',
      "  </div>",
      '  <div class="lab-preview-body">',
      '    <img class="lab-preview-img" alt="" />',
      "  </div>",
      "</div>",
    ].join("");
    document.body.appendChild(overlay);
    const titleEl = overlay.querySelector(".lab-preview-title");
    const close = () => {
      overlay.classList.remove("open");
    };
    overlay.querySelector(".lab-preview-close").addEventListener("click", close);
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) close();
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && overlay.classList.contains("open")) close();
    });
    labPreviewModalEls = {
      overlay,
      img: overlay.querySelector(".lab-preview-img"),
      titleEl,
      close,
    };
    return labPreviewModalEls;
  }

  function openLabPreviewModal() {
    const cfg = CONFIG.labMap;
    const rawPreview = cfg && String(cfg.previewImageUrl || "").trim();
    const rawImg = cfg && cfg.imageUrl ? String(cfg.imageUrl).trim() : "";
    const urlPath = rawPreview || rawImg;
    if (!cfg || !cfg.enabled || !urlPath) return;
    const m = ensureLabPreviewModal();
    const modalTitle =
      (cfg && String(cfg.previewModalTitle || "").trim()) || "Khu vực test - phòng Lab_DQA 2026";
    if (m.titleEl) m.titleEl.textContent = modalTitle;
    const url = resolveLabMapImageUrl(urlPath);
    m.img.alt = modalTitle;
    m.img.src = url;
    m.overlay.classList.add("open");
  }

  function initLabPreviewLink() {
    const btn = ui.labPreviewBtn;
    if (!btn) return;
    const cfg = CONFIG.labMap;
    const rawPreview = cfg && String(cfg.previewImageUrl || "").trim();
    const rawImg = cfg && cfg.imageUrl ? String(cfg.imageUrl).trim() : "";
    const urlPath = rawPreview || rawImg;
    if (!cfg || !cfg.enabled || !urlPath) {
      btn.hidden = true;
      return;
    }
    const label =
      (cfg && String(cfg.previewLinkLabel || "").trim()) || "Xem sơ đồ khu vực test (phòng lab)";
    btn.textContent = label;
    btn.hidden = false;
    btn.addEventListener("click", () => openLabPreviewModal());
  }

  function renderMetaSummary(parts, extraText) {
    const list = Array.isArray(parts) ? parts.filter(Boolean) : [];
    const fullParts = extraText ? list.concat([extraText]) : list;
    metaDetailText = fullParts.join(" | ");
    ui.metaBox.innerHTML = "";
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "meta-detail-btn";
    btn.textContent = "Chi tiết";
    btn.addEventListener("click", () => {
      const modal = ensureMetaModal();
      modal.content.innerHTML = formatMetaDetailHtml(metaDetailText || "Chưa có dữ liệu chi tiết.");
      modal.overlay.classList.add("open");
    });
    ui.metaBox.appendChild(btn);
  }

  function formatMetaDetailHtml(text) {
    const esc = (s) => String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
    const parts = String(text || "").split(" | ");
    const htmlParts = parts.map((p) => {
      const t = String(p || "").trim();
      const safe = esc(t);
      if (/^Thời gian:\s*\d+\s*ms$/i.test(t)) return "<strong>" + safe + "</strong>";
      return safe;
    });
    return htmlParts.join(" | ");
  }

  function addScanStats(target, source) {
    if (!target || !source) return;
    target.fileErrors += Number(source.fileErrors || 0);
    target.sheetsMatched += Number(source.sheetsMatched || 0);
    target.sheetsMissingDanhGia += Number(source.sheetsMissingDanhGia || 0);
    target.matchedRows += Number(source.matchedRows || 0);
  }

  function addSyncStats(target, source) {
    if (!target || !source) return;
    target.created += Number(source.created || 0);
    target.updated += Number(source.updated || 0);
    target.createErrors += Number(source.createErrors || 0);
    target.updateErrors += Number(source.updateErrors || 0);
  }

  function formatElapsedSince(ts) {
    const t = Number(ts || 0);
    if (!t) return "";
    const diffMs = nowMs() - t;
    if (!Number.isFinite(diffMs) || diffMs < 0) return "";
    const sec = Math.floor(diffMs / 1000);
    if (sec < 60) return sec + " giây trước";
    const min = Math.floor(sec / 60);
    if (min < 60) return min + " phút trước";
    const hour = Math.floor(min / 60);
    if (hour < 24) return hour + " giờ trước";
    const day = Math.floor(hour / 24);
    return day + " ngày trước";
  }

  function buildExcelFileStatsMetaParts(results, scannedCount, fileErrors, okWithTestBenCount) {
    const list = Array.isArray(results) ? results : [];
    const scanned = Math.max(0, Number(scannedCount || 0));
    const errors = Math.max(0, Number(fileErrors || 0));
    const okWithTestBen = Number.isFinite(Number(okWithTestBenCount))
      ? Math.max(0, Number(okWithTestBenCount))
      : (function () {
      const seen = new Set();
      for (const row of list) {
        const key = String((row && row.fileUrl) || "").trim();
        if (key) seen.add(key);
      }
      return seen.size;
    })();
    return [
      "Số file excel đã quét: " + scanned,
      "Số file excel lỗi: " + errors,
      "Số file excel OK - có trạng thái Test bền: " + okWithTestBen,
    ];
  }

  function computeExcelFilesOkWithTestBen(results) {
    const list = Array.isArray(results) ? results : [];
    const seen = new Set();
    for (const row of list) {
      const key = String((row && row.fileUrl) || "").trim();
      if (key) seen.add(key);
    }
    return seen.size;
  }

  /** Chuẩn hóa nhẹ URL để khớp Drive / Sheets khi query string khác nhau. */
  function excelUrlMatchKey(url) {
    const raw = String(url || "").trim();
    if (!raw) return "";
    try {
      const u = new URL(raw);
      const host = u.hostname.toLowerCase();
      const id = u.searchParams.get("id");
      if (id && host.includes("drive.google.com")) return "drive:" + id;
      const m = u.pathname.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (m) return "sheet:" + m[1];
      return raw.split("#")[0];
    } catch (_) {
      return raw.split("#")[0];
    }
  }

  function buildAllowedExcelUrlKeysByTaskCode(config, sourceRecords) {
    const byTask = new Map();
    const allKeys = new Set();
    const arr = Array.isArray(sourceRecords) ? sourceRecords : [];
    for (const record of arr) {
      const fields = dataService.getRecordFields(record);
      const meta = dataService.getTaskMetaFromRecord(config, record);
      const tc = String(meta.taskCode || "").trim();
      const excelVal = dataService.getExcelFieldValue(config, fields);
      const urls = dataService.collectExcelUrls(excelVal);
      for (const url of urls) {
        const k = excelUrlMatchKey(url);
        if (!k) continue;
        allKeys.add(k);
        if (tc) {
          if (!byTask.has(tc)) byTask.set(tc, new Set());
          byTask.get(tc).add(k);
        }
      }
    }
    return { byTask, allKeys };
  }

  function filterSnapshotRowsByCurrentSource(config, snapshotRows, allowed) {
    const { byTask, allKeys } = allowed;
    const rows = Array.isArray(snapshotRows) ? snapshotRows : [];
    return rows.filter((row) => {
      if (!row || typeof row !== "object") return false;
      const fk = excelUrlMatchKey(row.fileUrl);
      if (!fk) return false;
      const tc = String(row.taskCode || "").trim();
      if (tc && byTask.has(tc)) return byTask.get(tc).has(fk);
      return allKeys.has(fk);
    });
  }

  function getDataSourceMode() {
    if (forceExcelScanFromZeroThisRefresh) return "scan";
    if (forceScanThisRefresh) return "scan";
    if (serveStaleSnapshotNotice) return "snapshot";
    if (isPublisherRole()) return "scan";
    const m = String(CONFIG && CONFIG.dataSourceMode || "").trim().toLowerCase();
    if (m === "scan" || m === "snapshot" || m === "hybrid") return m;
    return "scan";
  }

  function getPublisherConfig() {
    return (CONFIG && CONFIG.publisher) || {};
  }

  function getPublisherRole() {
    const role = String(getPublisherConfig().role || "").trim().toLowerCase();
    if (role === "publisher" || role === "consumer" || role === "auto") return role;
    return "consumer";
  }

  function isPublisherRole() {
    return getPublisherRole() === "publisher";
  }

  function isPublisherAutoRole() {
    return getPublisherRole() === "auto";
  }

  function shouldAttemptPublisherFlow() {
    const role = getPublisherRole();
    return role === "publisher" || role === "auto";
  }

  async function fetchSnapshotMetaSafe() {
    try {
      const snapCfg = CONFIG && CONFIG.snapshot || {};
      const timeoutMs = Number(snapCfg.requestTimeoutMs || CONFIG?.nocodb?.requestTimeoutMs || 25000);
      const metaUrl = String(snapCfg.metaUrl || "").trim();
      if (!metaUrl) return null;
      return await fetchJsonWithTimeout(metaUrl, timeoutMs);
    } catch (_) {
      return null;
    }
  }

  function isSnapshotFreshByTtl(meta) {
    if (!isPublisherAutoRole()) return false;
    const pubCfg = getPublisherConfig();
    if (!pubCfg.schedule || pubCfg.schedule.enabled !== true) return false;
    const maxAgeMinutes = Math.max(1, Number(pubCfg.schedule.snapshotMaxAgeMinutes || 60));
    const builtAt = String(meta && meta.builtAt || "");
    const t = new Date(builtAt).getTime();
    if (!Number.isFinite(t)) return false;
    const ageMs = nowMs() - t;
    if (!Number.isFinite(ageMs) || ageMs < 0) return false;
    return ageMs <= maxAgeMinutes * 60 * 1000;
  }

  async function waitForSnapshotUpdateAfter(previousBuiltAtIso, timeoutMs, pollEveryMs) {
    const timeout = Math.max(15000, Number(timeoutMs || 180000));
    const step = Math.max(3000, Number(pollEveryMs || 7000));
    const start = nowMs();
    const prev = String(previousBuiltAtIso || "").trim();
    while (nowMs() - start < timeout) {
      const meta = await fetchSnapshotMetaSafe();
      const builtAt = String(meta && meta.builtAt || "").trim();
      if (builtAt && builtAt !== prev) return meta;
      await new Promise((resolve) => setTimeout(resolve, step));
    }
    return null;
  }

  function isLockBusyErrorText(messageText) {
    const s = String(messageText || "").toUpperCase();
    if (!s) return false;
    return s.includes("LOCKED") || s.includes("HTTP 409");
  }

  async function tryAcquirePublisherOrWait(snapshotMetaBefore) {
    try {
      const lock = await tryAcquirePublisherLock();
      return { lock, waitedOk: false, errorText: "", serveStaleNow: false };
    } catch (e) {
      const msg = String(e && e.message || e || "");
      if (!isPublisherRole() && isLockBusyErrorText(msg)) {
        return { lock: null, waitedOk: true, errorText: "", serveStaleNow: true };
      }
      return { lock: null, waitedOk: false, errorText: msg, serveStaleNow: false };
    }
  }

  async function hasFreshSnapshotByTtlForAutoWithMeta() {
    if (!isPublisherAutoRole()) return { isFresh: false, meta: null };
    const meta = await fetchSnapshotMetaSafe();
    return { isFresh: isSnapshotFreshByTtl(meta), meta };
  }

  function getPublisherOwnerId() {
    const cfg = getPublisherConfig();
    const fixed = String(cfg.ownerId || "").trim();
    if (fixed) return fixed;
    const machineName =
      (global.navigator && (global.navigator.userAgentData && global.navigator.userAgentData.platform)) ||
      (global.navigator && global.navigator.platform) ||
      "web";
    return "publisher-" + String(machineName).replace(/[^a-z0-9_-]+/gi, "_").toLowerCase();
  }

  async function postJsonWithTimeout(url, payload, timeoutMs) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), Math.max(3000, Number(timeoutMs || 25000)));
    try {
      const resp = await fetch(url, {
        method: "POST",
        signal: controller.signal,
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload || {}),
      });
      const text = await resp.text();
      let data = {};
      try {
        data = text ? JSON.parse(text) : {};
      } catch (_) {
        data = { raw: text };
      }
      if (!resp.ok) {
        const detail = data && (data.detail || data.error || data.raw) ? String(data.detail || data.error || data.raw) : "";
        throw new Error("HTTP " + resp.status + " từ " + url + (detail ? " | " + detail : ""));
      }
      return data;
    } finally {
      clearTimeout(timer);
    }
  }

  async function acquirePublisherLockSessionInternal() {
    const pubCfg = getPublisherConfig();
    const lockCfg = pubCfg.lock || {};
    const acquireUrl = String(lockCfg.acquireUrl || "").trim();
    const heartbeatUrl = String(lockCfg.heartbeatUrl || "").trim();
    if (!acquireUrl || !heartbeatUrl) {
      throw new Error("Thiếu cấu hình publisher.lock.acquireUrl hoặc heartbeatUrl.");
    }
    const ownerId = getPublisherOwnerId();
    const timeoutMs = Number(CONFIG?.snapshot?.requestTimeoutMs || CONFIG?.nocodb?.requestTimeoutMs || 25000);
    const acquire = await postJsonWithTimeout(acquireUrl, { ownerId }, timeoutMs);
    const realOwner = String(acquire && acquire.ownerId || ownerId).trim() || ownerId;
    const beatMs = Math.max(10000, Number(lockCfg.heartbeatIntervalMs || 60000));
    const timerId = setInterval(() => {
      void postJsonWithTimeout(heartbeatUrl, { ownerId: realOwner }, timeoutMs).catch((e) => {
        console.warn("Publisher heartbeat lỗi:", e);
      });
    }, beatMs);
    return { ownerId: realOwner, timerId };
  }

  async function tryAcquirePublisherLock() {
    if (!shouldAttemptPublisherFlow()) return null;
    return acquirePublisherLockSessionInternal();
  }

  /** Cho «Quét lại từ đầu»: consumer cũng có thể publish snapshot nếu worker chấp nhận lock/publish. */
  async function tryAcquirePublisherLockForForceScanPublish() {
    const snapCfg = CONFIG.snapshot || {};
    if (snapCfg.publishAfterForceScan === false) return null;
    const pubCfg = getPublisherConfig();
    if (!pubCfg || pubCfg.publish === false || (pubCfg.publish && pubCfg.publish.enabled === false)) {
      return null;
    }
    return acquirePublisherLockSessionInternal();
  }

  async function releasePublisherLock(lockSession) {
    if (!lockSession) return;
    try {
      if (lockSession.timerId) clearInterval(lockSession.timerId);
    } catch (_) {}
    const pubCfg = getPublisherConfig();
    const lockCfg = pubCfg.lock || {};
    const releaseUrl = String(lockCfg.releaseUrl || "").trim();
    if (!releaseUrl) return;
    const timeoutMs = Number(CONFIG?.snapshot?.requestTimeoutMs || CONFIG?.nocodb?.requestTimeoutMs || 25000);
    try {
      await postJsonWithTimeout(releaseUrl, { ownerId: lockSession.ownerId }, timeoutMs);
    } catch (e) {
      console.warn("Publisher release lock lỗi:", e);
    }
  }

  async function publishSnapshotRows(rows, builtAtIso, publisherStats) {
    if (!publisherLockSession || !publisherLockSession.ownerId) return { ok: false, skipped: true };
    const pubCfg = getPublisherConfig();
    if (!pubCfg.publish || pubCfg.publish.enabled === false) {
      return { ok: false, skipped: true };
    }
    const publishUrl = String(pubCfg.publish.url || "").trim();
    if (!publishUrl) throw new Error("Thiếu cấu hình publisher.publish.url.");
    if (!publisherLockSession || !publisherLockSession.ownerId) {
      throw new Error("Chưa acquire lock publisher.");
    }
    const timeoutMs = Number(CONFIG?.snapshot?.requestTimeoutMs || CONFIG?.nocodb?.requestTimeoutMs || 25000);
    const payload = {
      ownerId: publisherLockSession.ownerId,
      builtAt: builtAtIso || new Date().toISOString(),
      rows: Array.isArray(rows) ? rows : [],
      publisherStats: publisherStats && typeof publisherStats === "object" ? publisherStats : {},
    };
    const res = await postJsonWithTimeout(publishUrl, payload, timeoutMs);
    const builtAtOut = String((res && res.builtAt) || builtAtIso || "").trim();
    return {
      ok: !!(res && res.ok),
      rowsCount: Number(res && res.rowsCount || 0),
      hash: String(res && res.hash || ""),
      version: String(res && res.version || ""),
      builtAt: builtAtOut,
    };
  }

  async function fetchJsonWithTimeout(url, timeoutMs) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), Math.max(2000, Number(timeoutMs || 25000)));
    try {
      const resp = await fetch(url, { method: "GET", signal: controller.signal });
      const text = await resp.text();
      if (!resp.ok) {
        throw new Error("HTTP " + resp.status + " from " + url + (text ? "\n" + text.slice(0, 400) : ""));
      }
      try {
        return JSON.parse(text);
      } catch (_) {
        throw new Error("JSON parse lỗi từ " + url);
      }
    } finally {
      clearTimeout(timer);
    }
  }

  async function fetchSourceRecordsFromSnapshot() {
    const snapCfg = CONFIG && CONFIG.snapshot || {};
    const timeoutMs = Number(snapCfg.requestTimeoutMs || CONFIG?.nocodb?.requestTimeoutMs || 25000);
    const metaUrl = String(snapCfg.metaUrl || "").trim();
    const directDataUrl = String(snapCfg.dataUrl || "").trim();
    if (!metaUrl && !directDataUrl) {
      throw new Error("Thiếu cấu hình snapshot.metaUrl hoặc snapshot.dataUrl.");
    }

    let meta = {};
    let dataUrl = directDataUrl;
    if (metaUrl) {
      meta = await fetchJsonWithTimeout(metaUrl, timeoutMs);
      if (!dataUrl) dataUrl = String(meta && meta.dataUrl || "").trim();
    }
    if (!dataUrl) throw new Error("snapshot.meta.json không có dataUrl và cấu hình snapshot.dataUrl đang trống.");

    const data = await fetchJsonWithTimeout(dataUrl, timeoutMs);
    const rows = Array.isArray(data && data.rows) ? data.rows : [];
    const snapshotReady =
      !rows.length ||
      rows.every((r) =>
        !!r &&
        typeof r === "object" &&
        !!r.rowData &&
        typeof r.rowData === "object" &&
        (Object.prototype.hasOwnProperty.call(r, "fileUrl") || Object.prototype.hasOwnProperty.call(r, "taskCode"))
      );
    return {
      rows,
      source: "snapshot",
      snapshotReady,
      builtAt: String((data && data.builtAt) || (meta && meta.builtAt) || ""),
      version: String((meta && meta.version) || (data && data.snapshotVersion) || ""),
      hash: String((meta && meta.hash) || (data && data.hash) || ""),
      rowsCountMeta: Number((meta && meta.rowsCount) || 0),
      excelFilesScannedTotal: Number((meta && meta.excelFilesScannedTotal) || (data && data.excelFilesScannedTotal) || 0),
      excelFilesErrorTotal: Number((meta && meta.excelFilesErrorTotal) || (data && data.excelFilesErrorTotal) || 0),
      excelFilesOkWithTestBen: Number((meta && meta.excelFilesOkWithTestBen) || (data && data.excelFilesOkWithTestBen) || 0),
      nocoSourceRecordsTotal: Number((meta && meta.nocoSourceRecordsTotal) || (data && data.nocoSourceRecordsTotal) || 0),
      apiPath: dataUrl,
      authMode: "snapshot-public",
      fallbackReason: "",
    };
  }

  async function fetchSourceRecordsByMode() {
    const mode = getDataSourceMode();
    if (mode === "snapshot") {
      const snap = await fetchSourceRecordsFromSnapshot();
      if (!snap.snapshotReady) {
        throw new Error("Snapshot chưa ở định dạng kết quả cuối (thiếu rowData/fileUrl/taskCode).");
      }
      return snap;
    }

    if (mode === "hybrid") {
      try {
        const snap = await fetchSourceRecordsFromSnapshot();
        if (!snap.snapshotReady) {
          throw new Error("Snapshot chưa ở định dạng kết quả cuối (thiếu rowData/fileUrl/taskCode).");
        }
        if (!Array.isArray(snap.rows) || snap.rows.length === 0) {
          const localCache = cacheService.loadCache(CACHE_CONFIG);
          const cacheRows = Array.isArray(localCache && localCache.results) ? localCache.results : [];
          if (cacheRows.length > 0) {
            return {
              ...snap,
              rows: cacheRows,
              source: "snapshot",
              servedFromLocalCache: true,
              fallbackReason: "Snapshot tạm thời rỗng, đang dùng cache local gần nhất.",
            };
          }
          // Consumer mới (ẩn danh, máy mới) không có cache local:
          // vẫn phục vụ snapshot hiện tại để tránh tự động quét lại toàn bộ từ đầu.
          return {
            ...snap,
            rows: [],
            source: "snapshot",
            servedFromLocalCache: false,
            servedEmptySnapshot: true,
            fallbackReason: "Snapshot tạm thời rỗng, hệ thống đang cập nhật nền.",
          };
        }
        return snap;
      } catch (snapshotErr) {
        const fallback = await dataService.fetchAllRecords(CONFIG);
        return {
          rows: fallback.rows || [],
          source: "scan-fallback",
          snapshotReady: false,
          builtAt: "",
          version: "",
          hash: "",
          rowsCountMeta: 0,
          apiPath: fallback.apiPath,
          authMode: fallback.authMode,
          fallbackReason: String(snapshotErr && snapshotErr.message || snapshotErr || ""),
        };
      }
    }

    const fetchResult = await dataService.fetchAllRecords(CONFIG);
    return {
      rows: fetchResult.rows || [],
      source: "scan",
      builtAt: "",
      version: "",
      hash: "",
      rowsCountMeta: 0,
      apiPath: fetchResult.apiPath,
      authMode: fetchResult.authMode,
      fallbackReason: "",
    };
  }

  async function runRefresh(options = {}) {
    const silent = !!options.silent;
    const demoMode = !!options.demoMode;
    const forceExcelScanFromZero = !!options.forceExcelScanFromZero && !demoMode;
    const refreshStartedAtTs = nowMs();
    const refreshModeText = demoMode
      ? "DEMO"
      : forceExcelScanFromZero
        ? "Quét lại từ đầu"
        : (silent ? "Tự động" : "Thủ công");
    if (isRefreshing) return;
    isRefreshing = true;
    isRefreshBlockingInputs = !silent;
    if (forceExcelScanFromZero) {
      cacheService.clearFileCache(CACHE_CONFIG);
    }
    setSourceBadge("loading", getSnapshotFallbackCount());
    setTtlBadgeText("TTL snapshot: đang tính...", "");
    updateFrameState();
    applyManualLock();
    ui.refreshBtn.disabled = true;
    if (ui.refreshForceBtn) ui.refreshForceBtn.disabled = true;
    if (refreshLockModalEls && refreshLockModalEls.btn) {
      refreshLockModalEls.btn.disabled = true;
    }
    closeRefreshLockModal();
    const startedAt = performance.now();
    let publisherAcquireErrorText = "";
    let snapshotWasFreshByTtl = false;
    try {
      forceExcelScanFromZeroThisRefresh = !!forceExcelScanFromZero;
      forceScanThisRefresh = false;
      serveStaleSnapshotNotice = false;
      const ttlCheck = await hasFreshSnapshotByTtlForAutoWithMeta();
      const snapshotFreshByTtl = !!(ttlCheck && ttlCheck.isFresh);
      snapshotWasFreshByTtl = !!snapshotFreshByTtl;
      if (snapshotFreshByTtl) {
        forceScanThisRefresh = false;
        publisherLockSession = null;
      } else if (shouldAttemptPublisherFlow()) {
        const acquireResult = await tryAcquirePublisherOrWait(ttlCheck && ttlCheck.meta);
        publisherLockSession = acquireResult && acquireResult.lock ? acquireResult.lock : null;
        forceScanThisRefresh = !!(publisherLockSession && publisherLockSession.ownerId);
        serveStaleSnapshotNotice = !!(acquireResult && acquireResult.serveStaleNow);
        publisherAcquireErrorText = String(acquireResult && acquireResult.errorText || "");
        if (forceScanThisRefresh) {
          renderService.setStatus(ui, "Đang chạy chế độ publisher (đã acquire lock)...", "");
        } else if (serveStaleSnapshotNotice) {
          renderService.setStatus(
            ui,
            "Dữ liệu đang cập nhật nền, bạn vẫn dùng được snapshot gần nhất.",
            "warn"
          );
        } else if (publisherAcquireErrorText) {
          if (isPublisherRole()) {
            throw new Error(publisherAcquireErrorText);
          }
          console.warn("Auto publisher: không lấy được lock, chuyển sang consumer.", publisherAcquireErrorText);
        }
      }
      if (!silent) renderService.setStatus(ui, demoMode ? "Đang lấy dữ liệu demo..." : "Đang lấy dữ liệu...", "");
      ui.metaBox.textContent = "";

      const needTargetSnapshot =
        !!syncService &&
        typeof syncService.listAllTargetRecords === "function" &&
        (!!CONFIG.syncTarget?.enabled ||
          (manualService && typeof manualService.buildManualMapFromRecords === "function"));

      let sourceResult;
      let targetSnapshot = null;
      let sessionRowKeyMap = null;
      let manualMapPrepared = null;

      if (needTargetSnapshot) {
        const pair = await Promise.all([
          fetchSourceRecordsByMode(),
          syncService.listAllTargetRecords(CONFIG).catch((err) => {
            console.warn("Snapshot bảng đích (manual/sync) lỗi — dùng fallback từng bước:", err);
            return null;
          }),
        ]);
        sourceResult = pair[0];
        targetSnapshot = pair[1];
      } else {
        sourceResult = await fetchSourceRecordsByMode();
      }

      const records = Array.isArray(sourceResult && sourceResult.rows) ? sourceResult.rows : [];
      let fallbackCount = getSnapshotFallbackCount();
      if (sourceResult && sourceResult.source === "scan-fallback") {
        fallbackCount = increaseSnapshotFallbackCount();
      }
      setSourceBadge(sourceResult && sourceResult.source, fallbackCount);
      setTtlBadgeFromBuiltAt(sourceResult && sourceResult.builtAt);

      if (
        targetSnapshot !== null &&
        Array.isArray(targetSnapshot) &&
        CONFIG.syncTarget?.enabled &&
        syncService &&
        typeof syncService.buildRowKeyMapFromList === "function"
      ) {
        sessionRowKeyMap = syncService.buildRowKeyMapFromList(CONFIG, targetSnapshot);
      }
      if (
        targetSnapshot !== null &&
        Array.isArray(targetSnapshot) &&
        manualService &&
        typeof manualService.buildManualMapFromRecords === "function"
      ) {
        manualMapPrepared = manualService.buildManualMapFromRecords(CONFIG, targetSnapshot);
      }

      // Snapshot đúng định dạng kết quả cuối => bỏ toàn bộ scan Excel để tăng tốc.
      if (sourceResult && sourceResult.source === "snapshot" && sourceResult.snapshotReady) {
        let snapshotRows = records;
        let snapshotReconcileNote = "";
        const snapCfg = CONFIG.snapshot || {};
        if (snapCfg.reconcileWithNocoSource !== false) {
          try {
            const srcFetch = await dataService.fetchAllRecords(CONFIG);
            const sourceRecs = Array.isArray(srcFetch && srcFetch.rows) ? srcFetch.rows : [];
            const allowed = buildAllowedExcelUrlKeysByTaskCode(CONFIG, sourceRecs);
            const filtered = filterSnapshotRowsByCurrentSource(CONFIG, snapshotRows, allowed);
            if (filtered.length !== snapshotRows.length) {
              snapshotReconcileNote =
                "Lọc snapshot theo link Excel hiện tại (NocoDB nguồn): " +
                snapshotRows.length +
                " → " +
                filtered.length +
                " dòng.";
              console.info(snapshotReconcileNote);
            }
            snapshotRows = filtered;
          } catch (reconcileErr) {
            console.warn("Không lọc snapshot theo NocoDB nguồn — giữ nguyên snapshot.", reconcileErr);
          }
        }

        const results = snapshotRows.map((r) => (r && typeof r === "object" ? Object.assign({}, r) : r));
        let manualLoaded = 0;
        try {
          if (manualService && typeof manualService.fetchManualMap === "function") {
            let manualMap = manualMapPrepared;
            if (manualMap == null) {
              try {
                manualMap = await manualService.fetchManualMap(CONFIG);
              } catch (retryErr) {
                console.warn("Load manual fields failed (retry)", retryErr);
                manualMap = null;
              }
            }
            if (manualMap && typeof manualMap === "object") {
              for (const row of results) {
                const rowKey =
                  row && row.rowKey
                    ? String(row.rowKey)
                    : (typeof syncService.buildRowKey === "function"
                      ? syncService.buildRowKey(row)
                      : [row.fileUrl || "", row.sheetName || "", row.excelRowIndex || ""].join("__"));
                row.rowKey = rowKey;
                const saved = manualMap[rowKey];
                if (!saved) continue;
                manualLoaded += 1;
                row.targetRecordId = saved.recordId;
                row.manual_test_start_date = saved.fields.manual_test_start_date || "";
                row.manual_eta_date = saved.fields.manual_eta_date || "";
                row.manual_so_luong_mau = saved.fields.manual_so_luong_mau || "";
                row.manual_test_area = saved.fields.manual_test_area || "";
                row.manual_test_area_detail = saved.fields.manual_test_area_detail || "";
                row.manual_jig_code = saved.fields.manual_jig_code || "";
                row.manual_actual_done_date = saved.fields.manual_actual_done_date || "";
                row.manual_status = saved.fields.manual_status || "Pending";
                row.manual_ket_qua = saved.fields.manual_ket_qua || "";
                row.manual_ghi_chu = saved.fields.manual_ghi_chu || "";
                row.manual_updated_at = saved.fields.manual_updated_at || "";
                row.manual_updated_by = saved.fields.manual_updated_by || "";
                row.manual_version = Number(saved.fields.manual_version || 0);
                manualRecordIdByRowKey[rowKey] = saved.recordId;
                manualVersionByRowKey[rowKey] = row.manual_version;
              }
            }
          }
        } catch (manualErr) {
          console.warn("Load manual fields failed", manualErr);
        }

        await renderResults(results);
        const refreshEndedAtTs = nowMs();
        const elapsed = Math.round(performance.now() - startedAt);
        const excelFileStats = buildExcelFileStatsMetaParts(
          results,
          Number(sourceResult && sourceResult.excelFilesScannedTotal || 0),
          Number(sourceResult && sourceResult.excelFilesErrorTotal || 0),
          computeExcelFilesOkWithTestBen(results)
        );
        const sourceNocoTotal = Math.max(0, Number(sourceResult && sourceResult.nocoSourceRecordsTotal || 0));
        const metaParts = [
          latestSourceBadgeText,
          latestTtlBadgeText,
          "Số file từ nguồn NocoDB gốc: " + sourceNocoTotal,
          excelFileStats[0],
          excelFileStats[1],
          excelFileStats[2],
          "Loại refresh: " + refreshModeText,
          "Nguồn dữ liệu: Snapshot",
          (sourceResult && sourceResult.builtAt) ? ("Snapshot builtAt: " + sourceResult.builtAt) : "",
          (sourceResult && sourceResult.version) ? ("Snapshot version: " + sourceResult.version) : "",
          (sourceResult && sourceResult.hash) ? ("Snapshot hash: " + sourceResult.hash) : "",
          (sourceResult && Number(sourceResult.rowsCountMeta || 0) > 0)
            ? ("Snapshot rows(meta): " + Number(sourceResult.rowsCountMeta || 0))
            : "",
          "Bản ghi Snapshot: " + results.length,
          "Scan Excel: bỏ qua (đã dùng snapshot kết quả cuối)",
          "Thời gian: " + elapsed + " ms",
          "Nguồn endpoint: " + String(sourceResult && sourceResult.apiPath || ""),
          "Nguồn auth mode: " + String(sourceResult && sourceResult.authMode || ""),
          "Sheets cấu hình: " + CONFIG.sheetNames.join(", "),
          "Dòng có dữ liệu manual: " + manualLoaded,
          snapshotReconcileNote || "",
          (sourceResult && sourceResult.fallbackReason) ? ("Ghi chú: " + String(sourceResult.fallbackReason)) : "",
          serveStaleSnapshotNotice
            ? "Dữ liệu đang cập nhật nền, bạn vẫn dùng được snapshot gần nhất."
            : "",
          (!snapshotWasFreshByTtl && publisherAcquireErrorText)
            ? ("Không làm mới snapshot: " + publisherAcquireErrorText)
            : "",
          (isPublisherAutoRole() && getPublisherConfig().schedule && getPublisherConfig().schedule.enabled === true)
            ? ("Chính sách TTL: Snapshot mới hơn " + Math.max(1, Number((getPublisherConfig().schedule || {}).snapshotMaxAgeMinutes || 60)) + " phút thì bỏ qua scan/publish")
            : "",
        ].filter(Boolean);
        renderMetaSummary(metaParts);
        if (sourceResult && sourceResult.servedEmptySnapshot) {
          renderService.setStatus(
            ui,
            "Dữ liệu đang cập nhật nền, bạn vẫn dùng được snapshot gần nhất.",
            "warn"
          );
        } else if (serveStaleSnapshotNotice) {
          renderService.setStatus(
            ui,
            "Dữ liệu đang cập nhật nền, bạn vẫn dùng được snapshot gần nhất.",
            "warn"
          );
        } else if (!snapshotWasFreshByTtl && publisherAcquireErrorText) {
          renderService.setStatus(
            ui,
            "TTL snapshot đã hết hạn nhưng chưa lấy được lock publisher. Đang dùng snapshot cũ.\n" + publisherAcquireErrorText,
            "warn"
          );
        } else {
          renderService.setStatus(ui, "Hoàn tất quét dữ liệu", "ok");
        }
        if (!demoMode) {
          hasRefreshedThisSession = true;
          applyManualLock();
        }
        cacheService.saveCache(CACHE_CONFIG, {
          savedAt: nowMs(),
          results,
          metaText: metaParts.join(" | "),
          statusText: ui.statusBox.textContent,
          statusType: "ok",
        });
        return;
      }

      const files = [];
      const seenFileUrls = new Set();
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
            if (seenFileUrls.has(url)) continue;
            seenFileUrls.add(url);
            files.push({ url, taskMeta });
            // Demo mode: dung som ngay khi du file de scan nhanh.
            if (demoMode && files.length >= demoLimit) break;
          }
          if (demoMode && files.length >= demoLimit) break;
        }
      }

      if (!files.length) {
        await renderResults([]);
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
          return [{
            url: CONFIG.demo.sampleFileUrl,
            taskMeta: { taskCode: "DEMO", taskName: "Demo scan 1 file", assignee: "", completionActual: "" }
          }];
        }
        // Prefer a file that previously had matched rows (if cache exists),
        // so demo is more likely to show visible data.
        for (const file of files) {
          const c = cacheService.getFreshFileCache(CACHE_CONFIG, CONFIG, file.url);
          if (c && c.stats && Number(c.stats.matchedRows || 0) > 0) {
            return [file];
          }
        }
        return files.slice(0, demoLimit);
      })() : files;

      const syncOpts = sessionRowKeyMap ? { rowKeyMap: sessionRowKeyMap } : undefined;

      const results = [];
      const scanStats = { fileErrors: 0, sheetsMatched: 0, sheetsMissingDanhGia: 0, matchedRows: 0 };
      const cachedFiles = [];
      const needScanFiles = [];
      for (const file of filesToScan) {
        // Demo mode should always rescan to avoid stale cache confusion.
        if (demoMode) {
          needScanFiles.push(file);
          continue;
        }
        const cached = cacheService.getFreshFileCache(CACHE_CONFIG, CONFIG, file.url);
        if (cached) {
          cachedFiles.push({ file, cached });
        } else {
          needScanFiles.push(file);
        }
      }

      for (const item of cachedFiles) {
        results.push(...(item.cached.results || []));
        addScanStats(scanStats, item.cached.stats || {});
      }

      let doneCount = cachedFiles.length;
      renderService.setStatus(
        ui,
        (demoMode ? "Đang quét DEMO... " : "Đang quét Excel... ")
          + doneCount + "/" + filesToScan.length
          + (cachedFiles.length ? " | cache: " + cachedFiles.length : ""),
        ""
      );

      const updateProgress = () => {
        doneCount += 1;
        renderService.setStatus(
          ui,
          (demoMode ? "Đang quét DEMO... " : "Đang quét Excel... ")
            + doneCount + "/" + filesToScan.length
            + (cachedFiles.length ? " | cache: " + cachedFiles.length : ""),
          ""
        );
      };

      const concurrency = Math.max(1, Number(CONFIG.scan?.concurrency || 4));
      const deferSyncUntilEnd = CONFIG.scan && CONFIG.scan.deferSyncUntilEnd !== false;
      let cursor = 0;
      const syncStats = { created: 0, updated: 0, createErrors: 0, updateErrors: 0 };
      const syncBuffer = [];
      /** rowKey đã gửi sync thành công trong phiên — để reconcile không PATCH lặp dòng đã đồng bộ. */
      const syncedRowKeys = new Set();
      let syncLock = Promise.resolve();
      let syncFatalError = "";
      const syncBatchRows = Math.max(
        80,
        Number(CONFIG.scan?.syncFlushBatchRows) || 220
      );

      const flushSyncBuffer = async (force) => {
        if (!syncService || typeof syncService.syncResultsToTarget !== "function") return;
        if (syncFatalError) return;
        if (!force && syncBuffer.length < syncBatchRows) return;
        if (!syncBuffer.length) return;

        const payload = syncBuffer.splice(0, syncBuffer.length);
        syncLock = syncLock.then(async () => {
          renderService.setStatus(
            ui,
            (demoMode ? "Đang quét DEMO... " : "Đang quét Excel... ")
              + doneCount + "/" + filesToScan.length
              + " | Đang đồng bộ: " + payload.length + " dòng",
            ""
          );
          const res = await syncService.syncResultsToTarget(CONFIG, payload, syncOpts);
          if (res && !res.skipped) {
            addSyncStats(syncStats, res);
            if (typeof syncService.buildRowKey === "function") {
              for (let pi = 0; pi < payload.length; pi += 1) {
                syncedRowKeys.add(syncService.buildRowKey(payload[pi]));
              }
            }
          }
        }).catch((e) => {
          syncFatalError = String(e?.message || e || "Sync unknown error");
        });

        await syncLock;
      };

      const workersCount = Math.min(concurrency, needScanFiles.length);
      const workers = Array.from({ length: workersCount }, async () => {
        while (true) {
          const idx = cursor;
          cursor += 1;
          if (idx >= needScanFiles.length) break;
          const file = needScanFiles[idx];
          try {
            const fileResults = [];
            const fileStats = { fileErrors: 0, sheetsMatched: 0, sheetsMissingDanhGia: 0, matchedRows: 0 };
            await excelService.scanWorkbook(CONFIG, dataService, file.url, file.taskMeta, fileResults, () => {}, fileStats);
            results.push(...fileResults);
            addScanStats(scanStats, fileStats);
            if (fileResults.length) {
              syncBuffer.push(...fileResults);
              if (!deferSyncUntilEnd) {
                await flushSyncBuffer(false);
              }
            }
            cacheService.saveFileCache(CACHE_CONFIG, CONFIG, file.url, {
              results: fileResults,
              stats: fileStats,
            });
            updateProgress();
          } catch (error) {
            // scanWorkbook that throws sẽ khong goi progress() nen phai update o day
            updateProgress();
            scanStats.fileErrors += 1;
            console.warn("Scan workbook failed", file.url, error);
          }
        }
      });

      await Promise.all(workers);
      await flushSyncBuffer(true);

      // Sync to target NocoDB table (best-effort) — chỉ các dòng chưa sync (vd: từ cache file),
      // tránh PATCH lặp toàn bộ bảng khi results.length > syncedTotal do đếm lô.
      let syncMeta = "";
      try {
        if (syncFatalError) {
          syncMeta = "Sync lỗi: " + syncFatalError;
        } else if (syncService && typeof syncService.syncResultsToTarget === "function") {
          const seenReconcile = new Set();
          const reconcileRows = [];
          for (const r of results) {
            const rk = typeof syncService.buildRowKey === "function"
              ? syncService.buildRowKey(r)
              : [r.fileUrl || "", r.sheetName || "", r.excelRowIndex || ""].join("__");
            if (seenReconcile.has(rk)) continue;
            seenReconcile.add(rk);
            if (syncedRowKeys.has(rk)) continue;
            reconcileRows.push(r);
          }
          if (reconcileRows.length) {
            const reconcileRes = await syncService.syncResultsToTarget(CONFIG, reconcileRows, syncOpts);
            if (reconcileRes && !reconcileRes.skipped) addSyncStats(syncStats, reconcileRes);
          }

          syncMeta =
            "Sync tạo mới: " + syncStats.created +
            " | Sync cập nhật: " + syncStats.updated +
            " | Lỗi tạo: " + syncStats.createErrors +
            " | Lỗi cập nhật: " + syncStats.updateErrors;
        }
      } catch (e) {
        syncMeta = "Sync lỗi: " + String(e?.message || e || "");
      }

      // Always attach rowKey first so row selection/highlight works
      // even if manual map loading fails.
      for (const row of results) {
        row.rowKey = typeof syncService.buildRowKey === "function"
          ? syncService.buildRowKey(row)
          : [row.fileUrl || "", row.sheetName || "", row.excelRowIndex || ""].join("__");
      }

      let manualLoaded = 0;
      try {
        if (manualService && typeof manualService.fetchManualMap === "function") {
          let manualMap = manualMapPrepared;
          if (manualMap == null) {
            try {
              manualMap = await manualService.fetchManualMap(CONFIG);
            } catch (retryErr) {
              console.warn("Load manual fields failed (retry)", retryErr);
              manualMap = null;
            }
          }
          if (manualMap && typeof manualMap === "object") {
            for (const row of results) {
              const rowKey = row.rowKey;
              const saved = manualMap[rowKey];
              if (!saved) continue;
              manualLoaded += 1;
              row.targetRecordId = saved.recordId;
              row.manual_test_start_date = saved.fields.manual_test_start_date || "";
              row.manual_eta_date = saved.fields.manual_eta_date || "";
              row.manual_so_luong_mau = saved.fields.manual_so_luong_mau || "";
              row.manual_test_area = saved.fields.manual_test_area || "";
              row.manual_test_area_detail = saved.fields.manual_test_area_detail || "";
              row.manual_jig_code = saved.fields.manual_jig_code || "";
              row.manual_actual_done_date = saved.fields.manual_actual_done_date || "";
              row.manual_status = saved.fields.manual_status || "Pending";
              row.manual_ket_qua = saved.fields.manual_ket_qua || "";
              row.manual_ghi_chu = saved.fields.manual_ghi_chu || "";
              row.manual_updated_at = saved.fields.manual_updated_at || "";
              row.manual_updated_by = saved.fields.manual_updated_by || "";
              row.manual_version = Number(saved.fields.manual_version || 0);
              manualRecordIdByRowKey[rowKey] = saved.recordId;
              manualVersionByRowKey[rowKey] = row.manual_version;
            }
          }
        }
      } catch (manualErr) {
        console.warn("Load manual fields failed", manualErr);
      }

      await renderResults(results);
      const refreshEndedAtTs = nowMs();
      const elapsed = Math.round(performance.now() - startedAt);
      let publishMeta = "";
      if (!demoMode) {
        const wantForcePublishShare =
          !!forceExcelScanFromZero &&
          CONFIG.snapshot &&
          CONFIG.snapshot.publishAfterForceScan !== false &&
          !(getPublisherConfig().publish && getPublisherConfig().publish.enabled === false);

        if (wantForcePublishShare && !publisherLockSession) {
          try {
            const pubSession = await tryAcquirePublisherLockForForceScanPublish();
            if (pubSession && pubSession.ownerId) publisherLockSession = pubSession;
          } catch (forcePubErr) {
            const msg = String(forcePubErr && forcePubErr.message || forcePubErr || "");
            console.warn("Quét lại từ đầu: không publish snapshot cho mọi người dùng:", forcePubErr);
            publishMeta = "Publish snapshot (chia sẻ): thất bại — " + msg.slice(0, 220);
          }
        }

        if (publisherLockSession && publisherLockSession.ownerId) {
          const builtIso = new Date().toISOString();
          const publishStats = {
            nocoSourceRecordsTotal: records.length,
            excelFilesScannedTotal: filesToScan.length,
            excelFilesErrorTotal: scanStats.fileErrors,
            excelFilesOkWithTestBen: computeExcelFilesOkWithTestBen(results),
          };
          const publishRes = await publishSnapshotRows(results, builtIso, publishStats);
          if (publishRes && publishRes.ok) {
            const shareHint = forceExcelScanFromZero ? "Đã publish snapshot cho mọi người dùng | " : "";
            publishMeta =
              shareHint +
              "Publisher publish: OK | Rows: " +
              publishRes.rowsCount +
              (publishRes.builtAt ? (" | Snapshot builtAt (server): " + publishRes.builtAt) : "");
            if (publishRes.builtAt) setTtlBadgeFromBuiltAt(publishRes.builtAt);
          } else if (publishRes && publishRes.skipped) {
            if (!publishMeta) publishMeta = "Publisher publish: bỏ qua (cấu hình)";
          } else {
            publishMeta = publishMeta || "Publisher publish: không thành công";
          }
        }
      }

      const excelFileStats = buildExcelFileStatsMetaParts(
        results,
        filesToScan.length,
        scanStats.fileErrors
      );

      const metaParts = [
        latestSourceBadgeText,
        latestTtlBadgeText,
        "Số file từ nguồn NocoDB gốc: " + Math.max(0, Number(records.length || 0)),
        excelFileStats[0],
        excelFileStats[1],
        excelFileStats[2],
        "Loại refresh: " + refreshModeText,
        forceExcelScanFromZeroThisRefresh ? "Quét từ đầu: có (đã xóa cache Excel trình duyệt, bỏ snapshot)" : "",
        "Nguồn dữ liệu: " + (
          sourceResult && sourceResult.source === "snapshot" ? "Snapshot" :
          sourceResult && sourceResult.source === "scan-fallback" ? "Hybrid (fallback scan)" :
          "Scan trực tiếp"
        ),
        (sourceResult && sourceResult.builtAt) ? ("Snapshot builtAt: " + sourceResult.builtAt) : "",
        (sourceResult && sourceResult.version) ? ("Snapshot version: " + sourceResult.version) : "",
        (sourceResult && sourceResult.hash) ? ("Snapshot hash: " + sourceResult.hash) : "",
        (sourceResult && Number(sourceResult.rowsCountMeta || 0) > 0)
          ? ("Snapshot rows(meta): " + Number(sourceResult.rowsCountMeta || 0))
          : "",
        "Bản ghi NocoDB: " + records.length,
        "File Excel tổng: " + filesToScan.length + (demoMode ? " (demo)" : ""),
        "File lấy từ cache: " + cachedFiles.length,
        "File quét mới: " + needScanFiles.length,
        "Dòng Test bền tìm được: " + results.length,
        "File lỗi: " + scanStats.fileErrors,
        "Sheet có cột Đánh giá: " + scanStats.sheetsMatched,
        "Sheet thiếu cột Đánh giá: " + scanStats.sheetsMissingDanhGia,
        "Thời gian: " + elapsed + " ms",
        syncMeta,
        "Nguồn endpoint: " + String(sourceResult && sourceResult.apiPath || ""),
        "Nguồn auth mode: " + String(sourceResult && sourceResult.authMode || ""),
        "Sheets cấu hình: " + CONFIG.sheetNames.join(", "),
        "Dòng có dữ liệu manual: " + manualLoaded,
        publishMeta,
        serveStaleSnapshotNotice
          ? "Dữ liệu đang cập nhật nền, bạn vẫn dùng được snapshot gần nhất."
          : "",
        (!snapshotWasFreshByTtl && publisherAcquireErrorText)
          ? ("Không làm mới snapshot: " + publisherAcquireErrorText)
          : "",
        (isPublisherAutoRole() && getPublisherConfig().schedule && getPublisherConfig().schedule.enabled === true)
          ? ("Chính sách TTL: Snapshot mới hơn " + Math.max(1, Number((getPublisherConfig().schedule || {}).snapshotMaxAgeMinutes || 60)) + " phút thì bỏ qua scan/publish")
          : "",
      ].filter(Boolean);
      renderMetaSummary(metaParts);
      if (demoMode && results.length === 0) {
        renderService.setStatus(
          ui,
          "DEMO đã quét xong nhưng file được chọn không có dòng 'Đánh giá = Test bền'. Hãy bấm Refresh dữ liệu để quét toàn bộ.",
          "err"
        );
      } else {
        const baseDone = "Hoàn tất quét dữ liệu";
        if (serveStaleSnapshotNotice && sourceResult && sourceResult.source === "snapshot") {
          renderService.setStatus(
            ui,
            "Dữ liệu đang cập nhật nền, bạn vẫn dùng được snapshot gần nhất.",
            "warn"
          );
        } else if (sourceResult && sourceResult.source === "scan-fallback" && sourceResult.fallbackReason) {
          renderService.setStatus(
            ui,
            baseDone + " (fallback scan do snapshot lỗi).\n" + String(sourceResult.fallbackReason).slice(0, 280),
            "warn"
          );
        } else if (sourceResult && sourceResult.source === "snapshot") {
          renderService.setStatus(ui, baseDone + " (nguồn snapshot)", "ok");
        } else {
          renderService.setStatus(ui, baseDone, "ok");
        }
      }
      if (!demoMode) {
        hasRefreshedThisSession = true;
        applyManualLock();
      }
      if (sessionRowKeyMap && syncService && typeof syncService.persistSessionRowKeyMap === "function") {
        try {
          syncService.persistSessionRowKeyMap(sessionRowKeyMap);
        } catch (_) {}
      }
      cacheService.saveCache(CACHE_CONFIG, {
        savedAt: nowMs(),
        results,
        metaText: metaParts.join(" | "),
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
      await releasePublisherLock(publisherLockSession);
      publisherLockSession = null;
      forceScanThisRefresh = false;
      forceExcelScanFromZeroThisRefresh = false;
      serveStaleSnapshotNotice = false;
      isRefreshing = false;
      isRefreshBlockingInputs = false;
      updateFrameState();
      ui.refreshBtn.disabled = false;
      if (ui.refreshForceBtn) ui.refreshForceBtn.disabled = false;
      if (refreshLockModalEls && refreshLockModalEls.btn) {
        refreshLockModalEls.btn.disabled = false;
      }
      applyManualLock();
      if (!hasRefreshedThisSession) {
        openRefreshLockModal();
      }
    }
  }

  async function bootFromCacheThenRefresh() {
    const cache = cacheService.loadCache(CACHE_CONFIG);
    if (!cache) return;
    await renderResults(cache.results || []);
    const cachePartsBase = String(cache.metaText || "")
      .split(" | ")
      .map((x) => String(x || "").trim())
      .filter(Boolean);
    const cacheParts = cachePartsBase.filter((x) => !/^Refresh gần nhất:/i.test(x));
    const elapsedText = formatElapsedSince(cache.savedAt);
    if (elapsedText) cacheParts.push("Refresh gần nhất: " + elapsedText);
    renderMetaSummary(cacheParts, "Cache lúc: " + formatTime(cache.savedAt));
    renderService.setStatus(ui, "Tải lại trang để cập nhật dữ liệu mới nhất", "warn");
    // Khong auto full refresh luc mo trang de tranh gay nham lan voi mode DEMO.
  }

  ui.refreshBtn.addEventListener("click", () => runRefresh({ silent: false }));
  if (ui.refreshForceBtn) {
    ui.refreshForceBtn.addEventListener("click", () => {
      openForceScanConfirmModal(() => {
        void runRefresh({ silent: false, forceExcelScanFromZero: true }).catch((err) => console.error("runRefresh:", err));
      });
    });
  }
  initLabPreviewLink();
  if (ui.exportExcelBtn) {
    ui.exportExcelBtn.addEventListener("click", () => exportFilteredToExcel());
  }

  if (ui.filterField) {
    ui.filterField.addEventListener("change", () => {
      currentFilter.field = String(ui.filterField.value || "");
      // Doi truong thi reset gia tri loc cu de tranh lan ket qua.
      currentFilter.value = "";
      currentFilter.dateFrom = "";
      currentFilter.dateTo = "";
      updateFilterInputMode();
      void renderFiltered().catch((err) => console.error("renderFiltered:", err));
    });
  }
  if (ui.filterField2) {
    ui.filterField2.addEventListener("change", () => {
      currentFilter2.field = String(ui.filterField2.value || "");
      currentFilter2.value = "";
      currentFilter2.dateFrom = "";
      currentFilter2.dateTo = "";
      updateFilterInputMode();
      void renderFiltered().catch((err) => console.error("renderFiltered:", err));
    });
  }
  if (ui.filterValue) {
    ui.filterValue.addEventListener("input", () => {
      currentFilter.value = String(ui.filterValue.value || "");
      void renderFiltered().catch((err) => console.error("renderFiltered:", err));
    });
  }
  if (ui.filterValue2) {
    ui.filterValue2.addEventListener("input", () => {
      currentFilter2.value = String(ui.filterValue2.value || "");
      void renderFiltered().catch((err) => console.error("renderFiltered:", err));
    });
  }
  if (ui.filterDateFrom) {
    ui.filterDateFrom.addEventListener("change", () => {
      currentFilter.dateFrom = String(ui.filterDateFrom.value || "");
      void renderFiltered().catch((err) => console.error("renderFiltered:", err));
    });
  }
  if (ui.filterDateFrom2) {
    ui.filterDateFrom2.addEventListener("change", () => {
      currentFilter2.dateFrom = String(ui.filterDateFrom2.value || "");
      void renderFiltered().catch((err) => console.error("renderFiltered:", err));
    });
  }
  if (ui.filterDateTo) {
    ui.filterDateTo.addEventListener("change", () => {
      currentFilter.dateTo = String(ui.filterDateTo.value || "");
      void renderFiltered().catch((err) => console.error("renderFiltered:", err));
    });
  }
  if (ui.filterDateTo2) {
    ui.filterDateTo2.addEventListener("change", () => {
      currentFilter2.dateTo = String(ui.filterDateTo2.value || "");
      void renderFiltered().catch((err) => console.error("renderFiltered:", err));
    });
  }
  if (ui.filterClear) {
    ui.filterClear.addEventListener("click", () => {
      currentFilter.field = "";
      currentFilter.value = "";
      currentFilter.dateFrom = "";
      currentFilter.dateTo = "";
      currentFilter2.field = "";
      currentFilter2.value = "";
      currentFilter2.dateFrom = "";
      currentFilter2.dateTo = "";
      if (ui.filterField) ui.filterField.value = "";
      if (ui.filterValue) ui.filterValue.value = "";
      if (ui.filterDateFrom) ui.filterDateFrom.value = "";
      if (ui.filterDateTo) ui.filterDateTo.value = "";
      if (ui.filterField2) ui.filterField2.value = "";
      if (ui.filterValue2) ui.filterValue2.value = "";
      if (ui.filterDateFrom2) ui.filterDateFrom2.value = "";
      if (ui.filterDateTo2) ui.filterDateTo2.value = "";
      updateFilterInputMode();
      void renderFiltered().catch((err) => console.error("renderFiltered:", err));
    });
  }

  ui.tableRoot.addEventListener("click", (event) => {
    const tr = event.target && event.target.closest ? event.target.closest("tbody tr") : null;
    if (!tr) return;
    const newKey = String(tr.getAttribute("data-row-key") || "");
    if (newKey === selectedTableRowKey) return;
    selectedTableRowKey = newKey;
    const allRows = ui.tableRoot.querySelectorAll("tbody tr");
    allRows.forEach((r) => r.classList.remove("row-selected"));
    tr.classList.add("row-selected");
    ui.tableRoot.querySelectorAll(".manual-save-state").forEach((el) => {
      el.textContent = "";
    });
    setEditingStatusIfReady();
  });

  ui.tableRoot.addEventListener("click", (event) => {
    const openBtn = event.target && event.target.closest ? event.target.closest(".lab-map-open-btn") : null;
    if (!openBtn) return;
    event.preventDefault();
    if (openBtn.disabled) return;
    const cell = openBtn.closest(".manual-area-cell");
    const input = cell && cell.querySelector('input[name="manual_test_area"]');
    if (!input) return;
    openLabMapModal(input);
  });

  ui.tableRoot.addEventListener("change", (event) => {
    const el = event.target;
    if (!el || !el.classList || !el.classList.contains("manual-status-select")) return;
    el.classList.remove("status-pending", "status-doing", "status-done", "status-cancel");
    const value = String(el.value || "").toLowerCase();
    if (value) el.classList.add("status-" + value);
    setEditingStatusIfReady();
  });
  ui.tableRoot.addEventListener("input", (event) => {
    const el = event.target;
    if (!el || !el.classList || !el.classList.contains("manual-input")) return;
    setEditingStatusIfReady();
  });
  ui.tableRoot.addEventListener("click", async (event) => {
    const btn = event.target && event.target.closest ? event.target.closest(".manual-save-btn") : null;
    if (!btn) return;
    if (!hasRefreshedThisSession) {
      openRefreshLockModal();
      renderService.setStatus(ui, "Vui lòng bấm \"Refresh dữ liệu\" trước khi lưu.", "err");
      return;
    }
    if (!manualService || typeof manualService.saveManualFields !== "function") {
      renderService.setStatus(ui, "Thiếu manualService để lưu dữ liệu nhập tay.", "err");
      return;
    }

    const rowEl = btn.closest("tr");
    if (!rowEl) return;
    const rowKey = String(rowEl.getAttribute("data-row-key") || "").trim();
    const recordId = String(rowEl.getAttribute("data-record-id") || manualRecordIdByRowKey[rowKey] || "").trim();
    if (!rowKey || !recordId) {
      renderService.setStatus(ui, "Không tìm thấy record để lưu manual. Vui lòng Refresh dữ liệu.", "err");
      return;
    }

    const readValue = (name) => {
      const node = rowEl.querySelector('[name="' + name + '"]');
      if (!node) return "";
      const v = String(node.value ?? "");
      if (node.tagName === "TEXTAREA") return v;
      return v.trim();
    };
    const stateEl = rowEl.querySelector(".manual-save-state");
    const manualPayload = {
      manual_test_start_date: readValue("manual_test_start_date"),
      manual_eta_date: readValue("manual_eta_date"),
      manual_so_luong_mau: readValue("manual_so_luong_mau"),
      manual_test_area: readValue("manual_test_area"),
      manual_test_area_detail: readValue("manual_test_area_detail"),
      manual_jig_code: readValue("manual_jig_code"),
      manual_actual_done_date: readValue("manual_actual_done_date"),
      manual_status: readValue("manual_status") || "Pending",
      manual_ket_qua: readValue("manual_ket_qua"),
      manual_ghi_chu: readValue("manual_ghi_chu"),
      manual_updated_at: new Date().toISOString(),
      manual_updated_by: "web-user",
      manual_version: Number(manualVersionByRowKey[rowKey] || 0) + 1,
    };

    const inputs = rowEl.querySelectorAll(".manual-input");
    btn.disabled = true;
    inputs.forEach((el) => { el.disabled = true; });
    if (stateEl) stateEl.textContent = "Đang lưu...";
    try {
      await manualService.saveManualFields(CONFIG, recordId, manualPayload);
      manualVersionByRowKey[rowKey] = Number(manualPayload.manual_version || 0);
      const rowObj = latestResults.find((r) => String(r.rowKey || "") === rowKey);
      if (rowObj) {
        rowObj.manual_so_luong_mau = manualPayload.manual_so_luong_mau;
        rowObj.manual_ket_qua = manualPayload.manual_ket_qua;
        rowObj.manual_ghi_chu = manualPayload.manual_ghi_chu;
      }
      if (stateEl) stateEl.textContent = "Đã lưu";
      renderService.setStatus(ui, "Đã lưu thành công", "ok");
    } catch (error) {
      if (stateEl) stateEl.textContent = "Lỗi";
      renderService.setStatus(ui, "Lưu manual lỗi: " + String(error?.message || error || ""), "err");
    } finally {
      btn.disabled = false;
      inputs.forEach((el) => { el.disabled = false; });
    }
  });

  void (async function initAfterLoad() {
    setSourceBadge("loading", getSnapshotFallbackCount());
    setTtlBadgeText("TTL snapshot: đang tính...", "");
    await bootFromCacheThenRefresh();
    updateFrameState();
    applyManualLock();
    if (!hasRefreshedThisSession) {
      openRefreshLockModal();
    }
    if (Number(CACHE_CONFIG.autoRefreshMs || 0) > 0) {
      setInterval(() => {
        if (!document.hidden) runRefresh({ silent: true });
      }, CACHE_CONFIG.autoRefreshMs);
    }
  })().catch((err) => {
    console.error(err);
    updateFrameState();
    applyManualLock();
    if (!hasRefreshedThisSession) openRefreshLockModal();
  });
})(window);

