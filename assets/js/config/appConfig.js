(function initForm15Config(global) {
  const CONFIG = {
    /** scan | snapshot | hybrid (ưu tiên snapshot, lỗi thì fallback scan) */
    dataSourceMode: "hybrid",
    snapshot: {
      metaUrl: "https://form15-nocodb-proxy.superaguero1999.workers.dev/snapshot.meta.json",
      dataUrl: "https://form15-nocodb-proxy.superaguero1999.workers.dev/snapshot.json",
      requestTimeoutMs: 25000,
      /**
       * true: sau khi tải snapshot, lọc dòng theo link Excel đang có trên bản ghi NocoDB nguồn.
       * Tránh hiển thị dòng “Test bền” của file cũ khi tác vụ đã đổi Link BCexcel (snapshot server có thể chưa kịp rebuild).
       */
      reconcileWithNocoSource: true,
      /**
       * Sau «Quét lại từ đầu» + quét xong: thử acquire lock và publish snapshot (nếu API/worker cho phép).
       * Giúp mọi tab/người dùng tải snapshot mới; builtAt/hash phụ thuộc worker phản hồi publish.
       */
      publishAfterForceScan: true,
    },
    publisher: {
      /** auto | consumer | publisher */
      role: "consumer",
      ownerId: "",
      schedule: {
        enabled: true,
        snapshotMaxAgeMinutes: 60,
      },
      lock: {
        acquireUrl: "https://form15-nocodb-proxy.superaguero1999.workers.dev/snapshot/lock/acquire",
        heartbeatUrl: "https://form15-nocodb-proxy.superaguero1999.workers.dev/snapshot/lock/heartbeat",
        releaseUrl: "https://form15-nocodb-proxy.superaguero1999.workers.dev/snapshot/lock/release",
        heartbeatIntervalMs: 60000,
      },
      publish: {
        url: "https://form15-nocodb-proxy.superaguero1999.workers.dev/snapshot/publish",
        enabled: true,
      },
    },
    demo: {
      enabled: true,
      maxFiles: 1,
      sampleFileUrl: "",
    },
    nocodb: {
      proxyUrl: "https://form15-nocodb-proxy.superaguero1999.workers.dev",
      host: "https://iatzhxxuk.tino.page",
      apiPathCandidates: [
        "/api/v2/tables/mnwnhukgbu8zs9o/records",
        "/nc/api/v2/tables/mnwnhukgbu8zs9o/records",
      ],
      viewId: "vwnqck09ijylco3b",
      limit: 500,
      requestTimeoutMs: 40000,
      token: "G3QREWJrrT_E7QHPfteLJlI90zab7mUW_jNMwZIu",
      authModes: ["header", "query_xc_token", "query_token"],
    },
    excelFieldId: "Link BCexcel",
    excelFieldCandidates: ["Link BCexcel", "File attachment", "Link BCexcel "],
    sheetNames: ["TCKT", "THEM"],
    testBenKeyword: "test ben",
    danhGiaHeaderCandidates: ["Đánh giá", "Danh gia"],
    taskCodeCandidates: ["Mã tác vụ", "Ma tac vu", "Task code", "Task ID"],
    taskNameCandidates: ["Tên tác vụ", "Ten tac vu", "Task name", "Task"],
    assigneeCandidates: ["Assignee", "Asignee", "Người phụ trách", "Nguoi phu trach"],
    completionActualCandidates: [
      "Ngày trả báo cáo tức thời",
      "Ngay tra bao cao tuc thoi",
      "Ngày hoàn thành thực tế",
      "Ngay hoan thanh thuc te",
      "Ngày hoàn thành thực tê",
      "Ngay hoan thanh thuc te",
    ],
    ignoreColumnRange: { from: 16, to: 32 }, // 1-based
    // Chi lay cac cot trong file excel bao cao de hien thi.
    // Chu y: map theo header trong sheet (fuzzy theo normalizeCompact).
    excelDisplayColumns: [
      "STT",
      "Công chuẩn",
      "Mã danh mục",
      "Hạng mục kiểm tra (Index)",
      "Tiêu chuẩn (Standard)",
      "Công cụ (Tool)",
      "Hướng dẫn / Phương pháp (Document)",
    ],
    excelColumnHeaderCandidates: {
      "STT": ["STT", "So thu tu", "STT #"],
      "Công chuẩn": ["Công chuẩn", "Cong chuan", "Cong-Chuan", "Congchuan"],
      "Mã danh mục": ["Mã danh mục", "Ma danh muc", "Danh muc", "Danhmuc"],
      "Hạng mục kiểm tra (Index)": ["Hạng mục kiểm tra (Index)", "Hang muc kiem tra (Index)", "Hang muc kiem tra", "Index"],
      "Tiêu chuẩn (Standard)": ["Tiêu chuẩn (Standard)", "Tieu chuan (Standard)", "Tieu chuan", "Standard"],
      "Công cụ (Tool)": ["Công cụ (Tool)", "Cong cu (Tool)", "Cong cu", "Tool"],
      "Hướng dẫn / Phương pháp (Document)": [
        "Hướng dẫn / Phương pháp (Document)",
        "Huong dan / Phuong phap (Document)",
        "Huong dan / Phuong phap",
        "Document",
        "Hướng dẫn",
        "Phương pháp",
      ],
    },
    // Cot Excel (rowData) an tren bang; du lieu van quet / sync / loc noi bo neu can.
    tableHideExcelColumns: ["Công chuẩn", "Mã danh mục", "Công cụ (Tool)"],
    scan: {
      concurrency: 14, // so luong file excel quet cung luc
      /** So dong toi da truoc khi flush dong bo len NocoDB (chi dung khi deferSyncUntilEnd = false). */
      syncFlushBatchRows: 220,
      /**
       * true (mac dinh): quét xong hết file rồi mới đồng bộ 1–2 lần — giảm mạnh thời gian so với sync từng lô trong lúc quét.
       * false: đồng bộ dần (hành vi cũ), hữu ích nếu cần thấy dữ liệu trên NocoDB ngay khi chưa quét xong.
       */
      deferSyncUntilEnd: true,
    },
    // Sơ đồ: rect = % trên lab-map-map-area (sau contentInset). PNG có viền trắng → chỉnh contentInset.
    labMap: {
      enabled: true,
      imageUrl: "assets/lab-map/lab-map.PNG",
      /** Link dạng chữ trên topbar — mở popup xem ảnh sơ đồ phòng lab (khu vực test). */
      previewLinkLabel: "Xem sơ đồ khu vực test (phòng lab)",
      /** Tiêu đề trong popup (ảnh dùng previewImageUrl hoặc imageUrl). */
      previewModalTitle: "Khu vực test - phòng Lab_DQA 2026",
      /** Ảnh popup “Xem sơ đồ” trên topbar (khác imageUrl nếu cần). */
      previewImageUrl: "assets/lab-map/view-lab-map.PNG",
      // Tăng số này khi đổi value/label/hint trong zones: trình duyệt sẽ bỏ bản zones đã Lưu trong localStorage (nếu lệch phiên bản) và dùng zones trong file.
      zonesDataVersion: 2,
      // Nếu bản vẽ nằm lệch trong file ảnh: tăng left/top/right/bottom (%) để khớp ô xanh với nền
      contentInset: { left: 0, top: 0, right: 0, bottom: 0 },
      zones: [
        { value: "Điện.01", label: "Điện.01", hint: "Đ1 – Đ2 – Đ3 – QA.143", rect: { left: 17.723, top: 5.544, width: 25.972, height: 4.925 } },
        { value: "Điện.02", label: "Điện.02", hint: "QA.20 – QA.21 – QA.53 (cạnh ĐB2)", rect: { left: 23.725, top: 48.376, width: 22.028, height: 5.204 } },
        { value: "Hóa.01", label: "Hóa.01", hint: "HB1", rect: { left: 48.094, top: 25.724, width: 26.794, height: 7.191 } },
        { value: "Kho", label: "Kho", hint: "K2 – K5 – K7 (cột dọc)", rect: { left: 50.137, top: 36.343, width: 25.76, height: 15.866 } },
        { value: "Test rò rỉ", label: "Test rò rỉ", hint: "QA.119", rect: { left: 45.008, top: 53.226, width: 40.069, height: 7.48 } },
        { value: "Bền LK", label: "Bền LK", hint: "QA.120", rect: { left: 62.952, top: 74.727, width: 29.818, height: 7.676 } },
        { value: "Bền MLN", label: "Bền MLN", hint: "QA.121", rect: { left: 37.301, top: 74.565, width: 24.779, height: 7.838 } },
        { value: "Test lõi", label: "Test lõi", hint: "QA.138 – QA.137", rect: { left: 65.102, top: 88.811, width: 32.102, height: 8.673 } },
        { value: "Nhiệt.01", label: "Nhiệt.01", hint: "QA.122", rect: { left: 34.284, top: 53.683, width: 8.101, height: 7.456 } },
        { value: "MLN.01", label: "MLN.01", hint: "QA.142", rect: { left: 51.754, top: 63.315, width: 23.425, height: 7.793 } },
        { value: "Test nổ", label: "Test nổ", hint: "QA.118", rect: { left: 77.182, top: 63.315, width: 10.653, height: 7.631 } },
        { value: "Tủ UV", label: "Tủ UV", hint: "QA.144", rect: { left: 16.457, top: 77.907, width: 7.366, height: 5.469 } },
        { value: "Nhiệt.02", label: "Nhiệt.02", hint: "QA.51", rect: { left: 16.324, top: 84.539, width: 7.5, height: 4.498 } },
        { value: "NaCl", label: "NaCl", hint: "QA.18", rect: { left: 23.285, top: 90.85, width: 7.5, height: 4.498 } },
        { value: "MLN.02", label: "MLN.02", hint: "QA.141", rect: { left: 31.402, top: 89.814, width: 11.188, height: 6.705 } },
        { value: "MLN.03", label: "MLN.03", hint: "QA.140", rect: { left: 42.654, top: 89.814, width: 10.92, height: 7.029 } },
        { value: "MLN.04", label: "MLN.04", hint: "QA.139", rect: { left: 53.701, top: 89.976, width: 11.123, height: 6.867 } },
      ],
    },
    syncTarget: {
      enabled: true,
      // Anh xa key trong code -> ten cot tren NocoDB (key trong JSON khi doc/ghi). Cot hien thi "Kết quả" / "Ghi chú" thuong khac manual_ket_qua / manual_ghi_chu.
      manualFieldKeys: {
        manual_so_luong_mau: "Số lượng mẫu",
        manual_ket_qua: "Kết quả",
        manual_ghi_chu: "Ghi chú",
      },
      // Su dung Cloudflare worker de tranh CORS va bao mat token.
      proxyUrl: "https://form15-nocodb-proxy.superaguero1999.workers.dev",
      host: "https://iatzhxxuk.tino.page",
      tableId: "msc421bfqh1yjne",
      viewId: "vwu5d7inj77t4i1z",
      token: "7fi5obsEtrRFZzLuW9KIKwBOUCNnhtbW4jrV1oqh",
      batchSize: 100,
      updateConcurrency: 24,
      // Ten cot ky thuat trong table dich (ban can tao trong NocoDB).
      technicalColumns: {
        rowKey: "rowKey",
        excelRowIndex: "excelRowIndex",
        lastScannedAt: "lastScannedAt",
        syncSource: "syncSource",
      },
    },
  };

  const CACHE_CONFIG = {
    key: "form15.testben.cache.v1",
    fileKey: "form15.testben.file-cache.v1",
    ttlMs: 1000 * 60 * 30,
    fileTtlMs: 1000 * 60 * 60 * 6,
    autoRefreshMs: 1000 * 60 * 10,
  };

  global.Form15Config = { CONFIG, CACHE_CONFIG };
})(window);

