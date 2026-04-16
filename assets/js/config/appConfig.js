(function initForm15Config(global) {
  const CONFIG = {
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
      limit: 100,
      requestTimeoutMs: 25000,
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
  };

  const CACHE_CONFIG = {
    key: "form15.testben.cache.v1",
    ttlMs: 1000 * 60 * 30,
    autoRefreshMs: 1000 * 60 * 10,
  };

  global.Form15Config = { CONFIG, CACHE_CONFIG };
})(window);

