(function initForm15CacheService(global) {
  const { nowMs } = global.Form15Utils;

  function saveCache(cacheConfig, payload) {
    try {
      localStorage.setItem(cacheConfig.key, JSON.stringify(payload));
    } catch (_) {}
  }

  function loadCache(cacheConfig) {
    try {
      const raw = localStorage.getItem(cacheConfig.key);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (!parsed || !Array.isArray(parsed.results)) return null;
      return parsed;
    } catch (_) {
      return null;
    }
  }

  function isCacheFresh(cacheConfig, cache) {
    if (!cache || !cache.savedAt) return false;
    return nowMs() - Number(cache.savedAt) <= cacheConfig.ttlMs;
  }

  global.Form15CacheService = {
    saveCache,
    loadCache,
    isCacheFresh,
  };
})(window);

