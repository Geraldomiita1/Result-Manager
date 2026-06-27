// Provides window.storage.{get,set} backed by localStorage so the MKIS
// component runs identically outside the Lovable host environment.
// Build: 2026.06.27.1536
export function installStorageShim() {
  if (typeof window === "undefined") return;
  if (window.storage && window.storage.__mkisShim) return;
  const prefix = "mkis_shared::";
  window.storage = {
    __mkisShim: true,
    async get(key) {
      try {
        const value = window.localStorage.getItem(prefix + key);
        return value == null ? null : { value };
      } catch { return null; }
    },
    async set(key, value) {
      try { window.localStorage.setItem(prefix + key, String(value)); } catch {}
      return true;
    },
    async remove(key) {
      try { window.localStorage.removeItem(prefix + key); } catch {}
      return true;
    },
  };
}
