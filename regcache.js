// Registry cache, kept in IndexedDB.
// The statewide registry (~12,000 records, ~5 MB) is too big for localStorage, which is also
// shared with FieldPro and the LPP Log (same github.io origin) — filling it would break their saves.
// Every call fails soft: no cache just means the page waits for the network.
(function () {
  const DB = 'lpp_beeaware', STORE = 'kv', KEY = 'registry';

  // The old localStorage copy — drop it so it stops eating the shared quota.
  try { ['lpp_registry', 'lpp_registry_v2', 'lpp_registry_time', 'lpp_registry_v2_time'].forEach(k => localStorage.removeItem(k)); } catch (e) {}

  function open() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async function run(mode, fn) {
    const db = await open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, mode);
      const req = fn(tx.objectStore(STORE));
      tx.oncomplete = () => { db.close(); resolve(req && req.result); };
      tx.onerror = tx.onabort = () => { db.close(); reject(tx.error); };
    });
  }

  window.RegCache = {
    // → { records, time } or null
    async get() { try { return (await run('readonly', s => s.get(KEY))) || null; } catch (e) { return null; } },
    async set(records) { try { await run('readwrite', s => s.put({ records, time: Date.now() }, KEY)); } catch (e) {} },
    async clear() { try { await run('readwrite', s => s.delete(KEY)); } catch (e) {} }
  };
})();
