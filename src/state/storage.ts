/** Tiny IndexedDB key/value store (works in Android WebView and browsers). */
const DB = 'opencells';
const STORE = 'kv';

let dbp: Promise<IDBDatabase> | null = null;
function db(): Promise<IDBDatabase> {
  if (!dbp) {
    dbp = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  return dbp;
}

function tx<T>(mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  return db().then(
    (d) =>
      new Promise<T>((resolve, reject) => {
        const t = d.transaction(STORE, mode);
        const req = fn(t.objectStore(STORE));
        t.oncomplete = () => resolve(req.result);
        t.onerror = () => reject(t.error);
        t.onabort = () => reject(t.error);
      }),
  );
}

export const kv = {
  get: <T>(key: string) => tx<T | undefined>('readonly', (s) => s.get(key) as IDBRequest<T | undefined>),
  set: (key: string, value: unknown) => tx('readwrite', (s) => s.put(value, key)),
  del: (key: string) => tx('readwrite', (s) => s.delete(key)),
  delPrefix: (prefix: string) => tx('readwrite', (s) => s.delete(IDBKeyRange.bound(prefix, prefix + '\uffff'))),
};

/** Small synchronous preferences (theme, last panel) — localStorage with try/catch. */
export const prefs = {
  get(key: string, def = ''): string {
    try {
      return localStorage.getItem('oc.' + key) ?? def;
    } catch {
      return def;
    }
  },
  set(key: string, v: string) {
    try {
      localStorage.setItem('oc.' + key, v);
    } catch {
      /* private mode */
    }
  },
};
