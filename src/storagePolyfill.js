// Drop-in replacement for the window.storage API that Claude.ai artifacts
// provide. Same method signatures (get/set/list/delete), so the app's
// component code needs zero changes.
//
// - shared === true  -> backed by the Cloudflare Worker + D1 (everyone
//                        behind the shared password sees the same data)
// - shared === false -> backed by browser localStorage (stays private to
//                        this browser, e.g. "which teammate am I viewing as")

const LOCAL_PREFIX = "capacity-local:";

async function apiRequest(path, options) {
  const res = await fetch(path, {
    ...options,
    headers: { "Content-Type": "application/json", ...(options && options.headers) },
  });
  if (res.status === 401) {
    window.location.href = "/login";
    throw new Error("Not authenticated");
  }
  if (res.status === 404) {
    const err = new Error("Not found");
    err.code = "NOT_FOUND"; // normal/expected — key just hasn't been set yet
    throw err;
  }
  if (!res.ok) {
    const err = new Error(`Storage request failed: ${res.status}`);
    err.code = "REQUEST_FAILED"; // a genuine connectivity/server problem
    throw err;
  }
  return res.json();
}

window.storage = {
  async get(key, shared) {
    if (shared) {
      const data = await apiRequest(`/api/kv/get?key=${encodeURIComponent(key)}`, { method: "GET" });
      return { key: data.key, value: data.value, shared: true };
    }
    const raw = window.localStorage.getItem(LOCAL_PREFIX + key);
    if (raw === null) {
      const err = new Error("Not found");
      err.code = "NOT_FOUND";
      throw err;
    }
    return { key, value: raw, shared: false };
  },

  async set(key, value, shared) {
    if (shared) {
      const data = await apiRequest("/api/kv/set", {
        method: "POST",
        body: JSON.stringify({ key, value }),
      });
      return { key: data.key, value: data.value, shared: true };
    }
    window.localStorage.setItem(LOCAL_PREFIX + key, value);
    return { key, value, shared: false };
  },

  async delete(key, shared) {
    if (shared) {
      const data = await apiRequest("/api/kv/delete", {
        method: "POST",
        body: JSON.stringify({ key }),
      });
      return { key: data.key, deleted: true, shared: true };
    }
    window.localStorage.removeItem(LOCAL_PREFIX + key);
    return { key, deleted: true, shared: false };
  },

  async list(prefix, shared) {
    if (shared) {
      const data = await apiRequest(`/api/kv/list?prefix=${encodeURIComponent(prefix || "")}`, {
        method: "GET",
      });
      return { keys: data.keys, prefix: prefix || "", shared: true };
    }
    const keys = [];
    for (let i = 0; i < window.localStorage.length; i++) {
      const k = window.localStorage.key(i);
      if (k && k.startsWith(LOCAL_PREFIX + (prefix || ""))) {
        keys.push(k.slice(LOCAL_PREFIX.length));
      }
    }
    return { keys, prefix: prefix || "", shared: false };
  },
};
