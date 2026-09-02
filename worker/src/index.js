// Cloudflare Worker for the DELTA IT Resource and Project Allocation tracker.
//
// Responsibilities:
//   1. Gate every request behind a single shared password (server-side check,
//      not just a client-side JS prompt).
//   2. Serve a small key/value API backed by D1, matching the interface the
//      React app already expects (get/set/list/delete).
//   3. Serve the built frontend (the Vite output in /dist) for everything else.
//
// Required secrets — either method works, the code below handles both:
//   Plain Worker secret:   wrangler secret put SITE_PASSWORD / SESSION_SECRET
//   Secrets Store binding: added via the dashboard's Bindings tab
//   SITE_PASSWORD   - the plaintext password visitors must enter
//   SESSION_SECRET  - a random string used as the session cookie's value
//
// Required bindings (see wrangler.toml):
//   DB      - the D1 database
//   ASSETS  - the built static frontend

const COOKIE_NAME = "capacity_session";
const MAX_AGE_SECONDS = 60 * 60 * 24 * 30; // 30 days

// A Secrets Store binding is an object you must call .get() on (async).
// A plain `wrangler secret put` value is just a string already.
// This works with either, so it doesn't matter which one was used.
async function resolveSecret(value) {
  if (value && typeof value.get === "function") {
    return await value.get();
  }
  return value;
}

function getCookie(request, name) {
  const header = request.headers.get("Cookie") || "";
  const match = header.match(new RegExp(`(?:^|;\\s*)${name}=([^;]*)`));
  return match ? decodeURIComponent(match[1]) : null;
}

async function isAuthenticated(request, env) {
  const cookie = getCookie(request, COOKIE_NAME);
  if (cookie === null) return false;
  const sessionSecret = await resolveSecret(env.SESSION_SECRET);
  return Boolean(sessionSecret) && cookie === sessionSecret;
}

function loginPageHtml(error) {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>DELTA IT Resource and Project Allocation — Sign in</title>
  <style>
    body {
      font-family: -apple-system, "Segoe UI", sans-serif;
      background: linear-gradient(180deg, #DCEBF7 0%, #FFFFFF 55%);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 0;
    }
    form {
      background: #fff;
      border: 1px solid #D9DCD1;
      border-radius: 14px;
      padding: 32px;
      width: 280px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.06);
    }
    h1 { font-size: 16px; margin: 0 0 18px; color: #1B2320; }
    input {
      width: 100%;
      box-sizing: border-box;
      padding: 10px 12px;
      border: 1px solid #D9DCD1;
      border-radius: 8px;
      font-size: 14px;
      margin-bottom: 12px;
    }
    button {
      width: 100%;
      padding: 10px 12px;
      background: #1F6F63;
      color: #fff;
      border: none;
      border-radius: 8px;
      font-size: 14px;
      cursor: pointer;
    }
    button:hover { background: #185349; }
    .error { color: #C1443C; font-size: 12.5px; margin: -6px 0 12px; }
  </style>
</head>
<body>
  <form method="POST" action="/login">
    <h1>Enter password to continue</h1>
    ${error ? `<div class="error">Incorrect password — try again.</div>` : ""}
    <input type="password" name="password" placeholder="Password" autofocus />
    <button type="submit">Sign in</button>
  </form>
</body>
</html>`;
}

async function handleLogin(request, env) {
  if (request.method === "GET") {
    return new Response(loginPageHtml(false), { headers: { "Content-Type": "text/html" } });
  }

  const formData = await request.formData();
  const password = formData.get("password") || "";

  const sitePassword = await resolveSecret(env.SITE_PASSWORD);
  const sessionSecret = await resolveSecret(env.SESSION_SECRET);

  if (!sitePassword || !sessionSecret) {
    const missing = [!sitePassword && "SITE_PASSWORD", !sessionSecret && "SESSION_SECRET"]
      .filter(Boolean)
      .join(" and ");
    return new Response(
      `Server misconfigured: ${missing} not set. Add ${missing} as a Worker secret or Secrets Store binding (variable name must match exactly), then redeploy.`,
      { status: 500, headers: { "Content-Type": "text/plain" } }
    );
  }

  if (password !== sitePassword) {
    return new Response(loginPageHtml(true), {
      status: 401,
      headers: { "Content-Type": "text/html" },
    });
  }

  const headers = new Headers();
  headers.set("Location", "/");
  headers.append(
    "Set-Cookie",
    `${COOKIE_NAME}=${encodeURIComponent(sessionSecret)}; HttpOnly; Secure; Path=/; Max-Age=${MAX_AGE_SECONDS}; SameSite=Lax`
  );
  return new Response(null, { status: 302, headers });
}

async function handleKvApi(request, env, url) {
  const path = url.pathname;

  if (path === "/api/kv/get" && request.method === "GET") {
    const key = url.searchParams.get("key");
    if (!key) return jsonResponse({ error: "missing key" }, 400);
    const row = await env.DB.prepare("SELECT value FROM kv_store WHERE key = ?").bind(key).first();
    if (!row) return jsonResponse({ error: "not_found" }, 404);
    return jsonResponse({ key, value: row.value });
  }

  if (path === "/api/kv/set" && request.method === "POST") {
    const body = await request.json();
    if (!body || typeof body.key !== "string" || typeof body.value !== "string") {
      return jsonResponse({ error: "key and value (strings) required" }, 400);
    }
    await env.DB.prepare(
      `INSERT INTO kv_store (key, value, updated_at) VALUES (?, ?, datetime('now'))
       ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at`
    )
      .bind(body.key, body.value)
      .run();
    return jsonResponse({ key: body.key, value: body.value });
  }

  if (path === "/api/kv/delete" && request.method === "POST") {
    const body = await request.json();
    if (!body || typeof body.key !== "string") {
      return jsonResponse({ error: "key (string) required" }, 400);
    }
    await env.DB.prepare("DELETE FROM kv_store WHERE key = ?").bind(body.key).run();
    return jsonResponse({ key: body.key, deleted: true });
  }

  if (path === "/api/kv/list" && request.method === "GET") {
    const prefix = url.searchParams.get("prefix") || "";
    const { results } = await env.DB.prepare("SELECT key FROM kv_store WHERE key LIKE ?")
      .bind(`${prefix}%`)
      .all();
    return jsonResponse({ keys: results.map((r) => r.key) });
  }

  return jsonResponse({ error: "not_found" }, 404);
}

function jsonResponse(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    // Login page and its POST handler are always reachable, unauthenticated.
    if (url.pathname === "/login") {
      return handleLogin(request, env);
    }

    const authed = await isAuthenticated(request, env);

    if (url.pathname.startsWith("/api/")) {
      if (!authed) return jsonResponse({ error: "unauthorized" }, 401);
      return handleKvApi(request, env, url);
    }

    if (!authed) {
      // Not logged in — send everything else to the login page.
      return Response.redirect(`${url.origin}/login`, 302);
    }

    // Authenticated — serve the built frontend.
    return env.ASSETS.fetch(request);
  },
};
