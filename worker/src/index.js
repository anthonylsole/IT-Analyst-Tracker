// This replaces the QA/POC Worker's code entirely. Every request to this
// instance gets redirected straight to production, preserving the path and
// any query string (e.g. /manage-projects?foo=bar still redirects correctly).

const PRODUCTION_URL = "https://analyst-allocation.delta-it-tools.workers.dev";

export default {
  async fetch(request) {
    const url = new URL(request.url);
    const target = PRODUCTION_URL + url.pathname + url.search;
    // 301 = permanent redirect. Browsers and search engines will remember
    // this and stop hitting the QA URL directly over time. Use 302 instead
    // if you want this to stay easily reversible / not get cached long-term.
    return Response.redirect(target, 301);
  },
};
