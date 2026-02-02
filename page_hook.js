// page_hook.js
(() => {
  const ALLOWED_API_ORIGIN = "https://svc-scheduling.logistics.pdisoftware.com";
  const DEFAULT_CAPTURE_WINDOW_MS = 30000;

  let captureEnabled = false;
  let captureTimer = null;

  function enableCapture(windowMs = DEFAULT_CAPTURE_WINDOW_MS) {
    captureEnabled = true;
    if (captureTimer) clearTimeout(captureTimer);
    captureTimer = setTimeout(() => {
      captureEnabled = false;
      captureTimer = null;
    }, Math.max(1000, Number(windowMs) || DEFAULT_CAPTURE_WINDOW_MS));
  }

  window.addEventListener("message", (event) => {
    if (event.source !== window) return;
    if (!event.data) return;
    if (event.data.type === "PDI_CAPTURE_START") {
      enableCapture(event.data.windowMs);
      return;
    }

    if (event.data.type === "PDI_AUTH_REFRESH") {
      triggerAuthRefresh();
    }
  });

  function isAllowedApiUrl(input) {
    try {
      if (!input) return false;
      const url = typeof input === "string" ? input : input?.url;
      if (!url || typeof url !== "string") return false;
      return url.startsWith(ALLOWED_API_ORIGIN);
    } catch (_) {
      return false;
    }
  }

  function normaliseHeaders(h) {
    const out = {};
    try {
      if (!h) return out;

      // fetch: headers can be Headers, array, or object
      if (h instanceof Headers) {
        for (const [k, v] of h.entries()) out[k.toLowerCase()] = String(v);
        return out;
      }

      if (Array.isArray(h)) {
        for (const [k, v] of h) out[String(k).toLowerCase()] = String(v);
        return out;
      }

      if (typeof h === "object") {
        for (const [k, v] of Object.entries(h)) out[String(k).toLowerCase()] = String(v);
        return out;
      }
    } catch (_) {}
    return out;
  }

  function emitIfFound(headersObj) {
    const auth = headersObj["authorization"];
    const tenantId = headersObj["tenant-id"];

    if (auth && !/^bearer\s+/i.test(String(auth))) return;

    // Only emit when we have something useful
    if (!auth && !tenantId) return;
    if (!captureEnabled) return;

    window.postMessage(
      {
        type: "PDI_AUTH_CAPTURE",
        authorization: auth || null,
        tenantId: tenantId || null
      },
      "*"
    );
  }

  function readCookie(name) {
    try {
      const all = document.cookie || "";
      const parts = all.split(";").map(p => p.trim());
      const match = parts.find(p => p.toLowerCase().startsWith(name.toLowerCase() + "="));
      if (!match) return "";
      return decodeURIComponent(match.substring(match.indexOf("=") + 1));
    } catch (_) {
      return "";
    }
  }

  function getAccessToken() {
    try {
      const sso = sessionStorage.getItem("pdi-auth-result");
      if (sso) {
        const parsed = JSON.parse(sso);
        if (parsed?.accessToken) return String(parsed.accessToken);
      }
    } catch (_) {}

    const cookieToken = readCookie("auth_access_token");
    if (cookieToken) return cookieToken;

    try {
      const lsToken = localStorage.getItem("auth_access_token");
      if (lsToken) return lsToken;
    } catch (_) {}

    return "";
  }

  function getTenantId() {
    try {
      const lsTenant = localStorage.getItem("Tenant-Id") || localStorage.getItem("tenant-id");
      if (lsTenant) return String(lsTenant);
    } catch (_) {}

    const cookieTenant = readCookie("auth_tenant_id") || readCookie("Tenant-Id") || readCookie("tenant-id");
    if (cookieTenant) return cookieTenant;

    return "";
  }

  async function triggerAuthRefresh() {
    try {
      const token = getAccessToken();
      const tenantId = getTenantId();

      if (!token || !tenantId) {
        window.postMessage({
          type: "PDI_AUTH_REFRESH_RESULT",
          ok: false,
          reason: "Missing access token or tenant id"
        }, "*");
        return;
      }

      await fetch(`${ALLOWED_API_ORIGIN}/api/Orders/GetThirdPartyOrdersDashboardData`, {
        method: "GET",
        headers: {
          "accept": "application/json, text/plain, */*",
          "authorization": `Bearer ${token}`,
          "tenant-id": String(tenantId)
        },
        credentials: "include"
      });

      window.postMessage({ type: "PDI_AUTH_REFRESH_RESULT", ok: true }, "*");
    } catch (e) {
      window.postMessage({
        type: "PDI_AUTH_REFRESH_RESULT",
        ok: false,
        reason: e?.message || String(e)
      }, "*");
    }
  }

  // Patch fetch
  const originalFetch = window.fetch;
  window.fetch = async function patchedFetch(input, init) {
    try {
      if (!isAllowedApiUrl(input)) {
        return originalFetch.apply(this, arguments);
      }
      const headersObj = normaliseHeaders(init?.headers);
      emitIfFound(headersObj);
    } catch (_) {}
    return originalFetch.apply(this, arguments);
  };

  // Patch XHR
  const originalOpen = XMLHttpRequest.prototype.open;
  const originalSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;

  XMLHttpRequest.prototype.open = function patchedOpen() {
    this.__pdiHeaders = {};
    this.__pdiUrl = arguments?.[1];
    return originalOpen.apply(this, arguments);
  };

  XMLHttpRequest.prototype.setRequestHeader = function patchedSetRequestHeader(key, value) {
    try {
      if (!isAllowedApiUrl(this.__pdiUrl)) {
        return originalSetRequestHeader.apply(this, arguments);
      }
      this.__pdiHeaders[String(key).toLowerCase()] = String(value);
      emitIfFound(this.__pdiHeaders);
    } catch (_) {}
    return originalSetRequestHeader.apply(this, arguments);
  };
})();
