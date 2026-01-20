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
    if (!event.data || event.data.type !== "PDI_CAPTURE_START") return;
    enableCapture(event.data.windowMs);
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
