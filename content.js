// content.js
(() => {
  const script = document.createElement("script");
  script.src = chrome.runtime.getURL("page_hook.js");
  script.onload = () => script.remove();
  (document.head || document.documentElement).appendChild(script);

  window.addEventListener("message", (event) => {
    if (event.source !== window) return;
    if (!event.data || event.data.type !== "PDI_AUTH_CAPTURE") return;

    chrome.runtime.sendMessage({
      type: "AUTH_CAPTURED",
      payload: {
        authorization: event.data.authorization,
        tenantId: event.data.tenantId
      }
    });
  });

  const EXPORT_BTN_ID = "pdi-export-xlsx-btn";
  const TOAST_ID = "pdi-export-toast";

  function buildIsoDate(yyyy, mm, dd) {
    const year = Number(yyyy);
    const month = Number(mm);
    const day = Number(dd);
    if (!year || !month || !day) return "";
    const date = new Date(year, month - 1, day);
    if (Number.isNaN(date.getTime())) return "";
    if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return "";
    return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  function parseDateToIso(value) {
    if (!value) return "";
    const raw = value.trim();

    if (raw.includes("/")) {
      const parts = raw.split("/").map(p => p.trim());
      if (parts.length !== 3) return "";
      const [a, b, c] = parts;
      if (!a || !b || !c) return "";

      const aa = Number(a);
      const bb = Number(b);
      if (!aa || !bb) return "";

      if (aa > 12 && bb <= 12) return buildIsoDate(c, b, a); // dd/mm/yyyy
      if (bb > 12 && aa <= 12) return buildIsoDate(c, a, b); // mm/dd/yyyy

      // Ambiguous format (both parts <= 12): assume dd/mm/yyyy (EU format)
      // This handles dates like 01/02/2026 where we can't tell if it's Feb 1 or Jan 2
      if (aa <= 12 && bb <= 12) return buildIsoDate(c, b, a);
    }

    if (raw.includes("-")) {
      const parts = raw.split("-").map(p => p.trim());
      if (parts.length !== 3) return "";
      const [a, b, c] = parts;
      if (a.length === 4) return buildIsoDate(a, b, c);
      if (c.length === 4) return buildIsoDate(c, b, a);
      return "";
    }

    return "";
  }

  function findDateInput(selector, fallbackIdPrefix) {
    const host = document.querySelector(selector);
    if (host) {
      const innerInput = host.querySelector("input.k-input-inner, input.k-input");
      if (innerInput) return innerInput;
    }

    if (fallbackIdPrefix) {
      const fallback = document.querySelector(`input[id^="${fallbackIdPrefix}"]`);
      if (fallback) return fallback;
    }

    return null;
  }

  function getDateRangeFromToolbar() {
    const startInput = findDateInput("kendo-dateinput[kendodaterangestartinput]", "daterangestart-");
    const endInput = findDateInput("kendo-dateinput[kendodaterangeendinput]", "daterangeend-");
    const start = parseDateToIso(startInput?.value || "");
    const end = parseDateToIso(endInput?.value || "");
    return { start, end };
  }

  function ensureToastEl() {
    let el = document.getElementById(TOAST_ID);
    if (!el) {
      el = document.createElement("div");
      el.id = TOAST_ID;
      el.style.position = "fixed";
      el.style.right = "16px";
      el.style.bottom = "16px";
      el.style.zIndex = "99999";
      el.style.padding = "10px 14px";
      el.style.borderRadius = "6px";
      el.style.fontSize = "13px";
      el.style.fontFamily = "inherit";
      el.style.color = "#fff";
      el.style.background = "#333";
      el.style.boxShadow = "0 6px 18px rgba(0, 0, 0, 0.2)";
      el.style.whiteSpace = "pre-wrap";
      el.style.opacity = "0";
      el.style.transform = "translateY(6px)";
      el.style.transition = "opacity 140ms ease, transform 140ms ease";
      el.style.pointerEvents = "none";
      document.body.appendChild(el);
    }
    return el;
  }

  let toastTimer = null;
  function showToast(message, tone = "info", durationMs = 4000) {
    const el = ensureToastEl();
    el.textContent = message || "";
    if (tone === "error") el.style.background = "#c62828";
    else if (tone === "success") el.style.background = "#2e7d32";
    else el.style.background = "#333";

    if (toastTimer) clearTimeout(toastTimer);
    el.style.opacity = "1";
    el.style.transform = "translateY(0)";

    toastTimer = setTimeout(() => {
      el.style.opacity = "0";
      el.style.transform = "translateY(6px)";
    }, Math.max(1200, Number(durationMs) || 4000));
  }

  function injectExportButton() {
    if (document.getElementById(EXPORT_BTN_ID)) return;

    const searchBtn = document.querySelector(".toolbar--secondary .search-icon");
    if (!searchBtn) return;

    const wrapper = searchBtn.closest(".input-group") || searchBtn.parentElement;
    if (!wrapper) return;

    const btn = document.createElement("button");
    btn.id = EXPORT_BTN_ID;
    btn.className = "btn";
    btn.type = "button";
    btn.innerHTML = "<i class=\"fa fa-file-excel-o\"></i>";
    btn.title = "Export XLSX";
    btn.style.marginLeft = "8px";
    btn.style.display = "inline-flex";
    btn.style.alignItems = "center";
    btn.style.justifyContent = "center";
    btn.style.padding = "6px 10px";
    btn.style.borderRadius = "4px";
    btn.style.color = "#2e7d32";
    btn.style.fontSize = "18px";

    btn.addEventListener("click", async () => {
      if (btn.disabled) return;
      const originalHtml = btn.innerHTML;
      btn.disabled = true;
      btn.innerHTML = "<i class=\"fa fa-spinner fa-spin\"></i>";
      try {
        window.postMessage({ type: "PDI_CAPTURE_START", windowMs: 60000 }, "*");
        window.postMessage({ type: "PDI_AUTH_REFRESH" }, "*");

        const { start, end } = getDateRangeFromToolbar();
        if (!start || !end) {
          showToast("Select a valid start and end date (unambiguous format).", "error");
          btn.innerHTML = "<i class=\"fa fa-exclamation-triangle\"></i>";
          setTimeout(() => {
            btn.innerHTML = originalHtml;
            btn.disabled = false;
          }, 1200);
          return;
        }

        const { profileId: storedProfileId } = await chrome.storage.local.get(["profileId"]);
        const profileId = String(storedProfileId || "39");

        const resp = await chrome.runtime.sendMessage({
          type: "EXPORT_TRIPS",
          payload: { profileId, start, end, language: "en" }
        });

        if (!resp) {
          throw new Error("No response from background script - extension may need to be reloaded");
        }
        if (!resp.ok) throw new Error(resp.error || "Export failed");
        btn.innerHTML = "<i class=\"fa fa-check\"></i>";
        setTimeout(() => {
          btn.innerHTML = originalHtml;
          btn.disabled = false;
        }, 1200);
        showToast(resp.message || "Export complete", "success");
      } catch (e) {
        btn.innerHTML = "<i class=\"fa fa-exclamation-triangle\"></i>";
        setTimeout(() => {
          btn.innerHTML = originalHtml;
          btn.disabled = false;
        }, 1200);
        showToast(`Export error: ${e?.message || e}`, "error");
      }
    });

    wrapper.appendChild(btn);
  }

  const observer = new MutationObserver(() => injectExportButton());
  observer.observe(document.documentElement, { childList: true, subtree: true });
  injectExportButton();
})();
