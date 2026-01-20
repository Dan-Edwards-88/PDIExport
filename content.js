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

      if (aa > 12 && bb <= 12) return buildIsoDate(c, b, a);
      if (bb > 12 && aa <= 12) return buildIsoDate(c, a, b);

      // Ambiguous format; refuse to guess
      return "";
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

  function getDateRangeFromToolbar() {
    const startInput = document.querySelector("kendo-dateinput[kendodaterangestartinput] .k-input");
    const endInput = document.querySelector("kendo-dateinput[kendodaterangeendinput] .k-input");
    const start = parseDateToIso(startInput?.value || "");
    const end = parseDateToIso(endInput?.value || "");
    return { start, end };
  }

  function ensureStatusEl(wrapper) {
    let el = wrapper.querySelector(".pdi-export-status");
    if (!el) {
      el = document.createElement("div");
      el.className = "pdi-export-status";
      el.style.marginLeft = "8px";
      el.style.fontSize = "12px";
      el.style.color = "#333";
      el.style.whiteSpace = "pre-wrap";
      el.style.alignSelf = "center";
      wrapper.appendChild(el);
    }
    return el;
  }

  function showStatus(wrapper, message, tone) {
    const el = ensureStatusEl(wrapper);
    el.textContent = message || "";
    if (tone === "error") el.style.color = "#b00020";
    else if (tone === "success") el.style.color = "#1b5e20";
    else el.style.color = "#333";
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

    btn.addEventListener("click", async () => {
      if (btn.disabled) return;
      const originalHtml = btn.innerHTML;
      btn.disabled = true;
      btn.innerHTML = "<i class=\"fa fa-spinner fa-spin\"></i>";
      try {
        window.postMessage({ type: "PDI_CAPTURE_START", windowMs: 60000 }, "*");

        const { start, end } = getDateRangeFromToolbar();
        if (!start || !end) {
          showStatus(wrapper, "Select a valid start and end date (unambiguous format).", "error");
          return;
        }

        const { profileId: storedProfileId } = await chrome.storage.local.get(["profileId"]);
        const profileId = String(storedProfileId || "39");

        const resp = await chrome.runtime.sendMessage({
          type: "EXPORT_TRIPS",
          payload: { profileId, start, end, language: "en" }
        });

        if (!resp?.ok) throw new Error(resp?.error || "Export failed");
        btn.innerHTML = "<i class=\"fa fa-check\"></i>";
        setTimeout(() => {
          btn.innerHTML = originalHtml;
          btn.disabled = false;
        }, 1200);
        showStatus(wrapper, resp.message || "Export complete", "success");
      } catch (e) {
        btn.innerHTML = "<i class=\"fa fa-exclamation-triangle\"></i>";
        setTimeout(() => {
          btn.innerHTML = originalHtml;
          btn.disabled = false;
        }, 1200);
        showStatus(wrapper, `Export error: ${e?.message || e}`, "error");
        return;
      } finally {
        if (btn.disabled) {
          setTimeout(() => {
            if (btn.disabled) {
              btn.innerHTML = originalHtml;
              btn.disabled = false;
            }
          }, 30000);
        }
      }
    });

    wrapper.appendChild(btn);
  }

  const observer = new MutationObserver(() => injectExportButton());
  observer.observe(document.documentElement, { childList: true, subtree: true });
  injectExportButton();
})();
