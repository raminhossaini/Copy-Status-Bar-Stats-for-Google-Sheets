// background.js (MV3, Firefox)
// Click the toolbar button to copy a Sheets stat (Sum / Avg / Count / Min / Max).
// If multiple stats are visible, an in-page chooser appears so you can pick which one to copy.
const VERSION = "1.2.1"; console.log("[CopySheetsStat] Loaded", VERSION);

browser.action.onClicked.addListener(async (tab) => {
  if (!tab?.id) return;
  if (!/^https:\/\/docs\.google\.com\/spreadsheets\//.test(tab.url || "")) {
    console.warn("[CopySheetsStat] Not a Sheets tab:", tab.url);
    return;
  }

  try {
    const results = await browser.scripting.executeScript({
      target: { tabId: tab.id, allFrames: true },
      func: pageCopySheetsStatToClipboard,
      args: ["auto"]  // or "sum" | "avg" | "count" | "min" | "max"
    });

    if (!results || results.length === 0) {
      console.warn("[CopySheetsStat] executeScript returned no results.");
      await browser.scripting.executeScript({
        target: { tabId: tab.id },
        func: (msg) => {
          const d = document.createElement("div");
          d.textContent = msg;
          Object.assign(d.style, {
            position: "fixed", right: "12px", bottom: "12px",
            padding: "8px 10px", background: "rgba(0,0,0,0.85)",
            color: "#fff", fontSize: "12px", borderRadius: "6px",
            zIndex: 2147483647
          });
          document.body.appendChild(d);
          setTimeout(() => d.remove(), 1500);
        },
        args: ["Couldn’t inject into this page (no frames). Try reloading the Sheet."]
      });
      return;
    }

    const perFrame = results.map(r => r?.result ?? r);
    console.log("[CopySheetsStat] per-frame results:", perFrame);

    const winner = perFrame.find(x => x && x.ok);
    if (winner) {
      console.log(`[CopySheetsStat] Copied OK from frame: ${winner.frame}`);
      return;
    }

    // If we got here, injection ran but didn’t find stats. The page code already toasts,
    // but log again here for clarity.
    console.warn("[CopySheetsStat] No frame succeeded. Last results above.");
  } catch (e) {
    console.error("[CopySheetsStat] executeScript error:", e);
    // Try to surface it in-page too
    try {
      await browser.scripting.executeScript({
        target: { tabId: tab.id },
        func: (msg) => {
          const d = document.createElement("div");
          d.textContent = msg;
          Object.assign(d.style, {
            position: "fixed", right: "12px", bottom: "12px",
            padding: "8px 10px", background: "rgba(160,0,0,0.9)",
            color: "#fff", fontSize: "12px", borderRadius: "6px",
            zIndex: 2147483647
          });
          document.body.appendChild(d);
          setTimeout(() => d.remove(), 1800);
        },
        args: ["Extension error: see Service Worker console for details."]
      });
    } catch {}
  }
});


// ---- injected into the page/frame ----
async function pageCopySheetsStatToClipboard(requested = "auto") {
  // Labels for each metric (expand if needed for your locale)
  const METRIC_LABELS = {
    sum:   ["Sum", "Soma", "Somme", "Summe", "Suma", "Somma", "合計", "Сумма", "Σύνολο"],
    avg:   ["Avg", "Average", "Média", "Moyenne", "Durchschnitt", "Promedio", "Media", "平均", "Среднее"],
    count: ["Count", "Contagem", "Compte", "Anzahl", "Cuenta", "Conteggio", "Recuento", "カウント", "Количество"],
    min:   ["Min", "Mín", "Minimum", "Minimo", "Minimo", "Mínimo", "最小値", "Минимум"],
    max:   ["Max", "Máx", "Maximum", "Massimo", "Máximo", "最大値", "Максимум"]
  };
  const PREFERENCE_ORDER = ["sum", "avg", "count", "min", "max"]; // for "auto"

  // ---------- helpers ----------
  const ALL_LABELS = [];
  const LABEL_TO_METRIC = new Map();
  for (const [metric, labels] of Object.entries(METRIC_LABELS)) {
    for (const lbl of labels) {
      ALL_LABELS.push(lbl);
      LABEL_TO_METRIC.set(lbl.toLowerCase(), metric);
    }
  }
  function escapeReg(s){ return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
  const ALL_LABELS_GROUP = ALL_LABELS.map(escapeReg).join("|");
  const ANY_LABEL_RE = new RegExp(`(?:${ALL_LABELS_GROUP})\\s*:?\\s*`, "i"); // for quick tests
  const PAIR_RE = new RegExp(
    `(${ALL_LABELS_GROUP})\\s*:?\\s*([\\s\\S]*?)(?=(?:${ALL_LABELS_GROUP})\\s*:?\\s*|[\\u2014—|]|$)`,
    "gi"
  );

  function norm(t){ return String(t ?? "").replace(/\s+/g," ").trim(); }

  function parseStats(text) {
    const s = norm(text);
    const stats = {};
    let m;
    while ((m = PAIR_RE.exec(s)) !== null) {
      const rawLabel = m[1];
      const value = (m[2] || "").replace(/[,\s]+$/,"").trim();
      const metric = LABEL_TO_METRIC.get(rawLabel.toLowerCase());
      if (metric && value) stats[metric] = value;
    }
    return stats;
  }

  function isVisible(el) {
    const r = el.getBoundingClientRect?.();
    if (!r || r.width === 0 || r.height === 0) return false;
    const cs = getComputedStyle(el);
    return !(cs.visibility === "hidden" || cs.display === "none");
  }

  function isNearBottomRight(el) {
    const r = el.getBoundingClientRect?.();
    if (!r) return false;
    const closeToBottom = r.bottom > (window.innerHeight - 200);
    const closeToRight  = r.right  > (window.innerWidth  - 360);
    return closeToBottom && closeToRight;
  }

  function findCandidateElement() {
    // Prefer elements in bottom-right region with aria-label or text containing any label
    const roots = [document];
    document.querySelectorAll("*").forEach(n => { if (n.shadowRoot) roots.push(n.shadowRoot); });

    // 1) aria-label near bottom-right
    for (const root of roots) {
      const nodes = root.querySelectorAll("[aria-label]");
      for (const el of nodes) {
        if (!isVisible(el) || !isNearBottomRight(el)) continue;
        const al = el.getAttribute("aria-label") || "";
        if (ANY_LABEL_RE.test(al)) return { el, text: al, via: "aria" };
      }
    }

    // 2) visible text near bottom-right
    const candidates = [];
    for (const root of roots) {
      root.querySelectorAll("div,span,button,[role='status'],[aria-live]").forEach(el => {
        if (!isVisible(el) || !isNearBottomRight(el)) return;
        candidates.push(el);
      });
    }
    // Prefer smaller leaves first
    candidates.sort((a,b) => {
      const ra=a.getBoundingClientRect(), rb=b.getBoundingClientRect();
      return (ra.width*ra.height) - (rb.width*rb.height);
    });
    for (const el of candidates) {
      const txt = norm(el.textContent);
      if (txt && ANY_LABEL_RE.test(txt)) return { el, text: txt, via: "text" };
    }

    // 3) last resort: whole doc aria/text (capped)
    const anyAria = document.querySelectorAll("[aria-label]");
    for (const el of anyAria) {
      const al = el.getAttribute("aria-label") || "";
      if (ANY_LABEL_RE.test(al)) return { el, text: al, via: "aria-doc" };
    }
    const anyText = Array.from(document.querySelectorAll("div,span,button"))
      .filter(isVisible)
      .slice(0, 5000);
    for (const el of anyText) {
      const txt = norm(el.textContent);
      if (txt && ANY_LABEL_RE.test(txt)) return { el, text: txt, via: "text-doc" };
    }

    return null;
  }

  function toast(result) {
    try {
      const msg = result.ok ? (result.message || `Copied: ${result.value}`) : (result.message || "Copy failed.");
      const el = document.createElement("div");
      el.textContent = msg;
      Object.assign(el.style, {
        position: "fixed", right: "12px", bottom: "12px",
        padding: "8px 10px", background: "rgba(0,0,0,0.85)",
        color: "#fff", fontSize: "12px", borderRadius: "6px",
        zIndex: "2147483647", pointerEvents: "none"
      });
      document.body.appendChild(el);
      setTimeout(() => el.remove(), 1200);
    } catch {}
    return { ...result, frame: location.href, showToast: true };
  }

  async function copy(value, label) {
    await navigator.clipboard.writeText(value);
    return toast({ ok: true, value, message: `Copied ${label}: ${value}` });
  }

  function showChooser(stats) {
    const container = document.createElement("div");
    Object.assign(container.style, {
      position: "fixed",
      right: "12px",
      bottom: "12px",
      background: "rgba(0,0,0,0.9)",
      color: "#fff",
      padding: "10px",
      borderRadius: "10px",
      zIndex: "2147483647",
      display: "grid",
      gap: "6px",
      gridAutoFlow: "row",
      boxShadow: "0 6px 20px rgba(0,0,0,0.4)"
    });

    const title = document.createElement("div");
    title.textContent = "Copy which?";
    Object.assign(title.style, { fontSize: "12px", opacity: "0.9", marginBottom: "4px" });
    container.appendChild(title);

    const makeBtn = (metric, label, value) => {
      const btn = document.createElement("button");
      btn.textContent = `${label}: ${value}`;
      Object.assign(btn.style, {
        border: "none",
        padding: "6px 8px",
        borderRadius: "8px",
        cursor: "pointer",
        fontSize: "12px",
        background: "#1f6feb",
        color: "#fff",
      });
      btn.addEventListener("click", async (e) => {
        e.stopPropagation();
        try { await navigator.clipboard.writeText(value); } catch {}
        container.remove();
        toast({ ok: true, value, message: `Copied ${label}: ${value}` });
      });
      return btn;
    };

    const LABEL_FOR = {
      sum: METRIC_LABELS.sum[0],
      avg: METRIC_LABELS.avg[0],
      count: METRIC_LABELS.count[0],
      min: METRIC_LABELS.min[0],
      max: METRIC_LABELS.max[0],
    };

    for (const key of PREFERENCE_ORDER) {
      if (stats[key]) container.appendChild(makeBtn(key, LABEL_FOR[key], stats[key]));
    }

    // click-outside to close
    const close = (ev) => {
      if (!container.contains(ev.target)) {
        container.remove();
        document.removeEventListener("mousedown", close, true);
      }
    };
    document.addEventListener("mousedown", close, true);

    document.body.appendChild(container);
  }

  // ---------- main ----------
  try {
    const cand = findCandidateElement();
    if (!cand) {
      return toast({ ok: false, message: "Couldn't find Sheets status bar. Select numeric cells so stats appear." });
    }

    const stats = parseStats(cand.text);
    const keys = Object.keys(stats);
    if (keys.length === 0) {
      return toast({ ok: false, message: "No stats detected (Sum/Avg/Count/Min/Max). Select numeric cells." });
    }

    // If a specific metric was requested, try that first
    const wanted = String(requested || "").toLowerCase();
    if (["sum","avg","count","min","max"].includes(wanted) && stats[wanted]) {
      return copy(stats[wanted], METRIC_LABELS[wanted][0]);
    }

    // If only one stat is visible, copy it immediately
    if (keys.length === 1) {
      const k = keys[0];
      return copy(stats[k], METRIC_LABELS[k][0]);
    }

    // AUTO: if multiple stats are available, show chooser
    showChooser(stats);
    return toast({ ok: true, message: "Choose a stat to copy…" });
  } catch (err) {
    return toast({ ok: false, message: String(err) });
  }
}
