/* ============================================================================
 * AGL BUDGET — charts.js
 * Helpers de visualisation sans dépendance (canvas natif) — COMEX-ready :
 * bar, line, bubble (analyse concurrentielle §13), donut (parts de marché).
 * Palette sobre type cabinet de conseil. Rendu nettoyé/responsive simple.
 * Expose window.AGL.charts. Aucune lib externe (offline-safe).
 * ==========================================================================*/
(function (global) {
  "use strict";

  const PALETTE = ["#0b3d91", "#2563eb", "#16a34a", "#d97706", "#dc2626",
                   "#7c3aed", "#0891b2", "#65a30d", "#be185d", "#475569"];

  function _ctx(canvas) {
    const dpr = global.devicePixelRatio || 1;
    const rect = canvas.getBoundingClientRect();
    const w = rect.width || canvas.width || 600;
    const h = rect.height || canvas.height || 320;
    canvas.width = w * dpr; canvas.height = h * dpr;
    const ctx = canvas.getContext("2d");
    ctx.scale(dpr, dpr);
    ctx.clearRect(0, 0, w, h);
    ctx.font = "12px -apple-system, Segoe UI, Roboto, sans-serif";
    return { ctx, w, h };
  }

  function _niceMax(v) {
    if (v <= 0) return 1;
    const mag = Math.pow(10, Math.floor(Math.log10(v)));
    const n = v / mag;
    const step = n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10;
    return step * mag;
  }

  // data : [{label, value}]
  function bar(canvas, data, opts) {
    opts = opts || {};
    const { ctx, w, h } = _ctx(canvas);
    const padL = 48, padB = 40, padT = 16, padR = 12;
    const max = _niceMax(Math.max.apply(null, data.map((d) => d.value).concat([0])));
    const plotW = w - padL - padR, plotH = h - padT - padB;
    // axes + grille
    ctx.strokeStyle = "#e3e6ea"; ctx.fillStyle = "#6b7280"; ctx.textAlign = "right";
    for (let g = 0; g <= 4; g++) {
      const y = padT + plotH - (plotH * g / 4);
      ctx.beginPath(); ctx.moveTo(padL, y); ctx.lineTo(w - padR, y); ctx.stroke();
      ctx.fillText(Math.round(max * g / 4).toLocaleString("fr-FR"), padL - 6, y + 4);
    }
    const bw = plotW / data.length * 0.62;
    data.forEach((d, i) => {
      const x = padL + plotW * (i + 0.5) / data.length - bw / 2;
      const bh = max ? (d.value / max) * plotH : 0;
      ctx.fillStyle = opts.color || PALETTE[i % PALETTE.length];
      ctx.fillRect(x, padT + plotH - bh, bw, bh);
      ctx.fillStyle = "#6b7280"; ctx.textAlign = "center";
      ctx.fillText(String(d.label), x + bw / 2, h - padB + 16);
    });
    if (opts.title) { ctx.fillStyle = "#1f2937"; ctx.textAlign = "left"; ctx.fillText(opts.title, padL, 12); }
  }

  // series : [{label, points:[{x,y}]}] — utile projections multi-scénarios.
  function line(canvas, series, opts) {
    opts = opts || {};
    const { ctx, w, h } = _ctx(canvas);
    const padL = 52, padB = 36, padT = 16, padR = 12;
    const allY = series.reduce((a, s) => a.concat(s.points.map((p) => p.y)), [0]);
    const allX = series.reduce((a, s) => a.concat(s.points.map((p) => p.x)), []);
    const maxY = _niceMax(Math.max.apply(null, allY));
    const minX = Math.min.apply(null, allX), maxX = Math.max.apply(null, allX);
    const plotW = w - padL - padR, plotH = h - padT - padB;
    const X = (x) => padL + (maxX === minX ? 0 : (x - minX) / (maxX - minX)) * plotW;
    const Y = (y) => padT + plotH - (maxY ? (y / maxY) * plotH : 0);
    ctx.strokeStyle = "#e3e6ea"; ctx.fillStyle = "#6b7280"; ctx.textAlign = "right";
    for (let g = 0; g <= 4; g++) {
      const y = padT + plotH - (plotH * g / 4);
      ctx.beginPath(); ctx.moveTo(padL, y); ctx.lineTo(w - padR, y); ctx.stroke();
      ctx.fillText(Math.round(maxY * g / 4).toLocaleString("fr-FR"), padL - 6, y + 4);
    }
    series.forEach((s, i) => {
      ctx.strokeStyle = PALETTE[i % PALETTE.length]; ctx.lineWidth = 2; ctx.beginPath();
      s.points.forEach((p, j) => { const fn = j ? "lineTo" : "moveTo"; ctx[fn](X(p.x), Y(p.y)); });
      ctx.stroke();
    });
    ctx.textAlign = "center"; ctx.fillStyle = "#6b7280";
    series[0] && series[0].points.forEach((p) => ctx.fillText(String(p.x), X(p.x), h - padB + 16));
  }

  // points : [{x, y, r, label}] — bubble chart analyse concurrentielle (§13).
  function bubble(canvas, points, opts) {
    opts = opts || {};
    const { ctx, w, h } = _ctx(canvas);
    const padL = 52, padB = 40, padT = 16, padR = 16;
    const maxX = _niceMax(Math.max.apply(null, points.map((p) => p.x).concat([0])));
    const maxY = _niceMax(Math.max.apply(null, points.map((p) => p.y).concat([0])));
    const maxR = Math.max.apply(null, points.map((p) => p.r || 1).concat([1]));
    const plotW = w - padL - padR, plotH = h - padT - padB;
    const X = (x) => padL + (maxX ? x / maxX : 0) * plotW;
    const Y = (y) => padT + plotH - (maxY ? y / maxY : 0) * plotH;
    ctx.strokeStyle = "#e3e6ea";
    ctx.beginPath(); ctx.moveTo(padL, padT); ctx.lineTo(padL, padT + plotH); ctx.lineTo(w - padR, padT + plotH); ctx.stroke();
    points.forEach((p, i) => {
      const r = 6 + Math.sqrt((p.r || 1) / maxR) * 34;
      ctx.fillStyle = (PALETTE[i % PALETTE.length]) + "cc";
      ctx.beginPath(); ctx.arc(X(p.x), Y(p.y), r, 0, Math.PI * 2); ctx.fill();
      if (p.label) { ctx.fillStyle = "#1f2937"; ctx.textAlign = "center"; ctx.fillText(p.label, X(p.x), Y(p.y) - r - 4); }
    });
    if (opts.xLabel) { ctx.fillStyle = "#6b7280"; ctx.textAlign = "right"; ctx.fillText(opts.xLabel, w - padR, h - 6); }
    if (opts.yLabel) { ctx.save(); ctx.translate(12, padT + 6); ctx.textAlign = "left"; ctx.fillText(opts.yLabel, 0, 0); ctx.restore(); }
  }

  // data : [{label, value}] — donut parts de marché.
  function donut(canvas, data, opts) {
    opts = opts || {};
    const { ctx, w, h } = _ctx(canvas);
    const cx = w / 2, cy = h / 2, R = Math.min(w, h) / 2 - 12, r = R * 0.58;
    const total = data.reduce((a, d) => a + d.value, 0) || 1;
    let ang = -Math.PI / 2;
    data.forEach((d, i) => {
      const a2 = ang + (d.value / total) * Math.PI * 2;
      ctx.fillStyle = PALETTE[i % PALETTE.length];
      ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, R, ang, a2); ctx.closePath(); ctx.fill();
      ang = a2;
    });
    ctx.fillStyle = "#fff"; ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.fill();
    if (opts.center) { ctx.fillStyle = "#1f2937"; ctx.textAlign = "center"; ctx.font = "bold 16px -apple-system, sans-serif"; ctx.fillText(opts.center, cx, cy + 5); }
  }

  global.AGL = global.AGL || {};
  global.AGL.charts = { bar, line, bubble, donut, PALETTE };
})(typeof window !== "undefined" ? window : this);
