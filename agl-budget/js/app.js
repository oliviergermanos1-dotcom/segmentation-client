/* ============================================================================
 * AGL BUDGET — app.js  (contrôleur / bootstrap UI)
 * Câble la coquille : navigation 14 onglets, cartes d'upload des 6 bases,
 * progression, snapshots, aperçu, panneau macro. Sprint 1 : seul l'onglet 01
 * (Matching & RMC) et l'onglet Macro sont fonctionnels ; 02–14 sont des stubs.
 * Dépend de : AGL.utils, AGL.loader, AGL.store, AGL.charts, MACRO_DATA.
 * ==========================================================================*/
(function (global) {
  "use strict";
  const U = global.AGL.utils, loader = global.AGL.loader, store = global.AGL.store;
  const { $, el } = U;

  // Définition des 14 onglets (§1) + onglet Macro transverse.
  const TABS = [
    { id: "01", label: "01 · Matching & RMC" },
    { id: "02", label: "02 · Marché STATCOM" },
    { id: "03", label: "03 · Position AGL" },
    { id: "04", label: "04 · Budget PFA vs Réel" },
    { id: "05", label: "05 · Pipeline CRM" },
    { id: "06", label: "06 · Segmentation BCG" },
    { id: "07", label: "07 · White Spaces" },
    { id: "08", label: "08 · Carte PND" },
    { id: "09", label: "09 · Top 20 Projets PND" },
    { id: "10", label: "10 · Projections 2026-2030" },
    { id: "11", label: "11 · Préconisations Client" },
    { id: "12", label: "12 · Synthèse Executive" },
    { id: "13", label: "13 · Analyse Concurrentielle", star: true },
    { id: "14", label: "14 · Cohortes & Rétention", star: true },
    { id: "macro", label: "📈 Macro" }
  ];

  function buildTabs() {
    const nav = $("#tabs"), main = $("main");
    TABS.forEach((t) => {
      const btn = el("button", { class: "tab", "data-tab": t.id });
      btn.innerHTML = t.label + (t.star ? ' <span class="star">★</span>' : "");
      btn.addEventListener("click", () => activate(t.id));
      nav.appendChild(btn);
      // Crée les panneaux stub manquants (tous sauf 01 & macro déjà dans le HTML).
      if (!$('[data-panel="' + t.id + '"]')) {
        const panel = el("section", { class: "panel", "data-panel": t.id });
        panel.appendChild(el("div", { class: "section" }, [
          el("h2", { text: t.label.replace(/^\d+ · /, "") }),
          el("div", { class: "placeholder", html: "Onglet prévu au planning des sprints (cf. README §8). <br>Non encore implémenté — Sprint 1 = Fondations." })
        ]));
        main.appendChild(panel);
      }
    });
  }

  function activate(id) {
    U.$$(".tab").forEach((b) => b.classList.toggle("active", b.getAttribute("data-tab") === id));
    U.$$(".panel").forEach((p) => p.classList.toggle("active", p.getAttribute("data-panel") === id));
  }

  /* ----- Cartes d'upload des 6 bases ----- */
  function buildSourceCards() {
    const grid = $("#sources-grid");
    Object.keys(loader.SOURCES).forEach((key) => {
      const s = loader.SOURCES[key];
      const isLocked = /LOCKED/.test(s.priorite);
      const card = el("div", { class: "src-card" });
      card.appendChild(el("div", { class: "prio" + (isLocked ? " locked" : ""), text: s.priorite }));
      card.appendChild(el("strong", { text: s.label }));
      card.appendChild(el("div", { class: "role", text: s.role }));
      const input = el("input", { type: "file", accept: ".xlsx,.xls,.csv" });
      const status = el("div", { class: "src-status muted", text: "" });
      input.addEventListener("change", () => {
        if (!input.files || !input.files[0]) return;
        handleUpload(key, input.files[0], status);
      });
      card.appendChild(input);
      card.appendChild(status);
      grid.appendChild(card);
    });
  }

  function setProgress(pct, show) {
    const p = $("#progress");
    p.classList.toggle("show", show !== false);
    p.firstElementChild.style.width = (pct || 0) + "%";
    if (pct >= 100) setTimeout(() => p.classList.remove("show"), 400);
  }

  function handleUpload(key, file, statusEl) {
    statusEl.className = "src-status muted";
    statusEl.textContent = "Lecture…";
    const t0 = performance.now();
    loader.loadFile(key, file, {
      onProgress: (m) => { setProgress(m.pct); statusEl.textContent = m.phase + " — " + U.formatNumber(m.rows) + " lignes…"; },
      onDone: (snapshot) => {
        setProgress(100);
        const dt = ((performance.now() - t0) / 1000).toFixed(1);
        statusEl.className = "src-status ok";
        statusEl.textContent = "✓ " + U.formatNumber(snapshot.rows) + " lignes · " + snapshot.columns.length + " colonnes · " + dt + " s";
        renderPreview(snapshot);
        refreshLoaded();
      },
      onError: (err) => { setProgress(0, false); statusEl.className = "src-status err"; statusEl.textContent = "✗ " + err.message; }
    });
  }

  function renderPreview(snapshot) {
    const box = $("#preview");
    box.className = "";
    box.innerHTML = "";
    box.appendChild(el("div", { class: "meta-bar" }, [
      el("span", { html: "<span class='badge'>" + snapshot.source + "</span>" }),
      el("span", { text: snapshot.name }),
      el("span", { text: U.formatNumber(snapshot.rows) + " lignes" }),
      el("span", { text: snapshot.columns.length + " colonnes" }),
      el("span", { text: "en-tête ligne " + (snapshot.headerRow + 1) })
    ]));
    const wrap = el("div", { class: "table-wrap", style: "max-height:340px;overflow:auto;margin-top:8px" });
    const table = el("table"), thead = el("thead"), trh = el("tr");
    snapshot.columns.forEach((c) => trh.appendChild(el("th", { text: c })));
    thead.appendChild(trh); table.appendChild(thead);
    const tbody = el("tbody");
    snapshot.sample.slice(0, 20).forEach((row) => {
      const tr = el("tr");
      snapshot.columns.forEach((c) => tr.appendChild(el("td", { text: row[c] == null ? "" : String(row[c]) })));
      tbody.appendChild(tr);
    });
    table.appendChild(tbody); wrap.appendChild(table); box.appendChild(wrap);
  }

  function refreshLoaded() {
    Promise.all([store.list("sources"), store.quota()]).then(([items, q]) => {
      const box = $("#loaded-list");
      if (!items.length) { box.className = "placeholder"; box.textContent = "—"; return; }
      box.className = "";
      box.innerHTML = "";
      const table = el("table");
      table.innerHTML = "<thead><tr><th>Source</th><th>Fichier</th><th>Lignes</th><th>Colonnes</th><th>Chargé</th></tr></thead>";
      const tb = el("tbody");
      items.forEach((it) => {
        const v = it.value;
        const tr = el("tr");
        tr.innerHTML = "<td><span class='badge'>" + v.source + "</span></td><td>" + (v.name || "") +
          "</td><td>" + U.formatNumber(v.rows) + "</td><td>" + (v.columns ? v.columns.length : 0) +
          "</td><td>" + new Date(v.loadedAt).toLocaleString("fr-FR") + "</td>";
        tb.appendChild(tr);
      });
      table.appendChild(tb); box.appendChild(table);
      if (q && q.usage != null) $("#quota-bar").textContent =
        "Stockage local : " + U.formatBytes(q.usage) + " / " + U.formatBytes(q.quota);
    });
  }

  /* ----- Panneau macro ----- */
  function buildMacro() {
    const D = global.MACRO_DATA;
    const groups = [
      ["Macro Côte d'Ivoire", D.civ], ["Commerce mondial — OMC", D.commerce_mondial],
      ["Maritime conteneurs", D.maritime_conteneurs], ["Aérien cargo", D.aerien_cargo],
      ["Hinterland & AfCFTA", D.hinterland_afcfta]
    ];
    const host = $("#macro-tables");
    groups.forEach(([title, rows]) => {
      host.appendChild(el("h3", { text: title }));
      const t = el("table");
      t.innerHTML = "<thead><tr><th>Indicateur</th><th>Valeur</th><th>Source</th><th>Date</th></tr></thead>";
      const tb = el("tbody");
      rows.forEach((r) => {
        const tr = el("tr");
        tr.innerHTML = "<td>" + r.indicateur + "</td><td><strong>" + r.valeur + "</strong></td><td>" + r.source + "</td><td>" + r.date + "</td>";
        tb.appendChild(tr);
      });
      t.appendChild(tb); host.appendChild(t);
    });
    const ro = $("#ro-grid");
    D.risques_opportunites.forEach((x) => {
      ro.appendChild(el("div", { class: "item " + x.type }, [
        el("div", { class: "t", text: (x.type === "RISQUE" ? "⚠️ " : "🚀 ") + x.titre }),
        el("div", { class: "muted", text: x.detail })
      ]));
    });
  }

  function init() {
    buildTabs();
    buildSourceCards();
    buildMacro();
    activate("01");
    const w = loader.memoryWarning();
    if (w) $("#mem-warn").textContent = "⚠️ " + w;
    store.open().then(refreshLoaded).catch((e) => U.log("app", "IndexedDB: " + e.message, "warn"));
  }

  document.addEventListener("DOMContentLoaded", init);
})(typeof window !== "undefined" ? window : this);
