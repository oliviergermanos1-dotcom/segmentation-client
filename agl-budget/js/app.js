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
    // Rendu paresseux des onglets analytiques (chaque activation rafraîchit).
    if (id === "02") renderMarketTab();
    else if (id === "03") renderPositionTab();
    else if (id === "04") renderBudgetTab();
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

  /* ====================================================================== */
  /*  Import des artefacts du pipeline Python                                */
  /* ====================================================================== */
  function bindPipelineLoader() {
    const input = $("#pipeline-input");
    if (!input) return;
    const pipe = global.AGL.pipeline;
    input.addEventListener("change", () => {
      const files = input.files;
      if (!files || !files.length) return;
      const status = $("#pipeline-status");
      status.textContent = "Lecture…";
      pipe.loadFiles(files, (frac, name, st) => {
        status.textContent = "(" + Math.round(frac * 100) + " %) " + name + " → " + st;
      }).then((res) => {
        const ok = res.filter((r) => r.status === "ok").length;
        const ign = res.filter((r) => r.status === "ignored").length;
        const err = res.filter((r) => r.status === "error").length;
        status.innerHTML = "<span class='ok'>" + ok + " chargé(s)</span>" +
          (ign ? " · <span class='muted'>" + ign + " ignoré(s)</span>" : "") +
          (err ? " · <span class='err'>" + err + " erreur(s)</span>" : "");
        refreshPipelineLoaded();
      });
    });
    refreshPipelineLoaded();
  }
  function refreshPipelineLoaded() {
    const pipe = global.AGL.pipeline;
    pipe.listLoaded().then((items) => {
      const host = $("#pipeline-loaded");
      if (!items.length) { host.innerHTML = ""; return; }
      host.innerHTML = items.map((it) =>
        "<span class='badge'>" + it.key + "</span> " + (it.value.name || "") +
        " · " + U.formatNumber(it.value.rows.length) + " lignes"
      ).join(" · ");
    });
  }

  /* ====================================================================== */
  /*  Onglet 02 — Marché Global STATCOM                                      */
  /* ====================================================================== */
  function fillYearSelect(selectId, years, defaultYear) {
    const sel = $(selectId);
    if (!sel) return;
    const current = sel.value;
    sel.innerHTML = "";
    years.forEach((y) => sel.appendChild(el("option", { value: y, text: y })));
    sel.value = current && years.indexOf(+current) > -1 ? current : (defaultYear || years[years.length - 1]);
  }

  function renderMarketTab() {
    const pipe = global.AGL.pipeline;
    pipe.availableYears().then((years) => {
      if (!years.length) { $("#m02-placeholder").style.display = ""; $("#m02-content").style.display = "none"; return; }
      $("#m02-placeholder").style.display = "none";
      $("#m02-content").style.display = "";
      fillYearSelect("#m02-year", years);
      const sel = $("#m02-year");
      sel.onchange = () => drawMarket(+sel.value);
      drawMarket(+sel.value);
    });
  }
  function drawMarket(year) {
    const pipe = global.AGL.pipeline;
    Promise.all([pipe.marketTotalsByMetier(year), pipe.marketByMetierYear()]).then(([byM, all]) => {
      const sorted = byM.sort((a, b) => b.value - a.value);
      global.AGL.charts.bar($("#m02-chart-mix"), sorted, { title: "Marché total STATCOM " + year + " — volume principal par métier" });
      // Table — toutes années pour donner la vision 3 ans.
      const tbl = el("table");
      tbl.innerHTML = "<thead><tr><th>Métier</th><th>Année</th><th>Vol. TEU</th><th>Vol. Kg</th><th>Nb clients</th><th>Nb ops</th></tr></thead>";
      const tb = el("tbody");
      all.sort((a, b) => (a.METIER || "").localeCompare(b.METIER || "") || (+a.ANNEE) - (+b.ANNEE))
         .forEach((r) => {
        const tr = el("tr");
        tr.innerHTML = "<td>" + (r.METIER || "") + "</td><td>" + r.ANNEE + "</td>" +
          "<td>" + U.formatNumber(r.VOLUME_TEU) + "</td>" +
          "<td>" + U.formatNumber(r.VOLUME_KG) + "</td>" +
          "<td>" + U.formatNumber(r.NB_CLIENTS) + "</td>" +
          "<td>" + U.formatNumber(r.NB_OPERATIONS) + "</td>";
        tb.appendChild(tr);
      });
      tbl.appendChild(tb);
      const host = $("#m02-table"); host.innerHTML = ""; host.appendChild(tbl);
    });
  }

  /* ====================================================================== */
  /*  Onglet 03 — Position AGL                                                */
  /* ====================================================================== */
  function renderPositionTab() {
    const pipe = global.AGL.pipeline;
    pipe.get("PDM").then((rows) => {
      if (!rows || !rows.length) { $("#m03-placeholder").style.display = ""; $("#m03-content").style.display = "none"; return; }
      $("#m03-placeholder").style.display = "none";
      $("#m03-content").style.display = "";
      const years = Array.from(new Set(rows.map((r) => +r.ANNEE))).sort();
      fillYearSelect("#m03-year", years);
      const sel = $("#m03-year");
      sel.onchange = () => drawPosition(+sel.value);
      drawPosition(+sel.value);
    });
  }
  function drawPosition(year) {
    global.AGL.pipeline.pdmByMetier(year).then((rows) => {
      const data = rows.sort((a, b) => b.pdm - a.pdm).map((r) => ({ label: r.metier, value: r.pdm * 100 }));
      global.AGL.charts.bar($("#m03-chart-pdm"), data, { title: "PDM AGL " + year + " (%)" });
      const tbl = el("table");
      tbl.innerHTML = "<thead><tr><th>Métier</th><th>Vol. AGL</th><th>Vol. Marché</th><th>PDM</th><th>CAGR AGL</th><th>CAGR Marché</th></tr></thead>";
      const tb = el("tbody");
      rows.forEach((r) => {
        const tr = el("tr");
        tr.innerHTML = "<td>" + r.metier + "</td>" +
          "<td>" + U.formatNumber(r.volume_agl) + "</td>" +
          "<td>" + U.formatNumber(r.volume_marche) + "</td>" +
          "<td><strong>" + (r.pdm * 100).toFixed(1) + " %</strong></td>" +
          "<td>" + (r.cagr_agl != null ? (r.cagr_agl * 100).toFixed(1) + " %" : "—") + "</td>" +
          "<td>" + (r.cagr_marche != null ? (r.cagr_marche * 100).toFixed(1) + " %" : "—") + "</td>";
        tb.appendChild(tr);
      });
      tbl.appendChild(tb);
      const host = $("#m03-table"); host.innerHTML = ""; host.appendChild(tbl);
    });
  }

  /* ====================================================================== */
  /*  Onglet 04 — Budget PFA vs Réel                                          */
  /* ====================================================================== */
  function renderBudgetTab() {
    Promise.all([store.get("sources", "RUBRIKS"), store.get("sources", "IRIS")]).then(([ru, ir]) => {
      if (!ru || !ir) { $("#m04-placeholder").style.display = ""; $("#m04-content").style.display = "none"; return; }
      // Sprint 1 ne persiste que les snapshots légers (sample). Pour le calcul complet,
      // il faut re-charger les bases — message clair en attendant le pipeline 04.
      $("#m04-placeholder").style.display = "";
      $("#m04-content").style.display = "none";
      $("#m04-placeholder").innerHTML =
        "Onglet 04 — fonctionnel à brancher au Sprint 4 (jointure RUBRIKS × IRIS via RMC). " +
        "RUBRIKS chargé (" + U.formatNumber(ru.rows) + " lignes), IRIS chargé (" + U.formatNumber(ir.rows) + " lignes).";
    });
  }

  function init() {
    buildTabs();
    buildSourceCards();
    buildMacro();
    bindPipelineLoader();
    activate("01");
    const w = loader.memoryWarning();
    if (w) $("#mem-warn").textContent = "⚠️ " + w;
    store.open().then(refreshLoaded).catch((e) => U.log("app", "IndexedDB: " + e.message, "warn"));
  }

  document.addEventListener("DOMContentLoaded", init);
})(typeof window !== "undefined" ? window : this);
