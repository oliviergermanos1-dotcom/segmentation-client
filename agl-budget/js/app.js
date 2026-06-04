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

  // Messages restants pour onglets non branchés (aucun pour l'instant).
  const STUB_MSG = {};

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
      // Crée les panneaux stub manquants (tous sauf ceux déjà dans le HTML).
      if (!$('[data-panel="' + t.id + '"]')) {
        const panel = el("section", { class: "panel", "data-panel": t.id });
        const stubMsg = STUB_MSG[t.id] ||
          "Onglet prévu au planning des sprints (cf. README §8). Non encore implémenté.";
        panel.appendChild(el("div", { class: "section" }, [
          el("h2", { text: t.label.replace(/^\d+ · /, "") }),
          el("div", { class: "placeholder", html: stubMsg })
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
    else if (id === "05") renderPipelineCRMTab();
    else if (id === "06") renderBCGTab();
    else if (id === "07") renderWhiteSpacesTab();
    else if (id === "08") renderPNDMapTab();
    else if (id === "09") renderPNDTopTab();
    else if (id === "10") renderProjectionsTab();
    else if (id === "11") renderPreconisationsTab();
    else if (id === "12") renderExecutiveTab();
    else if (id === "13") renderConcurrentielleTab();
    else if (id === "14") renderCohortesTab();
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
      const allSorted = all.slice().sort((a, b) => (a.METIER || "").localeCompare(b.METIER || "") || (+a.ANNEE) - (+b.ANNEE));
      const tbl = el("table");
      tbl.innerHTML = "<thead><tr><th>Métier</th><th>Année</th><th>Vol. TEU</th><th>Vol. Kg</th><th>Nb clients</th><th>Nb ops</th></tr></thead>";
      const tb = el("tbody");
      allSorted.forEach((r) => {
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

      renderToolbar("02", "Marché STATCOM " + year, "Marche_STATCOM_" + year,
        () => global.AGL.prompts.market(sorted, year),
        () => [
          _tableSection("Marché total " + year + " — mix par métier",
            ["Métier", "Volume principal"],
            sorted.map((s) => [s.label, U.formatNumber(s.value)])),
          _tableSection("Marché total 3 ans (détail)",
            ["Métier", "Année", "Vol. TEU", "Vol. Kg", "Nb clients", "Nb ops"],
            allSorted.map((r) => [r.METIER, r.ANNEE, U.formatNumber(r.VOLUME_TEU),
              U.formatNumber(r.VOLUME_KG), U.formatNumber(r.NB_CLIENTS), U.formatNumber(r.NB_OPERATIONS)]))
        ]);
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
      const sorted = rows.slice().sort((a, b) => b.pdm - a.pdm);
      global.AGL.charts.bar($("#m03-chart-pdm"), sorted.map((r) => ({ label: r.metier, value: r.pdm * 100 })),
        { title: "PDM AGL " + year + " (%)" });
      const tbl = el("table");
      tbl.innerHTML = "<thead><tr><th>Métier</th><th>Vol. AGL</th><th>Vol. Marché</th><th>PDM</th><th>CAGR AGL</th><th>CAGR Marché</th></tr></thead>";
      const tb = el("tbody");
      sorted.forEach((r) => {
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

      renderToolbar("03", "Position AGL " + year, "Position_AGL_" + year,
        () => global.AGL.prompts.position(sorted, year),
        () => [
          _tableSection("PDM AGL " + year + " par métier",
            ["Métier", "Vol. AGL", "Vol. Marché", "PDM", "CAGR AGL", "CAGR Marché"],
            sorted.map((r) => [r.metier, U.formatNumber(r.volume_agl), U.formatNumber(r.volume_marche),
              (r.pdm * 100).toFixed(1) + " %",
              r.cagr_agl != null ? (r.cagr_agl * 100).toFixed(1) + " %" : "—",
              r.cagr_marche != null ? (r.cagr_marche * 100).toFixed(1) + " %" : "—"]))
        ]);
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

  /* ====================================================================== */
  /*  Onglet 07 — White Spaces                                                */
  /* ====================================================================== */
  function renderWhiteSpacesTab() {
    global.AGL.pipeline.get("WHITESPACES").then((rows) => {
      if (!rows || !rows.length) { $("#m07-placeholder").style.display = ""; $("#m07-content").style.display = "none"; return; }
      $("#m07-placeholder").style.display = "none";
      $("#m07-content").style.display = "";
      const sorted = rows.slice().sort((a, b) => U.toNumber(b.VOLUME_PRINCIPAL) - U.toNumber(a.VOLUME_PRINCIPAL));
      const top10 = sorted.slice(0, 10).map((r) => ({ label: (r.NOM_BASE || "?").slice(0, 18), value: U.toNumber(r.VOLUME_PRINCIPAL) }));
      global.AGL.charts.bar($("#m07-chart"), top10, { title: "Top 10 white spaces — volume non capturé" });
      const tbl = el("table");
      tbl.innerHTML = "<thead><tr><th>Client</th><th>Métier</th><th>Statut</th><th>Volume</th><th>Nb ops</th><th>Années</th></tr></thead>";
      const tb = el("tbody");
      sorted.slice(0, 50).forEach((r) => {
        const tr = el("tr");
        const cls = r.STATUT === "MATCH_AMBIGU" ? "warning" : "";
        tr.innerHTML = "<td>" + (r.NOM_BASE || "") + "</td><td>" + (r.METIER || "") +
          "</td><td><span class='badge'>" + (r.STATUT || "") + "</span></td>" +
          "<td>" + U.formatNumber(r.VOLUME_PRINCIPAL) + "</td>" +
          "<td>" + U.formatNumber(r.NB_OPERATIONS) + "</td>" +
          "<td>" + U.formatNumber(r.ANNEES_PRESENT) + "</td>";
        tb.appendChild(tr);
      });
      tbl.appendChild(tb);
      const host = $("#m07-table"); host.innerHTML = ""; host.appendChild(tbl);

      renderToolbar("07", "White Spaces", "White_Spaces",
        () => global.AGL.prompts.whitespaces(sorted),
        () => [
          _tableSection("Top 50 white spaces",
            ["Client", "Métier", "Statut", "Volume", "Nb ops"],
            sorted.slice(0, 50).map((r) => [r.NOM_BASE, r.METIER, r.STATUT,
              U.formatNumber(r.VOLUME_PRINCIPAL), U.formatNumber(r.NB_OPERATIONS)]))
        ]);
    });
  }

  /* ====================================================================== */
  /*  Onglet 10 — Projections 2026-2030                                       */
  /* ====================================================================== */
  function renderProjectionsTab() {
    global.AGL.pipeline.get("PROJECTIONS").then((rows) => {
      if (!rows || !rows.length) { $("#m10-placeholder").style.display = ""; $("#m10-content").style.display = "none"; return; }
      $("#m10-placeholder").style.display = "none";
      $("#m10-content").style.display = "";
      const metiers = Array.from(new Set(rows.map((r) => r.METIER))).sort();
      const sel = $("#m10-metier");
      const current = sel.value;
      sel.innerHTML = "";
      metiers.forEach((m) => sel.appendChild(el("option", { value: m, text: m })));
      sel.value = current && metiers.indexOf(current) > -1 ? current : metiers[0];
      sel.onchange = () => drawProjections(rows, sel.value);
      drawProjections(rows, sel.value);
    });
  }
  function drawProjections(rows, metier) {
    const sub = rows.filter((r) => r.METIER === metier);
    const scenarios = ["REEL", "PESSIMISTE", "BASE", "OPTIMISTE"];
    const series = scenarios.map((s) => ({
      label: s,
      points: sub.filter((r) => r.SCENARIO === s)
                 .map((r) => ({ x: +r.ANNEE, y: U.toNumber(r.VOLUME) }))
                 .sort((a, b) => a.x - b.x)
    })).filter((s) => s.points.length);
    // Pour relier le réel à chaque scénario : on injecte le dernier point réel
    // au début de chaque trajectoire scénario.
    const reel = series.find((s) => s.label === "REEL");
    if (reel && reel.points.length) {
      const anchor = reel.points[reel.points.length - 1];
      ["PESSIMISTE", "BASE", "OPTIMISTE"].forEach((sc) => {
        const s = series.find((x) => x.label === sc);
        if (s) s.points.unshift(anchor);
      });
    }
    global.AGL.charts.line($("#m10-chart"), series, { title: metier + " — projections 3 scénarios" });
    // Table
    const tbl = el("table");
    tbl.innerHTML = "<thead><tr><th>Scénario</th><th>Probabilité</th><th>Année</th><th>Volume</th><th>CAGR appliqué</th></tr></thead>";
    const tb = el("tbody");
    sub.sort((a, b) => (a.SCENARIO || "").localeCompare(b.SCENARIO || "") || (+a.ANNEE) - (+b.ANNEE))
       .forEach((r) => {
      const tr = el("tr");
      const prob = U.toNumber(r.PROBABILITE);
      const cagr = r.CAGR_APPLIQUE !== "" && r.CAGR_APPLIQUE != null ? U.toNumber(r.CAGR_APPLIQUE) : null;
      tr.innerHTML = "<td><strong>" + r.SCENARIO + "</strong></td>" +
        "<td>" + (prob * 100).toFixed(0) + " %</td>" +
        "<td>" + r.ANNEE + "</td>" +
        "<td>" + U.formatNumber(U.toNumber(r.VOLUME)) + "</td>" +
        "<td>" + (cagr != null ? (cagr * 100).toFixed(1) + " %" : "—") + "</td>";
      tb.appendChild(tr);
    });
    tbl.appendChild(tb);
    const host = $("#m10-table"); host.innerHTML = ""; host.appendChild(tbl);

    renderToolbar("10", "Projections " + metier, "Projections_" + _slug(metier),
      () => global.AGL.prompts.projections(sub, metier),
      () => [
        _tableSection(metier + " — projections 3 scénarios",
          ["Scénario", "Probabilité", "Année", "Volume", "CAGR appliqué"],
          sub.map((r) => [r.SCENARIO,
            (U.toNumber(r.PROBABILITE) * 100).toFixed(0) + " %",
            r.ANNEE, U.formatNumber(U.toNumber(r.VOLUME)),
            r.CAGR_APPLIQUE !== "" && r.CAGR_APPLIQUE != null ? (U.toNumber(r.CAGR_APPLIQUE) * 100).toFixed(1) + " %" : "—"]))
      ]);
  }

  /* ====================================================================== */
  /*  Onglet 13 ★ Analyse Concurrentielle (bubble chart)                      */
  /* ====================================================================== */
  function renderConcurrentielleTab() {
    global.AGL.pipeline.get("PDM").then((rows) => {
      if (!rows || !rows.length) { $("#m13-placeholder").style.display = ""; $("#m13-content").style.display = "none"; return; }
      $("#m13-placeholder").style.display = "none";
      $("#m13-content").style.display = "";
      const years = Array.from(new Set(rows.map((r) => +r.ANNEE))).sort();
      fillYearSelect("#m13-year", years);
      const sel = $("#m13-year");
      sel.onchange = () => drawConcurrentielle(rows, +sel.value);
      drawConcurrentielle(rows, +sel.value);
    });
  }
  function drawConcurrentielle(rows, year) {
    const sub = rows.filter((r) => +r.ANNEE === year);
    const points = sub.map((r) => ({
      x: U.toNumber(r.PDM) * 100,
      y: r.CAGR_MARCHE != null && r.CAGR_MARCHE !== "" ? U.toNumber(r.CAGR_MARCHE) * 100 : 0,
      r: U.toNumber(r.VOLUME_MARCHE),
      label: r.METIER
    }));
    global.AGL.charts.bubble($("#m13-chart"), points, { xLabel: "PDM AGL (%)", yLabel: "CAGR marché (%)" });

    const pdmShape = sub.map((r) => ({
      metier: r.METIER, pdm: U.toNumber(r.PDM),
      volume_agl: U.toNumber(r.VOLUME_AGL), volume_marche: U.toNumber(r.VOLUME_MARCHE),
      cagr_agl: r.CAGR_AGL != null && r.CAGR_AGL !== "" ? U.toNumber(r.CAGR_AGL) : null,
      cagr_marche: r.CAGR_MARCHE != null && r.CAGR_MARCHE !== "" ? U.toNumber(r.CAGR_MARCHE) : null
    }));
    renderToolbar("13", "Analyse Concurrentielle " + year, "Concurrentielle_" + year,
      () => global.AGL.prompts.position(pdmShape, year),
      () => [
        _tableSection("Positionnement AGL " + year + " — par métier",
          ["Métier", "Vol. AGL", "Vol. Marché", "PDM", "CAGR Marché"],
          pdmShape.map((r) => [r.metier, U.formatNumber(r.volume_agl), U.formatNumber(r.volume_marche),
            (r.pdm * 100).toFixed(1) + " %",
            r.cagr_marche != null ? (r.cagr_marche * 100).toFixed(1) + " %" : "—"]))
      ]);
  }

  /* ====================================================================== */
  /*  Onglet 06 — Matrice BCG                                                 */
  /* ====================================================================== */
  function renderBCGTab() {
    global.AGL.analyses.bcg().then((cells) => {
      if (!cells || !cells.length) { $("#m06-placeholder").style.display = ""; $("#m06-content").style.display = "none"; return; }
      $("#m06-placeholder").style.display = "none";
      $("#m06-content").style.display = "";
      const points = cells.slice(0, 50).map((c) => ({
        x: c.pdm * 100,
        y: (c.cagr == null ? 0 : c.cagr * 100),
        r: c.volume,
        label: (c.client || "?").slice(0, 12)
      }));
      global.AGL.charts.bubble($("#m06-chart"), points, { xLabel: "PDM client (%)", yLabel: "CAGR client (%)" });
      const tbl = el("table");
      tbl.innerHTML = "<thead><tr><th>Client</th><th>Métier</th><th>Volume</th><th>PDM</th><th>CAGR</th><th>Cellule BCG</th></tr></thead>";
      const tb = el("tbody");
      cells.slice(0, 30).forEach((c) => {
        const tr = el("tr");
        tr.innerHTML = "<td>" + (c.client || "") + "</td><td>" + (c.metier || "") + "</td>" +
          "<td>" + U.formatNumber(c.volume) + "</td>" +
          "<td>" + (c.pdm * 100).toFixed(1) + " %</td>" +
          "<td>" + (c.cagr == null ? "—" : (c.cagr * 100).toFixed(1) + " %") + "</td>" +
          "<td><span class='badge'>" + (c.cellule || "?") + "</span></td>";
        tb.appendChild(tr);
      });
      tbl.appendChild(tb);
      const host = $("#m06-table"); host.innerHTML = ""; host.appendChild(tbl);
    });
  }

  /* ====================================================================== */
  /*  Onglet 11 — Préconisations Client                                       */
  /* ====================================================================== */
  let _precoCache = null;
  function renderPreconisationsTab() {
    global.AGL.analyses.preconisations().then((rows) => {
      if (!rows || !rows.length) { $("#m11-placeholder").style.display = ""; $("#m11-content").style.display = "none"; return; }
      $("#m11-placeholder").style.display = "none";
      $("#m11-content").style.display = "";
      _precoCache = rows;
      const search = $("#m11-search"), sel = $("#m11-client");
      function refresh() {
        const q = (search.value || "").toUpperCase();
        const filtered = rows.filter((r) => !q || (r.nom || "").toUpperCase().indexOf(q) > -1);
        sel.innerHTML = "";
        filtered.slice(0, 200).forEach((r) =>
          sel.appendChild(el("option", { value: r.id_rmc, text: r.nom + " · " + (r.secteur || "?") })));
        if (filtered.length) drawFiche(filtered[0]);
      }
      search.oninput = U.debounce(refresh, 150);
      sel.onchange = () => {
        const cur = _precoCache.find((r) => r.id_rmc === sel.value);
        if (cur) drawFiche(cur);
      };
      refresh();
    });
  }
  function drawFiche(c) {
    const host = $("#m11-fiche");
    host.innerHTML =
      "<div style='display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:10px'>" +
      _ficheCard("Nom canonique", c.nom) +
      _ficheCard("Secteur", c.secteur + (c.locked ? " · 🔒 LOCKED" : "") + " (" + (c.source_secteur || "?") + ")") +
      _ficheCard("ID CRM", c.id_crm || "—") +
      _ficheCard("ID IRIS", (c.id_iris || "—") + (c.iris_score ? " · score " + (+c.iris_score).toFixed(2) : "")) +
      _ficheCard("ID STATCOM", (c.id_statcom || "—") + (c.statcom_score ? " · score " + (+c.statcom_score).toFixed(2) : "")) +
      _ficheCard("Volume total 3 ans", U.formatNumber(c.volume_total)) +
      _ficheCard("Métier principal", c.metier_principal) +
      _ficheCard("PDM (métier principal)", c.pdm_principal != null ? (c.pdm_principal * 100).toFixed(1) + " %" : "—") +
      _ficheCard("CAGR client", c.cagr_principal != null ? (c.cagr_principal * 100).toFixed(1) + " %" : "—") +
      _ficheCard("Position BCG", c.bcg) +
      "</div>" +
      "<div class='section' style='margin-top:12px;background:#f8fafc;border-color:var(--primary)'>" +
      "<strong>Recommandation :</strong> " + c.recommandation + "</div>";
    renderToolbar("11", "Fiche client · " + (c.nom || ""), "Fiche_" + _slug(c.nom),
      () => global.AGL.prompts.clientFiche(c),
      () => _ficheSections(c));
  }
  function _ficheCard(label, value) {
    return "<div style='border:1px solid var(--border);border-radius:6px;padding:8px'>" +
           "<div class='muted' style='font-size:11px;text-transform:uppercase'>" + label + "</div>" +
           "<div><strong>" + (value == null ? "—" : value) + "</strong></div></div>";
  }
  function _ficheSections(c) {
    return [
      { heading: "Identité",
        table: { headers: ["Champ", "Valeur"], rows: [
          ["Nom canonique", c.nom], ["Secteur", c.secteur + (c.locked ? " 🔒" : "")],
          ["Source secteur", c.source_secteur || "—"],
          ["ID CRM", c.id_crm || "—"], ["ID IRIS", c.id_iris || "—"], ["ID STATCOM", c.id_statcom || "—"]
        ]}
      },
      { heading: "Position commerciale",
        table: { headers: ["Indicateur", "Valeur"], rows: [
          ["Volume total 3 ans", U.formatNumber(c.volume_total)],
          ["Métier principal", c.metier_principal],
          ["PDM (métier principal)", c.pdm_principal != null ? (c.pdm_principal * 100).toFixed(1) + " %" : "—"],
          ["CAGR client", c.cagr_principal != null ? (c.cagr_principal * 100).toFixed(1) + " %" : "—"],
          ["Position BCG", c.bcg]
        ]}
      },
      { heading: "Recommandation stratégique", text: c.recommandation }
    ];
  }
  function _slug(s) { return (s || "client").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^A-Za-z0-9]+/g, "_").slice(0, 40); }
  function _bindAct(id, fn) {
    U.$$('[data-act="' + id + '"]').forEach((b) => b.onclick = (e) => { e.preventDefault(); fn(); });
  }

  /* ----- Toolbar standard (Claude prompt / Word / PPT / HTML) -----
   * Place 4 boutons dans <div class="toolbar" data-toolbar="XX">.
   * promptFn() → string ; sectionsFn() → [{heading,text/bullets/table}].
   * --------------------------------------------------------------------- */
  function renderToolbar(panelId, title, fileBase, promptFn, sectionsFn) {
    const host = U.$('[data-toolbar="' + panelId + '"]');
    if (!host) return;
    host.innerHTML = "";
    const mk = (txt, fn) => { const b = el("button", { text: txt }); b.onclick = (e) => { e.preventDefault(); fn(); }; return b; };
    if (promptFn) host.appendChild(mk("📋 Copier prompt Claude", () => {
      Promise.resolve(promptFn()).then((txt) =>
        global.AGL.prompts.copyToClipboard(txt).then(() => _toast("Prompt Claude copié")));
    }));
    host.appendChild(el("span", { class: "sep" }));
    host.appendChild(mk("📄 Word", () => {
      Promise.resolve(sectionsFn()).then((s) => global.AGL.exports.exportWord(fileBase + ".docx", title, s));
    }));
    host.appendChild(mk("📊 PPT", () => {
      Promise.resolve(sectionsFn()).then((s) => global.AGL.exports.exportPPT(fileBase + ".pptx", title, s));
    }));
    host.appendChild(mk("🖨️ HTML", () => {
      Promise.resolve(sectionsFn()).then((s) => global.AGL.exports.exportHTML(fileBase + ".html", title, s));
    }));
  }
  function _toast(msg) {
    const t = el("div", { text: msg, style:
      "position:fixed;bottom:20px;right:20px;background:#16a34a;color:#fff;padding:10px 16px;border-radius:6px;z-index:9999;box-shadow:0 4px 12px rgba(0,0,0,.15)" });
    document.body.appendChild(t);
    setTimeout(() => t.remove(), 2200);
  }

  /* ----- Builders de sections génériques (tables) ----- */
  function _tableSection(heading, headers, rows, text) {
    const sec = { heading: heading };
    if (text) sec.text = text;
    sec.table = { headers: headers, rows: rows };
    return sec;
  }

  /* ====================================================================== */
  /*  Onglet 12 — Synthèse Executive                                          */
  /* ====================================================================== */
  function renderExecutiveTab() {
    global.AGL.analyses.executive().then((snap) => {
      if (!snap || (snap.marche_total == null && snap.pdm_globale == null && !snap.clients_rmc)) {
        $("#m12-placeholder").style.display = "";
        $("#m12-content").style.display = "none";
        return;
      }
      $("#m12-placeholder").style.display = "none";
      $("#m12-content").style.display = "";
      const host = $("#m12-snap");
      host.innerHTML = "";
      const kpi = (l, v) => "<div style='border:1px solid var(--border);border-radius:6px;padding:10px'>" +
                            "<div class='muted' style='font-size:11px;text-transform:uppercase'>" + l +
                            "</div><div style='font-size:18px'><strong>" + v + "</strong></div></div>";
      let html = "<div style='display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px;margin-bottom:12px'>";
      if (snap.marche_total != null) html += kpi("Marché total " + (snap.derniere_annee || ""), U.formatNumber(snap.marche_total));
      if (snap.pdm_globale != null)  html += kpi("PDM AGL globale", (snap.pdm_globale * 100).toFixed(1) + " %");
      if (snap.clients_rmc != null)  html += kpi("Clients RMC", U.formatNumber(snap.clients_rmc));
      if (snap.horizon)              html += kpi("Horizon projections", snap.horizon);
      html += "</div>";
      if (snap.top_croissance && snap.top_croissance.length) {
        html += "<h3>Top métiers en croissance</h3><table><thead><tr><th>Métier</th><th>CAGR Marché</th></tr></thead><tbody>" +
          snap.top_croissance.map((m) => "<tr><td>" + m.metier + "</td><td><strong>" + (m.cagr * 100).toFixed(1) + " %</strong></td></tr>").join("") +
          "</tbody></table>";
      }
      if (snap.top_white_spaces && snap.top_white_spaces.length) {
        html += "<h3>Top white spaces</h3><table><thead><tr><th>Client</th><th>Métier</th><th>Volume</th></tr></thead><tbody>" +
          snap.top_white_spaces.map((w) => "<tr><td>" + w.client + "</td><td>" + w.metier + "</td><td>" + U.formatNumber(w.volume) + "</td></tr>").join("") +
          "</tbody></table>";
      }
      if (snap.risques_opportunites) {
        html += "<h3>Risques & opportunités macro (juin 2026)</h3><ul>";
        snap.risques_opportunites.slice(0, 6).forEach((x) =>
          html += "<li><strong>[" + x.type + "] " + x.titre + "</strong> — " + x.detail + "</li>");
        html += "</ul>";
      }
      host.innerHTML = html;

      renderToolbar("12", "Note COMEX — Synthèse Executive AGL CIV " + (snap.derniere_annee || ""), "Note_COMEX",
        () => global.AGL.prompts.executive(snap),
        () => global.AGL.exports.buildExecutiveSections(snap));
    });
  }

  /* ====================================================================== */
  /*  Onglet 14 — Cohortes & Rétention                                        */
  /* ====================================================================== */
  function renderCohortesTab() {
    global.AGL.analyses.cohortes().then((res) => {
      if (!res || !res.retention.length) { $("#m14-placeholder").style.display = ""; $("#m14-content").style.display = "none"; return; }
      $("#m14-placeholder").style.display = "none";
      $("#m14-content").style.display = "";

      // Rétention.
      const years = res.years;
      const tbl = el("table");
      tbl.innerHTML = "<thead><tr><th>Cohorte</th><th>Taille</th>" +
        years.map((y) => "<th>" + y + "</th>").join("") + "</tr></thead>";
      const tb = el("tbody");
      res.retention.forEach((r) => {
        const tr = el("tr");
        let row = "<td>" + r.cohorte + "</td><td>" + r.taille_initiale + "</td>";
        years.forEach((y) => {
          const v = r["r" + y];
          row += "<td>" + (v == null ? "" : (v * 100).toFixed(0) + " %") + "</td>";
        });
        tr.innerHTML = row;
        tb.appendChild(tr);
      });
      tbl.appendChild(tb);
      $("#m14-retention").innerHTML = ""; $("#m14-retention").appendChild(tbl);

      // Churn.
      const tbl2 = el("table");
      tbl2.innerHTML = "<thead><tr><th>Année</th><th>Base N-1</th><th>Base N</th><th>Perdus</th><th>Gagnés</th><th>Churn</th></tr></thead>";
      const tb2 = el("tbody");
      res.churn.forEach((r) => {
        const tr = el("tr");
        tr.innerHTML = "<td>" + r.annee + "</td><td>" + r.base_n1 + "</td><td>" + r.base_n + "</td>" +
          "<td>" + r.perdus + "</td><td>" + r.gagnes + "</td>" +
          "<td><strong>" + (r.churn * 100).toFixed(1) + " %</strong></td>";
        tb2.appendChild(tr);
      });
      tbl2.appendChild(tb2);
      $("#m14-churn").innerHTML = ""; $("#m14-churn").appendChild(tbl2);

      // LTV top 20.
      const tbl3 = el("table");
      tbl3.innerHTML = "<thead><tr><th>#</th><th>Client</th><th>LTV volume</th><th>Années actives</th></tr></thead>";
      const tb3 = el("tbody");
      res.ltv.slice(0, 20).forEach((r, i) => {
        const tr = el("tr");
        tr.innerHTML = "<td>" + (i + 1) + "</td><td>" + r.client + "</td>" +
          "<td>" + U.formatNumber(r.ltv_volume) + "</td><td>" + r.annees_actives + "</td>";
        tb3.appendChild(tr);
      });
      tbl3.appendChild(tb3);
      $("#m14-ltv").innerHTML = ""; $("#m14-ltv").appendChild(tbl3);

      const retSections = [
        _tableSection("Rétention par cohorte",
          ["Cohorte", "Taille"].concat(years.map(String)),
          res.retention.map((r) => [r.cohorte, r.taille_initiale]
            .concat(years.map((y) => r["r" + y] == null ? "" : (r["r" + y] * 100).toFixed(0) + " %")))),
        _tableSection("Churn année par année",
          ["Année", "Base N-1", "Base N", "Perdus", "Gagnés", "Churn"],
          res.churn.map((r) => [r.annee, r.base_n1, r.base_n, r.perdus, r.gagnes, (r.churn * 100).toFixed(1) + " %"])),
        _tableSection("Top 20 LTV (volume cumulé)",
          ["#", "Client", "LTV volume", "Années actives"],
          res.ltv.slice(0, 20).map((r, i) => [i + 1, r.client, U.formatNumber(r.ltv_volume), r.annees_actives]))
      ];
      renderToolbar("14", "Cohortes & Rétention", "Cohortes_Retention",
        null, () => retSections);
    });
  }

  /* ====================================================================== */
  /*  Onglet 05 — Pipeline CRM (opportunites.xlsx)                            */
  /* ====================================================================== */
  function renderPipelineCRMTab() {
    global.AGL.pipeline.get("PIPELINE_CRM").then((rows) => {
      if (!rows || !rows.length) { $("#m05-placeholder").style.display = ""; $("#m05-content").style.display = "none"; return; }
      $("#m05-placeholder").style.display = "none";
      $("#m05-content").style.display = "";
      const C = {
        cap:   "CAP Potentiel Global (€)",
        proba: "Probabilité de succès",
        stat:  "Statut",
        sst:   "Sous-statut",
        type:  "Type",
        vert:  "Verticale (Compte) (Compte)",
        compte:"Compte",
        nom:   "Nom de l'opportunité",
        proprio: "Propriétaire",
      };
      const num = (v) => U.toNumber(v);
      const pct = (v) => {
        if (v == null) return 0;
        const s = String(v).replace("%", "").replace(",", ".").trim();
        const n = parseFloat(s);
        return isNaN(n) ? 0 : (n > 1 ? n / 100 : n);
      };
      const capTotal = rows.reduce((a, r) => a + num(r[C.cap]), 0);
      const capPond  = rows.reduce((a, r) => a + num(r[C.cap]) * pct(r[C.proba]), 0);
      const nbOuvert = rows.filter((r) => /ouvert/i.test(r[C.stat] || "")).length;
      const nbGagne  = rows.filter((r) => /gagn/i.test(r[C.sst] || "")).length;
      const nbPerdu  = rows.filter((r) => /perdu|annul/i.test((r[C.sst] || "") + (r[C.stat] || ""))).length;
      const winRate  = (nbGagne + nbPerdu) ? nbGagne / (nbGagne + nbPerdu) : 0;

      const kpi = (l, v) => "<div style='border:1px solid var(--border);border-radius:6px;padding:10px'>" +
        "<div class='muted' style='font-size:11px;text-transform:uppercase'>" + l + "</div>" +
        "<div style='font-size:18px'><strong>" + v + "</strong></div></div>";
      $("#m05-kpis").innerHTML =
        kpi("Opportunités", U.formatNumber(rows.length)) +
        kpi("CAP potentiel total", new Intl.NumberFormat("fr-FR", { maximumFractionDigits: 0 }).format(capTotal) + " €") +
        kpi("CAP pondéré (× proba)", new Intl.NumberFormat("fr-FR", { maximumFractionDigits: 0 }).format(capPond) + " €") +
        kpi("Ouvertes", U.formatNumber(nbOuvert)) +
        kpi("Win rate (clôturées)", (winRate * 100).toFixed(1) + " %");

      // Funnel par sous-statut
      const funnel = {};
      rows.forEach((r) => { const k = r[C.sst] || "?"; funnel[k] = (funnel[k] || 0) + 1; });
      const funnelData = Object.keys(funnel).map((k) => ({ label: k, value: funnel[k] }))
                          .sort((a, b) => b.value - a.value);
      global.AGL.charts.bar($("#m05-funnel"), funnelData, { title: "Nb opportunités par sous-statut" });

      // Verticale
      const byVert = {};
      rows.forEach((r) => { const k = r[C.vert] || "?"; byVert[k] = (byVert[k] || 0) + num(r[C.cap]); });
      const vertData = Object.keys(byVert).map((k) => ({ label: k.slice(0, 14), value: byVert[k] }))
                        .sort((a, b) => b.value - a.value).slice(0, 8);
      global.AGL.charts.bar($("#m05-vert"), vertData, { title: "CAP potentiel par verticale (top 8) — €" });

      // Top 20 opportunités
      const top = rows.slice().sort((a, b) => num(b[C.cap]) - num(a[C.cap])).slice(0, 20);
      const tbl = el("table");
      tbl.innerHTML = "<thead><tr><th>Compte</th><th>Verticale</th><th>Opportunité</th>" +
        "<th>CAP €</th><th>Proba</th><th>Sous-statut</th><th>Type</th><th>Propriétaire</th></tr></thead>";
      const tb = el("tbody");
      top.forEach((r) => {
        const tr = el("tr");
        tr.innerHTML = "<td>" + (r[C.compte] || "") + "</td><td>" + (r[C.vert] || "") + "</td>" +
          "<td>" + (r[C.nom] || "") + "</td>" +
          "<td>" + U.formatNumber(num(r[C.cap])) + "</td>" +
          "<td>" + (pct(r[C.proba]) * 100).toFixed(0) + " %</td>" +
          "<td><span class='badge'>" + (r[C.sst] || "") + "</span></td>" +
          "<td>" + (r[C.type] || "") + "</td>" +
          "<td>" + (r[C.proprio] || "") + "</td>";
        tb.appendChild(tr);
      });
      tbl.appendChild(tb);
      $("#m05-top").innerHTML = ""; $("#m05-top").appendChild(tbl);

      renderToolbar("05", "Pipeline CRM — analyse", "Pipeline_CRM", null, () => [
        _tableSection("KPIs pipeline",
          ["Indicateur", "Valeur"],
          [["Opportunités", U.formatNumber(rows.length)],
           ["CAP potentiel total (€)", U.formatNumber(capTotal)],
           ["CAP pondéré (× proba) (€)", U.formatNumber(capPond)],
           ["Win rate (clôturées)", (winRate * 100).toFixed(1) + " %"]]),
        _tableSection("Funnel par sous-statut",
          ["Sous-statut", "Nb"], funnelData.map((f) => [f.label, f.value])),
        _tableSection("Mix par verticale — CAP €",
          ["Verticale", "CAP €"], vertData.map((v) => [v.label, U.formatNumber(v.value)])),
        _tableSection("Top 20 opportunités",
          ["Compte", "Verticale", "Opportunité", "CAP €", "Proba", "Sous-statut"],
          top.map((r) => [r[C.compte], r[C.vert], r[C.nom], U.formatNumber(num(r[C.cap])),
            (pct(r[C.proba]) * 100).toFixed(0) + " %", r[C.sst]]))
      ]);
    });
  }

  /* ====================================================================== */
  /*  Onglet 08 — PND vue par axe & sectoriel (faute de GPS, treemap/donut)   */
  /* ====================================================================== */
  function renderPNDMapTab() {
    Promise.all([global.AGL.pipeline.get("PND_AXES"), global.AGL.pipeline.get("PND_SECTORIELS")])
      .then(([axes, sect]) => {
      if (!axes) { $("#m08-placeholder").style.display = ""; $("#m08-content").style.display = "none"; return; }
      $("#m08-placeholder").style.display = "none";
      $("#m08-content").style.display = "";

      const data = axes.map((r) => ({ label: "Axe " + r.axe, value: U.toNumber(r.cout_total) }));
      global.AGL.charts.donut($("#m08-donut"), data, { center: "PND" });

      const tblA = el("table");
      tblA.innerHTML = "<thead><tr><th>Axe</th><th>Libellé</th><th>Nb actions</th><th>Coût 5 ans (M FCFA)</th><th>% total</th><th>Actions taggées AGL</th></tr></thead>";
      const total = axes.reduce((a, r) => a + U.toNumber(r.cout_total), 0) || 1;
      const tbA = el("tbody");
      axes.sort((a, b) => U.toNumber(b.cout_total) - U.toNumber(a.cout_total)).forEach((r) => {
        const tr = el("tr");
        tr.innerHTML = "<td><strong>" + r.axe + "</strong></td><td>" + r.axe_libelle + "</td>" +
          "<td>" + U.formatNumber(r.nb_actions) + "</td>" +
          "<td>" + U.formatNumber(U.toNumber(r.cout_total)) + "</td>" +
          "<td>" + ((U.toNumber(r.cout_total) / total) * 100).toFixed(1) + " %</td>" +
          "<td>" + U.formatNumber(r.nb_agl) + "</td>";
        tbA.appendChild(tr);
      });
      tblA.appendChild(tbA);
      $("#m08-axes").innerHTML = ""; $("#m08-axes").appendChild(tblA);

      if (sect && sect.length) {
        const tblS = el("table");
        tblS.innerHTML = "<thead><tr><th>Axe</th><th>Résultat sectoriel</th><th>Nb actions</th><th>Coût 5 ans</th><th>AGL</th></tr></thead>";
        const tbS = el("tbody");
        sect.slice().sort((a, b) => U.toNumber(b.cout_total) - U.toNumber(a.cout_total)).slice(0, 30).forEach((r) => {
          const tr = el("tr");
          tr.innerHTML = "<td>" + r.axe + "</td><td>" + r.code_sectoriel + "</td>" +
            "<td>" + U.formatNumber(r.nb_actions) + "</td>" +
            "<td>" + U.formatNumber(U.toNumber(r.cout_total)) + "</td>" +
            "<td>" + U.formatNumber(r.nb_agl) + "</td>";
          tbS.appendChild(tr);
        });
        tblS.appendChild(tbS);
        $("#m08-sect").innerHTML = ""; $("#m08-sect").appendChild(tblS);
      }

      renderToolbar("08", "PND 2026-2030 — vue d'ensemble", "PND_Vue_Ensemble", null, () => [
        _tableSection("Coût PND par axe (M FCFA, 5 ans)",
          ["Axe", "Libellé", "Nb actions", "Coût 5 ans", "% total", "AGL"],
          axes.map((r) => [r.axe, r.axe_libelle, U.formatNumber(r.nb_actions),
            U.formatNumber(U.toNumber(r.cout_total)),
            ((U.toNumber(r.cout_total) / total) * 100).toFixed(1) + " %",
            U.formatNumber(r.nb_agl)]))
      ]);
    });
  }

  /* ====================================================================== */
  /*  Onglet 09 — Top Projets PND (filtres axe + AGL + recherche)             */
  /* ====================================================================== */
  let _pndRowsCache = null;
  function renderPNDTopTab() {
    global.AGL.pipeline.get("PND_PROJETS").then((rows) => {
      if (!rows || !rows.length) { $("#m09-placeholder").style.display = ""; $("#m09-content").style.display = "none"; return; }
      $("#m09-placeholder").style.display = "none";
      $("#m09-content").style.display = "";
      _pndRowsCache = rows.filter((r) => r.niveau === "Action");
      const axes = Array.from(new Set(_pndRowsCache.map((r) => +r.axe))).sort();
      const sel = $("#m09-axe");
      sel.innerHTML = "<option value=''>Tous</option>";
      axes.forEach((a) => sel.appendChild(el("option", { value: a, text: "Axe " + a })));
      ["#m09-axe", "#m09-agl", "#m09-search", "#m09-top"].forEach((id) => {
        const e = $(id);
        const evt = e.tagName === "INPUT" && e.type === "text" ? "input" : "change";
        e.oninput = e.onchange = U.debounce(drawPND, 120);
      });
      drawPND();
    });
  }
  function drawPND() {
    if (!_pndRowsCache) return;
    const axe = $("#m09-axe").value;
    const aglOnly = $("#m09-agl").checked;
    const q = ($("#m09-search").value || "").toUpperCase();
    const topN = parseInt($("#m09-top").value, 10) || 50;
    let rows = _pndRowsCache.slice();
    if (axe) rows = rows.filter((r) => +r.axe === +axe);
    if (aglOnly) rows = rows.filter((r) => +r.pertinence_agl === 1);
    if (q) rows = rows.filter((r) => (r.intitule || "").toUpperCase().indexOf(q) > -1);
    rows.sort((a, b) => U.toNumber(b.cout_total) - U.toNumber(a.cout_total));
    rows = rows.slice(0, topN);

    const data = rows.slice(0, 15).map((r) => ({
      label: (r.intitule || r.code || "?").slice(0, 22),
      value: U.toNumber(r.cout_total)
    }));
    global.AGL.charts.bar($("#m09-chart"), data, { title: "Top 15 actions PND (coût 5 ans, M FCFA)" });

    const tbl = el("table");
    tbl.innerHTML = "<thead><tr><th>Axe</th><th>Code</th><th>Intitulé</th><th>Structure</th>" +
      "<th>2026</th><th>2027</th><th>2028</th><th>2029</th><th>2030</th><th>Total</th><th>AGL</th></tr></thead>";
    const tb = el("tbody");
    rows.forEach((r) => {
      const tr = el("tr");
      tr.innerHTML = "<td>" + r.axe + "</td><td>" + r.code + "</td>" +
        "<td>" + (r.intitule || "").slice(0, 80) + "</td>" +
        "<td>" + (r.structure_responsable || "") + "</td>" +
        "<td>" + U.formatNumber(U.toNumber(r.cout_2026)) + "</td>" +
        "<td>" + U.formatNumber(U.toNumber(r.cout_2027)) + "</td>" +
        "<td>" + U.formatNumber(U.toNumber(r.cout_2028)) + "</td>" +
        "<td>" + U.formatNumber(U.toNumber(r.cout_2029)) + "</td>" +
        "<td>" + U.formatNumber(U.toNumber(r.cout_2030)) + "</td>" +
        "<td><strong>" + U.formatNumber(U.toNumber(r.cout_total)) + "</strong></td>" +
        "<td>" + (+r.pertinence_agl ? "✓" : "") + "</td>";
      tb.appendChild(tr);
    });
    tbl.appendChild(tb);
    $("#m09-table").innerHTML = ""; $("#m09-table").appendChild(tbl);

    renderToolbar("09", "Top Projets PND" + (axe ? " — Axe " + axe : "") + (aglOnly ? " — pertinence AGL" : ""),
      "PND_Top_Projets",
      null,
      () => [
        _tableSection("Top " + rows.length + " actions PND" + (axe ? " — Axe " + axe : ""),
          ["Axe", "Code", "Intitulé", "Structure", "Coût total 5 ans (M FCFA)"],
          rows.map((r) => [r.axe, r.code, r.intitule, r.structure_responsable,
            U.formatNumber(U.toNumber(r.cout_total))]))
      ]);
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
