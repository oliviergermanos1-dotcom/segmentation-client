/* ============================================================================
 * AGL BUDGET — exports.js
 * Génère, à partir des analyses chargées, des livrables COMEX-ready :
 *   - PowerPoint (.pptx) via pptxgenjs (CDN)
 *   - Word (.docx) via docx (CDN)
 *   - Fallback : rapport HTML imprimable (Ctrl+P → PDF), 100 % offline-safe.
 *
 * Les libs CDN sont chargées paresseusement à la première demande.
 * Si le réseau est indisponible, l'utilisateur conserve le rapport HTML.
 *
 * Dépend de : AGL.utils, AGL.charts, AGL.pipeline, AGL.analyses, MACRO_DATA.
 * Expose window.AGL.exports.
 * ==========================================================================*/
(function (global) {
  "use strict";
  const U = global.AGL.utils;

  const CDN = {
    pptx: "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.min.js",
    docx: "https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js",
  };

  function _loadScript(src) {
    return new Promise((resolve, reject) => {
      if ([...document.scripts].some((s) => s.src === src)) return resolve();
      const s = document.createElement("script");
      s.src = src; s.onload = () => resolve(); s.onerror = () => reject(new Error("CDN " + src));
      document.head.appendChild(s);
    });
  }

  /* --------------------------------------------------------------------- */
  /*  HTML imprimable (fallback)                                            */
  /* --------------------------------------------------------------------- */
  function htmlReport(title, sections) {
    const css = `body{font:14px/1.5 -apple-system,Segoe UI,Roboto,sans-serif;color:#1f2937;max-width:900px;margin:24px auto;padding:0 16px}
      h1{color:#0b3d91;font-size:24px;margin:0 0 4px}h2{font-size:16px;border-bottom:2px solid #0b3d91;padding-bottom:4px;margin-top:24px}
      .sub{color:#6b7280;margin-bottom:16px;font-size:12px}table{border-collapse:collapse;width:100%;font-size:12px;margin:8px 0}
      th,td{border:1px solid #e3e6ea;padding:4px 8px;text-align:left}th{background:#f8fafc}
      .badge{background:#eef2ff;color:#2563eb;border-radius:999px;padding:2px 8px;font-size:11px}
      .conf{color:#dc2626;font-weight:600;letter-spacing:1px;font-size:11px}
      @media print{h2{page-break-after:avoid}table{page-break-inside:avoid}}`;
    let html = "<!doctype html><html><head><meta charset='utf-8'><title>" + title + "</title>" +
               "<style>" + css + "</style></head><body>" +
               "<div class='conf'>CONFIDENTIEL · AGL Direction Stratégie · " + new Date().toLocaleDateString("fr-FR") + "</div>" +
               "<h1>" + title + "</h1>" +
               "<div class='sub'>Sources macro à date juin 2026 — OMC, Drewry, IATA, AfCFTA, Banque Mondiale, FMI</div>";
    sections.forEach((sec) => {
      html += "<h2>" + sec.heading + "</h2>";
      if (sec.text) html += "<p>" + sec.text.replace(/\n/g, "<br>") + "</p>";
      if (sec.bullets) html += "<ul>" + sec.bullets.map((b) => "<li>" + b + "</li>").join("") + "</ul>";
      if (sec.table) {
        html += "<table><thead><tr>" + sec.table.headers.map((h) => "<th>" + h + "</th>").join("") + "</tr></thead><tbody>";
        sec.table.rows.forEach((r) => html += "<tr>" + r.map((c) => "<td>" + (c == null ? "" : c) + "</td>").join("") + "</tr>");
        html += "</tbody></table>";
      }
    });
    html += "</body></html>";
    return html;
  }

  function exportHTML(filename, title, sections) {
    const html = htmlReport(title, sections);
    U.downloadBlob(filename, new Blob([html], { type: "text/html;charset=utf-8" }));
  }

  /* --------------------------------------------------------------------- */
  /*  Word (.docx) via lib `docx` CDN                                       */
  /* --------------------------------------------------------------------- */
  function exportWord(filename, title, sections) {
    return _loadScript(CDN.docx).then(() => {
      if (!global.docx) throw new Error("Lib docx non disponible — fallback HTML");
      const D = global.docx;
      const children = [
        new D.Paragraph({ text: "CONFIDENTIEL · AGL Direction Stratégie",
                          alignment: D.AlignmentType.RIGHT,
                          spacing: { after: 100 } }),
        new D.Paragraph({ children: [new D.TextRun({ text: title, bold: true, size: 36, color: "0B3D91" })] }),
        new D.Paragraph({ children: [new D.TextRun({ text: "Sources macro à date juin 2026", italics: true, size: 18, color: "6B7280" })],
                          spacing: { after: 300 } }),
      ];
      sections.forEach((sec) => {
        children.push(new D.Paragraph({ children: [new D.TextRun({ text: sec.heading, bold: true, size: 26, color: "0B3D91" })],
                                        spacing: { before: 300, after: 100 } }));
        if (sec.text) children.push(new D.Paragraph({ text: sec.text }));
        if (sec.bullets) sec.bullets.forEach((b) =>
          children.push(new D.Paragraph({ text: b, bullet: { level: 0 } })));
        if (sec.table) {
          const headerRow = new D.TableRow({
            children: sec.table.headers.map((h) => new D.TableCell({
              children: [new D.Paragraph({ children: [new D.TextRun({ text: h, bold: true })] })],
              shading: { type: "clear", color: "auto", fill: "F8FAFC" }
            }))
          });
          const dataRows = sec.table.rows.map((r) => new D.TableRow({
            children: r.map((c) => new D.TableCell({
              children: [new D.Paragraph({ text: c == null ? "" : String(c) })]
            }))
          }));
          children.push(new D.Table({ rows: [headerRow, ...dataRows], width: { size: 100, type: D.WidthType.PERCENTAGE } }));
        }
      });
      const doc = new D.Document({ sections: [{ children: children }] });
      return D.Packer.toBlob(doc).then((blob) => U.downloadBlob(filename, blob));
    }).catch((err) => {
      U.log("exports", "Word KO (" + err.message + ") — fallback HTML", "warn");
      exportHTML(filename.replace(/\.docx$/i, ".html"), title, sections);
    });
  }

  /* --------------------------------------------------------------------- */
  /*  PowerPoint (.pptx) via pptxgenjs                                       */
  /* --------------------------------------------------------------------- */
  function exportPPT(filename, title, sections) {
    return _loadScript(CDN.pptx).then(() => {
      if (!global.PptxGenJS) throw new Error("pptxgenjs non disponible — fallback HTML");
      const pptx = new global.PptxGenJS();
      pptx.layout = "LAYOUT_WIDE";
      // Slide 1 — titre.
      const s0 = pptx.addSlide();
      s0.addText("CONFIDENTIEL", { x: 0.5, y: 0.3, w: 12.3, h: 0.3, fontSize: 10, color: "DC2626", bold: true });
      s0.addText(title, { x: 0.5, y: 1.8, w: 12.3, h: 1.0, fontSize: 36, color: "0B3D91", bold: true });
      s0.addText("AGL Direction Stratégie — Sources macro à date juin 2026",
        { x: 0.5, y: 3.0, w: 12.3, h: 0.5, fontSize: 14, color: "6B7280" });
      s0.addText("Généré le " + new Date().toLocaleDateString("fr-FR"),
        { x: 0.5, y: 6.5, w: 12.3, h: 0.3, fontSize: 10, color: "6B7280" });

      sections.forEach((sec) => {
        const s = pptx.addSlide();
        s.addText(sec.heading, { x: 0.5, y: 0.3, w: 12.3, h: 0.6, fontSize: 24, color: "0B3D91", bold: true });
        let y = 1.1;
        if (sec.text) { s.addText(sec.text, { x: 0.5, y: y, w: 12.3, h: 1.5, fontSize: 12, color: "1F2937" }); y += 1.5; }
        if (sec.bullets && sec.bullets.length) {
          s.addText(sec.bullets.map((b) => ({ text: b, options: { bullet: true } })),
            { x: 0.5, y: y, w: 12.3, h: 4.5, fontSize: 14, color: "1F2937" });
          y += 4.5;
        }
        if (sec.table) {
          const tableRows = [sec.table.headers.map((h) => ({
            text: h, options: { bold: true, fill: { color: "F8FAFC" } }
          }))].concat(sec.table.rows.map((r) => r.map((c) => ({ text: c == null ? "" : String(c) }))));
          s.addTable(tableRows, { x: 0.5, y: y, w: 12.3, fontSize: 10, border: { type: "solid", color: "E3E6EA" } });
        }
      });
      return pptx.writeFile({ fileName: filename });
    }).catch((err) => {
      U.log("exports", "PPT KO (" + err.message + ") — fallback HTML", "warn");
      exportHTML(filename.replace(/\.pptx$/i, ".html"), title, sections);
    });
  }

  /* --------------------------------------------------------------------- */
  /*  Builders de contenu par onglet                                         */
  /* --------------------------------------------------------------------- */
  function buildExecutiveSections(snap) {
    const secs = [];
    const headline = [];
    if (snap.marche_total != null) headline.push("Marché total CIV " + (snap.derniere_annee || "") + " : " + U.formatNumber(snap.marche_total));
    if (snap.pdm_globale != null)  headline.push("PDM AGL globale : " + (snap.pdm_globale * 100).toFixed(1) + " %");
    if (snap.clients_rmc != null)  headline.push("RMC : " + U.formatNumber(snap.clients_rmc) + " clients (locked " + (snap.clients_locked || 0) + ")");
    secs.push({ heading: "Headline", bullets: headline.length ? headline : ["Pas encore d'agrégats — charger le pipeline."] });

    if (snap.top_croissance && snap.top_croissance.length) {
      secs.push({
        heading: "Top métiers en croissance",
        table: {
          headers: ["Métier", "CAGR Marché"],
          rows: snap.top_croissance.map((m) => [m.metier, (m.cagr * 100).toFixed(1) + " %"])
        }
      });
    }
    if (snap.top_white_spaces && snap.top_white_spaces.length) {
      secs.push({
        heading: "Top white spaces — marché non capturé",
        table: {
          headers: ["Client", "Métier", "Volume"],
          rows: snap.top_white_spaces.map((w) => [w.client, w.metier, U.formatNumber(w.volume)])
        }
      });
    }
    if (snap.projection_base_2030 && snap.projection_base_2030.length) {
      secs.push({
        heading: "Projections BASE " + (snap.horizon || ""),
        table: {
          headers: ["Métier", "Volume " + (snap.horizon || ""), "CAGR/an"],
          rows: snap.projection_base_2030.map((p) => [p.metier, U.formatNumber(p.volume), (p.cagr * 100).toFixed(1) + " %"])
        }
      });
    }
    if (snap.risques_opportunites && snap.risques_opportunites.length) {
      secs.push({
        heading: "Risques & opportunités — macro juin 2026",
        bullets: snap.risques_opportunites.slice(0, 6).map((x) => "[" + x.type + "] " + x.titre + " — " + x.detail)
      });
    }
    return secs;
  }

  global.AGL = global.AGL || {};
  global.AGL.exports = {
    htmlReport, exportHTML, exportWord, exportPPT,
    buildExecutiveSections
  };
})(typeof window !== "undefined" ? window : this);
