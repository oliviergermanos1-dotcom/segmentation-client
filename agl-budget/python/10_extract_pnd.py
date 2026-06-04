#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — 10 : Extraction PND 2026-2030 Tome 3 (Matrice d'Actions)
=====================================================================
Parse le PDF "PND 2026-2030 Tome 3 — Matrices d'actions prioritaires et cadre
de résultats" et produit :
  - projets_pnd.csv    : 1 ligne par Action/Produit/Effet/Résultat avec
                         axe, code hiérarchique, intitulé, structure responsable,
                         coûts 2026-2030 et Total (millions FCFA)
  - axes_pnd.csv       : agrégat par axe (somme 5 ans)
  - sectoriels_pnd.csv : agrégat par résultat sectoriel (X.YY)

6 axes attendus :
  1 Paix, sécurité
  2 Agriculture
  3 Investissement privé
  4 Capital humain
  5 Infrastructures stratégiques (cœur logistique AGL)
  6 Gouvernance

Méthode : pdftotext -layout puis parsing ligne par ligne avec regex
sur les marqueurs hiérarchiques (Action/Produit/Effet/Résultat + code dotted).

Dépendance externe : `pdftotext` (paquet poppler-utils).
"""
from __future__ import annotations
import argparse
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pandas as pd


AXES_LIBELLE = {
    1: "Paix, sécurité et stabilité durables",
    2: "Modernisation de l'agriculture, sécurisation foncière, chaînes de valeurs",
    3: "Promotion de l'investissement privé, champions nationaux, réduction de l'informalité",
    4: "Développement du capital humain, compétences, emplois décents",
    5: "Infrastructures stratégiques, pôles économiques, transition écologique",
    6: "Bonne gouvernance et modernisation de l'État",
}

# Mots-clés AGL pour scorer la pertinence logistique d'un projet.
KEYWORDS_AGL = re.compile(
    r"\b(port|portuaire|aéroport|aerien|aerodrome|maritime|fluvial|logistique|"
    r"corridor|hinterland|route|autoroute|rail|ferroviaire|fret|transport|"
    r"entrepôt|magasin|terminal|conteneur|TEU|chaîne de valeur|export|import|"
    r"mine|mines|cacao|anacarde|coton|hydrocarbures|pétrole|"
    r"agro[- ]industriel|zone industrielle|raffinerie)\b",
    re.IGNORECASE,
)


# --- Regex de parsing -------------------------------------------------------
# Une ligne de code peut commencer par : Action / Produit / Effet / Résultat
# (avec accent variable) puis un code 1.01.1 ... 6.07.4.5.12
_NIVEAU_RE = re.compile(
    r"^\s*(Action|Produit|Effet|R[ée]sultat[^\d]*?)"
    r"\s+(\d+(?:\.\d+)+)"            # code dotted
    r"\s+(.*)$"
)
# Atome "valeur monétaire FR" : "-", "1 234,5", "85,0", "12 345 678", possessessivement borné.
_NUM = r"(?:-|\d{1,3}(?:[ ]\d{3})*(?:,\d+)?)"
# Bloc de 6 colonnes en fin de ligne (5 années + total) — non-gourmand.
_COSTS_TAIL_RE = re.compile(
    rf"({_NUM})[ ]+({_NUM})[ ]+({_NUM})[ ]+({_NUM})[ ]+({_NUM})[ ]+({_NUM})\s*$"
)
# Lignes de bruit (en-têtes répétés, pieds de page) à ignorer.
_NOISE = re.compile(
    r"^\s*(Matrice d.?Actions|PND 2026|PLAN NATIONAL|"
    r"R[ée]sultat Sectoriel|Effet/Produit/Action|Intitul[ée]|Structure|"
    r"Co[ûu]t \(en millions|^\s*[0-9 ]+\s*$|\s*P\.\d+\s*$|TOME 3|"
    r"MATRICE D.ACTIONS)",
    re.IGNORECASE,
)


def _to_float(s: str) -> float:
    if s is None: return 0.0
    s = s.strip()
    if s in ("", "-", "—"): return 0.0
    s = s.replace(" ", "").replace(" ", "").replace(" ", "")
    # FR : virgule décimale, points en séparateur de milliers (rare ici).
    if s.count(",") == 1 and s.count(".") >= 1:
        # "12.345,67" → 12345.67
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",") == 1:
        s = s.replace(",", ".")
    elif s.count(".") >= 1 and len(s.split(".")[-1]) == 3:
        # "12.345" : milliers
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def pdf_to_text(pdf_path: Path) -> str:
    """Convertit avec pdftotext -layout pour préserver la structure tabulaire."""
    if not shutil.which("pdftotext"):
        sys.exit("ERREUR : pdftotext absent. Installe poppler-utils "
                 "(apt-get install poppler-utils).")
    out = subprocess.run(
        ["pdftotext", "-layout", str(pdf_path), "-"],
        capture_output=True, text=True, check=False
    )
    if out.returncode != 0:
        sys.exit(f"ERREUR pdftotext : {out.stderr[:300]}")
    return out.stdout


def parse(text: str):
    """Parse le texte layout en records. Gère les intitulés multi-lignes."""
    records = []
    current = None

    def flush():
        nonlocal current
        if current is not None:
            records.append(current)
            current = None

    # Buffer pour les lignes de texte sans code marker. Elles peuvent appartenir
    # soit au record précédent (continuation d'intitulé), soit au record suivant
    # (préambule d'intitulé qui sera repris quand on rencontre le code).
    # Heuristique : si la dernière action a un intitulé non-trivial (>= 8 chars
    # et 2+ mots), la continuation va au record SUIVANT. Sinon au courant.
    pending_continuations = []

    for raw in text.splitlines():
        if _NOISE.search(raw):
            flush()
            pending_continuations.clear()
            continue
        m = _NIVEAU_RE.match(raw)
        if m:
            flush()
            niveau = m.group(1).strip().rstrip(":")
            code = m.group(2)
            rest = m.group(3).strip()
            axe = int(code.split(".")[0])
            cm = _COSTS_TAIL_RE.search(rest)
            costs = ("0", "0", "0", "0", "0", "0")
            head = rest
            if cm:
                costs = cm.groups()
                head = rest[:cm.start()].strip()
            # Préfixe : ce qui s'est accumulé en attente.
            prefix = " ".join(pending_continuations).strip()
            pending_continuations.clear()
            intitule = (prefix + " " + head).strip() if prefix and not head else head or prefix
            current = {
                "axe": axe,
                "axe_libelle": AXES_LIBELLE.get(axe, ""),
                "niveau": "Résultat" if niveau.lower().startswith(("résultat", "resultat")) else niveau,
                "code": code,
                "code_sectoriel": ".".join(code.split(".")[:2]),
                "intitule": intitule,
                "structure_responsable": "",
                "cout_2026": _to_float(costs[0]),
                "cout_2027": _to_float(costs[1]),
                "cout_2028": _to_float(costs[2]),
                "cout_2029": _to_float(costs[3]),
                "cout_2030": _to_float(costs[4]),
                "cout_total": _to_float(costs[5]),
            }
        else:
            # Ligne sans code marker. On la met en attente.
            extra = re.sub(r"\s{2,}", " ", raw.strip())
            if not extra:
                continue
            # Si l'intitulé du record courant est court ou vide, c'est probablement
            # une continuation du courant. Sinon, c'est le préambule du prochain.
            if current and (not current["intitule"] or len(current["intitule"]) < 12 or
                            len(current["intitule"].split()) < 3):
                current["intitule"] = (current["intitule"] + " " + extra).strip()
            else:
                pending_continuations.append(extra)
                # On limite à 3 lignes en attente (au-delà, vraisemblablement
                # texte non relié — on jette les plus anciennes).
                if len(pending_continuations) > 3:
                    pending_continuations.pop(0)
    flush()

    # Post-traitement : la "structure responsable" est souvent l'isolat textuel
    # situé entre l'intitulé et la 1ère colonne numérique sur la même ligne ;
    # avec pdftotext layout c'est souvent collé après plusieurs espaces.
    # Heuristique : si l'intitulé contient un séparateur "    " (>=4 espaces),
    # la dernière partie courte (< 60 chars sans virgule décimale) est promue
    # en structure_responsable.
    for r in records:
        m = re.split(r"\s{4,}", r["intitule"])
        if len(m) >= 2 and 1 <= len(m[-1]) <= 60 and not re.search(r"\d{3,}", m[-1]):
            r["structure_responsable"] = m[-1].strip()
            r["intitule"] = " ".join(p.strip() for p in m[:-1])
        r["intitule"] = re.sub(r"\s+", " ", r["intitule"]).strip()
        # Tag AGL : 1 si l'intitulé matche les mots-clés logistique.
        r["pertinence_agl"] = 1 if KEYWORDS_AGL.search(r["intitule"]) else 0

    return records


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — 10 extraction PND Tome 3")
    ap.add_argument("--pdf",    required=True, help="PND Tome 3 (PDF)")
    ap.add_argument("--outdir", default="agreg", help="Dossier de sortie")
    args = ap.parse_args(argv)

    pdf = Path(args.pdf)
    if not pdf.exists():
        sys.exit(f"ERREUR : PDF introuvable : {pdf}")

    print(f"[1/3] Conversion PDF → texte (pdftotext -layout) — {pdf.name}")
    text = pdf_to_text(pdf)
    print(f"      {len(text):,} caractères extraits, {text.count(chr(10)):,} lignes")

    print("[2/3] Parsing hiérarchie Action/Produit/Effet/Résultat")
    rows = parse(text)
    df = pd.DataFrame(rows)
    print(f"      {len(df):,} enregistrements parsés "
          f"({(df.niveau=='Action').sum()} Actions, "
          f"{(df.niveau=='Produit').sum()} Produits, "
          f"{(df.niveau=='Effet').sum()} Effets, "
          f"{(df.niveau=='Résultat').sum()} Résultats)")

    # Sanity check : on s'attend à ~1480 Actions selon le sondage.
    actions = df[df.niveau == "Action"]
    print(f"      Coût total Actions 2026-2030 = "
          f"{actions['cout_total'].sum():,.0f} M FCFA "
          f"({actions['cout_total'].sum()/1e6:,.1f} Mds FCFA)")

    outdir = Path(args.outdir); outdir.mkdir(parents=True, exist_ok=True)
    df.to_csv(outdir / "projets_pnd.csv", index=False)

    # Agrégat par axe.
    axes = (actions
            .groupby(["axe", "axe_libelle"], as_index=False)
            .agg(nb_actions=("code", "count"),
                 cout_2026=("cout_2026", "sum"),
                 cout_2027=("cout_2027", "sum"),
                 cout_2028=("cout_2028", "sum"),
                 cout_2029=("cout_2029", "sum"),
                 cout_2030=("cout_2030", "sum"),
                 cout_total=("cout_total", "sum"),
                 nb_agl=("pertinence_agl", "sum"))
            .sort_values("cout_total", ascending=False))
    axes.to_csv(outdir / "axes_pnd.csv", index=False)

    # Agrégat par résultat sectoriel (X.YY).
    sectoriels = (actions
                  .groupby(["axe", "code_sectoriel"], as_index=False)
                  .agg(nb_actions=("code", "count"),
                       cout_total=("cout_total", "sum"),
                       nb_agl=("pertinence_agl", "sum"))
                  .sort_values("cout_total", ascending=False))
    sectoriels.to_csv(outdir / "sectoriels_pnd.csv", index=False)

    print("[3/3] Récap par axe (somme 5 ans Actions, M FCFA) :")
    for _, r in axes.iterrows():
        agl_tag = f" · {int(r.nb_agl)} actions taggées AGL" if r.nb_agl else ""
        print(f"    {int(r.axe)} {r.axe_libelle[:55]:<55} "
              f"{int(r.nb_actions):>4} actions · {r.cout_total:>15,.0f} M FCFA{agl_tag}")
    print(f"\n[ok] écrit : {outdir/'projets_pnd.csv'} (toutes lignes)")
    print(f"[ok] écrit : {outdir/'axes_pnd.csv'}, {outdir/'sectoriels_pnd.csv'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
