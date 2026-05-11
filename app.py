"""Segmentation client - app Streamlit.

Workflow :
1. Upload d'un ou plusieurs fichiers Excel (extraction CRM Siebel).
   La ligne d'en-tete reelle est detectee automatiquement (les lignes
   "Filtres appliques" et le warning "Exported data limited to 150000 rows"
   sont ignorees). Les fichiers sont concatenes et dedupliques par
   JobfileNumber - utile pour contourner la limite Siebel de 150k lignes.
2. Filtres sur la periode (Year / Month) et autres dimensions.
3. Agregation par client (Customer / HQ / Code CTO au choix) :
   - CA total (Turnover EUR), GM total, TEU, Freight Ton
   - nombre d'operations, nb de mois actifs, premiere / derniere activite
4. Definition des segments via regles sur les metriques agregees.
5. Classification automatique + export Excel / CSV.
"""

from __future__ import annotations

import io
import json
import re
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st


# Fichier local de persistance des regles. Cree automatiquement a cote
# de app.py. Permet de retrouver ses regles d'une session a l'autre.
RULES_DIR = Path(__file__).parent / "configs"
RULES_DIR.mkdir(exist_ok=True)
RULES_FILE = RULES_DIR / "segments_saved.json"


def _load_persisted_segments() -> list[dict[str, Any]]:
    if RULES_FILE.exists():
        try:
            return json.loads(RULES_FILE.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            return []
    return []


def _persist_segments(segments: list[dict[str, Any]]) -> None:
    RULES_FILE.write_text(
        json.dumps(segments, indent=2, ensure_ascii=False), encoding="utf-8"
    )


OPERATORS_NUMERIC = ["==", "!=", ">", ">=", "<", "<=", "between", "est vide", "n'est pas vide"]
OPERATORS_TEXT = [
    "==",
    "!=",
    "contient",
    "ne contient pas",
    "commence par",
    "finit par",
    "dans la liste",
    "pas dans la liste",
    "regex",
    "est vide",
    "n'est pas vide",
]
OPERATORS_DATE = ["==", "!=", ">", ">=", "<", "<=", "between", "est vide", "n'est pas vide"]

HEADER_MARKERS = {"Year", "Customer", "JobfileNumber"}


st.set_page_config(page_title="Segmentation client", layout="wide")
st.title("Segmentation client - export CRM logistique")


# ---------------------------------------------------------------------------
# Session state init
# ---------------------------------------------------------------------------
def _init_state() -> None:
    if "segments" not in st.session_state:
        st.session_state["segments"] = _load_persisted_segments()
    st.session_state.setdefault("raw_df", None)
    st.session_state.setdefault("agg_df", None)
    st.session_state.setdefault("classified", None)
    st.session_state.setdefault("loaded_files", [])


_init_state()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _detect_header_row(xls: pd.ExcelFile, sheet: str, max_scan: int = 15) -> int:
    """Retourne l'index de la ligne d'en-tete reelle d'une feuille Siebel.

    On scanne les premieres lignes et on retient celle qui contient les
    marqueurs caracteristiques de l'export (Year, Customer, JobfileNumber).
    Fallback : 0.
    """
    probe = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=max_scan, dtype=str)
    for i, row in probe.iterrows():
        values = {str(v).strip() for v in row.tolist() if pd.notna(v)}
        if HEADER_MARKERS.issubset(values):
            return int(i)
    return 0


@st.cache_data(show_spinner=False)
def _load_excel(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    bio = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(bio)
    sheet = xls.sheet_names[0]
    header_row = _detect_header_row(xls, sheet)
    df = pd.read_excel(xls, sheet_name=sheet, header=header_row)
    df["__source_file__"] = file_name
    return df


def _concat_dedup(frames: list[pd.DataFrame]) -> pd.DataFrame:
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True, sort=False)
    if "JobfileNumber" in df.columns:
        before = len(df)
        df = df.drop_duplicates(subset=["JobfileNumber"], keep="first")
        st.session_state["_dedup_removed"] = before - len(df)
    return df


def _ensure_period(df: pd.DataFrame) -> pd.DataFrame:
    """Construit une colonne date 'Period' a partir de Year/Month si possible."""
    if {"Year", "Month"}.issubset(df.columns):
        try:
            df = df.copy()
            df["Period"] = pd.to_datetime(
                df["Year"].astype("Int64").astype(str)
                + "-"
                + df["Month"].astype("Int64").astype(str).str.zfill(2)
                + "-01",
                errors="coerce",
            )
        except Exception:  # noqa: BLE001
            pass
    return df


# ---------------------------------------------------------------------------
# Step 1 : upload
# ---------------------------------------------------------------------------
st.header("1. Importer les fichiers Excel")
st.caption(
    "L'export Siebel est limite a 150 000 lignes : importez plusieurs extractions "
    "(par exemple une par annee) - elles seront concatenees et dedupliquees par "
    "JobfileNumber. La ligne d'en-tete reelle est detectee automatiquement."
)

uploaded_files = st.file_uploader(
    "Fichiers .xlsx", type=["xlsx", "xls"], accept_multiple_files=True
)

if uploaded_files:
    frames = []
    progress = st.progress(0.0, text="Chargement...")
    for i, f in enumerate(uploaded_files):
        try:
            df_part = _load_excel(f.getvalue(), f.name)
            frames.append(df_part)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Erreur sur {f.name} : {exc}")
        progress.progress((i + 1) / len(uploaded_files), text=f"{f.name}")
    progress.empty()
    df = _concat_dedup(frames)
    df = _ensure_period(df)
    st.session_state.raw_df = df
    st.session_state.loaded_files = [f.name for f in uploaded_files]

raw_df = st.session_state.raw_df

if raw_df is not None:
    c1, c2, c3 = st.columns(3)
    c1.metric("Lignes totales", f"{len(raw_df):,}".replace(",", " "))
    c2.metric("Colonnes", len(raw_df.columns))
    removed = st.session_state.get("_dedup_removed", 0)
    c3.metric("Doublons supprimes", f"{removed:,}".replace(",", " "))
    with st.expander("Apercu (20 lignes)"):
        st.dataframe(raw_df.head(20), use_container_width=True)
    with st.expander("Colonnes detectees"):
        st.write(list(raw_df.columns))


# ---------------------------------------------------------------------------
# Step 2 : filters
# ---------------------------------------------------------------------------
st.header("2. Filtrer la periode et les dimensions")

filtered_df = None
if raw_df is None:
    st.info("Importez d'abord au moins un fichier.")
else:
    filtered_df = raw_df.copy()

    fcol1, fcol2 = st.columns(2)
    if "Year" in filtered_df.columns:
        years_avail = sorted([int(y) for y in filtered_df["Year"].dropna().unique()])
        years_sel = fcol1.multiselect("Annees", years_avail, default=years_avail)
        filtered_df = filtered_df[filtered_df["Year"].isin(years_sel)]
    if "Month" in filtered_df.columns:
        months = sorted([int(m) for m in filtered_df["Month"].dropna().unique()])
        if months:
            m_min, m_max = fcol2.select_slider(
                "Mois (plage)",
                options=months,
                value=(min(months), max(months)),
            )
            filtered_df = filtered_df[
                (filtered_df["Month"].astype("Int64") >= m_min)
                & (filtered_df["Month"].astype("Int64") <= m_max)
            ]

    cat_dims = [
        c
        for c in [
            "Reporting Entity Label",
            "Zone 3 Label",
            "Market Segment Name",
            "Product Name",
            "Activity Name (EN)",
            "Shipping Type Name",
        ]
        if c in filtered_df.columns
    ]
    if cat_dims:
        with st.expander("Filtres dimensions (optionnel)"):
            for dim in cat_dims:
                opts = sorted([v for v in filtered_df[dim].dropna().unique()])
                sel = st.multiselect(dim, opts, default=opts, key=f"flt_{dim}")
                if len(sel) != len(opts):
                    filtered_df = filtered_df[filtered_df[dim].isin(sel)]

    st.metric("Lignes apres filtres", f"{len(filtered_df):,}".replace(",", " "))


# ---------------------------------------------------------------------------
# Step 3 : aggregation by client
# ---------------------------------------------------------------------------
st.header("3. Agreger par client")

agg_df: pd.DataFrame | None = None
client_key: str | None = None

if filtered_df is None or filtered_df.empty:
    st.info("Filtrez d'abord la base.")
else:
    candidate_keys = [
        c
        for c in [
            "Customer",
            "HQ CTO Customer Name",
            "CTO / RCU Customer Code",
            "CTO Name And Code",
        ]
        if c in filtered_df.columns
    ]
    client_key = st.selectbox(
        "Colonne identifiant le client",
        candidate_keys or filtered_df.columns.tolist(),
        index=0,
    )

    numeric_metrics = {
        "Turnover in EUR": "CA_EUR",
        "Direct GM in EUR": "GM_EUR",
        "Gross Billing Tax Excl. in EUR": "Billing_HT_EUR",
        "TEU": "TEU",
        "Freight Ton": "Freight_Ton",
    }
    available_metrics = {k: v for k, v in numeric_metrics.items() if k in filtered_df.columns}

    metrics_sel = st.multiselect(
        "Metriques a agreger (somme)",
        list(available_metrics.keys()),
        default=list(available_metrics.keys()),
    )

    if st.button("Lancer l'agregation", type="primary"):
        with st.spinner("Agregation en cours..."):
            grp = filtered_df.groupby(client_key, dropna=False)
            agg_spec: dict[str, Any] = {}
            for col in metrics_sel:
                agg_spec[available_metrics[col]] = (col, "sum")
            if "JobfileNumber" in filtered_df.columns:
                agg_spec["Nb_operations"] = ("JobfileNumber", "nunique")
            if "Period" in filtered_df.columns:
                agg_spec["Premiere_activite"] = ("Period", "min")
                agg_spec["Derniere_activite"] = ("Period", "max")
                agg_spec["Nb_mois_actifs"] = ("Period", "nunique")
            for extra in ["HQ CTO Customer Name", "Market Segment Name", "Reporting Entity Label"]:
                if extra in filtered_df.columns and extra != client_key:
                    agg_spec[extra] = (extra, lambda s: s.dropna().mode().iloc[0] if not s.dropna().empty else None)
            agg = grp.agg(**agg_spec).reset_index()
            agg = agg.sort_values(
                by=next(iter([v for v in available_metrics.values() if v in agg.columns]), agg.columns[1]),
                ascending=False,
            )
            st.session_state.agg_df = agg

agg_df = st.session_state.agg_df

if agg_df is not None:
    st.success(f"{len(agg_df):,} clients agreges.".replace(",", " "))
    st.dataframe(agg_df.head(50), use_container_width=True)


# ---------------------------------------------------------------------------
# Step 4 : segment rules
# ---------------------------------------------------------------------------
st.header("4. Definir les regles de segmentation")

if agg_df is None:
    st.info("Lancez d'abord l'agregation.")
else:
    st.markdown(
        """
        Les regles sont **sauvegardees automatiquement** dans
        `configs/segments_saved.json` et rechargees a chaque ouverture de l'app.
        Vous pouvez **ajouter**, **modifier** ou **supprimer** une regle a tout moment.

        Les segments sont evalues **dans l'ordre** : le premier qui matche est
        attribue au client. Reordonnez avec les boutons monter / descendre.
        """
    )

    # ---- barre d'etat
    bar = st.columns([3, 1, 1, 1])
    bar[0].metric("Regles definies", len(st.session_state.segments))
    if bar[1].button("Tout effacer", help="Supprime toutes les regles"):
        st.session_state.segments = []
        _persist_segments([])
        st.rerun()
    if bar[2].button(
        "Quartiles CA",
        help="Cree 4 regles automatiques basees sur les quartiles du CA",
        disabled="CA_EUR" not in agg_df.columns,
    ):
        q = agg_df["CA_EUR"].quantile([0.25, 0.5, 0.75]).tolist()
        st.session_state.segments = [
            {"name": "VIP", "logic": "AND",
             "conditions": [{"column": "CA_EUR", "op": ">=", "value": str(int(q[2]))}]},
            {"name": "Gold", "logic": "AND",
             "conditions": [{"column": "CA_EUR", "op": ">=", "value": str(int(q[1]))}]},
            {"name": "Silver", "logic": "AND",
             "conditions": [{"column": "CA_EUR", "op": ">=", "value": str(int(q[0]))}]},
            {"name": "Bronze", "logic": "AND",
             "conditions": [{"column": "CA_EUR", "op": ">", "value": "0"}]},
        ]
        _persist_segments(st.session_state.segments)
        st.rerun()
    if bar[3].button(
        "Sauver sous...",
        help="Telecharge la config actuelle en JSON (sauvegarde / partage)",
        disabled=not st.session_state.segments,
    ):
        st.session_state["_show_export"] = True

    # ---- import / export json (replie dans un expander)
    with st.expander("Import / export JSON (sauvegarde, partage entre postes)"):
        io1, io2 = st.columns(2)
        with io1:
            if st.session_state.segments:
                st.download_button(
                    "Telecharger config (JSON)",
                    data=json.dumps(
                        st.session_state.segments, indent=2, ensure_ascii=False
                    ),
                    file_name="segments.json",
                    mime="application/json",
                )
        with io2:
            cfg_file = st.file_uploader(
                "Charger une config", type=["json"], key="cfg_uploader"
            )
            if cfg_file is not None:
                try:
                    st.session_state.segments = json.load(cfg_file)
                    _persist_segments(st.session_state.segments)
                    st.success("Config chargee et sauvegardee localement.")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"JSON invalide : {exc}")

    st.markdown("---")
    st.subheader("Ajouter une nouvelle regle")
    with st.form("add_segment_form", clear_on_submit=True):
        new_name = st.text_input(
            "Nom du segment",
            placeholder="ex : VIP, Gold, Strategique, A_relancer...",
        )
        submitted = st.form_submit_button("Ajouter")
        if submitted and new_name.strip():
            st.session_state.segments.append(
                {"name": new_name.strip(), "logic": "AND", "conditions": []}
            )
            _persist_segments(st.session_state.segments)
            st.rerun()

    if st.session_state.segments:
        st.markdown("---")
        st.subheader("Regles actuelles")

    for idx, seg in enumerate(st.session_state.segments):
        with st.container(border=True):
            st.markdown(f"**Regle #{idx + 1}**  -  {len(seg['conditions'])} condition(s)")
            top = st.columns([4, 1, 1, 1, 1])
            new_name = top[0].text_input(
                "Nom du segment",
                value=seg["name"],
                key=f"name_{idx}",
                label_visibility="collapsed",
            )
            if new_name != seg["name"]:
                seg["name"] = new_name
                _persist_segments(st.session_state.segments)
            if top[1].button("monter", key=f"up_{idx}", disabled=idx == 0):
                st.session_state.segments[idx - 1], st.session_state.segments[idx] = (
                    st.session_state.segments[idx],
                    st.session_state.segments[idx - 1],
                )
                _persist_segments(st.session_state.segments)
                st.rerun()
            if top[2].button(
                "descendre",
                key=f"down_{idx}",
                disabled=idx == len(st.session_state.segments) - 1,
            ):
                st.session_state.segments[idx + 1], st.session_state.segments[idx] = (
                    st.session_state.segments[idx],
                    st.session_state.segments[idx + 1],
                )
                _persist_segments(st.session_state.segments)
                st.rerun()
            new_logic = top[3].selectbox(
                "Logique",
                ["AND", "OR"],
                index=0 if seg.get("logic", "AND") == "AND" else 1,
                key=f"logic_{idx}",
                label_visibility="collapsed",
                help="AND = toutes les conditions doivent etre vraies. OR = au moins une.",
            )
            if new_logic != seg.get("logic"):
                seg["logic"] = new_logic
                _persist_segments(st.session_state.segments)
            if top[4].button("Supprimer", key=f"del_{idx}"):
                st.session_state.segments.pop(idx)
                _persist_segments(st.session_state.segments)
                st.rerun()

            st.markdown("**Conditions**")
            for c_idx, cond in enumerate(seg["conditions"]):
                cols = st.columns([3, 2, 4, 1])
                col_name = cols[0].selectbox(
                    "Colonne",
                    agg_df.columns.tolist(),
                    index=(
                        agg_df.columns.tolist().index(cond["column"])
                        if cond.get("column") in agg_df.columns
                        else 0
                    ),
                    key=f"col_{idx}_{c_idx}",
                    label_visibility="collapsed",
                )
                cond["column"] = col_name
                series = agg_df[col_name]
                if pd.api.types.is_numeric_dtype(series):
                    ops = OPERATORS_NUMERIC
                elif pd.api.types.is_datetime64_any_dtype(series):
                    ops = OPERATORS_DATE
                else:
                    ops = OPERATORS_TEXT
                op_index = ops.index(cond["op"]) if cond.get("op") in ops else 0
                cond["op"] = cols[1].selectbox(
                    "Operateur",
                    ops,
                    index=op_index,
                    key=f"op_{idx}_{c_idx}",
                    label_visibility="collapsed",
                )

                if cond["op"] in ("est vide", "n'est pas vide"):
                    cond["value"] = None
                    cols[2].markdown("&nbsp;")
                elif cond["op"] == "between":
                    sub = cols[2].columns(2)
                    cond["value"] = [
                        sub[0].text_input(
                            "min",
                            value=str(cond.get("value", ["", ""])[0])
                            if isinstance(cond.get("value"), list)
                            else "",
                            key=f"vmin_{idx}_{c_idx}",
                            label_visibility="collapsed",
                        ),
                        sub[1].text_input(
                            "max",
                            value=str(cond.get("value", ["", ""])[1])
                            if isinstance(cond.get("value"), list)
                            else "",
                            key=f"vmax_{idx}_{c_idx}",
                            label_visibility="collapsed",
                        ),
                    ]
                elif cond["op"] in ("dans la liste", "pas dans la liste"):
                    cond["value"] = cols[2].text_input(
                        "valeurs separees par des virgules",
                        value=cond.get("value", "") or "",
                        key=f"v_{idx}_{c_idx}",
                        label_visibility="collapsed",
                    )
                else:
                    cond["value"] = cols[2].text_input(
                        "valeur",
                        value=str(cond.get("value", "")) if cond.get("value") is not None else "",
                        key=f"v_{idx}_{c_idx}",
                        label_visibility="collapsed",
                    )

                if cols[3].button("retirer", key=f"delcond_{idx}_{c_idx}"):
                    seg["conditions"].pop(c_idx)
                    _persist_segments(st.session_state.segments)
                    st.rerun()

            if st.button("+ ajouter une condition", key=f"addcond_{idx}"):
                seg["conditions"].append(
                    {"column": agg_df.columns[0], "op": "==", "value": ""}
                )
                _persist_segments(st.session_state.segments)
                st.rerun()

    # ---- persiste les valeurs/operateurs modifies sur place (text_input,
    # selectbox sans rerun explicite). Idempotent et rapide.
    _persist_segments(st.session_state.segments)


# ---------------------------------------------------------------------------
# Step 5 : classification engine
# ---------------------------------------------------------------------------
def _coerce(series: pd.Series, value: Any) -> Any:
    if value is None or value == "":
        return value
    if pd.api.types.is_numeric_dtype(series):
        try:
            return float(value)
        except (TypeError, ValueError):
            return value
    if pd.api.types.is_datetime64_any_dtype(series):
        try:
            return pd.to_datetime(value)
        except (TypeError, ValueError):
            return value
    return str(value)


def _evaluate_condition(df: pd.DataFrame, cond: dict[str, Any]) -> pd.Series:
    col = cond["column"]
    if col not in df.columns:
        return pd.Series([False] * len(df), index=df.index)
    op = cond["op"]
    raw = cond.get("value")
    series = df[col]

    if op == "est vide":
        return series.isna() | (series.astype(str).str.strip() == "")
    if op == "n'est pas vide":
        return ~(series.isna() | (series.astype(str).str.strip() == ""))
    if op == "between":
        vmin = _coerce(series, raw[0] if isinstance(raw, list) else None)
        vmax = _coerce(series, raw[1] if isinstance(raw, list) else None)
        return series.between(vmin, vmax)
    if op in ("dans la liste", "pas dans la liste"):
        items = [v.strip() for v in str(raw or "").split(",") if v.strip()]
        if pd.api.types.is_numeric_dtype(series):
            items = [_coerce(series, v) for v in items]
        result = series.isin(items)
        return ~result if op == "pas dans la liste" else result
    if op == "contient":
        return series.astype(str).str.contains(str(raw or ""), case=False, na=False)
    if op == "ne contient pas":
        return ~series.astype(str).str.contains(str(raw or ""), case=False, na=False)
    if op == "commence par":
        return series.astype(str).str.startswith(str(raw or ""), na=False)
    if op == "finit par":
        return series.astype(str).str.endswith(str(raw or ""), na=False)
    if op == "regex":
        try:
            return series.astype(str).str.contains(str(raw or ""), regex=True, na=False)
        except re.error:
            return pd.Series([False] * len(series), index=series.index)

    value = _coerce(series, raw)
    if op == "==":
        if pd.api.types.is_numeric_dtype(series) or pd.api.types.is_datetime64_any_dtype(series):
            return series == value
        return series.astype(str).str.lower() == str(value).lower()
    if op == "!=":
        if pd.api.types.is_numeric_dtype(series) or pd.api.types.is_datetime64_any_dtype(series):
            return series != value
        return series.astype(str).str.lower() != str(value).lower()
    if op == ">":
        return series > value
    if op == ">=":
        return series >= value
    if op == "<":
        return series < value
    if op == "<=":
        return series <= value
    return pd.Series([False] * len(series), index=series.index)


def _segment_mask(df: pd.DataFrame, segment: dict[str, Any]) -> pd.Series:
    if not segment["conditions"]:
        return pd.Series([False] * len(df), index=df.index)
    masks = [_evaluate_condition(df, c) for c in segment["conditions"]]
    combined = masks[0]
    for m in masks[1:]:
        combined = combined & m if segment.get("logic", "AND") == "AND" else combined | m
    return combined.fillna(False)


def classify(df: pd.DataFrame, segments: list[dict[str, Any]]) -> pd.DataFrame:
    result = df.copy()
    seg_col = pd.Series(["Non classe"] * len(df), index=df.index)
    for seg in segments:
        mask = _segment_mask(df, seg)
        seg_col = seg_col.mask((seg_col == "Non classe") & mask, seg["name"])
    result["segment"] = seg_col
    return result


# ---------------------------------------------------------------------------
# Step 6 : run classification + display + export
# ---------------------------------------------------------------------------
st.header("5. Classifier et exporter")

if agg_df is None:
    st.info("Lancez d'abord l'agregation.")
elif not st.session_state.segments:
    st.info("Definissez au moins un segment.")
else:
    if st.button("Lancer la classification", type="primary"):
        try:
            st.session_state.classified = classify(agg_df, st.session_state.segments)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Erreur pendant la classification : {exc}")

    classified = st.session_state.classified
    if classified is not None:
        st.subheader("Repartition")
        counts = (
            classified["segment"].value_counts().rename_axis("segment").reset_index(name="nb_clients")
        )
        counts["%"] = (counts["nb_clients"] / len(classified) * 100).round(1)
        if "CA_EUR" in classified.columns:
            ca_by_seg = (
                classified.groupby("segment")["CA_EUR"].sum().rename("CA_total_EUR").reset_index()
            )
            counts = counts.merge(ca_by_seg, on="segment", how="left")
            counts["%_CA"] = (counts["CA_total_EUR"] / counts["CA_total_EUR"].sum() * 100).round(1)
        c1, c2 = st.columns([1, 2])
        c1.dataframe(counts, use_container_width=True, hide_index=True)
        if "CA_EUR" in classified.columns:
            c2.bar_chart(counts.set_index("segment")["CA_total_EUR"])
        else:
            c2.bar_chart(counts.set_index("segment")["nb_clients"])

        st.subheader("Clients classes")
        segs_filter = st.multiselect(
            "Filtrer par segment",
            options=classified["segment"].unique().tolist(),
            default=classified["segment"].unique().tolist(),
        )
        view = classified[classified["segment"].isin(segs_filter)]
        st.dataframe(view, use_container_width=True)

        csv_bytes = view.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Telecharger en CSV",
            data=csv_bytes,
            file_name="clients_segmentes.csv",
            mime="text/csv",
        )
        xls_buf = io.BytesIO()
        with pd.ExcelWriter(xls_buf, engine="xlsxwriter") as writer:
            view.to_excel(writer, sheet_name="clients", index=False)
            counts.to_excel(writer, sheet_name="repartition", index=False)
        st.download_button(
            "Telecharger en Excel",
            data=xls_buf.getvalue(),
            file_name="clients_segmentes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
