"""Segmentation client - app Streamlit.

Workflow :
1. Upload du fichier Excel brut (extraction CRM).
2. Definition des segments via des regles (conditions sur les colonnes).
3. Classification automatique de chaque client dans le premier segment matche.
4. Export du resultat en Excel / CSV.
"""

from __future__ import annotations

import io
import json
import re
from typing import Any

import pandas as pd
import streamlit as st


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


st.set_page_config(page_title="Segmentation client", layout="wide")
st.title("Segmentation client a partir d'un export CRM")

# ---------------------------------------------------------------------------
# Session state init
# ---------------------------------------------------------------------------
if "segments" not in st.session_state:
    st.session_state.segments: list[dict[str, Any]] = []
if "df" not in st.session_state:
    st.session_state.df = None


# ---------------------------------------------------------------------------
# Step 1 : upload
# ---------------------------------------------------------------------------
st.header("1. Importer le fichier Excel")
uploaded = st.file_uploader("Fichier .xlsx ou .xls", type=["xlsx", "xls"])

if uploaded is not None:
    try:
        xls = pd.ExcelFile(uploaded)
        sheet = st.selectbox("Feuille a utiliser", xls.sheet_names)
        header_row = st.number_input(
            "Ligne d'en-tete (0 = premiere ligne)", min_value=0, value=0, step=1
        )
        df = pd.read_excel(xls, sheet_name=sheet, header=int(header_row))
        st.session_state.df = df
        st.success(f"{len(df)} lignes chargees, {len(df.columns)} colonnes.")
        with st.expander("Apercu des 20 premieres lignes"):
            st.dataframe(df.head(20), use_container_width=True)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Erreur de lecture : {exc}")


df = st.session_state.df

# ---------------------------------------------------------------------------
# Step 2 : segment rules
# ---------------------------------------------------------------------------
st.header("2. Definir les segments")

if df is None:
    st.info("Importez d'abord un fichier Excel.")
else:
    st.caption(
        "Chaque segment est compose d'une ou plusieurs conditions. "
        "Choisissez si toutes (ET) ou au moins une (OU) doivent etre vraies. "
        "Les segments sont evalues dans l'ordre : le **premier** segment qui matche "
        "est attribue au client. Reordonnez-les via les fleches."
    )

    # ---- import / export json
    col_io1, col_io2 = st.columns(2)
    with col_io1:
        if st.session_state.segments:
            st.download_button(
                "Telecharger la config des segments (JSON)",
                data=json.dumps(st.session_state.segments, indent=2, ensure_ascii=False),
                file_name="segments.json",
                mime="application/json",
            )
    with col_io2:
        cfg_file = st.file_uploader(
            "Charger une config segments (JSON)", type=["json"], key="cfg_uploader"
        )
        if cfg_file is not None:
            try:
                st.session_state.segments = json.load(cfg_file)
                st.success("Config chargee.")
            except Exception as exc:  # noqa: BLE001
                st.error(f"JSON invalide : {exc}")

    # ---- add segment
    with st.form("add_segment_form", clear_on_submit=True):
        new_name = st.text_input("Nom du nouveau segment", placeholder="ex : VIP")
        submitted = st.form_submit_button("Ajouter le segment")
        if submitted and new_name.strip():
            st.session_state.segments.append(
                {"name": new_name.strip(), "logic": "AND", "conditions": []}
            )
            st.rerun()

    # ---- existing segments
    for idx, seg in enumerate(st.session_state.segments):
        with st.container(border=True):
            top = st.columns([4, 1, 1, 1, 1])
            seg["name"] = top[0].text_input(
                "Nom", value=seg["name"], key=f"name_{idx}", label_visibility="collapsed"
            )
            if top[1].button("↑", key=f"up_{idx}", disabled=idx == 0):
                st.session_state.segments[idx - 1], st.session_state.segments[idx] = (
                    st.session_state.segments[idx],
                    st.session_state.segments[idx - 1],
                )
                st.rerun()
            if top[2].button(
                "↓",
                key=f"down_{idx}",
                disabled=idx == len(st.session_state.segments) - 1,
            ):
                st.session_state.segments[idx + 1], st.session_state.segments[idx] = (
                    st.session_state.segments[idx],
                    st.session_state.segments[idx + 1],
                )
                st.rerun()
            seg["logic"] = top[3].selectbox(
                "Logique",
                ["AND", "OR"],
                index=0 if seg.get("logic", "AND") == "AND" else 1,
                key=f"logic_{idx}",
                label_visibility="collapsed",
            )
            if top[4].button("Supprimer", key=f"del_{idx}"):
                st.session_state.segments.pop(idx)
                st.rerun()

            st.markdown("**Conditions**")
            for c_idx, cond in enumerate(seg["conditions"]):
                cols = st.columns([3, 2, 4, 1])
                col_name = cols[0].selectbox(
                    "Colonne",
                    df.columns.tolist(),
                    index=(
                        df.columns.tolist().index(cond["column"])
                        if cond.get("column") in df.columns
                        else 0
                    ),
                    key=f"col_{idx}_{c_idx}",
                    label_visibility="collapsed",
                )
                cond["column"] = col_name

                series = df[col_name]
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

                if cols[3].button("✕", key=f"delcond_{idx}_{c_idx}"):
                    seg["conditions"].pop(c_idx)
                    st.rerun()

            if st.button("+ ajouter une condition", key=f"addcond_{idx}"):
                seg["conditions"].append(
                    {"column": df.columns[0], "op": "==", "value": ""}
                )
                st.rerun()


# ---------------------------------------------------------------------------
# Step 3 : classification
# ---------------------------------------------------------------------------
def _coerce(series: pd.Series, value: Any) -> Any:
    """Convertit la valeur saisie au type de la colonne quand c'est possible."""
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
    matches_col = pd.Series([[] for _ in range(len(df))], index=df.index)
    for seg in segments:
        mask = _segment_mask(df, seg)
        seg_col = seg_col.where(seg_col != "Non classe", other=seg["name"]).where(
            ~(mask & (seg_col == "Non classe")), other=seg["name"]
        )
        for i in df.index[mask]:
            matches_col.loc[i] = matches_col.loc[i] + [seg["name"]]
    # premier match gagnant
    seg_col = pd.Series(["Non classe"] * len(df), index=df.index)
    for seg in segments:
        mask = _segment_mask(df, seg)
        seg_col = seg_col.mask((seg_col == "Non classe") & mask, seg["name"])
    result["segment"] = seg_col
    result["segments_matches"] = matches_col.apply(lambda lst: ", ".join(lst))
    return result


st.header("3. Classifier")

if df is None:
    st.info("Importez un fichier.")
elif not st.session_state.segments:
    st.info("Definissez au moins un segment.")
else:
    if st.button("Lancer la classification", type="primary"):
        try:
            classified = classify(df, st.session_state.segments)
            st.session_state.classified = classified
        except Exception as exc:  # noqa: BLE001
            st.error(f"Erreur pendant la classification : {exc}")

    classified = st.session_state.get("classified")
    if classified is not None:
        st.subheader("Repartition")
        counts = classified["segment"].value_counts().rename_axis("segment").reset_index(
            name="nb_clients"
        )
        counts["%"] = (counts["nb_clients"] / len(classified) * 100).round(1)
        c1, c2 = st.columns([1, 2])
        c1.dataframe(counts, use_container_width=True, hide_index=True)
        c2.bar_chart(counts.set_index("segment")["nb_clients"])

        st.subheader("Clients classes")
        segs_filter = st.multiselect(
            "Filtrer par segment",
            options=classified["segment"].unique().tolist(),
            default=classified["segment"].unique().tolist(),
        )
        view = classified[classified["segment"].isin(segs_filter)]
        st.dataframe(view, use_container_width=True)

        # exports
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
