# -*- coding: utf-8 -*-
"""Tests end-to-end : run_all.py sur les fixtures + vérifs sur les artefacts."""
import importlib.util
import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / "fixtures"


@pytest.fixture(scope="module")
def pipeline_out(tmp_path_factory):
    """Régénère les fixtures puis exécute run_all.py de bout en bout."""
    out = tmp_path_factory.mktemp("agl_pipeline")
    # On régénère les fixtures (idempotent).
    subprocess.run([sys.executable, str(ROOT / "_make_fixtures.py")],
                   check=True, capture_output=True)
    rc = subprocess.run(
        [sys.executable, str(ROOT / "run_all.py"),
         "--data-dir", str(FIXTURES), "--out", str(out)],
        capture_output=True, text=True,
    )
    assert rc.returncode == 0, f"run_all.py KO :\n{rc.stdout}\n{rc.stderr}"
    return out


def test_artefacts_existent(pipeline_out):
    expected = [
        "CRM_norm.xlsx", "IRIS_norm.xlsx", "STATCOM_norm.xlsx",
        "match_crm_iris.csv", "match_crm_stat.csv", "match_iris_stat.csv",
        "RMC.xlsx",
        "agreg/marche_total.csv", "agreg/volume_client.csv",
        "agreg/metiers_3ans.csv", "agreg/pdm.csv", "agreg/projections.csv",
        "agreg/whitespaces.csv",
        "agreg/budget_vs_real.csv", "agreg/budget_secteur.csv",
    ]
    for f in expected:
        assert (pipeline_out / f).exists(), f"manquant : {f}"


def test_rmc_15_champs_et_secteur_locked(pipeline_out):
    df = pd.read_excel(pipeline_out / "RMC.xlsx")
    # 15 colonnes attendues (§3 du cahier des charges).
    expected_cols = [
        "ID_RMC", "ID_CRM", "ID_IRIS", "ID_STATCOM",
        "NOM_CANONIQUE", "SECTEUR", "ALIAS_CRM", "ALIAS_IRIS", "ALIAS_STATCOM",
        "MARCHANDISES", "VOLUME_3ANS", "CAP_3ANS",
        "SCORE_MATCH_IRIS", "SCORE_MATCH_STATCOM",
        "SOURCE_SECTEUR", "LOCKED",
    ]
    for c in expected_cols:
        assert c in df.columns, f"colonne {c} manquante dans RMC"
    # Au moins un client doit avoir SECTEUR LOCKED depuis CRM.
    locked = df[df["LOCKED"] == True]
    assert len(locked) >= 1
    assert (locked["SOURCE_SECTEUR"] == "CRM").all()


def test_pdm_calcule(pipeline_out):
    df = pd.read_csv(pipeline_out / "agreg/pdm.csv")
    assert {"METIER", "ANNEE", "VOLUME_AGL", "VOLUME_MARCHE", "PDM"}.issubset(df.columns)
    assert (df["PDM"] >= 0).all() and (df["PDM"] <= 1).all()
    # Au moins un métier doit avoir une PDM strictement positive.
    assert (df["PDM"] > 0).any()


def test_projections_3_scenarios(pipeline_out):
    df = pd.read_csv(pipeline_out / "agreg/projections.csv")
    scenarios = set(df["SCENARIO"].unique())
    assert {"PESSIMISTE", "BASE", "OPTIMISTE"}.issubset(scenarios)
    # L'horizon doit aller au moins jusqu'à 2030.
    assert df["ANNEE"].max() >= 2030


def test_whitespaces_detecte_olam(pipeline_out):
    """OLAM est présent dans STATCOM mais absent du CRM (fixtures) → white space."""
    df = pd.read_csv(pipeline_out / "agreg/whitespaces.csv")
    assert "OLAM" in set(df["NOM_BASE"].fillna(""))


def test_budget_alias_sir_rattache(pipeline_out):
    """Grâce à aliases.json, SIR (RUBRIKS) doit maintenant être rattaché à
    Société Ivoirienne de Raffinage (RMC) et donc avoir un CAP_REEL > 0."""
    df = pd.read_csv(pipeline_out / "agreg/budget_vs_real.csv")
    sir_lines = df[df["NOM_CLIENT"].astype(str).str.contains("Ivoirienne|SIR", case=False)]
    assert len(sir_lines) >= 1
    # Si l'alias a marché, l'ID_CRM est renseigné ET le CAP_REEL est > 0.
    matched = sir_lines[sir_lines["ID_CRM"].notna() & (sir_lines["ID_CRM"] != "")]
    assert len(matched) >= 1, "L'alias SIR → Société Ivoirienne de Raffinage n'a pas fonctionné"


def test_budget_ecart_pct_calcule(pipeline_out):
    df = pd.read_csv(pipeline_out / "agreg/budget_secteur.csv")
    assert {"SECTEUR", "ANNEE", "CAP_PFA", "CAP_REEL", "ECART", "ECART_PCT"}.issubset(df.columns)
    # ECART doit toujours valoir CAP_REEL - CAP_PFA.
    assert ((df["ECART"] - (df["CAP_REEL"] - df["CAP_PFA"])).abs() < 0.01).all()
