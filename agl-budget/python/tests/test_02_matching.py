# -*- coding: utf-8 -*-
"""Tests du module matching_lib (blocking + scoring + classification)."""
import pandas as pd


def _df(records):
    return pd.DataFrame(records)


def test_blocking_keys_multi_token(matching):
    keys = matching._blocking_keys("SOCIETE IVOIRIENNE RAFFINAGE", 3)
    # Doit générer 1 clé par token significatif → permet match "SIR" via aucun token,
    # mais permet à "IVOIRIENNE..." de croiser avec "SIR ..." si l'autre côté a "IVO".
    assert any(k.startswith("IVO") for k in keys)
    assert any(k.startswith("RAF") for k in keys)


def test_match_exact_score_eleve(matching):
    L = _df([{"id": "L1", "NOM_BASE": "NESTLE"}])
    R = _df([{"id": "R1", "NOM_BASE": "NESTLE"}])
    res = matching.match_two_tables(L, R, "id", "id", flux_label="TEST")
    assert len(res) == 1
    assert res.iloc[0]["score"] >= 0.95
    assert res.iloc[0]["classe"] == "AUTO"


def test_match_typo_classe_ambigu_ou_auto(matching):
    L = _df([{"id": "L1", "NOM_BASE": "CARGIL"}])
    R = _df([{"id": "R1", "NOM_BASE": "CARGILL"}])  # double L
    res = matching.match_two_tables(L, R, "id", "id", flux_label="TEST")
    assert len(res) == 1
    assert res.iloc[0]["classe"] in ("AUTO", "AMBIGU")
    assert res.iloc[0]["score"] >= 0.65


def test_match_aucun_blocking_resultat_vide(matching):
    L = _df([{"id": "L1", "NOM_BASE": "ALPHA"}])
    R = _df([{"id": "R1", "NOM_BASE": "ZULU"}])
    res = matching.match_two_tables(L, R, "id", "id", flux_label="TEST")
    # Préfixes différents → pas de paire candidate.
    assert len(res) == 0


def test_match_unicite_par_id_left(matching):
    """Un client gauche ne doit avoir qu'un seul (meilleur) match droit."""
    L = _df([{"id": "L1", "NOM_BASE": "ORANGE"}])
    R = _df([
        {"id": "R1", "NOM_BASE": "ORANGE"},
        {"id": "R2", "NOM_BASE": "ORANGE BANK"},
    ])
    res = matching.match_two_tables(L, R, "id", "id", flux_label="TEST")
    assert len(res) == 1
    assert res.iloc[0]["id_R"] == "R1"            # match exact gagne


def test_match_bonus_colonne_secondaire(matching):
    """Avec une 2e colonne (marchandise), le score est boosté."""
    L = _df([{"id": "L1", "NOM_BASE": "NESTLE"}])
    R = _df([{"id": "R1", "NOM_BASE": "NESTLE", "marchandise": "Cacao"}])
    res_seul = matching.match_two_tables(L, R, "id", "id")
    L2 = L.copy(); L2["sec"] = "Cacao"
    res_avec = matching.match_two_tables(L2, R, "id", "id",
                                          secondary_left="sec", secondary_right="marchandise")
    # Avec bonus ne peut pas être inférieur au sans bonus.
    assert res_avec.iloc[0]["score"] >= res_seul.iloc[0]["score"]
