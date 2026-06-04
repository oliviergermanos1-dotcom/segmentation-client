# -*- coding: utf-8 -*-
"""Tests du module 01_normalise (Cleanco + OHADA + aliases)."""


def test_normalise_cargil(norm):
    cleaned, base = norm.normalise("SARL CARGIL COTE D'IVOIRE")
    assert "CARGIL" in cleaned
    assert "SARL" not in cleaned                  # forme juridique retirée
    assert base == "CARGIL"                       # géo retiré + articles


def test_normalise_kouassi(norm):
    cleaned, base = norm.normalise("Ets KOUASSI & Frères SA")
    assert "ETS" not in cleaned and "SA" not in cleaned
    assert "KOUASSI" in base and "FRERES" in base


def test_normalise_accents_et_ponctuation(norm):
    cleaned, _ = norm.normalise("NESTLÉ CÔTE D'IVOIRE")
    assert "NESTLE" in cleaned                    # accents normalisés
    assert "'" not in cleaned                     # ponctuation supprimée


def test_normalise_vide(norm):
    assert norm.normalise("") == ("", "")
    assert norm.normalise(None) == ("", "")


def test_normalise_articles_strip(norm):
    _, base = norm.normalise("Cie Fruitière de l'Ouest GIE")
    # GIE + Cie retirés, articles "de l" retirés.
    assert "GIE" not in base
    assert "DE" not in base.split()
    assert "FRUITIERE" in base


def test_alias_rubriks(norm):
    # L'alias par défaut dans aliases.json : SIR → Société Ivoirienne de Raffinage
    norm.load_aliases.__globals__["_ALIASES_CACHE"] = None  # reset cache
    out = norm.apply_alias("SIR", "RUBRIKS")
    assert "Ivoirienne" in out or "IVOIRIENNE" in out.upper()


def test_alias_inconnu(norm):
    # Un nom non aliasé ressort tel quel.
    assert norm.apply_alias("XYZ_INCONNU", "RUBRIKS") == "XYZ_INCONNU"


def test_alias_source_inexistante(norm):
    assert norm.apply_alias("X", "INEXISTANT") == "X"
