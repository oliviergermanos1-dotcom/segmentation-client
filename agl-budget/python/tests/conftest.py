# -*- coding: utf-8 -*-
"""Fixtures pytest partagées — pipeline AGL BUDGET."""
import importlib.util
import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))   # rend `config` importable depuis les scripts


def _mod(name, filename):
    spec = importlib.util.spec_from_file_location(name, ROOT / filename)
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    return m


@pytest.fixture(scope="session")
def norm():    return _mod("norm", "01_normalise.py")
@pytest.fixture(scope="session")
def m02():     return _mod("m02", "02_match_crm_iris.py")
@pytest.fixture(scope="session")
def m03():     return _mod("m03", "03_match_crm_stat.py")
@pytest.fixture(scope="session")
def m05():     return _mod("m05", "05_build_rmc.py")
@pytest.fixture(scope="session")
def m06():     return _mod("m06", "06_agreg_statcom.py")
@pytest.fixture(scope="session")
def m07():     return _mod("m07", "07_calcul_pdm.py")
@pytest.fixture(scope="session")
def m08():     return _mod("m08", "08_projections.py")
@pytest.fixture(scope="session")
def m09():     return _mod("m09", "09_whitespaces.py")
@pytest.fixture(scope="session")
def m10():     return _mod("m10", "10_extract_pnd.py")
@pytest.fixture(scope="session")
def m11():     return _mod("m11", "11_budget_real.py")
@pytest.fixture(scope="session")
def matching():
    import matching_lib
    return matching_lib


@pytest.fixture
def tmp_outdir(tmp_path):
    """Dossier de sortie temporaire pour les scripts qui écrivent des CSV/XLSX."""
    d = tmp_path / "out"
    d.mkdir()
    return d
