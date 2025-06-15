import os

import pytest
from py.path import local as LocalPath


@pytest.fixture
def datadir():
    """DATADIR as a LocalPath"""
    here = os.path.split(__file__)[0]
    data_dir = os.path.join(here, "data")
    return LocalPath(data_dir)
