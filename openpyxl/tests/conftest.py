# Fixtures (pre-configured objects) for tests
import os

import pytest
from py.path import LocalPath


@pytest.fixture
def datadir():
    """DATADIR as a LocalPath"""
    here = os.path.split(__file__)[0]
    data_dir = os.path.join(here, "data")
    return LocalPath(data_dir)


# @pytest.fixture
# def ws(Workbook):
#     """Empty worksheet titled 'data'"""
#     wb = Workbook()
#     ws = wb.active
#     ws.title = "data"
#     return ws


# objects under test
# @pytest.fixture
# def Image():
#     """Image class"""
#     from openpyxl.drawing import Image
#
#     return Image
