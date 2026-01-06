# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import os
from unittest import mock

import pytest
from openpyxl.utils import column_index_from_string

from pycel.excelutil import AddressCell
from pycel.excelwrapper import ExcelOpxWrapper as ExcelWrapperImpl


@pytest.fixture(scope='session')
def ATestCell():

    class ATestCell:

        def __init__(self, col, row, sheet='', excel=None, value=None):
            self.row = row
            self.col = col
            self.col_idx = column_index_from_string(col)
            self.sheet = sheet
            self.excel = excel
            self.address = AddressCell(f'{col}{row}', sheet=sheet)
            self.value = value

    return ATestCell


@pytest.fixture(scope='session')
def fixture_dir():
    return os.path.join(os.path.dirname(__file__), 'fixtures')


@pytest.fixture(scope='session')
def fixture_xls_path(fixture_dir):
    return os.path.join(fixture_dir, 'excelcompiler.xlsx')


@pytest.fixture(scope='session')
def unconnected_excel(fixture_xls_path):
    import openpyxl.worksheet._reader as orw
    old_warn = orw.warn

    def new_warn(msg, *args, **kwargs):
        if 'Unknown' not in msg:
            old_warn(msg, *args, **kwargs)

    with mock.patch('openpyxl.worksheet._reader.warn', new_warn):
        yield ExcelWrapperImpl(fixture_xls_path)


@pytest.fixture
def excel(unconnected_excel):
    unconnected_excel.load()
    return unconnected_excel


@pytest.fixture
def cond_format_ws(fixture_dir, ATestCell):
    path = os.path.join(fixture_dir, 'cond-format.xlsx')
    excel = ExcelWrapperImpl(path)
    excel.load()
    return ATestCell('A', 1, excel=excel)
