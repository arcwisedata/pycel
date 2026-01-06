# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import os
import pickle
from collections import namedtuple

import pytest
from openpyxl.utils import quote_sheetname

from pycel.excelutil import (
    AddressCell,
    AddressMultiAreaRange,
    AddressRange,
    flatten,
    is_address,
    list_like,
    MAX_COL,
    MAX_ROW,
    NULL_ERROR,
    PyCelException,
    range_boundaries,
    split_sheetname,
    structured_reference_boundaries,
    uniqueify,
    unquote_sheetname,
    VALUE_ERROR,
)


def test_address_range():
    a = AddressRange('a1:b2')
    b = AddressRange('A1:B2')
    c = AddressRange(a)

    assert a == b
    assert b == c

    assert b == AddressRange(b)
    assert b == AddressRange.create(b)

    assert AddressRange('sh!a1:b2') == AddressRange(a, sheet='sh')
    assert AddressCell('C13') == AddressCell('R13C3')

    with pytest.raises(ValueError):
        AddressRange(AddressRange('sh!a1:b2'), sheet='sheet')

    a = AddressRange('A:A')
    assert 'A' == a.start.column
    assert 'A' == a.end.column
    assert 0 == a.start.row
    assert 0 == a.end.row

    b = AddressRange('1:1')
    assert '' == b.start.column
    assert '' == b.end.column
    assert 1 == b.start.row
    assert 1 == b.end.row

    c = b.start
    assert c.start == c.end == c


def test_address_range_errors():

    with pytest.raises(ValueError):
        AddressRange('B32:B')

    with pytest.raises(ValueError):
        AddressRange('B32:B33:B')


@pytest.mark.parametrize(
    'address, expected', (
        ('s!D2:F4:E3', 's!D2:F4'),
        ('s!D2:E3:F4', 's!D2:F4'),
        ('s!E3:D2:F4', 's!D2:F4'),
        ('s!D2:F4:G3', 's!D2:G4'),
        ('s!D2:G3:F4', 's!D2:G4'),
        ('s!G3:D2:F4', 's!D2:G4'),
        ('s!G3:G3:G3', 's!G3'),
    )
)
def test_address_range_multi_colon(address, expected):
    a_range = AddressRange(address)
    assert a_range == AddressRange(expected)


@pytest.mark.parametrize(
    'left, right, expected', (
        ('a1', 'a1', 'a1'),
        ('a1:b2', 'b1:c3', 'b1:b2'),
        ('a1:d5', 'b3', 'b3'),
        ('d4:e5', 'c3', NULL_ERROR),
        ('d4:e5', 'd3', NULL_ERROR),
        ('d4:e5', 'e3', NULL_ERROR),
        ('d4:e5', 'f3', NULL_ERROR),
        ('d4:e5', 'c4', NULL_ERROR),
        ('d4:e5', 'd4', 'd4'),
        ('d4:e5', 'e4', 'e4'),
        ('d4:e5', 'f4', NULL_ERROR),
        ('d4:e5', 'c5', NULL_ERROR),
        ('d4:e5', 'd5', 'd5'),
        ('d4:e5', 'e5', 'e5'),
        ('d4:e5', 'f5', NULL_ERROR),
        ('d4:e5', 'c6', NULL_ERROR),
        ('d4:e5', 'd6', NULL_ERROR),
        ('d4:e5', 'e6', NULL_ERROR),
        ('d4:e5', 'f6', NULL_ERROR),
        ('c4:e5', 'd1', NULL_ERROR),
        ('c4:e6', 'a5', NULL_ERROR),
        ('c4:e6', 's!a5', NULL_ERROR),
        ('s!c4:e6', 'a5', NULL_ERROR),
        ('s!c4:e6', 's!a5', NULL_ERROR),
        ('s!c4:e6', 't!a5', VALUE_ERROR),
        ('d4:e5', 's!e5', 's!e5'),
        ('s!d4:e5', 'e5', 's!e5'),
        ('s!d4:e5', 's!e5', 's!e5'),
        ('s!d4:e5', 't!e5', VALUE_ERROR),
    )
)
def test_address_range_intersection(left, right, expected):
    expected = AddressRange(expected)
    assert AddressRange(left) & AddressRange(right) == expected
    assert AddressRange(left) & right == expected
    assert left & AddressRange(right) == expected


@pytest.mark.parametrize(
    'left, right, expected', (
        ('a1', 'a1', 'a1'),
        ('a1:b2', 'b1:c3', 'a1:c3'),
        ('a1:b2', 'd5', 'a1:d5'),
        ('a1:d5', 'b3', 'a1:d5'),
        ('d4:e5', 'a1', 'a1:e5'),
        ('c4:e5', 'd1', 'c1:e5'),
        ('c4:e6', 'a5', 'a4:e6'),
        ('c4:e5', 'd9', 'c4:e9'),
        ('c4:e6', 'j5', 'c4:j6'),
        ('c4:e6', 's!a5', 's!a4:e6'),
        ('s!c4:e6', 'a5', 's!a4:e6'),
        ('s!c4:e6', 's!a5', 's!a4:e6'),
        ('s!c4:e6', 't!a5', VALUE_ERROR),
    )
)
def test_address_range_union(left, right, expected):
    expected = AddressRange(expected)
    assert AddressRange(left) ** AddressRange(right) == expected
    assert AddressRange(left) ** right == expected
    assert left ** AddressRange(right) == expected


@pytest.mark.parametrize(
    'a_range, address, expected', (
        ('s!D2:F4', 's!D2', True),
        ('s!D2:F4', 's!F2', True),
        ('s!D2:F4', 's!D4', True),
        ('s!D2:F4', 's!F4', True),
        ('s!D2:F4', 's!C2', False),
        ('s!D2:F4', 's!D1', False),
        ('s!D2:F4', 's!G4', False),
        ('s!D2:F4', 's!F5', False),
    )
)
def test_address_range_contains(a_range, address, expected):
    a_range = AddressRange(a_range)
    assert expected == (address in a_range)
    address = AddressCell(address)
    assert expected == (address in a_range)
    assert address in address


def test_is_range():

    assert AddressRange('a1:b2').is_range
    assert not AddressRange('a1').is_range


def test_has_sheet():

    assert AddressRange('Sheet1!a1').has_sheet
    assert not AddressRange('a1').has_sheet
    assert AddressRange('Sheet1!a1:b2').has_sheet
    assert not AddressRange('a1:b2').has_sheet

    assert AddressCell('sh!A2') == AddressRange(AddressRange('A2'), sheet='sh')

    with pytest.raises(ValueError, match='Mismatched sheets'):
        AddressRange(AddressRange('shx!a1'), sheet='sh')


def test_address_range_size():

    assert (1, 1) == AddressRange('B1').size
    assert (1, 2) == AddressRange('B1:C1').size
    assert (2, 1) == AddressRange('B1:B2').size
    assert (2, 2) == AddressRange('B1:C2').size

    assert (MAX_ROW, 2) == AddressRange('B:C').size
    assert (3, MAX_COL) == AddressRange('2:4').size


def test_address_cell_addr_inc():

    cell_addr = AddressCell('sh!C2')

    assert MAX_COL - 1 == cell_addr.inc_col(-4)
    assert MAX_COL == cell_addr.inc_col(-3)
    assert 1 == cell_addr.inc_col(-2)
    assert 5 == cell_addr.inc_col(2)
    assert 6 == cell_addr.inc_col(3)

    assert MAX_ROW - 1 == cell_addr.inc_row(-3)
    assert MAX_ROW == cell_addr.inc_row(-2)
    assert 1 == cell_addr.inc_row(-1)
    assert 5 == cell_addr.inc_row(3)
    assert 6 == cell_addr.inc_row(4)


@pytest.mark.parametrize(
    'row_inc, col_inc, expected', (
        (-3, -4, 'sh!XFC1048575'),
        (-2, -3, 'sh!XFD1048576'),
        (-1, -2, 'sh!A1'),
        (3, 2, 'sh!E5'),
        (4, 3, 'sh!F6'),
    )
)
def test_address_cell_addr_offset(row_inc, col_inc, expected):
    assert AddressCell('sh!C2').address_at_offset(row_inc, col_inc).address == expected
    assert AddressRange('sh!C2:D3').address_at_offset(row_inc, col_inc).address == expected


def test_address_sort_keys():

    a1_b2 = AddressRange('sh!A1:B2')
    a1 = AddressRange('sh!A1')
    b2 = AddressRange('sh!B2')

    assert a1.sort_key == a1_b2.sort_key
    assert a1.sort_key < b2.sort_key


def test_address_range_columns():
    columns = list(list(x) for x in AddressRange('sh!A1:C3').cols)
    assert 3 == len(columns)
    assert 3 == len(columns[0])

    assert all('A' == addr.column for addr in columns[0])
    assert all('C' == addr.column for addr in columns[-1])


def test_address_pickle(tmpdir):
    addrs = [
        AddressRange('B1'),
        AddressRange('B1:C1'),
        AddressRange('B1:B2'),
        AddressRange('B1:C2'),
        AddressRange('sh!B1'),
        AddressRange('sh!B1:C1'),
        AddressRange('sh!B1:B2'),
        AddressRange('sh!B1:C2'),
        AddressRange('B:C'),
        AddressRange('2:4'),
        AddressCell('sh!XFC1048575'),
        AddressCell('sh!XFD1048576'),
        AddressCell('sh!A1'),
        AddressCell('sh!E5'),
        AddressCell('sh!F6'),
    ]

    filename = os.path.join(str(tmpdir), 'test_addrs.pkl')
    with open(filename, 'wb') as f:
        pickle.dump(addrs, f)

    with open(filename, 'rb') as f:
        new_addrs = pickle.load(f)

    assert addrs == new_addrs


@pytest.mark.parametrize(
    'sheet_name',
    [
        u'In Dusseldorf',
        u'My-Sheet',
        u"Demande d'autorisation",
        "1sheet",
        ".sheet",
        '"',
    ]
)
def test_unquote_sheetname(sheet_name):
    assert sheet_name == unquote_sheetname(quote_sheetname(sheet_name))


@pytest.mark.parametrize(
    'sheet_name',
    [
        u'In Dusseldorf',
        u'My-Sheet',
        u"Demande d'autorisation",
        "1sheet",
        ".sheet",
        '"',
    ]
)
def test_quoted_address(sheet_name):
    addr = AddressCell('A2', sheet=sheet_name)
    assert addr.quoted_address == f'{addr.quote_sheet(sheet_name)}!A2'


@pytest.mark.parametrize(
    'address, expected', (
        ('s!D2', 's!$D$2'),
        ('s!D2:F4', 's!$D$2:$F$4'),
        (AddressRange("D2:F4", sheet='sh 1'), "'sh 1'!$D$2:$F$4"),
    )
)
def test_address_absolute(address, expected):
    assert AddressRange.create(address).abs_address == expected


def test_split_sheetname():

    assert ('', 'B1') == split_sheetname('B1')
    assert ('sheet', 'B1') == split_sheetname('sheet!B1')
    assert ('', 'B1:C2') == split_sheetname('B1:C2')
    assert ('sheet', 'B1:C2') == split_sheetname('sheet!B1:C2')

    assert ('sheet', 'B1:C2') == split_sheetname('sheet!B1:C2')
    assert ('SheetA', 'A1:A9') == split_sheetname("SheetA!A1:'SheetA'!A9")

    assert ("shee't", 'B1:C2') == split_sheetname("'shee''t'!B1:C2")
    assert ("Sheet' A", 'A1:A9') == split_sheetname("'Sheet'' A'!A1:'Sheet'' A'!A9")

    assert ("shee t", 'B1:C2') == split_sheetname("'shee t'!B1:C2")

    with pytest.raises(ValueError):
        split_sheetname('sh!B1', sheet='shx')

    with pytest.raises(NotImplementedError):
        split_sheetname('sh!B1:C2:sh2!B1:C2')


def test_address_cell_enum(ATestCell):
    assert ('B1', '', 2, 1, 'B1') == AddressCell('B1')
    assert ('sheet!B1', 'sheet', 2, 1, 'B1') == AddressCell('sheet!B1')

    assert ('A1', '', 1, 1, 'A1') == AddressCell('R1C1')
    assert ('sheet!A1', 'sheet', 1, 1, 'A1') == AddressCell('sheet!R1C1')

    cell = ATestCell('A', 1)
    assert ('B2', '', 2, 2, 'B2') == AddressCell.create(
        'R[1]C[1]', cell=cell)
    assert ('sheet!B2', 'sheet', 2, 2, 'B2') == AddressCell.create(
        'sheet!R[1]C[1]', cell=cell)

    with pytest.raises(ValueError):
        AddressCell('B1:C2')

    with pytest.raises(ValueError):
        AddressCell('sheet!B1:C2')

    with pytest.raises(ValueError):
        AddressCell('xyzzy')


def test_resolve_range():
    a = AddressRange.create

    assert ((a('B1'), ), ) == a('B1').resolve_range
    assert ((a('B1'), a('C1')),) == a('B1:C1').resolve_range
    assert ((a('B1'),), (a('B2'), )) == a('B1:B2').resolve_range
    assert ((a('B1'), a('C1')), (a('B2'), a('C2'))) == a('B1:C2').resolve_range

    assert ((a('sh!B1'),),) == a('sh!B1').resolve_range
    assert ((a('sh!B1'), a('sh!C1')),) == a('sh!B1:C1').resolve_range
    assert ((a('sh!B1'),), (a('sh!B2'),)) == a('sh!B1:B2').resolve_range
    assert ((a('sh!B1'), a('sh!C1')),
            (a('sh!B2'), a('sh!C2'))) == (a('sh!B1:C2')).resolve_range

    assert ((a('sh!B1'),),) == a('sh!B1', sheet='sh').resolve_range
    assert ((a('sh!B1'), a('sh!C1')),) == (
        a('sh!B1:C1', sheet='sh')).resolve_range
    assert ((a('sh!B1'),), (a('sh!B2'),)) == (
        a('sh!B1:B2', sheet='sh')).resolve_range
    assert ((a('sh!B1'), a('sh!C1')), (a('sh!B2'), a('sh!C2'))) == \
        (a('sh!B1:C2', sheet='sh')).resolve_range

    with pytest.raises(AssertionError):
        a('B:C').resolve_range

    with pytest.raises(AssertionError):
        a('1:2').resolve_range


addr_cr = AddressRange.create


@pytest.mark.parametrize(
    'address, string, mar', (
        (((addr_cr('B1'), ), (addr_cr('B1'), addr_cr('C1'),)),
         'B1,B1:C1',
         AddressMultiAreaRange((addr_cr('B1'), addr_cr('B1:C1')))),
        (((addr_cr('B1'),), (addr_cr('B2'),),
          (addr_cr('B1'), addr_cr('C1')), (addr_cr('B2'), addr_cr('C2'))),
         'B1:B2,B1:C2',
         AddressMultiAreaRange((addr_cr('B1:B2'), addr_cr('B1:C2')))),
    )
)
def test_multi_area_range(address, string, mar):
    assert address == tuple(mar.resolve_range)
    assert not mar.is_unbounded_range
    assert address[0][0] in mar
    assert AddressRange('Z99') not in mar
    assert str(mar) == string


@pytest.mark.parametrize(
    'ref, expected', (
        ('a_table[[#This Row], [col5]]', 'E5'),
        ('a_table[[#All],[col3]]', 'C1:C8'),
        ('a_table[[#All],[col3]:[col4]]', 'C1:D8'),
        ('a_table[[#Headers],[col4]]', 'D1'),
        ('a_table[[#Headers],[col2]:[col5]]', 'B1:E1'),
        ('a_table[[#Headers],[#Data],[col4]]', PyCelException('D1:D7')),
        ('a_table[[#Data],[col4]:[col4]]', 'D2:D7'),
        ('a_table[[#Data],[col4]:[col5]]', 'D2:E7'),
        ('a_table[[#Totals],[col2]]', 'B8'),
        ('a_table[[#Totals],[col3]:[col5]]', 'C8:E8'),
        ('a_table[[#This Row], [col5]]', 'E5'),
        ('a_table[[col4]:[col4]]', 'D2:D7'),
        ('a_table[@col5]', 'E5'),
        ('a_table[@[col2]]', 'B5'),
        ('a_table[#This Row]', 'A5:E5'),
        ('a_table[@]', 'A5:E5'),
        ('a_table[]', 'A2:E7'),
        ('JUNK[]', PyCelException()),
        ('a_table[]', None),
        ('a_table[[#JUNK]]', PyCelException()),
        ('a_table[[#Data],[JUNK]]', PyCelException()),
        ('a_table[[#Data],[JUNK]:[col4]]', PyCelException()),
        ('a_table[[#Data],[col5]:[col4]]', PyCelException()),
        ('a_table[[]', PyCelException()),
        ('a_table[[[col4]:[col4]]', PyCelException()),
    )
)
def test_structured_table_reference_boundaries(ref, expected):

    Column = namedtuple('Column', 'name')

    class Table:
        def __init__(self, ref, header_rows, totals_rows):
            self.ref = ref
            self.headerRowCount = header_rows
            self.totalsRowCount = totals_rows
            self.tableColumns = tuple(
                Column(name) for name in 'col1 col2 col3 col4 col5'.split())

    class Excel:
        def __init__(self, table):
            self.a_table = table

        def table(self, name):
            if name == 'a_table':
                return self.a_table, None
            else:
                return None, None

    class Cell:
        def __init__(self, table, address):
            self.excel = Excel(table)
            self.address = AddressCell(address)

    cell = Cell(Table('A1:E8', 1, 1), 'E5')

    if isinstance(expected, PyCelException):
        with pytest.raises(PyCelException):
            structured_reference_boundaries(ref, cell=cell)

    elif expected is None:
        with pytest.raises(PyCelException):
            structured_reference_boundaries(ref, cell=None)

    else:
        ref_bound = structured_reference_boundaries(ref, cell=cell)
        expected_bound = range_boundaries(expected, cell=cell)
        assert ref_bound == expected_bound

        expected_ref = range_boundaries(ref, cell=cell)
        assert ref_bound == expected_ref


@pytest.mark.parametrize(
    'expected, address', (

        ((1, 2) * 2, 'A2'),
        ((2, 1) * 2, 'B1'),
        ((1, 2) * 2, 'R2C1'),
        ((2, 1) * 2, 'R1C2'),
        ((2, 3) * 2, 'R[2]C[1]'),
        ((3, 2) * 2, 'R[1]C[2]'),

        ((1, 1, 2, 2), 'A1:B2'),
        ((1, 1, 2, 2), 'R1C1:R2C2'),
        ((2, 1, 2, 3), 'R1C2:R[2]C[1]'),

        ((3, 13) * 2, 'R13C3'),

        ((1, 1, 1, 1), 'RC'),

        ((None, 1, None, 4), 'R:R[3]'),
        ((None, 1, None, 4), 'R1:R[3]'),
        ((None, 2, None, 4), 'R2:R[3]'),

        ((1, None, 4, None), 'C:C[3]'),
        ((1, None, 4, None), 'C1:C[3]'),
        ((2, None, 4, None), 'C2:C[3]'),

        ((4, 2, 6, 4), 's!D2:F4:E3'),
        ((4, 2, 6, 4), 's!D2:E3:F4'),
        ((4, 2, 6, 4), 's!E3:D2:F4'),
        ((4, 2, 7, 4), 's!D2:F4:G3'),
        ((4, 2, 7, 4), 's!D2:G3:F4'),
        ((4, 2, 7, 4), 's!G3:D2:F4'),
        ((7, 3, 7, 3), 's!G3:G3:G3'),
    )
)
def test_extended_range_boundaries(expected, address, ATestCell):
    assert range_boundaries(address, cell=ATestCell('A', 1))[0] == expected


def test_range_boundaries_defined_names(excel, ATestCell):
    cell = ATestCell('A', 1, excel=excel)

    assert ((3, 1, 3, 18), 'Sheet1') == range_boundaries('SINUS', cell)
    assert ((2, 1, 5, 18), 'Sheet1') == range_boundaries('B2:E5:SINUS', cell)


@pytest.mark.parametrize(
    'address_string',
    [
        'R',
        'C',
        ':',
        'R:',
        'C:',
        ':R',
        ':C',
        'RC:',
        ':RC',
        'R:C1',
        'C:R1',
        'C1:RC',
        'R1:RC',
        'RC:R1',
        'RC:C1',
        'sheet!B1',
        'xyzzy',
    ]
)
def test_extended_range_boundaries_errors(address_string, ATestCell):
    cell = ATestCell('A', 1)

    with pytest.raises(ValueError, match='not a valid coordinate or range'):
        range_boundaries(address_string, cell)


def test_multi_area_ranges(excel, ATestCell):
    cell = ATestCell('A', 1, excel=excel)
    from unittest import mock
    with mock.patch.object(excel, '_defined_names', {
            'dname': (('$A$1', 's1'), ('$A$3:$A$4', 's2'))}):

        multi_area_range = AddressMultiAreaRange(
            tuple(AddressRange(addr, sheet=sh))
            for addr, sh in excel._defined_names['dname'])

        assert (multi_area_range, None) == range_boundaries('dname', cell)
        assert multi_area_range == AddressRange.create('dname', cell=cell)


def test_flatten():
    assert ['ddd'] == list(flatten(['ddd']))
    assert ['ddd', 1, 2, 3] == list(flatten(['ddd', 1, (2, 3)]))
    assert ['ddd', 1, 2, 3] == list(flatten(['ddd', (1, (2, 3))]))
    assert ['ddd', 1, 2, 3] == list(flatten(['ddd', (1, 2), 3]))

    assert [None] == list(flatten(None))
    assert [True] == list(flatten(True))
    assert [1.0] == list(flatten(1.0))


def test_uniqueify():
    assert (1, 2, 3, 4) == uniqueify((1, 2, 3, 4, 3))
    assert (4, 1, 2, 3) == uniqueify((4, 1, 2, 3, 4, 3))


@pytest.mark.parametrize(
    'data, expected', (
        (AddressCell('A1'), True),
        (AddressRange('A1:B2'), True),
        ('A1', False),
        ('A1:B2', False),

        (1, False),
        (0, False),
        (-1, False),
        (1.0, False),
        ('-1.0', False),
        (True, False),
        (False, False),
        (None, False),
        ('x', False),
    )
)
def test_is_address(data, expected):
    assert is_address(data) == expected


@pytest.mark.parametrize(
    'value, expected', (
        ('xyzzy', False),
        (AddressRange('A1:B2'), False),
        (AddressCell('A1'), False),
        ([1, 2], True),
        ((1, 2), True),
        ({1: 2, 3: 4}, True),
        ((a for a in range(2)), True),
    )
)
def test_list_like(value, expected):
    assert list_like(value) == expected
