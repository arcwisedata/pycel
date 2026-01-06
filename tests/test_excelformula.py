# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import collections
from unittest import mock

import pytest

from pycel.excelformula import (
    ASTNode,
    ExcelFormula,
    FormulaParserError,
    FunctionNode,
    OperandNode,
    OperatorNode,
    RangeNode,
    Token,
)
from pycel.excelutil import AddressCell


FormulaTest = collections.namedtuple('FormulaTest', 'formula rpn')


def stringify_rpn(e):
    return "|".join([str(x) for x in e])


range_inputs = [
    FormulaTest('=$A1', '$A1'),
    FormulaTest('=$B$2', '$B$2'),
    FormulaTest('=SUM(B5:B15)', 'B5:B15|SUM'),
    FormulaTest('=SUM(B5:B15,D5:D15)', 'B5:B15|D5:D15|SUM'),
    FormulaTest('=SUM(B5:B15 A7:D7)', 'B5:B15|A7:D7| |SUM'),
    FormulaTest('=SUM((A:A,1:1))', 'A:A|1:1|,|SUM'),
    FormulaTest('=SUM((A:A A1:B1))', 'A:A|A1:B1| |SUM'),
    FormulaTest('=SUM(D9:D11,E9:E11,F9:F11)', 'D9:D11|E9:E11|F9:F11|SUM'),
    FormulaTest('=SUM((D9:D11,(E9:E11,F9:F11)))', 'D9:D11|E9:E11|F9:F11|,|,|SUM'),
    FormulaTest('={SUM(B2:D2*B3:D3)}', 'B2:D2|B3:D3|*|SUM|ARRAYROW|ARRAY'),
]

basic_inputs = [
    FormulaTest('=SUM((A:A 1:1))', 'A:A|1:1| |SUM'),
    FormulaTest('=A1', 'A1'),
    FormulaTest('=50', '50'),
    FormulaTest('=1+1', '1|1|+'),
    FormulaTest('=atan2(A1,B1)', 'A1|B1|atan2'),
    FormulaTest('=5*log(sin()+2)', '5|sin|2|+|log|*'),
    FormulaTest('=5*log(sin(3,7,9)+2)', '5|3|7|9|sin|2|+|log|*'),
    FormulaTest('="x"="y"', '"x"|"y"|='),
    FormulaTest('="x"=1', '"x"|1|='),
    FormulaTest('=3 +1-5', '3|1|+|5|-'),
    FormulaTest('=3 + 4 * 5', '3|4|5|*|+'),
    FormulaTest('=+3', '3'),
    FormulaTest('=PI()', 'PI'),
    FormulaTest('=_xlfn.FUNCTION(L45)', 'L45|_xlfn.FUNCTION'),
    FormulaTest('=FLOOR.MATH(L45)', 'L45|FLOOR.MATH'),
    FormulaTest('=100%', '100|%'),
    FormulaTest('=100^100%', '100|100|%|^'),
    FormulaTest('=SUM(B5:B15,D5:D15)%', 'B5:B15|D5:D15|SUM|%'),
    FormulaTest('=AND(G3, 1)', 'G3|1|AND'),
    FormulaTest('=OR(TRUE, TRUE(), FALSE, FALSE())', 'TRUE|TRUE|FALSE|FALSE|OR'),
    FormulaTest('=--4', '4|-|-'),
]

whitespace_inputs = [
    FormulaTest('=3 + 4 * 2 / ( 1 - 5 ) ^ 2 ^ 3', '3|4|2|*|1|5|-|2|^|3|^|/|+'),
    FormulaTest('=1+3+5', '1|3|+|5|+'),
    FormulaTest('=3 * 4 + 5', '3|4|*|5|+'),
    FormulaTest('= (1,5 * (1 + B11 *B3 ^ B12) + 5) + 10 ',
                '1|5|,|1|B11|B3|B12|^|*|+|*|5|+|10|+'),
    FormulaTest('=f(,1)', '|1|f'),
    FormulaTest('=f(1,,)', '1|||f'),
]

if_inputs = [
    FormulaTest(
        '=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &'
        '   "  more ""test"" text"',
        '"a"|"a"|"b"|ARRAYROW|"c"|#N/A|ARRAYROW|1|-|TRUE|ARRAYROW|ARRAY|=|'
        '"yes"|"no"|IF|"  more ""test"" text"|&'),
    FormulaTest('=IF(AI119="","",E119)', 'AI119|""|=|""|E119|IF'),
]

reference_inputs = [
    FormulaTest('=ROW(4:7)', '4:7|ROW'),
    FormulaTest('=ROW(D1:E1)', 'D1:E1|ROW'),
    FormulaTest('=COLUMN(D1:D2)', 'D1:D2|COLUMN'),
    FormulaTest('=ROW(D1:E2)', 'D1:E2|ROW'),
    FormulaTest('=ROW(B53:D54 C54:E54)', 'B53:D54|C54:E54| |ROW'),
    FormulaTest('=COLUMN(L45)', 'L45|COLUMN'),
    FormulaTest('=OFFSET(L45,1,2,3,4)', 'L45|1|2|3|4|OFFSET'),
    FormulaTest('=OFFSET(L45:O50,1,2,,4)', 'L45:O50|1|2||4|OFFSET'),
]

test_names = (
    'range_inputs', 'basic_inputs', 'whitespace_inputs', 'if_inputs',
    'reference_inputs',
)

test_data = []
for test_name in test_names:
    for i, test in enumerate(globals()[test_name]):
        test_data.append((f'{test_name}_{i + 1}', test[0], test[1]))


@pytest.mark.parametrize('test_number, formula, rpn', test_data)
def test_tokenizer(test_number, formula, rpn):
    assert rpn == stringify_rpn(ExcelFormula(formula).rpn)


def test_str():
    excel_formula = ExcelFormula('=E54-E48')
    assert '=E54-E48' == str(excel_formula)


def test_descendants():

    excel_formula = ExcelFormula('=E54-E48')
    descendants = excel_formula.ast.descendants
    assert descendants == excel_formula.ast.descendants

    assert 2 == len(descendants)
    assert 'OPERAND' == descendants[0][0].type
    assert 'OPERAND' == descendants[1][0].type
    assert {'E48', 'E54'} == {
        descendants[0][0].value, descendants[1][0].value
    }


def test_ast_node():
    with pytest.raises(FormulaParserError):
        ASTNode.create(Token('a_value', None, None))

    node = ASTNode(Token('a_value', None, None))
    assert 'ASTNode<a_value>' == repr(node)
    assert 'a_value' == str(node)


@pytest.mark.parametrize(
    'formula', (
        '=if(1',
        '=G11;',
        '=G11,',
        '=(G11;',
        '=;',
        '=,',
        '=-',
    )
)
def test_parser_error(formula):
    with pytest.raises(FormulaParserError):
        ExcelFormula(formula).ast


def test_ast_children():
    excel_formula = ExcelFormula('=SUM(A1, B1)')
    ast = excel_formula.ast

    assert isinstance(ast, FunctionNode)
    assert ast.func_name == 'sum'
    assert len(ast.children) == 2

    for child in ast.children:
        assert isinstance(child, RangeNode)


def test_operator_node():
    excel_formula = ExcelFormula('=A1+B1')
    ast = excel_formula.ast

    assert isinstance(ast, OperatorNode)
    assert ast.operator == '+'
    assert not ast.is_unary


def test_unary_operator_node():
    excel_formula = ExcelFormula('=-A1')
    ast = excel_formula.ast

    assert isinstance(ast, OperatorNode)
    assert ast.is_unary


def test_function_node():
    excel_formula = ExcelFormula('=SUM(A1:B2)')
    ast = excel_formula.ast

    assert isinstance(ast, FunctionNode)
    assert ast.func_name == 'sum'
    assert ast.num_args == 1


def test_operand_node():
    excel_formula = ExcelFormula('=123')
    ast = excel_formula.ast

    assert isinstance(ast, OperandNode)
    assert ast.operand_type == 'NUMBER'


def test_range_node():
    excel_formula = ExcelFormula('=A1:B2')
    ast = excel_formula.ast

    assert isinstance(ast, RangeNode)
    address = ast.get_address()
    assert str(address) == 'A1:B2'


def test_getstate():
    excel_formula = ExcelFormula('=A1+B1')
    _ = excel_formula.ast
    state = excel_formula.__getstate__()

    assert state['base_formula'] == '=A1+B1'
    assert state['_ast'] is None
    assert state['_rpn'] is None


def test_complex_formula_ast():
    formula = '=IF(AND(A1>0, B1<10), SUM(C1:C10), AVERAGE(D1:D10))'
    excel_formula = ExcelFormula(formula)
    ast = excel_formula.ast

    assert isinstance(ast, FunctionNode)
    assert ast.func_name == 'if'
    assert ast.num_args == 3


def test_array_formula():
    formula = '={1,2,3;4,5,6}'
    excel_formula = ExcelFormula(formula)
    ast = excel_formula.ast

    assert isinstance(ast, FunctionNode)


def test_parent_property():
    excel_formula = ExcelFormula('=A1+B1')
    ast = excel_formula.ast

    assert ast.parent is None

    for child in ast.children:
        assert child.parent == ast


def test_token_properties():
    from pycel.excelformula import Tokenizer

    tokenizer = Tokenizer('=SUM(A1:B2)')
    tokens = tokenizer.items

    func_token = next(t for t in tokens if t.type == Token.FUNC)
    assert func_token.is_funcopen

    operand_token = next(t for t in tokens if t.type == Token.OPERAND)
    assert not operand_token.is_operator


def test_token_matches():
    from pycel.excelformula import Tokenizer

    tokenizer = Tokenizer('=A1+B1')
    tokens = tokenizer.items

    operand_token = next(t for t in tokens if t.type == Token.OPERAND)
    assert operand_token.matches(type_=Token.OPERAND)
    assert operand_token.matches(type_=Token.OPERAND, subtype=Token.RANGE)
    assert not operand_token.matches(type_=Token.FUNC)


def test_token_precedence():
    from pycel.excelformula import Tokenizer

    tokenizer = Tokenizer('=A1+B1*C1')
    tokens = tokenizer.items

    plus_token = next(t for t in tokens if t.value == '+')
    mult_token = next(t for t in tokens if t.value == '*')

    assert mult_token.precedence.precedence > plus_token.precedence.precedence
