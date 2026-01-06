# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Excel formula parsing module.

This module provides AST parsing for Excel formulas WITHOUT any code
compilation or evaluation capabilities. The evaluation code has been
removed for security reasons.
"""

import logging

import openpyxl.formula.tokenizer as tokenizer
from networkx.classes.digraph import DiGraph
from networkx.exception import NetworkXError

from pycel.excelutil import (
    AddressRange,
    NAME_ERROR,
    PyCelException,
)


class FormulaParserError(PyCelException):
    """Error during parsing"""


class Tokenizer(tokenizer.Tokenizer):
    """Amend openpyxl tokenizer"""

    def __init__(self, formula):
        super(Tokenizer, self).__init__(formula)
        self.items = self._items()

    def _items(self):
        """Convert to use our Token"""
        t = [None] + [Token.from_token(t) for t in self.items] + [None]

        # convert or remove unneeded whitespace
        tokens = []
        for prev_token, token, next_token in zip(t, t[1:], t[2:]):
            if token.type != Token.WSPACE or not prev_token or not next_token:
                # ::HACK:: this is code to make the tokenizer behave like
                # this change to the openpyxl tokenizer.
                # https://bitbucket.org/openpyxl/openpyxl/pull-requests/345
                # If the pull request gets merged, we can the update our
                # openpyxl requirements and remove this code.
                if (token.matches(type_=Token.FUNC, subtype=Token.OPEN) and
                        ':' in token.value):

                    # split the address on the ':'
                    addr, func = token.value.rsplit(':', maxsplit=1)
                    tokens.append(Token(addr, Token.OPERAND, Token.RANGE))
                    tokens.append(Token(':', Token.OP_IN, ''))
                    token.value = func
                    tokens.append(token)

                elif (token.matches(type_=Token.OPERAND,
                                    subtype=Token.RANGE) and
                      token.value.startswith(':')):
                    # split the address on the ':'
                    tokens.append(Token(':', Token.OP_IN, ''))
                    token.value = token.value[1:]
                    tokens.append(token)

                # drop unary +
                elif not token.matches(type_=Token.OP_PRE, value='+'):
                    tokens.append(token)

            elif (
                prev_token.matches(type_=Token.FUNC, subtype=Token.CLOSE) or
                prev_token.matches(type_=Token.PAREN, subtype=Token.CLOSE) or
                prev_token.type == Token.OPERAND
            ) and (
                next_token.matches(type_=Token.FUNC, subtype=Token.OPEN) or
                next_token.matches(type_=Token.PAREN, subtype=Token.OPEN) or
                next_token.type == Token.OPERAND
            ):
                # this whitespace is an intersect operator
                tokens.append(Token(token.value, Token.OP_IN, Token.INTERSECT))

        return tokens


class Token(tokenizer.Token):
    """Amend openpyxl token"""

    INTERSECT = "INTERSECT"
    ARRAYROW = "ARRAYROW"
    EMPTY = "EMPTY"

    class Precedence:
        """Small wrapper class to manage operator precedence during parsing"""

        def __init__(self, precedence, associativity):
            self.precedence = precedence
            self.associativity = associativity

        def __lt__(self, other):
            return (self.precedence < other.precedence or
                    self.associativity == "left" and
                    self.precedence == other.precedence
                    )

    precedences = {
        # http://office.microsoft.com/en-us/excel-help/
        #   calculation-operators-and-precedence-HP010078886.aspx
        ':': Precedence(8, 'left'),
        ' ': Precedence(8, 'left'),  # range intersection
        ',': Precedence(8, 'left'),
        'u': Precedence(7, 'right'),  # unary operator
        '%': Precedence(6, 'left'),
        '^': Precedence(5, 'left'),
        '*': Precedence(4, 'left'),
        '/': Precedence(4, 'left'),
        '+': Precedence(3, 'left'),
        '-': Precedence(3, 'left'),
        '&': Precedence(2, 'left'),
        '=': Precedence(1, 'left'),
        '<': Precedence(1, 'left'),
        '>': Precedence(1, 'left'),
        '<=': Precedence(1, 'left'),
        '>=': Precedence(1, 'left'),
        '<>': Precedence(1, 'left'),
    }

    @classmethod
    def from_token(cls, token, value=None, type_=None, subtype=None):
        return cls(
            token.value if value is None else value,
            token.type if type_ is None else type_,
            token.subtype if subtype is None else subtype
        )

    @property
    def is_operator(self):
        return self.type in (Token.OP_PRE, Token.OP_IN, Token.OP_POST)

    @property
    def is_funcopen(self):
        return self.subtype == Token.OPEN and self.type in (
            Token.FUNC, Token.ARRAY, Token.ARRAYROW)

    def matches(self, type_=None, subtype=None, value=None):
        return ((type_ is None or self.type == type_) and
                (subtype is None or self.subtype == subtype) and
                (value is None or self.value == value))

    @property
    def precedence(self):
        assert self.is_operator
        return self.precedences[
            'u' if self.type == Token.OP_PRE else self.value]


class ASTNode:
    """A generic node in the AST used to represent a cell's formula"""

    def __init__(self, token, cell=None):
        super(ASTNode, self).__init__()
        self.token = token
        self.cell = cell
        self._ast = None
        self._parent = None
        self._children = None
        self._descendants = None

    @classmethod
    def create(cls, token, cell=None):
        """Simple factory function"""
        if token.type == Token.OPERAND:
            if token.subtype == Token.RANGE:
                return RangeNode(token, cell)
            else:
                return OperandNode(token, cell)

        elif token.is_funcopen:
            return FunctionNode(token, cell)

        elif token.is_operator:
            return OperatorNode(token, cell)

        raise FormulaParserError(f'Unknown token type: {repr(token)}')

    def __str__(self):
        return str(self.token.value.strip('('))

    def __repr__(self):
        return f"{type(self).__name__}<{self.token.value.strip('(')}>"

    @property
    def ast(self):
        return self._ast

    @ast.setter
    def ast(self, value):
        self._ast = value

    @property
    def value(self):
        return self.token.value

    @property
    def type(self):
        return self.token.type

    @property
    def subtype(self):
        return self.token.subtype

    @property
    def children(self):
        if self._children is None:
            try:
                args = self.ast.predecessors(self)
            except NetworkXError:
                args = []
            self._children = sorted(
                args, key=lambda x: self.ast.nodes[x]['pos'])
        return self._children

    @property
    def descendants(self):
        if self._descendants is None:
            self._descendants = list(
                n for n in self.ast.nodes(self) if n[0] != self)
        return self._descendants

    @property
    def parent(self):
        if self._parent is None:
            self._parent = next(self.ast.successors(self), None)
        return self._parent


class OperatorNode(ASTNode):
    """AST node representing an operator (+, -, *, /, etc.)"""

    op_map = {
        # convert the operator to python equivalents
        "^": "**",
        "=": "==",
        "<>": "!=",
    }

    @property
    def operator(self):
        """Return the operator symbol"""
        return self.op_map.get(self.value, self.value)

    @property
    def is_unary(self):
        """Check if this is a unary operator"""
        return self.type == Token.OP_PRE


class OperandNode(ASTNode):
    """AST node representing an operand (number, string, boolean, etc.)"""

    @property
    def operand_type(self):
        """Return the type of operand"""
        return self.subtype


class RangeNode(OperandNode):
    """Represents a spreadsheet cell or range, e.g., A5 or B3:C20"""

    def get_address(self):
        """Get the AddressRange for this node"""
        sheet = self.cell and self.cell.sheet or ''
        value = self.value
        if '!' in value:
            sheet = ''
        try:
            addr_str = value.replace('$', '')
            address = AddressRange.create(addr_str, sheet=sheet, cell=self.cell)
            return address
        except ValueError:
            # check for table relative address
            table_name = None
            if self.cell:
                excel = self.cell.excel
                if excel and '[' in addr_str:
                    table_name = excel.table_name_containing(self.cell.address)

            if not table_name:
                logging.getLogger('pycel').warning(f'Table Name not found: {addr_str}')
                return NAME_ERROR

            addr_str = f'{table_name}{addr_str}'
            address = AddressRange.create(
                addr_str, sheet=self.cell.address.sheet, cell=self.cell)
            return address


class FunctionNode(ASTNode):
    """AST node representing a function call"""

    def __init__(self, *args):
        super(FunctionNode, self).__init__(*args)
        self.num_args = 0

    @property
    def func_name(self):
        """Get the function name (normalized)"""
        func = self.value.lower().strip('(')
        if func and func[0] == func[-1] == '_':
            func = func.upper()
        if func.startswith('_xlfn.'):
            func = func[6:]
        func = func.replace('.', '_')
        return func


class ExcelFormula:
    """Parse an Excel formula into an AST.
    
    This class provides ONLY parsing functionality. Code compilation
    and evaluation have been removed for security reasons.
    """

    def __init__(self, formula, cell=None):
        self.base_formula = formula
        self.cell = cell

        self._rpn = None
        self._ast = None

    def __str__(self):
        return self.base_formula

    def __repr__(self):
        return f'ExcelFormula({self.base_formula})'

    def __getstate__(self):
        # Throw everything away except the base formula
        state = dict(self.__dict__)
        for to_remove in ('_ast', '_rpn'):
            if to_remove in state:
                state[to_remove] = None
        return state

    @property
    def rpn(self):
        """Parse formula to Reverse Polish Notation"""
        if self._rpn is None:
            self._rpn = self._parse_to_rpn(self.base_formula)
        return self._rpn

    @property
    def ast(self):
        """Build and return the Abstract Syntax Tree"""
        if self._ast is None and self.rpn:
            self._ast = self._build_ast(self.rpn)
        return self._ast

    def _ast_node(self, token):
        return ASTNode.create(token, self.cell)

    def _parse_to_rpn(self, expression):
        """
        Parse an excel formula expression into reverse polish notation

        Core algorithm taken from wikipedia with varargs extensions from
        http://www.kallisti.net.nz/blog/2008/02/extension-to-the-shunting-yard-
            algorithm-to-allow-variable-numbers-of-arguments-to-functions/
        """

        lexer = Tokenizer(expression)

        # amend token stream to ease code production
        tokens = []
        for token, next_token in zip(lexer.items, lexer.items[1:] + [None]):

            if token.matches(Token.FUNC, Token.OPEN):
                tokens.append(token)
                token = Token('(', Token.PAREN, Token.OPEN)
                if next_token.matches(Token.SEP, Token.ARG):
                    tokens.append(token)
                    token = Token('', Token.OPERAND, Token.EMPTY)

            elif token.matches(Token.FUNC, Token.CLOSE):
                token = Token(')', Token.PAREN, Token.CLOSE)

            elif token.matches(Token.ARRAY, Token.OPEN):
                tokens.append(token)
                tokens.append(Token('(', Token.PAREN, Token.OPEN))
                tokens.append(Token('', Token.ARRAYROW, Token.OPEN))
                token = Token('(', Token.PAREN, Token.OPEN)

            elif token.matches(Token.ARRAY, Token.CLOSE):
                tokens.append(token)
                token = Token(')', Token.PAREN, Token.CLOSE)

            elif token.matches(Token.SEP, Token.ROW):
                tokens.append(Token(')', Token.PAREN, Token.CLOSE))
                tokens.append(Token(',', Token.SEP, Token.ARG))
                tokens.append(Token('', Token.ARRAYROW, Token.OPEN))
                token = Token('(', Token.PAREN, Token.OPEN)

            elif token.matches(Token.SEP, Token.ARG):
                if next_token.matches(Token.SEP, Token.ARG) or \
                        next_token.matches(Token.FUNC, Token.CLOSE):
                    tokens.append(token)
                    token = Token('', Token.OPERAND, Token.EMPTY)

            elif token.matches(Token.PAREN, Token.OPEN):
                token.value = '('

            elif token.matches(Token.PAREN, Token.CLOSE):
                token.value = ')'

            tokens.append(token)

        output = []
        stack = []
        were_values = []
        arg_count = []

        for token in tokens:
            if token.type == token.OPERAND:

                output.append(self._ast_node(token))
                if were_values:
                    were_values[-1] = True

            elif token.type != token.PAREN and token.subtype == token.OPEN:

                if token.type in (token.ARRAY, Token.ARRAYROW):
                    token = Token(token.type, token.type, token.subtype)

                stack.append(token)
                arg_count.append(0)
                if were_values:
                    were_values[-1] = True
                were_values.append(False)

            elif token.type == token.SEP:

                while stack and (stack[-1].subtype != token.OPEN):
                    output.append(self._ast_node(stack.pop()))

                if not len(were_values):
                    raise FormulaParserError("Mismatched or misplaced parentheses")

                were_values.pop()
                arg_count[-1] += 1
                were_values.append(False)

            elif token.is_operator:

                while stack and stack[-1].is_operator and (
                        token.precedence < stack[-1].precedence):
                    output.append(self._ast_node(stack.pop()))

                stack.append(token)

            elif token.subtype == token.OPEN:
                assert token.type in (token.FUNC, token.PAREN, token.ARRAY)
                stack.append(token)

            elif token.subtype == token.CLOSE:

                while stack and stack[-1].subtype != Token.OPEN:
                    output.append(self._ast_node(stack.pop()))

                if not stack:
                    raise FormulaParserError("Mismatched or misplaced parentheses")

                stack.pop()

                if stack and stack[-1].is_funcopen:
                    f = self._ast_node(stack.pop())
                    f.num_args = arg_count.pop() + int(were_values.pop())
                    output.append(f)

            else:
                assert token.type == token.WSPACE, f'Unexpected token: {token}'

        while stack:
            if stack[-1].subtype in (Token.OPEN, Token.CLOSE):
                raise FormulaParserError("Mismatched or misplaced parentheses")

            output.append(self._ast_node(stack.pop()))

        return output

    @classmethod
    def _build_ast(cls, rpn_expression):
        """build an AST from an Excel formula

        :param rpn_expression: a string formula or the result of parse_to_rpn()
        :return: AST which can be used to inspect the formula structure
        """

        # use a directed graph to store the syntax tree
        tree = DiGraph()

        # production stack
        stack = []

        for node in rpn_expression:
            # The graph does not maintain the order of adding nodes/edges, so
            # add an attribute 'pos' so we can always sort to the correct order

            node.ast = tree
            if isinstance(node, OperatorNode):
                if node.token.type == node.token.OP_IN:
                    try:
                        arg2 = stack.pop()
                        arg1 = stack.pop()
                    except IndexError:
                        raise FormulaParserError(
                            f"'{node.token.value}' operator missing operand")
                    tree.add_node(arg1, pos=0)
                    tree.add_node(arg2, pos=1)
                    tree.add_edge(arg1, node)
                    tree.add_edge(arg2, node)
                else:
                    try:
                        arg1 = stack.pop()
                    except IndexError:
                        raise FormulaParserError(
                            f"'{node.token.value}' operator missing operand")
                    tree.add_node(arg1, pos=1)
                    tree.add_edge(arg1, node)

            elif isinstance(node, FunctionNode):
                if node.num_args:
                    args = stack[-node.num_args:]
                    del stack[-node.num_args:]
                    for i, a in enumerate(args):
                        tree.add_node(a, pos=i)
                        tree.add_edge(a, node)
            else:
                tree.add_node(node, pos=0)

            stack.append(node)

        assert 1 == len(stack)
        return stack[0]
