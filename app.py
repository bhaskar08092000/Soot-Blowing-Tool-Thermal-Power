#!/usr/bin/env python
"""Streamlit performance dashboard for boiler and soot-blowing workbooks."""

from __future__ import annotations

import ast
import io
import math
import re
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
)
from openpyxl.utils.exceptions import InvalidFileException


APP_TITLE = " Soot Blowing Tool "
DEFAULT_WORKBOOK = Path(__file__).with_name("Copy of SB_01-07-2026 2.xlsx")
KNOWN_WORKBOOK_MODULES = (
    "Boiler Basic Inputs",
    "Boiler Inputs Dynamic",
    "Boiler Inputs",
    "Curves",
    "Calc Module",
    "Heat Balance",
    "SB Benefit Optimisation Engine",
    "Eco HT CF Model",
    "Boiler Efficiency PTC4",
    "Heat Balance",
)

st.set_page_config(page_title=APP_TITLE, page_icon="", layout="wide")


def safe_float(value: Any, default: float | None = None) -> float | None:
    """Return a finite numeric value, or the supplied default."""
    if value is None or isinstance(value, bool):
        return default
    try:
        result = float(value)
    except (TypeError, ValueError, OverflowError):
        return default
    return result if math.isfinite(result) else default


def compact_text(value: Any, limit: int = 240) -> Any:
    if isinstance(value, str) and len(value) > limit:
        return value[:limit] + "..."
    return value


def cell_ref_parts(reference: str) -> tuple[int, int]:
    reference = reference.replace("$", "")
    column, row = coordinate_from_string(reference)
    return int(row), column_index_from_string(column)


def flatten(values: Iterable[Any]) -> list[Any]:
    flat: list[Any] = []
    for value in values:
        if isinstance(value, (list, tuple, np.ndarray)):
            flat.extend(flatten(value))
        else:
            flat.append(value)
    return flat


def excel_truth(value: Any) -> bool:
    if isinstance(value, str):
        return value.upper() not in ("", "FALSE", "0")
    return bool(value)


def excel_convert(value: Any, source: Any, target: Any) -> float:
    """Small, explicit subset of Excel CONVERT units used in this workbook."""
    amount = float(value)
    source_unit = str(source).strip().lower()
    target_unit = str(target).strip().lower()
    if source_unit == target_unit:
        return amount
    if source_unit in ("c", "degc", "celsius") and target_unit in ("f", "degf", "fahrenheit"):
        return amount * 9 / 5 + 32
    if source_unit in ("f", "degf", "fahrenheit") and target_unit in ("c", "degc", "celsius"):
        return (amount - 32) * 5 / 9
    if source_unit in ("c", "degc", "celsius") and target_unit in ("k", "kel"):
        return amount + 273.15
    if source_unit in ("k", "kel") and target_unit in ("c", "degc", "celsius"):
        return amount - 273.15
    if source_unit in ("f", "degf", "fahrenheit") and target_unit in ("k", "kel"):
        return (amount - 32) * 5 / 9 + 273.15
    if source_unit in ("k", "kel") and target_unit in ("f", "degf", "fahrenheit"):
        return (amount - 273.15) * 9 / 5 + 32

    unit_groups = {
        "mass": {
            "kg": 1.0, "g": 0.001, "gram": 0.001, "lbm": 0.45359237,
            "lb": 0.45359237, "ozm": 0.028349523125, "ton": 907.18474,
            "t": 1000.0,
        },
        "pressure": {
            "pa": 1.0, "kpa": 1000.0, "bar": 100000.0, "atm": 101325.0,
            "psi": 6894.757293168, "mmhg": 133.322387415,
        },
        "energy": {
            "j": 1.0, "kj": 1000.0, "cal": 4.184, "kcal": 4184.0,
            "btu": 1055.05585262, "wh": 3600.0, "kwh": 3600000.0,
        },
        "length": {
            "m": 1.0, "km": 1000.0, "cm": 0.01, "mm": 0.001,
            "in": 0.0254, "ft": 0.3048, "mi": 1609.344,
        },
        "time": {
            "s": 1.0, "sec": 1.0, "min": 60.0, "hr": 3600.0,
            "h": 3600.0, "day": 86400.0,
        },
        "volume": {
            "m3": 1.0, "l": 0.001, "liter": 0.001, "gal": 0.003785411784,
        },
        "power": {"w": 1.0, "kw": 1000.0, "hp": 745.699871582},
    }
    for units in unit_groups.values():
        if source_unit in units and target_unit in units:
            return amount * units[source_unit] / units[target_unit]
    raise ValueError(f"Unsupported CONVERT units: {source!s} to {target!s}")


class FormulaEngine:
    """Evaluate common workbook formula syntax without executing uploaded code."""

    FORMULA_ERRORS = (
        ArithmeticError, AttributeError, IndexError, KeyError, RecursionError,
        SyntaxError, TypeError, ValueError,
    )
    _quoted_sheet_ref = re.compile(
        r"'((?:[^']|'')+)'!\$?([A-Z]{1,3}\$?\d+)(?::\$?([A-Z]{1,3}\$?\d+))?",
        re.IGNORECASE,
    )
    _plain_sheet_ref = re.compile(
        r"\b([A-Za-z_][A-Za-z0-9_.]*)!\$?([A-Z]{1,3}\$?\d+)"
        r"(?::\$?([A-Z]{1,3}\$?\d+))?"
    )
    _range_ref = re.compile(
        r"(?<![A-Z0-9_])\$?([A-Z]{1,3}\$?\d+):\$?([A-Z]{1,3}\$?\d+)(?![A-Z0-9_])",
        re.IGNORECASE,
    )
    _cell_ref = re.compile(r"(?<![A-Z0-9_])\$?([A-Z]{1,3}\$?\d+)(?![A-Z0-9_])", re.IGNORECASE)

    def __init__(self, formula_book: openpyxl.Workbook, value_book: openpyxl.Workbook):
        self.formula_book = formula_book
        self.value_book = value_book
        self.memo: dict[tuple[str, str], Any] = {}
        self.visiting: set[tuple[str, str]] = set()
        self.failures: Counter[str] = Counter()
        self.fallback_cells: set[tuple[str, str]] = set()
        self.fallback_reasons: dict[tuple[str, str], str] = {}
        self.evaluated = 0
        self.cached_fallbacks = 0

    @staticmethod
    def _value_formula(cell: Any) -> str | None:
        value = cell.value
        if isinstance(value, str) and value.startswith("="):
            return value[1:]
        if isinstance(value, ArrayFormula):
            formula_text = value.text
            if isinstance(formula_text, str):
                return formula_text[1:] if formula_text.startswith("=") else formula_text
        return None

    def cached(self, sheet: str, coordinate: str) -> Any:
        if sheet not in self.value_book.sheetnames:
            return None
        return self.value_book[sheet][coordinate.replace("$", "")].value

    def cell(self, sheet: str, coordinate: str) -> Any:
        coordinate = coordinate.replace("$", "").upper()
        key = (sheet, coordinate)
        if key in self.memo:
            return self.memo[key]
        if sheet not in self.formula_book.sheetnames or key in self.visiting:
            return self.cached(sheet, coordinate)
        formula = self._value_formula(self.formula_book[sheet][coordinate])
        if formula is None:
            value = self.formula_book[sheet][coordinate].value
            self.memo[key] = value
            return value

        cached_value = self.cached(sheet, coordinate)
        if "RANDBETWEEN(" in formula.upper():
            result = cached_value
            if result is None:
                match = re.search(r"RANDBETWEEN\(\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*\)", formula, re.I)
                result = (float(match.group(1)) + float(match.group(2))) / 2 if match else None
            else:
                self.fallback_cells.add(key)
                self.fallback_reasons[key] = "Volatile formula; retained workbook's cached value"
                self.cached_fallbacks += 1
            self.memo[key] = result
            return result

        self.visiting.add(key)
        try:
            result = self._evaluate(formula, sheet)
            self.evaluated += 1
        except self.FORMULA_ERRORS as exc:
            self.failures[type(exc).__name__] += 1
            self.fallback_cells.add(key)
            self.fallback_reasons[key] = compact_text(str(exc) or type(exc).__name__)
            result = cached_value
            self.cached_fallbacks += 1
        finally:
            self.visiting.discard(key)
        self.memo[key] = result
        return result

    def range(self, sheet: str, reference: str) -> list[list[Any]]:
        first, last = reference.split(":")
        start_row, start_col = cell_ref_parts(first)
        end_row, end_col = cell_ref_parts(last)
        values = []
        for row in range(start_row, end_row + 1):
            values.append(
                [self.cell(sheet, get_column_letter(col) + str(row))
                 for col in range(start_col, end_col + 1)]
            )
        return values

    @staticmethod
    def _outside_string_parts(expression: str) -> list[str]:
        return re.split(r'("(?:[^"]|"")*")', expression)

    def _translate(self, formula: str, current_sheet: str) -> str:
        pieces = self._outside_string_parts(formula)
        for index in range(0, len(pieces), 2):
            text = pieces[index]

            def replace_quoted(match: re.Match[str]) -> str:
                sheet = match.group(1).replace("''", "'")
                if sheet.startswith("["):
                    raise ValueError("Formula depends on an external workbook")
                first = match.group(2).replace("$", "").upper()
                last = match.group(3)
                if last:
                    return f'_range("{sheet}","{first}:{last.replace("$", "").upper()}")'
                return f'_cell("{sheet}","{first}")'

            text = self._quoted_sheet_ref.sub(replace_quoted, text)

            def replace_plain_sheet(match: re.Match[str]) -> str:
                sheet = match.group(1)
                first = match.group(2).replace("$", "").upper()
                last = match.group(3)
                if last:
                    return f'_range("{sheet}","{first}:{last.replace("$", "").upper()}")'
                return f'_cell("{sheet}","{first}")'

            text = self._plain_sheet_ref.sub(replace_plain_sheet, text)
            translated_parts = self._outside_string_parts(text)
            for part_index in range(0, len(translated_parts), 2):
                range_translated = self._range_ref.sub(
                    lambda m: f'_range("{current_sheet}","{m.group(1).replace("$", "").upper()}:{m.group(2).replace("$", "").upper()}")',
                    translated_parts[part_index],
                )
                reference_parts = self._outside_string_parts(range_translated)
                for reference_index in range(0, len(reference_parts), 2):
                    reference_parts[reference_index] = self._cell_ref.sub(
                        lambda m: f'_cell("{current_sheet}","{m.group(1).replace("$", "").upper()}")',
                        reference_parts[reference_index],
                    )
                translated_parts[part_index] = "".join(reference_parts)
            text = "".join(translated_parts)
            pieces[index] = text

        expression_parts = self._outside_string_parts("".join(pieces))
        for index in range(0, len(expression_parts), 2):
            expression = expression_parts[index].replace("<>", "!=")
            expression = re.sub(r"(?<![<>=])=(?!=)", "==", expression)
            expression = expression.replace("^", "**")
            expression = re.sub(r"(?<![A-Z0-9_])(\d+(?:\.\d+)?)%(?![A-Za-z])", r"(\1/100)", expression, flags=re.IGNORECASE)
            expression = re.sub(r"\bTRUE\b", "True", expression, flags=re.IGNORECASE)
            expression = re.sub(r"\bFALSE\b", "False", expression, flags=re.IGNORECASE)
            expression_parts[index] = expression
        return "".join(expression_parts)

    @staticmethod
    def _numeric(values: Iterable[Any]) -> list[float]:
        out = []
        for value in flatten(values):
            number = safe_float(value)
            if number is not None:
                out.append(number)
        return out

    def _function(self, name: str, args: list[Any]) -> Any:
        fn = name.upper()
        values = flatten(args)
        numeric = self._numeric(args)
        if fn == "SUM":
            return sum(numeric)
        if fn == "AVERAGE":
            return sum(numeric) / len(numeric) if numeric else 0
        if fn == "MIN":
            return min(numeric) if numeric else 0
        if fn == "MAX":
            return max(numeric) if numeric else 0
        if fn == "COUNT":
            return len(numeric)
        if fn == "COUNTA":
            return sum(value is not None and value != "" for value in values)
        if fn == "AND":
            return all(excel_truth(value) for value in values)
        if fn == "OR":
            return any(excel_truth(value) for value in values)
        if fn == "NOT":
            return not excel_truth(args[0])
        if fn == "ISTEXT":
            return isinstance(args[0], str)
        if fn == "ISNUMBER":
            return safe_float(args[0]) is not None
        if fn == "ABS":
            return abs(float(args[0]))
        if fn == "ROUND":
            return round(float(args[0]), int(args[1]) if len(args) > 1 else 0)
        if fn == "EXP":
            return math.exp(float(args[0]))
        if fn == "LN":
            return math.log(float(args[0]))
        if fn == "PI":
            return math.pi
        if fn == "CONVERT":
            return excel_convert(args[0], args[1], args[2])
        if fn == "COUNTIF":
            criteria = args[1]
            pattern = str(criteria)
            def criteria_matches(value: Any) -> bool:
                if pattern.startswith("<>"):
                    return str(value).upper() != pattern[2:].upper()
                if pattern.startswith(">="):
                    number, boundary = safe_float(value), safe_float(pattern[2:])
                    return number is not None and boundary is not None and number >= boundary
                if pattern.startswith("<="):
                    number, boundary = safe_float(value), safe_float(pattern[2:])
                    return number is not None and boundary is not None and number <= boundary
                if pattern.startswith(">"):
                    number, boundary = safe_float(value), safe_float(pattern[1:])
                    return number is not None and boundary is not None and number > boundary
                if pattern.startswith("<"):
                    number, boundary = safe_float(value), safe_float(pattern[1:])
                    return number is not None and boundary is not None and number < boundary
                return str(value).upper() == pattern.upper()
            return sum(criteria_matches(value) for value in flatten([args[0]]))
        if fn == "MATCH":
            needle, haystack = args[0], flatten([args[1]])
            for index, value in enumerate(haystack, start=1):
                left_number, right_number = safe_float(value), safe_float(needle)
                if value == needle or (left_number is not None and left_number == right_number):
                    return index
            raise ValueError("MATCH did not find the requested value")
        if fn == "INDEX":
            values_arg = args[0]
            if isinstance(values_arg, list):
                matrix = values_arg if values_arg and isinstance(values_arg[0], list) else [values_arg]
                row = max(1, int(args[1])) - 1 if len(args) > 1 else 0
                col = max(1, int(args[2])) - 1 if len(args) > 2 else 0
                return matrix[row][col]
            raise ValueError("INDEX requires a range")
        if fn == "OFFSET":
            if len(args) < 3 or not isinstance(args[0], tuple):
                raise ValueError("OFFSET range unavailable")
            sheet, reference = args[0]
            row, col = cell_ref_parts(reference)
            target_row = row + int(args[1])
            target_col = col + int(args[2])
            height = int(args[3]) if len(args) > 3 else 1
            width = int(args[4]) if len(args) > 4 else 1
            first = f"{get_column_letter(target_col)}{target_row}"
            if height == width == 1:
                return self.cell(sheet, first)
            last = f"{get_column_letter(target_col + width - 1)}{target_row + height - 1}"
            return self.range(sheet, f"{first}:{last}")
        if fn == "RANK":
            value, group = args[0], self._numeric([args[1]])
            descending = len(args) < 3 or safe_float(args[2], 0) == 0
            return 1 + sum((item > value if descending else item < value) for item in group)
        if fn == "IFERROR":
            raise ValueError("IFERROR is evaluated lazily")
        if fn == "RANDBETWEEN":
            return (int(args[0]) + int(args[1])) // 2
        if fn == "IF":
            raise ValueError("IF is evaluated lazily")
        if fn == "LOOKUP":
            needle, vector, result_vector = args[0], flatten([args[1]]), flatten([args[2]]) if len(args) > 2 else flatten([args[1]])
            numeric_needle = safe_float(needle)
            if numeric_needle is None:
                raise ValueError("LOOKUP requires a numeric lookup value")
            match_indices = []
            for index, item in enumerate(vector):
                item_number = safe_float(item)
                if item_number is not None and item_number <= numeric_needle:
                    match_indices.append(index)
            if not match_indices:
                raise ValueError("LOOKUP did not find a value")
            return result_vector[match_indices[-1]]
        raise ValueError(f"Unsupported Excel function: {fn}")

    def _evaluate(self, formula: str, current_sheet: str) -> Any:
        expression = self._translate(formula, current_sheet)
        tree = ast.parse(expression, mode="eval")

        allowed = (
            ast.Expression, ast.Constant, ast.Name, ast.Call, ast.Load, ast.BinOp,
            ast.UnaryOp, ast.Compare, ast.BoolOp, ast.IfExp, ast.Add, ast.Sub,
            ast.Mult, ast.Div, ast.Pow, ast.Mod, ast.USub, ast.UAdd, ast.Not,
            ast.Eq, ast.NotEq, ast.Lt, ast.LtE, ast.Gt, ast.GtE, ast.And, ast.Or,
            ast.BitAnd,
        )
        for node in ast.walk(tree):
            if not isinstance(node, allowed):
                raise ValueError("Unsupported formula syntax")
            if isinstance(node, ast.Call) and (
                not isinstance(node.func, ast.Name) or node.func.id is None
            ):
                raise ValueError("Unsupported formula call")

        def run(node: ast.AST) -> Any:
            if isinstance(node, ast.Expression):
                return run(node.body)
            if isinstance(node, ast.Constant):
                return node.value
            if isinstance(node, ast.Name):
                if node.id == "True":
                    return True
                if node.id == "False":
                    return False
                raise ValueError(f"Unknown formula identifier: {node.id}")
            if isinstance(node, ast.Call):
                if not isinstance(node.func, ast.Name):
                    raise ValueError("Unsupported formula call")
                name = node.func.id.upper()
                if name == "_CELL":
                    return self.cell(run(node.args[0]), run(node.args[1]))
                if name == "_RANGE":
                    return self.range(run(node.args[0]), run(node.args[1]))
                if name == "IF":
                    condition = run(node.args[0])
                    return run(node.args[1] if excel_truth(condition) else node.args[2]) if len(node.args) > 2 else (
                        run(node.args[1]) if excel_truth(condition) else False
                    )
                if name == "IFERROR":
                    try:
                        return run(node.args[0])
                    except FormulaEngine.FORMULA_ERRORS:
                        return run(node.args[1])
                if name == "OFFSET":
                    first = node.args[0]
                    if not (
                        isinstance(first, ast.Call)
                        and isinstance(first.func, ast.Name)
                        and first.func.id == "_CELL"
                    ):
                        raise ValueError("OFFSET requires a cell reference")
                    reference = (run(first.args[0]), run(first.args[1]))
                    return self._function(name, [reference] + [run(arg) for arg in node.args[1:]])
                return self._function(name, [run(arg) for arg in node.args])
            if isinstance(node, ast.BinOp):
                left, right = run(node.left), run(node.right)
                def apply_binary(first: Any, second: Any) -> Any:
                    if isinstance(first, list) and isinstance(second, list):
                        if len(first) != len(second):
                            raise ValueError("Array formula dimensions do not match")
                        return [
                            apply_binary(left_value, right_value)
                            for left_value, right_value in zip(first, second)
                        ]
                    if isinstance(first, list):
                        return [apply_binary(value, second) for value in first]
                    if isinstance(second, list):
                        return [apply_binary(first, value) for value in second]
                    try:
                        if isinstance(node.op, ast.Add):
                            return first + second
                        if isinstance(node.op, ast.Sub):
                            return first - second
                        if isinstance(node.op, ast.Mult):
                            return first * second
                        if isinstance(node.op, ast.Div):
                            return first / second
                        if isinstance(node.op, ast.Pow):
                            return first ** second
                        if isinstance(node.op, ast.Mod):
                            return first % second
                        if isinstance(node.op, ast.BitAnd):
                            return str(first) + str(second)
                    except (ArithmeticError, TypeError, ValueError):
                        return None
                    raise ValueError("Unsupported formula operator")
                return apply_binary(left, right)
            if isinstance(node, ast.UnaryOp):
                value = run(node.operand)
                if isinstance(node.op, ast.USub):
                    return -value
                if isinstance(node.op, ast.UAdd):
                    return +value
                return not excel_truth(value)
            if isinstance(node, ast.BoolOp):
                values = [excel_truth(run(value)) for value in node.values]
                return all(values) if isinstance(node.op, ast.And) else any(values)
            if isinstance(node, ast.Compare):
                left = run(node.left)
                for operator, comparator in zip(node.ops, node.comparators):
                    right = run(comparator)
                    def apply_compare(first: Any, second: Any) -> Any:
                        if isinstance(first, list) and isinstance(second, list):
                            if len(first) != len(second):
                                raise ValueError("Array formula dimensions do not match")
                            return [
                                apply_compare(left_value, right_value)
                                for left_value, right_value in zip(first, second)
                            ]
                        if isinstance(first, list):
                            return [apply_compare(value, second) for value in first]
                        if isinstance(second, list):
                            return [apply_compare(first, value) for value in second]
                        if isinstance(operator, ast.Eq):
                            return first == second or (
                                safe_float(first) is not None and safe_float(first) == safe_float(second)
                            )
                        if isinstance(operator, ast.NotEq):
                            if first is None and second == "":
                                return False
                            return first != second
                        if first is None or second is None:
                            return False
                        if isinstance(operator, ast.Lt):
                            return first < second
                        if isinstance(operator, ast.LtE):
                            return first <= second
                        if isinstance(operator, ast.Gt):
                            return first > second
                        return first >= second

                    valid = apply_compare(left, right)
                    if isinstance(valid, list):
                        if len(node.ops) > 1:
                            raise ValueError("Chained array comparisons are not supported")
                        return valid
                    if not valid:
                        return False
                    left = right
                return True
            raise ValueError("Unsupported formula node")

        return run(tree)

    def value(self, sheet: str, coordinate: str) -> Any:
        return self.cell(sheet, coordinate)

    def formula_audit(self) -> pd.DataFrame:
        rows = []
        for sheet in self.formula_book.worksheets:
            formulas = [
                (cell.coordinate, self._value_formula(cell))
                for row in sheet.iter_rows()
                for cell in row
                if self._value_formula(cell) is not None
            ]
            if formulas:
                for coordinate, _ in formulas:
                    self.cell(sheet.title, coordinate)
                rows.append({
                    "Sheet": sheet.title,
                    "Formula cells": len(formulas),
                    "Recomputed": sum((sheet.title, coord) not in self.fallback_cells for coord, _ in formulas),
                    "Cached fallback": sum((sheet.title, coord) in self.fallback_cells for coord, _ in formulas),
                    "Fallback without cached value": sum(
                        (sheet.title, coord) in self.fallback_cells and self.cached(sheet.title, coord) is None
                        for coord, _ in formulas
                    ),
                })
        return pd.DataFrame(rows)


@st.cache_resource(show_spinner="Reading workbook...", max_entries=3)
def load_excel_workbooks(content: bytes) -> tuple[openpyxl.Workbook, openpyxl.Workbook]:
    formula_book = openpyxl.load_workbook(io.BytesIO(content), data_only=False, read_only=False)
    value_book = openpyxl.load_workbook(io.BytesIO(content), data_only=True, read_only=False)
    return formula_book, value_book


@st.cache_resource(show_spinner="Evaluating workbook formulas...", max_entries=3)
def load_formula_engine(content: bytes) -> FormulaEngine:
    formula_book, value_book = load_excel_workbooks(content)
    return FormulaEngine(formula_book, value_book)


def worksheet_value(engine: FormulaEngine | None, formula_book: openpyxl.Workbook | None,
                    value_book: openpyxl.Workbook | None, sheet: str, cell: str) -> Any:
    if engine and sheet in engine.formula_book.sheetnames:
        return engine.value(sheet, cell)
    if formula_book and sheet in formula_book.sheetnames:
        return formula_book[sheet][cell].value
    if value_book and sheet in value_book.sheetnames:
        return value_book[sheet][cell].value
    return None


def sheet_frame(engine: FormulaEngine, sheet_name: str, min_row: int = 1,
                max_row: int | None = None, max_col: int | None = None) -> pd.DataFrame:
    ws = engine.formula_book[sheet_name]
    end_row = min(max_row or ws.max_row, ws.max_row)
    end_col = min(max_col or ws.max_column, ws.max_column)
    rows = []
    for row_index in range(min_row, end_row + 1):
        values = [engine.value(sheet_name, get_column_letter(col) + str(row_index))
                  for col in range(1, end_col + 1)]
        if any(value is not None for value in values):
            rows.append(values)
    return pd.DataFrame(rows)


def get_label_value(engine: FormulaEngine, sheet_name: str, label: str,
                    value_column: str = "D") -> Any:
    ws = engine.formula_book[sheet_name]
    normalized = label.strip().upper()
    for row in range(1, ws.max_row + 1):
        cell_value = ws.cell(row, 2).value
        if isinstance(cell_value, str) and cell_value.strip().upper() == normalized:
            return engine.value(sheet_name, f"{value_column}{row}")
    return None


def build_dynamic_frame(engine: FormulaEngine) -> pd.DataFrame:
    sheet = "Boiler Inputs Dynamic"
    if sheet not in engine.formula_book.sheetnames:
        return pd.DataFrame()
    ws = engine.formula_book[sheet]
    headers = [ws.cell(3, col).value for col in range(1, ws.max_column + 1)]
    columns = [
        str(header).strip() if header is not None else f"Column {get_column_letter(index + 1)}"
        for index, header in enumerate(headers)
    ]
    records = []
    for row in range(5, ws.max_row + 1):
        values = [engine.value(sheet, f"{get_column_letter(col)}{row}")
                  for col in range(1, ws.max_column + 1)]
        if any(value is not None for value in values):
            records.append(dict(zip(columns, values)))
    frame = pd.DataFrame(records)
    if "TIME" in frame.columns:
        frame["TIME"] = frame["TIME"].map(lambda value: str(value) if value is not None else "")
    return frame


def curve_frame(engine: FormulaEngine) -> pd.DataFrame:
    sheet = "Curves"
    if sheet not in engine.formula_book.sheetnames:
        return pd.DataFrame()
    ws = engine.formula_book[sheet]
    headers = []
    for col in range(1, ws.max_column + 1):
        header = engine.value(sheet, f"{get_column_letter(col)}20")
        headers.append(str(header).strip() if header is not None else f"Column {col}")
    records = []
    for row in range(21, ws.max_row + 1):
        values = [engine.value(sheet, f"{get_column_letter(col)}{row}")
                  for col in range(1, ws.max_column + 1)]
        if any(value is not None for value in values):
            records.append(values)
    frame = pd.DataFrame(records, columns=headers)
    if frame.shape[1] > 1:
        frame = frame.rename(columns={frame.columns[1]: "Description"})
    return frame


def heat_balance_frames(
    engine: FormulaEngine,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    sheet = "Heat Balance"
    if sheet not in engine.formula_book.sheetnames:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    ws = engine.formula_book[sheet]
    sections = [
        ws.cell(3, col).value
        for col in range(2, 9)
    ]
    heat_absorbed = []
    heat_released = []
    effectiveness = []
    rows = [("Design", 6)]
    rows.extend(
        (f"Operating {engine.value(sheet, f'A{row}') or row - 7}", row)
        for row in range(8, 11)
    )
    for label, row in rows:
        absorbed = {
            "Case": label,
            **{
                str(sections[index]): safe_float(engine.value(sheet, f"{get_column_letter(index + 2)}{row}"))
                for index in range(len(sections))
                if sections[index] is not None
            },
            "Total heat absorbed (kW)": safe_float(engine.value(sheet, f"I{row}")),
        }
        released = {
            "Case": label,
            **{
                str(ws.cell(3, index + 11).value): safe_float(engine.value(sheet, f"{get_column_letter(index + 11)}{row}"))
                for index in range(7)
                if ws.cell(3, index + 11).value is not None
            },
            "Total heat released (kW)": safe_float(engine.value(sheet, f"R{row}")),
        }
        effectiveness_values = {
            "Case": label,
            **{
                str(ws.cell(3, index + 20).value): safe_float(engine.value(sheet, f"{get_column_letter(index + 20)}{row}"))
                for index in range(7)
                if ws.cell(3, index + 20).value is not None
            },
            "Overall effectiveness": safe_float(engine.value(sheet, f"AA{row}")),
        }
        heat_absorbed.append(absorbed)
        heat_released.append(released)
        effectiveness.append(effectiveness_values)

    performance = []
    for col in range(2, 10):
        label = engine.value(sheet, f"{get_column_letter(col)}15")
        value = engine.value(sheet, f"{get_column_letter(col)}17")
        unit = engine.value(sheet, f"{get_column_letter(col)}16")
        if label is not None or value is not None:
            performance.append({
                "Parameter": label,
                "Unit": unit,
                "Value": safe_float(value) if col > 2 else value,
            })
    return pd.DataFrame(heat_absorbed), pd.DataFrame(heat_released), pd.DataFrame(effectiveness), pd.DataFrame(performance)


def get_ptc_efficiency(engine: FormulaEngine) -> tuple[float | None, float | None]:
    sheet = "Boiler Efficiency PTC4"
    if sheet not in engine.formula_book.sheetnames:
        return None, None
    design = safe_float(engine.value(sheet, "D174"))
    operational = safe_float(engine.value(sheet, "F174"))
    return design, operational


def get_heat_rate(dynamic: pd.DataFrame, fallback_load: float | None = None) -> tuple[float | None, dict[str, float | None]]:
    data: dict[str, float | None] = {}
    if dynamic.empty:
        return None, data
    numeric = dynamic.copy()
    for column in numeric.columns:
        numeric[column] = pd.to_numeric(numeric[column], errors="coerce")
    candidates = [
        ("LOAD", "COAL GCV", "COAL FLOW"),
        ("LOAD", "COAL GCV", "COAL FLOW"),
    ]
    for load_col, gcv_col, coal_col in candidates:
        if all(column in numeric.columns for column in (load_col, gcv_col, coal_col)):
            valid = numeric[[load_col, gcv_col, coal_col]].dropna()
            valid = valid[(valid[load_col] > 0) & (valid[gcv_col] > 0) & (valid[coal_col] > 0)]
            if not valid.empty:
                record = valid.iloc[-1]
                load, gcv, coal = map(float, (record[load_col], record[gcv_col], record[coal_col]))
                return coal * gcv * 4.186 / load, {"Load (MW)": load, "Coal GCV (kcal/kg)": gcv, "Coal flow (t/h)": coal}
    if fallback_load and fallback_load > 0:
        data["Load (MW)"] = fallback_load
    return None, data


def optimizer_frame(engine: FormulaEngine, fuel_cost_multiplier: float = 1.0,
                    baseline_override: float | None = None, tau_multiplier: float = 1.0,
                    clean_factor_adjustment: float = 0.0) -> tuple[pd.DataFrame, pd.DataFrame]:
    sheet = "SB Benefit Optimisation Engine"
    if sheet not in engine.formula_book.sheetnames:
        return pd.DataFrame(), pd.DataFrame()
    ws = engine.formula_book[sheet]
    fuel_cost = safe_float(engine.value(sheet, "C8"), 0.0) or 0.0
    availability = safe_float(engine.value(sheet, "C15"), 0.85) or 0.85
    minimum_cycle = safe_float(engine.value(sheet, "C18"), 8.0) or 8.0
    workbook_native = (
        baseline_override is None
        and fuel_cost_multiplier == 1.0
        and tau_multiplier == 1.0
        and clean_factor_adjustment == 0.0
    )
    min_candidates: list[float] = []
    for row in range(35, 65):
        value = safe_float(engine.value(sheet, f"A{row}"))
        if value is not None and value >= minimum_cycle:
            min_candidates.append(value)
    if not min_candidates:
        min_candidates = [float(value) for value in np.arange(max(4, minimum_cycle), 121, 4)]

    results = []
    sweeps = []
    for row in range(23, 30):
        name = ws[f"A{row}"].value
        if not name:
            continue
        q_clean = safe_float(engine.value(sheet, f"B{row}"))
        cf_clean = safe_float(engine.value(sheet, f"C{row}"))
        cf_fouled = safe_float(engine.value(sheet, f"D{row}"))
        tau = safe_float(engine.value(sheet, f"E{row}"))
        live_cf = safe_float(engine.value(sheet, f"F{row}"))
        stack_factor = safe_float(engine.value(sheet, f"G{row}"))
        lances = safe_float(engine.value(sheet, f"I{row}"))
        blow_seconds = safe_float(engine.value(sheet, f"J{row}"))
        event_cost = safe_float(engine.value(sheet, f"K{row}"))
        configured_baseline = safe_float(engine.value(sheet, f"Y{row}"))
        if (
            q_clean is None
            or cf_clean is None
            or cf_fouled is None
            or tau is None
            or stack_factor is None
            or lances is None
            or blow_seconds is None
        ):
            continue
        event_cost_value = (
            event_cost
            if event_cost is not None
            else (safe_float(engine.value(sheet, f"X{row}")) or 0.0) * lances
        )
        baseline = float(baseline_override if baseline_override is not None else (configured_baseline or 12.0))
        tau_value = max(0.01, float(tau) * tau_multiplier)
        clean_factor = min(1.0, max(0.0, float(cf_clean) + clean_factor_adjustment))
        coefficient = (
            float(q_clean)
            * max(0, clean_factor - float(cf_fouled))
            * max(0, float(stack_factor))
            * 0.0036
            * fuel_cost
            * fuel_cost_multiplier
        )

        def cost_rate(cycle: float) -> float:
            degradation = cycle - tau_value * (1 - math.exp(-cycle / tau_value))
            return max(0.0, coefficient * degradation / cycle + event_cost_value / cycle)

        evaluated_sweep = [(cycle, cost_rate(cycle)) for cycle in min_candidates]
        optimum, minimum_cost = min(evaluated_sweep, key=lambda pair: pair[1])
        baseline_cost = cost_rate(baseline)
        workbook_saving = None
        if workbook_native:
            workbook_optimum = safe_float(engine.value(sheet, f"L{row}"))
            workbook_minimum_cost = safe_float(engine.value(sheet, f"N{row}"))
            workbook_baseline_cost = safe_float(engine.value(sheet, f"O{row}"))
            workbook_saving = safe_float(engine.value(sheet, f"P{row}"))
            if workbook_optimum is not None:
                optimum = workbook_optimum
            if workbook_minimum_cost is not None:
                minimum_cost = workbook_minimum_cost
            if workbook_baseline_cost is not None:
                baseline_cost = workbook_baseline_cost
        critical_cf = float(cf_fouled) + (clean_factor - float(cf_fouled)) * math.exp(-optimum / tau_value)
        current_cf = live_cf if live_cf is not None else clean_factor
        current_loss_rate = max(0, float(q_clean) * (clean_factor - current_cf) * max(0, float(stack_factor)) * 0.0036 * fuel_cost * fuel_cost_multiplier)
        results.append({
            "Section": str(name),
            "Live CF": current_cf,
            "Clean CF": clean_factor,
            "Fouled CF": float(cf_fouled),
            "Stack-loss factor": float(stack_factor),
            "Fuel-loss rate (INR/h)": current_loss_rate,
            "Lances": float(lances),
            "Blow duration (s)": float(blow_seconds),
            "Blow cost / event (INR)": event_cost_value,
            "Optimal interval (h)": optimum,
            "CF at optimum": critical_cf,
            "Cost at optimum (INR/h)": minimum_cost,
            "Cost at baseline (INR/h)": baseline_cost,
            "Annual saving (INR)": (
                workbook_saving
                if workbook_saving is not None
                else max(0, baseline_cost - minimum_cost) * 8760 * availability
            ),
            "Recommendation": "BLOW NOW" if current_cf <= critical_cf else "MONITOR",
        })
        for cycle, cost in evaluated_sweep:
            sweeps.append({"Section": str(name), "Cycle (h)": cycle, "Cost rate (INR/h)": cost})
    return pd.DataFrame(results), pd.DataFrame(sweeps)


def make_gauge(title: str, value: float | None, suffix: str = "", minimum: float = 0,
               maximum: float = 100, thresholds: tuple[float, float] | None = None) -> go.Figure:
    shown = value if value is not None else 0
    low, high = thresholds or (maximum * 0.65, maximum * 0.85)
    figure = go.Figure(go.Indicator(
        mode="gauge+number",
        value=shown,
        number={"suffix": suffix, "valueformat": ".2f"},
        title={"text": title},
        gauge={
            "axis": {"range": [minimum, maximum]},
            "bar": {"color": "#1677ff"},
            "steps": (
                [
                    {"range": [minimum, low], "color": "#fde2e2"},
                    {"range": [low, high], "color": "#fff2cc"},
                    {"range": [high, maximum], "color": "#e9f4ed"},
                ]
                if "efficiency" in title.lower()
                else [
                    {"range": [minimum, low], "color": "#e9f4ed"},
                    {"range": [low, high], "color": "#fff2cc"},
                    {"range": [high, maximum], "color": "#fde2e2"},
                ]
            ),
        },
    ))
    figure.update_layout(height=250, margin={"l": 20, "r": 20, "t": 50, "b": 10})
    return figure


def metric_value(value: float | None, suffix: str = "", precision: int = 1) -> str:
    return "N/A" if value is None else f"{value:,.{precision}f}{suffix}"


def get_sheet_summary(formula_book: openpyxl.Workbook) -> pd.DataFrame:
    rows = []
    for sheet in formula_book.worksheets:
        formulas = 0
        populated = 0
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    populated += 1
                    if FormulaEngine._value_formula(cell) is not None:
                        formulas += 1
        rows.append({"Worksheet": sheet.title, "Rows": sheet.max_row, "Columns": sheet.max_column,
                     "Populated cells": populated, "Formula cells": formulas})
    return pd.DataFrame(rows)


def report_pdf(project: str, kpis: dict[str, Any], optimizer: pd.DataFrame) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    target = io.BytesIO()
    document = SimpleDocTemplate(target, pagesize=landscape(A4), title=f"{project} Performance Report")
    styles = getSampleStyleSheet()
    story = [
        Paragraph(f"{project} — Plant Performance Report", styles["Title"]),
        Paragraph(f"Generated {datetime.now().astimezone().strftime('%Y-%m-%d %H:%M %Z')}", styles["Normal"]),
        Spacer(1, 12),
    ]
    kpi_rows = [["KPI", "Value"]] + [[name, str(value)] for name, value in kpis.items()]
    table = Table(kpi_rows, colWidths=[220, 260], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#15324b")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f6f9")]),
    ]))
    story.extend([Paragraph("Executive KPIs", styles["Heading2"]), table, Spacer(1, 16)])
    if not optimizer.empty:
        report_columns = ["Section", "Live CF", "Optimal interval (h)", "Recommendation", "Annual saving (INR)"]
        rows = [report_columns]
        for _, record in optimizer[report_columns].iterrows():
            rows.append([
                str(record["Section"]), f'{record["Live CF"]:.3f}',
                f'{record["Optimal interval (h)"]:.1f}', str(record["Recommendation"]),
                f'{record["Annual saving (INR)"]:,.0f}',
            ])
        opt_table = Table(rows, repeatRows=1)
        opt_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1677ff")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.lightgrey),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f6f9")]),
        ]))
        story.extend([Paragraph("Soot Blowing Optimization", styles["Heading2"]), opt_table])
    story.append(Spacer(1, 14))
    story.append(Paragraph(
        "Turbine/condenser results require measured inputs. Workbook formulas with broken external links or unsupported "
        "functions are reported with their cached workbook values where available; verify calculations before operational use.",
        styles["Italic"],
    ))
    document.build(story)
    return target.getvalue()


def report_xlsx(kpis: dict[str, Any], optimizer: pd.DataFrame, trend: pd.DataFrame,
                sheets: pd.DataFrame, audit: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame({"KPI": list(kpis.keys()), "Value": list(kpis.values())}).to_excel(writer, sheet_name="Executive KPIs", index=False)
        optimizer.to_excel(writer, sheet_name="Soot Blowing Optimizer", index=False)
        trend.to_excel(writer, sheet_name="Trend Data", index=False)
        sheets.to_excel(writer, sheet_name="Workbook Sheets", index=False)
        audit.to_excel(writer, sheet_name="Formula Audit", index=False)
    return output.getvalue()


def condenser_saturation_temp(pressure_kpa: float) -> float | None:
    """Approximate water saturation temperature from 1–100 C Antoine data."""
    pressure_mmhg = pressure_kpa * 760 / 101.325
    if pressure_mmhg <= 0:
        return None
    a, b, c = 8.07131, 1730.63, 233.426
    temp = b / (a - math.log10(pressure_mmhg)) - c
    return temp if -20 <= temp <= 110 else None


st.markdown(
    """
    <style>
    .block-container {padding-top: 1.4rem; padding-bottom: 2rem;}
    [data-testid="stMetric"] {background: #f5f8fc; border: 1px solid #e4ebf3; padding: 14px; border-radius: 12px;}
    .small-note {color: #637386; font-size: 0.9rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Soot Blowing Tool")
st.caption("Workbook-driven boiler diagnostics, heat-rate monitoring, and soot-blowing economics.")

with st.sidebar:
    st.header("Navigation")
    page = st.radio(
        "Select module",
        [
            "Executive Dashboard",
            "Boiler Module",
            "Heat Balance",
            "Turbine Module",
            "Condenser Module",
            "Soot Blowing Optimizer",
            "Trend Analysis",
            "KPI Monitoring",
            "What-if Analysis",
            "Data Upload",
            "Report Generation",
        ],
        label_visibility="collapsed",
    )
    st.divider()
    uploaded_file = st.file_uploader(
        "Workbook or telemetry CSV",
        type=["xlsx", "xlsm", "csv"],
        key="workbook_upload",
        help="Upload a workbook to replace the bundled workbook, or a CSV for operational trends.",
    )

formula_book: openpyxl.Workbook | None = None
value_book: openpyxl.Workbook | None = None
engine: FormulaEngine | None = None
source_name = ""
csv_frame = pd.DataFrame()
load_error: str | None = None

if uploaded_file is not None:
    source_name = uploaded_file.name
    content = uploaded_file.getvalue()
    if source_name.lower().endswith(".csv"):
        try:
            csv_frame = pd.read_csv(io.BytesIO(content))
        except (UnicodeDecodeError, pd.errors.ParserError, OSError) as exc:
            load_error = f"Could not read the CSV: {exc}"
    else:
        try:
            engine = load_formula_engine(content)
            formula_book, value_book = engine.formula_book, engine.value_book
        except (OSError, ValueError, KeyError, InvalidFileException) as exc:
            load_error = f"Could not open the workbook: {exc}"
elif DEFAULT_WORKBOOK.is_file():
    source_name = DEFAULT_WORKBOOK.name
    try:
        engine = load_formula_engine(DEFAULT_WORKBOOK.read_bytes())
        formula_book, value_book = engine.formula_book, engine.value_book
    except (OSError, ValueError, KeyError, InvalidFileException) as exc:
        load_error = f"Could not open the bundled workbook: {exc}"

if load_error:
    st.error(load_error)

if engine:
    st.sidebar.caption(f"Loaded: {source_name}")
    st.sidebar.caption(f"{len(engine.formula_book.sheetnames)} worksheets")
elif not csv_frame.empty:
    st.sidebar.caption(f"Loaded CSV: {source_name}")
elif page != "Data Upload":
    st.info("Upload the Excel workbook or a telemetry CSV to populate workbook-based modules.")

dynamic = build_dynamic_frame(engine) if engine else csv_frame.copy()
design_efficiency, operational_efficiency = get_ptc_efficiency(engine) if engine else (None, None)
heat_rate, heat_rate_inputs = get_heat_rate(dynamic) if not dynamic.empty else (None, {})
optimizer, optimizer_sweep = optimizer_frame(engine) if engine else (pd.DataFrame(), pd.DataFrame())
curves = curve_frame(engine) if engine and "Curves" in engine.formula_book.sheetnames else pd.DataFrame()
sheets_summary = get_sheet_summary(formula_book) if formula_book else pd.DataFrame()
formula_audit = engine.formula_audit() if engine else pd.DataFrame()

project_name = "Uploaded plant"
unit_number: Any = "—"
design_load = None
if engine and "Boiler Basic Inputs" in engine.formula_book.sheetnames:
    project_name = str(engine.value("Boiler Basic Inputs", "D2") or project_name)
    unit_number = engine.value("Boiler Basic Inputs", "D3") or unit_number
    design_load = safe_float(engine.value("Boiler Basic Inputs", "D4"))

latest_load = heat_rate_inputs.get("Load (MW)")
if latest_load is None and not dynamic.empty and "LOAD" in dynamic.columns:
    load_values = pd.to_numeric(dynamic["LOAD"], errors="coerce").dropna()
    latest_load = float(load_values.iloc[-1]) if not load_values.empty else None

if not optimizer.empty:
    annual_savings = float(optimizer["Annual saving (INR)"].sum())
    mean_live_cf = float(optimizer["Live CF"].mean())
else:
    annual_savings = None
    mean_live_cf = None

dashboard_kpis = {
    "Project": project_name,
    "Unit": unit_number,
    "Load": metric_value(latest_load, " MW"),
    "Boiler efficiency (design)": metric_value(design_efficiency, "%", 2),
    "Heat rate": metric_value(heat_rate, " kJ/kWh", 0),
    "Annual SB savings (model)": metric_value(annual_savings, " INR", 0),
}


if page == "Executive Dashboard":
    st.header("Executive Dashboard")
    left, right = st.columns([1.5, 1])
    with left:
        kpi_cols = st.columns(3)
        kpi_cols[0].metric("Project / Unit", f"{project_name} / {unit_number}")
        kpi_cols[1].metric("Operating load", metric_value(latest_load, " MW"))
        kpi_cols[2].metric("Design load", metric_value(design_load, " MW"))
        kpi_cols[0].metric("Boiler efficiency — design", metric_value(design_efficiency, "%", 2))
        kpi_cols[1].metric("Heat rate — latest case", metric_value(heat_rate, " kJ/kWh", 0))
        kpi_cols[2].metric("Soot-blowing savings", metric_value(annual_savings, " INR/yr", 0))
    with right:
        st.plotly_chart(make_gauge("Boiler efficiency (design)", design_efficiency, "%"), width="stretch")
    if operational_efficiency is None and engine and "Boiler Efficiency PTC4" in engine.formula_book.sheetnames:
        st.warning("Operational PTC-4 efficiency is unavailable in this workbook: the actual/derived columns contain broken #REF! links. The design efficiency is shown separately.")
    if not dynamic.empty and {"TIME", "LOAD"}.issubset(dynamic.columns):
        plot_data = dynamic.copy()
        plot_data["LOAD"] = pd.to_numeric(plot_data["LOAD"], errors="coerce")
        st.plotly_chart(px.line(plot_data, x="TIME", y="LOAD", markers=True, title="Boiler load by operating case"), width="stretch")
    if not optimizer.empty:
        st.subheader("Soot-blowing actions")
        st.dataframe(
            optimizer[["Section", "Live CF", "CF at optimum", "Optimal interval (h)", "Recommendation", "Annual saving (INR)"]]
            .sort_values("Annual saving (INR)", ascending=False),
            width="stretch",
            hide_index=True,
        )
    if formula_book:
        st.subheader("Workbook inventory")
        st.dataframe(sheets_summary, width="stretch", hide_index=True)

elif page == "Boiler Module":
    st.header("Boiler Module")
    if engine and "Boiler Inputs" in engine.formula_book.sheetnames:
        st.caption("Design and operating values are read from the workbook. PTC-4 loss calculations are evaluated in Python where formula dependencies are available.")
        values = {
            "Design efficiency": design_efficiency,
            "Operational efficiency": operational_efficiency,
            "Load (MW)": latest_load,
            "Heat rate (kJ/kWh)": heat_rate,
        }
        cols = st.columns(len(values))
        for col, (name, value) in zip(cols, values.items()):
            col.metric(name, metric_value(value, "%" if "efficiency" in name.lower() else "", 2 if "efficiency" in name.lower() else 0))
        if operational_efficiency is None:
            st.error("The workbook's operational PTC-4 values resolve to #REF! because dependent links are missing. No operational efficiency is inferred.")
        ptc = engine.formula_book["Boiler Efficiency PTC4"] if "Boiler Efficiency PTC4" in engine.formula_book.sheetnames else None
        if ptc:
            st.subheader("PTC-4 loss breakdown — design case")
            labels = [
                ("Dry gas loss", "D166"), ("Unburnt carbon loss", "D167"),
                ("Fuel moisture loss", "D168"), ("Hydrogen loss", "D169"),
                ("Air moisture loss", "D170"), ("Radiation / unaccounted", "D171"),
            ]
            loss_data = [{"Loss": label, "Loss (%)": safe_float(engine.value("Boiler Efficiency PTC4", cell))}
                         for label, cell in labels]
            loss_data = [item for item in loss_data if item["Loss (%)"] is not None]
            if loss_data:
                st.plotly_chart(px.bar(pd.DataFrame(loss_data), x="Loss", y="Loss (%)", title="PTC-4 design heat losses"), width="stretch")
                st.dataframe(pd.DataFrame(loss_data), hide_index=True, width="stretch")
        st.subheader("Boiler inputs")
        boiler = engine.formula_book["Boiler Inputs"]
        records = []
        for row in range(1, boiler.max_row + 1):
            label = boiler.cell(row, 2).value
            if isinstance(label, str) and label.strip() and row not in (2, 3):
                records.append({
                    "Parameter": label,
                    "Unit": boiler.cell(row, 3).value,
                    "Design": "" if engine.value("Boiler Inputs", f"D{row}") is None else str(engine.value("Boiler Inputs", f"D{row}")),
                    "Operating": "" if engine.value("Boiler Inputs", f"H{row}") is None else str(engine.value("Boiler Inputs", f"H{row}")),
                })
        st.dataframe(pd.DataFrame(records), width="stretch", hide_index=True)
    else:
        st.warning("This workbook has no recognized boiler inputs sheet. Upload the plant workbook to enable the boiler model.")

elif page == "Heat Balance":
    st.header("Heat Balance")
    if engine and "Heat Balance" in engine.formula_book.sheetnames:
        absorbed, released, effectiveness, degradation = heat_balance_frames(engine)
        if not absorbed.empty:
            design = absorbed.loc[absorbed["Case"] == "Design"]
            operating = absorbed.loc[absorbed["Case"] != "Design"]
            design_total = safe_float(design["Total heat absorbed (kW)"].iloc[0]) if not design.empty else None
            latest_total = safe_float(operating["Total heat absorbed (kW)"].iloc[-1]) if not operating.empty else None
            latest_effectiveness = safe_float(effectiveness["Overall effectiveness"].iloc[-1]) if not effectiveness.empty else None
            cards = st.columns(3)
            cards[0].metric("Design heat absorbed", metric_value(design_total, " kW", 0))
            cards[1].metric("Latest operating heat absorbed", metric_value(latest_total, " kW", 0))
            cards[2].metric("Latest overall effectiveness", metric_value(latest_effectiveness, "", 3))

            st.subheader("Heat absorbed by section")
            st.dataframe(absorbed, width="stretch", hide_index=True)
            absorbed_chart = absorbed.melt(
                id_vars=["Case"],
                value_vars=[column for column in absorbed.columns if column not in ("Case", "Total heat absorbed (kW)")],
                var_name="Heating section",
                value_name="Heat absorbed (kW)",
            )
            st.plotly_chart(
                px.bar(
                    absorbed_chart,
                    x="Case",
                    y="Heat absorbed (kW)",
                    color="Heating section",
                    title="Heat absorbed across boiler heating surfaces",
                ),
                width="stretch",
            )

            st.subheader("Heat released and section effectiveness")
            left, right = st.columns(2)
            with left:
                st.dataframe(released, width="stretch", hide_index=True)
            with right:
                st.dataframe(effectiveness, width="stretch", hide_index=True)
            st.caption("Section effectiveness and totals use the workbook's own formulas: heat absorbed / heat released and the workbook's per-case section sums.")

        if not degradation.empty:
            st.subheader("Heat-rate / efficiency degradation")
            st.dataframe(degradation, width="stretch", hide_index=True)
            st.caption("This table reproduces the Heat Balance sheet's user-entered Eco Actual CF scenario and its TGCHR, boiler-efficiency, gross-heat-rate, and replacement-coal calculations.")
        if absorbed.empty and degradation.empty:
            st.warning("The Heat Balance worksheet exists but contains no readable calculation results.")
    else:
        st.warning("No Heat Balance worksheet was found in the loaded workbook.")

elif page == "Turbine Module":
    st.header("Turbine Module")
    st.info("The supplied workbook has no dedicated turbine worksheet. Enter turbine test measurements below; values are kept in this session and are not represented as workbook measurements.")
    with st.form("turbine_inputs"):
        c1, c2, c3 = st.columns(3)
        steam_flow = c1.number_input("Turbine steam flow (t/h)", min_value=0.0, value=0.0, step=1.0)
        inlet_enthalpy = c2.number_input("Inlet steam enthalpy (kJ/kg)", min_value=0.0, value=0.0, step=10.0)
        exhaust_enthalpy = c3.number_input("Actual exhaust enthalpy (kJ/kg)", min_value=0.0, value=0.0, step=10.0)
        c4, c5 = st.columns(2)
        isentropic_exhaust = c4.number_input("Isentropic exhaust enthalpy (kJ/kg)", min_value=0.0, value=0.0, step=10.0)
        gross_power = c5.number_input("Generator gross output (MW)", min_value=0.0, value=0.0, step=1.0)
        calculate_turbine = st.form_submit_button("Calculate turbine KPIs")
    if calculate_turbine:
        st.session_state["turbine_inputs"] = {
            "steam_flow": steam_flow, "inlet_enthalpy": inlet_enthalpy,
            "exhaust_enthalpy": exhaust_enthalpy, "isentropic_exhaust": isentropic_exhaust,
            "gross_power": gross_power,
        }
    turbine = st.session_state.get("turbine_inputs", {})
    if turbine and turbine.get("steam_flow", 0) > 0:
        mass_flow = turbine["steam_flow"] * 1000 / 3600
        actual_drop = turbine["inlet_enthalpy"] - turbine["exhaust_enthalpy"]
        isentropic_drop = turbine["inlet_enthalpy"] - turbine["isentropic_exhaust"]
        steam_power = mass_flow * actual_drop / 1000
        efficiency = (actual_drop / isentropic_drop * 100) if isentropic_drop > 0 else None
        electric_efficiency = (turbine["gross_power"] / steam_power * 100) if steam_power > 0 else None
        metrics = st.columns(4)
        metrics[0].metric("Actual enthalpy drop", metric_value(actual_drop, " kJ/kg", 1))
        metrics[1].metric("Steam-side power", metric_value(steam_power, " MW", 2))
        metrics[2].metric("Isentropic efficiency", metric_value(efficiency, "%", 2))
        metrics[3].metric("Generator / steam power", metric_value(electric_efficiency, "%", 2))
        st.caption("Isentropic efficiency = (h_in − h_actual,out) / (h_in − h_isentropic,out). Steam-side power is an idealized enthalpy-flow estimate; validate boundary conditions and losses.")
    else:
        st.warning("Enter non-zero measured turbine data and submit the form to calculate results.")

elif page == "Condenser Module":
    st.header("Condenser Module")
    st.info("The supplied workbook has no condenser worksheet. Enter measured absolute pressure and cooling-water conditions for the calculation.")
    with st.form("condenser_inputs"):
        c1, c2, c3 = st.columns(3)
        pressure = c1.number_input("Condenser absolute pressure (kPa)", min_value=0.01, value=10.0, step=0.5)
        cw_in = c2.number_input("Cooling-water inlet temperature (°C)", value=25.0, step=1.0)
        cw_out = c3.number_input("Cooling-water outlet temperature (°C)", value=35.0, step=1.0)
        cw_flow = st.number_input("Cooling-water flow (kg/s)", min_value=0.0, value=0.0, step=10.0)
        calculate_condenser = st.form_submit_button("Calculate condenser KPIs")
    if calculate_condenser:
        st.session_state["condenser_inputs"] = {
            "pressure": pressure, "cw_in": cw_in, "cw_out": cw_out, "cw_flow": cw_flow,
        }
    condenser = st.session_state.get("condenser_inputs", {})
    if condenser:
        sat_temp = condenser_saturation_temp(condenser["pressure"])
        vacuum = max(0.0, 101.325 - condenser["pressure"])
        rise = condenser["cw_out"] - condenser["cw_in"]
        duty = condenser["cw_flow"] * 4.186 * rise
        approach = sat_temp - condenser["cw_out"] if sat_temp is not None else None
        metrics = st.columns(4)
        metrics[0].metric("Vacuum (gauge equivalent)", metric_value(vacuum, " kPa", 2))
        metrics[1].metric("Saturation temperature¹", metric_value(sat_temp, " °C", 2))
        metrics[2].metric("Cooling-water rise", metric_value(rise, " °C", 2))
        metrics[3].metric("Cooling duty", metric_value(duty, " kW", 0))
        st.metric("Condenser terminal temperature difference²", metric_value(approach, " °C", 2))
        st.caption("¹ Antoine correlation is an approximation near atmospheric pressure and should not be used outside its stated range. ² Outlet cooling-water temperature minus estimated saturation temperature. Duty uses cp = 4.186 kJ/kg·K.")
        if rise < 0:
            st.error("Cooling-water outlet is below inlet; check the entered measurements.")
        if condenser["pressure"] > 101.325:
            st.warning("Condenser absolute pressure is above standard atmospheric pressure; the vacuum-equivalent gauge value is shown as zero.")
    else:
        st.warning("Submit measured condenser inputs to calculate KPIs.")

elif page == "Soot Blowing Optimizer":
    st.header("Soot Blowing Optimizer")
    if engine and not optimizer.empty:
        st.caption("The objective follows the workbook's fouling-recovery cost-rate model: recovered heat loss plus steam/erosion event cost, minimized across the cycle-time sweep.")
        opt_sheet = engine.formula_book["SB Benefit Optimisation Engine"]
        copied_baseline_formula = any(
            isinstance(opt_sheet[f"O{row}"].value, str)
            and "$B$23" in opt_sheet[f"O{row}"].value
            for row in range(24, 30)
        )
        if copied_baseline_formula:
            st.warning("Workbook formula parity is active: annual savings use each section's Excel P-column formula and match the sheet's P30 total. The workbook's O-column formulas for several sections reference Economiser inputs, so review those Excel formulas before using the reported value as a validated business case.")
        controls = st.columns(3)
        fuel_multiplier = controls[0].slider("Fuel-cost multiplier", 0.5, 2.0, 1.0, 0.05)
        tau_factor = controls[1].slider("Fouling time-constant multiplier", 0.5, 2.0, 1.0, 0.05)
        clean_adjustment = controls[2].slider("Clean-CF adjustment", -0.05, 0.0, 0.0, 0.005)
        use_uniform_cycle = st.checkbox("Use one comparison cycle for all sections", value=False)
        baseline_cycle = st.number_input("Uniform comparison cycle (h)", 1.0, 168.0, 12.0, 1.0, disabled=not use_uniform_cycle)
        scenario, sweep = optimizer_frame(
            engine,
            fuel_multiplier,
            baseline_cycle if use_uniform_cycle else None,
            tau_factor,
            clean_adjustment,
        )
        st.metric("Modelled annual savings", metric_value(float(scenario["Annual saving (INR)"].sum()), " INR/yr", 0))
        st.dataframe(
            scenario.sort_values(["Recommendation", "Annual saving (INR)"], ascending=[True, False]),
            width="stretch",
            hide_index=True,
        )
        section = st.selectbox("Section cost curve", scenario["Section"].tolist())
        selection = sweep[sweep["Section"] == section]
        optimum = scenario.loc[scenario["Section"] == section, "Optimal interval (h)"].iloc[0]
        figure = px.line(selection, x="Cycle (h)", y="Cost rate (INR/h)", markers=True, title=f"{section}: total cost rate vs blow interval")
        figure.add_vline(x=optimum, line_dash="dash", annotation_text=f"optimum {optimum:g} h")
        st.plotly_chart(figure, width="stretch")
        st.caption("Savings are model estimates annualized using the optimizer sheet's availability assumption; validate costs, steam diversion, erosion and operating constraints before changing a schedule.")
    else:
        st.warning("No soot-blowing optimization model was found in the workbook.")

elif page == "Trend Analysis":
    st.header("Trend Analysis")
    if not dynamic.empty:
        numeric_columns = [
            column for column in dynamic.columns
            if pd.to_numeric(dynamic[column], errors="coerce").notna().any()
        ]
        if numeric_columns:
            x_default = "TIME" if "TIME" in dynamic.columns else dynamic.columns[0]
            x_options = dynamic.columns.tolist()
            x_index = next((index for index, column in enumerate(x_options) if column == x_default), 0)
            x_column = st.selectbox("Horizontal axis", x_options, index=x_index)
            default_metrics = [column for column in ("LOAD", "COAL FLOW", "COAL GCV") if column in numeric_columns]
            y_columns = st.multiselect("Trend metrics", numeric_columns, default=default_metrics or numeric_columns[:1])
            if y_columns:
                plot_data = dynamic.copy()
                for column in y_columns:
                    plot_data[column] = pd.to_numeric(plot_data[column], errors="coerce")
                plot_data = plot_data.melt(id_vars=[x_column], value_vars=y_columns, var_name="Metric", value_name="Value")
                st.plotly_chart(px.line(plot_data, x=x_column, y="Value", color="Metric", markers=True), width="stretch")
            st.dataframe(dynamic, width="stretch", hide_index=True)
        else:
            st.warning("The uploaded data contains no numeric trend fields.")
    else:
        st.info("No dynamic operating cases are available. Upload a workbook or CSV containing timestamped operating values.")
    if not curves.empty:
        st.subheader("Thermal curves")
        desc_values = curves["Description"].fillna("").astype(str)
        parameters = sorted(desc_values[desc_values.str.strip() != ""].unique().tolist())
        if parameters:
            selected = st.selectbox("Curve parameter", parameters)
            row = curves.loc[desc_values == selected].iloc[0]
            numeric = {column: safe_float(row[column]) for column in curves.columns[2:]}
            curve_data = pd.DataFrame([{"Operating case": key, "Value": val} for key, val in numeric.items() if val is not None])
            if not curve_data.empty:
                st.plotly_chart(px.line(curve_data, x="Operating case", y="Value", markers=True, title=selected), width="stretch")

elif page == "KPI Monitoring":
    st.header("KPI Monitoring")
    low_cf_threshold = st.slider("Low cleanliness-factor alert threshold", 0.70, 0.99, 0.90, 0.01)
    min_efficiency = st.slider("Minimum boiler efficiency target (%)", 50.0, 100.0, 85.0, 0.5)
    cells = st.columns(4)
    cells[0].metric("Design boiler efficiency", metric_value(design_efficiency, "%", 2))
    cells[1].metric("Latest heat rate", metric_value(heat_rate, " kJ/kWh", 0))
    cells[2].metric("Average live CF", metric_value(mean_live_cf, "", 3))
    cells[3].metric("Annualized SB saving", metric_value(annual_savings, " INR", 0))
    if design_efficiency is not None:
        if design_efficiency < min_efficiency:
            st.error(f"Design efficiency {design_efficiency:.2f}% is below the selected target of {min_efficiency:.1f}%.")
        else:
            st.success("Design efficiency meets the selected target.")
    if not optimizer.empty:
        alerts = optimizer[optimizer["Live CF"] < low_cf_threshold]
        if not alerts.empty:
            st.warning(f"{len(alerts)} section(s) are below the cleanliness alert threshold.")
            st.dataframe(alerts[["Section", "Live CF", "CF at optimum", "Recommendation"]], width="stretch", hide_index=True)
        else:
            st.success("No section is below the selected cleanliness-factor threshold.")
        st.subheader("Soot-blowing decision status")
        st.plotly_chart(px.bar(optimizer, x="Section", y="Fuel-loss rate (INR/h)", color="Recommendation", title="Current modeled fouling-loss cost"), width="stretch")
    if operational_efficiency is None and engine:
        st.info("Operational boiler efficiency KPI is unavailable because the workbook's actual PTC-4 formula references are broken.")

elif page == "What-if Analysis":
    st.header("What-if Analysis")
    if engine and not optimizer.empty:
        st.write("Explore economic and fouling assumptions without editing or overwriting the source workbook.")
        c1, c2, c3 = st.columns(3)
        fuel_multiplier = c1.slider("Fuel price multiplier", 0.5, 2.0, 1.0, 0.05, key="whatif_fuel")
        tau_multiplier = c2.slider("Fouling rate multiplier", 0.5, 2.0, 1.0, 0.05, key="whatif_tau")
        cycle = c3.slider("Fixed schedule interval (h)", 4, 120, 12, 4, key="whatif_cycle")
        base, base_sweep = optimizer_frame(engine)
        scenario, scenario_sweep = optimizer_frame(engine, fuel_multiplier, cycle, tau_multiplier)
        comparison = base[["Section", "Optimal interval (h)", "Annual saving (INR)"]].merge(
            scenario[["Section", "Optimal interval (h)", "Annual saving (INR)"]],
            on="Section", suffixes=(" (base)", " (scenario)"),
        )
        comparison["Savings delta (INR)"] = comparison["Annual saving (INR) (scenario)"] - comparison["Annual saving (INR) (base)"]
        st.dataframe(comparison, width="stretch", hide_index=True)
        left, right = st.columns(2)
        left.metric("Base annualized savings", metric_value(float(base["Annual saving (INR)"].sum()), " INR", 0))
        right.metric("Scenario annualized savings", metric_value(float(scenario["Annual saving (INR)"].sum()), " INR", 0))
        selected = st.selectbox("Compare cost curves", scenario["Section"].tolist(), key="whatif_section")
        curve = pd.concat([
            base_sweep[base_sweep["Section"] == selected].assign(Scenario="Workbook baseline"),
            scenario_sweep[scenario_sweep["Section"] == selected].assign(Scenario="What-if"),
        ])
        st.plotly_chart(px.line(curve, x="Cycle (h)", y="Cost rate (INR/h)", color="Scenario", markers=True, title=f"{selected}: baseline vs what-if"), width="stretch")
    else:
        st.info("Upload the soot-blowing optimization workbook to run the fouling and cost sensitivity analysis.")

elif page == "Data Upload":
    st.header("Data Upload")
    st.write("Load the workbook used by the dashboard or a CSV containing time-series operating data.")
    if uploaded_file is None and DEFAULT_WORKBOOK.is_file():
        st.success(f"Using the workbook alongside app.py: {DEFAULT_WORKBOOK.name}")
    if engine:
        st.subheader("Workbook sheet inventory")
        st.dataframe(sheets_summary, width="stretch", hide_index=True)
        st.subheader("Formula evaluation status")
        st.caption("The evaluator recomputes supported Excel formulas in Python and falls back to cached Excel values for unsupported formulas, external links, or broken references.")
        st.dataframe(formula_audit, width="stretch", hide_index=True)
        if formula_book:
            available = engine.formula_book.sheetnames
            absent = [name for name in KNOWN_WORKBOOK_MODULES if name not in available]
            if absent:
                st.warning("Not present in this workbook: " + ", ".join(absent))
            st.info("The supplied workbook has no dedicated turbine or condenser calculation sheets. Those modules use explicit measured inputs rather than inferred workbook values.")
            if engine.fallback_cells:
                fallback_rows = []
                for sheet_name, coordinate in sorted(engine.fallback_cells)[:100]:
                    formula = FormulaEngine._value_formula(engine.formula_book[sheet_name][coordinate])
                    cached_value = engine.cached(sheet_name, coordinate)
                    fallback_rows.append({
                        "Sheet": sheet_name,
                        "Cell": coordinate,
                        "Formula": "" if formula is None else compact_text(formula),
                        "Cached value": "" if cached_value is None else str(compact_text(cached_value)),
                        "Fallback reason": compact_text(engine.fallback_reasons.get((sheet_name, coordinate), "No cached result")),
                    })
                with st.expander(f"Formula fallbacks ({len(engine.fallback_cells)} total; showing up to 100)"):
                    st.dataframe(pd.DataFrame(fallback_rows), width="stretch", hide_index=True)
    elif not csv_frame.empty:
        st.subheader("CSV preview")
        st.write(f"{len(csv_frame):,} rows · {len(csv_frame.columns):,} columns")
        st.dataframe(csv_frame.head(100), width="stretch", hide_index=True)
    else:
        st.info("Choose an .xlsx, .xlsm, or .csv file from the sidebar to load data.")
    if engine:
        with st.expander("Formula audit detail"):
            st.dataframe(formula_audit, width="stretch", hide_index=True)

elif page == "Report Generation":
    st.header("Report Generation")
    st.caption("Create downloadable KPI and optimization reports from the currently loaded data.")
    st.subheader("Report preview")
    st.json(dashboard_kpis)
    if not optimizer.empty:
        st.dataframe(optimizer, width="stretch", hide_index=True)
    if engine or not dynamic.empty:
        xlsx_bytes = report_xlsx(dashboard_kpis, optimizer, dynamic, sheets_summary, formula_audit)
        st.download_button(
            "Download Excel report",
            data=xlsx_bytes,
            file_name=f"{re.sub(r'[^A-Za-z0-9_-]+', '_', project_name)}_performance_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        try:
            pdf_bytes = report_pdf(project_name, dashboard_kpis, optimizer)
        except ImportError:
            st.warning("PDF export requires reportlab. Excel report export is available.")
        else:
            st.download_button(
                "Download PDF report",
                data=pdf_bytes,
                file_name=f"{re.sub(r'[^A-Za-z0-9_-]+', '_', project_name)}_performance_report.pdf",
                mime="application/pdf",
            )
    else:
        st.warning("Load a workbook or telemetry CSV before generating a performance report.")

st.divider()
st.caption("Engineering decision-support only. Validate instruments, formula assumptions, and operating limits before applying recommendations.")
