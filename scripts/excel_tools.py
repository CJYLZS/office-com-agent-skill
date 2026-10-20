"""Guarded live Excel edits for an already-open workbook.

The read-only adapter remains excel_com.py. This module never opens, creates,
saves in place, closes, or quits Excel, and never changes application settings.

Preview is the default. A commit rebinds the target and re-reads local evidence
before writing once, then verifies native values/formulas/styles. COM is not
transactional: a failed or timed-out commit may have partially completed.
WriteOutcomeUnknown must be inspected manually; never retry or roll back
automatically.

Structural edits fail closed. The default spill_policy="native" requires a
working per-cell HasSpill == False. This Excel runtime exposes the property name
but throws when reading it, so native edits are refused there. An explicit
spill_policy="bounded_scalar_sheet" instead proves that every formula in the
whole UsedRange (<=100 cells) is a bounded scalar formula and that no cell is an
array formula. That is a complete local scalar inventory, not a claim that
HasSpill is false or that a dynamic-array engine was validated.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from excel_com import ExcelBridge, parse_address, read_range
from office_com import WriteOutcomeUnknown, rebind, rebind_after_write

MAX_CELLS = 100
MAX_FORMULA_CHARS = 512
MAX_TEXT_UNITS = 32767
MAX_TOTAL_TEXT_UNITS = 1_000_000
LEGACY_NO_SPILL = (12, 13, 14, 15)

STYLE_FIELDS = (
    'number_format', 'bold', 'font_size', 'font_name', 'color',
    'fill_color', 'horizontal_alignment', 'column_width',
)
ALIGNMENTS = {
    'general': 1, 'left': -4131, 'center': -4108, 'right': -4152,
    'fill': 5, 'justify': -4130, 'center_across': 7, 'distributed': -4117,
}
SPILL_POLICIES = ('native', 'bounded_scalar_sheet')


def _hash(obj):
    payload = json.dumps(
        obj, ensure_ascii=False, allow_nan=False, sort_keys=True,
        separators=(',', ':'),
    )
    return hashlib.sha256(payload.encode('utf-8', errors='surrogatepass')).hexdigest()


def _put_bstr(obj, prop, text):
    """Length-counted UTF-16 property put; avoids comtypes non-BMP truncation."""
    import ctypes
    from comtypes.automation import (
        IDispatch, VARIANT, DISPPARAMS, EXCEPINFO, IID_NULL,
    )
    from excel_com import automation_api, clear_exception_strings, check_hresult

    api = automation_api()
    alloc = api.SysAllocStringLen
    alloc.argtypes = [ctypes.c_void_p, ctypes.c_uint]
    alloc.restype = ctypes.c_void_p
    data = text.encode('utf-16-le', errors='strict')
    buf = ctypes.create_string_buffer(data)
    value = VARIANT()
    value.vt = 8
    value._.c_void_p = alloc(ctypes.cast(buf, ctypes.c_void_p), len(data) // 2)
    if not value._.c_void_p:
        raise MemoryError('Cannot allocate BSTR')
    result, exc = VARIANT(), EXCEPINFO()
    try:
        disp = obj.QueryInterface(IDispatch)
        ids = disp.GetIDsOfNames(prop)
        dispid = ids[0] if isinstance(ids, (tuple, list)) else ids
        named, arg = ctypes.c_long(-3), ctypes.c_uint()
        params = DISPPARAMS()
        params.rgvarg = ctypes.pointer(value)
        params.rgdispidNamedArgs = ctypes.pointer(named)
        params.cArgs = params.cNamedArgs = 1
        disp._IDispatch__com_Invoke(
            dispid, ctypes.byref(IID_NULL), 0, 4, ctypes.byref(params),
            ctypes.byref(result), ctypes.byref(exc), ctypes.byref(arg),
        )
    finally:
        primary = sys.exc_info()[1]
        try:
            try:
                check_hresult(api.VariantClear(ctypes.byref(result)))
            finally:
                try:
                    clear_exception_strings(exc)
                finally:
                    check_hresult(api.VariantClear(ctypes.byref(value)))
        except Exception as cleanup:
            if primary is None:
                raise
            if hasattr(primary, 'add_note'):
                primary.add_note('COM cleanup failed: ' + str(cleanup))


def _set_literal(cell, text):
    """Write a literal string. A leading apostrophe prevents formula coercion.

    Real Excel receives a length-counted BSTR and keeps non-BMP characters.
    A plain mock/lite object receives the exact text so postconditions exercise
    the same value without depending on COM.
    """
    if hasattr(cell, 'QueryInterface'):
        _put_bstr(cell, 'Value2', "'" + text)
    else:
        setattr(cell, 'Value2', text)


def _sheet(wb, spec):
    if (not isinstance(spec, dict) or set(spec) != {'name', 'code_name'}
            or not isinstance(spec['name'], str) or not spec['name']
            or not isinstance(spec['code_name'], str)):
        raise ValueError('Exact sheet name and code_name required')
    sh = wb.Worksheets.Item(spec['name'])
    if str(sh.Name) != spec['name'] or str(sh.CodeName) != spec['code_name']:
        raise ValueError('Worksheet identity mismatch')
    return sh


def _shape(value, rows, columns, key):
    if (not isinstance(value, list) or len(value) != rows
            or any(not isinstance(row, list) or len(row) != columns
                   for row in value)):
        raise ValueError(f'{key} must be exact {rows}x{columns} 2D array')
    return value


def _text_units(text):
    return len(text.encode('utf-16-le')) // 2


def _value(value):
    if value is None or type(value) is bool:
        return 0
    if type(value) in (int, float):
        if type(value) is float and not math.isfinite(value):
            raise ValueError('Non-finite number')
        if type(value) is int and abs(value) > 2 ** 53:
            raise ValueError('Integer cannot be represented exactly in Excel')
        return 0
    if type(value) is str:
        if '\x00' in value:
            raise ValueError('NUL cannot be stored in an Excel cell')
        if value == '':
            raise ValueError('Empty strings are unsupported; use null to clear')
        units = _text_units(value)
        if units > MAX_TEXT_UNITS:
            raise ValueError('Text exceeds the Excel cell limit')
        value.encode('utf-8')
        if value.lstrip().startswith(('=', '+', '-', '@')):
            raise ValueError('Formula-like string rejected as a value')
        return units
    raise ValueError('Values must be null, bool, finite number, or string')


_TOKEN = re.compile(
    r'\s*(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}'
    r'|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?'
    r'|[()+*/^%-])'
)
_SUM = re.compile(
    r'SUM\(\s*(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}'
    r':\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})\s*\)',
    re.IGNORECASE,
)


def _formula(formula):
    """Accept only scalar local arithmetic or one bounded SUM rectangle.

    Rejected: names, external/structured references, unions/intersections,
    range operators outside SUM, array/dynamic functions, implicit intersection
    (@), spill operators (#), unary operators, and adjacent operands.
    """
    if (not isinstance(formula, str) or len(formula) > MAX_FORMULA_CHARS
            or not formula.startswith('=')):
        raise ValueError('Formula must be a bounded string starting with =')
    expr = formula[1:].strip()
    if not expr:
        raise ValueError('Empty formula')
    summed = _SUM.fullmatch(expr)
    if summed:
        parse_address(summed.group(1), max_cells=MAX_CELLS)
        return formula
    tokens, position = [], 0
    while position < len(expr):
        match = _TOKEN.match(expr, position)
        if not match:
            raise ValueError('Formula outside the restricted scalar grammar')
        tokens.append(match.group(1))
        position = match.end()
    if not tokens:
        raise ValueError('Empty formula')
    expecting_operand, depth = True, 0
    for token in tokens:
        if token == '(':
            if not expecting_operand:
                raise ValueError('Formula has adjacent operands')
            depth += 1
            expecting_operand = True
        elif token == ')':
            if expecting_operand or depth == 0:
                raise ValueError('Formula is unbalanced')
            depth -= 1
            expecting_operand = False
        elif token in ('+', '-', '*', '/', '^', '%'):
            if expecting_operand:
                raise ValueError('Formula uses an unsupported unary operator')
            expecting_operand = True
        else:
            if not expecting_operand:
                raise ValueError('Formula has adjacent operands')
            expecting_operand = False
    if depth or expecting_operand:
        raise ValueError('Formula is incomplete or unbalanced')
    return formula


def _style(cell):
    return {
        'number_format': cell.NumberFormat,
        'bold': cell.Font.Bold,
        'font_size': cell.Font.Size,
        'font_name': cell.Font.Name,
        'color': cell.Font.Color,
        'fill_color': cell.Interior.Color,
        'horizontal_alignment': cell.HorizontalAlignment,
        'column_width': cell.ColumnWidth,
    }


def _scalar_inventory(sh):
    """Prove the complete UsedRange contains only bounded scalar formulas."""
    used = sh.UsedRange
    address = used.Address
    if callable(address) and not isinstance(address, str):
        # Excel's Address is a parameterized property; (0, 0) is relative A1
        # notation without dollar signs.
        address = address(0, 0)
    bounds = parse_address(str(address), max_cells=MAX_CELLS)
    rng = sh.Range(bounds['address'])
    if (int(rng.Rows.Count) != bounds['rows']
            or int(rng.Columns.Count) != bounds['columns']
            or int(rng.Areas.Count) != 1):
        raise ValueError('UsedRange shape is inconsistent')
    cells = []
    for row in range(1, bounds['rows'] + 1):
        for column in range(1, bounds['columns'] + 1):
            cell = rng.Cells.Item(row, column)
            has_formula = cell.HasFormula
            has_array = cell.HasArray
            if type(has_formula) is not bool or type(has_array) is not bool:
                raise ValueError('Scalar inventory has an unknown formula/array flag')
            if has_array:
                raise ValueError('Scalar inventory found an array formula')
            if has_formula:
                formula = str(cell.Formula)
                _formula(formula)
                cells.append([row, column, formula])
    return {
        'kind': 'bounded_scalar_sheet_inventory',
        'used_address': bounds['address'],
        'rows': bounds['rows'],
        'columns': bounds['columns'],
        'formula_cells': cells,
    }


def _structure(wb, sh, rng, bounds, spill_policy, *, diagnostic=False):
    checks, blockers = {}, []

    def require_false(label, get):
        try:
            value = get()
            checks[label] = value if type(value) is bool else None
            if type(value) is not bool or value is not False:
                blockers.append(label)
        except Exception:
            checks[label] = None
            blockers.append(label + ': unavailable')

    for name in ('ReadOnly', 'ProtectStructure', 'HasVBProject', 'MultiUserEditing'):
        require_false('workbook.' + name, lambda name=name: getattr(wb, name))
    for name in ('ProtectContents', 'ProtectDrawingObjects', 'ProtectScenarios'):
        require_false('sheet.' + name, lambda name=name: getattr(sh, name))
    for name in ('ListObjects', 'PivotTables', 'QueryTables'):
        try:
            collection = getattr(sh, name)
            collection = collection() if name == 'PivotTables' else collection
            count = collection.Count
            if type(count) is not int or count < 0:
                raise ValueError('Unknown collection count')
            checks['sheet.' + name] = count
            if count:
                blockers.append('sheet.' + name)
        except Exception:
            checks['sheet.' + name] = None
            blockers.append('sheet.' + name + ': unavailable')
    try:
        major = int(str(wb.Application.Version).split('.')[0])
        checks['excel_major_version'] = major
    except Exception:
        major = None
        checks['excel_major_version'] = None
    for row in range(1, bounds['rows'] + 1):
        for column in range(1, bounds['columns'] + 1):
            cell = rng.Cells.Item(row, column)
            for name in ('MergeCells', 'HasArray'):
                require_false(
                    f'cell[{row},{column}].{name}',
                    lambda name=name, cell=cell: getattr(cell, name),
                )
            try:
                has_formula = cell.HasFormula
                checks[f'cell[{row},{column}].HasFormula'] = (
                    has_formula if type(has_formula) is bool else None
                )
                if type(has_formula) is not bool:
                    blockers.append(f'cell[{row},{column}].HasFormula: unknown')
            except Exception:
                checks[f'cell[{row},{column}].HasFormula'] = None
                blockers.append(f'cell[{row},{column}].HasFormula: unavailable')
            try:
                spilled = cell.HasSpill
                checks[f'cell[{row},{column}].HasSpill'] = (
                    spilled if type(spilled) is bool else None
                )
                if type(spilled) is not bool:
                    blockers.append(f'cell[{row},{column}].HasSpill: unknown')
                elif spilled:
                    blockers.append(f'cell[{row},{column}].HasSpill')
            except Exception:
                checks[f'cell[{row},{column}].HasSpill'] = None
                if spill_policy == 'bounded_scalar_sheet':
                    checks[f'cell[{row},{column}].HasSpill'] = (
                        'deferred_to_bounded_scalar_inventory'
                    )
                elif major in LEGACY_NO_SPILL:
                    checks[f'cell[{row},{column}].HasSpill'] = (
                        'not_applicable_legacy'
                    )
                else:
                    blockers.append(f'cell[{row},{column}].HasSpill: unavailable')
    if blockers and not diagnostic:
        raise ValueError('Structural safety checks failed: ' + ', '.join(blockers))
    checks['spill_policy'] = spill_policy
    if spill_policy == 'bounded_scalar_sheet':
        try:
            checks['spill_proof'] = _scalar_inventory(sh)
        except Exception as exc:
            if not diagnostic:
                raise
            blockers.append('bounded_scalar_inventory: ' + type(exc).__name__)
    if diagnostic:
        return {'checks': checks, 'blockers': blockers,
                'checked_structure_clear': not blockers, 'write_authorized': False}
    return checks


def read_edit_range(wb, sheet, address, spill_policy='native'):
    if spill_policy not in SPILL_POLICIES:
        raise ValueError('Unsupported spill_policy')
    bounds = parse_address(address, max_cells=MAX_CELLS)
    sh = _sheet(wb, sheet)
    rng = sh.Range(bounds['address'])
    if (int(rng.Rows.Count) != bounds['rows']
            or int(rng.Columns.Count) != bounds['columns']
            or int(rng.Areas.Count) != 1):
        raise ValueError('Range shape mismatch')
    checks = _structure(wb, sh, rng, bounds, spill_policy)
    snapshot = read_range(
        wb, sheet['name'], bounds['address'], mode='both', formula_api='Formula'
    )
    mask, formulas, styles = [], [], []
    for row in range(1, bounds['rows'] + 1):
        mask_row, formula_row, style_row = [], [], []
        for column in range(1, bounds['columns'] + 1):
            cell = rng.Cells.Item(row, column)
            has_formula = cell.HasFormula
            if type(has_formula) is not bool:
                raise ValueError('HasFormula is not a scalar boolean')
            mask_row.append(has_formula)
            formula_row.append(str(cell.Formula) if has_formula else None)
            style_row.append(_style(cell))
        mask.append(mask_row)
        formulas.append(formula_row)
        styles.append(style_row)
    evidence = {
        'sheet': snapshot['sheet'],
        'address': bounds['address'],
        'snapshot': snapshot,
        'has_formula': mask,
        'formulas': formulas,
        'structural_checks': checks,
        'styles': styles,
        'spill_policy': spill_policy,
    }
    return {
        'sheet': snapshot['sheet'],
        **bounds,
        'snapshot': snapshot,
        'has_formula': mask,
        'formulas': formulas,
        'structural_checks': checks,
        'structural_risks': [],
        'styles': styles,
        'spill_policy': spill_policy,
        'hash_scope': 'excel_edit_evidence_v2',
        'sha256': _hash(evidence),
    }


REFUSED = ('Workbook cannot be rebound natively ({stage}); no write attempted '
           '(未尝试写入). Reread and replan: {error}')
IDENTITY_CHANGED = ('Workbook identity changed ({stage}); no write attempted '
                    '(未尝试写入). Reread and replan.')
WRITE_LOST = ('Write attempted, but the workbook cannot be rebound afterwards, so the '
              'outcome may be partial. Do not retry or undo automatically; inspect '
              '{label} first: {error}')


def _target(bridge, target, stage='读取前'):
    if not isinstance(target, dict) or target.get('application') != 'excel':
        raise ValueError('Full Excel target identity required')
    return rebind(bridge, target, stage, refusal=REFUSED, mismatch=IDENTITY_CHANGED)


def _save_path(value):
    if not isinstance(value, str) or not value:
        raise ValueError('save_copy needs an explicit path')
    path = Path(os.path.abspath(os.path.expanduser(value)))
    if path.suffix.lower() != '.xlsx':
        raise ValueError('Only an absolute .xlsx destination is supported')
    if not path.parent.is_dir():
        raise ValueError('Destination parent directory must already exist')
    if path.exists():
        raise FileExistsError(str(path))
    return path


def _save_preconditions(wb, path):
    if bool(wb.ReadOnly):
        raise ValueError('Workbook is read-only')
    if bool(wb.HasVBProject):
        raise ValueError('Macro-containing workbooks are unsupported')
    if int(wb.FileFormat) != 51:
        raise ValueError('Source must be a native macro-free .xlsx workbook')
    if path.exists():
        raise FileExistsError(str(path))
    if not path.parent.is_dir():
        raise ValueError('Destination parent directory must already exist')


def _save_evidence(wb, identity, path):
    return {
        'identity': identity,
        'path': str(path),
        'format': int(wb.FileFormat),
        'vba': bool(wb.HasVBProject),
        'read_only': bool(wb.ReadOnly),
        'saved': bool(wb.Saved),
    }


def _validate(request):
    if not isinstance(request, dict) or set(request) != {'target', 'operation'}:
        raise ValueError('Request needs exactly target and operation')
    op = request['operation']
    kind = op.get('kind') if isinstance(op, dict) else None
    if kind not in (
        'set_values', 'set_formulas', 'format_range', 'calculate_range', 'save_copy'
    ):
        raise ValueError('Unsupported operation')
    if kind == 'save_copy':
        if set(op) != {'kind', 'path'}:
            raise ValueError('save_copy accepts only kind and path')
        _save_path(op['path'])
        return op
    base = {'kind', 'sheet', 'address', 'expected_sha256'}
    policy = op.get('spill_policy', 'native')
    if policy not in SPILL_POLICIES:
        raise ValueError('Unsupported spill_policy')
    sheet = op.get('sheet', {})
    if (set(sheet) != {'name', 'code_name'}
            or not isinstance(sheet.get('name'), str) or not sheet['name']
            or not isinstance(sheet.get('code_name'), str)):
        raise ValueError('Exact sheet identity required')
    bounds = parse_address(op.get('address'), max_cells=MAX_CELLS)
    if not re.fullmatch('[0-9a-f]{64}', str(op.get('expected_sha256', ''))):
        raise ValueError('expected_sha256 required')
    if kind == 'set_values':
        if set(op) - (base | {'values', 'spill_policy'}) or 'values' not in op:
            raise ValueError('Invalid set_values fields')
        matrix = _shape(op['values'], bounds['rows'], bounds['columns'], 'values')
        total = 0
        for row in matrix:
            for item in row:
                total += _value(item)
        if total > MAX_TOTAL_TEXT_UNITS:
            raise ValueError('Total text budget exceeded')
    elif kind == 'set_formulas':
        if set(op) - (base | {'formulas', 'spill_policy'}) or 'formulas' not in op:
            raise ValueError('Invalid set_formulas fields')
        matrix = _shape(op['formulas'], bounds['rows'], bounds['columns'], 'formulas')
        for row in matrix:
            for item in row:
                _formula(item)
    elif kind == 'format_range':
        allowed = base | {'spill_policy'} | set(STYLE_FIELDS)
        if set(op) - allowed or not set(op) & set(STYLE_FIELDS):
            raise ValueError('Invalid or empty format fields')
        if 'bold' in op and type(op['bold']) is not bool:
            raise ValueError('bold must be bool')
        if 'font_size' in op and (
            type(op['font_size']) not in (int, float)
            or not 1 <= op['font_size'] <= 409
        ):
            raise ValueError('font_size out of range')
        if 'font_name' in op and (
            not isinstance(op['font_name'], str)
            or not 1 <= len(op['font_name']) <= 128
        ):
            raise ValueError('Invalid font_name')
        for key in ('color', 'fill_color'):
            if key in op and (
                type(op[key]) is not int or not 0 <= op[key] <= 0xffffff
            ):
                raise ValueError('Invalid ' + key)
        if 'number_format' in op and (
            not isinstance(op['number_format'], str)
            or len(op['number_format']) > 128
        ):
            raise ValueError('Invalid number_format')
        if ('horizontal_alignment' in op
                and op['horizontal_alignment'] not in ALIGNMENTS):
            raise ValueError('Invalid alignment')
        if 'column_width' in op:
            if (type(op['column_width']) not in (int, float)
                    or not 0 <= op['column_width'] <= 255):
                raise ValueError('Invalid column_width')
            if bounds['columns'] != 1:
                raise ValueError('column_width requires a one-column range')
    elif kind == 'calculate_range':
        if set(op) - (base | {'spill_policy'}):
            raise ValueError('Invalid calculate_range fields')
    return op


def _write(wb, sh, op):
    kind = op['kind']
    policy = op.get('spill_policy', 'native')
    if kind in ('set_values', 'set_formulas'):
        bounds = parse_address(op['address'], max_cells=MAX_CELLS)
        rng = sh.Range(bounds['address'])
        key = 'values' if kind == 'set_values' else 'formulas'
        matrix = _shape(op[key], int(rng.Rows.Count), int(rng.Columns.Count), key)
        for row_index, row in enumerate(matrix, 1):
            for column_index, value in enumerate(row, 1):
                cell = rng.Cells.Item(row_index, column_index)
                if kind == 'set_values':
                    if type(value) is str:
                        number_format = cell.NumberFormat
                        _set_literal(cell, value)
                        if (cell.Value2 != value or cell.HasFormula is not False
                                or cell.NumberFormat != number_format):
                            raise RuntimeError(
                                'Literal text setter postcondition mismatch'
                            )
                    else:
                        setattr(cell, 'Value2', value)
                else:
                    setattr(cell, 'Formula', value)
        after = read_edit_range(wb, op['sheet'], op['address'], policy)
        actual = (after['snapshot']['data']['values']
                  if kind == 'set_values' else after['formulas'])
        if actual != matrix:
            raise RuntimeError('Write postcondition mismatch')
        return after
    if kind == 'format_range':
        bounds = parse_address(op['address'], max_cells=MAX_CELLS)
        rng = sh.Range(bounds['address'])
        if int(rng.Columns.Count) != 1 and 'column_width' in op:
            raise ValueError('column_width requires a one-column range')
        before = read_edit_range(wb, op['sheet'], op['address'], policy)
        if 'number_format' in op:
            rng.NumberFormat = op['number_format']
        for key, attribute in (
            ('bold', 'Bold'), ('font_size', 'Size'),
            ('font_name', 'Name'), ('color', 'Color'),
        ):
            if key in op:
                setattr(rng.Font, attribute, op[key])
        if 'fill_color' in op:
            rng.Interior.Color = op['fill_color']
        if 'horizontal_alignment' in op:
            rng.HorizontalAlignment = ALIGNMENTS[op['horizontal_alignment']]
        if 'column_width' in op:
            rng.EntireColumn.ColumnWidth = op['column_width']
        after = read_edit_range(wb, op['sheet'], op['address'], policy)
        if (after['snapshot']['data'] != before['snapshot']['data']
                or after['formulas'] != before['formulas']):
            raise RuntimeError('Formatting changed cell content')
        for row in after['styles']:
            for style in row:
                for key in STYLE_FIELDS:
                    if key not in op:
                        continue
                    actual, expected = style[key], op[key]
                    if key == 'horizontal_alignment':
                        expected = ALIGNMENTS[expected]
                    if key == 'column_width':
                        if abs(float(actual) - float(expected)) > 1 / 64:
                            raise RuntimeError('Column width postcondition mismatch')
                    elif key == 'font_size':
                        if abs(float(actual) - float(expected)) > 0.05:
                            raise RuntimeError('Font size postcondition mismatch')
                    elif actual != expected:
                        raise RuntimeError('Style postcondition mismatch: ' + key)
        return {'verified': True, 'address': op['address'], 'after': after}
    if kind == 'calculate_range':
        bounds = parse_address(op['address'], max_cells=MAX_CELLS)
        sh.Range(bounds['address']).Calculate()
        return read_edit_range(wb, op['sheet'], op['address'], policy)
    if kind == 'save_copy':
        path = _save_path(op['path'])
        _save_preconditions(wb, path)
        saved_before = bool(wb.Saved)
        directory = Path(tempfile.mkdtemp(prefix='.excel-tools-', dir=path.parent))
        stage = directory / path.name
        published = False
        try:
            wb.SaveCopyAs(str(stage))
            if not stage.is_file() or stage.stat().st_size <= 0:
                raise RuntimeError('SaveCopyAs produced no non-empty file')
            if bool(wb.Saved) != saved_before:
                raise RuntimeError('SaveCopyAs changed the source Saved state')
            os.link(stage, path)
            published = True
        except Exception as exc:
            raise WriteOutcomeUnknown(
                f'Copy write/publication outcome unknown; inspect {path} and '
                f'{stage}: {exc}'
            ) from exc
        finally:
            if published:
                try:
                    stage.unlink()
                    directory.rmdir()
                except OSError:
                    pass
        return {
            'path': str(path),
            'bytes': path.stat().st_size,
            'source_saved': bool(wb.Saved),
        }
    raise ValueError('Unsupported operation')


def apply_request(bridge, request, commit=False):
    """Preview by default; commit rebinds, rechecks evidence, writes once, verifies.

    Raises WriteOutcomeUnknown after an attempted write fails or its local
    postcondition fails. Never retries, rolls back, or saves in place.
    """
    if type(commit) is not bool:
        raise ValueError('commit must be explicit bool')
    request = json.loads(json.dumps(request, ensure_ascii=False, allow_nan=False))
    op = _validate(request)
    wb, identity = _target(bridge, request['target'])
    if op['kind'] == 'save_copy':
        path = _save_path(op['path'])
        _save_preconditions(wb, path)
        evidence = _save_evidence(wb, identity, path)
        sh = None
    else:
        policy = op.get('spill_policy', 'native')
        sh = _sheet(wb, op['sheet'])
        current = read_edit_range(wb, op['sheet'], op['address'], policy)
        evidence = {
            'sheet': current['sheet'],
            'address': current['address'],
            'sha256': current['sha256'],
            'checks': current['structural_checks'],
            'spill_policy': policy,
        }
        if current['sha256'] != op['expected_sha256']:
            raise ValueError('Evidence hash mismatch')
    result = {
        'target': identity,
        'operation': op,
        'committed': False,
        'saved_in_place': False,
        'evidence': evidence,
    }
    if not commit:
        return result
    wb, identity = _target(bridge, request['target'], '提交前')
    if op['kind'] == 'save_copy':
        path = _save_path(op['path'])
        _save_preconditions(wb, path)
        fresh = _save_evidence(wb, identity, path)
        sh = None
    else:
        policy = op.get('spill_policy', 'native')
        sh = _sheet(wb, op['sheet'])
        current = read_edit_range(wb, op['sheet'], op['address'], policy)
        fresh = {
            'sheet': current['sheet'],
            'address': current['address'],
            'sha256': current['sha256'],
            'checks': current['structural_checks'],
            'spill_policy': policy,
        }
        if current['sha256'] != op['expected_sha256']:
            raise ValueError('Evidence changed before commit')
    if fresh != evidence:
        raise ValueError('Target/evidence changed before commit')
    try:
        data = _write(wb, sh, op)
        rebind_after_write(bridge, request['target'], label='Excel', message=WRITE_LOST)
    except WriteOutcomeUnknown:
        raise
    except Exception as exc:
        raise WriteOutcomeUnknown(
            'Write attempted; unknown/partial outcome. No retry or rollback: '
            + str(exc)
        ) from exc
    return {**result, 'committed': True, 'data': data}


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command', choices=('read', 'capabilities', 'apply'))
    parser.add_argument('--request', type=Path)
    parser.add_argument('--pid', type=int)
    parser.add_argument('--hwnd', type=int)
    parser.add_argument('--document')
    parser.add_argument('--sheet')
    parser.add_argument('--range', dest='address')
    parser.add_argument(
        '--spill-policy', choices=SPILL_POLICIES, default='native'
    )
    parser.add_argument('--commit', action='store_true')
    parser.add_argument('--timeout', type=int, default=9)
    parser.add_argument('--worker', action='store_true', help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if not 1 <= args.timeout <= 120:
        parser.error('--timeout must be 1..120 seconds')
    if args.commit and args.command != 'apply':
        parser.error('--commit applies only to apply')
    if not args.worker:
        forwarded = list(sys.argv[1:] if argv is None else argv)
        try:
            return subprocess.run(
                [sys.executable, str(Path(__file__).resolve()), *forwarded, '--worker'],
                timeout=args.timeout,
            ).returncode
        except subprocess.TimeoutExpired:
            print(json.dumps({
                'status': 'unknown' if args.commit else 'error',
                'error': 'Worker timed out; Excel may continue processing. '
                         'Inspect the workbook before another request; no retry.',
            }), file=sys.stderr)
            return 2
    try:
        bridge = ExcelBridge()
        if args.command in ('read', 'capabilities'):
            if (args.commit or args.pid is None or args.hwnd is None
                    or not args.sheet or not args.address):
                raise ValueError(
                    'read requires --pid --hwnd --sheet --range; --commit forbidden'
                )
            wb, target = bridge.select(args.pid, args.hwnd, args.document)
            sh = wb.Worksheets.Item(args.sheet)
            if args.command == 'capabilities':
                from excel_capabilities import probe_capabilities
                reader = probe_capabilities
            else:
                reader = read_edit_range
            data = reader(
                wb,
                {'name': args.sheet, 'code_name': str(sh.CodeName)},
                args.address,
                args.spill_policy,
            )
            _target(bridge, target)
            out = {'target': target, 'data': data}
        else:
            if args.request is None:
                raise ValueError('apply requires --request')
            out = apply_request(
                bridge,
                json.loads(args.request.read_text(encoding='utf-8-sig')),
                args.commit,
            )
        print(json.dumps({'status': 'ok', **out}, ensure_ascii=False, allow_nan=False))
        return 0
    except Exception as exc:
        print(json.dumps({
            'status': 'unknown' if isinstance(exc, WriteOutcomeUnknown) else 'error',
            'type': type(exc).__name__,
            'error': ('Capability probe failed; inspect error type/HRESULT; no writes.'
                      if args.command == 'capabilities' else str(exc)),
            **({'hresult': getattr(exc, 'hresult', None)}
               if args.command == 'capabilities'
               and type(getattr(exc, 'hresult', None)) is int else {}),
        }, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
