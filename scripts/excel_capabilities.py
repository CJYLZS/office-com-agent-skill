"""Read-only, target-local Excel capability evidence; never a write permit.

No setters, scratch cells, application setting changes, or version-based feature
claims. Formula2 read support does not prove setter or dynamic-array support.
"""
import time

from excel_com import parse_address
from excel_tools import MAX_CELLS, SPILL_POLICIES, _sheet, _structure


def _error(exc):
    result = {'type': type(exc).__name__}
    code = getattr(exc, 'hresult', None)
    if type(code) is int:
        result['hresult'] = code
    return result


def _resolve_name(obj, name):
    from comtypes.automation import IDispatch
    disp = obj.QueryInterface(IDispatch)
    ids = disp.GetIDsOfNames(name)
    code = ids[0] if isinstance(ids, (list, tuple)) else ids
    return type(code) is int and code != -1


def _property(obj, name):
    """Resolve dispatch name separately from actually invoking the getter."""
    result = {'name_resolved': None, 'readable': False,
              'write_tested': False, 'setter_supported': None}
    try:
        result['name_resolved'] = _resolve_name(obj, name)
    except Exception as exc:
        result['resolution_error'] = _error(exc)
    try:
        value = getattr(obj, name)
        if type(value) not in (bool, int, float, str, type(None)):
            raise TypeError('Property did not return a scalar')
        result.update(readable=True, value_type=type(value).__name__)
        # Do not disclose formula text or cell contents in a capability report.
        if name in ('HasSpill', 'HasArray', 'HasFormula'):
            result['boolean_value'] = value if type(value) is bool else None
    except Exception as exc:
        result['read_error'] = _error(exc)
    return result


def probe_capabilities(wb, sheet_spec, address, spill_policy='native'):
    """Check one bounded target, read representative getters and current guards.

    The guard scan covers every target cell (<=100); getter capability samples
    only its top-left cell. No policy fallback, no evidence hash/write permit.
    The CLI supplies a 9-second watchdog; direct imports have no hard timeout.
    """
    started = time.perf_counter()
    if spill_policy not in SPILL_POLICIES:
        raise ValueError('Unsupported spill_policy')
    bounds = parse_address(address, max_cells=MAX_CELLS)
    sh = _sheet(wb, sheet_spec)
    rng = sh.Range(bounds['address'])
    if (int(rng.Rows.Count) != bounds['rows']
            or int(rng.Columns.Count) != bounds['columns']
            or int(rng.Areas.Count) != 1):
        raise ValueError('Range shape mismatch')
    saved_before = wb.Saved
    runtime = {}
    for name in ('Version', 'Build'):
        try:
            value = getattr(wb.Application, name)
            if type(value) not in (str, int, float):
                raise TypeError('Non-scalar runtime metadata')
            runtime[name.lower()] = str(value)[:80]
        except Exception as exc:
            runtime[name.lower()] = None
            runtime[name.lower() + '_error'] = _error(exc)
    cell = rng.Cells.Item(1, 1)
    properties = {name: _property(cell, name) for name in
                  ('Value2', 'Formula', 'Formula2', 'HasFormula', 'HasArray', 'HasSpill')}
    safety = _structure(wb, sh, rng, bounds, spill_policy, diagnostic=True)
    # Editing evidence retains formulas for optimistic conflict checks. A capability
    # report only needs inventory size, not content from outside its target.
    proof = safety['checks'].get('spill_proof')
    if proof is not None:
        formulas = proof.pop('formula_cells')
        proof['formula_cell_count'] = len(formulas)
    # The historical editing guard permits confirmed pre-16 versions. Capability
    # evidence must not turn that version assumption into native HasSpill proof.
    if any(value == 'not_applicable_legacy' for value in safety['checks'].values()):
        safety['blockers'].append('legacy_version_assumption_not_native_spill_evidence')
        safety['checked_structure_clear'] = False
    saved_after = wb.Saved
    return {
        'schema_version': 1, 'runtime': runtime,
        'sheet': dict(sheet_spec), 'address': bounds['address'],
        'sample': {'relative_row': 1, 'relative_column': 1, 'properties': properties},
        'safety': safety, 'spill_policy': spill_policy,
        'dynamic_arrays_validated': False, 'write_authorized': False,
        'saved_before': saved_before if type(saved_before) is bool else None,
        'saved_after': saved_after if type(saved_after) is bool else None,
        'saved_state_unchanged': type(saved_before) is bool and
                                 type(saved_after) is bool and saved_before == saved_after,
        'elapsed_seconds': round(time.perf_counter() - started, 3),
        'warnings': [
            'Name resolution does not prove getter or setter support.',
            'Top-left getter sample is not a workbook-wide capability guarantee.',
            'Clear structural checks are not edit authorization; re-read evidence and preview.',
            'No automatic policy fallback. Version/build are descriptive only.',
            'Concurrent users, events and add-ins are not locked out.',
        ],
    }
