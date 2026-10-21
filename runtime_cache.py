"""Bounded, session-local reuse of source parsing and financial calculations."""
import hashlib
from core import load_workbook, build_plan


def parsed_source(state, raw):
    digest = hashlib.sha256(raw).hexdigest()
    cached = state.get('source_bundle')
    if cached is None or cached[0] != digest:
        cached = (digest, *load_workbook(raw))
        state['source_bundle'] = cached
    return cached


def monthly_plan(state, store, records, headers, source_hash, month, contact, errors):
    signature = (source_hash, month, contact, store.revision())
    cached = state.get('monthly_plan_cache')
    if cached is None or cached[0] != signature:
        plan = build_plan(records, headers, store, month, contact, source_hash, errors=errors, snapshot=True)
        if store.revision() != signature[-1]:
            plan = build_plan(records, headers, store, month, contact, source_hash, errors=errors, snapshot=True)
        state['monthly_plan_cache'] = ((source_hash, month, contact, store.revision()), plan)
    return state['monthly_plan_cache'][1]


def cached_analytics(state, store, source_hash, records, headers, months, contacts, calculate):
    signature = (source_hash, tuple(sorted(months)), tuple(sorted(contacts)), store.revision())
    cached = state.get('analytics_cache')
    if cached is None or cached[0] != signature:
        result = calculate(records, headers, store, months, contacts)
        if store.revision() != signature[-1]:
            result = calculate(records, headers, store, months, contacts)
        state['analytics_cache'] = ((source_hash, tuple(sorted(months)), tuple(sorted(contacts)), store.revision()), result)
    return state['analytics_cache'][1]


def relevant_archives(files, months):
    """New exports advertise scope; keep newest per person/month and settings.

    Legacy archives have no scope metadata, so retain them for compatibility.
    """
    selected, seen = [], set()
    for index, file in enumerate(files):
        properties = file.get('appProperties', {})
        month, person = properties.get('report_month'), properties.get('report_person')
        if not month or not person:
            selected.append(file)
        elif (month in months and (month,person) not in seen) or index == 0:
            selected.append(file)
        seen.add((month, person))
    return selected


def drive_objects(state, folder, files, loader, key, limit, newest_first=True):
    cache = state.setdefault(key, {})
    if cache.get('folder') != folder:
        cache.clear(); cache['folder'] = folder; cache['items'] = {}
    items = cache['items']
    live = {f['id'] for f in files}
    for old in set(items) - live:
        items.pop(old, None)
    result = []
    for file in files:
        # Missing version metadata must never be treated as immutable.
        signature = (file.get('modifiedTime'), file.get('size'),
                     repr(sorted(file.get('appProperties', {}).items())))
        cached = items.get(file['id'])
        if not file.get('modifiedTime') or cached is None or cached[0] != signature:
            value = loader(file)
            items[file['id']] = (signature, value)
        else:
            value = cached[1]
        result.append(value)
    keep = {f['id'] for f in (files[:limit] if newest_first else files[-limit:])}
    for old in set(items) - keep:
        items.pop(old, None)
    return result
