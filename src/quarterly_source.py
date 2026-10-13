"""Optional, read-only SEC EX-99.1 standalone GAAP table supplement.

Supports explicit Three Months Ended tables with dates, stated millions or
thousands units, total EPS and weighted shares. Unsupported layouts stay absent;
annual/YTD averages are never converted. No issuer-specific concepts or values.
"""
from datetime import date
import hashlib
import json
from pathlib import Path
import re
from urllib.parse import urlparse

from lxml import html

NONADDITIVE = frozenset({'Basic Shares', 'Diluted Shares', 'Basic EPS', 'Diluted EPS'})
ANCHORS = ('Revenue', 'Gross Profit', 'Operating Income', 'Pre-tax Income', 'Income Tax', 'Net Income')
MONTHS = {name: i for i, name in enumerate(
    ('jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'), 1)}


def _date_columns(table):
    """Tie each date interval to its duration group through HTML colspan.

    Repeated dates have no duration meaning. Year-first and quarter-first layouts
    are both supported; rowspans or unmatched date intervals are rejected.
    """
    groups = []
    for row in table.xpath('./tr|./tbody/tr')[:8]:
        cells = []
        column = 0
        for cell in row.xpath('./td|./th'):
            if cell.get('rowspan', '1') != '1':
                return [], []
            try:
                span = int(cell.get('colspan', '1'))
            except ValueError:
                return [], []
            if span < 1:
                return [], []
            text = ' '.join(cell.text_content().split())
            cells.append((column, column + span, text))
            column += span
        durations = [(start, stop, 'quarter' if text.lower() == 'three months ended' else 'cumulative')
                     for start, stop, text in cells
                     if re.fullmatch(r'(?:three months|six months|nine months|year|years) ended', text, re.I)]
        if durations:
            groups = durations
        dates, quarters = [], []
        for start, stop, text in cells:
            m = re.fullmatch(r'([A-Za-z]+)\.?\s*(\d{1,2}),?\s*(20\d{2})', text)
            if m and m[1][:3].lower() in MONTHS:
                matches = [kind for left, right, kind in groups if left <= start and stop <= right]
                if len(matches) != 1:
                    return [], []
                if matches[0] == 'quarter':
                    quarters.append(len(dates))
                dates.append(date(int(m[3]), MONTHS[m[1][:3].lower()], int(m[2])).isoformat())
        if len(dates) >= 2:
            return dates, quarters
    return [], []


def _numbers(cells):
    result = []
    i = 0
    while i < len(cells):
        token = cells[i]
        i += 1
        if token in ('$', 'US$'):
            continue
        if token.startswith('(') and not token.endswith(')') and i < len(cells) and cells[i] == ')':
            token += ')'
            i += 1
        if token in ('—', '–', '-', '\x97'):
            result.append(0.0)  # explicit reported dash only; a missing cell is not zero
        elif re.fullmatch(r'\(?-?[\d,]+(?:\.\d+)?\)?', token):
            result.append(float(token.strip('()').replace(',', '')) * (-1 if token.startswith('(') else 1))
        else:
            return []
    return result


def _scope_label(label):
    return ' '.join(re.sub(r'[,()]', ' ', label.lower().strip().rstrip(':')).split())


def _key(label, mode):
    s = label.lower().strip().rstrip(':')
    full = _scope_label(label)
    if s in ('revenue', 'revenues', 'net revenue', 'net revenues', 'sales', 'net sales'):
        return 'Revenue'
    if s in ('total cost of sales', 'total cost of revenue', 'total cost of revenues'):
        return 'Cost of Revenue'
    if s in ('gross profit', 'gross margin'):
        return 'Gross Profit'
    if re.fullmatch(r'operating (?:income|loss)(?: \(loss\)| \(income\))?', s):
        return 'Operating Income'
    if 'before' in s and ('income tax' in s or 'for income taxes' in s):
        return 'Pre-tax Income'
    if ('income tax' in s or 'for income taxes' in s) and 'before' not in s:
        return 'Income Tax'
    if full in ('equity income', 'equity loss', 'equity income in investee',
                 'equity loss in investee', 'income loss from equity method investments'):
        return 'Equity Method Income/(Loss)'
    if re.fullmatch(r'(?:(?:income|loss|income loss|loss income) from )?discontinued operations net of tax', full):
        return 'Discontinued Operations, net of tax'
    if re.fullmatch(r'(?:income|loss|income loss|loss income) from continuing operations net of tax', full):
        return 'Net Income from Continuing Operations'
    if full in ('net income', 'net loss', 'net income loss', 'net loss income'):
        return 'Net Income'
    if s in ('basic', 'diluted') and mode:
        return s.title() + (' Shares' if mode == 'shares' else ' EPS')
    if re.fullmatch(r'(basic|diluted) earnings(?: \(loss\))? per share', s):
        return s.split()[0].title() + ' EPS'
    return None


def _table_records(table, index, source):
    rows = [[' '.join(c.text_content().split()) for c in tr.xpath('./td|./th')
             if ' '.join(c.text_content().split())] for tr in table.xpath('./tr|./tbody/tr')]
    text = ' '.join(' '.join(row) for row in rows)
    if len(text) > 15000 or not re.search(r'three months ended', text, re.I):
        return []
    heading = ' '.join(table.xpath('preceding::text()')[-8:])[-600:]
    if re.search(r'non[ -]?gaap|reconcil|adjusted', text, re.I):
        return []
    if not re.search(r'consolidated statements? of (?:operations|income|earnings)', heading + text[:300], re.I):
        return []
    if re.search(r'non[ -]?gaap.*statements?', heading, re.I):
        return []
    unit_text = heading + ' ' + text[:300]
    # Monetary and share units may differ. This limited reader does not infer
    # mixed units; explicitly excepted share counts are also unsupported.
    if re.search(r'\bshares\b', re.sub(r'per share', '', unit_text, flags=re.I), re.I):
        return []
    unit = re.search(r'\b(millions|thousands)\b.*?except.*?per share', unit_text, re.I)
    if not unit:
        return []
    scale = 1_000_000 if unit[1].lower() == 'millions' else 1_000
    dates, quarter_indices = _date_columns(table)
    if not dates or not quarter_indices:
        return []
    matrix, labels = {}, {}
    mode = None
    for row in rows:
        if not row:
            continue
        label = row[0]
        low = label.lower()
        if re.search(r'shares used in per share calculation|weighted average (?:number of )?shares', low) and len(row) == 1:
            mode = 'shares'
            continue
        if 'earnings' in low and 'per share' in low and len(row) == 1:
            mode = ('eps' if re.fullmatch(r'earnings(?: \(loss\))? per (?:common )?share:?', low) else None)
            continue
        if mode == 'eps' and len(row) == 1 and re.search(r'continuing|discontinued|attribut|noncontrolling', low):
            mode = None
            continue
        key = _key(label, mode)
        if key is None:
            continue
        values = _numbers(row[1:])
        if len(values) != len(dates):
            continue
        if key in matrix and matrix[key] != values:
            return []  # ambiguous repeated rows, not first-row-wins
        matrix[key], labels[key] = values, label
    if not all(k in matrix for k in NONADDITIVE | set(ANCHORS)):
        return []
    records = []
    for col in quarter_indices:
        end = dates[col]
        if any(matrix[k][col] <= 0 for k in ('Basic Shares', 'Diluted Shares')):
            continue
        reported = {k: v[col] for k, v in matrix.items()}
        values = {k: v * (1 if k.endswith('EPS') else scale) for k, v in reported.items()}
        equity_label = _scope_label(labels.get('Equity Method Income/(Loss)', ''))
        if equity_label in ('equity loss', 'equity loss in investee'):
            values['Equity Method Income/(Loss)'] = -abs(values['Equity Method Income/(Loss)'])
        records.append(dict(period_end=end, source_current_end=max(dates[c] for c in quarter_indices),
                            is_comparative=end != max(dates[c] for c in quarter_indices),
                            values=values, reported_values=reported, labels=labels, table=index,
                            column=col + 1, scale=scale, **source))
    return records


def read_quarter_source_pack(folder, expected_cik):
    folder = Path(folder).resolve()
    manifest = json.loads((folder / 'manifest.json').read_text(encoding='utf-8'))
    if not expected_cik or int(manifest['cik']) != int(expected_cik):
        raise ValueError('Quarter source pack issuer does not match the requested company')
    records = {}
    # Earliest supported official table wins. A later comparative can supply
    # an otherwise unsupported quarter; the publication and comparative basis
    # remain explicit. Same-disclosure conflicts are rejected.
    for source in sorted(manifest['sources'], key=lambda s: (s['filing_date'], s['accession'], s['filename'])):
        if source.get('document_type') != 'EX-99.1' or source.get('form') != '8-K':
            continue
        url = urlparse(source['url'])
        accession = source['accession']
        if url.scheme != 'https' or url.hostname != 'www.sec.gov' or not re.fullmatch(r'\d{10}-\d{2}-\d{6}', accession):
            raise ValueError('Invalid SEC quarterly source identity')
        prefix = f'/Archives/edgar/data/{int(expected_cik)}/{accession.replace("-", "")}/'
        if not url.path.startswith(prefix):
            raise ValueError('Quarterly source URL does not match issuer/accession')
        path = (folder / source['filename']).resolve()
        if not path.is_relative_to(folder):
            raise ValueError('Quarterly source path escapes the pack')
        data = path.read_bytes()
        if hashlib.sha256(data).hexdigest() != source['sha256']:
            raise ValueError('Quarterly source hash mismatch')
        try:
            doc = html.fromstring(data.decode('utf-8'))
        except UnicodeDecodeError:
            doc = html.fromstring(data)
        provenance = {k: source[k] for k in ('url', 'accession', 'sha256', 'filename', 'filing_date')}
        for index, table in enumerate(doc.xpath('//table'), 1):
            for record in _table_records(table, index, provenance):
                end = record['period_end']
                previous = records.get(end)
                if previous and previous['filing_date'] == record['filing_date'] and previous['values'] != record['values']:
                    raise ValueError('Conflicting official standalone quarterly tables')
                records.setdefault(end, record)
    return records


def supplement_quarter_table(table, records):
    """Fill only missing cells, with same-period statement anchors as scope guards.

    Each returned trace identifies the exact raw table/column. A source disagreement
    is reported and leaves existing values intact. Annual tables are not inputs.
    """
    from dataclasses import replace
    values = [list(row) for row in table.values]
    indices = {name: i for i, name in enumerate(table.concepts)}
    trace = []
    for col, end in enumerate(table.period_ends):
        record = records.get(end)
        if record is None:
            continue
        differences = [key for key in ANCHORS if key not in indices or values[indices[key]][col] is None
                       or abs(values[indices[key]][col] - record['values'][key]) > 1e-5]
        if differences:
            trace.append(dict(period=table.quarter_labels[col], status='scope-conflict', metrics=differences,
                              **{k: v for k, v in record.items() if k not in ('values', 'labels', 'reported_values')}))
            continue
        for key, raw_value in record['values'].items():
            if key not in indices:
                continue
            row = indices[key]
            existing = values[row][col]
            if existing is None:
                values[row][col] = raw_value
                trace.append(dict(period=table.quarter_labels[col], metric=key, value=raw_value,
                                  raw_label=record['labels'][key], reported_value=record['reported_values'][key], status='filled',
                                  **{k: v for k, v in record.items() if k not in ('values', 'labels', 'reported_values')}))
            elif key in NONADDITIVE and abs(existing - raw_value) > 1e-5:
                trace.append(dict(period=table.quarter_labels[col], metric=key, status='value-conflict',
                                  existing=existing, official=raw_value,
                                  **{k: v for k, v in record.items() if k not in ('values', 'labels', 'reported_values')}))
    return replace(table, values=values), trace
