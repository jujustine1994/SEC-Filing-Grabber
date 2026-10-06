"""Compare Revenue audits while retaining fixed-section and source-row identity."""
import compare_fiscal_audits as base
from collections import Counter
from verify_fixed_financial_rows import source_contract


def cells(result):
    values, labels, sources = base.cells(result)
    identities = {}
    contract = source_contract()
    for table in result['tables']:
        name = table['sheet_name']
        concepts = table['concepts']
        fixed = {}
        if name in ('Data_Financials(Q)', 'Data_Financials(Y)'):
            for section, key in [('Income Statement', 'IS_TEMPLATE'),
                                 ('Balance Sheet', 'BS_TEMPLATE'), ('Cash Flow', 'CF_TEMPLATE')]:
                start = concepts.index(section) + 1
                names = contract[key]
                assert concepts[start:start + len(names)] == names, (name, section)
                fixed.update({start + i: f'template:{section}:{i}' for i in range(len(names))})
        occurrences = Counter()
        source_occurrences = Counter()
        for row, concept in enumerate(concepts):
            occurrence = occurrences[concept]
            occurrences[concept] += 1
            source_labels = table.get('labels') or []
            source = source_labels[row] if row < len(source_labels) else ''
            if row in fixed:
                identity = fixed[row]
            elif name.startswith('Data_Financials') and source:
                source_key = (concept, source)
                identity = f'source:{source}:{source_occurrences[source_key]}'
                source_occurrences[source_key] += 1
            else:
                identity = f'row:{occurrence}'
            identities[(name, concept, occurrence)] = identity
    stable = {}
    for key, value in values.items():
        identity = identities[key[:3]]
        stable[(key[0], key[1], identity, *key[3:])] = value
    return stable, labels, sources


def compare(before, after):
    result = base.compare(before, after, cell_reader=cells)
    for category in ('value_changes', 'header_changes', 'metadata_changes'):
        for change in result[category]:
            change['row_identity'] = change.pop('occurrence')
    return result


if __name__ == '__main__':
    raise SystemExit(base.main(comparator=compare))
