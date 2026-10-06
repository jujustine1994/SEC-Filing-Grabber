"""Verify Revenue repair examples directly against original SEC XML, offline.

Checks bytes, entity, non-dimensional duration, USD unit and exact QName. Never
imports the selection code or reads/writes the financial database.
"""
import argparse
from decimal import Decimal
import hashlib
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET

EVIDENCE = Path(__file__).resolve().parents[1] / 'tests/fixtures/revenue-resolution-cases.json'
NS = {'x': 'http://www.xbrl.org/2003/instance'}


def verify(case, folder):
    path = folder / f"{case['ticker']}-{case['accession']}-{case['url'].rsplit('/', 1)[-1]}"
    assert hashlib.sha256(path.read_bytes()).hexdigest() == case['instance_sha256'], path
    prefixes = dict(item for _, item in ET.iterparse(path, events=['start-ns']))
    root = ET.parse(path).getroot()
    cik = int(re.search(r'/data/(\d+)/', case['url']).group(1))
    usd = set()
    for unit in root.findall('x:unit', NS):
        measure = unit.findtext('x:measure', namespaces=NS)
        if len(unit) == 1 and measure and ':' in measure:
            prefix, name = measure.split(':', 1)
            if name == 'USD' and prefixes.get(prefix) == 'http://www.xbrl.org/2003/iso4217':
                usd.add(unit.get('id'))
    total = Decimal(0)
    for expected in case['facts']:
        contexts = set()
        for context in root.findall('x:context', NS):
            if any(n.tag.rsplit('}', 1)[-1] in ('explicitMember', 'typedMember') for n in context.iter()):
                continue
            identifier = context.findtext('x:entity/x:identifier', namespaces=NS)
            if not identifier or int(identifier) != cik:
                continue
            if (context.findtext('x:period/x:startDate', namespaces=NS) == expected['start']
                    and context.findtext('x:period/x:endDate', namespaces=NS) == expected['end']):
                contexts.add(context.get('id'))
        local = expected['concept'].split('_', 1)[-1]
        qname = '{' + expected['namespace'] + '}' + local
        values = {Decimal(f.text) for f in root if f.tag == qname
                  and f.get('contextRef') in contexts and f.get('unitRef') in usd and f.text}
        assert values == {Decimal(expected['value'])}, (case['ticker'], local, values)
        total += Decimal(expected['value'])
    assert total == Decimal(case['expected_value']), case['ticker']
    link = 'http://www.xbrl.org/2003/linkbase'
    xl = 'http://www.w3.org/1999/xlink'
    calculations = set()
    for source in case['sources']:
        if not source['url'].endswith('_cal.xml'):
            continue
        cal_path = folder / f"{case['ticker']}-{case['accession']}-{source['url'].rsplit('/', 1)[-1]}"
        assert hashlib.sha256(cal_path.read_bytes()).hexdigest() == source['sha256'], cal_path
        for group in ET.parse(cal_path).findall(f'.//{{{link}}}calculationLink'):
            locators = {node.get(f'{{{xl}}}label'): node.get(f'{{{xl}}}href').split('#')[-1]
                        for node in group.findall(f'{{{link}}}loc')}
            for arc in group.findall(f'{{{link}}}calculationArc'):
                calculations.add((source['url'], group.get(f'{{{xl}}}role'),
                                  locators[arc.get(f'{{{xl}}}from')], locators[arc.get(f'{{{xl}}}to')],
                                  Decimal(arc.get('weight'))))
    for check in case['calculation_checks']:
        assert (check['url'], check['role'], check['parent'], check['child'],
                Decimal(check['weight'])) in calculations, check
    return total


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('folder', type=Path)
    args = parser.parse_args()
    cases = json.loads(EVIDENCE.read_text(encoding='utf-8'))
    for case in cases:
        print(case['ticker'], case['accession'], verify(case, args.folder), 'OK')
    print(len(cases), 'original SEC Revenue repair examples verified')


if __name__ == '__main__':
    main()
