"""Read-only verification against downloaded original SEC XBRL instances.

Files: TICKER-accession-URL_basename under the supplied folder. No DB access,
network calls, standardized concepts or template selection functions are used.
"""
import argparse
import hashlib
import json
from decimal import Decimal
from pathlib import Path
import xml.etree.ElementTree as ET


def verify(case, folder):
    path = folder / (case['ticker'] + '-' + case['accession'] + '-' + case['url'].split('/')[-1])
    assert hashlib.sha256(path.read_bytes()).hexdigest() == case['instance_sha256'], path
    namespaces = dict(item for _, item in ET.iterparse(path, events=['start-ns']))
    root = ET.parse(path).getroot()
    ns = {'x': 'http://www.xbrl.org/2003/instance'}
    usd = set()
    for unit in root.findall('x:unit', ns):
        measure = unit.findtext('x:measure', namespaces=ns)
        if len(unit) != 1 or not measure or ':' not in measure:
            continue
        prefix, name = measure.split(':', 1)
        if name == 'USD' and namespaces.get(prefix) == 'http://www.xbrl.org/2003/iso4217':
            usd.add(unit.get('id'))
    contexts = set()
    for context in root.findall('x:context', ns):
        if any(node.tag.rsplit('}', 1)[-1] in ('explicitMember', 'typedMember') for node in context.iter()):
            continue
        if (context.findtext('x:period/x:startDate', namespaces=ns) == case['start']
                and context.findtext('x:period/x:endDate', namespaces=ns) == case['end']):
            contexts.add(context.get('id'))
    values = set()
    for fact in root:
        if not fact.tag.startswith('{'):
            continue
        uri, name = fact.tag[1:].split('}', 1)
        if ('/us-gaap/' in uri and name == case['concept']
                and fact.get('contextRef') in contexts and fact.get('unitRef') in usd
                and fact.text):
            values.add(Decimal(fact.text))
    assert values == {Decimal(case['value'])}, (case['ticker'], values)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('folder', type=Path)
    args = parser.parse_args()
    evidence = Path(__file__).resolve().parents[1] / 'docs/superpowers/evidence/revenue-source-examples.json'
    cases = json.loads(evidence.read_text(encoding='utf-8'))
    for case in cases:
        verify(case, args.folder)
        print(case['ticker'], case['start'], case['end'], case['value'], 'OK')
    print(len(cases), 'original SEC examples verified')


if __name__ == '__main__':
    main()
