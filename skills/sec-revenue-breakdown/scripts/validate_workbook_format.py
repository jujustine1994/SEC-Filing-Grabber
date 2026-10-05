"""Check the shared revenue-breakdown XLSX sheet contract without dependencies."""
import argparse
import xml.etree.ElementTree as ET
import zipfile

EXPECTED = ['Quarterly History', 'Annual History', 'Coverage', 'Sources']

def validate(path):
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read('xl/workbook.xml'))
    actual = [s.attrib['name'] for s in root.findall('{*}sheets/{*}sheet')]
    if actual != EXPECTED:
        raise ValueError(f'Expected {EXPECTED}; got {actual}')
    return actual

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('xlsx')
    args = parser.parse_args()
    print('PASS:', ', '.join(validate(args.xlsx)))
