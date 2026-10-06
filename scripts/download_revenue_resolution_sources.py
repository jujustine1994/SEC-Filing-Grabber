"""Download certified SEC instance/calculation evidence to a separate folder.

Never changes the financial database or replaces an existing evidence file.
Only the SEC identity setting is read; URLs and SHA-256 come from the fixture.
"""
import argparse
import hashlib
import json
from pathlib import Path
import sys
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('folder', type=Path)
    args = parser.parse_args()
    sys.path.insert(0, str(ROOT / 'src'))
    from config import load_config
    from edgar import set_identity
    from edgar.httprequests import download_file
    set_identity(load_config()['identity'])
    args.folder.mkdir(parents=True, exist_ok=True)
    cases = json.loads((ROOT / 'tests/fixtures/revenue-resolution-cases.json').read_text(encoding='utf-8'))
    count = 0
    for case in cases:
        sources = [s for s in case['sources'] if s['url']==case['url'] or s['url'].endswith('_cal.xml')]
        for source in sources:
            url = urlparse(source['url'])
            if url.scheme!='https' or url.hostname!='www.sec.gov' or not url.path.startswith('/Archives/edgar/data/'):
                raise ValueError('Evidence must come from SEC Archives')
            path = args.folder / f"{case['ticker']}-{case['accession']}-{url.path.rsplit('/', 1)[-1]}"
            if not path.exists():
                download_file(source['url'], path=path)
            if hashlib.sha256(path.read_bytes()).hexdigest()!=source['sha256']:
                raise RuntimeError(f'Evidence hash mismatch; existing file is not replaced: {path.name}')
            count += 1
        print(case['ticker'], case['accession'], 'evidence ready', flush=True)
    print(count, 'SEC source files verified')


if __name__ == '__main__':
    main()
