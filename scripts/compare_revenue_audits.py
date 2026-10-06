"""Compare Revenue audits while retaining fixed-section and source-row identity."""
import compare_fiscal_audits as base


def compare(before, after):
    return base.compare(before, after)


if __name__ == '__main__':
    raise SystemExit(base.main())
