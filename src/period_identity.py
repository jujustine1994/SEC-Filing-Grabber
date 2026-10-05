"""Validated document fiscal focus from the same filing's SEC cover statement.

Focus applies only to the document's exact end date, never comparative periods.
Missing/conflicting/invalid metadata returns None; no month heuristic here.
"""
from datetime import date, timedelta
import re
import pandas as pd


# Exact, manually verified source errors. Never mutate source cache entries.
# The registrant explicitly defines its fiscal years in Item 1; DEI contradicts
# that definition. An accession AND original triple must match for correction.
_VERIFIED_FOCUS_CORRECTIONS = {
    '0001558370-24-004603': (('2024-02-03', 2024, 'FY'), ('2024-02-03', 2023, 'FY')),
    '0001558370-25-004267': (('2025-02-01', 2025, 'FY'), ('2025-02-01', 2024, 'FY')),
}


def corrected_focus(accession, focus):
    correction = _VERIFIED_FOCUS_CORRECTIONS.get(accession)
    return correction[1] if correction and focus == correction[0] else focus


def cover_focus(financials):
    try:
        statement=financials.cover()
        frame=statement.to_dataframe() if statement is not None else None
    except (AttributeError, TypeError, ValueError):
        return None
    if not isinstance(frame,pd.DataFrame) or 'concept' not in frame:
        return None
    candidates=set()
    names={'DocumentPeriodEndDate':'end','DocumentFiscalYearFocus':'year','DocumentFiscalPeriodFocus':'period'}
    for column in frame.columns:
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}(?:\s+\(\w+\))?',str(column)):
            continue
        values={}
        for concept,key in names.items():
            rows=frame[frame['concept'].astype(str).str.replace(':','_',regex=False).str.endswith('_'+concept)]
            # Dimensioned cover values must not establish a consolidated identity.
            if 'dimension' in rows:
                rows=rows[~rows['dimension'].fillna(False).astype(bool)]
            found={str(v).strip() for v in rows[column] if pd.notna(v)}
            if len(found)!=1: break
            values[key]=found.pop()
        if len(values)!=3: continue
        try:
            end=date.fromisoformat(values['end']).isoformat()
            year=float(values['year'])
            if not year.is_integer() or not 1900<=year<=2200: continue
            period=values['period'].upper()
            if period not in ('FY','Q1','Q2','Q3','Q4'): continue
            if str(column)[:10]!=end: continue
            candidates.add((end,int(year),period))
        except (ValueError,TypeError): continue
    return next(iter(candidates)) if len(candidates)==1 else None


def reported_label(focus,column,*,annual=False):
    if focus is None or str(column)[:10]!=focus[0]:
        return None
    end,year,period=focus
    is_annual=annual or str(column).endswith('(FY)')
    if is_annual:
        return f'FY{year}' if period=='FY' else None
    return f'FY{year}{period}' if period.startswith('Q') else None


def build_period_map(records):
    """Company-local authority map. Infer only complete, contiguous fiscal years.

    records: (SEC report date, form, optional validated cover focus).
    Do not infer a quarter from ordinal position when an intervening filing
    is missing, or carry a year through a fiscal transition/long date gap.
    """
    periods={}; conflict=set(); annual_dates=set(); quarterly_dates=set()
    def put(key,label):
        if key in periods and periods[key]!=label:
            conflict.add(key)
        else: periods[key]=label
    for end,form,focus in records:
        try: date.fromisoformat(end)
        except (ValueError,TypeError): continue
        if form in ('10-K','20-F'): annual_dates.add(end)
        if form in ('10-Q','6-K'): quarterly_dates.add(end)
        if focus and focus[0]==end:
            if focus[2]=='FY' and form in ('10-K','20-F'):
                put((end,True),f'FY{focus[1]}')
                put((end,False),f'FY{focus[1]}Q4')
            elif focus[2].startswith('Q') and form in ('10-Q','6-K'):
                put((end,False),f'FY{focus[1]}{focus[2]}')
    dates=sorted(annual_dates)
    # A chain is a continuous set of normal-length fiscal years. Named fiscal
    # years may differ from end-calendar years; verify known anchors agree.
    chains=[]
    for end in dates:
        if not chains or not 330 <= (date.fromisoformat(end)-date.fromisoformat(chains[-1][-1])).days <= 400:
            chains.append([])
        chains[-1].append(end)
    for chain in chains:
        offsets={int(periods[(end,True)][2:])-i for i,end in enumerate(chain)
                 if (end,True) in periods and (end,True) not in conflict}
        if len(offsets)!=1:
            if offsets:
                # Conflicting annual focus is not authority. Keeping its
                # direct mappings would deduplicate away a different year.
                ceiling=(date.fromisoformat(chain[-1])+timedelta(days=400)).isoformat()
                for key in list(periods):
                    if chain[0] <= key[0] <= ceiling:
                        periods.pop(key,None)
                        conflict.discard(key)
            continue
        offset=offsets.pop()
        for i,end in enumerate(chain):
            put((end,True),f'FY{offset+i}')
            put((end,False),f'FY{offset+i}Q4')
        for previous,end in zip(chain,chain[1:]):
            quarters=sorted(d for d in quarterly_dates if previous<d<end)
            if len(quarters)!=3: continue
            ends=[previous,*quarters,end]
            if not all(50 <= (date.fromisoformat(b)-date.fromisoformat(a)).days <= 140
                       for a,b in zip(ends,ends[1:])): continue
            year=periods[(end,True)][2:]
            proposed={(d,False):f'FY{year}Q{q}' for q,d in enumerate(quarters,1)}
            # A complete sequence bracketed by consistent annual anchors is
            # stronger than a single quarter's stale DEI focus. The selected
            # dates establish quarter ranks; annual anchors establish names.
            # Missing dates or inconsistent annual years never reach this path.
            for key,label in proposed.items():
                periods[key]=label
                conflict.discard(key)
    anchors=sorted(end for end in annual_dates if (end,True) in periods and (end,True) not in conflict)
    for end in quarterly_dates:
        key=(end,False)
        if key not in periods or end in annual_dates: continue
        previous=next((d for d in reversed(anchors) if d<end),None)
        following=next((d for d in anchors if d>end),None)
        expected=None
        if previous and following and 330 <= (date.fromisoformat(following)-date.fromisoformat(previous)).days <= 400:
            expected=int(periods[(following,True)][2:])
        elif previous and not following and 50 <= (date.fromisoformat(end)-date.fromisoformat(previous)).days <= 300:
            expected=int(periods[(previous,True)][2:])+1
        if expected is not None and int(periods[key][2:6])!=expected:
            # Do not rename/rank a partial year. Reject impossible focus and
            # let the caller retain its existing fallback, visibly uncertain.
            periods.pop(key,None)
    result={k:v for k,v in periods.items() if k not in conflict}
    identities={}
    for (end,annual),label in result.items():
        identity=(annual,label)
        if identity in identities and identities[identity]!=end:
            # Separate calendar chains can reuse a fiscal name during a
            # transition (LHX has two FY2020Q1 dates). The current FY keys
            # cannot represent both. Do not mix partial authority and fallback
            # names: retain the company's previous policy until period keys
            # can represent transition durations explicitly.
            return {}
        identities[identity]=end
    return result
