"""Pure period-label diagnostics. An incomplete fiscal year is inconclusive."""
from collections import defaultdict
import re


def label_anomalies(labels, ends):
    years, dates, by_label = defaultdict(list), defaultdict(list), defaultdict(list)
    unknown = []
    for label, end in zip(labels, ends):
        if not end:
            continue
        dates[end].append(label)
        by_label[label].append(end)
        match = re.fullmatch(r'FY(\d{4})Q([1-4])', str(label))
        if match:
            years[match[1]].append((end, int(match[2])))
        else:
            unknown.append(label)
    return dict(unclassified=unknown,
                duplicate_ends={d: ls for d, ls in dates.items() if len(ls)>1},
                collisions={l: ds for l, ds in by_label.items() if len(set(ds))>1},
                incomplete_years=[y for y, items in years.items() if len(items)<4],
                misordered_years={y: sorted(items) for y, items in years.items()
                    if len(items)>=4 and [q for _,q in sorted(items)] != [1,2,3,4]})
