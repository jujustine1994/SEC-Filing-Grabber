from fiscal_audit import label_anomalies


def test_audit_reports_unclassified_duplicate_and_colliding_labels():
    issues = label_anomalies(['2010-05-22','FY2011Q2','FY2011Q2'],
                             ['2010-05-22','2010-08-14','2010-11-06'])
    assert issues['unclassified'] == ['2010-05-22']
    assert 'FY2011Q2' in issues['collisions']


def test_audit_does_not_call_an_incomplete_year_correct():
    issues = label_anomalies(['FY2020Q1','FY2020Q3'], ['2020-03-31','2020-09-30'])
    assert issues['incomplete_years'] == ['2020']
    assert not issues['misordered_years']
