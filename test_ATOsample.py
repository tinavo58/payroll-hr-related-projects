import pytest
from src.tax.tax import TaxWeekly, TaxFortnightly, TaxMonthly
from sample_data import *
# from test.ATO_tax_tables.read_pdf import w_data, f_data, m_data

### test weekly pay
@pytest.mark.parametrize("earnings, taxWithheld", wScale1_data)
def test_TaxCalcWeekly_scale1(earnings, taxWithheld):
    tc = TaxWeekly(earnings, claimThreshold=False)
    assert tc.calcTaxWithholding() ==  taxWithheld

@pytest.mark.parametrize("earnings, taxWithheld", wScale2_data)
def test_TaxCalcWeekly_scale2(earnings, taxWithheld):
    tc = TaxWeekly(earnings)
    assert tc.calcTaxWithholding() ==  taxWithheld

@pytest.mark.parametrize("earnings, taxWithheld", wScale3_data)
def test_TaxCalcWeekly_scale3(earnings, taxWithheld):
    tc = TaxWeekly(earnings, isResident=False)
    assert tc.amountWithheld ==  taxWithheld

### test fortnightly pay
@pytest.mark.parametrize("earnings, taxWithheld", fortnightly_scale2_data)
def test_TaxCalcFortnightly(earnings, taxWithheld):
    tcf = TaxFortnightly(earnings)
    assert tcf.calcTaxWithholding() ==  taxWithheld


### test monthly pay
@pytest.mark.parametrize("earnings, taxWithheld", monthly_scale2_data)
def test_TaxCalcMonthly(earnings, taxWithheld):
    tcm = TaxMonthly(earnings)
    assert tcm.amountWithheld ==  taxWithheld


### test weekly/fortnightly/monthly tax tables
# @pytest.mark.parametrize("earning, scale2, scale1", w_data)
# def test_weekly_tax_table(earning, scale2, scale1):
#     assert WeeklyTax(earning).amountWithheld == scale2
#     assert WeeklyTax(earning, claimThreshold=False).amountWithheld == scale1

# @pytest.mark.parametrize("earning, scale2, scale1", f_data)
# def test_fortnightly_tax_table(earning, scale2, scale1):
#     assert TaxFortnightly(earning).amountWithheld == scale2
#     assert TaxFortnightly(earning, claimThreshold=False).amountWithheld == scale1

# @pytest.mark.parametrize("earning, scale2, scale1", m_data)
# def test_monthly_tax_table(earning, scale2, scale1):
#     assert TaxMonthly(earning).amountWithheld == scale2
#     assert TaxMonthly(earning, claimThreshold=False).amountWithheld == scale1