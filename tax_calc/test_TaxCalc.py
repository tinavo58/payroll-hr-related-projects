from src.tax.tax import TaxWeekly

def test_TaxCals():
    tc = TaxWeekly(353.33)

    assert tc.roundEarning() == 353.99
    assert tc.isResident == True
    assert tc.isTfnProvided == True
    assert tc.scaleNum == 2
    assert tc._coefficients() == (0, 0)
    assert tc.calcTaxWithholding() == 0