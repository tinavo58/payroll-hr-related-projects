# tax.py
"""
y = ax - b

y: weekly withholding amount expressed in dollars
x: number of whole dollars in the weekly earnings plus 99 cents

This Tax Calculator does not take into account
53 weeks and 27 fortnights pay at this stage.
This occurs when a pay day falls on 31 Dec or
in a leap year, on 30 or 31 Dec.
"""
from coefficients import coefficients
from dataclasses import dataclass, field
import re

SCALE_4_RESIDENT = 0.47
SCALE_4_FOREGIN_RESIDENT = 0.45

@dataclass
class TaxCalc:
    earning: float
    isResident: bool = True # default resident for tax purpose
    isTfnProvided: bool = True # default TFN provided
    claimThreshold: bool = field(default=True) # default claim tax-free threshold

    scaleNum : int = field(init=False)
    a: float = field(init=False)
    b: float = field(init=False)
    roundingEarning: int = field(init=False)
    amountWithheld: int = field(init=False)

    def __post_init__(self):
        self.scaleNum = self._determineScale()
        if self.scaleNum == 4:
            self.amountWithheld = int(int(self.earning) * SCALE_4_RESIDENT) if self.isResident else int(int(self.earning) * SCALE_4_FOREGIN_RESIDENT)
        else:
            self.a, self.b = self._coefficients()
            self.roundingEarning = self.roundEarning()
            self.amountWithheld = self.calcTaxWithholding()

    def _determineScale(self):
        if not self.isTfnProvided:
            return 4
        if not self.isResident:
            return 3
        if not self.claimThreshold:
            return 1
        return 2

    def _coefficients(self):
        '''get coefficients for calculation of amounts to be withheld'''
        scale = coefficients.get(self.scaleNum)

        if scale:
            for e, ab in scale.items():
                if self.roundEarning() < e:
                    return ab
            else:
                return scale[max(scale.keys())] # select tax rate per max earning

    def roundEarning(self):
        return int(self.earning) + 0.99

    def calcTaxWithholding(self):
        return round(self.a * self.roundingEarning - self.b)


class TaxWeekly(TaxCalc):
    pass


class TaxFortnightly(TaxCalc):
    def roundEarning(self):
        return int(self.earning / 2) + 0.99

    def calcTaxWithholding(self):
        return round(self.a * self.roundingEarning - self.b) * 2


class TaxMonthly(TaxCalc):
    def checkCents(self):
        patternCompiled = re.compile("\d*.33\d*")
        if patternCompiled.match(str(self.earning)):
            return self.earning + 0.01
        else:
            return self.earning

    def roundEarning(self):
        return int(self.checkCents() * 3 / 13) + 0.99

    def calcTaxWithholding(self):
        return round(round(self.a * self.roundingEarning - self.b) * 13 / 3)


if __name__ == '__main__':
    tax_calc_dict = {
        'w': TaxWeekly,
        'f': TaxFortnightly,
        'm': TaxMonthly
    }
    print("Please enter your earning followed by 'w', 'f' or 'm':")
    print("\t- w: weekly\n\t- f: fortnightly\n\t- m: monthly\n")
    print("e.g. 548.45 w (meaning your weekly earning is $548.45)")

    while True:
        try:
            earningInput = input('> ')
            earning, paycycle = earningInput.split()
            earning = float(earning.strip())
            assert paycycle.strip().lower() in ('w', 'f', 'm')
        except Exception as ex:
            print(ex, "\n")
        else:
            tc = tax_calc_dict[paycycle](earning)
            break

    print('-' * 25)
    print(f"Your withholding amount is: ${tc.amountWithheld:,}")
