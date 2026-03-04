import sys
sys.path.insert(0, '.')
from generate_bond import BillOfEntryExtractor, BondCalculator
import os

print('Testing extraction for all PDFs...')
print('='*110)
print(f"{'File':<50} {'BE No':<10} {'DEBT AMT':<12} {'Bond Amt':<12} {'TOT.ASS VAL':<12} {'INV':<5} {'PKG'}")
print('-'*110)

for f in sorted(os.listdir('.')):
    if f.endswith('.pdf'):
        extractor = BillOfEntryExtractor(f)
        data = extractor.extract_all()
        calc = BondCalculator()
        bond_amt = calc.calculate_bond_amount(data['debt_amt'])
        print(f"{f:<50} {data['be_no']:<10} {data['debt_amt']:<12} {bond_amt:<12} {data['total_assessed_value']:<12} {data['invoice_count']:<5} {data['total_packages']}")
