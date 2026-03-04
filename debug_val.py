import pdfplumber
import re

files = [
    'Processed BE_IR_51939_25-26_20260205_1028.pdf', 
    'Processed BE_IR_52242_25-26_20260205_1028.pdf',
    'Processed BE_IR_52522_25-26_20260205_1027.pdf',
    'Processed BE_IR_52886_25-26_20260205_1027.pdf',
    'Processed BE_IR_53112_25-26_20260204_1652.pdf',
]

for f in files:
    with pdfplumber.open(f) as pdf:
        text = pdf.pages[0].extract_text()
        print(f"\n{'='*70}")
        print(f"File: {f}")
        print('='*70)
        
        # Find the duty summary header line
        # Pattern: 1.BCD 2.ACD ... 8.G.CESS 18.TOT.ASS VAL
        match = re.search(r'(1\.BCD.*?18\.TOT\.ASS VAL)\s*\n([^\n]+)', text, re.DOTALL)
        if match:
            header = match.group(1).replace('\n', ' ')
            values = match.group(2)
            print(f"Header: {header[:80]}...")
            print(f"Values: {values}")
            
            # Parse values
            nums = re.findall(r'[\d.]+', values)
            print(f"Parsed numbers: {nums}")
            if len(nums) >= 8:
                print(f"TOT.ASS VAL (8th value): {nums[7]}")
        else:
            print("Pattern not found - trying fallback")
            
            # Try finding just the line with values
            idx = text.find('18.TOT.ASS VAL')
            if idx > 0:
                context = text[idx:idx+200]
                print(f"Context: {repr(context)}")
