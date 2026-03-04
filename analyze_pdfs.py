import pdfplumber
import re
import os

pdf_files = sorted([f for f in os.listdir('.') if f.endswith('.pdf')])

results = []

for pdf_file in pdf_files:
    result = {"file": pdf_file}
    
    with pdfplumber.open(pdf_file) as pdf:
        text = pdf.pages[0].extract_text()
        
        # BE Number
        be_match = re.search(r'INBOM4\s+(\d+)\s+(\d{2}/\d{2}/\d{4})', text)
        if be_match:
            result["be_no"] = be_match.group(1)
            result["be_date"] = be_match.group(2)
        
        # DEBT AMT
        debt_match = re.search(r'INBOM4\s+WH\s+(\d+)', text)
        if debt_match:
            result["debt_amt"] = debt_match.group(1)
        
        # Packages
        pkg_match = re.search(r'BE PKG\s+(\d+)', text)
        if pkg_match:
            result["packages"] = pkg_match.group(1)
        
        # All 7-digit numbers
        all_7digit = re.findall(r'\b(\d{7})\b', text)
        result["7digit_nums"] = all_7digit[:10]
        
        # Pattern after 8.G.CESS - this should be TOT.ASS VAL
        match4 = re.search(r'8\.G\.CESS.*?(\d{7})', text, re.DOTALL)
        if match4:
            result["tot_ass_val"] = match4.group(1)
    
    results.append(result)

# Print summary
print("="*100)
print("SUMMARY OF ALL PDFs")
print("="*100)
print(f"{'File':<55} {'BE No':<10} {'DEBT AMT':<12} {'TOT.ASS VAL':<12} {'PKG'}")
print("-"*100)

for r in results:
    print(f"{r.get('file',''):<55} {r.get('be_no',''):<10} {r.get('debt_amt',''):<12} {r.get('tot_ass_val',''):<12} {r.get('packages','')}")

print("\n" + "="*100)
print("7-DIGIT NUMBERS IN EACH PDF (first 8)")
print("="*100)
for r in results:
    print(f"{r.get('file','')}")
    print(f"  {r.get('7digit_nums', [])}")
