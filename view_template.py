from docx import Document
doc = Document('Triple_Duty_Bond_template__IR53112.docx')
print("Template Table Content:")
for table in doc.tables:
    for row_idx, row in enumerate(table.rows):
        row_data = []
        for cell_idx, cell in enumerate(row.cells):
            text = cell.text.replace('\n', '|').strip()
            row_data.append(f"[{cell_idx}]={text[:30]}")
        print(f"Row {row_idx}: {row_data}")
