import openpyxl
import docxtpl

CONGRAT_TEMPLATE = ('Награждается {{ surname }} {{ name }}, обучающийся(яся) {{ school }} '
                    'за достижения.')

wb = openpyxl.load_workbook("Люди.xlsx")
ws = wb.active
vals = list(ws.values)

rec_tpl = {head: None for head in vals[0]}
records = [rec_tpl.copy() for row in vals]
records = records[:-1]

for i, row in enumerate(vals[1:]):
    for j, val in enumerate(row):
        records[i][vals[0][j]] = val
doc = docxtpl.DocxTemplate('word_template.docx')
ctx = {'records': records}
doc.render(ctx)
doc.save('resultat.docx')