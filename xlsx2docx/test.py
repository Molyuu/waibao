from docxtpl import DocxTemplate
from openpyxl import load_workbook,Workbook

filename = ''

doc = DocxTemplate('gork.docx')
wb = load_workbook('gork.xlsx', data_only = True)
sheet = wb.active
for row in sheet:
	rowMap = map(lambda x:x.value, row)
	if row[0].row ==  1:
		title = list(rowMap)
	else:
		context = dict(zip(title, rowMap))
		filename = context['filename']
		print('Rendering', filename, '...', end = '\r' )
		doc.render(context)
		doc.save(filename)
		print('Rendering', filename, 'done.')
