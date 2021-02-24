import openpyxl, pathlib, pandas as pd
from pathlib import Path
from openpyxl import Workbook
from sklearn.feature_extraction.text import ENGLISH_STOP_WORDS

origDataPath = Path(r"C:\Documents\Projects\Data Analytics\Procurement_Report.xlsx")
origDatadf = pd.read_excel(origDataPath, sheet_name='Procurement_Report__All')

articleList = []
prepList = []
operWordset = []

possWords = origDatadf['Procurement Description'].loc[origDatadf['Type of Procurement'] == 'Design and Construction/Maintenance']
for x in possWords:
    operWordset.extend([y.lower() for y in filter(lambda f: (f.isalpha()) and (f.lower() not in ENGLISH_STOP_WORDS) and (len(f) >= 2), x.split(' '))])
        
wb = openpyxl.load_workbook(str(origDataPath))
ws = wb.create_sheet(title='ProjNameWords')

ws['A1'].value = 'ProjNameWords'
for n,x in enumerate(operWordset,2):
    ws[f'A{n}'].value = x

wb.save(r"C:\Documents\Projects\Data Analytics\Procurement_Report.xlsx")
print('Finished')
