from docxtpl import DocxTemplate
import pandas as pd

data = pd.read_excel('exam.xlsx')

for i, row in data.iterrows():
    cd = {}
    cd = {'name': str(row[1]) + " " + str(row[2]), 'ar':row[3], 'aw':row[4],'quran':row[5],'is':row[6], 'result':row[7], 'mark':row[8]}
    #print(cd)
    #print(row[1] + row[2] + " " + str(row[3]))
    doc = DocxTemplate('slip.docx')
    doc.render(cd)
    doc.save(str(row[1]) + str(row[2]) + '.docx')
