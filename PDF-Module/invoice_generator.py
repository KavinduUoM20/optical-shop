from docxtpl import DocxTemplate
import pandas as pd

from win32com import

df = pd.read_csv('Book1.csv')
for row_index,row in df.iterrows():
    cid = row['CID']
    cname = row['CNAME']
    print(cid,cname)

    tpl = DocxTemplate('Template.docx')
    df_to_dict = df.to_dict()
    x = df.to_dict(orient='records')
    context = x
    tpl.render(context[row_index])
    tpl.save('Doc\\'+cname+".docx")
