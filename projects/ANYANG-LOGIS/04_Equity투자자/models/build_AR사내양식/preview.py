import openpyxl, json, datetime
vals=json.load(open('values.json'))
w2=openpyxl.load_workbook('stage_a.xlsx'); s2=w2['A&R(사내양식)']
for k,v in vals.items():
    if s2[k].is_date and isinstance(v,(int,float)):
        v=datetime.datetime(1899,12,30)+datetime.timedelta(days=v)
    s2[k]=v
w2.save('preview.xlsx')
