import openpyxl
from io import BytesIO
 
def unmerge_excel(file):
 
    wb = openpyxl.load_workbook(file)
 
    sheet = wb.active
 
    for merged_range in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged_range))
 
    temp = BytesIO()
 
    wb.save(temp)
 
    temp.seek(0)
 
    return temp