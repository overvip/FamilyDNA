import openpyxl

tot_families = coll_families = tot_persons = coll_persons = 0
wb = openpyxl.load_workbook('C:\\Users\shine\\Desktop\\东孚派出所采血名单.xlsx')
sheetHK = wb.get_sheet_by_name('后柯村')
# sheetHK = wb.get_active_sheet()
# print(sheetHK['g15'].value==None)
# print(sheetHK['a3'].row)
prev_fam_complete = 0
for row in sheetHK.rows:
    if row[0].row > 2:
        if row[1].value is not None:
            tot_families += 1
            if prev_fam_complete == 1:
                coll_families += 1
            else:
                prev_fam_complete = 1
                # prev_fam_complete = 1
        if row[6].value is not None:
            coll_persons += 1
        else:
            prev_fam_complete = 0
        tot_persons += 1
if prev_fam_complete == 1:
    coll_families += 1
print(coll_families)