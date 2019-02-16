# -*- coding:utf-8 -*-
import openpyxl

villages = ('鼎美村',)  #, '后柯村', '芸美村')
wb = openpyxl.load_workbook('C:\\Users\shine\\OneDrive\\社区采血名单.xlsx')
for village in villages:
    tot_families = coll_families = tot_persons = coll_persons = 0
    sheetHK = wb[village]
    # sheetHK = wb.get_active_sheet()
    # print(sheetHK['g15'].value==None)
    # print(sheetHK['a3'].row)
    prev_fam_complete = 0
    family_name = None
    for row in sheetHK.rows:
        if row[0].row > 2:
            if row[1].value is not None:
                tot_families += 1
                if prev_fam_complete == 1:
                    coll_families += 1
                    print(family_name)
                else:
                    prev_fam_complete = 1
                    # prev_fam_complete = 1
            if row[6].value is not None:
                coll_persons += 1
            else:
                prev_fam_complete = 0
            tot_persons += 1
        if row[1].value is not None:
            family_name = row[1].value
    if prev_fam_complete == 1:
        coll_families += 1
        print(family_name)
    print(village + '家系数' + str(coll_families) + '个')