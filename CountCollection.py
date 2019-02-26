# -*- coding:utf-8 -*-
import openpyxl
from openpyxl.utils import column_index_from_string

def is_cell_merged(i,s):
    """判断单元格是否在合并的单元
    格内！！！


    s是字符串列名，
    需转换为列号"""
    j = column_index_from_string(s)
    merged_list = wsVillage.merged_cells
    for cell in merged_list:
        if i != cell.min_row and j != cell.min_col:
                if i > cell.min_row and i <= cell.max_row and \
                j > cell.min_col and j <= cell.max_col:
                    return True
    return False


f='C:\\Users\\shine\\\onedrive\\社区采血名单.xlsx'
villages = ('鼎美村', '后柯村', '芸美村')
wb = openpyxl.load_workbook(f)
for village in villages:
    tot_families = coll_families = tot_persons = coll_persons = 0
    wsVillage = wb[village]
    # print(wsVillage.merged_cells)
    # wsVillage = wb.get_active_sheet()
    # print(wsVillage['g15'].value==None)
    # print(wsVillage['a3'].row)
    prev_fam_complete = 0
    family_name = None
    for row in wsVillage.rows:
        if row[0].row > 2:
            # if is_cell_merged(row[1].row,row[1].column) == False:
            if str(row[1].value).find('家系'):
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
        # if is_cell_merged(row[1].row,row[1].column) == False:
        if str(row[1].value).find('家系'):
            family_name = row[1].value
    if prev_fam_complete == 1:
        coll_families += 1
        print(family_name)
    msgResult = village + '家系总数' + str(tot_families) + '个,已采集' + \
        str(coll_families) + '个，采集总数' + str(tot_persons) + \
        '个，已采集' + str(coll_persons) + '个'
    print(msgResult)
    input('按回车显示下一个村')
    wsVillage['a1'].value = msgResult  # village + '家系数' + str(coll_families) + '个'
    wb.save(f)

