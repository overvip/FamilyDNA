# -*- coding:utf-8 -*-
import openpyxl
# from openpyxl.utils import column_index_from_string

def rows_merged(i):
    """判断单元格是否在合并的单元
    格内！！！

    s是字符串列名，
    需转换为列号"""
    # j = column_index_from_string(s)
    merged_list = wsVillage.merged_cells
    for cell in merged_list:
        if cell.max_col == cell.min_col == 2:
            if i >= cell.min_row and i <= cell.max_row:
                return cell.max_row - cell.min_row
    return 0


f='C:\\Users\\hx\\\onedrive\\社区采血名单.xlsx'
villages = ('鼎美村', '后柯村', '芸美村')
wb = openpyxl.load_workbook(f)
for village in villages:
    tot_families = coll_families = tot_persons = coll_persons = 0
    wsVillage = wb[village]

    # todo : 从cell[b3]开始判断海家系，若该行b列非合并单元，则为单人家系，
    #        否则为合并单元，该家系需采集多于1人，继续判断该合并单元占几行，
    #        即该家系要采集几名成员。再检查j列标记是否采集。

    merged_list = wsVillage.merged_cells
    row = 3
    while row < wsVillage.max_row:
        r = rows_merged(row)
        if r > 0:
            flag_completion = 1
            for j in range(r+1):
                if wsVillage.cell(row+j, 6) is not None:
                    coll_persons += 1
                else:
                    flag_completion = 0
                tot_persons += 1
            if flag_completion == 1:
                coll_families += 1
            row = row + r
        else:
            tot_families += 1
            tot_persons += 1
            if wsVillage.cell(row, 2).value is not None:
                coll_persons += 1
                coll_families += 1
            row += 1


    msgResult = village + '家系总数' + str(tot_families) + '个,已采集' + \
        str(coll_families) + '个，采集总数' + str(tot_persons) + \
        '个，已采集' + str(coll_persons) + '个'
    print(msgResult)
    input('按回车显示下一个村')
    wsVillage['a1'].value = msgResult
    wb.save(f)

