import os
import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.styles import Font, colors, Alignment
from tkinter import filedialog

table_range = ('A1:A2', 'B1:B2', 'C1:C2', 'D1:D2', 'E1:G1', 'H1:J1', 'K1:M1', 'N1:P1')

Calibri_10_font = Font(name='Calibri', size=10)
border1 = Border(left=Side(border_style='thin', color='000000'),
                 right=Side(border_style='thin', color='000000'),
                 top=Side(border_style='thin', color='000000'),
                 bottom=Side(border_style='thin', color='000000'))

path = filedialog.askopenfilename(title='请选择Excel文件', filetypes=[('Excel', '.xlsx')],
                                  defaultextension='.xlsx')

wb = openpyxl.load_workbook(path)
ws = wb.active

# 获取全部学校名字，填充学校名字
schools = []
for row in range(4, ws.max_row + 1):
    if ws.cell(row, 2).value == '' or ws.cell(row, 2).value is None:
        ws.cell(row, 2, ws.cell(row-1, 2).value)
    else:
        schools.append(ws.cell(row, 2).value)

for school in schools:
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = '班级科目均分'

    # 添加前3行
    for i, row in enumerate(ws.values):
        if i == 3:
            break
        ws2.append(row)

    # 添加学校数据
    for row in ws.values:
        if row[1] == school:
            ws2.append(row)
    # 下面是经典方法
    # for row in range(4, ws.max_row + 1):
    #     if ws.cell(row, 2).value == school:
    #         line = ['']
    #         for col in range(2, ws.max_column + 1):
    #             line.append(ws.cell(row, col).value)
    #         ws2.append(line)

    # 添加边框、对齐方式、字体
    for row in range(1, ws2.max_row + 1):
        for col in range(1, ws2.max_column + 1):
            ws2.cell(row, col).border = border1
            ws2.cell(row, col).font = Calibri_10_font
            ws2.cell(row, col).alignment = Alignment(horizontal='center', vertical='center')

    # 合并单元格
    for rg in table_range:
        ws2.merge_cells(rg)

    if not os.path.exists(f'F:/用户目录/桌面/全部学校/{school}'):
        os.makedirs(f'F:/用户目录/桌面/全部学校/{school}')

    wb2.save(f'F:/用户目录/桌面/全部学校/{school}/{school}_班级科目均分.xlsx')
    wb2.close()

wb.close()
