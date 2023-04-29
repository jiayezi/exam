import os
import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.styles import Font, Alignment
from tkinter import filedialog

table_range = ('A1:A2', 'B1:B2', 'C1:C2', 'D1:D2', 'E1:E2', 'F1:F2', 'G1:H1', 'I1:J1', 'K1:L1', 'M1:N1', 'O1:P1', 'Q1:Q2')

Calibri_10_font = Font(name='Calibri', size=10)
border1 = Border(left=Side(border_style='thin', color='000000'),
                 right=Side(border_style='thin', color='000000'),
                 top=Side(border_style='thin', color='000000'),
                 bottom=Side(border_style='thin', color='000000'))

path = filedialog.askopenfilename(title='请选择Excel文件', filetypes=[('Excel', '.xlsx')],
                                  defaultextension='.xlsx')

wb = openpyxl.load_workbook(path)
ws = wb.worksheets[0]

# 获取全部学校名字
schools = []
for row in range(3, ws.max_row + 1, 100):
    if ws.cell(row, 2).value not in schools:
        schools.append(ws.cell(row, 2).value)

titles = []
for i, row in enumerate(ws.values):
    if i == 2:
        break
    titles.append(row)

# 处理每个学校的数据
for school in schools:
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = '学生科目AB卷成绩'

    # 添加开头固定部分
    ws2.append(titles[0])
    ws2.append(titles[1])

    # 添加学校部分
    for row in ws.values:
        if row[1] == school:
            ws2.append(row)

    # 添加边框、对齐方式、字体
    for row in range(1, ws2.max_row + 1):
        for col in range(1, ws2.max_column + 1):
            ws2.cell(row, col).border = border1
            ws2.cell(row, col).font = Calibri_10_font
            ws2.cell(row, col).alignment = Alignment(horizontal='center', vertical='center')

    # 合并单元格
    for r in table_range:
        ws2.merge_cells(r)

    if not os.path.exists(f'E:/库/桌面/全部学校/{school}'):
        os.makedirs(f'E:/库/桌面/全部学校/{school}')

    wb2.save(f'E:/库/桌面/全部学校/{school}/{school}_学生科目AB分卷成绩.xlsx')
    wb2.close()

wb.close()
