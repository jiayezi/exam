import os
import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.styles import Font, colors, Alignment
from tkinter import messagebox, simpledialog, filedialog

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
for row in range(2, ws.max_row + 1):
    if ws.cell(row, 2).value not in schools:
        schools.append(ws.cell(row, 2).value)

# 处理每个学校的数据
for school in schools:
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = '学生科目总成绩'

    # 添加开头固定部分
    line = []
    for col in range(1, ws.max_column+1):
        line.append(ws.cell(1, col).value)
    ws2.append(line)

    # 添加学校部分
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, 2).value == school:
            line = []
            for col in range(1, ws.max_column + 1):
                line.append(ws.cell(row, col).value)
            ws2.append(line)

    # 添加边框、对齐方式、字体
    for row in range(1, ws2.max_row + 1):
        for col in range(1, ws2.max_column + 1):
            ws2.cell(row, col).border = border1
            ws2.cell(row, col).alignment = Alignment(horizontal='center', vertical='center')
            ws2.cell(row, col).font = Calibri_10_font

    if not os.path.exists(f'F:/用户目录/桌面/全部学校/{school}'):
        os.makedirs(f'F:/用户目录/桌面/全部学校/{school}')

    wb2.save(f'F:/用户目录/桌面/全部学校/{school}/{school} 学生科目总分成绩 文科.xlsx')
    wb2.close()

wb.close()
