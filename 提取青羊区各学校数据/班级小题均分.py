import os
import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.styles import Font, colors, Alignment
from tkinter import filedialog

# 设置每个工作表的合并范围
range1 = ('A1:A2', 'B1:B2', 'C1:C2', 'D1:D2', 'AI1:AI2', 'AJ1:AJ2')
range2 = ('A1:A2', 'B1:B2', 'C1:C2', 'D1:D2', 'W1:W2', 'X1:X2')
range3 = ('A1:A2', 'B1:B2', 'C1:C2', 'D1:D2', 'CH1:CH2', 'CI1:CI2')
range4 = ('A1:A2', 'B1:B2', 'C1:C2', 'D1:D2', 'AI1:AI2', 'AJ1:AJ2')
range5 = ('A1:A2', 'B1:B2', 'C1:C2', 'D1:D2', 'Y1:AA1', 'AG1:AG2', 'AH1:AH2')
# range6 = ('A1:A2', 'B1:B2', 'C1:C2', 'D1:D2', 'AW1:AW2', 'AX1:AX2')
table_ranges = (range1, range2, range3, range4, range5)

# 定义格式
Calibri_10_font = Font(name='Calibri', size=10)
border1 = Border(left=Side(border_style='thin', color='000000'),
                 right=Side(border_style='thin', color='000000'),
                 top=Side(border_style='thin', color='000000'),
                 bottom=Side(border_style='thin', color='000000'))

path = filedialog.askopenfilename(title='请选择Excel文件', filetypes=[('Excel', '.xlsx')],
                                  defaultextension='.xlsx')

wb = openpyxl.load_workbook(path)

# 获取全部学校名字，填充学校名字
schools = []
for index, sheet in enumerate(wb.sheetnames):
    ws = wb[sheet]
    for row in range(4, ws.max_row + 1):
        if ws.cell(row, 2).value == '' or ws.cell(row, 2).value is None:
            ws.cell(row, 2, ws.cell(row-1, 2).value)
        else:
            # 只获取第一个工作表里的全部学校名字
            if index == 0:
                schools.append(ws.cell(row, 2).value)

# 提取每个学校的数据
for school in schools:
    wb_new = openpyxl.Workbook()

    # 根据原始工作簿的工作表创建新工作簿的工作表，然后添加数据
    for index, sheet in enumerate(wb.sheetnames):
        ws = wb[sheet]
        if index == 0:
            ws_new = wb_new.active
            ws_new.title = sheet
        else:
            ws_new = wb_new.create_sheet(sheet)

        # 添加前3行
        for i, row in enumerate(ws.values):
            if i == 3:
                break
            ws_new.append(row)

        # 添加学校数据
        for row in ws.values:
            if row[1] == school:
                ws_new.append(row)

        # 设置边框、对齐方式、字体
        for row in range(1, ws_new.max_row + 1):
            for col in range(1, ws_new.max_column + 1):
                ws_new.cell(row, col).border = border1
                ws_new.cell(row, col).font = Calibri_10_font
                ws_new.cell(row, col).alignment = Alignment(horizontal='center', vertical='center')

        # 合并单元格
        for rg in table_ranges[index]:
            ws_new.merge_cells(rg)

    # 保存并关闭工作簿
    if not os.path.exists(f'F:/用户目录/桌面/全部学校/{school}'):
        os.makedirs(f'F:/用户目录/桌面/全部学校/{school}')
    wb_new.save(f'F:/用户目录/桌面/全部学校/{school}/{school}_班级小题均分.xlsx')
    wb_new.close()

wb.close()
