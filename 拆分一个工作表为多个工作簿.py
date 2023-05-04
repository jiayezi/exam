from tkinter import filedialog
import openpyxl

# 拆分的列，存储目录
school_index = 0
save_path = 'F:/用户目录/桌面/全部学校'

path = filedialog.askopenfilename(title='请选择Excel文件', initialdir='F:/用户目录/桌面/',
                                  filetypes=[('Excel', '.xlsx')], defaultextension='.xlsx')

wb = openpyxl.load_workbook(path, read_only=True)
ws = wb.active

# 获取全部学校名字
schools = set()
for i, row in enumerate(ws.values):
    if i == 0:
        continue
    schools.add(row[school_index])
school_list = list(schools)

# 提取每个学校的数据
for school in school_list:
    wb_new = openpyxl.Workbook(write_only=True)

    # 根据原始工作簿的工作表创建新工作簿的工作表，然后添加数据
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        header = next(ws.values)
        ws_new = wb_new.create_sheet(sheet)
        # 添加数据
        ws_new.append(header)
        for row in ws.values:
            if row[school_index] == school:
                ws_new.append(row)

    # 保存并关闭工作簿
    wb_new.save(f'{save_path}/{school}.xlsx')
    wb_new.close()
