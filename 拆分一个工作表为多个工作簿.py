from tkinter import filedialog
import openpyxl

# 指定拆分的列和存储目录
name_index = 1
save_path = 'F:/用户目录/桌面/全部学校'

path = filedialog.askopenfilename(title='请选择Excel文件', initialdir='F:/用户目录/桌面/',
                                  filetypes=[('Excel', '.xlsx')], defaultextension='.xlsx')

wb = openpyxl.load_workbook(path)
ws = wb.active

# 获取全部名字
names = set()
for i, row in enumerate(ws.values):
    if i == 0:
        continue
    names.add(row[name_index])

# 提取数据
for name in names:
    wb_new = openpyxl.Workbook()

    # 根据原始工作簿的工作表创建新工作簿的工作表，然后添加数据
    for ws in wb:
        ws_new = wb_new.create_sheet(ws.title)
        # ws_new = wb_new.active
        for i, row in enumerate(ws.values):
            # 添加标题
            if i < 2:
                ws_new.append(row)
            # 添加数据
            if row[name_index] == name:
                ws_new.append(row)
        # 合并单元格
        for rg in ws.merged_cells:
            ws_new.merge_cells(str(rg))

    # 保存并关闭工作簿
    wb_new.save(f'{save_path}/{name}.xlsx')
    wb_new.close()

wb.close()
