from tkinter import filedialog
from openpyxl import load_workbook, Workbook

extra = 5  # 前5列数据用不上

# 读取Excel文件
file_path = filedialog.askopenfilename(title='请选择Excel文件', initialdir='F:/用户目录/桌面/',
                                       filetypes=[('Excel', '.xlsx')], defaultextension='.xlsx')
if not file_path:
    exit()
wb = load_workbook(file_path)
ws = wb.active
ws_rows = []
for row in ws.values:
    ws_rows.append(list(row))
student_data = ws_rows[1:]
ws_title = ws_rows[0]
subjects = ws_title[extra:]
student_num = len(student_data)

rateT = (1, 0.85, 0.5, 0.15, 0.02, 0)
rateY = (100, 86, 71, 56, 41, 30, 0)

for i, subject in enumerate(subjects):
    student_data.sort(key=lambda x: float(x[extra + i]), reverse=True)

    # 获取原始分临界值
    rateS = [int(student_data[0][extra + i])]
    temp_dj = 1
    for j, row in enumerate(student_data):
        rate = (student_num - j - 1) / student_num  # 领先率
        for index, value in enumerate(rateT):
            if rate > value:
                # dj = index
                if temp_dj != index:
                    temp_dj = index
                    rateS.append(int(row[extra + i]))
                break
    rateS.append(int(student_data[-1][extra + i]))
    print(f'{subject}原始分临界值：{rateS}')

    # 计算赋分成绩
    for row in student_data:
        score = int(row[extra + i])
        xsdj = 1
        for index, dj_score in enumerate(rateS):
            if index == 0:
                continue
            if score >= dj_score:
                xsdj = index
                break
        m = rateS[xsdj - 1]
        n = rateS[xsdj]
        a = rateY[xsdj - 1]
        b = rateY[xsdj]
        converts = (b * (score - m) + a * (n - score)) / (n - m)
        converts = round(converts)
        print(f'等级：{xsdj}\tm：{m}\tn：{n}\ta：{a}\tb：{b}\t原始分：{score:0>2d}\t转换分：{converts}')
        row.append(converts)
    ws_title.append(f'{subject}转换分')

# 写入Excel文件
wb = Workbook()
ws = wb.active
ws.append(ws_title)
for row in student_data:
    ws.append(row)
file_path = filedialog.asksaveasfilename(title='请选择文件存储路径', initialdir='F:/用户目录/桌面/', initialfile='赋分成绩',
                                         filetypes=[('Excel', '.xlsx')], defaultextension='.xlsx')
if file_path:
    wb.save(file_path)

"""
  * @param m 原始分开始
  * @param n 原始分结束
  * @param a 等级赋值分开始
  * @param b 等级赋值分结束
  * @param score 实考分数
"""
