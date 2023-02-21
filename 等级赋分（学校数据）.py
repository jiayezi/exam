from tkinter import filedialog
from openpyxl import load_workbook, Workbook

extra = 5  # 前9列数据用不上

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

rateT = (1, 0.85, 0.5, 0.15, 0.02, 0)
rateY = ((100, 86), (85, 71), (70, 56), (55, 41), (40, 30))
dict_dj = {0: 'A', 1: 'B', 2: 'C', 3: 'D', 4: 'E'}


def getp1(data, index, score):  # 小于某值的个数
    d1 = filter(lambda x: score > x[index] > 0, data)
    return len(list(d1))


def getp2(data, index, score):  # 大于某值的个数
    d1 = filter(lambda x: x[index] > score, data)
    return len(list(d1))


for i, subject in enumerate(subjects):

    #########获取科目非0最小原始分#########
    student_data.sort(key=lambda x: float(x[extra + i]), reverse=False)
    minscore = 0
    for j in range(len(student_data)):
        if student_data[j][extra + i] > minscore:
            minscore = student_data[j][extra + i]
            break
    student_data.sort(key=lambda x: float(x[extra + i]), reverse=True)

    # 获取原始分等级区间
    rateS = [[float(student_data[0][extra + i])]]
    temp_dj = 0
    A = 0
    T = getp2(student_data, extra + i, 0)
    rate = 1
    for j, row in enumerate(student_data):
        current_score = float(row[extra + i])
        A = getp1(student_data, extra + i, current_score)
        rate = A / T  # 领先率
        ########原始分为0分不参与原始分对照表#######
        if current_score == 0:
            continue

        for index, value in enumerate(rateT):
            if index == 0:
                continue
            if rate >= value:
                if temp_dj != index - 1:
                    temp_dj = index - 1
                    rateS[temp_dj - 1].append(float(student_data[j - 1][extra + i]))
                    rateS.append([float(row[extra + i])])
                break
        # print(temp_dj)

    rateS[-1].append(minscore)  # 转换分对照的原始分区间最后一个值为最后一个不为0分的最小值
    print(f'\n{subject}原始分等级区间：{rateS}')

    # 计算赋分成绩
    for row in student_data:
        score = float(row[extra + i])
        xsdj = -1
        converts = 0  # 转换分初始为0
        for index, dj_score in enumerate(rateS):
            if dj_score[0] >= score >= dj_score[1]:
                xsdj = index
                break
        m = 0
        n = 0
        a = 0
        b = 0
        if score > 0:  # 只有原始分不为0时才计算转换分，否则转换分为初始值0
            m = rateS[xsdj][1]
            n = rateS[xsdj][0]
            a = rateY[xsdj][1]
            b = rateY[xsdj][0]
            converts = (b * (score - m) + a * (n - score)) / (n - m)
            converts = round(converts)
        # print(f'等级：{xsdj}\tm：{m}\tn：{n}\ta：{a}\tb：{b}\t原始分：{score:.1f}\t转换分：{converts}')
        row.append(converts)
        if score > 0:
            row.append(dict_dj[xsdj])
        else:
            row.append('-')
    ws_title.append(f'{subject}转换分')
    ws_title.append(f'{subject}等级')

# 写入Excel文件
wb = Workbook()
ws = wb.active
ws.append(ws_title)
for row in student_data:
    ws.append(row)
file_path = filedialog.asksaveasfilename(title='请选择文件存储路径', initialdir='F:/用户目录/桌面/',
                                         initialfile='赋分成绩',
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
