import openpyxl
from tkinter import filedialog


def get_closest(num, collection):
    """查找最接近的值"""
    return min(collection, key=lambda x: abs(x-num))


start_col, end_col = (9, 16)  # 修改1
path = filedialog.askopenfilename(title='选择Excel工作簿',
                                  filetypes=[('Excel工作簿', '.xlsx')],
                                  defaultextension='.xlsx')
wb = openpyxl.load_workbook(path, read_only=True)

school_sheet = wb['学校考试成绩']
city_sheet1 = wb['市级考试成绩1']
city_sheet2 = wb['市级考试成绩2']

# 读取市级考试的每个科目的领先百分比、原始分、转换分、等级
dict_list1, dict_list2 = [], []
for col_index in range(start_col, end_col, 4):  # 每个科目信息占4列
    temp_dict = {}
    for row_index, row in enumerate(city_sheet1.values):
        if row_index == 0:
            continue
        temp_dict[row[col_index]] = row[col_index+1:col_index+4]
    dict_list1.append(temp_dict)
    temp_dict = {}
    for row_index, row in enumerate(city_sheet2.values):
        if row_index == 0:
            continue
        temp_dict[row[col_index]] = row[col_index+1:col_index+4]
    dict_list2.append(temp_dict)

# 计算学校的赋分成绩
start_col, end_col = (11, 13)  # 修改2
all_data = []
for row in school_sheet.values:
    all_data.append(list(row))
student_data = all_data[1:]
ws_title = all_data[0]
subjects = ws_title[start_col:end_col]

for sub_index, subject in enumerate(subjects):
    city_dict1 = dict_list1[sub_index]
    city_dict2 = dict_list2[sub_index]
    # 计算赋分成绩
    for row in student_data:
        score = row[start_col + sub_index]
        percent_rank = row[start_col + sub_index-2]  # 修改3
        if not isinstance(score, (int, float)) or score == 0:
            row.append('')
            row.append('')
            continue
        city_percent_rank1 = get_closest(percent_rank, city_dict1.keys())
        city_percent_rank2 = get_closest(percent_rank, city_dict2.keys())
        convert1 = city_dict1[city_percent_rank1][1]
        convert2 = city_dict2[city_percent_rank2][1]
        convert = round(convert1*0.6 + convert2*0.4)
        grade = city_dict1[city_percent_rank1][2]
        row.append(convert)
        row.append(grade)
    ws_title.append(f'{subject}转换分')
    ws_title.append(f'{subject}等级')
wb.close()

# 写入Excel文件
wb = openpyxl.Workbook()
ws = wb.active
ws.append(ws_title)
for row in student_data:
    ws.append(row)
file_path = filedialog.asksaveasfilename(title='请选择文件存储路径', initialdir='F:/用户目录/桌面/',
                                         initialfile='赋分成绩',
                                         filetypes=[('Excel', '.xlsx')], defaultextension='.xlsx')
if file_path:
    wb.save(file_path)
    wb.close()
