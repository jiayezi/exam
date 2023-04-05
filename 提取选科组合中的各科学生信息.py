from tkinter import filedialog
from openpyxl import Workbook, load_workbook


# 科目范围，是列数，不是索引
subject_range = (1, 3)


def get_student_info():
    # 提取
    file_path = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel', '.xlsx')],
                                           defaultextension='.xlsx')
    if not file_path:
        return
    subject_dict = {'历史': [], '物理': [], '政治': [], '地理': [], '化学': [], '生物': []}
    wb = load_workbook(file_path)
    ws = wb.active
    title = next(ws.values)
    for col in range(subject_range[0]-1, subject_range[1]):
        for i, row in enumerate(ws.values):
            if i == 0:
                continue
            subject_name = row[col]
            subject_dict[subject_name].append(row)
    wb.close()

    # 保存
    save_dir = filedialog.askdirectory(title='选择存储文件夹', initialdir='F:/用户目录/桌面/')
    if not save_dir:
        return
    for subject_name in subject_dict.keys():
        wb = Workbook()
        ws = wb.active
        ws.append(title)
        student_list = subject_dict[subject_name]
        for row in student_list:
            ws.append(row)
        wb.save(f'{save_dir}/{subject_name}.xlsx')
        wb.close()

    # 计算每个考场的人数
    room_index = 5
    wb = Workbook()
    ws = wb.active
    ws.append(('科目', '考室', '人数'))
    for subject_name in subject_dict.keys():
        student_list = subject_dict[subject_name]
        exam_room_dict = {}
        for row in student_list:
            room_name = row[room_index]
            if room_name in exam_room_dict.keys():
                exam_room_dict[room_name] += 1
            else:
                exam_room_dict[room_name] = 1
        counter = 0
        for k, v in exam_room_dict.items():
            ws.append((subject_name, f'第{k}考室', v))
            counter += v
        ws.append((subject_name, '加印', 10))
        ws.append((subject_name, '总计', counter+10))
        ws.append([])
    wb.save(f'{save_dir}/印刷安排.xlsx')


if __name__ == '__main__':
    get_student_info()
