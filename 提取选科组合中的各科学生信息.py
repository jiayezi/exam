from tkinter import filedialog
from openpyxl import Workbook, load_workbook


# 科目范围
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


if __name__ == '__main__':
    get_student_info()
