import json
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from threading import Thread
import openpyxl


class Subject:
    """记录转换分、等级、原始分市排名和校排名、转换分市排名和校排名"""

    def __init__(self, name, score):
        self.name = name
        self.score = score
        self.convert = 0
        self.grade = ''

        self.city_rank = 0
        self.area_rank = 0
        self.school_rank = 0
        self.class_rank = 0
        self.convert_city_rank = 0
        self.convert_area_rank = 0
        self.convert_school_rank = 0
        self.convert_class_rank = 0


class Student:
    def __init__(self, area, school, class_, exam_id):
        self.area = area
        self.school = school
        self.class_ = class_
        self.exam_id = exam_id
        self.subject = []


class MyThread(Thread):
    def __init__(self, func, *args):
        super().__init__()
        self.func = func
        self.args = args
        self.setDaemon(True)
        self.start()  # 在这里开始

    def run(self):
        self.func(*self.args)


class App(ttk.Frame):

    def __init__(self, master):
        super().__init__(master, padding=(20, 10, 20, 10))
        self.subject_checkbutton = []
        self.checkbutton_var = []
        self.checkbutton_name = []
        self.selected_subject_index = []
        self.selected_subject_name = []
        self.title = []
        self.student_objs = []
        self.wb = None
        self.ws = None

        self.createUI()

    def createUI(self):
        """创建界面元素"""
        self.grid(row=0, column=0)

        self.cbox_frame = ttk.Frame(master=self, padding=(10, 0, 10, 0))
        self.cbox_frame.grid(row=0, column=0, pady=10, sticky='n')
        self.item_frame = ttk.Frame(master=self, padding=(10, 0, 10, 0))
        self.item_frame.grid(row=0, column=1, pady=10, sticky='n')
        self.item_frame.grid_columnconfigure(0, minsize=65)
        self.btn_frame = ttk.Frame(master=self, padding=(10, 0, 10, 0))
        self.btn_frame.grid(row=0, column=2, pady=10, sticky='n')
        ttk.Label(master=self.item_frame, text='数据项', font=('黑体', 12)).grid(row=0, column=0, pady=10)
        self.select_all_var = ttk.StringVar()
        cb = ttk.Checkbutton(master=self.item_frame, text='全选', variable=self.select_all_var,
                             onvalue='全选', offvalue='', command=self.select_all)
        cb.grid(row=1, column=0, pady=16, sticky='w')

        ttk.Label(master=self.cbox_frame, text='系统字段', font=('黑体', 12)).grid(row=0, column=0, pady=10)
        ttk.Label(master=self.cbox_frame, text='Excel字段', font=('黑体', 12)).grid(row=0, column=1, pady=10)
        ttk.Label(master=self.cbox_frame, text='区市', font=('黑体', 10)).grid(row=1, column=0, pady=10)
        ttk.Label(master=self.cbox_frame, text='学校', font=('黑体', 10)).grid(row=2, column=0, pady=10)
        ttk.Label(master=self.cbox_frame, text='班级', font=('黑体', 10)).grid(row=3, column=0, pady=10)
        ttk.Label(master=self.cbox_frame, text='考号', font=('黑体', 10)).grid(row=4, column=0, pady=10)
        self.area_cbox = ttk.Combobox(master=self.cbox_frame, width=10)
        self.area_cbox.grid(row=1, column=1, pady=10)
        self.school_cbox = ttk.Combobox(master=self.cbox_frame, width=10)
        self.school_cbox.grid(row=2, column=1, pady=10)
        self.class_cbox = ttk.Combobox(master=self.cbox_frame, width=10)
        self.class_cbox.grid(row=3, column=1, pady=10)
        self.id_cbox = ttk.Combobox(master=self.cbox_frame, width=10)
        self.id_cbox.grid(row=4, column=1, pady=10)

        ttk.Label(master=self.btn_frame, text='功能', font=('黑体', 12)).grid(row=0, column=0, pady=10)
        self.open_btn = ttk.Button(master=self.btn_frame, text='选择文档', command=self.open_file)
        self.open_btn.grid(row=1, column=0, pady=10)
        self.submit_btn = ttk.Button(master=self.btn_frame, text='提交数据',
                                     command=lambda: MyThread(self.extract_data), state=DISABLED)
        self.submit_btn.grid(row=2, column=0, pady=10)
        self.convert_btn = ttk.Button(master=self.btn_frame, text='成绩赋分',
                                      command=lambda: MyThread(self.convert_score),
                                      state=DISABLED)
        self.convert_btn.grid(row=3, column=0, pady=10)
        self.rank_btn = ttk.Button(master=self.btn_frame, text='计算排名', command=lambda: MyThread(self.rank),
                                   state=DISABLED)
        self.rank_btn.grid(row=4, column=0, pady=10)
        self.save_btn = ttk.Button(master=self.btn_frame, text='保存文件', command=lambda: MyThread(self.save_file),
                                   state=DISABLED)
        self.save_btn.grid(row=5, column=0, pady=10)

    def open_file(self):
        """打开Excel，创建复选框"""
        path = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel', '.xlsx')],
                                          defaultextension='.xlsx')
        if not path:
            return
        self.wb = openpyxl.load_workbook(path, read_only=True)
        self.ws = self.wb.active
        self.title = next(self.ws.values)

        # 添加复选框之前先删除上次创建的复选框
        for cb in self.subject_checkbutton:
            cb.destroy()

        possible_subjects = (
            '语文', '数学', '数学文', '数学理', '英语', '外语', '政治', '历史', '地理', '物理', '化学', '生物')
        self.checkbutton_var = []
        self.checkbutton_name = []
        for i, item in enumerate(self.title):
            if item in possible_subjects:
                self.checkbutton_var.append(ttk.StringVar())
                cb = ttk.Checkbutton(master=self.item_frame, text=f'{i + 1:0>2d} {item}',
                                     variable=self.checkbutton_var[-1],
                                     onvalue=item, offvalue='')
                cb.grid(row=i + 2, column=0, pady=3, sticky='w')
                self.checkbutton_name.append(item)

        # 配置下拉列表
        self.area_cbox.config(values=self.title)
        if '区市' in self.title:
            self.area_cbox.current(self.title.index('区市'))
        self.school_cbox.config(values=self.title)
        if '学校' in self.title:
            self.school_cbox.current(self.title.index('学校'))
        self.class_cbox.config(values=self.title)
        if '班级' in self.title:
            self.class_cbox.current(self.title.index('班级'))
        self.id_cbox.config(values=self.title)
        if '考号' in self.title:
            self.id_cbox.current(self.title.index('考号'))

        self.submit_btn.config(state=NORMAL)

    def select_all(self):
        if self.select_all_var.get():
            for i, item in enumerate(self.checkbutton_var):
                value = self.checkbutton_name[i]
                item.set(value)
        else:
            for item in self.checkbutton_var:
                item.set('')

    def extract_data(self):
        """提取数据"""
        # 获取选中的科目的下标和名字
        self.selected_subject_index = []
        self.selected_subject_name = []
        for i, value in enumerate(self.checkbutton_var):
            data = value.get()
            if data:
                index = self.title.index(data)
                self.selected_subject_index.append(int(index))
                self.selected_subject_name.append(data)

        # 获取下拉列表的值
        area = self.area_cbox.get()
        school = self.school_cbox.get()
        class_ = self.class_cbox.get()
        id = self.id_cbox.get()
        if not (area and school and class_ and id and self.selected_subject_name):
            messagebox.showwarning(message='请选择对应字段并勾选数据项！')
            return

        self.open_btn.config(state=DISABLED)
        self.submit_btn.config(state=DISABLED)

        # 获取下拉列表的值的下标
        area_index = self.title.index(area)
        school_index = self.title.index(school)
        class_index = self.title.index(class_)
        id_index = self.title.index(id)

        print('区、校、班、考号：', area_index, school_index, class_index, id_index)
        print('选择科目：', self.selected_subject_index)

        # 存为对象
        self.student_objs = []
        for i, row in enumerate(self.ws.values):
            if i == 0:
                continue
            stu_obj = Student(row[area_index], row[school_index], row[class_index], row[id_index])
            # 添加科目分数
            for index, cell_index in enumerate(self.selected_subject_index):
                subject_obj = Subject(self.selected_subject_name[index], row[cell_index])
                stu_obj.subject.append(subject_obj)
            self.student_objs.append(stu_obj)

        # 存为json文件
        # with open('DataKv.json', 'w') as file_obj:
        #     json.dump(self.student_objs, file_obj, ensure_ascii=False, indent=4, default=lambda t: t.__dict__)

        messagebox.showinfo(message='成功')
        self.open_btn.config(state=NORMAL)
        self.submit_btn.config(state=NORMAL)
        self.convert_btn.config(state=NORMAL)
        self.rank_btn.config(state=NORMAL)
        self.save_btn.config(state=NORMAL)

    def convert_score(self):
        """成绩赋分"""

        def sort_rule(score):
            """定义排序规则"""
            if score is None or score == '':
                return 0
            else:
                return float(score)

        # 配置领先率、赋分区间和等级
        rateT = (1, 0.97, 0.9, 0.74, 0.50, 0.26, 0.1, 0.03, 0)
        rateY = ((100, 91), (90, 81), (80, 71), (70, 61), (60, 51), (50, 41), (40, 31), (30, 21))
        dict_dj = {0: 'A', 1: 'B+', 2: 'B', 3: 'C+', 4: 'C', 5: 'D+', 6: 'D', 7: 'E'}

        for sub_index, subject in enumerate(self.selected_subject_name):
            self.student_objs.sort(key=lambda x: sort_rule(x.subject[sub_index].score), reverse=True)

            # 获取得分大于0分的人数、获取大于0分的最小原始分
            student_data_reverse = self.student_objs[::-1]
            student_num = len(student_data_reverse)
            min_score = 0.0
            for w_index, row in enumerate(student_data_reverse):
                if row.subject[sub_index].score is None or row.subject[sub_index].score == '':
                    continue
                if float(row.subject[sub_index].score) > 0.0:
                    student_num -= w_index
                    min_score = float(row.subject[sub_index].score)
                    break

            # 获取原始分等级区间
            rateS = [[float(self.student_objs[0].subject[sub_index].score)]]
            temp_dj = 0
            rate = (student_num - 1) / student_num
            for row_index, row in enumerate(self.student_objs):
                current_score_str = row.subject[sub_index].score
                if (current_score_str is None or current_score_str == '') and row_index != 0:
                    continue
                current_score = float(current_score_str)
                # 原始分为0分不参与原始分对照表
                if current_score < 0.001:
                    continue

                previous_score_str = self.student_objs[row_index - 1].subject[sub_index].score
                if row_index == 0:
                    previous_score_str = 0

                previous_score = float(previous_score_str)
                if current_score != previous_score:
                    rate = (student_num - row_index - 1) / student_num  # 领先率
                    for v_index, value in enumerate(rateT):
                        if v_index == 0:
                            continue
                        if rate >= value:
                            if temp_dj != v_index - 1:
                                temp_dj = v_index - 1
                                rateS[temp_dj - 1].append(
                                    float(self.student_objs[row_index - 1].subject[sub_index].score))
                                rateS.append([float(row.subject[sub_index].score)])
                            break
            # rateS[-1].append(float(student_data[-1][extra + i]))
            rateS[-1].append(min_score)
            print(f'\n{subject}原始分等级区间：{rateS}')

            # 计算赋分成绩
            for row in self.student_objs:
                score_str = row.subject[sub_index].score
                if score_str is None or score_str == '' or float(score_str) < 0.001:
                    continue
                score = float(score_str)
                xsdj = 0
                for index, dj_score in enumerate(rateS):
                    if index == 0:
                        continue
                    if dj_score[0] >= score >= dj_score[1]:
                        xsdj = index
                        break
                m = rateS[xsdj][1]
                n = rateS[xsdj][0]
                a = rateY[xsdj][1]
                b = rateY[xsdj][0]
                converts = (b * (score - m) + a * (n - score)) / (n - m)
                row.subject[sub_index].convert = round(converts)
                row.subject[sub_index].grade = dict_dj[xsdj]

        messagebox.showinfo(message='赋分完成')

    def rank(self):
        pass

        messagebox.showinfo(message='排名计算完成')

    def save_file(self):
        wb = openpyxl.Workbook(write_only=True)
        ws = wb.create_sheet()

        ws_title = ['区市', '学校', '班级', '考号']
        for subject in self.selected_subject_name:
            ws_title.extend([subject, f'{subject}转换分', f'{subject}等级'])
        ws.append(ws_title)

        for row in self.student_objs:
            ws_row = [row.area, row.school, row.class_, row.exam_id]
            for subject in row.subject:
                ws_row.extend([subject.score, subject.convert, subject.grade])
            ws.append(ws_row)

        path = filedialog.asksaveasfilename(title='请选择文件存储位置',
                                            initialdir='F:/用户目录/桌面/',
                                            initialfile='计算结果',
                                            filetypes=[('Excel', '.xlsx')],
                                            defaultextension='.xlsx')
        if path:
            wb.save(path)
        messagebox.showinfo(message='文件保存成功')


if __name__ == "__main__":
    app = ttk.Window(title="成绩计算程序")
    # app.geometry(f'650x380')  # 窗口大小
    app.minsize(220, 330)
    app.iconbitmap('green_apple.ico')
    App(app)
    app.place_window_center()
    app.mainloop()
