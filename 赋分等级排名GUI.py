import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from threading import Thread
import openpyxl


class Subject:
    """记录转换分、等级、原始分市排名和校排名、转换分市排名和校排名"""

    def __int__(self):
        self.score = 0
        self.convert = 0
        self.grade = 0
        self.city_rank = 0
        self.school_rank = 0
        self.convert_city_rank = 0
        self.convert_school_rank = 0


class Student:
    def __int__(self):
        self.id = ''
        self.class_ = ''
        self.school = ''
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
        super().__init__(master, padding=20)
        self.subject_Checkbutton = []
        self.subject_name = []
        self.selected_subject_name = []
        self.item_var = []
        self.wb = None
        self.ws = None
        self.student_data = []

        self.createUI()

    def createUI(self):
        """创建界面元素"""
        self.grid(row=0, column=0)
        self.label1 = ttk.Label(master=self, text='数据项', font=('黑体', 12))
        self.label1.grid(row=0, column=0, pady=10)
        self.label2 = ttk.Label(master=self, text='功能', font=('黑体', 12))
        self.label2.grid(row=0, column=1, pady=10)

        self.item_frame = ttk.Frame(master=self, padding=(10, 0, 10, 0))
        self.item_frame.grid(row=1, column=0, pady=10)
        cb = ttk.Checkbutton(master=self.item_frame, text='全选', command=self.select_all)
        cb.grid(row=0, column=0, pady=3, sticky='w')

        self.btn_frame = ttk.Frame(master=self, padding=(10, 0, 10, 0))
        self.btn_frame.grid(row=1, column=1, pady=10, sticky='n')
        self.open_btn = ttk.Button(master=self.btn_frame, text='选择文档', command=self.open_file)
        self.open_btn.grid(row=0, column=0, pady=20)
        self.convert_btn = ttk.Button(master=self.btn_frame, text='成绩赋分',
                                      command=lambda: MyThread(self.convert_score), state=DISABLED)
        self.convert_btn.grid(row=1, column=0, pady=10)
        self.rank_btn = ttk.Button(master=self.btn_frame, text='计算排名', command=lambda: MyThread(self.convert_score),
                                   state=DISABLED)
        self.rank_btn.grid(row=2, column=0, pady=10)

    def open_file(self):
        """打开Excel，创建复选框"""
        path = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel', '.xlsx')],
                                          defaultextension='.xlsx')
        if not path:
            return
        self.wb = openpyxl.load_workbook(path, read_only=True)
        self.ws = self.wb.active

        title = []
        data1 = []
        for i, row in enumerate(self.ws.values):
            if i == 0:
                title = row
            if i == 1:
                data1 = row
                break

        # 添加复选框之前先删除上次创建的复选框
        for cb in self.subject_Checkbutton:
            cb.destroy()

        self.item_var = []
        for i, item in enumerate(title):
            cell_data = data1[i]
            if cell_data is None or cell_data == '':
                continue
            if isinstance(cell_data, float) or isinstance(cell_data, int):
                self.item_var.append(ttk.StringVar())
                cb = ttk.Checkbutton(master=self.item_frame, text=f'{i + 1:0>2d} {item}', variable=self.item_var[-1],
                                     onvalue='1', offvalue='0')
                cb.grid(row=i + 1, column=0, pady=3, sticky='w')
                self.subject_Checkbutton.append(cb)
                self.subject_name.append(item)

        # 全选
        # for i, item in enumerate(title):
        #     self.item_var[i].set(item)

        # 存储读取的数据
        self.student_data = []
        for row in self.ws.values:
            self.student_data.append(row)
        print(len(self.student_data))

        self.convert_btn.config(state=NORMAL)
        self.rank_btn.config(state=NORMAL)

    def select_all(self):
        for item in self.item_var:
            item.set('1')

    def convert_score(self):
        """计算赋分成绩和等级"""
        self.open_btn.config(state=DISABLED)
        self.convert_btn.config(state=DISABLED)
        self.rank_btn.config(state=DISABLED)

        for i, value in enumerate(self.item_var):
            if value.get() == '1':
                self.selected_subject_name.append(self.subject_name[i])
        print(self.selected_subject_name)

        messagebox.showinfo(message='计算完成')
        self.open_btn.config(state=NORMAL)
        self.convert_btn.config(state=NORMAL)
        self.rank_btn.config(state=NORMAL)


if __name__ == "__main__":
    app = ttk.Window(title="成绩计算程序")
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()
    offset_x = int((screen_width - 650) / 2)
    offset_y = int((screen_height - 380) / 2)
    # app.geometry(f'650x380+{offset_x}+{offset_y}')  # 窗口大小
    app.minsize(220, 330)
    app.iconbitmap('green_apple.ico')
    App(app)
    app.mainloop()
