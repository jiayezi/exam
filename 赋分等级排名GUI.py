import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import threading
import openpyxl
import time


class MyThread(threading.Thread):
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
        self.pack(fill=BOTH, expand=YES)
        self.btn1 = ttk.Button(master=self, text='选择文档', command=self.open_file)
        self.btn1.pack(pady=10)
        self.label1 = ttk.Label(master=self, text='计算项目：', font=('黑体', 12))
        self.label2 = ttk.Label(master=self, text='计算项目：', font=('黑体', 12))
        self.sj_frame = ttk.Frame(master=self)
        self.xm_frame = ttk.Frame(master=self)
        self.submit_btn = ttk.Button(master=self, text='提交', command=lambda: MyThread(self.deal_data))

        self.all_subjects = ('语文', '数学', '数学文', '数学理', '英语', '政治', '历史', '地理', '物理', '化学', '生物')
        self.ws_subjects = []
        self.selected_subjects = []
        self.var = []
        self.var2 = []

    def open_file(self):
        path = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel', '.xlsx')],
                                          defaultextension='.xlsx')
        if not path:
            return

        self.btn1.config(state=DISABLED)

        wb = openpyxl.load_workbook(path, read_only=True)
        ws = wb.active
        title = next(ws.values)
        self.ws_subjects = []
        for item in title:
            if item in self.all_subjects:
                self.ws_subjects.append(item)

        self.label1.pack(pady=10)
        self.sj_frame.pack(padx=0, pady=10)
        # 添加复选框
        self.var = []
        for subject in self.ws_subjects:
            self.var.append(ttk.StringVar())
            cb = ttk.Checkbutton(self.sj_frame, text=subject, variable=self.var[-1], onvalue=subject, offvalue='')
            cb.pack(side=LEFT, padx=5)
        # 默认全选
        for i, subject in enumerate(self.ws_subjects):
            self.var[i].set(subject)

        self.label2.pack(pady=10)

        self.xm_frame.pack(padx=0, pady=10)
        item_list = ('赋分成绩', '赋分等级', '市排名', '校排名')
        self.var2 = []
        for item in item_list:
            self.var2.append(ttk.StringVar())
            cb = ttk.Checkbutton(self.xm_frame, text=item, variable=self.var2[-1], onvalue=item, offvalue='')
            cb.pack(side=LEFT, padx=5)
        # 默认全选
        for i, item in enumerate(item_list):
            self.var2[i].set(item)

        self.submit_btn.pack(pady=20)

    def deal_data(self):
        self.submit_btn.config(state=DISABLED)

        print('计算科目：', end=' ')
        for v in self.var:
            if v.get():
                print(v.get(), end=' ')
        print()

        print('计算项目：', end=' ')
        for v in self.var2:
            if v.get():
                print(v.get(), end=' ')
        print()

        print('正在模拟计算')
        time.sleep(5)
        print('计算完成')
        self.submit_btn.config(state=NORMAL)


if __name__ == "__main__":
    app = ttk.Window("赋分", resizable=(False, False))
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()
    offset_x = int((screen_width - 600) / 2)
    offset_y = int((screen_height - 350) / 2)
    app.geometry(f'600x350+{offset_x}+{offset_y}')  # 窗口大小
    app.iconbitmap('green_apple.ico')
    App(app)
    app.mainloop()
