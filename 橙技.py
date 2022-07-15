"""
图像界面程序，辅助处理试题结构和小分表的信息
版本：1.0
"""

from tkinter import messagebox, simpledialog, filedialog  # 消息框，对话框，文件访问对话框
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from re import search
from openpyxl import Workbook, load_workbook
from passing_rate import student_rate


def initialize():
    """释放文本框，清空文本框"""
    text3.config(state=NORMAL)
    text2.delete(1.0, END)  # 删除文本框里的内容
    text3.delete(1.0, END)


def timu():
    """提取数字"""
    initialize()
    data = text1.get(1.0, END)
    data = data.strip()
    if data:
        data_list = data.split('\n')

        for i, s in enumerate(data_list):
            data_obj = search(r'\d{1,2}', s)
            data = data_obj.group()
            text2.insert('end', f'{data}\n')
        text3.insert('end', '提取完成\n')
        text2.focus()
    over()


def nandu():
    """把数字放大100倍"""
    initialize()
    counter = 0

    data = text1.get(1.0, END)
    data = data.strip()

    if data:
        data_list = data.split('\n')

        for i, s in enumerate(data_list):
            try:
                num = float(s)
                num *= 100
            except ValueError:
                text2.insert('end', '\n')
                text3.insert('end', f'第 {i + 1} 行不是纯数字，处理失败\n')
            else:
                text2.insert('end', f'{str(int(num))}\n')
                counter += 1
                text3.insert('end', f'处理了 {counter} 个难度值\n')

        text3.insert('end', '全部处理完成\n')
        text2.focus()
    over()


def xuanzeti():
    """提取字符串里的A、B、C、D、E、F、G"""
    initialize()
    data = text1.get(1.0, END)  # 获取文本框里的数据
    data = data.strip()
    if data:
        counter = 0
        for i in data:
            if i in ('A', 'B', 'C', 'D', 'E', 'F', 'G'):
                text2.insert('end', f'{i}\n')
                counter += 1
        text3.insert('end', f'全部提取完成，发现 {counter} 个答案\n')
        text2.focus()
    over()


def nengli():
    """把√替换成该列对应的文字"""
    initialize()

    data = text1.get(1.0, 2.0)  # 获取第一行
    data = data.strip()

    if data:
        title_list = data.split('\t')
        row = 2.0
        while data:
            data = text1.get(row, row + 1)
            data = data[:-1]
            data_list = data.split('\t')
            if len(data_list) > 1:  # 跳过末尾的换行符
                try:
                    n = data_list.index('√')
                except ValueError:
                    text2.insert('end', '\n')
                    text3.insert('end', f'第 {int(row)} 行没有“√”\n')
                else:
                    text2.insert('end', f'{title_list[n]}\n')
            row += 1
        text3.insert('end', '全部处理完成\n')
        text2.focus()
    over()


def omr():
    """删除制表符，把长度不是1的字符串替换成."""
    initialize()

    data = text1.get(1.0, END)
    data = data.strip()
    data = data.replace(' ', '.')

    data_list = data.split('\n')

    counter = 0
    for line in data_list:
        line_list = line.split('\t')
        for s in line_list:
            if len(s) == 1:
                text2.insert('end', f'{s}')
            else:
                text2.insert('end', '.')
                counter += 1
        text2.insert('end', '\n')

    text3.insert('end', f'替换 {counter} 处多选\n')
    text2.focus()
    over()


def buding():
    """合并不定向选择答案"""
    initialize()

    data = text1.get(1.0, END)
    data = data.strip()
    data = data.replace(' ', '.')

    data_list = data.split('\n')

    for line in data_list:
        line_list = line.split('\t')
        for s in line_list:
            text2.insert('end', f'[{s}]')
        text2.insert('end', '\n')

    text2.focus()
    over()


def xiaofen():
    """把小分表修改成指定格式的Excel文档，方便上传"""
    initialize()

    # 打开Excel表格
    open_path = filedialog.askopenfilename(title='请选择Excel文件', filetypes=[('Excel', '.xlsx')],
                                           defaultextension='.xlsx')
    if open_path:
        wb_old = load_workbook(open_path)
        ws_old = wb_old.worksheets[0]

        # 检查选择题答案的数量
        for row in range(3, ws_old.max_row + 1):
            try:
                length = len(ws_old.cell(row, 2).value)
                if length != len(ws_old.cell(2, 2).value):
                    text3.insert(END, '选择题答案数量不一样，请检查这个科目是否有多选题 Σ(ŎдŎ|||)ﾉﾉ\n')
                    break
            except TypeError:
                text3.insert(END, '表格第二列没有文字，无法计算长度 (ー_ー)!!\n')
                break
        else:
            text3.insert(END, f'选择题答案数量一样\n')

        # 找出每个小题的最大得分
        max_list = []
        max_num = 0.0
        for col in range(3, ws_old.max_column + 1):
            for row in range(2, ws_old.max_row + 1):
                if ws_old.cell(row, col).value is None:
                    continue
                try:
                    singe_num = float(ws_old.cell(row, col).value)
                except ValueError:
                    text3.insert('end', f'第 {row} 行 {col} 列不是纯数字\n')
                    continue
                if singe_num > max_num:
                    max_num = singe_num
            max_list.append(max_num)
            max_num = 0.0
        for n in max_list:
            text2.insert(END, f'{n}\n')
        text3.insert('end', '题目最高分查找完成\n')
        text3.tag_add('forever', 1.0, END)

        # 新建Excel表格
        wb_new = Workbook()
        ws_new = wb_new.active

        # 读取原始文档，写入到新文档
        for row in range(1, ws_old.max_row + 1):  # 遍历整个工作表
            for col in range(1, ws_old.max_column + 1):
                if ws_old.cell(row, col).value is None:
                    ws_new.cell(row, col, str(0))
                else:
                    ws_new.cell(row, col, str(ws_old.cell(row, col).value))
        # 插入一列
        ws_new.insert_cols(1)
        ws_new.cell(1, 1, '班级')
        for row in range(1, ws_old.max_row):
            ws_new.cell(row + 1, 1, '0' * 16)

        # 插入行
        ws_new.insert_rows(1)
        num = simpledialog.askstring(' ', '请输入科目编号')
        if num:
            num = num.strip()
            ws_new.cell(1, 1, num)

        text3.insert(END, '小分表已生成\n')
        text3.tag_add('forever', 1.0, END)
        text3.yview_moveto(1)
        wb_old.close()

        save_path = filedialog.asksaveasfilename(title='请选择文件存储路径',
                                                 initialdir='F:/用户目录/桌面/',
                                                 initialfile=num,
                                                 filetypes=[('Excel', '.xlsx')],
                                                 defaultextension='.xlsx')
        if save_path:
            wb_new.save(save_path)
            text3.insert('end', '小分表保存成功\n')
            wb_new.close()
        text2.focus()
    else:
        text3.insert('end', '没有打开Excel文件\n')

    over()


def total_score():
    """把每个学生的单科成绩相加，计算总分"""
    initialize()

    # 存储考号和分数的字典
    student_dict = {}
    counter = 0
    titles = []
    while True:
        data = simpledialog.askstring('输入成绩', '请输入考号和单科成绩：                                                  ')
        if data:
            data = data.strip()
            data_list = data.split('\n')

            try:
                for line in data_list:
                    singe_list = line.split('\t')

                    # 判断考号
                    student_id = singe_list[0]
                    if len(student_id) < 5 or not student_id.isdigit():
                        titles.append(singe_list[1])
                        continue
                    score = float(singe_list[1])

                    # 如果是新增的考号，为保证科目与分数对应，需要让该考号的之前的科目为0分
                    if student_id not in student_dict:
                        student_dict[student_id] = [0.0 for _ in range(counter + 1)]
                    student_dict[student_id][counter] = score
            except IndexError:
                text3.insert(END, '数据不完整，处理失败，请同时提交考号和成绩 (ー_ー)!!\n')
                break
            except ValueError:
                text3.insert(END, '成绩字段不是纯数字，处理失败 (ー_ー)!!\n')
                break

            counter += 1
            text3.insert(END, f'已提交 {counter} 个科目成绩\n')
            text3.tag_add('forever', 1.0, END)
            text3.yview_moveto(1)

            choice = messagebox.askyesno('添加确认', '是否继续添加其他科目成绩？')
            if choice:
                # 每添加一个科目，就添加一列0分，如果考号的科目有分，就用分数替换0，否则就保持0分
                for key in student_dict:
                    student_dict[key].append(0.0)
            else:
                break
        else:
            text3.insert(END, '没有输入考号和分数\n')
            break

    if len(student_dict):

        # 输出成绩到文本框
        for key in student_dict:
            text2.insert(END, f'{key}\t')
            for single_score in student_dict[key]:
                text2.insert(END, f'{str(single_score)}\t')
            total = str(sum(student_dict[key]))
            text2.insert(END, total)
            text2.insert(END, '\n')
        text3.insert(END, '总分计算完成\n')
        text3.tag_add('forever', 1.0, END)
        text3.yview_moveto(1)

        # 输出成绩到Excel表格
        wb = Workbook()
        ws = wb.active
        ws.title = '总分表'

        # 添加首行
        ws.cell(1, 1, '考号')
        if len(titles) > 1:
            for col, title in enumerate(titles):
                ws.cell(1, col + 2, title)
        else:
            for i in range(counter):
                ws.cell(1, i + 2, f'科目{i + 1}')
        ws.cell(1, counter + 2, '总分')

        # 添加数据
        for row, key in enumerate(student_dict):
            ws.cell(row + 2, 1, key)
            for col, single_score in enumerate(student_dict[key]):
                ws.cell(row + 2, col + 2, single_score)
            total = sum(student_dict[key])
            ws.cell(row + 2, len(student_dict[key]) + 2, total)

        file_path = filedialog.asksaveasfilename(title='请选择文件存储路径',
                                                 initialdir='F:/用户目录/桌面/',
                                                 initialfile='总分表',
                                                 filetypes=[('Excel', '.xlsx')],
                                                 defaultextension='.xlsx')
        if file_path:
            wb.save(file_path)
            text3.insert(END, '文件保存成功\n')
            wb.close()
    over()


def chaifen():
    """按小题的分数拆分总分"""
    initialize()

    total = text1.get(1.0, END)
    total = total.strip()
    total_list = total.split('\n')

    score = simpledialog.askstring('提交分数', '请输入试题结构的题目分数：                                       ')

    if score:
        score = score.strip()
        score_list = score.split('\n')
        counter = 0
        try:
            for total in total_list:
                num_total = float(total)
                for score in score_list:
                    num_score = float(score)
                    if num_total >= num_score:
                        text2.insert('end', f'{score}\t')
                        num_total -= num_score
                    else:
                        text2.insert('end', f'{str(num_total)}\t')
                        num_total = 0
                counter += 1
                text2.insert('end', '\n')
                text3.insert('end', f'拆散 {counter} 个总分\n')
        except ValueError:
            text3.insert('end', '总分或题目分不是纯数字，拆分失败 (ー_ー)!!\n')
        else:
            text3.insert('end', '全部拆分完成\n')
            text2.focus()
    else:
        text3.insert('end', '没有输入题目分数\n')

    over()


def over():
    """改变文本颜色，禁用文本框"""
    text3.tag_add('forever', 1.0, END)
    text3.config(state=DISABLED)
    text3.yview_moveto(1)  # 文本更新滚动显示


def show_message():
    top = ttk.Toplevel()
    top.title('软件介绍')
    top.geometry('600x320+680+300')  # 窗口大小
    top.maxsize(700, 400)
    top.minsize(500, 200)

    text0 = ttk.Text(top, width=100, height=20, spacing1=10, spacing2=10)
    text0.pack()
    text0.insert(END, '本软件用于处理考试成绩之类的外部数据，减少工作过程中经常遇到的复杂和重复的操作。\n\n'
                      '复制需要处理的文本，粘贴到第一个文本框，点击下方对应的按钮，第二个文本框会显示处理结果。\n\n'
                      '数字：提取文本里的数字\n'
                      '难度值：处理双向细目表里的不大于1的难度值，把数字放大100倍\n'
                      '单选答案：处理试卷文档里的单选题答案，从文字里提取大写字母A-G\n'
                      '能力要求：处理双向细目表里的能力层次信息，把“√”替换成第一行对应的能力层次\n'
                      'OMR：处理原始小分表里的单选题答案，把多选答案和空白替换成“.”\n'
                      '不定项OMR：处理原始小分表里的不定项选选题答案\n'
                      '小分表：读取原始小分表，检查题目分数，在第一列插入16个0，第一行插入科目编号，另存为新小分表\n'
                      '总分：输入考号和单科成绩，对每个学生的单科成绩求和，输出单科成绩和总分，另存为总分表\n'
                      '拆分：按照小题分数把每个学生的总分拆分成小分\n'
                      'ctrl+y：计算及格率和不及格率')

    text0.tag_config('forever', foreground='green', font=('黑体', 12), spacing3=5)
    text0.tag_add('forever', 1.0, END)
    text0.config(state=DISABLED)

    top.mainloop()


def about():
    messagebox.showinfo(title='关于', message='橙技 1.0\n'
                                            'by 李清萍\n'
                                            'QQ 1601235906\n')


def select_all(event):
    text3.config(state=NORMAL)
    text3.insert(END, '选中全部\n')
    over()


def cp_msg(event):
    text3.config(state=NORMAL)
    text3.insert(END, '已复制到剪贴板\n')
    over()


def close_handle():
    if messagebox.askyesno(title='退出确认', message='确定要退出吗？'):
        root.destroy()


# 窗口
root = ttk.Window(themename='cerculean', title='橙技')
root.geometry('1080x800+450+100')  # 窗口大小
root.minsize(1000, 750)
root.bind('<Control-y>', student_rate)

# 菜单
menubar = ttk.Menu(root)
help_menu = ttk.Menu(menubar, tearoff=0)
help_menu.add_command(label='介绍', command=show_message)
help_menu.add_command(label='关于', command=about)
menubar.add_cascade(label='帮助', menu=help_menu)

label1 = ttk.Label(root, text='原始数据', font=('黑体', 12), width=10)
label1.pack(pady=10, side=TOP)  # 按布局方式放置标签

text1 = ttk.Text(root, width=120, height=14)
text1.pack()
text1.focus()

label2 = ttk.Label(root, text='处理结果', font=('黑体', 12), width=10)
label2.pack(pady=10, side=TOP)

text2 = ttk.Text(root, width=120, height=14)
text2.pack()
text2.bind('<Control-a>', select_all)  # 绑定事件
text2.bind('<Control-A>', select_all)
text2.bind('<Control-c>', cp_msg)
text2.bind('<Control-C>', cp_msg)
text2.bind('<Control-x>', cp_msg)
text2.bind('<Control-X>', cp_msg)

text3 = ttk.Text(root, width=120, height=6, border=-1)
text3.pack(pady=10, side=TOP)
text3.tag_config('forever', foreground='green', font=('黑体', 11), spacing3=5)
text3.config(state=DISABLED)

# 按钮区域
buttonbar = ttk.Frame(root)
buttonbar.pack(padx=10, pady=20, side=BOTTOM)

btn = ttk.Button(master=buttonbar, text='数字', compound=LEFT, command=timu)
btn.pack(side=LEFT, ipadx=15, padx=10, pady=5)

btn = ttk.Button(master=buttonbar, text='难度值', compound=LEFT, command=nandu)
btn.pack(side=LEFT, ipadx=15, padx=10, pady=5)

btn = ttk.Button(master=buttonbar, text='单选答案', compound=LEFT, command=xuanzeti)
btn.pack(side=LEFT, ipadx=15, padx=10, pady=5)

btn = ttk.Button(master=buttonbar, text='能力要求', compound=LEFT, command=nengli)
btn.pack(side=LEFT, ipadx=15, padx=10, pady=5)

btn = ttk.Button(master=buttonbar, text='OMR', compound=LEFT, command=omr)
btn.pack(side=LEFT, ipadx=15, padx=10, pady=5)

btn = ttk.Button(master=buttonbar, text='不定项OMR', compound=LEFT, command=buding)
btn.pack(side=LEFT, ipadx=15, padx=10, pady=5)

btn = ttk.Button(master=buttonbar, text='小分表', compound=LEFT, command=xiaofen)
btn.pack(side=LEFT, ipadx=15, padx=10, pady=5)

btn = ttk.Button(master=buttonbar, text='总分表', compound=LEFT, command=total_score)
btn.pack(side=LEFT, ipadx=15, padx=10, pady=5)

btn = ttk.Button(master=buttonbar, text='拆分', compound=LEFT, command=chaifen)
btn.pack(side=LEFT, ipadx=15, padx=10, pady=5)

root.protocol('WM_DELETE_WINDOW', close_handle)  # 点击关闭按钮，触发事件

root.config(menu=menubar)
root.mainloop()
