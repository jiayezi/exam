﻿"""
图像界面程序，辅助处理试题结构和小分表的信息
版本：1.3
"""
import os
from tkinter import messagebox, simpledialog, filedialog  # 消息框，对话框，文件访问对话框
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from re import search
from openpyxl import Workbook
from win32com import client


def initialize():
    """释放文本框，清空文本框"""
    info_text.config(state=NORMAL)
    output_text.delete(1.0, END)  # 删除文本框里的内容
    info_text.delete(1.0, END)


def heading():
    """提取数字"""
    initialize()
    data = input_text.get(1.0, END)
    data = data.strip()
    if data:
        data_list = data.split('\n')
        text = ''
        for i, s in enumerate(data_list):
            data_obj = search(r'\d{1,2}', s)
            if not data_obj:
                info_text.insert('end', '没有找到数字\n')
                over()
                return
            data = data_obj.group()
            text += f'{data}\n'
        output_text.insert('end', text)
        info_text.insert('end', '提取完成\n')
        output_text.focus()
    over()


def difficulty_level():
    """把数字放大100倍"""
    initialize()
    counter = 0

    data = input_text.get(1.0, END)
    data = data.strip()

    if data:
        data_list = data.split('\n')

        text = ''
        for i, s in enumerate(data_list):
            try:
                num = float(s)
                num *= 100
            except ValueError:
                text += '\n'
                info_text.insert('end', f'第 {i + 1} 行不是纯数字，处理失败\n')
            else:
                text += f'{str(int(num))}\n'
                counter += 1

        output_text.insert('end', text[:-1])
        info_text.insert('end', f'处理了 {counter} 个难度值\n')
        output_text.focus()
    over()


def single_choice():
    """提取字符串里的A、B、C、D、E、F、G"""
    initialize()
    data = input_text.get(1.0, END)  # 获取文本框里的数据
    data = data.strip()
    if data:
        counter = 0
        text = ''
        for s in data:
            if s in ('A', 'B', 'C', 'D', 'E', 'F', 'G'):
                text += f'{s}\n'
                counter += 1
        output_text.insert('end', text[:-1])
        info_text.insert('end', f'提取了 {counter} 个答案\n')
        output_text.focus()
    over()


def skill_requirements():
    """把符号替换成该列对应的文字"""
    initialize()

    data = input_text.get(1.0, END)
    data = data.strip()
    if not data:
        return

    data_list = data.split('\n')
    title = data_list[0]
    rows = data_list[1:]
    title_list = title.split('\t')

    text = ''
    for row_index, row in enumerate(rows):
        row_list = row.split('\t')
        blank = True
        for i, mark in enumerate(row_list):
            if mark.strip():
                blank = False
                text += f'{title_list[i]}/'
        if blank:
            text += f'\n'
            info_text.insert('end', f'第 {row_index + 2} 行没有符号\n')
        text = text[:-1]+'\n'
    output_text.insert('end', text[:-1])

    info_text.insert('end', '全部处理完成\n')
    output_text.focus()
    over()


def OMR():
    """删除制表符，把长度不是1的字符串替换成."""
    initialize()

    data = input_text.get(1.0, END)
    data = data.strip().replace(' ', '.')
    data_list = data.split('\n')

    counter = 0
    for line in data_list:
        line_list = line.split('\t')
        for s in line_list:
            if len(s) == 1:
                output_text.insert('end', f'{s}')
            else:
                output_text.insert('end', '.')
                counter += 1
        output_text.insert('end', '\n')

    info_text.insert('end', f'替换 {counter} 处多选\n')
    output_text.focus()
    over()


def multiple_OMR():
    """合并不定向选择答案"""
    initialize()

    data = input_text.get(1.0, END)
    data = data.strip().replace(' ', '.')
    data_list = data.split('\n')

    for line in data_list:
        line_list = line.split('\t')
        for s in line_list:
            if len(s) > 0:
                output_text.insert('end', f'[{s}]')
            else:
                output_text.insert('end', '[.]')
        output_text.insert('end', '\n')

    output_text.focus()
    over()


def format_table():
    """把小分表修改成指定格式的Excel文档，方便上传"""
    initialize()

    # 打开Excel表格
    open_path = filedialog.askopenfilename(title='请选择Excel文件', filetypes=[('Excel', '.xlsx')],
                                           defaultextension='.xlsx')
    if not open_path:
        return
    excel = client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(open_path, False)
    ws = wb.Worksheets(1)
    max_row = ws.UsedRange.Rows.Count

    # 从文件名提取编号
    file_name = os.path.split(open_path)[1]
    num = file_name[:file_name.rfind('.')]

    # 检查选择题答案的数量
    range_obj = ws.Range('C2')
    range_obj.EntireColumn.Insert()
    ws.Cells(2, 3).Value = '=len(B2)'

    sourceRange = ws.Range(ws.Cells(2, 3), ws.Cells(2, 3))
    fillRange = ws.Range(ws.Cells(2, 3), ws.Cells(max_row, 3))
    sourceRange.AutoFill(Destination=fillRange)

    ws.Cells(max_row + 1, 3).Value = f'=MIN(C2:C{max_row})'
    ws.Cells(max_row + 2, 3).Value = f'=MAX(C2:C{max_row})'

    if ws.Cells(max_row + 1, 3).Value == ws.Cells(max_row + 2, 3).Value:
        info_text.insert(END, f'选择题答案数量一样\n')
    else:
        info_text.insert(END, '选择题答案数量不一样，请检查这个科目是否有多选题 Σ(ŎдŎ|||)ﾉﾉ\n')
    ws.Cells(max_row + 1, 3).Value = None
    ws.Cells(max_row + 2, 3).Value = None
    ws.Columns(3).Delete()

    # 找出每个小题的最大得分
    for col in range(3, ws.UsedRange.Columns.Count + 1):
        # 通过数字获取列号
        col_name_fun = f'=SUBSTITUTE(ADDRESS(1,{col},4),1,"")'
        ws.Cells(max_row + 1, col).Value = col_name_fun
        col_name = ws.Cells(max_row + 1, col).Value
        # 计算最大值
        ws.Cells(max_row + 2, col).Value = f'=MAX({col_name}2:{col_name}{max_row})'
        output_text.insert(END, f'{ws.Cells(max_row + 2, col).Value}\n')
    info_text.insert('end', '题目最高分查找完成\n')
    # 删除2行临时数据
    ws.Rows(max_row + 1).Delete()
    ws.Rows(max_row + 1).Delete()

    # 设置表格为文本格式
    ws.Cells.NumberFormatLocal = "@"

    # 插入单行单列
    range_obj = ws.Range('A1')
    range_obj.EntireRow.Insert()
    range_obj.EntireColumn.Insert()

    ws.Cells(1, 1).Value = num
    ws.Cells(2, 1).Value = '班级'
    ws.Cells(3, 1).Value = '0' * 16

    # 模拟自动填充
    sourceRange = ws.Range(ws.Cells(3, 1), ws.Cells(3, 1))
    fillRange = ws.Range(ws.Cells(3, 1), ws.Cells(ws.UsedRange.Rows.Count, 1))
    sourceRange.AutoFill(Destination=fillRange)

    wb.Close(SaveChanges=1)  # 保存并关闭
    excel.Quit()

    info_text.insert(END, '小分表修改完成\n')
    info_text.yview_moveto(1)
    output_text.focus()

    over()


def calculate_total_score():
    """把每个学生的单科成绩相加，计算总分"""
    initialize()

    all_data_list = []
    counter = 0
    while True:
        data = simpledialog.askstring('输入成绩', '请输入考号和单科成绩：')
        if not data:
            return
        all_data_list.append(data)
        counter += 1
        info_text.insert(END, f'已提交 {counter} 个科目成绩\n')
        info_text.tag_add('forever', 1.0, END)
        info_text.tag_config('forever', foreground='green', font=('黑体', 11), justify="center", spacing3=5)
        choice = messagebox.askyesno('添加确认', '是否继续添加其他科目成绩？')
        if not choice:
            break

    # 存储考号和分数的字典
    student_dict = {}
    titles = ['考号']
    for index, data in enumerate(all_data_list):
        data = data.strip()
        row_list = data.split('\n')

        try:
            for row in row_list:
                singe_list = row.split('\t')
                # 判断考号
                student_id = singe_list[0]
                if len(student_id) < 5 or not student_id.isdigit():
                    titles.append(singe_list[1])
                    continue
                score = float(singe_list[1])

                # 如果是新增的考号，为保证科目与分数对应，需要让该考号的之前的科目为0分
                if student_id not in student_dict:
                    student_dict[student_id] = [0.0 for _ in range(index + 1)]
                student_dict[student_id][index] = score
        except IndexError:
            info_text.insert(END, '数据不完整，处理失败，请同时提交考号和成绩 (ー_ー)!!\n')
            return
        except ValueError:
            info_text.insert(END, '成绩字段不是纯数字，处理失败 (ー_ー)!!\n')
            return
        # 把所有学生的下一个科目的分数初始化为0分
        if index < len(all_data_list) - 1:
            for key in student_dict:
                student_dict[key].append(0.0)

    titles.append('总分')
    wb = Workbook()
    ws = wb.active
    ws.title = '总分表'

    # 添加首行
    if len(titles) > 2:
        ws.append(titles)
    else:
        row_data = ['考号'] + list(range(len(all_data_list))) + ['总分']
        ws.append(row_data)
    # 添加数据
    for row_index, key in enumerate(student_dict):
        row_data = [key]
        row_data.extend(student_dict[key])
        row_data.append(sum(student_dict[key]))
        ws.append(row_data)

    file_path = filedialog.asksaveasfilename(title='请选择文件存储路径',
                                             initialdir='F:/用户目录/桌面/',
                                             initialfile='总分表',
                                             filetypes=[('Excel', '.xlsx')],
                                             defaultextension='.xlsx')
    if file_path:
        wb.save(file_path)
        info_text.insert(END, '文件保存成功\n')
        wb.close()
    over()


def split_score():
    """按小题的分数拆分总分"""
    initialize()

    data = input_text.get(1.0, END).strip()
    total_score_list = data.split('\n')

    small_data = simpledialog.askstring('提交分数', '请输入试题结构的题目分数：')
    if not small_data:
        return
    small_data = small_data.strip()
    small_score_list = small_data.split('\n')
    try:
        for total_score in total_score_list:
            total_score = float(total_score)
            for small_score in small_score_list:
                small_score = float(small_score)
                if total_score >= small_score:
                    output_text.insert('end', f'{small_score}\t')
                    total_score -= small_score
                else:
                    output_text.insert('end', f'{total_score}\t')
                    total_score = 0
            output_text.insert('end', '\n')
    except ValueError:
        info_text.insert('end', '总分或题目分不是纯数字，拆分失败 (ー_ー)!!\n')
    else:
        info_text.insert('end', f'拆分完毕\n')
        output_text.focus()

    over()


def over():
    """改变文本颜色，禁用文本框"""
    info_text.tag_add('forever', 1.0, END)
    # 使用 tag_config() 来改变标签"forever"的文字颜色和大小
    info_text.tag_config('forever', foreground='green', font=('黑体', 11), justify="center", spacing3=5)
    info_text.config(state=DISABLED)
    info_text.yview_moveto(1)  # 文本更新滚动显示


def show_message():
    top = ttk.Toplevel()
    top.title('软件介绍')
    top.geometry(f'600x320+{offset_x+240}+{offset_y+180}')  # 窗口大小
    top.maxsize(700, 400)
    top.minsize(500, 200)

    text0 = ttk.Text(top, width=100, height=20, spacing2=10, spacing3=15)
    text0.pack()
    text0.insert(END, '题目：每行提取一个最多两位的数字\n'
                      '难度值：把数字放大100倍\n'
                      '单选答案：从文子里提取A-G的大写字母\n'
                      '能力要求：把文字里的“√”替换成第一行对应的能力要求\n'
                      'OMR：合并所有列，把多选题答案和空白替换成“.”\n'
                      '多选OMR：合并所有列，把每个多选题答案放进中括号里\n'
                      '小分表：读取原始小分表，检查题目分数，在第一列插入16个0，第一行插入科目编号，另存为新的小分表\n'
                      '总分：输入考号和单科成绩，生成总分表\n'
                      '拆分：按照小题分数把每个学生的总分拆分成小分\n')

    text0.tag_add('forever', 1.0, END)
    text0.tag_config('forever', foreground='green', font=('黑体', 12))
    text0.config(state=DISABLED)

    top.mainloop()


def about():
    messagebox.showinfo(title='关于', message='橙技 1.0\n')


def select_all(event):
    info_text.config(state=NORMAL)
    info_text.insert(END, '选中全部\n')
    over()


def cp_msg(event):
    info_text.config(state=NORMAL)
    info_text.insert(END, '已复制到剪贴板\n')
    over()


def close_handle():
    if messagebox.askyesno(title='退出确认', message='确定要退出吗？'):
        root.destroy()


# 窗口
root = ttk.Window(themename='cerculean', title='橙技')
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
offset_x = int((screen_width-1080)/2)
offset_y = int((screen_height-800)/2)
root.geometry(f'1080x800+{offset_x}+{offset_y}')  # 窗口大小
root.minsize(1000, 750)
root.iconbitmap('green_apple.ico')

# 菜单
main_menu = ttk.Menu(root)
sub_menu = ttk.Menu(main_menu, tearoff=0)
sub_menu.add_command(label='介绍', command=show_message)
sub_menu.add_command(label='关于', command=about)
main_menu.add_cascade(label='帮助', menu=sub_menu)

label1 = ttk.Label(root, text='原始数据', font=('黑体', 12))
label1.pack(pady=10)  # 按布局方式放置标签

input_text = ttk.Text(root, height=12)
input_text.pack(fill=X, padx=100)  # 文本框宽度沿水平方向自适应填充，左右两边空100像素
input_text.focus()

label2 = ttk.Label(root, text='计算结果', font=('黑体', 12))
label2.pack(pady=10)

output_text = ttk.Text(root, height=12)
output_text.pack(fill=X, padx=100)
output_text.bind('<Control-a>', select_all)  # 绑定事件
output_text.bind('<Control-A>', select_all)
output_text.bind('<Control-c>', cp_msg)
output_text.bind('<Control-C>', cp_msg)
output_text.bind('<Control-x>', cp_msg)
output_text.bind('<Control-X>', cp_msg)

info_text = ttk.Text(root, height=6, border=-1)
info_text.pack(pady=10, padx=100, fill=X)
info_text.config(state=DISABLED)

# 按钮区域
buttonbar = ttk.Labelframe(root, text='选择功能', labelanchor="n")
buttonbar.pack(pady=0,  padx=100, ipady=20)

btn = ttk.Button(master=buttonbar, text='题目', compound=LEFT, command=heading)
btn.pack(side=LEFT, padx=20)

btn = ttk.Button(master=buttonbar, text='难度值', compound=LEFT, command=difficulty_level)
btn.pack(side=LEFT, padx=20)

btn = ttk.Button(master=buttonbar, text='单选答案', compound=LEFT, command=single_choice)
btn.pack(side=LEFT, padx=18)

btn = ttk.Button(master=buttonbar, text='能力要求', compound=LEFT, command=skill_requirements)
btn.pack(side=LEFT, padx=15)

btn = ttk.Button(master=buttonbar, text='OMR', compound=LEFT, command=OMR)
btn.pack(side=LEFT, padx=15)

btn = ttk.Button(master=buttonbar, text='多选OMR', compound=LEFT, command=multiple_OMR)
btn.pack(side=LEFT, padx=15)

btn = ttk.Button(master=buttonbar, text='小分表', compound=LEFT, command=format_table)
btn.pack(side=LEFT, padx=18)

btn = ttk.Button(master=buttonbar, text='总分表', compound=LEFT, command=calculate_total_score)
btn.pack(side=LEFT, padx=20)

btn = ttk.Button(master=buttonbar, text='拆分', compound=LEFT, command=split_score)
btn.pack(side=LEFT, padx=20)

root.protocol('WM_DELETE_WINDOW', close_handle)  # 启用协议处理机制，点击关闭时按钮，触发事件

root.config(menu=main_menu)
root.mainloop()
