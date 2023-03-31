import os
import shutil
from tkinter import messagebox, simpledialog, filedialog  # 消息框，对话框，文件访问对话框

import fitz  # pymupdf库，操作PDF文件，可转换成图片
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image
from PIL import UnidentifiedImageError
from openpyxl import load_workbook

from win32com import client  # 操作office文档，转换格式


def unfreeze():
    """取消冻结文本框，清空文本框"""
    info_text.config(state=NORMAL)
    info_text.delete(1.0, END)


def word_to_pdf(word_path):
    """word转pdf"""
    file_name = os.path.basename(word_path)
    file_name = file_name[:file_name.rfind('.')]
    if not os.path.exists('tmp'):
        os.mkdir('tmp')
    pdf_path = os.path.join(os.getcwd(), 'tmp', f'{file_name}.pdf')  # 需要使用绝对路径，使用相对路径会出错

    word = client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False
    doc = word.Documents.Open(word_path)
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    word.Quit()
    return pdf_path


def long_png(path, output_name):
    """纵向拼接图片"""
    img_list = []
    for file in os.listdir(path):
        if file.endswith('.png'):
            img_list.append(Image.open(path + os.sep + file))

    # 获取总的高度以及最大的宽度
    width = 0
    height = 0
    for img in img_list:
        # 单幅图像尺寸
        w, h = img.size
        # 获取总高度
        height += h
        # 取最大的宽度作为拼接图的宽度
        width = max(width, w)

    # 创建白色的空白长图
    result = Image.new(mode='RGB', size=(width, height), color=0xffffff)

    # 拼接图片
    height = 0
    for img in img_list:
        w, h = img.size
        # 图片水平居中
        result.paste(img, box=(round(width / 2 - w / 2), height))
        height += h

    # 保存图片
    save_path = filedialog.asksaveasfilename(title='请选择图片存储路径',
                                             initialdir='F:/用户目录/桌面/',
                                             initialfile=output_name,
                                             filetypes=[('PNG', '.png')],
                                             defaultextension='.png')
    if save_path:
        result.save(save_path)
        info_text.insert(END, '图片保存成功\n', 'center')


def word_to_images():
    """word转图片"""

    word_path = filedialog.askopenfilename(title='请选择Word文档',
                                           filetypes=[('Word文档', '.docx'), ('Word文档', '.doc')],
                                           defaultextension='.docx')
    if word_path:
        pdf_path = word_to_pdf(word_path)
        pdf_to_images(pdf_path)


def pdf_to_images(pdf_path=None):
    """pdf转图片"""

    if not os.path.exists('tmp'):
        os.mkdir('tmp')

    if pdf_path is None:
        # 打开PDF文件，生成一个对象
        pdf_path = filedialog.askopenfilename(title='请选择PDF文档',
                                              filetypes=[('PDF', '.pdf')],
                                              defaultextension='.pdf')

    if pdf_path:
        unfreeze()
        file_name = os.path.basename(pdf_path)
        file_name = file_name[:file_name.rfind('.')]

        pdf = fitz.open(pdf_path)
        for page in pdf:
            pm = page.get_pixmap(dpi=150)
            pm.save(f'tmp/{page.number:0>3d}.png')
        pdf.close()

        long_png('tmp', file_name)

        # 删除临时文件
        for file in os.listdir('tmp'):
            try:
                os.remove(f'tmp/{file}')
            except PermissionError:
                pass

    freeze()


def cut_out_level():
    """裁剪图片"""

    def cut_out():
        img_list = filedialog.askopenfilenames(title='请选择图片文件',
                                               filetypes=[('PNG', '.png'), ('JPG', '.jpg')])
        if not img_list:
            top.destroy()
            return

        unfreeze()
        data_tuple = (left_sb.get(), top_sb.get(), right_sb.get(), bottom_sb.get())
        pixel_list = [0, 0, 0, 0]
        for i, value in enumerate(data_tuple):
            if value.isdigit():
                pixel_list[i] = int(value)

        for img in img_list:
            image = Image.open(img)
            # 前两个坐标点是左上角坐标，后两个坐标点是右下角坐标，width在前， height在后
            box = (pixel_list[0], pixel_list[1], image.width - pixel_list[2], image.height - pixel_list[3])
            image = image.crop(box)
            image.save(img, quality=95, optimize=True)
        top.destroy()
        info_text.insert(END, '裁剪完毕\n', 'center')
        freeze()

    top = ttk.Toplevel()
    top.title('裁剪图片')
    top.geometry(f'400x300')  # 窗口大小
    top.iconbitmap('green_apple.ico')
    top.place_window_center()

    lb = ttk.Label(top, text='请输入四边需要裁剪的像素：', font=('微软雅黑', 12))
    lb.pack(pady=10)

    entry_bar = ttk.Frame(top)
    entry_bar.pack(padx=0, pady=10)
    top_sb = ttk.Spinbox(entry_bar, from_=10, to=1000, increment=10, width=3)
    top_sb.pack(side=TOP, padx=20, pady=10)
    bottom_sb = ttk.Spinbox(entry_bar, from_=10, to=1000, increment=10, width=3)
    bottom_sb.pack(side=BOTTOM, padx=20, pady=10)
    left_sb = ttk.Spinbox(entry_bar, from_=10, to=1000, increment=10, width=3)
    left_sb.pack(side=LEFT, padx=20, pady=10)
    right_sb = ttk.Spinbox(entry_bar, from_=10, to=1000, increment=10, width=3)
    right_sb.pack(side=RIGHT, padx=20, pady=10)

    btn = ttk.Button(master=top, text='裁剪', compound=CENTER, command=cut_out)
    btn.pack(side=TOP, ipadx=12, pady=10)

    top.mainloop()


def choice_level():
    """根据输入的字母输出对应的图片"""

    def choice():
        unfreeze()
        data = text0.get(1.0, END).strip()
        img_path = filedialog.askdirectory(title='请选择答案文件夹', initialdir='F:/用户目录/桌面/')
        if not (data and img_path):
            top.destroy()
            info_text.insert(END, '没有提供足够的数据\n', 'center')
            freeze()
            return
        sb_data = start_sb.get()
        start = 0
        if sb_data.isdigit():
            start = int(sb_data) - 1
        add = ''
        if add_text.get():
            add = add_text.get()
        data_list = data.split('\n')

        counter = 0
        for row in data_list:
            counter += 1
            if len(row) == 1 and row in ('A', 'B', 'C', 'D', 'E', 'F', 'G'):
                shutil.copyfile(f'img/{row}.png', f'{img_path}/{add}{counter + start}.png')

            # 处理多选题的答案
            elif len(row) > 1:
                img_list = []
                for item in row:
                    if item in ('A', 'B', 'C', 'D', 'E', 'F', 'G'):
                        img_list.append(Image.open(f'img/{item}.png'))

                width = 0
                height = 0
                for img in img_list:
                    w, h = img.size
                    width += w
                    height = max(height, h)

                result = Image.new(mode='RGB', size=(width, height), color=0xffffff)
                width = 0
                for img in img_list:
                    w, h = img.size
                    result.paste(img, box=(width, round(height / 2 - h / 2)))
                    width += w
                result.save(f'{img_path}/{add}{counter + start}.png')
        info_text.insert(END, f'制作了{counter}个答案\n', 'center')
        freeze()
        top.destroy()

    top = ttk.Toplevel()
    top.title('制作答案')
    top.geometry(f'500x350')  # 窗口大小
    top.iconbitmap('green_apple.ico')
    top.place_window_center()

    lb = ttk.Label(top, text='选择题答案：', font=('微软雅黑', 12))
    lb.pack(pady=10)
    text0 = ttk.Text(top, width=50, height=8)
    text0.pack()
    text0.focus()
    setbar = ttk.Frame(top)
    setbar.pack(pady=10)
    ttk.Label(setbar, text='起始题号：', font=('微软雅黑', 12)).pack(side=LEFT)
    start_sb = ttk.Spinbox(setbar, from_=1, to=100, increment=1, width=3)
    start_sb.pack(side=LEFT)
    ttk.Label(setbar, text='文件名前添加：', font=('微软雅黑', 12), padding=(20, 0, 0, 0)).pack(side=LEFT)
    add_text = ttk.Entry(setbar, width=3)
    add_text.pack(side=LEFT)
    ttk.Button(master=top, text='提交', compound=CENTER, command=choice).pack(ipadx=12, pady=15)

    top.mainloop()


def splice():
    """拼接两个文件夹里的名字相同的图片，拼接成功后删除第二个文件夹的图片"""
    img_path = filedialog.askdirectory(title='请选择题目文件夹', initialdir='F:/用户目录/桌面/')
    if not img_path:
        return
    img_path2 = img_path + '答案'
    if not os.path.exists(img_path2):
        img_path2 = filedialog.askdirectory(title='请选择答案文件夹', initialdir='F:/用户目录/桌面/')
        if not img_path2:
            return
    unfreeze()
    img_list = os.listdir(img_path)
    for img in img_list:
        if os.path.exists(img_path2 + '/' + img):
            try:
                img1 = Image.open(img_path + '/' + img)
                img2 = Image.open(img_path2 + '/' + img)
            except UnidentifiedImageError:
                info_text.insert(END, '一个文件不是图片格式，打开失败\n', 'center')
                continue
            img1_height = img1.height
            new_width = max(img1.width, img2.width)
            new_height = img1.height + img2.height
            # 多留10像素，左右两边各留5像素的空白
            result = Image.new(mode='RGB', size=(new_width + 10, new_height), color=(255, 255, 255))
            result.paste(img1, box=(5, 0))
            result.paste(img2, box=(5, img1_height))
            result.save(img_path + '/' + img)
            os.remove(img_path2 + '/' + img)
        else:
            info_text.insert(END, f'{img[:-4]}没有答案\n', 'center')
    # 删除空文件夹
    if not os.listdir(img_path2):
        os.rmdir(img_path2)

    info_text.insert(END, '拼接完成\n', 'center')
    freeze()


def copy_rename():
    """根据分隔符两边的数字确认图片的数量，复制图片并改名"""
    img_list = filedialog.askopenfilenames(title='请选择图片文件',
                                           filetypes=[('PNG', '.png'), ('JPG', '.jpg')],
                                           defaultextension='.png')
    if not img_list:
        return
    unfreeze()
    file_path = os.path.dirname(img_list[0])
    for img in img_list:
        name, extension = os.path.basename(img).split('.')  # 文件名，扩展名
        fs = ''
        if name[0] in ('A', 'B'):
            fs = name[0]
            name = name[1:]
        num_list = name.split('-')
        if len(num_list) == 2 and num_list[0].isdigit() and num_list[1].isdigit():
            start = int(num_list[0])
            end = int(num_list[1])
            for i in range(start, end):
                shutil.copyfile(img, f'{file_path}/{fs}{i}.{extension}')
            os.rename(img, f'{file_path}/{fs}{end}.{extension}')
        else:
            info_text.insert(END, f'文件名 {name} 格式不正确，跳过修改 (ー_ー)!!\n', 'center')
    info_text.insert(END, '修改完成\n', 'center')
    freeze()


def add_point():
    """根据小题数量复制图片并改名"""
    img_list = filedialog.askopenfilenames(title='请选择图片文件',
                                           filetypes=[('PNG', '.png'), ('JPG', '.jpg')],
                                           defaultextension='.png')
    if not img_list:
        return
    unfreeze()
    point_str = simpledialog.askstring('输入', '请输入小题数量：')
    if point_str is None or not point_str.isdigit():
        info_text.insert(END, '必须输入纯数字\n', 'center')
        freeze()
        return
    point_num = int(point_str)
    file_path = os.path.dirname(img_list[0])
    for img in img_list:
        name, extension = os.path.basename(img).split('.')  # 文件名，扩展名
        for i in range(1, point_num):
            shutil.copyfile(img, f'{file_path}/{name}-{i}.{extension}')
        os.rename(img, f'{file_path}/{name}-{point_num}.{extension}')
    info_text.insert(END, '完成\n', 'center')
    freeze()


def rename_id():
    """把文件夹里的图片名改成Excel里的编号，复制文件夹并改名"""
    wb_file = filedialog.askopenfilename(title='请选择Excel文件',
                                         initialdir='F:/用户目录/桌面/',
                                         filetypes=[('Excel', '.xlsx')],
                                         defaultextension='.xlsx')

    img_dir = filedialog.askdirectory(title='请选择图片文件夹', initialdir='F:/用户目录/桌面/')
    if not (wb_file and img_dir):
        return

    unfreeze()
    subject_id = {'语文': '01', '数学': '02', '数学文': '03', '数学理': '04', '英语': '05', '政治': '06', '历史': '07',
                  '地理': '08', '物理': '09', '化学': '10', '生物': '11', '科学': '13', '品德与社会': '14',
                  '道德与法治': '15'}
    wb = load_workbook(wb_file)
    ws = wb.worksheets[0]
    subjects = []
    for i, row in enumerate(ws.values):
        if row[1] not in subjects:
            subjects.append(row[1])
    subjects.pop(0)  # 删除标题

    for subject in subjects:
        subject_dir = img_dir + '/' + subject
        complete = True
        for row in ws.values:
            if row[1] == subject:
                img_name = row[2]
                img_path = subject_dir + f'/{img_name}.png'
                img_id_name = row[0]
                img_id_path = subject_dir + f'/{img_id_name}.png'
                if os.path.exists(img_path):
                    shutil.copyfile(img_path, img_id_path)
                else:
                    info_text.insert(END, f'图片 {img_name}.png 不存在 (ー_ー)!!\n', 'center')
                    go_on = messagebox.askyesno(message='是否继续？')
                    if not go_on:
                        complete = False

        # 复制图片到指定目录
        if complete:
            imgs = os.listdir(subject_dir)
            os.mkdir(subject_dir + '/03')
            for img in imgs:
                os.rename(subject_dir + '/' + img, subject_dir + '/03/' + img)
            shutil.copytree(subject_dir + '/03', subject_dir + '/13')
            os.rename(subject_dir, img_dir + '/' + subject_id[subject])
            info_text.insert(END, f'{subject}处理完成\n', 'center')
    info_text.insert(END, '全部完成\n', 'center')
    wb.close()
    freeze()


def subject_dir(event):
    def make_dir():
        path = filedialog.askdirectory(title='请选择目录')
        if not path:
            return
        try:
            os.mkdir(path + os.sep + '题干图片')
            for item in var:
                dir_name = item.get()
                if not dir_name:
                    continue
                os.mkdir(f'{path}/题干图片/{dir_name}')
                os.mkdir(f'{path}/题干图片/{dir_name}答案')
        except FileExistsError:
            pass
        top.destroy()

    top = ttk.Toplevel()
    top.title('创建科目目录')
    top.geometry(f'500x200')  # 窗口大小
    top.iconbitmap('green_apple.ico')
    top.place_window_center()

    lb = ttk.Label(top, text='选择科目：', font=('微软雅黑', 12))
    lb.pack(pady=10)

    # 用循环方式添加两排复选框
    subjects = ('语文', '数学', '数学文', '数学理', '英语', '政治', '历史', '地理', '物理', '化学', '生物')
    var = []
    sj_frame = ttk.Frame(top)
    sj_frame.pack(padx=0, pady=10)
    for subject in subjects[:5]:
        var.append(ttk.StringVar())
        cb = ttk.Checkbutton(sj_frame, text=subject, variable=var[-1], onvalue=subject, offvalue='')
        cb.pack(side=LEFT, padx=5)
    sj_frame2 = ttk.Frame(top)
    sj_frame2.pack(padx=0, pady=10)
    for subject in subjects[5:]:
        var.append(ttk.StringVar())
        cb = ttk.Checkbutton(sj_frame2, text=subject, variable=var[-1], onvalue=subject, offvalue='')
        cb.pack(side=LEFT, padx=5)

    # 默认全选
    for i, subject in enumerate(subjects):
        var[i].set(subject)

    btn = ttk.Button(top, text='创建', compound=CENTER, command=make_dir)
    btn.pack(ipadx=12, pady=15)

    top.mainloop()


def freeze():
    """禁用文本框"""
    info_text.config(state=DISABLED)
    info_text.yview_moveto(1)  # 滚动到文本末尾


def show_message():
    top = ttk.Toplevel()
    top.title('软件介绍')
    top.geometry(f'500x250')  # 窗口大小
    top.maxsize(600, 350)
    top.minsize(350, 180)
    top.place_window_center()

    text0 = ttk.Text(top, width=800, height=20, spacing1=10, spacing2=10)
    text0.pack()
    text0.insert(END, 'Word/PDF转长图：字面意思\n'
                      '制作答案：读取单项选择题答案，生成答案图片。\n'
                      '拼接图片：打开两个图片文件夹，把文件名相同的图片拼接，拼接成功后会删除原始图片。\n'
                      '拆文件名：找出文件名以“-”分隔的图片，复制图片并修改图片文件名，如把A1-3.png改成A1.png，A2.png，A3.png。\n'
                      '添加编号：复制题干图片并把图片名字改成试题结构里的编号，然后把图片放进 03 和 13 文件夹，修改上级文件夹为科目编号。\n')

    text0.tag_config('forever', foreground='green', font=('黑体', 12), spacing3=5)
    text0.tag_add('forever', 1.0, END)
    text0.config(state=DISABLED)

    top.mainloop()


def about():
    messagebox.showinfo(title='关于', message='橙图 1.1\n')


root = ttk.Window(themename='cerculean', title='橙图')
root.geometry(f'600x400')  # 窗口大小
root.resizable(False, False)
root.iconbitmap('green_apple.ico')
root.bind('<Control-m>', subject_dir)
root.bind('<Control-M>', subject_dir)

menubar = ttk.Menu(root)
help_menu = ttk.Menu(menubar, tearoff=0)
help_menu.add_command(label='介绍', command=show_message)
help_menu.add_command(label='关于', command=about)
menubar.add_cascade(label='帮助', menu=help_menu)

info_text = ttk.Text(root, width=60, height=5, border=-1, font=('黑体', 12), spacing3=8)
info_text.pack(pady=20)
info_text.tag_config('center', foreground='green', justify='center')
info_text.insert(END, '请选择功能\n', 'center')
info_text.config(state=DISABLED)

buttonbar1 = ttk.Frame(root)
buttonbar1.pack(padx=0, pady=10)

btn = ttk.Button(master=buttonbar1, text='Word转长图', command=word_to_images)
btn.pack(side=LEFT, ipadx=2, padx=10)

btn = ttk.Button(master=buttonbar1, text='PDF转长图', command=pdf_to_images)
btn.pack(side=LEFT, ipadx=6, padx=10)

btn = ttk.Button(master=buttonbar1, text='图片裁剪', command=cut_out_level)
btn.pack(side=LEFT, ipadx=12, padx=10)

buttonbar2 = ttk.Frame(root)
buttonbar2.pack(padx=0, pady=10)

btn = ttk.Button(master=buttonbar2, text='制作答案', command=choice_level)
btn.pack(side=LEFT, ipadx=12, padx=10)

btn = ttk.Button(master=buttonbar2, text='拼接图片', command=splice)
btn.pack(side=LEFT, ipadx=12, padx=10)

btn = ttk.Button(master=buttonbar2, text='拆文件名', command=copy_rename)
btn.pack(side=LEFT, ipadx=12, padx=10)

btn = ttk.Button(master=buttonbar2, text='增加小题', command=add_point)
btn.pack(side=LEFT, ipadx=12, padx=10)

btn = ttk.Button(master=buttonbar2, text='添加编号', command=rename_id)
btn.pack(side=LEFT, ipadx=12, padx=10)

root.config(menu=menubar)
root.place_window_center()
root.mainloop()
