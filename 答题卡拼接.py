import os
import time
from PIL import Image
from concurrent.futures import ThreadPoolExecutor
import ttkbootstrap as ttk
from tkinter import filedialog, messagebox


def get_src_dir():
    path = filedialog.askdirectory(title='请选择图片目录', initialdir='F:/用户目录/桌面/')
    src_text.set(path)


def get_dst_dir():
    path = filedialog.askdirectory(title='请选择存储目录', initialdir='F:/用户目录/桌面/')
    dst_text.set(path)


def submit_task():
    def splice(filesname):
        """上下拼接两张图片"""
        img1 = Image.open(input_path + '/' + filesname[0])
        img2 = Image.open(input_path + '/' + filesname[1])
        img1_height = img1.height
        new_width = max(img1.width, img2.width)
        new_height = img1.height + img2.height
        result = Image.new(mode='RGB', size=(new_width, new_height), color=(255, 255, 255))
        result.paste(img1, box=(0, 0))
        result.paste(img2, box=(0, img1_height))
        save_name = filesname[0].split('.')[0]
        save_name = save_name.replace(del_text, '')
        save_path = output_path + '/' + save_name
        result.save(save_path, format='JPEG', optimize=True, quality=quality)
        img1.close()
        img2.close()
        result.close()
        os.rename(save_path, save_path + '.png')

    # 使用递归调用，迟早会把栈的空间用完，处理大量数据时不推荐
    # def update_progress():
    #     """定时更新进度条"""
    #     done = [f for f in futures if f.done()]
    #     for f in done:
    #         del futures[f]
    #         progress['value'] += step
    #     if len(futures) == 0:
    #         progress['value'] = 100
    #         messagebox.showinfo(message='全部处理完成')
    #         submit_btn.config(state=ttk.NORMAL)
    #         return
    #     root.after(200, update_progress)

    def update_progress():
        """定时更新进度条"""
        # 等待futures列表不为空
        time.sleep(1)
        # 查找并删除已完成的线程对象，增加进度条的值
        while futures:
            done = [f for f in futures if f.done()]
            for f in done:
                del futures[f]
                progress['value'] += step
            time.sleep(0.2)

        progress['value'] = 100
        messagebox.showinfo(message='全部处理完成')
        submit_btn.config(state=ttk.NORMAL)

    # 获取用户输入的数据，判断是否有效
    input_path = src_text.get()
    output_path = dst_text.get()
    if not os.path.exists(output_path):
        os.mkdir(output_path)
    del_text = del_entry.get()
    max_number = 32
    if cbox.get().isdigit():
        max_number = int(cbox.get())
    quality = round(sc.get())  # 如果sc.get()的值是小数，PIl可能无法保存图片

    if not os.path.isdir(input_path):
        messagebox.showerror('错误', '读取目录不存在！')
        return
    if input_path == output_path:
        messagebox.showwarning('警告', '读取目录和保存目录不能相同！')
        return
    img_list = os.listdir(input_path)
    img_list_count = len(img_list)
    if img_list_count < 2:
        messagebox.showerror('错误', '该目录下没有足够的文件！')
        return
    if img_list_count & 1 == 1:
        messagebox.showerror('错误', '图片数量不是偶数！')
        return

    submit_btn.config(state=ttk.DISABLED)
    progress['value'] = 0

    # 每两张图片为一组
    file_count = 0
    splice_file_list = []
    for i in range(0, len(img_list), 2):
        splice_file_list.append((img_list[i], img_list[i + 1]))
        file_count += 1

    # 每处理完一张图片，进度条需要增加的值
    step = 100.0 / file_count

    # 创建进度条线程和图片拼接线程
    pool = ThreadPoolExecutor(max_workers=max_number)
    pool.submit(update_progress)
    futures = {}
    for filesname in splice_file_list:
        future = pool.submit(splice, filesname=filesname)  # 提交任务
        futures[future] = filesname
    pool.shutdown(wait=False)


# 设置窗口
root = ttk.Window(themename='cerculean', title='答题卡拼接')

# 设置控件
frame = ttk.Frame(root, padding=20)
frame.grid(row=0, column=0)
ttk.Label(frame, text='读取目录：', font=('黑体', 12)).grid(row=0, column=0, pady=5)
ttk.Label(frame, text='存储目录：', font=('黑体', 12)).grid(row=1, column=0, pady=5)
src_text = ttk.StringVar()
dst_text = ttk.StringVar()
ttk.Entry(frame, textvariable=src_text, width=35).grid(row=0, column=1, pady=5)
ttk.Entry(frame, textvariable=dst_text, width=35).grid(row=1, column=1, pady=5)

ttk.Button(master=frame, text='浏览', command=get_src_dir).grid(row=0, column=2, padx=(10, 0), pady=5)
ttk.Button(master=frame, text='浏览', command=get_dst_dir).grid(row=1, column=2, padx=(10, 0), pady=5)

ttk.Label(frame, text='删除字符：', font=('黑体', 12)).grid(row=2, column=0, pady=5)
del_entry = ttk.Entry(frame, width=35)
del_entry.grid(row=2, column=1, pady=5)

ttk.Label(frame, text='最大线程：', font=('黑体', 12)).grid(row=3, column=0, pady=5)
cbox = ttk.Combobox(frame, values=('4', '8', '16', '32', '64', '128'), width=3)
cbox.grid(row=3, column=1, pady=5, sticky='w')
cbox.current(3)

ttk.Label(frame, text='图片质量：', font=('黑体', 12)).grid(row=4, column=0, pady=5)
sc = ttk.Scale(frame, from_=0, to=100, value=33, length=260, command=lambda value: sc_value.set(f"{float(value):.0f}"))
sc.grid(row=4, column=1, pady=5)
sc_value = ttk.StringVar()
sc_value.set('33')
ttk.Label(frame, textvariable=sc_value).grid(row=4, column=2, pady=5)

submit_btn = ttk.Button(frame, text='提交', command=submit_task)
submit_btn.grid(row=5, column=0, columnspan=3, ipadx=10, pady=20)

progress = ttk.Progressbar(frame, length=380, value=0)
progress.grid(row=6, column=0, columnspan=3, pady=10)

root.place_window_center()
root.iconbitmap('green_apple.ico')
root.mainloop()
