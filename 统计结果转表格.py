import ttkbootstrap as ttk


def paste_from_clipboard(event):
    clipboard_text = root.clipboard_get()
    input_text.insert('end', clipboard_text)


def copy_to_clipboard(event):
    selected_text = output_text.get(1.0, 'end')
    root.clipboard_clear()
    root.clipboard_append(selected_text)

    info_text.config(state='normal')
    info_text.insert('end', '已复制到剪贴板\n', 'center')


def unfreeze():
    """释放文本框，清空文本框"""
    info_text.config(state='normal')
    output_text.delete(1.0, 'end')  # 删除文本框里的内容
    info_text.delete(1.0, 'end')


def freeze():
    """改变文本颜色，禁用文本框"""
    info_text.config(state='disabled')
    info_text.yview_moveto(1)  # 滚动到文本末尾
    input_text.delete(1.0, 'end')


def table():
    def split_list(input_list, chunk_size):
        result = []
        for i in range(0, len(input_list), chunk_size):
            result.append(input_list[i:i + chunk_size])
        return result

    unfreeze()
    data = input_text.get(1.0, 'end')
    if data:
        data_list = data.split('\n')
        # 获取题目的选项数量
        key_row = data_list[1].split('\t')
        key_count = len(set(key_row))
        # 按题目拆分行
        title_row = data_list[0].split('\t')
        title_row_split = split_list(title_row, key_count)
        value_row = data_list[2].split('\t')
        value_row_split = split_list(value_row, key_count)
        # 生成表格
        text = '\t'+'\t'.join(key_row[:key_count])+'\n'
        for titles, values in zip(title_row_split, value_row_split):
            _, title = titles[0].split('.')
            text += title+'\t'+'\t'.join(values)+'\n'
        output_text.insert('end', text[:-1])
    freeze()


# 窗口
root = ttk.Window(themename='cerculean', title='橙技')
root.geometry(f'1180x820')  # 窗口大小
root.iconbitmap(bitmap='green_apple.ico')
root.iconbitmap(default='green_apple.ico')
root.place_window_center()

label1 = ttk.Label(root, text='原始数据', font=('黑体', 12))
label1.pack(pady=(20, 10))  # 按布局方式放置标签

input_text = ttk.Text(root, height=12)
input_text.pack(fill='x', padx=100)  # 文本框宽度沿水平方向自适应填充，左右两边空100像素
input_text.focus()
input_text.bind("<Double-Button-1>", paste_from_clipboard)

label2 = ttk.Label(root, text='计算结果', font=('黑体', 12))
label2.pack(pady=(20, 10))

output_text = ttk.Text(root, height=12)
output_text.pack(fill='x', padx=100)
# 为文本框绑定鼠标双击事件
output_text.bind("<Double-Button-1>", copy_to_clipboard)

info_text = ttk.Text(root, height=5, font=('黑体', 12), spacing3=8, border=-1, state='disabled')
info_text.pack(pady=10, padx=100, fill='x')
info_text.tag_config('center', foreground='green', justify='center')

# 按钮区域
buttonbar = ttk.Labelframe(root, text='选择功能', labelanchor='n', padding=20)
buttonbar.pack(pady=10,  padx=100)

btn = ttk.Button(master=buttonbar, text='生成', command=table)
btn.pack(side='left', padx=10)

root.mainloop()
