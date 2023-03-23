import os
from PIL import Image
from concurrent.futures import ThreadPoolExecutor
from tkinter import filedialog

#  制作可视化界面


def splice(filesname, output_path, quality, del_name=''):
    img1 = Image.open(input_path+'/'+filesname[0])
    img2 = Image.open(input_path+'/'+filesname[1])
    img1_height = img1.height
    new_width = max(img1.width, img2.width)
    new_height = img1.height + img2.height
    result = Image.new(mode='RGB', size=(new_width, new_height), color=(255, 255, 255))
    result.paste(img1, box=(0, 0))
    result.paste(img2, box=(0, img1_height))
    save_name = filesname[0].split('.')[0]
    save_name = save_name.replace(del_name, '')
    save_path = output_path + '/' + save_name
    result.save(save_path, format='JPEG', optimize=True, quality=quality)
    os.rename(save_path, save_path+'.png')


input_path = filedialog.askdirectory(title='请选择图片文件夹', initialdir='F:/用户目录/桌面/')
img_list = os.listdir(input_path)
splice_file_list = []
for i in range(0, len(img_list), 2):
    splice_file_list.append((img_list[i], img_list[i+1]))

output_path = filedialog.askdirectory(title='请选择合并文件夹', initialdir='F:/用户目录/桌面/')

with ThreadPoolExecutor(max_workers=32) as pool:
    for filesname in splice_file_list:
        pool.submit(splice, filesname=filesname, output_path=output_path, quality=33, del_name='_full_1')  # 提交任务

