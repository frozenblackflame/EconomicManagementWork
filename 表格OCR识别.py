from datetime import datetime
import os

from lineless_table_rec import LinelessTableRecognition
from lineless_table_rec.utils_table_recover import format_html, plot_rec_box_with_logic_info, plot_rec_box
from table_cls import TableCls
from wired_table_rec import WiredTableRecognition

import fitz  # PyMuPDF
from PIL import Image
import json
import pandas as pd
from bs4 import BeautifulSoup

def pdf_to_jpg(pdf_path, output_folder="image"):
    # 确保输出文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 打开PDF文件
    pdf_document = fitz.open(pdf_path)
    for page_num in range(len(pdf_document)):
        # 获取页面
        page = pdf_document.load_page(page_num)
        # 将页面转换为像素地图
        pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
        # 创建图像对象
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        #当前日期
        date = datetime.now().strftime("%H-%M-%S-%f")
        # 保存图像
        img.save(os.path.join(output_folder, f"{date}_page_{page_num + 1}.jpg"))

def jpg_to_html(img_path):
    lineless_engine = LinelessTableRecognition()
    wired_engine = WiredTableRecognition()
    # 默认小yolo模型(0.1s)，可切换为精度更高yolox(0.25s),更快的qanything(0.07s)模型
    table_cls = TableCls(model_type="yolo") # TableCls(model_type="yolox"),TableCls(model_type="q")

    cls,elasp = table_cls(img_path)
    if cls == 'wired':
        table_engine = wired_engine
    else:
        table_engine = lineless_engine
    
    html, elasp, polygons, logic_points, ocr_res = table_engine(img_path)
    print(f"elasp: {elasp}")
    return html

def main():
    pdf_path = input("请输入pdf文件夹路径: ").replace("\"", "")
    pdf_list = os.listdir(pdf_path)
    for pdf in pdf_list:
        pdf_to_jpg(os.path.join(pdf_path, pdf))
    #获取返回的html列表
    html_list = []
    #循环image文件夹下的jpg文件,传入完整路径
    for file in os.listdir("image"):
        if file.endswith(".jpg"):
            html_list.append(jpg_to_html(os.path.join("image", file)))
    #将html列表写入html.json文件并保存到桌面
    with open(os.path.join(os.path.expanduser("~"), "Desktop", "html.json"), "w", encoding="utf-8") as file:
        json.dump(html_list, file, ensure_ascii=False, indent=4)

    #强制删除image文件夹包括文件夹下的所有文件
    for root, dirs, files in os.walk("image"):
        for file in files:
            os.remove(os.path.join(root, file))
        for dir in dirs:
            os.rmdir(os.path.join(root, dir))
    os.rmdir("image")
    print("文件处理完成")

if __name__ == "__main__":
    main()