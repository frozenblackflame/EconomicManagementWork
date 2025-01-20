import json
from rapidocr_onnxruntime import RapidOCR
import os
from UniversalToolbox import pdf_to_jpg
import tkinter as tk
from tkinter import filedialog

def ocr_text(img_path):
    engine = RapidOCR()
    result, elapse = engine(img_path)
    extracted_texts = [item[1] for item in result]
    print(f"识别结果: {extracted_texts}")
    return extracted_texts

def main():
    # 创建主窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    pdf_path = filedialog.askdirectory(title='选择pdf文件夹')
    pdf_list = os.listdir(pdf_path)
    for pdf in pdf_list:
        pdf_to_jpg(os.path.join(pdf_path, pdf))
    #获取返回的text列表
    text_list = []
    #循环image文件夹下的jpg文件,传入完整路径
    for file in os.listdir("image"):
        if file.endswith(".jpg"):
            text_list.append(ocr_text(os.path.join("image", file)))
    #将text列表写入text.json文件并保存到桌面
    with open(os.path.join(os.path.expanduser("~"), "Desktop", "text.json"), "w", encoding="utf-8") as file:
        json.dump(text_list, file, ensure_ascii=False, indent=4)

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