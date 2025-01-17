import os
from datetime import datetime
import fitz  # PyMuPDF
from PIL import Image

def pdf_to_jpg(pdf_path, output_folder="image"):
    """
    将PDF文件转换为JPG图像并保存到指定文件夹。

    参数:
    pdf_path (str): PDF文件的路径。
    output_folder (str): 保存图像的文件夹路径，默认为"image"。

    返回:
    None
    """
    # 确保输出文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 打开PDF文件
    pdf_document = fitz.open(pdf_path)
    for page_num in range(len(pdf_document)):
        # 获取页面
        page = pdf_document.load_page(page_num)
        # 将页面转换为像素地图
        pix = page.get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72))
        # 创建图像对象
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        # 生成日期
        date = datetime.now().strftime("%H-%M-%S-%f")
        # 保存图像
        img.save(os.path.join(output_folder, f"{date}_page_{page_num + 1}.jpg"))

# 测试代码（仅在直接运行此文件时执行）
if __name__ == "__main__":
    pdf_path = "path/to/your/pdf.pdf"  # 替换为你的PDF文件路径
    pdf_to_jpg(pdf_path, "output_images")  # 替换为你想要的输出文件夹路径
