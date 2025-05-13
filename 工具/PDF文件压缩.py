import os
from pdf2image import convert_from_path
from PIL import Image
from PyPDF2 import PdfMerger
import tempfile
from tkinter import filedialog
import tkinter as tk
from pathlib import Path

def compress_pdf(input_path, output_path, target_size_mb=2):
    with tempfile.TemporaryDirectory() as temp_dir:
        # 降低DPI至120（这个参数可以根据需要调整）
        images = convert_from_path(input_path, dpi=120)
        
        compressed_images = []
        quality = 10  # 降低初始质量（这个参数可以根据需要调整）
        for i, image in enumerate(images):
            img_path = os.path.join(temp_dir, f'page_{i}.jpg')
            # 降低图片最大尺寸至1200像素
            width, height = image.size
            if width > 1200 or height > 1200:
                ratio = min(1200/width, 1200/height)
                new_size = (int(width * ratio), int(height * ratio))
                image = image.resize(new_size, Image.LANCZOS)
            
            while True:
                image.save(img_path, 'JPEG', quality=quality, optimize=True)
                if os.path.getsize(img_path) > (target_size_mb * 1024 * 1024) / len(images):
                    quality -= 15  # 更激进的质量降低步长
                    if quality < 10:  # 允许更低的质量下限
                        # 如果质量已经很低仍然太大，尝试进一步缩小尺寸
                        width, height = image.size
                        new_size = (int(width * 0.8), int(height * 0.8))
                        image = image.resize(new_size, Image.LANCZOS)
                        quality = 20  # 重置质量继续尝试
                        if width < 800:  # 防止图片过小
                            break
                else:
                    break
            compressed_images.append(img_path)

        # 将压缩后的图片转换回PDF
        merger = PdfMerger()
        for img_path in compressed_images:
            img = Image.open(img_path)
            pdf_path = img_path.replace('.jpg', '.pdf')
            img.save(pdf_path, 'PDF', resolution=72.0)
            merger.append(pdf_path)

        merger.write(output_path)
        merger.close()

if __name__ == '__main__':
    # 创建GUI根窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    # 打开文件选择对话框
    input_file = filedialog.askopenfilename(
        title='选择要压缩的PDF文件',
        filetypes=[('PDF文件', '*.pdf')]
    )
    
    if input_file:
        # 创建桌面上的"压缩后"文件夹
        desktop_path = str(Path.home() / "Desktop")
        output_folder = os.path.join(desktop_path, "压缩后")
        os.makedirs(output_folder, exist_ok=True)
        
        # 设置输出文件路径
        input_filename = os.path.basename(input_file)
        output_filename = f"压缩_{input_filename}"
        output_file = os.path.join(output_folder, output_filename)
        
        # 压缩PDF
        compress_pdf(input_file, output_file)
        print(f"文件已保存至: {output_file}")