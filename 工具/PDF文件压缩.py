import os
from pdf2image import convert_from_path
from PIL import Image
from PyPDF2 import PdfMerger
import tempfile

def compress_pdf(input_path, output_path, target_size_mb=2):
    with tempfile.TemporaryDirectory() as temp_dir:
        # 降低DPI至120
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
    input_file = r"C:\Users\biyun\Desktop\工作\1.pdf"
    output_file = 'compressed_output.pdf'
    compress_pdf(input_file, output_file)