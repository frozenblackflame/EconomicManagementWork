import os
import pandas as pd

def convert_xls_to_xlsx(directory):
    # 切换到指定目录
    os.chdir(directory)
    
    # 获取所有的 .xls 文件
    xls_files = [f for f in os.listdir() if f.endswith('.xls')]
    
    for xls_file in xls_files:
        # 生成对应的 .xlsx 文件名
        xlsx_file = xls_file.replace('.xls', '.xlsx')
        
        # 读取 .xls 文件并保存为 .xlsx 文件
        try:
            df = pd.read_excel(xls_file)
            df.to_excel(xlsx_file, index=False)
            print(f'已转换: {xls_file} -> {xlsx_file}')
            
            # 删除原本的 .xls 文件
            os.remove(xls_file)
            print(f'已删除原文件: {xls_file}')
        except Exception as e:
            print(f'转换 {xls_file} 时出错: {e}')

if __name__ == "__main__":
    # 手动输入目录
    directory = input("请输入包含 .xls 文件的目录: ")
    convert_xls_to_xlsx(directory)
