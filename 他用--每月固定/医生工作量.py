import os
import pandas as pd
from tkinter import filedialog
import tkinter as tk
import zipfile
from datetime import datetime

def process_excel_files():
    # 创建一个临时的root窗口（但不显示）
    root = tk.Tk()
    root.withdraw()
    
    # 弹出文件选择对话框，允许多选
    files = filedialog.askopenfilenames(
        title='请选择Excel文件',
        filetypes=[('Excel文件', '*.xlsx')]
    )
    
    if not files:  # 如果用户取消选择，则退出
        print("未选择文件，程序退出")
        return
    
    # 获取桌面路径
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    output_folder = os.path.join(desktop_path, '医生的工作量')
    
    # 如果输出文件夹不存在，则创建
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # 遍历选择的文件
    for file_path in files:
        file_name = os.path.basename(file_path)
        date_info = file_name.split('工作量.xlsx')[0]  # 提取日期信息
        
        # 读取Excel文件，不直接指定列，因为需要动态查找
        df = pd.read_excel(file_path, header=None, skiprows=3)

        # 提取第4行和第5行的内容
        header_row_1 = df.iloc[0].tolist()  # 第4行
        header_row_2 = df.iloc[1].tolist()  # 第5行
        
        # 找到需要的列索引
        indices = {
            '出院人次': header_row_1.index('出院人次'),
            '门诊人次': header_row_1.index('门诊人次'),
            '1级手术': header_row_1.index('1级手术'),
            '2级手术': header_row_1.index('2级手术'),
            '3级手术': header_row_1.index('3级手术'),
            '4级手术': header_row_1.index('4级手术'),
            '3级微创手术': header_row_1.index('3级微创手术'),
            '4级微创手术': header_row_1.index('4级微创手术'),
            '医生中医适宜技术': header_row_1.index('医生中医适宜技术')
        }
        
        # 收集每个科室的数据
        for index, row in df.iterrows():
            if index >= 2:  # 从第6行开始读取
                department_name = row[0]  # 科室名称
                if pd.notna(department_name):  # 确保科室名称不是NaN
                    data = {
                        '日期': date_info,
                        '出院人次': row[indices['出院人次']+2],
                        '门诊人次': row[indices['门诊人次']+2],
                        '1级手术': row[indices['1级手术']+2],
                        '2级手术': row[indices['2级手术']+2],
                        '3级手术': row[indices['3级手术']+2],
                        '4级手术': row[indices['4级手术']+2],
                        '3级微创手术': row[indices['3级微创手术']+2],
                        '4级微创手术': row[indices['4级微创手术']+2],
                        '医生中医适宜技术': row[indices['医生中医适宜技术']+2]
                    }
                    
                    # 确保数值类型一致，转换为数值类型
                    try:
                        微创3级 = float(data['3级微创手术']) if pd.notna(data['3级微创手术']) else 0
                        微创4级 = float(data['4级微创手术']) if pd.notna(data['4级微创手术']) else 0
                        data['微创手术'] = 微创3级 + 微创4级
                        data['医生中医适宜技术'] = float(data['医生中医适宜技术']) if pd.notna(data['医生中医适宜技术']) else 0
                        # 转换手术数据为数值类型
                        for level in ['1级手术', '2级手术', '3级手术', '4级手术']:
                            data[level] = float(data[level]) if pd.notna(data[level]) else 0
                    except (ValueError, TypeError):
                        if '微创手术' not in data:
                            data['微创手术'] = 0
                        if '医生中医适宜技术' not in data:
                            data['医生中医适宜技术'] = 0
                        for level in ['1级手术', '2级手术', '3级手术', '4级手术']:
                            if level not in data:
                                data[level] = 0

                    # 去除data['3级微创手术'] 和 data['4级微创手术']
                    del data['3级微创手术']
                    del data['4级微创手术']
                    
                    # 将数据转化为DataFrame
                    department_df = pd.DataFrame([data])
                    
                    # 保存到新的Excel文件中
                    output_file_path = os.path.join(output_folder, f'{department_name}.xlsx')
                    if os.path.exists(output_file_path):
                        # 读取现有的数据
                        existing_df = pd.read_excel(output_file_path, sheet_name='工作量明细')
                        # 追加新的数据
                        final_df = pd.concat([existing_df, department_df], ignore_index=True)
                    else:
                        final_df = department_df
                    
                    # 保存到新的Excel文件中
                    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                        final_df.to_excel(writer, sheet_name='工作量明细', index=False)
                        # 设置列宽
                        worksheet = writer.sheets['工作量明细']
                        for column in worksheet.columns:
                            worksheet.column_dimensions[column[0].column_letter].width = 10

    # 创建ZIP文件
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    zip_filename = os.path.join(desktop_path, f'医生的工作量_{timestamp}.zip')
    
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # 遍历输出文件夹中的所有文件
        for foldername, subfolders, filenames in os.walk(output_folder):
            for filename in filenames:
                # 获取文件的完整路径
                file_path = os.path.join(foldername, filename)
                # 获取相对于output_folder的路径
                arcname = os.path.relpath(file_path, output_folder)
                # 将文件添加到ZIP中
                zipf.write(file_path, arcname)
    
    print(f"处理完成！文件已保存到：{zip_filename}")

# 直接调用函数
if __name__ == "__main__":
    process_excel_files()
