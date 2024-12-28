import pandas as pd
import os
import re
from datetime import datetime

def extract_date(filename):
    # 从文件名中提取年月信息
    match = re.search(r'(\d{4})年(\d{1,2})月', filename)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        return datetime(year, month, 1)
    return None

def process_excel_files(folder_path):
    # 存储所有月份的数据
    all_data = {}
    
    # 遍历文件夹中的所有Excel文件
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            date = extract_date(filename)
            
            if date:
                try:
                    # 读取Excel文件的第16行，设置header=None避免将数据行作为列名
                    df = pd.read_excel(file_path, skiprows=15, nrows=1, header=None)
                    
                    # 打印列名和数据，用于调试
                    print(f"\n处理文件：{filename}")
                    print(f"B列(总数)：{df.iloc[0, 1]}")
                    print(f"H列(出院人次)：{df.iloc[0, 7-2]}")
                    print(f"N列(手术)：{df.iloc[0, 13-2]}")
                    print(f"Q列(手术)：{df.iloc[0, 16-2]}")
                    print(f"T列(手术)：{df.iloc[0, 19-2]}")
                    print(f"W列(手术)：{df.iloc[0, 22-2]}")
                    print(f"FK列(手术)：{df.iloc[0, 166-2]}")
                    print(f"FN列(手术)：{df.iloc[0, 169-2]}")
                    print(f"手术合计：{df.iloc[0, [13-2, 16-2, 19-2, 22-2, 166-2, 169-2]].sum()}")
                    print(f"Z列(门诊人次)：{df.iloc[0, 25-2]}")
                    
                    # 使用数字索引来获取列数据
                    surgery = df.iloc[0, [13-2, 16-2, 19-2, 22-2, 166-2, 169-2]].sum()  # F=13-2, Q=16-2, T=19-2, W=22-2, FK=166-2, FN=169-2
                    outpatient = df.iloc[0, 25-2]  # Z列
                    discharge = df.iloc[0, 7-2]    # H列
                    total = df.iloc[0, 1]       # B列
                    
                    # 计算比例
                    surgery_ratio = surgery / total
                    discharge_ratio = discharge / total
                    outpatient_ratio = outpatient / total
                    
                    # 存储数据
                    month_data = pd.DataFrame({
                        '实际数': [surgery, discharge, outpatient, total],
                        '比例': [surgery_ratio, discharge_ratio, outpatient_ratio, 1]
                    }, index=['手术', '出院人次', '门诊人次', '总数'])
                    
                    all_data[date] = month_data
                    
                except Exception as e:
                    print(f"处理文件 {filename} 时出错：{str(e)}")

    # 修改输出路径到桌面
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    output_file = os.path.join(desktop_path, '统计结果.xlsx')
    
    # 创建新的Excel文件，按月份排序
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for date in sorted(all_data.keys()):
            sheet_name = f"{date.year}年{date.month}月"
            all_data[date].to_excel(writer, sheet_name=sheet_name)

# 使用示例
folder_path = r"C:\Users\biyun\Desktop\临床积分明细（2018.06-2024.10）\2024年绩效"
process_excel_files(folder_path)