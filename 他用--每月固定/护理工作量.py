import os
import json
import pandas as pd
from datetime import datetime
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilenames

def get_date_from_filename(filename):
    # 从文件名中提取日期
    match = re.search(r'(\d{4})\.(\d{2})', filename)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        return datetime(year, month, 1)
    return None

def get_previous_year_date(current_date):
    # 获取上一年同月的日期
    return datetime(current_date.year - 1, current_date.month, 1)

def format_sheet_name(date):
    # 格式化工作表名称
    return f"{date.year}年{date.month}月绩效-核算数据"

def save_to_json(data, filename):
    # 保存数据到JSON文件
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def process_dataframe(df):
    # 重命名 'Unnamed: 0' 列为 '科室名称'
    if 'Unnamed: 0' in df.columns:
        df = df.rename(columns={'Unnamed: 0': '科室名称'})
    
    # 如果存在科室名称列，删除科室名称为NaN的行
    if '科室名称' in df.columns:
        df = df.dropna(subset=['科室名称'])
    
    return df

def main():
    # 创建Tk根窗口
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    
    # 打开文件选择对话框
    files = askopenfilenames(
        title="选择Excel文件",
        filetypes=[("Excel files", "*.xlsx *.xlsm")]
    )
    
    if not files:
        print("未选择文件")
        return
        
    # 分离护理绩效数据文件和工作量文件
    nursing_performance_file = None
    workload_files = []
    
    for file in files:
        filename = os.path.basename(file)
        if filename == "护理绩效数据.xlsx":
            nursing_performance_file = file
        else:
            workload_files.append(file)
    
    if not nursing_performance_file or not workload_files:
        print("缺少必要的文件")
        return
    
    # 找到最近日期的工作量文件
    current_dates = [get_date_from_filename(os.path.basename(f)) for f in workload_files]
    current_dates = [d for d in current_dates if d is not None]
    
    if not current_dates:
        print("无法从文件名中提取日期")
        return
    
    current_date = max(current_dates)
    previous_date = get_previous_year_date(current_date)
    
    # 处理工作量文件
    for file in workload_files:
        filename = os.path.basename(file)
        file_date = get_date_from_filename(filename)
        
        if file_date:
            df = pd.read_excel(file)
            df = process_dataframe(df)  # 处理数据框
            output_filename = f"{file_date.year}.{file_date.month:02d}.json"
            save_to_json(df.to_dict(orient='records'), output_filename)
            print(f"已保存工作量数据到: {output_filename}")
    
    # 处理护理绩效数据文件
    sheet_name = format_sheet_name(previous_date)
    try:
        df = pd.read_excel(nursing_performance_file, sheet_name=sheet_name)
        df = process_dataframe(df)  # 处理数据框
        output_filename = f"{previous_date.year}.{previous_date.month:02d}.json"
        save_to_json(df.to_dict(orient='records'), output_filename)
        print(f"已保存护理绩效数据到: {output_filename}")
    except Exception as e:
        print(f"处理护理绩效数据时出错: {str(e)}")

if __name__ == "__main__":
    main()
