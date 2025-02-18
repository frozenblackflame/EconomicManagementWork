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

def get_latest_json_files():
    # 获取当前目录下所有JSON文件
    json_files = [f for f in os.listdir() if f.endswith('.json')]
    
    # 解析文件名中的日期
    dated_files = []
    for file in json_files:
        try:
            year, month = map(int, file.split('.')[0:2])
            date = datetime(year, month, 1)
            dated_files.append((date, file))
        except:
            continue
    
    # 按日期排序
    dated_files.sort(key=lambda x: x[0], reverse=True)
    
    return dated_files[:2] if len(dated_files) >= 2 else []

def get_data_from_performance_json(dept_name, data):
    """从绩效数据中提取指定科室的工作量数据"""
    dept_data = {}
    
    # 检查数据格式
    sample_item = data[0] if data else None
    if not sample_item:
        return dept_data
    
    # 定义需要查找的字段和对应的关键词
    field_keywords = {
        '出院人次（护理）': '出院人次',
        '年龄（护理）': '年龄',
        '护士中医适宜技术（护理）': '护士中医适宜技术',
        'Ⅱ级护理（护理）': 'Ⅱ级护理',
        'Ⅰ级护理（护理）': 'Ⅰ级护理',
        '出院患者占用床日（护理）': '出院患者占用床日'
    }
        
    # 如果是绩效数据格式（包含"项目名称"字段）
    if '项目名称' in sample_item:
        for item in data:
            if item['科室名称'] == dept_name:
                item_name = item['项目名称']
                # 对每个字段进行模糊匹配
                for field, keyword in field_keywords.items():
                    if keyword in item_name:
                        dept_data[field] = item['工作量']
                        break
    # 如果是基础数据格式（直接包含指标字段）
    else:
        dept_item = next((item for item in data if item['科室名称'] == dept_name), None)
        if dept_item:
            for field in field_keywords.keys():
                if field in dept_item:
                    dept_data[field] = dept_item[field]
                    
    return dept_data

def create_department_excel(dept_name, current_data, performance_data, latest_files):
    # 创建Excel文件
    output_dir = "护理工作量"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    filename = os.path.join(output_dir, f"{dept_name}.xlsx")
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    
    # 获取当前日期
    current_date = max(latest_files[0][0], latest_files[1][0])
    
    # 准备数据
    rows = []
    
    # 获取当前月数据
    current_dept_data = next((item for item in current_data if item['科室名称'] == dept_name), None)
    if current_dept_data:
        row = {
            '日期': current_date.strftime('%Y.%m'),
            '出院人次（护理）': current_dept_data.get('出院人次（护理）', ''),
            '年龄（护理）': current_dept_data.get('年龄（护理）', ''),
            '护士中医适宜技术（护理）': current_dept_data.get('护士中医适宜技术（护理）', ''),
            'Ⅱ级护理（护理）': current_dept_data.get('Ⅱ级护理（护理）', ''),
            'Ⅰ级护理（护理）': current_dept_data.get('Ⅰ级护理（护理）', ''),
            '出院患者占用床日（护理）': current_dept_data.get('出院患者占用床日（护理）', '')
        }
        rows.append(row)
    
    # 获取上月数据
    last_month = datetime(current_date.year, current_date.month - 1 if current_date.month > 1 else 12, 1)
    if current_date.month == 1:
        last_month = last_month.replace(year=last_month.year - 1)
    
    # 获取去年同月数据
    last_year = datetime(current_date.year - 1, current_date.month, 1)
    
    # 尝试读取去年同月的JSON文件
    last_year_filename = f"{last_year.year}.{last_year.month:02d}.json"
    if os.path.exists(last_year_filename):
        with open(last_year_filename, 'r', encoding='utf-8') as f:
            last_year_data = json.load(f)
            last_year_dept_data = get_data_from_performance_json(dept_name, last_year_data)
            if last_year_dept_data:
                row = {
                    '日期': last_year.strftime('%Y.%m'),
                    **{k: last_year_dept_data.get(k, '') for k in [
                        '出院人次（护理）', '年龄（护理）', '护士中医适宜技术（护理）',
                        'Ⅱ级护理（护理）', 'Ⅰ级护理（护理）', '出院患者占用床日（护理）'
                    ]}
                }
                rows.append(row)
    
    # 从performance_data中获取上月数据
    dept_data = get_data_from_performance_json(dept_name, performance_data)
    if dept_data:
        row = {
            '日期': last_month.strftime('%Y.%m'),
            **{k: dept_data.get(k, '') for k in [
                '出院人次（护理）', '年龄（护理）', '护士中医适宜技术（护理）',
                'Ⅱ级护理（护理）', 'Ⅰ级护理（护理）', '出院患者占用床日（护理）'
            ]}
        }
        rows.append(row)
    
    # 创建DataFrame并保存到Excel
    if rows:
        df = pd.DataFrame(rows)
        df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    writer.close()

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

    # 获取最新的两个JSON文件
    latest_files = get_latest_json_files()
    if len(latest_files) < 2:
        print("找不到足够的JSON文件")
        return
    
    # 读取当前月份数据
    with open(latest_files[0][1], 'r', encoding='utf-8') as f:
        current_data = json.load(f)
    
    # 读取上月绩效数据
    with open(latest_files[1][1], 'r', encoding='utf-8') as f:
        performance_data = json.load(f)
    
    # 获取所有科室名称
    departments = set(item['科室名称'] for item in current_data)
    
    # 为每个科室创建Excel文件
    for dept in departments:
        create_department_excel(dept, current_data, performance_data, latest_files)
        print(f"已生成科室 {dept} 的Excel文件")

if __name__ == "__main__":
    main()
