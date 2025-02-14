import json
import pandas as pd
from pathlib import Path
import os

def get_desktop_path():
    """获取桌面路径"""
    return os.path.join(os.path.expanduser("~"), "Desktop")

def read_json_data():
    """读取JSON文件"""
    desktop_path = get_desktop_path()
    json_path = os.path.join(desktop_path, "指标统计结果.json")
    
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def create_excel_data(data):
    """处理数据并创建Excel文件"""
    desktop_path = get_desktop_path()
    excel_path = os.path.join(desktop_path, "指标.xlsx")
    
    # 创建ExcelWriter对象
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        # 遍历每个科室
        for dept_name, dept_data in data.items():
            # 收集有效的指标数据
            valid_indicators = []
            
            # 遍历科室下的所有指标
            for indicator_name, indicator_data in dept_data.items():
                # 检查是否有月度数据
                if indicator_data.get('monthly_data') and len(indicator_data['monthly_data']) > 0:
                    monthly_data = indicator_data['monthly_data']
                    stats = indicator_data['statistics']
                    
                    # 创建一行数据
                    row_data = {
                        '序号': len(valid_indicators) + 1,
                        '考核指标': indicator_name,
                        '目标值': '',  # 目标值未提供
                        '1月': monthly_data.get('01月', ''),
                        '2月': monthly_data.get('02月', ''),
                        '3月': monthly_data.get('03月', ''),
                        '4月': monthly_data.get('04月', ''),
                        '5月': monthly_data.get('05月', ''),
                        '6月': monthly_data.get('06月', ''),
                        '7月': monthly_data.get('07月', ''),
                        '8月': monthly_data.get('08月', ''),
                        '9月': monthly_data.get('09月', ''),
                        '10月': monthly_data.get('10月', ''),
                        '11月': monthly_data.get('11月', ''),
                        '12月': monthly_data.get('12月', ''),
                        '全年均值': stats.get('平均值', '')
                    }
                    valid_indicators.append(row_data)
            
            if valid_indicators:
                # 创建DataFrame
                df = pd.DataFrame(valid_indicators)
                
                # 写入Excel，包括标题行
                worksheet = writer.sheets.get(dept_name)
                
                # 写入表头
                header = f"2024年{dept_name}开始指标统计表"
                df.to_excel(writer, sheet_name=dept_name, index=False, startrow=1)
                
                # 获取worksheet对象并写入标题
                worksheet = writer.sheets[dept_name]
                worksheet.cell(row=1, column=1, value=header)

def main():
    # 读取JSON数据
    data = read_json_data()
    
    # 创建Excel文件
    create_excel_data(data)
    
    print("Excel文件已生成到桌面")

if __name__ == "__main__":
    main()
