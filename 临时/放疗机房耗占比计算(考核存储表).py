import pandas as pd
import os
import json
from datetime import datetime

def calculate_monthly_ratio():
    # 设置文件路径
    base_path = r"C:\Users\biyun\Desktop\工作\存储表"
    
    # 存储每月的耗占比结果
    monthly_ratios = {}
    
    # 遍历目录下的所有文件
    for filename in os.listdir(base_path):
        if filename.startswith("考核数据存储") and filename.endswith(".xlsx"):
            try:
                # 从文件名中提取年月
                date_str = filename[6:12]  # 提取文件名中的年月部分
                file_date = datetime.strptime(date_str, "%Y%m")
                
                # 只处理2024年的文件
                if file_date.year == 2024:
                    file_path = os.path.join(base_path, filename)
                    print(f"正在处理文件: {filename}")
                    
                    # 读取Excel文件，跳过第一行，使用第二行作为表头
                    df = pd.read_excel(file_path, usecols='A:Z', header=1)
                    
                    # 找到"放疗机房"所在的行
                    radiotherapy_row = df[df['科室名称'] == "放疗机房"]
                    
                    if radiotherapy_row.empty:
                        print(f"警告：在文件 {filename} 的科室列中未找到'放疗机房'")
                        continue
                    
                    # 检查所需列是否存在
                    required_columns = ["耗材", "门诊执行收入", "住院执行收入"]
                    for col in required_columns:
                        if col not in df.columns:
                            raise ValueError(f"未找到所需列：{col}")
                    
                    # 获取需要的数据
                    haocai = float(radiotherapy_row['耗材'].iloc[0])
                    menzhen = float(radiotherapy_row['门诊执行收入'].iloc[0])
                    zhuyuan = float(radiotherapy_row['住院执行收入'].iloc[0])
                    
                    print(f"获取到的数据 - 耗材: {haocai}, 门诊: {menzhen}, 住院: {zhuyuan}")
                    
                    # 计算耗占比
                    ratio = haocai / (menzhen + zhuyuan) if (menzhen + zhuyuan) != 0 else 0
                    
                    # 存储结果
                    month_key = f"{file_date.year}年{file_date.month}月"
                    # 保存为带百分号的字符串，保留2位小数
                    monthly_ratios[month_key] = f"{round(ratio * 100, 2)}%"
                    print(f"{month_key}的耗占比为：{monthly_ratios[month_key]}")
                    
            except Exception as e:
                print(f"处理文件 {filename} 时出错: {str(e)}")
                print(f"错误类型: {type(e).__name__}")
                import traceback
                print(traceback.format_exc())
    
    # 将结果写入JSON文件到桌面
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    output_path = os.path.join(desktop_path, "放疗机房耗占比.json")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(monthly_ratios, f, ensure_ascii=False, indent=4)
    
    print(f"计算完成，结果已保存到: {output_path}")

if __name__ == "__main__":
    calculate_monthly_ratio()
