import os
import pandas as pd
import numpy as np
import json
from collections import OrderedDict

def calculate_department_metrics():
    # 定义目录路径
    folder_path = r"C:\Users\biyun\Desktop\work\科室积分（2018.06-2024.10）\2024年绩效"
    # 定义输出JSON文件路径
    output_path = os.path.join(os.path.expanduser("~"), "Desktop", "医疗业务指标统计结果.json")
    no_data_output_path = os.path.join(os.path.expanduser("~"), "Desktop", "无数据科室统计.json")
    simple_output_path = os.path.join(os.path.expanduser("~"), "Desktop", "医疗业务指标简化统计.json")
    
    # 科室列表
    departments = [
        "胸外一科", "胸外二科", "介入治疗科", "妇瘤一科", "妇瘤二科",
        "骨与软组织一科", "骨与软组织二科", "乳腺一科", "乳腺二科", "头颈一科",
        "头颈二科", "胃肿瘤外科", "肝胆胰肿瘤外科", "泌尿外科", "结直肠肿瘤外科"
    ]
    
    # 需要统计的指标
    metrics = ["次均门诊费用", "次均出院费用", "住院西成药占比"]
    
    # 初始化结果字典，添加月份信息
    results = {dept: {metric: {} for metric in metrics} for dept in departments}
    
    print("\n=== 开始处理数据 ===")
    print(f"将处理以下科室的数据：{', '.join(departments)}")
    print(f"将统计以下指标：{', '.join(metrics)}\n")
    
    # 记录处理的文件数
    processed_files = []
    
    # 遍历目录下的所有Excel文件
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx') and '年' in filename and '月' in filename:
            file_path = os.path.join(folder_path, filename)
            print(f"\n正在处理文件：{filename}")
            processed_files.append(filename)
            
            try:
                # 判断是否为11月的数据
                is_november = '11月' in filename
                
                # 读取Excel时不指定header
                df = pd.read_excel(file_path, header=None)
                # 删除B列
                df = df.drop(df.columns[1], axis=1)
                
                # 获取第4行作为列名（索引为3）
                header_row = df.iloc[3]
                
                # 处理合并单元格：将每个非空列名向右填充3列
                new_headers = []
                current_header = None
                for h in header_row:
                    if pd.notna(h):
                        current_header = h
                    new_headers.append(current_header)
                
                # 设置列名
                df.columns = new_headers
                
                # 重命名第一列为'科室'
                df.columns.values[0] = '科室'
                
                # 从第6行开始读取数据（跳过前5行）
                df = df.iloc[5:]
                
                # 提取月份
                parts = filename.split('年')
                if len(parts) == 2:
                    month_part = parts[1].split('月')[0]
                    try:
                        month = str(int(month_part)).zfill(2)
                        print(f"提取到月份：{month}")
                    except ValueError:
                        print(f"警告：无法从文件名 {filename} 中提取月份")
                        continue
                else:
                    print(f"警告：文件名 {filename} 格式不正确")
                    continue
                
                # 对每个科室进行数据统计
                for dept in departments:
                    dept_data = df[df['科室'] == dept]
                    if not dept_data.empty:
                        print(f"\n科室：{dept}")
                        for metric in metrics:
                            if metric in df.columns:
                                try:
                                    # 获取指标列索引
                                    metric_idx = list(df.columns).index(metric)
                                    # 非11月数据需要向右偏移3列
                                    if not is_november:
                                        metric_idx += 3
                                    
                                    # 确保索引不超出列数
                                    if metric_idx < len(df.columns):
                                        value = dept_data.iloc[0, metric_idx]
                                        if pd.notna(value):  # 检查是否为空值
                                            value = float(value)
                                            results[dept][metric][month] = value
                                            print(f"  - {metric}: {value:.4f}")
                                        else:
                                            print(f"  - {metric}: 数据为空")
                                    else:
                                        print(f"  - {metric}: 列索引超出范围")
                                except Exception as e:
                                    print(f"  - {metric}: 处理出错 - {str(e)}")
                    else:
                        print(f"\n警告：未找到科室 {dept} 的数据")
                
            except Exception as e:
                print(f"\n错误：处理文件 {filename} 时出错")
                print(f"错误详情：{str(e)}")
                print(f"出错位置：", e.__traceback__.tb_lineno)
    
    # 生成最终JSON格式的报告
    final_report = OrderedDict()
    no_data_report = OrderedDict()
    simple_report = OrderedDict()  # 新增简化格式报告
    
    for dept in departments:
        dept_data = OrderedDict()
        no_data_metrics = OrderedDict()
        simple_dept_data = OrderedDict()  # 新增科室简化数据
        
        for metric in metrics:
            metric_data = OrderedDict()
            
            # 获取月度数据并排序
            monthly_data = results[dept][metric]
            sorted_months = sorted(monthly_data.keys())
            
            # 月度数据
            metric_data["monthly_data"] = OrderedDict()
            for month in sorted_months:
                metric_data["monthly_data"][f"{month}月"] = round(monthly_data[month], 4)
            
            # 计算统计数据
            values = list(monthly_data.values())
            if values:
                metric_data["statistics"] = {
                    "平均值": round(sum(values) / len(values), 4),
                    "数据月份数": len(values),
                    "总值": round(sum(values), 4)
                }
                # 添加到简化格式
                simple_dept_data[metric] = {
                    "平均值": round(sum(values) / len(values), 4),
                    "数据月份数": len(values)
                }
            else:
                # 记录无数据的指标
                no_data_metrics[metric] = {
                    "状态": "无数据",
                    "说明": "未找到任何月份的数据"
                }
            
            dept_data[metric] = metric_data
        
        final_report[dept] = dept_data
        simple_report[dept] = simple_dept_data  # 添加科室简化数据
        
        # 如果该科室有无数据的指标，添加到无数据报告中
        if no_data_metrics:
            no_data_report[dept] = no_data_metrics
    
    # 将结果保存为JSON文件
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(final_report, f, ensure_ascii=False, indent=2)
        print(f"\n结果已保存到：{output_path}")
        
        # 保存无数据科室的报告
        if no_data_report:
            with open(no_data_output_path, 'w', encoding='utf-8') as f:
                json.dump(no_data_report, f, ensure_ascii=False, indent=2)
            print(f"无数据科室统计已保存到：{no_data_output_path}")
        
        # 保存简化格式报告
        with open(simple_output_path, 'w', encoding='utf-8') as f:
            json.dump(simple_report, f, ensure_ascii=False, indent=2)
        print(f"简化统计结果已保存到：{simple_output_path}")
            
    except Exception as e:
        print(f"\n保存JSON文件时出错：{str(e)}")
    
    # 打印到控制台
    print("\n=== 统计结果 ===")
    print(json.dumps(final_report, ensure_ascii=False, indent=2))
    
    if no_data_report:
        print("\n=== 无数据科室统计 ===")
        print(json.dumps(no_data_report, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    calculate_department_metrics()