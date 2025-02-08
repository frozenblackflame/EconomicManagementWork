import json
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def json_to_excel():
    # 创建主窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    # 打开文件选择对话框
    json_file = filedialog.askopenfilename(
        title='选择JSON文件',
        filetypes=[('JSON files', '*.json')]
    )
    
    if not json_file:  # 如果用户没有选择文件
        print("未选择文件")
        return
    
    try:
        # 读取JSON文件
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 准备Excel数据
        excel_data = []
        
        # 处理不同的JSON格式
        if isinstance(data, dict):
            # 如果是简单的月份-数值对应格式
            if all(isinstance(k, str) and "月" in k for k in data.keys()):
                row = {}
                for month, value in data.items():
                    row[month] = value
                excel_data.append(row)
            else:
                # 原有的科室-指标格式处理逻辑
                for dept, metrics in data.items():
                    if not isinstance(metrics, dict):
                        print(f"警告：科室 '{dept}' 的数据格式不正确，已跳过")
                        continue
                    
                    for metric_name, metric_data in metrics.items():
                        try:
                            if isinstance(metric_data, dict) and "平均值" in metric_data:
                                row = {
                                    "科室": dept,
                                    "指标": metric_name,
                                    "平均值": metric_data["平均值"],
                                    "数据月份数": metric_data["数据月份数"]
                                }
                                excel_data.append(row)
                            elif isinstance(metric_data, dict) and "monthly_data" in metric_data:
                                stats = metric_data.get("statistics", {})
                                monthly = metric_data.get("monthly_data", {})
                                
                                row = {
                                    "科室": dept,
                                    "指标": metric_name,
                                    "平均值": stats.get("平均值", ""),
                                    "数据月份数": stats.get("数据月份数", ""),
                                    "总值": stats.get("总值", "")
                                }
                                
                                for month, value in monthly.items():
                                    row[month] = value
                                
                                excel_data.append(row)
                            else:
                                print(f"警告：指标 '{metric_name}' 的数据格式不正确，已跳过")
                        except Exception as e:
                            print(f"处理指标 '{metric_name}' 时出错：{str(e)}")
                            continue
        else:
            raise ValueError("JSON数据格式错误：需要是一个字典（对象）格式")
            
        if not excel_data:
            raise ValueError("没有有效的数据可以转换")

        # 创建DataFrame
        df = pd.DataFrame(excel_data)
        
        # 生成输出文件名
        json_filename = os.path.splitext(os.path.basename(json_file))[0]
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        excel_file = os.path.join(desktop_path, f"{json_filename}.xlsx")
        
        # 保存为Excel文件
        df.to_excel(excel_file, index=False)
        print(f"转换完成！文件已保存至：{excel_file}")
        
        # 显示成功消息框
        root = tk.Tk()
        root.withdraw()
        tk.messagebox.showinfo("转换成功", f"文件已保存至：\n{excel_file}")
        
    except Exception as e:
        # 显示错误消息框
        root = tk.Tk()
        root.withdraw()
        tk.messagebox.showerror("错误", f"转换过程中出现错误：\n{str(e)}")
        print(f"错误：{str(e)}")

if __name__ == "__main__":
    json_to_excel()