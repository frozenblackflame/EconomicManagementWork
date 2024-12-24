import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
from datetime import datetime
import os
import re

class PerformanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("绩效结果生成")
        
        self.current_month_data = {}
        self.last_month_data = {}
        self.last_year_data = {}
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建输入框框架
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=5)
        
        # 输入框标签
        ttk.Label(input_frame, text="上年日期格式（2024-12）：").pack(side=tk.LEFT)
        
        # 设置默认日期为上年同月
        current_date = datetime.now()
        last_year_date = (current_date.replace(year=current_date.year-1)).strftime("%Y-%m")
        
        # 输入框
        self.date_entry = ttk.Entry(input_frame)
        self.date_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.date_entry.insert(0, last_year_date)
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        # 按钮
        self.current_month_button = ttk.Button(button_frame, text="选择当前月文件", command=self.load_current_month_file)
        self.current_month_button.pack(fill=tk.X, pady=2)
        
        self.last_month_button = ttk.Button(button_frame, text="选择上个月文件", command=self.load_last_month_file)
        self.last_month_button.pack(fill=tk.X, pady=2)
        
        self.last_year_button = ttk.Button(button_frame, text="选择上年文件", command=self.load_last_year_file)
        self.last_year_button.pack(fill=tk.X, pady=2)
        
        self.parse_button = ttk.Button(button_frame, text="开始解析", command=self.parse_data)
        self.parse_button.pack(fill=tk.X, pady=2)

    def load_current_month_file(self):
        file_path = filedialog.askopenfilename()
        self.current_month_data = self.read_file(file_path)

    def load_last_month_file(self):
        file_path = filedialog.askopenfilename()
        self.last_month_data = self.read_file(file_path)

    def load_last_year_file(self):
        file_path = filedialog.askopenfilename()
        self.last_year_data = self.read_last_year_file(file_path)

    def read_file(self, file_path):
        # 读取当前月和上月的文件
        df = pd.read_excel(file_path, header=[0, 1])  # 读取合并单元格的列名
        data_dict = {}
        for index, row in df.iterrows():
            department = row.iloc[0]  # 获取科室名称（A列）
            if department not in data_dict:
                data_dict[department] = []
            
            # 获取项目名称和工作量
            for col in df.columns[1:]:  # 跳过第一列（科室名称）
                project_name = col[0]  # 项目名称在第一层
                workload = row[col]    # 工作量
                data_dict[department].append({
                    "项目名称": project_name,
                    "工作量": workload
                })
        return data_dict

    def read_last_year_file(self, file_path):
        # 获取输入框中的日期
        date_str = self.date_entry.get()
        sheet_name = f"{date_str[:4]}年{date_str[5:7]}月绩效-核算数据"
        
        # 读取指定工作表
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        data_dict = {}
        
        # 获取列名（第一行）
        columns = df.columns.tolist()
        
        
        for index, row in df.iterrows():
            # 用于记录优势病种的计数
            advantage_disease_count = 1
            department = row[columns[0]]  # 科室名称
            if department not in data_dict:
                data_dict[department] = []
            
            # 获取项目名称和工作量
            project_name = row[columns[1]]  # 项目名称
            workload = row[columns[3]]      # 工作量（第4列）
            
            # 处理项目名称
            if isinstance(project_name, str):  # 确保项目名称是字符串
                # 替换罗马数字
                project_name = project_name.replace("Ⅰ级", "Ⅰ级").replace("Ⅱ级", "Ⅱ级")
                
                # 使用正则表达式匹配单个大写字母开头且（护理）结尾的项目
                if re.match(r'^[A-Z].*（护理）$', project_name):
                    project_name = f"优势病种{advantage_disease_count}"
                    advantage_disease_count += 1
                # 为其他项目添加（护理）后缀
                elif not project_name.endswith('（护理）'):
                    project_name = f"{project_name}（护理）"
            
            data_dict[department].append({
                "项目名称": project_name,
                "工作量": workload
            })
        
        return data_dict

    def parse_data(self):
        try:
            # 创建保存文件的文件夹
            output_dir = "护理科室数据"
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 获取日期
            last_year_date = self.date_entry.get()
            last_month_date = (datetime.strptime(last_year_date, "%Y-%m") - pd.DateOffset(months=1)).strftime("%Y-%m")
            current_date = datetime.now().strftime("%Y-%m")
            
            # 遍历当前月的所有科室
            for department in self.current_month_data:
                try:
                    # 使用当前月的项目名称作为列名
                    current_month_projects = []
                    for item in self.current_month_data[department]:
                        project_name = item["项目名称"]
                        if project_name not in current_month_projects:
                            current_month_projects.append(project_name)
                    
                    # 按数字排序项目名称
                    current_month_projects.sort(key=lambda x: int(x.split('.')[0]) if x.split('.')[0].isdigit() else float('inf'))
                    
                    # 创建DataFrame数据
                    data = {
                        "日期": [last_year_date, last_month_date, current_date]
                    }
                    
                    # 使用当前月的项目作为列名，并填充数据
                    for project in current_month_projects:
                        data[project] = []
                        
                        # 上年数据
                        value = 0
                        if department in self.last_year_data:
                            for item in self.last_year_data[department]:
                                if item["项目名称"] == project:
                                    value = item["工作量"]
                                    break
                        data[project].append(value)
                        
                        # 上月数据
                        value = 0
                        if department in self.last_month_data:
                            for item in self.last_month_data[department]:
                                if item["项目名称"] == project:
                                    value = item["工作量"]
                                    break
                        data[project].append(value)
                        
                        # 当月数据
                        value = 0
                        for item in self.current_month_data[department]:
                            if item["项目名称"] == project:
                                value = item["工作量"]
                                break
                        data[project].append(value)
                    
                    # 创建DataFrame并保存为Excel文件
                    df = pd.DataFrame(data)
                    output_path = os.path.join(output_dir, f"{department}.xlsx")
                    
                    # 如果文件已存在，先尝试删除
                    if os.path.exists(output_path):
                        try:
                            os.remove(output_path)
                        except:
                            import time
                            # 如果删除失败，使用时间戳创建新文件名
                            output_path = os.path.join(output_dir, f"{department}_{int(time.time())}.xlsx")
                    
                    df.to_excel(output_path, index=False)
                    
                except Exception as e:
                    print(f"处理科室 {department} 时出错: {str(e)}")
                    continue
                
        except Exception as e:
            print(f"程序执行出错: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("400x300")  # 设置窗口大小
    app = PerformanceApp(root)
    root.mainloop()