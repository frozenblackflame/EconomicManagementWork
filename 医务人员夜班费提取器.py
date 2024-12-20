import os
import tkinter as tk
from collections import defaultdict
from tkinter import filedialog, messagebox

import pandas as pd


class ExcelProcessor:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("夜班费统计程序")
        self.window.geometry("300x150")
        self.setup_gui()

    def setup_gui(self):
        self.select_button = tk.Button(
            self.window,
            text="选择Excel文件目录",
            command=self.process_directory
        )
        self.select_button.pack(pady=20)

    def process_directory(self):
        directory = filedialog.askdirectory()
        if not directory:
            return

        try:
            excel_files = [f for f in os.listdir(directory)
                           if f.endswith(('.xlsx', '.xls'))]

            all_department_data = defaultdict(dict)  # 存储所有科室的数据

            for file in excel_files:
                try:
                    file_path = os.path.join(directory, file)
                    # 跳过隐藏文件和临时文件
                    if file.startswith('~$') or file.startswith('.'):
                        continue

                    df = pd.read_excel(file_path, header=None)
                    if df.empty:
                        continue

                    # 全表检索单元格值为“部门”（去空格）的单元格，并读取所在行的右边单元格的值
                    department = None
                    for row_idx, row in df.iterrows():
                        for col_idx, cell in enumerate(row):
                            if str(cell).strip().lower() == '部门':
                                department_col = col_idx + 1
                                if department_col < len(df.columns):
                                    department = str(df.iat[row_idx, department_col]).replace('\\n', '').strip()
                                    break
                        if department:
                            break

                    if not department:
                        raise ValueError(f"在文件 {file} 中未找到科室名称")

                    # 查找数据开始行（序号为1的位置）
                    data_start_mask = df[0] == 1
                    if not data_start_mask.any():
                        raise ValueError(f"在文件 {file} 中未找到数据开始行")

                    data_start = data_start_mask.idxmax()

                    # 查找合计行
                    total_mask = df[0].astype(str).str.contains('合计', na=False)
                    if not total_mask.any():
                        raise ValueError(f"在文件 {file} 中未找到合计行")

                    data_end = total_mask.idxmax()

                    # 提取人员数据
                    for idx in range(data_start, data_end):
                        name = str(df.iloc[idx, 1]).strip()
                        try:
                            amount = float(df.iloc[idx, 4])
                            if pd.notna(amount):
                                all_department_data[department][name] = amount
                        except (ValueError, TypeError):
                            continue

                except Exception as e:
                    print(f"处理文件 {file} 时出错: {str(e)}")
                    continue

            # 创建结果DataFrame
            rows = []
            for department, names in all_department_data.items():
                for name, value in names.items():
                    rows.append({
                        '科室': department,
                        '姓名': name,
                        '小计': value
                    })

            result_df = pd.DataFrame(rows)

            if result_df.empty:
                raise ValueError("汇总数据为空")

            # 保存结果
            output_path = os.path.join(directory, '夜班费汇总表.xlsx')
            result_df.to_excel(output_path, index=False)

            messagebox.showinfo("成功", f"汇总表已保存至:\n{output_path}")

        except Exception as e:
            messagebox.showerror("错误", f"处理错误：\n{str(e)}")

    def run(self):
        self.window.mainloop()


if __name__ == "__main__":
    app = ExcelProcessor()
    app.run()
