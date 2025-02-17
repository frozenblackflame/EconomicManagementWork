import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd


def process_files(file_paths):
    results = {}
    all_names = []
    # 定义排除项列表，方便后期手动添加
    exclude_list = ["合计", "小计", "科主任", "总计", "备注", "病区"]
    for file_path in file_paths:
        # 获取文件所在文件夹名称
        folder_name = os.path.basename(os.path.dirname(file_path))
        current_names = set()
        # 读取 Excel 文件
        excel_file = pd.ExcelFile(file_path)
        # 获取所有表名
        sheet_names = excel_file.sheet_names
        for sheet_name in sheet_names:
            # 读取工作表数据
            df = excel_file.parse(sheet_name)
            # 去除空行
            df = df.dropna(how='all')
            # 遍历每个单元格
            for col in df.columns:
                for index, value in df[col].items():
                    if isinstance(value, str):
                        # 去除所有空格
                        clean_value = value.replace(" ", "")
                        if clean_value == "姓名":
                            # 将列名转换为整数索引
                            col_index = df.columns.get_loc(col)
                            # 从“姓名”单元格下一行开始获取姓名列数据
                            name_column = df.iloc[index + 1:, col_index]
                            for name in name_column.dropna():
                                if isinstance(name, str):
                                    clean_name = name.replace(" ", "")
                                    # 排除以排除项列表中元素开头的内容
                                    should_exclude = False
                                    for item in exclude_list:
                                        if clean_name.startswith(item):
                                            should_exclude = True
                                            break
                                    if not should_exclude:
                                        current_names.add(clean_name)
                                        all_names.append(clean_name)

        # 统计去重后的姓名数量
        count = len(current_names)
        if folder_name in results:
            results[folder_name] += count
        else:
            results[folder_name] = count

    return results, all_names


def select_and_process():
    # 选择多个.xlsx 文件
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    if file_paths:
        # 处理文件
        results, all_names = process_files(file_paths)

        # 清空之前的结果显示
        names_text.delete(1.0, tk.END)
        counts_text.delete(1.0, tk.END)

        # 显示统计到的姓名
        names_text.insert(tk.END, "统计到的姓名如下：\n")
        for name in all_names:
            names_text.insert(tk.END, f"{name}\n")

        # 显示各文件夹的人数统计结果
        counts_text.insert(tk.END, "各文件夹的人数统计结果如下：\n")
        for folder, count in results.items():
            counts_text.insert(tk.END, f"文件夹名称: {folder}, 人数: {count}\n")


# 创建主窗口
root = tk.Tk()
root.title("Excel 姓名人数统计")

# 创建选择文件按钮
select_button = tk.Button(root, text="选择文件", command=select_and_process)
select_button.pack(pady=20)

# 创建一个文本框用于显示姓名
names_text = tk.Text(root, height=10, width=30)
names_text.pack()

# 创建一个文本框用于显示各文件夹人数
counts_text = tk.Text(root, height=10, width=30)
counts_text.pack()

# 运行主循环
root.mainloop()