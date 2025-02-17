import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd


def select_files():
    root = tk.Tk()
    root.withdraw()
    # 选择多个.xlsx 文件
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    return file_paths


def process_files(file_paths):
    results = {}
    # 定义排除项列表，方便后期手动添加
    exclude_list = ["合计", "小计"]
    for file_path in file_paths:
        # 获取文件所在文件夹名称
        folder_name = os.path.basename(os.path.dirname(file_path))
        all_names = set()
        # 读取 Excel 文件
        excel_file = pd.ExcelFile(file_path)
        # 获取所有表名
        sheet_names = excel_file.sheet_names
        for sheet_name in sheet_names:
            # 读取工作表数据
            df = excel_file.parse(sheet_name)
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
                                        all_names.add(clean_name)
                                        print(clean_name)  # 打印姓名日志

        # 统计去重后的姓名数量
        count = len(all_names)
        if folder_name in results:
            results[folder_name] += count
        else:
            results[folder_name] = count

    return results


def save_results(results):
    # 获取桌面路径
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    output_file = os.path.join(desktop_path, "人数.xlsx")

    data = [[folder, count] for folder, count in results.items()]
    new_df = pd.DataFrame(data, columns=["文件夹名称", "人数"])

    if os.path.exists(output_file):
        # 如果文件已存在，读取原文件内容
        existing_df = pd.read_excel(output_file)
        # 追加新数据
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        combined_df = new_df

    # 保存结果到 Excel 文件
    combined_df.to_excel(output_file, index=False)
    print(f"结果已保存到 {output_file}")


if __name__ == "__main__":
    # 选择文件
    file_paths = select_files()
    if file_paths:
        # 处理文件
        results = process_files(file_paths)
        # 保存结果
        save_results(results)
    else:
        print("未选择任何文件。")