import tkinter as tk
import tkinter.messagebox as messagebox
from tkinter import ttk, filedialog

import pandas as pd
from openpyxl import load_workbook


class NameComparisonApp:
    def __init__(self, master):
        self.master = master
        master.title("姓名对比工具")
        master.geometry("800x850")

        # 黑名单
        self.blacklist = ["小计", "nan", "科主任：", "合计"]

        # 主容器
        main_frame = tk.Frame(master, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件选择
        self.file_frame = tk.Frame(main_frame)
        self.file_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=5)

        self.file_path = tk.StringVar()
        self.file_path_entry = tk.Entry(self.file_frame, textvariable=self.file_path, width=50)
        self.file_path_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))

        self.select_file_button = tk.Button(self.file_frame, text="选择文件", command=self.select_file)
        self.select_file_button.pack(side=tk.LEFT)

        self.load_data_button = tk.Button(self.file_frame, text="加载数据", command=self.load_excel_data)
        self.load_data_button.pack(side=tk.LEFT, padx=(10, 0))

        # 工作表和列配置
        tk.Label(main_frame, text="工作表名称").grid(row=1, column=0, sticky='w')
        self.sheet_entry = tk.Entry(main_frame, width=20)
        self.sheet_entry.grid(row=1, column=1, sticky='w')
        self.sheet_entry.insert(0, "Sheet1")

        tk.Label(main_frame, text="绩效列名").grid(row=2, column=0, sticky='w')
        self.performance_column_entry = tk.Entry(main_frame, width=20)
        self.performance_column_entry.grid(row=2, column=1, sticky='w')

        # 原有的文本框和表格部分
        tk.Label(main_frame, text="二次分配表").grid(row=4, column=0, sticky='w')
        tk.Label(main_frame, text="对比表").grid(row=4, column=1, sticky='w')

        self.text1 = tk.Text(main_frame, height=10, width=40)
        self.text1.grid(row=5, column=0, padx=10, sticky='nsew')

        self.text2 = tk.Text(main_frame, height=10, width=40)
        self.text2.grid(row=5, column=1, padx=10, sticky='nsew')

        # 金额输入
        tk.Label(main_frame, text="金额（对应对比表顺序）").grid(row=6, column=0, columnspan=2, pady=5)
        self.text3 = tk.Text(main_frame, height=5, width=40)
        self.text3.grid(row=7, column=0, columnspan=2, padx=10, sticky='nsew')

        # 姓名差异
        tk.Label(main_frame, text="姓名差异").grid(row=8, column=0, columnspan=2, pady=5)
        self.text4 = tk.Text(main_frame, height=5, width=40, state='disabled')
        self.text4.grid(row=9, column=0, columnspan=2, padx=10, sticky='nsew')

        # 按钮行
        self.copy_button = tk.Button(main_frame, text="复制金额", command=self.copy_amounts)
        self.copy_button.grid(row=10, column=0, columnspan=2, pady=10)

        # 结果显示区域
        self.result_tree = ttk.Treeview(main_frame, columns=('Name', 'Amount'), show='headings')
        self.result_tree.heading('Name', text='姓名')
        self.result_tree.heading('Amount', text='金额')
        self.result_tree.grid(row=11, column=0, columnspan=2, sticky='nsew', padx=10, pady=10)

        # 配置网格权重
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_rowconfigure(11, weight=1)

        # 绑定文本框内容变化事件
        self.text1.bind('<KeyRelease>', self.compare_names)
        self.text2.bind('<KeyRelease>', self.compare_names)
        self.text3.bind('<KeyRelease>', self.compare_names)

        # 存储当前剪切板金额
        self.current_amounts = []

    def select_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.file_path.set(filename)
            self.load_excel_data()  # 自动加载数据


    def load_excel_data(self):
        try:
            # 获取文件路径
            file_path = self.file_path.get()

            # 使用 openpyxl 加载工作簿以获取激活的工作表名称
            workbook = load_workbook(file_path, read_only=True)
            active_sheet_name = workbook.active.title
            workbook.close()

            # 将激活工作表名称回显到工作表名称输入框
            self.sheet_entry.delete(0, tk.END)
            self.sheet_entry.insert(0, active_sheet_name)

            # 获取工作表名称和绩效列名
            sheet_name = active_sheet_name
            performance_column = self.performance_column_entry.get()

            # 使用 Pandas 加载指定工作表的数据
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine="openpyxl")

            # 查找"姓名"列所在的列名行
            header_row_index = None
            for i, row in df.iterrows():
                cleaned_row = [str(cell).replace(" ", "") for cell in row.values]
                if "姓名" in cleaned_row:
                    header_row_index = i
                    break

            if header_row_index is None:
                messagebox.showerror("错误", "未找到姓名列")
                return

            # 设置列名行
            df.columns = df.iloc[header_row_index].apply(lambda x: str(x).replace(" ", ""))
            df = df[(header_row_index + 1):].reset_index(drop=True)

            # 检查绩效列是否存在
            if performance_column not in df.columns:
                messagebox.showerror("错误", "指定的绩效列名不存在")
                return

            # 清空对比表和金额文本框
            self.text2.delete('1.0', tk.END)
            self.text3.delete('1.0', tk.END)

            # 创建姓名与金额字典
            name_amount_dict = {}
            for _, row in df.iterrows():
                name = str(row.get("姓名", "")).strip()
                amount = str(row.get(performance_column, "")).strip()
                if name and name not in self.blacklist:  # 跳过姓名为空或黑名单的行
                    name_amount_dict[name] = amount

            # 填入对比表和金额框
            for name, amount in name_amount_dict.items():
                self.text2.insert(tk.END, name + '\n')
                self.text3.insert(tk.END, amount + '\n')

            # 触发名称比较
            self.compare_names()

            # 自动复制金额
            self.copy_amounts()

        except Exception as e:
            messagebox.showerror("错误", f"加载数据时发生错误: {str(e)}")

    def compare_names(self, event=None):
        # 获取输入
        allocation_names = [name.strip() for name in self.text1.get('1.0', tk.END).split('\n') if name.strip()]
        comparison_names = [name.strip() for name in self.text2.get('1.0', tk.END).split('\n') if name.strip()]
        amounts = [amount.strip() for amount in self.text3.get('1.0', tk.END).split('\n') if amount.strip()]

        # 计算两个表中互相没有的姓名
        names_only_in_allocation = list(set(allocation_names) - set(comparison_names))
        names_only_in_comparison = list(set(comparison_names) - set(allocation_names))

        # 清空结果表
        for i in self.result_tree.get_children():
            self.result_tree.delete(i)

        # 更新差异输出文本框
        self.text4.config(state='normal')
        self.text4.delete('1.0', tk.END)

        # 显示两个表中互相没有的姓名
        if names_only_in_allocation or names_only_in_comparison:
            diff_message = "差异姓名：\n"
            if names_only_in_allocation:
                diff_message += "仅在二次分配表的姓名：" + "，".join(names_only_in_allocation) + "\n"
            if names_only_in_comparison:
                diff_message += "仅在对比表的姓名：" + "，".join(names_only_in_comparison)
            self.text4.insert(tk.END, diff_message)
        else:
            self.text4.insert(tk.END, "两个表的姓名完全一致")

        self.text4.config(state='disabled')

        # 创建姓名和金额映射
        comparison_dict = {}
        for i, name in enumerate(comparison_names):
            if i < len(amounts):
                comparison_dict[name] = amounts[i]

        # 按二次分配表顺序重排并填充结果表
        self.current_amounts = []
        for name in allocation_names:
            amount = comparison_dict.get(name, '')
            self.current_amounts.append(str(amount))

            # 设置颜色
            if name not in comparison_names:
                self.result_tree.insert('', 'end', values=(name, amount), tags=('red',))
            else:
                self.result_tree.insert('', 'end', values=(name, amount))

        # 添加标签样式
        self.result_tree.tag_configure('red', foreground='red')

    def copy_amounts(self):
        if self.current_amounts:
            self.master.clipboard_clear()
            self.master.clipboard_append('\n'.join(self.current_amounts))
        else:
            messagebox.showwarning("复制失败", "没有可复制的金额")

def main():
    root = tk.Tk()
    app = NameComparisonApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
