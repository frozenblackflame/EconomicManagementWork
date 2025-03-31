import pandas as pd
import os

# 读取Excel文件
df = pd.read_excel(r"C:\Users\biyun\Desktop\work\点值数.xlsx")

# 1. 按照B列(科室)分类创建工作表
# 获取B列中所有唯一的科室名称
departments = df['B'].unique()

# 创建一个ExcelWriter对象
output_path = r"C:\Users\biyun\Desktop\work\分类结果.xlsx"
with pd.ExcelWriter(output_path) as writer:
    # 对每个科室进行处理
    for dept in departments:
        # 获取该科室的所有数据
        dept_data = df[df['B'] == dept].copy()
        
        # 2. 根据D列进行分类
        # 按D列排序，这样相同类型的会排在一起
        dept_data = dept_data.sort_values('D')
        
        # 在相同类型之间添加空行
        previous_type = None
        rows_to_add = []
        
        for index, row in dept_data.iterrows():
            current_type = row['D']
            # 比较时考虑 NaN 的情况
            if previous_type is not None and pd.notna(current_type) and pd.notna(previous_type):
                if current_type != previous_type:
                    empty_row = pd.Series([None] * len(row), index=row.index)
                    rows_to_add.append((index, empty_row))
            elif previous_type is not None and (pd.isna(current_type) != pd.isna(previous_type)):
                empty_row = pd.Series([None] * len(row), index=row.index)
                rows_to_add.append((index, empty_row))
            previous_type = current_type
        
        # 插入空行
        for idx, empty_row in sorted(rows_to_add, reverse=True):
            dept_data = pd.concat([dept_data.iloc[:idx], 
                                 pd.DataFrame([empty_row]), 
                                 dept_data.iloc[idx:]]).reset_index(drop=True)
        
        # 重命名列
        column_mapping = {
            'A': '核算单元名称',
            'B': '科室名称',
            'C': '科室名称',
            'D': '名称',
            'E': '项目名称',
            'F': '点数'
        }
        dept_data = dept_data.rename(columns=column_mapping)
        
        # 将处理后的数据写入到对应的工作表中
        dept_data.to_excel(writer, sheet_name=dept, index=False)

print("处理完成！结果已保存到：", output_path)
