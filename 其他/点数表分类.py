import pandas as pd
import os

# 读取Excel文件
df = pd.read_excel(r"C:\Users\biyun\Desktop\work\点值数.xlsx")

# 统计数据行数（不包括表头）
total_rows = len(df)
print(f"总数据行数（不含表头）：{total_rows} 行")

# 1. 按照B列(科室)分类创建工作表
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
        
        # 创建新的DataFrame来存储结果
        result_data = []
        
        # 遍历数据行（除了最后一行）
        for i in range(len(dept_data) - 1):
            current_row = dept_data.iloc[i]
            next_row = dept_data.iloc[i + 1]
            
            # 添加当前行
            result_data.append(current_row)
            
            # 如果当前行和下一行的"名称"列值不同，添加空行
            if current_row['D'] != next_row['D']:
                empty_row = pd.Series([None] * len(current_row), index=current_row.index)
                result_data.append(empty_row)
        
        # 添加最后一行
        result_data.append(dept_data.iloc[-1])
        
        # 将结果转换为DataFrame
        dept_data = pd.DataFrame(result_data)
        
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
