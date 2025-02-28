import pandas as pd

# 读取Excel文件，从第二行开始读取数据
file_path = r"C:\Users\biyun\Desktop\新建文件夹\2024年度目标责任制考核绩效发放表.xlsx"
df = pd.read_excel(file_path, header=1)  # header=1表示从第二行开始读取

# 检查列名是否存在，如果不存在则使用默认列名
if 'F' not in df.columns:
    df.columns = ['A', 'B', 'C', 'D', 'E', 'F'] + list(df.columns[6:])  # 假设前6列是A-F

# 处理F列（备注列）：清空“不在岗xx天”中天数小于30的记录的F列和E列的值
df['days'] = df['F'].str.extract(r'不在岗(\d+)天').astype(float)
df.loc[(df['days'] < 30) & (df['days'].notna()), ['F']] = None
df.drop(columns=['days'], inplace=True)

# 处理E列（金额列）：将所有金额除以2
df['E'] = df['E'] / 2

# 保存修改后的Excel文件
output_path = r"C:\Users\biyun\Desktop\新建文件夹\2024年度目标责任制考核绩效发放表_processed.xlsx"
df.to_excel(output_path, index=False)