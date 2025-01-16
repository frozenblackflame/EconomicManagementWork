import os
import pandas as pd

def process_excel_files(directory, output_folder):
    # 确保输出文件夹存在
    os.makedirs(output_folder, exist_ok=True)
    
    # 遍历目录下的所有Excel文件
    for file_name in os.listdir(directory):
        if file_name.endswith('.xlsx') and '临床积分明细' in file_name:
            file_path = os.path.join(directory, file_name)
            date_info = file_name.split('临床积分明细')[0].strip()  # 提取日期信息
            
            # 读取Excel文件，不直接指定列，因为需要动态查找
            df = pd.read_excel(file_path, header=None, skiprows=3)
            
            # 提取第4行和第5行的内容
            header_row_1 = df.iloc[0].tolist()  # 第4行
            header_row_2 = df.iloc[1].tolist()  # 第5行
            
            # 找到需要的列索引
            indices = {
                '出院人次': header_row_1.index('出院人次'),
                '门诊人次': header_row_1.index('门诊人次'),
                '3级手术': header_row_1.index('3级手术'),
                '4级手术': header_row_1.index('4级手术'),
                '3级微创手术': header_row_1.index('3级微创手术'),
                '4级微创手术': header_row_1.index('4级微创手术')
            }
            
            # 收集每个科室的数据
            for index, row in df.iterrows():
                if index >= 2:  # 从第6行开始读取
                    department_name = row[0]  # 科室名称
                    if pd.notna(department_name):  # 确保科室名称不是NaN
                        data = {
                            '日期': date_info,
                            '出院人次': row[indices['出院人次']+2],
                            '门诊人次': row[indices['门诊人次']+2],
                            '3级手术': row[indices['3级手术']+2],
                            '4级手术': row[indices['4级手术']+2],
                            '3级微创手术': row[indices['3级微创手术']+2],
                            '4级微创手术': row[indices['4级微创手术']+2],
                            '优势病种1': row[indices['优势病种1']+2],
                            '优势病种2': row[indices['优势病种2']+2],
                        }
                        data['微创手术'] = data['3级微创手术'] + data['4级微创手术']
                        del data['3级微创手术']
                        del data['4级微创手术']
                        
                        # 将数据转化为DataFrame
                        department_df = pd.DataFrame([data])
                        
                        # 构建输出文件路径
                        output_path = os.path.join(output_folder, f"{department_name}.xlsx")
                        
                        # 将DataFrame写入Excel文件
                        if os.path.exists(output_path):
                            with pd.ExcelWriter(output_path, mode='a', if_sheet_exists='overlay') as writer:
                                # 获取当前最大行数
                                current_max_row = writer.sheets['Sheet1'].max_row
                                
                                # 如果文件中已经有数据，从下一行开始写入
                                if current_max_row > 0:
                                    department_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=current_max_row)
                                else:
                                    department_df.to_excel(writer, sheet_name='Sheet1', index=False, header=True)
                        else:
                            department_df.to_excel(output_path, sheet_name='Sheet1', index=False, header=True)

# 使用示例
directory = r'C:\Users\biyun\Desktop\医生同比环比工作量提取'
output_folder = r'C:\Users\biyun\Desktop\医生同比环比工作量'
process_excel_files(directory, output_folder)
# 目录格式
'''
Mode                 LastWriteTime         Length Name
----                 -------------         ------ ----
-a----          2024/9/4     15:42          43282 2023.12临床积分明细.xlsx
-a----         2025/1/16     14:56          52291 2024.11临床积分明细.xlsx
-a----         2025/1/16     14:56          51795 2024.12临床积分明细.xlsx
'''