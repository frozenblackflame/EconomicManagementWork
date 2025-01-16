import os
import pandas as pd

def process_excel_files(folder_path):
    # 获取桌面路径
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    output_folder = os.path.join(desktop_path, '医生的工作量')
    
    # 如果输出文件夹不存在，则创建
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # 遍历文件夹中的所有文件
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(folder_path, file_name)
            date_info = file_name.split('临床积分明细.xlsx')[0]  # 提取日期信息
            
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
            department_data = {}
            for index, row in df.iterrows():
                if index >= 2:  # 从第6行开始读取
                    department_name = row[0]  # 科室名称
                    if pd.notna(department_name):  # 确保科室名称不是NaN
                        if department_name not in department_data:
                            department_data[department_name] = {
                                '出院人次': row[indices['出院人次']],
                                '门诊人次': row[indices['门诊人次']],
                                '3级手术': row[indices['3级手术']],
                                '4级手术': row[indices['4级手术']],
                                '3级微创手术': row[indices['3级微创手术']],
                                '4级微创手术': row[indices['4级微创手术']],
                            }

            # 将计算结果添加到部门数据
            for department, data in department_data.items():
                data['微创手术'] = data['3级微创手术'] + data['4级微创手术']

                # 将数据转化为DataFrame
                department_df = pd.DataFrame([data])
                department_df.insert(0, '科室', department)  # 插入科室列

                # 保存到新的Excel文件中
                output_file_path = os.path.join(output_folder, f'{department}.xlsx')
                with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a' if os.path.exists(output_file_path) else 'w') as writer:
                    department_df.to_excel(writer, sheet_name=date_info, index=False)

# 使用方法
folder_path = r'C:\Users\biyun\Desktop\新建文件夹'  # 替换为你的文件夹地址
process_excel_files(folder_path)
