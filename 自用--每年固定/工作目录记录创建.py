import os
from datetime import datetime

def create_work_directories(year):
    # 获取桌面路径
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    
    # 创建主目录
    main_dir = os.path.join(desktop_path, f"{year}年工作记录")
    
    try:
        # 创建主目录
        if not os.path.exists(main_dir):
            os.makedirs(main_dir)
            print(f"创建主目录：{main_dir}")
        
        # 创建1-12月的子目录
        for month in range(1, 13):
            month_dir = os.path.join(main_dir, f"{month}月工作记录")
            if not os.path.exists(month_dir):
                os.makedirs(month_dir)
                print(f"创建子目录：{month_dir}")
                
                # 创建“原始”，“修改”，“其他”子目录
                for sub_dir in ["原始", "修改", "其他"]:
                    sub_dir_path = os.path.join(month_dir, sub_dir)
                    if not os.path.exists(sub_dir_path):
                        os.makedirs(sub_dir_path)
                        print(f"创建子目录：{sub_dir_path}")
        
        # 创建年度总结报告子目录
        report_dir = os.path.join(main_dir, "年度总结报告")
        if not os.path.exists(report_dir):
            os.makedirs(report_dir)
            print(f"创建年度总结报告目录：{report_dir}")
        
        print("\n目录创建完成！")
        
    except Exception as e:
        print(f"创建目录时出错：{str(e)}")

# 获取用户输入的年份
current_year = datetime.now().year

while True:
    user_input = input(f"请输入年份（直接回车使用当前年份 {current_year}）：").strip()
    try:
        if user_input == "":
            year = current_year
            break
        year = int(user_input)
        if 1900 <= year <= 9999:  # 设置合理的年份范围
            break
        else:
            print("请输入有效的年份（1900-9999）")
    except ValueError:
        print("请输入有效的数字年份")

# 创建目录
create_work_directories(year)