import os
import re
import shutil
import sys
import time
import zipfile
from datetime import datetime, timedelta

import pandas as pd
import win32con
import win32gui
from win32com.client import Dispatch


def extract_resources():
    """提取打包的资源文件到当前目录"""
    # 获取资源文件路径
    if getattr(sys, 'frozen', False):
        # 如果是打包后的程序
        base_path = sys._MEIPASS
    else:
        # 如果是开发环境
        base_path = os.path.abspath(os.path.dirname(__file__))

    # 需要提取的文件列表
    files_to_extract = [
        "积分_模板.xlsm",
        "绩效_模板.xlsm",
        "实际值_模板.xlsm"
    ]

    # 提取文件
    for filename in files_to_extract:
        source = os.path.join(base_path, filename)
        destination = os.path.join(os.getcwd(), filename)

        # 如果目标文件已存在，先删除
        if os.path.exists(destination):
            try:
                os.remove(destination)
            except Exception as e:
                print(f"删除已存在的文件 {filename} 失败: {str(e)}")
                continue

        # 复制文件
        try:
            shutil.copy2(source, destination)
            print(f"已释放文件: {filename}")
        except Exception as e:
            print(f"释放文件 {filename} 失败: {str(e)}")

def rename_excel_files():
    """重命名Excel文件，移除日期部分"""
    # 获取当前目录下所有xlsx文件
    excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]

    for file in excel_files:
        # 处理类似"2024年11月绩效业务实际值实际值上传模板_20241122155535.xlsx"的情况
        if '_' in file:
            new_name = file.split('_')[0]
            # 移除年月
            new_name = re.sub(r'\d{4}年\d{1,2}月绩效', '', new_name)
            new_name = new_name + '.xlsx'
        elif '服务人次工作量不通用上传模板' in file:
            new_name = '服务人次工作量不通用上传模板' + '.xlsx'
        elif '行政科室评分表' in file:
            new_name = '行政科室评分表.xlsx'
        else:
            # 处理类似"科室奖罚数据2024011.xlsx"的情况
            new_name = re.sub(r'\d+', '', file)

        # 如果新文件名不同于原文件名，则重命名
        if new_name != file:
            try:
                os.rename(file, new_name)
                print(f'已重命名: {file} -> {new_name}')
            except Exception as e:
                print(f'重命名失败 {file}: {str(e)}')


def process_penalty_reward_data():
    """处理科室奖罚数据，生成汇总报告"""
    try:
        # 读取科室奖罚数据文件
        df = pd.read_excel('科室奖罚数据.xlsx', header=1)  # 第二行为列名

        # 按科室名称分组并计算金额总和
        summary = df.groupby('科室名称')['金额'].sum().reset_index()

        # 重命名列以符合要求
        summary.columns = ['科室', '金额']

        # 保存为新的Excel文件
        summary.to_excel('奖罚总计.xlsx', index=False)
        print('已生成奖罚总计.xlsx')

    except FileNotFoundError:
        print('未找到科室奖罚数据.xlsx文件')
    except Exception as e:
        print(f'处理科室奖罚数据时出错: {str(e)}')


def find_dialog_and_click_yes():
    """查找Excel对话框并自动点击"是"按钮"""

    def callback(handle, dialog_list):
        title = win32gui.GetWindowText(handle)
        if 'Microsoft Excel' in title:
            dialog_list.append(handle)
        return True

    dialog_list = []
    time.sleep(1)
    win32gui.EnumWindows(callback, dialog_list)

    for dialog in dialog_list:
        yes_button = win32gui.FindWindowEx(dialog, 0, "Button", "是(&Y)")
        if yes_button:
            win32gui.PostMessage(yes_button, win32con.WM_LBUTTONDOWN, 0, 0)
            win32gui.PostMessage(yes_button, win32con.WM_LBUTTONUP, 0, 0)
            return True
    return False


def get_previous_month_info():
    """获取上个月的年份和月份"""
    current_date = datetime.now()
    first_day_of_current_month = current_date.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
    return last_day_of_previous_month.year, last_day_of_previous_month.month


def create_folder_name():
    """创建文件夹名称，格式：YYYY年MM月绩效文件"""
    year, month = get_previous_month_info()
    return f"{year}年{month:02d}月绩效文件"


def create_file_name(original_name):
    """根据原始文件名创建新的文件名"""
    year, month = get_previous_month_info()

    # 提取文件类型（积分或绩效）
    if "积分" in original_name:
        file_type = "积分"
    elif "绩效" in original_name:
        file_type = "绩效"
    elif "实际值" in original_name:
        file_type = "实际值"
    else:
        file_type = "未知类型"

    return f"{year}年{month:02d}月{file_type}文件.xlsx"


def process_excel_files(files_to_convert):
    """处理Excel文件并创建ZIP包"""
    # 创建Excel应用实例
    excel = Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False
    excel.AlertBeforeOverwriting = False

    try:
        # 设置Excel安全级别
        app = excel.Application
        app.AutomationSecurity = 1

        # 创建目标文件夹
        folder_name = create_folder_name()
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        target_folder = os.path.join(desktop_path, folder_name)

        # 如果文件夹已存在，先删除它
        if os.path.exists(target_folder):
            shutil.rmtree(target_folder)

        # 创建新文件夹
        os.makedirs(target_folder)

        saved_files = []  # 记录保存的文件路径

        for file_path in files_to_convert:
            # 获取原始文件名
            original_name = os.path.basename(file_path)

            # 创建新的文件名
            new_filename = create_file_name(original_name)
            new_file_path = os.path.join(target_folder, new_filename)

            # 打开工作簿
            workbook = excel.Workbooks.Open(file_path)

            try:
                # 更新所有计算
                workbook.Application.CalculateFull()

                # 遍历所有工作表
                for sheet in workbook.Worksheets:
                    used_range = sheet.UsedRange
                    used_range.Value = used_range.Value

                # 保存文件
                workbook.SaveAs(new_file_path, FileFormat=51, ConflictResolution=2)

                # 处理可能出现的对话框
                find_dialog_and_click_yes()

                saved_files.append(new_file_path)
                print(f"已保存文件: {new_file_path}")

            finally:
                workbook.Close(SaveChanges=False)

        # 创建ZIP文件
        zip_path = os.path.join(desktop_path, f"{folder_name}.zip")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in saved_files:
                arcname = os.path.basename(file_path)
                zipf.write(file_path, arcname)

        print(f"已创建ZIP文件: {zip_path}")

    finally:
        excel.Quit()

    print("所有文件处理完成!")


def main():
    """主函数 - 按顺序执行所有步骤"""
    print('开始释放必要文件...')
    extract_resources()

    print('开始处理Excel文件...')

    # 第一步：重命名文件和处理奖罚数据
    print('\n===== 执行第一步：重命名文件和处理奖罚数据 =====')
    rename_excel_files()
    process_penalty_reward_data()

    # 第二步：处理模板文件并创建ZIP包
    print('\n===== 执行第二步：处理模板文件并创建ZIP包 =====')
    current_dir = os.getcwd()
    files_to_convert = [
        os.path.join(current_dir, "积分_模板.xlsm"),
        os.path.join(current_dir, "绩效_模板.xlsm"),
        os.path.join(current_dir, "实际值_模板.xlsm")
    ]

    # 检查文件是否存在
    for file_path in files_to_convert:
        if not os.path.exists(file_path):
            print(f"错误：文件 {file_path} 不存在！")
            exit(1)

    process_excel_files(files_to_convert)

    print('\n所有步骤执行完成!')


# 创建 performance.spec 文件
# 运行 pyinstaller performance.spec
"""
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# 定义数据文件
added_files = [
    ('template/积分_模板.xlsm', '.'),
    ('template/绩效_模板.xlsm', '.'),
    ('template/实际值_模板.xlsm', '.')
]

a = Analysis(
    ['绩效计算生成.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=['win32timezone'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='绩效计算生成',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None
)
"""

if __name__ == '__main__':
    main()
