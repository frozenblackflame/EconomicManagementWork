# 首先导入所需的包
import os
from datetime import datetime
from typing import List, Dict, Optional, Tuple, Any

import pandas as pd

current_date = datetime.now()

# DateValidator类定义
class DateValidator:
    @staticmethod
    def is_date_in_range(date_str: str) -> bool:
        """检查日期是否在2023.10到2024.10之间"""
        try:
            year = int(date_str.split(".")[0])
            month = int(date_str.split(".")[1])
            date_num = year * 100 + month
            return int(current_date.replace(year=current_date.year - 1, month=current_date.month - 1).strftime(
                "%Y%m")) <= date_num <= int(current_date.replace(month=current_date.month - 1).strftime("%Y%m"))
        except:
            return False

    @staticmethod
    def normalize_date(date_str: str) -> str:
        """将不同格式的日期统一转换为YYYY.MM格式"""
        date_str = date_str.strip()
        if "年" in date_str and "月" in date_str:
            year = date_str.split("年")[0]
            month = date_str.split("年")[1].replace("月", "")
            return f"{year}.{month.zfill(2)}"
        elif len(date_str) == 6:
            return f"{date_str[:4]}.{date_str[4:]}"
        return date_str


# ExcelDataReader类定义
class ExcelDataReader:
    def __init__(self, base_path: str, department_name: str):
        self.base_path = base_path
        self.department_name = department_name
        self.date_validator = DateValidator()

    def find_department_row(self, df: pd.DataFrame) -> Optional[pd.Series]:
        """在DataFrame中查找指定科室所在的行"""
        dept_row = df[df.iloc[:, 0] == self.department_name]
        return dept_row.iloc[0] if not dept_row.empty else None

    def read_excel_file(
        self, file_path: str
    ) -> Tuple[Optional[pd.DataFrame], Optional[pd.Series], Optional[pd.Series]]:
        """读取Excel文件并返回相关的行数据"""
        try:
            df = pd.read_excel(file_path)
            dept_row = self.find_department_row(df)
            if dept_row is not None:
                header_row = df.iloc[2]  # 项目名称行
                workload_row = df.iloc[3]  # 工作量行
                return df, header_row, workload_row
        except Exception as e:
            print(f"读取文件 {file_path} 时出错: {str(e)}")
        return None, None, None

    def process_excel_files(self, data_extractor) -> List[Dict[str, Any]]:
        """使用提供的数据提取器处理Excel文件"""
        all_data = []
        for root, dirs, files in os.walk(self.base_path):
            for file in files:
                if file.startswith("~$") or not file.endswith(".xlsx"):
                    continue
                file_path = os.path.join(root, file)
                date_str = self.date_validator.normalize_date(
                    file.split("绩效")[0].strip()
                )
                if not self.date_validator.is_date_in_range(date_str):
                    continue
                df, header_row, workload_row = self.read_excel_file(file_path)
                if all(x is not None for x in (df, header_row, workload_row)):
                    data = data_extractor(date_str, df, header_row, workload_row)
                    if data:
                        all_data.append(data)
        return all_data


# DepartmentAnalyzer类定义
class DepartmentAnalyzer:
    def __init__(self, base_path: str, department_name: str):
        self.reader = ExcelDataReader(base_path, department_name)
        self.department_name = department_name

    def extract_visits_data(
        self,
        date_str: str,
        df: pd.DataFrame,
        header_row: pd.Series,
        workload_row: pd.Series,
    ) -> Dict[str, Any]:
        """从Excel文件中提取出院人次和门诊人次数据"""
        dept_data = self.reader.find_department_row(df)
        if dept_data is None:
            return None

        data = {"月份": date_str}  # 保持原始日期格式
        metrics = ["出院人次", "门诊人次"]

        for i in range(len(header_row)):
            header_value = str(header_row.iloc[i]).strip()
            if header_value in metrics:
                try:
                    workload_index = i + 2
                    if workload_index < len(workload_row):
                        workload_value = str(workload_row.iloc[workload_index]).strip()
                        if workload_value == "工作量":
                            raw_value = dept_data.iloc[workload_index]
                            data[header_value] = self._convert_to_float(raw_value)
                except IndexError:
                    continue

        # 确保所有指标都有值
        for metric in metrics:
            if metric not in data:
                data[metric] = 0

        return data

    def extract_clinical_points_data(
        self,
        date_str: str,
        df: pd.DataFrame,
        header_row: pd.Series,
        workload_row: pd.Series,
    ) -> Dict[str, Any]:
        """从Excel文件中提取临床积分数据"""
        dept_data = self.reader.find_department_row(df)
        if dept_data is None:
            return None

        # 基础指标
        base_metrics = [
            "出院人次",
            "门诊人次",
            "3级手术",
            "4级手术",
        ]

        # 查找所有优势病种
        advantage_diseases = []
        for i in range(len(header_row)):
            header_value = str(header_row.iloc[i]).strip()
            # 检查是否符合一个大写字母后跟两位数字的模式
            if (
                len(header_value) == 3
                and header_value[0].isalpha()
                and header_value[0].isupper()
                and header_value[1:].isdigit()
            ):
                try:
                    workload_index = i + 2
                    if workload_index < len(workload_row):
                        workload_value = str(workload_row.iloc[workload_index]).strip()
                        if workload_value == "工作量":
                            raw_value = dept_data.iloc[workload_index]
                            converted_value = self._convert_to_float(raw_value)
                            if converted_value > 0:  # 只添加有数据的优势病种
                                advantage_diseases.append(header_value)
                except IndexError:
                    continue

        # 将找到的优势病种添加到指标列表中
        metrics = base_metrics + [f"优势病种{code}" for code in advantage_diseases]

        # 标准化日期格式
        display_date = self._normalize_display_date(date_str)
        if not display_date:
            return None

        data = {"月份": display_date}

        # 提取基础指标数据
        for i in range(len(header_row)):
            header_value = str(header_row.iloc[i]).strip()
            if header_value in base_metrics:
                try:
                    workload_index = i + 2
                    if workload_index < len(workload_row):
                        workload_value = str(workload_row.iloc[workload_index]).strip()
                        if workload_value == "工作量":
                            raw_value = dept_data.iloc[workload_index]
                            converted_value = self._convert_to_float(raw_value)
                            if converted_value is not None:
                                data[header_value] = converted_value
                except IndexError:
                    continue

        # 提取优势病种数据
        for i in range(len(header_row)):
            header_value = str(header_row.iloc[i]).strip()
            if header_value in advantage_diseases:
                try:
                    workload_index = i + 2
                    if workload_index < len(workload_row):
                        workload_value = str(workload_row.iloc[workload_index]).strip()
                        if workload_value == "工作量":
                            raw_value = dept_data.iloc[workload_index]
                            converted_value = self._convert_to_float(raw_value)
                            if converted_value is not None:
                                data[f"优势病种{header_value}"] = converted_value
                except IndexError:
                    continue

        # 确保所有指标都有值
        for metric in metrics:
            if metric not in data:
                data[metric] = 0

        return data

    @staticmethod
    def _normalize_display_date(date_str: str) -> Optional[str]:
        """标准化显示日期格式（仅用于临床积分数据）"""
        date_str = date_str.strip()
        if current_date.replace(year=current_date.year - 1, month=current_date.month - 1).strftime("%Y.%m") in date_str:
            return "去年" + current_date.replace(year=current_date.year - 1, month=current_date.month - 1).strftime(
                "%m") + "月"
        elif current_date.replace(month=current_date.month - 2).strftime("%Y.%m") in date_str:
            return current_date.replace(month=current_date.month - 2).strftime("%m") + "月"
        elif current_date.replace(month=current_date.month - 1).strftime("%Y.%m") in date_str:
            return current_date.replace(month=current_date.month - 1).strftime("%m") + "月"
        return None

    @staticmethod
    def _convert_year_month(date_str: str) -> int:
        """转换日期为数值格式（用于绩效数据）"""
        try:
            year = int(date_str.split(".")[0])
            month = int(date_str.split(".")[1])
            return year * 100 + month
        except:
            return 0

    @staticmethod
    def _convert_to_float(value: Any) -> float:
        """将值转换为浮点数，如果转换失败则返回0"""
        try:
            if pd.notnull(value) and str(value).strip() not in ["", "--"]:
                return float(value) if str(value).replace(".", "").isdigit() else 0
        except:
            pass
        return 0

    def _create_dataframe(
        self, data: List[Dict[str, Any]], output_file: str
    ) -> Optional[pd.DataFrame]:
        """从提取的数据创建并保存DataFrame"""
        if not data:
            print(f"未找到{output_file}的相关数据")
            return None

        # 创建DataFrame
        df = pd.DataFrame(data)

        # 根据不同的文件类型设置列顺序
        if "临床积分" in output_file:
            # 确定所有出现的列
            all_columns = df.columns.tolist()

            # 基础列
            base_columns = ["月份", "出院人次", "门诊人次", "3级手术", "4级手术"]

            # 找到所有优势病种列并排序
            advantage_columns = [
                col for col in all_columns if col.startswith("优势病种")
            ]
            advantage_columns.sort()  # 按字母顺序排序优势病种

            # 合并所有列
            columns = base_columns + advantage_columns

            # 临床积分数据的特殊排序
            month_order = {"去年" + current_date.replace(month=current_date.month - 1).strftime("%m") + "月": 1,
                           current_date.replace(month=current_date.month - 2).strftime("%m") + "月": 2,
                           current_date.replace(month=current_date.month - 1).strftime("%m") + "月": 3}
            df["排序"] = df["月份"].map(month_order)
            df = df.sort_values("排序").drop("排序", axis=1)
        elif "出院人次" in output_file:
            columns = ["月份", "出院人次", "门诊人次"]
            df = df.sort_values("月份")
        else:
            # 绩效数据
            columns = [
                "月份",
                "合计得分",
                "床均产值",
                "护均担负床日",
                "出院人次绩酬率",
                "领用耗材占开单收入比",
                "住院西成药占比",
                "开单收入成本率",
            ]
            df = df.sort_values("月份")

        # 确保所有列都存在
        for col in columns:
            if col not in df.columns:
                df[col] = 0

        # 按指定列顺序排序
        df = df[columns]
        df = df.reset_index(drop=True)

        # 保存到Excel
        df.to_excel(output_file, index=False)
        print(f"已生成文件: {output_file}")
        return df

    def analyze_all(
        self,
    ) -> tuple[Any, Any]:
        """运行所有分析并返回结果"""
        print("\n开始分析数据...")

        print("\n1. 分析出院人次和门诊人次...")
        visits_data = self.reader.process_excel_files(self.extract_visits_data)
        # C:\Users\biyun\Desktop\work\出院人次、门诊人次\
        df_visits = self._create_dataframe(
            visits_data,
            f"C:\\Users\\biyun\\Desktop\\work\\出院人次门诊人次\\{self.department_name}出院人次、门诊人次.xlsx"
        )

        print("\n2. 分析临床积分...")
        clinical_data = self.reader.process_excel_files(
            self.extract_clinical_points_data
        )
        df_clinical = self._create_dataframe(
            clinical_data, f"C:\\Users\\biyun\\Desktop\\work\\临床积分\\{self.department_name}临床积分.xlsx"
        )

        print("\n分析完成！")
        return df_visits, df_clinical


# main函数定义
def main():
    """主函数，用于运行分析"""
    base_path = input("请输入Excel文件夹所在路径（临床积分明细（2018.06-2024.10））：")
    department_list = [
        "头颈一科",
        "头颈二科",
        "妇瘤一科",
        "妇瘤二科",
        "骨与软组织一科",
        "骨与软组织二科",
        "乳腺一科",
        "乳腺二科",
        "胸外一科",
        "胸外二科",
        "胃肿瘤外科",
        "肝胆胰肿瘤外科",
        "泌尿外科",
        "结直肠肿瘤外科",
        "乳腺内科",
        "放射治疗科",
        "中西医结合科",
        "血液病科",
        "消化肿瘤内一科",
        "消化肿瘤内二科",
        "呼吸肿瘤内科",
        "重症医学科",
        "麻醉手术科",
        "麻醉手术科护士",
        "放疗机房",
        "介入治疗科",
        "门诊部",
        "眼科",
        "口腔科",
        "皮肤科",
        "便民门诊",
        "针灸理疗室",
        "多学科联合(MDT)门诊",
        "名老中医工作室",
        "特需门诊",
    ]
    for department_name in department_list:
        print(department_name)
        analyzer = DepartmentAnalyzer(base_path, department_name)
        analyzer.analyze_all()


if __name__ == "__main__":
    main()
