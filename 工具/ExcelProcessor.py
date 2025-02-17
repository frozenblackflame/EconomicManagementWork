import os
import pandas as pd
import json
from collections import OrderedDict

class ExcelProcessor:
    @staticmethod
    def get_default_output_path():
        """
        获取默认的JSON报告输出路径。

        :return: 默认的JSON报告路径（用户桌面）
        """
        return os.path.join(os.path.expanduser("~"), "Desktop", "指标统计结果.json")

    def __init__(self, folder_path, departments, metrics=None, output_path=None):
        """
        初始化ExcelProcessor类。

        :param folder_path: Excel文件所在的目录路径
        :param departments: 科室列表
        :param metrics: 需要统计的指标列表，如果为None则读取所有指标
        :param output_path: 最终JSON报告的输出路径
        :param no_data_output_path: 无数据科室统计报告的输出路径
        :param simple_output_path: 简化格式报告的输出路径
        """
        self.folder_path = folder_path
        self.departments = departments
        self.metrics = metrics  # 现在可以为None
        self.output_path = output_path or self.get_default_output_path()
        self.results = {}  # 将在process_files中初始化
        self.no_data_report = OrderedDict()
        self.simple_report = OrderedDict()
        self.processed_files = []

    def process_files(self):
        """
        遍历目录下的所有Excel文件并处理数据。
        """
        print("\n=== 开始处理数据 ===")
        print(f"将处理以下科室的数据：{', '.join(self.departments)}")
        
        # 处理第一个文件以获取所有指标
        first_file = None
        for filename in os.listdir(self.folder_path):
            if filename.endswith('.xlsx') and '年' in filename and '月' in filename:
                first_file = os.path.join(self.folder_path, filename)
                break
        
        if first_file:
            df = pd.read_excel(first_file, header=None)
            # 不再删除B列
            
            header_row = df.iloc[3]
            new_headers = []
            current_header = None
            for h in header_row:
                if pd.notna(h):
                    current_header = h
                new_headers.append(current_header)
            
            # 添加"合计得分"到指标列表
            if self.metrics is None:
                self.metrics = list(set(h for h in new_headers if pd.notna(h) and h != '科室'))
                self.metrics.append("合计得分")

        print(f"将统计以下指标：{', '.join(self.metrics)}\n")
        
        # 初始化results字典
        self.results = {dept: {metric: {} for metric in self.metrics} for dept in self.departments}

        for filename in os.listdir(self.folder_path):
            if filename.endswith('.xlsx') and '年' in filename and '月' in filename:
                file_path = os.path.join(self.folder_path, filename)
                print(f"\n正在处理文件：{filename}")
                self.processed_files.append(filename)

                try:
                    is_november = '11月' in filename or '12月' in filename
                    df = pd.read_excel(file_path, header=None)
                    
                    header_row = df.iloc[3]
                    new_headers = []
                    current_header = None
                    for h in header_row:
                        if pd.notna(h):
                            current_header = h
                        new_headers.append(current_header)
                    
                    df.columns = new_headers
                    df.columns.values[0] = '科室'
                    df = df.iloc[5:]

                    month = self.extract_month(filename)
                    if not month:
                        continue

                    self.aggregate_data(df, is_november, month)
                    
                except Exception as e:
                    print(f"\n错误：处理文件 {filename} 时出错")
                    print(f"错误详情：{str(e)}")
                    print(f"出错位置：", e.__traceback__.tb_lineno)

    def extract_month(self, filename):
        """
        从文件名中提取月份。

        :param filename: 文件名
        :return: 月份字符串，格式为两位数字
        """
        parts = filename.split('年')
        if len(parts) == 2:
            month_part = parts[1].split('月')[0]
            try:
                month = str(int(month_part)).zfill(2)
                print(f"提取到月份：{month}")
                return month
            except ValueError:
                print(f"警告：无法从文件名 {filename} 中提取月份")
                return None
        else:
            print(f"警告：文件名 {filename} 格式不正确")
            return None

    def aggregate_data(self, df, is_november, month):
        """
        对每个科室进行数据统计。

        :param df: DataFrame对象
        :param is_november: 是否为11月或12月的数据
        :param month: 当前处理的月份
        """
        for dept in self.departments:
            dept_data = df[df['科室'] == dept]
            if not dept_data.empty:
                print(f"\n科室：{dept}")
                
                # 处理合计得分（B列）
                try:
                    total_score = dept_data.iloc[0, 1]  # B列的索引是1
                    if pd.notna(total_score):
                        total_score = float(total_score)
                        self.results[dept]["合计得分"][month] = total_score
                        print(f"  - 合计得分: {total_score:.4f}")
                except Exception as e:
                    print(f"  - 合计得分处理出错: {str(e)}")
                
                # 处理其他指标
                for metric in self.metrics:
                    if metric != "合计得分" and metric in df.columns:
                        try:
                            metric_idx = list(df.columns).index(metric)
                            if not is_november:
                                metric_idx += 2
                            
                            if metric_idx < len(df.columns):
                                value = dept_data.iloc[0, metric_idx]
                                if pd.notna(value):
                                    value = float(value)
                                    self.results[dept][metric][month] = value
                                    print(f"  - {metric}: {value:.4f}")
                                else:
                                    print(f"  - {metric}: 数据为空")
                            else:
                                print(f"  - {metric}: 列索引超出范围")
                        except Exception as e:
                            print(f"  - {metric}: 处理出错 - {str(e)}")
            else:
                print(f"\n警告：未找到科室 {dept} 的数据")

    def generate_reports(self):
        """
        生成最终的JSON报告。
        """
        final_report = OrderedDict()
        
        for dept in self.departments:
            dept_data = OrderedDict()
            no_data_metrics = OrderedDict()
            
            for metric in self.metrics:
                metric_data = OrderedDict()
                monthly_data = self.results[dept][metric]
                sorted_months = sorted(monthly_data.keys())
                
                metric_data["monthly_data"] = OrderedDict()
                for month in sorted_months:
                    metric_data["monthly_data"][f"{month}月"] = round(monthly_data[month], 4)
                
                values = list(monthly_data.values())
                if values:
                    metric_data["statistics"] = {
                        "平均值": round(sum(values) / len(values), 4),
                        "数据月份数": len(values),
                        "总值": round(sum(values), 4)
                    }
                else:
                    no_data_metrics[metric] = {
                        "状态": "无数据",
                        "说明": "未找到任何月份的数据"
                    }
                
                dept_data[metric] = metric_data
            
            final_report[dept] = dept_data
            if no_data_metrics:
                self.no_data_report[dept] = no_data_metrics
        
        self.save_json(final_report, self.output_path, "结果已保存到")
        self.print_reports(final_report)

    def save_json(self, data, path, message):
        """
        将数据保存为JSON文件。

        :param data: 需要保存的数据
        :param path: 保存路径
        :param message: 保存成功后的提示信息
        """
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"\n{message}：{path}")
        except Exception as e:
            print(f"\n保存JSON文件时出错：{str(e)}")

    def print_reports(self, final_report):
        """
        将报告打印到控制台。

        :param final_report: 最终的统计报告
        """
        print("\n=== 统计结果 ===")
        print(json.dumps(final_report, ensure_ascii=False, indent=2))

    def run(self):
        """
        运行整个处理流程。
        """
        self.process_files()
        self.generate_reports()


if __name__ == "__main__":
    departments=[
        "胃肿瘤外科", "内窥镜中心", "头颈一科", "口腔科", 
        "肝胆胰肿瘤外科", "头颈二科", "皮肤科",
        "泌尿外科", "病理科", "妇瘤一科", "便民门诊",
        "结直肠肿瘤外科","功能科", "妇瘤二科",
        "针灸理疗室","乳腺内科", "核医学科",
        "骨与软组织一科", "重症医学科", "放射治疗科", "检验科", "骨与软组织二科",
        "麻醉手术科", "中西医结合科", "放射科", "乳腺一科",
        "血液病科", "乳腺二科", "消化肿瘤内一科",
        "胸外一科", "消化肿瘤内二科", "门诊部",
        "胸外二科", "介入治疗科", "呼吸肿瘤内科", "眼科"
    ]
    
    folder_path = r"C:\Users\biyun\Desktop\work\2024年绩效"
    
    processor = ExcelProcessor(
        folder_path=folder_path,
        departments=departments,
        # metrics参数不再需要指定
    )
    processor.run() 