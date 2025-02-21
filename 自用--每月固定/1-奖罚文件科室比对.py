import tkinter as tk
from tkinter import scrolledtext, messagebox

def compare_departments():
    # 获取已有科室数据
    existing_departments = [
        "行政单位系数", "行政绩效人员", "院办公室", "人事科", "财务科", "党委办公室", 
        "纪检监察室", "科教科", "医务科", "护理部", "感染管理科", "医疗保险办公室", 
        "审计科", "宣传科", "后勤保障科", "保卫科", "网络中心", "动力管理科", 
        "医学工程部", "采购办公室", "经济管理办公室", "健康促进科", "对外联络办公室", 
        "公共卫生科", "临床营养科", "收费管理科", "工会办公室", "患者维权办公室", 
        "肿瘤流行病学研究中心", "医疗质量控制科", "基建办", "药学部（职能）", "财务部", 
        "医务部", "头颈一科", "头颈二科", "妇瘤一科", "妇瘤二科", "骨与软组织一科", 
        "骨与软组织二科", "乳腺一科", "乳腺二科", "胸外一科", "胸外二科", "胃肿瘤外科", 
        "肝胆胰肿瘤外科", "泌尿外科", "结直肠肿瘤外科", "乳腺内科", "放疗一病区", 
        "放疗二病区", "放疗三病区", 
        "中西医结合科", "血液病科", "消化肿瘤内一科", "消化肿瘤内二科", "呼吸肿瘤内科", 
        "重症医学科", "麻醉手术科", "麻醉手术科护士", "放疗机房", "介入治疗科", 
        "门诊收费室", "住院收费室", "内窥镜中心", "消毒供应中心", "病理科", 
        "功能科", "核医学科", "检验科", "放射科", "药学部", "胸外一科护理", 
        "胸外二、介入治疗科护理", "胃肿瘤外科护理", "中西医、乳腺内科护理", 
        "放疗一病区护理","放疗二病区护理","放疗三病区护理",
        "血液病科护理", "消化肿瘤内一科护理", "消化肿瘤内二科护理", 
        "呼吸肿瘤内科护理", "头颈一科护理", "头颈二科护理", "妇瘤一科护理", 
        "骨与软组织一科护理", "乳腺一科护理", "乳腺二科护理", "结直肠肿瘤外科护理", 
        "肝胆胰、泌尿外科护理部", "骨软二、妇瘤二科护理部", "门诊部", "口腔科", 
        "皮肤科", "便民门诊", "针灸理疗室"
    ]

    # 获取用户输入的奖惩科室数据
    reward_penalty_input = reward_penalty_text.get("1.0", tk.END).strip()
    reward_penalty_departments = [line for line in reward_penalty_input.split('\n') if line]
    # 将输入的科室数据去重
    reward_penalty_departments = list(set(reward_penalty_departments))

    # 找出奖惩科室中有而已有科室中没有的数据
    result = [department for department in reward_penalty_departments if department not in existing_departments]

    # 在输出框中显示结果
    output_text.delete("1.0", tk.END)
    if result:
        output_text.insert(tk.END, "\n".join(result))
    else:
        output_text.insert(tk.END, "所有科室都已存在于已有科室中。")

# 创建主窗口
root = tk.Tk()
root.title("科室比对工具")

# 创建输入框
reward_penalty_label = tk.Label(root, text="请输入奖惩科室（每行一个）：")
reward_penalty_label.pack()

reward_penalty_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=50, height=15)
reward_penalty_text.pack()

# 创建比对按钮
compare_button = tk.Button(root, text="比对", command=compare_departments)
compare_button.pack()

# 创建输出框
output_label = tk.Label(root, text="比对结果：")
output_label.pack()

output_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=50, height=15)
output_text.pack()

# 运行主循环
root.mainloop()
