# wechat_analyzer.py
import os
import sys
import re
import pandas as pd
from datetime import datetime
import webbrowser
from threading import Timer

def classify_activity(title, summary):
    """活动类型分类"""
    text = (str(title) + " " + str(summary)).lower()
    
    # 定义关键词规则（可扩展）
    categories = {
        "联谊交流类": ["年会", "春茗", "返校", "聚会", "reunion", "班级", "迎新", "团建", "联谊"],
        "企业参访类": ["走进", "参访", "走访", "考察", "调研", "园区", "企业", "工厂", "政府"],
        "论坛讲座类": ["论坛", "讲座", "沙龙", "研讨会", "峰会", "分享会", "对话", "公开课"],
        "公益服务类": ["公益", "捐赠", "支教", "环保", "慈善", "志愿服务", "社区服务", "助学"],
        "跨界合作类": ["联合", "合作", "共建", "携手", "联盟", "商会", "高校", "兄弟校友会"],
        "内部治理类": ["换届", "理事会", "章程", "选举", "工作会议", "制度", "组织架构"],
        "学术研究类": ["研究", "课题", "白皮书", "报告发布", "学术", "智库"],  # 新增
        "文体赛事类": ["比赛", "运动会", "文艺", "演出", "摄影", "书画", "体育"],  # 新增
    }
    
    for cat, keywords in categories.items():
        if any(kw in text for kw in keywords):
            return cat
    return "其他未分类"

def extract_date(pub_time, summary):
    """从发布时间或摘要中提取最早有效日期，以摘要为准"""
    # 尝试从摘要中找日期（如“2024年3月活动”）
    summary_date = None
    date_match = re.search(r'(\d{4})[年\-](\d{1,2})', str(summary))
    if date_match:
        year, month = int(date_match.group(1)), int(date_match.group(2))
        if 2020 <= year <= 2026 and 1 <= month <= 12:
            summary_date = datetime(year, month, 1)
    
    # 解析发布时间
    pub_date = None
    try:
        pub_date = pd.to_datetime(pub_time)
    except:
        pass
    
    # 优先用摘要日期，否则用发布日期
    final_date = summary_date or pub_date
    return final_date

def assign_branch(title, summary, official_account):
    """判断文章归属的分支机构"""
    text = (str(title) + " " + str(summary) + " " + str(official_account)).lower()
    
    # 明确归属
    if "广东" in text or "emba广东" in text:
        return "光华EMBA广东校友会"
    if "华南" in text and "广东" not in text:  # 避免广东被误判为华南
        return "光华EMBA华南校友会"
    if "香港" in text:
        return "北京大學光華管理學院香港校友會"
    if "华东" in text or ("上海" in text and ("华东" in text or "校友会" in text)):
        return "北大光华华东校友会"
    
    # 总部文章中涉及分支机构的识别
    if official_account.strip() == "北大光华校友会":
        if any(kw in text for kw in ["广东", "华南", "香港", "华东", "上海校友"]):
            # 需人工复核，但先尝试分配
            if "广东" in text: return "光华EMBA广东校友会"
            if "华南" in text: return "光华EMBA华南校友会"
            if "香港" in text: return "北京大學光華管理學院香港校友會"
            if "华东" in text or "上海" in text: return "北大光华华东校友会"
    
    return "北大光华校友会总部"  # 默认归属总部

def main_analysis(input_path):
    """主分析流程"""
    df = pd.read_excel(input_path)
    
    # 必需字段检查
    required_cols = ["标题", "摘要", "发布时间", "作者", "文章链接"]
    if not all(col in df.columns for col in required_cols):
        raise ValueError(f"Excel 缺少必要列！请包含：{required_cols}")
    
    # 清洗数据
    df = df.dropna(subset=["标题"]).copy()
    df["发布时间"] = pd.to_datetime(df["发布时间"], errors="coerce")
    
    # 新增分析列
    df["活动类型"] = df.apply(lambda x: classify_activity(x["标题"], x["摘要"]), axis=1)
    df["有效日期"] = df.apply(lambda x: extract_date(x["发布时间"], x["摘要"]), axis=1)
    df["分支机构"] = df.apply(lambda x: assign_branch(x["标题"], x["摘要"], x["作者"]), axis=1)
    
    # 时间范围过滤：2023-01 至 2025-10
    start = datetime(2023, 1, 1)
    end = datetime(2025, 10, 31)
    df = df[(df["有效日期"] >= start) & (df["有效日期"] <= end)]
    
    # 生成统计报表
    branch_summary = df.groupby("分支机构")["活动类型"].value_counts().unstack(fill_value=0)
    type_summary = df["活动类型"].value_counts()
    time_trend = df.groupby(df["有效日期"].dt.to_period("M")).size()
    
    # 保存结果
    output_path = os.path.splitext(input_path)[0] + "_分析结果.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="原始数据清洗版", index=False)
        branch_summary.to_excel(writer, sheet_name="各分会活动对比")
        type_summary.to_excel(writer, sheet_name="活动类型统计")
        time_trend.to_excel(writer, sheet_name="时间趋势分析")
    
    return output_path

# ============ GUI 部分（打包用）============
import tkinter as tk
from tkinter import filedialog, messagebox

def run_gui():
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    file_path = filedialog.askopenfilename(
        title="选择公众号文章汇总 Excel 文件",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    if not file_path:
        return
    
    try:
        output_file = main_analysis(file_path)
        messagebox.showinfo("成功", f"分析完成！\n结果已保存至：\n{output_file}")
        os.startfile(os.path.dirname(output_file))  # 打开文件夹
    except Exception as e:
        messagebox.showerror("错误", f"分析失败：\n{str(e)}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # 命令行模式（调试用）
        main_analysis(sys.argv[1])
    else:
        # 图形界面模式（用户友好）
        run_gui()
