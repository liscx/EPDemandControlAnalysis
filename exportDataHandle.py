import pandas as pd
import os
import re
from datetime import datetime, date, timedelta
from chinese_calendar import is_workday

def extract_zone(text):
    if pd.isna(text):
        return None
    match = re.search(r'【(.*)】', str(text))
    if match:
        return re.sub(r'[【】]', '', match.group(1))
    return None

def is_workday_exceeded(date_series, days_limit):
    today = datetime.now().date()
    def count_bus_days(start_date):
        if pd.isna(start_date):
            return 0
        try:
            current = start_date.date()
            count = 0
            while current < today:
                current += timedelta(days=1)
                if is_workday(current):
                    count += 1
            return count
        except Exception:
            return None
    result = date_series.apply(count_bus_days)
    return result.notna() & (result > days_limit)

def is_calendar_exceeded(date_series, days_limit):
    today = datetime.now().date()
    def count_cal_days(start_date):
        if pd.isna(start_date):
            return 0
        try:
            return (today - start_date.date()).days
        except Exception:
            return 0
    return date_series.apply(count_cal_days) > days_limit

def process_logic(source_file="raw_data.xlsx", output_dir=None, timestamp=None):
    if not os.path.exists(source_file):
        print(f"错误: 找不到原始文件 '{source_file}'。")
        return None

    if not timestamp:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    
    if not output_dir:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(base_dir, "result")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    print(f"[*] 正在处理原始数据: {source_file}")
    df = pd.read_excel(source_file)

    # 预处理
    df['专区'] = pd.Series(dtype='object')
    df['登记日期'] = pd.to_datetime(df['登记日期'], errors='coerce')
    df = df.dropna(subset=['登记日期'])
    df = df[df['登记日期'] >= '2026-01-01']

    # 逻辑分析
    cond_review = (df['需求评审状态'].isin(['需求设计', '规划评审'])) | (df['开发状态'] == '等待任务分配')
    cond_not_online = df['开发状态'] != '已上线'
    cond1 = cond_review & cond_not_online & (is_workday_exceeded(df['登记日期'], 5) | is_calendar_exceeded(df['登记日期'], 7))
    res1 = df[cond1][['专区', '合同编号', '需求编号', '需求名称', '产品名称', '登记人员', '登记日期', '需求评审状态']]

    cond2 = ((df['需求评审状态'] == '评审完成') & (df['需求实际状态'].isna() | (df['需求实际状态'] == '开发中')) & (df['开发状态'] == '开发中') & (is_workday_exceeded(df['登记日期'], 10)))
    res2 = df[cond2][['专区', '合同编号', '需求编号', '需求名称', '产品名称', '登记人员', '登记日期', '需求评审状态', '需求责任人', '开发状态', '开发工作量评审']]

    exclude_dev = ['开发中', '已上线', '待任务分配', '作废', '终止', '设计评审', '已完成']
    exclude_actual = ['已上线', '作废', '暂停']
    cond3 = ((df['需求评审状态'] == '评审完成') & (~df['开发状态'].isin(exclude_dev)) & (~df['需求实际状态'].isin(exclude_actual)) & (is_workday_exceeded(df['登记日期'], 15)))
    res3 = df[cond3][['专区', '合同编号', '需求编号', '需求名称', '产品名称', '登记人员', '登记日期', '需求评审状态', '需求责任人', '开发状态', '开发工作量评审', '需求实际状态']]

    output_filename = f"需求分析结果_{timestamp}.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    with pd.ExcelWriter(output_path) as writer:
        res1.to_excel(writer, sheet_name="1_评审超时", index=False)
        res2.to_excel(writer, sheet_name="2_开发超时", index=False)
        res3.to_excel(writer, sheet_name="3_上线超期", index=False)

    print(f"【分析完成】结果保存至: {output_path}")
    return output_path

if __name__ == "__main__":
    pass