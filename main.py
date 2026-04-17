import pandas as pd
import numpy as np
import os
import re
from datetime import datetime

# ================= 配置区 =================
SOURCE_FILE = "raw_data.xlsx"  # 读取脚本A生成的原始文件名


# ==========================================

def extract_zone(text):
    """
    需求：提取第一个【】中的内容作为‘专区’。
    例如：‘【电子交易】XXX功能开发’ -> ‘电子交易’
    """
    if pd.isna(text):
        return None
    # 使用正则表达式匹配第一个【】内部的非贪婪内容
    match = re.search(r'【(.*?)】', str(text))
    return match.group(1) if match else None


def is_workday_exceeded(date_series, days_limit):
    """
    计算是否超过指定工作日（周一至周五）。
    基于 numpy 的 busday_count 实现，不计周六日。
    """
    today = datetime.now()

    def count_bus_days(start_date):
        if pd.isna(start_date):
            return 0
        try:
            # 计算从登记日期到今天的工作日天数
            return np.busday_count(start_date.date(), today.date())
        except Exception:
            return 0

    return date_series.apply(count_bus_days) > days_limit


def process_logic():
    # 1. 检查原始文件是否存在
    if not os.path.exists(SOURCE_FILE):
        print(f"错误: 找不到原始文件 '{SOURCE_FILE}'。")
        print("请先确保脚本 A (fetch_data.py) 运行成功并下载了数据。")
        return

    print(f"[*] 正在读取 {SOURCE_FILE}...")
    # 读取全量 Excel 数据
    df = pd.read_excel(SOURCE_FILE)

    # 2. 预处理数据
    print("[*] 正在预处理数据：提取‘专区’及转换日期...")
    # 提取专区列
    df['专区'] = df['需求名称'].apply(extract_zone)
    # 转换日期格式，解析失败的变为 NaT
    df['登记日期'] = pd.to_datetime(df['登记日期'], errors='coerce')
    # 剔除没有登记日期的脏数据
    df = df.dropna(subset=['登记日期'])

    # 3. 业务逻辑分析
    print("[*] 开始按照业务逻辑进行筛选分析...")

    # --- 逻辑 1：评审超时 ---
    # 条件：(评审状态为：需求设计/规划评审，或，开发状态为：等待任务分配) 并且 登记日期超过 5 工作日
    cond1 = (
                    (df['需求评审状态'].isin(['需求设计', '规划评审'])) |
                    (df['开发状态'] == '等待任务分配')
            ) & (is_workday_exceeded(df['登记日期'], 5))

    cols1 = ['专区', '需求编号', '需求名称', '产品名称', '登记人员', '登记日期', '需求评审状态']
    res1 = df[cond1][cols1]

    # --- 逻辑 2：开发超时 ---
    # 条件：评审完成 并且 (实际状态为：空 或 开发中) 并且 开发状态为：开发中 并且 超过 10 工作日
    cond2 = (
            (df['需求评审状态'] == '评审完成') &
            (df['需求实际状态'].isna() | (df['需求实际状态'] == '开发中')) &
            (df['开发状态'] == '开发中') &
            (is_workday_exceeded(df['登记日期'], 10))
    )

    cols2 = ['专区', '需求编号', '需求名称', '产品名称', '登记人员', '登记日期', '需求评审状态', '需求责任人',
             '开发状态', '开发工作量评审']
    res2 = df[cond2][cols2]

    # --- 逻辑 3：需求上线超期 ---
    # 条件：评审完成
    # 且 开发状态 不等于 (开发中/已上线/待任务分配/作废/终止/设计评审/已完成)
    # 且 需求实际状态 不等于 (已上线/作废/暂停)
    # 且 登记日期超过 15 工作日
    exclude_dev = ['开发中', '已上线', '待任务分配', '作废', '终止', '设计评审', '已完成']
    exclude_actual = ['已上线', '作废', '暂停']

    cond3 = (
            (df['需求评审状态'] == '评审完成') &
            (~df['开发状态'].isin(exclude_dev)) &
            (~df['需求实际状态'].isin(exclude_actual)) &
            (is_workday_exceeded(df['登记日期'], 15))
    )

    cols3 = ['专区', '需求编号', '需求名称', '产品名称', '登记人员', '登记日期', '需求评审状态', '需求责任人',
             '开发状态', '开发工作量评审', '需求实际状态']
    res3 = df[cond3][cols3]

    # 4. 结果导出
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    output_filename = f"需求分析结果_{timestamp}.xlsx"

    print(f"[*] 正在生成结果报表: {output_filename}")
    with pd.ExcelWriter(output_filename) as writer:
        res1.to_excel(writer, sheet_name="1_评审超时", index=False)
        res2.to_excel(writer, sheet_name="2_开发超时", index=False)
        res3.to_excel(writer, sheet_name="3_上线超期", index=False)

    print("\n" + "=" * 40)
    print("【分析完成】")
    print(f"原始文件: {SOURCE_FILE}")
    print(f"分析结果: {output_filename}")
    print(f" - 1_评审超时: {len(res1)} 条")
    print(f" - 2_开发超时: {len(res2)} 条")
    print(f" - 3_上线超期: {len(res3)} 条")
    print("=" * 40)


if __name__ == "__main__":
    process_logic()