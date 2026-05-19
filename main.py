import pandas as pd
from datetime import datetime
import sys
import os


def export_high_risk_overdue(file_path):
    # --- 配置区 ---
    COL_STATUS = '需求实际状态'
    COL_TIME = '期望完成时间'
    EXCLUDE_LIST = ['已上线', '待更新正式', '暂停', '作废', '待更新仿真']
    # --------------

    # 处理路径：去除用户可能误打的引号，并转换为绝对路径
    clean_path = os.path.abspath(file_path.strip('"'))

    if not os.path.exists(clean_path):
        print(f"【错误】找不到文件，请检查路径是否正确：")
        print(f"输入路径: {clean_path}")
        return

    try:
        print(f"正在读取文件: {clean_path}")
        df = pd.read_excel(clean_path, engine='openpyxl')
    except Exception as e:
        print(f"【读取失败】: {e}")
        return

    # 1. 基础校验
    if COL_STATUS not in df.columns or COL_TIME not in df.columns:
        print(f"【列名错误】Excel 缺少必要字段。")
        print(f"当前字段有: {list(df.columns)}")
        return

    # 2. 转换日期格式
    df[COL_TIME] = pd.to_datetime(df[COL_TIME], errors='coerce')
    now = datetime.now()

    # 3. 设定过滤条件
    mask_status = ~df[COL_STATUS].astype(str).str.strip().isin(EXCLUDE_LIST)
    mask_overdue = (now - df[COL_TIME]).dt.days > 10

    # 4. 提取数据
    risk_df = df[mask_status & mask_overdue].copy()
    risk_df['当前逾期天数'] = (now - risk_df[COL_TIME]).dt.days

    # 5. 导出结果到脚本同级目录
    if not risk_df.empty:
        # 生成带时间戳的文件名
        output_name = f"严重超期明细_{now.strftime('%Y%m%d_%H%M')}.xlsx"
        # 获取脚本所在的文件夹路径，确保结果保存在脚本旁边
        output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), output_name)

        try:
            risk_df.to_excel(output_path, index=False, engine='openpyxl')
            print("-" * 40)
            print(f"分析完成！找到 {len(risk_df)} 条异常需求。")
            print(f"结果已保存至: {output_path}")
            print("-" * 40)
        except Exception as e:
            print(f"导出失败: {e}")
    else:
        print("-" * 40)
        print("未发现符合条件的超期需求。")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # 获取命令行传入的全量路径
        input_argument = " ".join(sys.argv[1:])  # 处理路径中带空格但不加引号的情况
        export_high_risk_overdue(input_argument)
    else:
        # 默认文件名
        export_high_risk_overdue('新点e交易专区需求管控表_限额saas.xlsx')