import pandas as pd
import os


def final_solution():
    # --- 文件名定义 ---
    master_file = '01-新点电子交易专区&项目跟进表（重要）.xlsx'
    analysis_file = '需求分析结果_20260417_1058.xlsx'
    output_file = '需求分析结果_最终关联版.xlsx'

    print("Step 1: 建立分公司映射表...")
    try:
        # 使用 calamine 引擎，它只读数据不读样式，完美避开 'Fill' 报错
        # 如果没安装 calamine，请将 engine 改回 'openpyxl'
        df_master = pd.read_excel(master_file, sheet_name='专区管控表', engine='calamine')

        # 01表的结构：第4列(索引3)是专区名称，第6列(索引5)是分公司
        # 使用 iloc 绝对定位列，无视表头文字是否包含换行/空格
        mapping_df = df_master.iloc[:, [3, 5]].copy()
        mapping_df.columns = ['Zone', 'Branch']

        # 清洗：去掉空行，去掉名称前后的空格
        mapping_df = mapping_df.dropna(subset=['Zone']).drop_duplicates()
        mapping_df['Zone'] = mapping_df['Zone'].astype(str).str.strip()

        # 转换为字典提高查询速度
        mapping_dict = dict(zip(mapping_df['Zone'], mapping_df['Branch']))
        print(f"   [成功] 映射建立完成，包含 {len(mapping_dict)} 个专区。")
    except Exception as e:
        print(f"   [失败] 读取基准表时出错，请确保文件未被占用且已另存为过: {e}")
        return

    print(f"\nStep 2: 处理需求分析文件...")
    try:
        # 读取源文件所有 Sheet (评审超时、开发超时、上线超期)
        source_dict = pd.read_excel(analysis_file, sheet_name=None, engine='calamine')

        # 使用 ExcelWriter 统一写入，解决 "At least one sheet" 报错
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in source_dict.items():
                print(f"   正在处理页签: {sheet_name}...")

                # 检查是否存在“专区”这一列（有些 Sheet 列名可能就是“专区”）
                zone_col = None
                if '专区' in df.columns:
                    zone_col = '专区'
                elif '专区名称' in df.columns:
                    zone_col = '专区名称'

                if zone_col:
                    # 匹配分公司
                    df[zone_col] = df[zone_col].astype(str).str.strip()
                    df['分公司'] = df[zone_col].map(mapping_dict)

                    # 移动“分公司”到第一列，方便查看
                    cols = list(df.columns)
                    if '分公司' in cols:
                        cols.insert(0, cols.pop(cols.index('分公司')))
                        df = df[cols]
                else:
                    print(f"      - 警告: 页签 {sheet_name} 未找到“专区”列，跳过匹配直接保留。")

                # 强制回写 Sheet
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"\n[大功告成] 文件已生成: {output_file}")

    except Exception as e:
        print(f"   [严重错误] 处理过程中发生紊乱: {e}")


if __name__ == "__main__":
    final_solution()