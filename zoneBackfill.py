import pandas as pd
import os

def process_zone_backfill(target_file, master_file):
    print(f"[*] 正在进行专区信息最终回填: {target_file} (参考: {master_file})")
    try:
        # 1. 加载主表（专区管控表）
        # 列索引参考：1: 专区码, 2: 专区名称, 5: 分公司/省份
        df_master = pd.read_excel(master_file, sheet_name='专区管控表', engine='calamine')
        
        code_map = {} # code -> (name, company)
        valid_names = set() # valid names from master
        for _, row in df_master.iterrows():
            code = str(row.iloc[1]).strip()
            name = str(row.iloc[2]).strip()
            comp = str(row.iloc[5]).strip()
            if code and code != 'nan' and code != 'None':
                code_map[code] = (name, comp)
            if name and name != 'nan' and name != 'None':
                valid_names.add(name)

        if not code_map and not valid_names:
            print("  [!] 未能建立有效的主表映射，跳过。")
            return

        # 2. 读取目标表
        xl = pd.ExcelFile(target_file, engine='calamine')
        sheets_data = {}
        for sheet in xl.sheet_names:
            df = pd.read_excel(target_file, sheet_name=sheet, engine='calamine')
            
            if '专区码' not in df.columns:
                sheets_data[sheet] = df
                continue
            
            if '专区码' not in df.columns: df['专区码'] = pd.Series(dtype='object')
            else: df['专区码'] = df['专区码'].astype('object')
            if '专区' not in df.columns: df['专区'] = pd.Series(dtype='object')
            else: df['专区'] = df['专区'].astype('object')
            if '分公司' not in df.columns: df['分公司'] = pd.Series(dtype='object')
            else: df['分公司'] = df['分公司'].astype('object')

            match_count = 0
            for idx, row in df.iterrows():
                code = str(row['专区码']).strip() if pd.notna(row['专区码']) else ''
                current_name = str(df.at[idx, '专区']).strip() if pd.notna(df.at[idx, '专区']) else ''
                
                # 情况 A：有专区码 -> 直接以专区码对应的官网名称为准，覆盖掉括号提取的数据
                if code and code in code_map:
                    name, comp = code_map[code]
                    df.at[idx, '专区'] = name
                    df.at[idx, '分公司'] = comp
                    match_count += 1
                else:
                    # 情况 B：没有专区码 -> 检查当前填写的“专区”是否是主表中的正规名称
                    # 如果当前名称不在正规名单里（说明它只是括号里提取的关键词，且匹配失败了），则清空
                    if current_name and current_name not in valid_names:
                        df.at[idx, '专区'] = None
                        df.at[idx, '分公司'] = None
            
            # 调整列顺序：将“专区码”移到第二列 (索引 1)
            cols = list(df.columns)
            if '专区码' in cols:
                cols.insert(1, cols.pop(cols.index('专区码')))
                df = df[cols]

            print(f"  [+] 工作表 [{sheet}] 成功回填 {match_count} 条记录")
            sheets_data[sheet] = df

        # 3. 写回
        with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
            for sheet, data in sheets_data.items():
                data.to_excel(writer, sheet_name=sheet, index=False)
        print(f"[√] 专区及分公司信息回填完成")
        
    except Exception as e:
        print(f"[X] 专区信息回填执行失败: {e}")

if __name__ == "__main__":
    pass
