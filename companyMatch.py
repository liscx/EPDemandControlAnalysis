import pandas as pd
from difflib import SequenceMatcher

def similar(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def process_company_match(target_file, master_file):
    print(f"[*] 正在匹配分公司信息: {target_file} (参考: {master_file})")
    try:
        # 1. 加载主表（专区管控表）
        df_master = pd.read_excel(master_file, sheet_name='专区管控表', engine='calamine')
        # 假设：B列(1)是专区码，C列(2)是专区名，F列(5)是分公司
        master_data = []
        for _, row in df_master.iterrows():
            master_data.append({
                'code': str(row.iloc[1]).strip(),
                'name': str(row.iloc[2]).strip(),
                'company': str(row.iloc[5]).strip()
            })

        # 2. 读取目标表
        xl = pd.ExcelFile(target_file, engine='calamine')
        sheets_data = {}
        for sheet in xl.sheet_names:
            df = pd.read_excel(target_file, sheet_name=sheet, engine='calamine')
            if '专区' not in df.columns:
                sheets_data[sheet] = df
                continue
            
            if '分公司' not in df.columns:
                df['分公司'] = None
            if '专区码' not in df.columns:
                df['专区码'] = None

            for idx, row in df.iterrows():
                zone_name = str(row['专区']).strip()
                if not zone_name or zone_name == 'None':
                    continue
                
                best_match = None
                max_ratio = 0
                
                for master in master_data:
                    # 规则：如果主表名字包含在目标专区名中，或者反之，或者相似度极高
                    ratio = similar(zone_name, master['name'])
                    if master['name'] in zone_name or zone_name in master['name'] or ratio > 0.8:
                        if ratio > max_ratio:
                            max_ratio = ratio
                            best_match = master
                
                if best_match:
                    df.at[idx, '分公司'] = best_match['company']
                    df.at[idx, '专区码'] = best_match['code']

            sheets_data[sheet] = df

        # 3. 写回
        with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
            for sheet, data in sheets_data.items():
                data.to_excel(writer, sheet_name=sheet, index=False)
        print(f"[√] 分公司匹配完成")
    except Exception as e:
        print(f"[X] 分公司匹配执行失败: {e}")

if __name__ == "__main__":
    pass