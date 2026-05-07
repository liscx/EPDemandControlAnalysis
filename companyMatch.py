import pandas as pd
from difflib import SequenceMatcher

def similar(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def process_company_match(target_file, master_file):
    print(f"[*] 正在进行专区名称模糊匹配专区码: {target_file} (参考: {master_file})")
    try:
        # 1. 加载主表（针对名称到码的映射）
        # B(1): 专区码, C(2): 专区名称
        df_master = pd.read_excel(master_file, sheet_name='专区管控表', engine='calamine', usecols=[1, 2])
        df_master.columns = ['专区码', '专区名称']
        
        master_data = []
        for _, row in df_master.iterrows():
            master_data.append({
                'code': str(row['专区码']).strip(),
                'name': str(row['专区名称']).strip()
            })

        # 2. 读取目标表
        xl = pd.ExcelFile(target_file, engine='calamine')
        sheets_data = {}
        for sheet in xl.sheet_names:
            df = pd.read_excel(target_file, sheet_name=sheet, engine='calamine')
            
            if '专区' not in df.columns:
                sheets_data[sheet] = df
                continue
            
            if '专区码' not in df.columns:
                df['专区码'] = None

            match_count = 0
            for idx, row in df.iterrows():
                # 只有当专区码为空时，尝试通过“专区”列的名称进行模糊匹配
                if pd.isna(df.at[idx, '专区码']) or str(df.at[idx, '专区码']).strip() in ['', 'None', 'nan']:
                    zone_name = str(row['专区']).strip()
                    if not zone_name or zone_name == 'None':
                        continue
                    
                    best_match_code = None
                    max_ratio = 0
                    
                    for master in master_data:
                        # 简单的匹配逻辑：包含关系或相似度
                        ratio = similar(zone_name, master['name'])
                        if master['name'] in zone_name or zone_name in master['name'] or ratio > 0.8:
                            if ratio > max_ratio:
                                max_ratio = ratio
                                best_match_code = master['code']
                    
                    if best_match_code:
                        df.at[idx, '专区码'] = best_match_code
                        match_count += 1

            print(f"  [+] 工作表 [{sheet}] 模糊匹配成功 {match_count} 条专区码")
            sheets_data[sheet] = df

        # 3. 写回
        with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
            for sheet, data in sheets_data.items():
                data.to_excel(writer, sheet_name=sheet, index=False)
        print(f"[√] 名称模糊匹配映射完成")
    except Exception as e:
        print(f"[X] 名称模糊匹配执行失败: {e}")

if __name__ == "__main__":
    pass