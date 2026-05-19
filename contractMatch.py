import pandas as pd
import os
import re

def process_contract_match(target_file, master_file):
    print(f"[*] 正在进行合同编号匹配专区码: {target_file} (参考: {master_file})")
    try:
        # 1. 加载主表（专区管控表）
        # 使用第 0, 1, 2 列作为 合同编号, 专区码, 专区名称
        df_master = pd.read_excel(master_file, sheet_name='专区管控表', engine='calamine', usecols=[0, 1, 2])
        df_master.columns = ['合同编号', '专区码', '专区名称']
        
        # 建立映射列表 (保留原始字符串用于包含匹配)
        master_list = []
        for _, row in df_master.iterrows():
            raw_cid = str(row['合同编号']).strip()
            code = str(row['专区码']).strip() if pd.notna(row['专区码']) else ''
            name = str(row['专区名称']).strip() if pd.notna(row['专区名称']) else ''
            
            if raw_cid and raw_cid != 'nan' and raw_cid != 'None':
                master_list.append({
                    'cid': raw_cid,
                    'code': code if code not in ['nan', 'None'] else '',
                    'name': name if name not in ['nan', 'None'] else ''
                })

        # 2. 读取目标表
        xl = pd.ExcelFile(target_file, engine='calamine')
        sheets_data = {}
        for sheet in xl.sheet_names:
            df = pd.read_excel(target_file, sheet_name=sheet, engine='calamine')
            
            if '合同编号' not in df.columns:
                sheets_data[sheet] = df
                continue
            
            if '专区码' not in df.columns: df['专区码'] = pd.Series(dtype='object')
            else: df['专区码'] = df['专区码'].astype('object')
            if '专区' not in df.columns: df['专区'] = pd.Series(dtype='object')
            else: df['专区'] = df['专区'].astype('object')

            match_count = 0
            for idx, row in df.iterrows():
                # 只有当专区码和专区都为空时，尝试匹配
                has_code = pd.notna(df.at[idx, '专区码']) and str(df.at[idx, '专区码']).strip() not in ['', 'None', 'nan']
                has_name = pd.notna(df.at[idx, '专区']) and str(df.at[idx, '专区']).strip() not in ['', 'None', 'nan']
                
                if not has_code:
                    target_cid = str(row['合同编号']).strip()
                    if not target_cid or target_cid == 'nan' or target_cid == 'None':
                        continue
                        
                    for master_item in master_list:
                        if target_cid in master_item['cid']: 
                            if master_item['code']:
                                df.at[idx, '专区码'] = master_item['code']
                            elif master_item['name']:
                                # 如果没有专区码但有专区名称，直接填写名称
                                df.at[idx, '专区'] = master_item['name']
                            
                            match_count += 1
                            break
            
            print(f"  [+] 工作表 [{sheet}] 合同匹配成功 {match_count} 条专区码")
            sheets_data[sheet] = df

        # 3. 写回
        with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
            for sheet, data in sheets_data.items():
                data.to_excel(writer, sheet_name=sheet, index=False)
        print(f"[√] 合同编号匹配映射完成")
        
    except Exception as e:
        print(f"[X] 合同编号匹配执行失败: {e}")


if __name__ == "__main__":
    pass
