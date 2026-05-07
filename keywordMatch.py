import pandas as pd
import math
import os

def process_keyword_match(target_file, master_file, kw_file='映射表.xlsx'):
    print(f"[*] 正在通过关键词映射匹配专区码: {target_file}")
    
    # 1. 读关键字映射表
    try:
        df_kw = pd.read_excel(kw_file, usecols=[0, 1])
        df_kw.columns = ['关键字', '专区码']
        df_kw = df_kw.dropna(subset=['关键字'])
        kw_dict = {}
        for _, row in df_kw.iterrows():
            kw = str(row['关键字']).strip()
            code = row['专区码']
            if isinstance(code, float) and not math.isnan(code):
                code = int(code)
            kw_dict[kw] = str(code).strip()
    except Exception as e:
        print(f"警告：无法读取映射表 {kw_file}: {e}")
        return

    # 2. 处理目标表
    try:
        xl = pd.ExcelFile(target_file, engine='calamine')
        sheets_data = {}
        for sheet in xl.sheet_names:
            df = pd.read_excel(target_file, sheet_name=sheet, engine='calamine')
            
            if '专区码' not in df.columns:
                df['专区码'] = None

            match_count = 0
            for idx, row in df.iterrows():
                # 只有当专区码仍然为空时，尝试通过关键词映射
                val_code = str(df.at[idx, '专区码']).strip() if pd.notna(df.at[idx, '专区码']) else ''
                if val_code in ['', 'None', 'nan']:
                    req_name = str(row.get('需求名称', ''))
                    for kw, code_str in kw_dict.items():
                        if kw in req_name:
                            df.at[idx, '专区码'] = code_str
                            match_count += 1
                            break
                    
            print(f"  [+] 工作表 [{sheet}] 映射成功 {match_count} 条专区码")
            sheets_data[sheet] = df

        with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
            for sheet, data in sheets_data.items():
                data.to_excel(writer, sheet_name=sheet, index=False)
        print(f"[√] 关键词映射匹配完成")
    except Exception as e:
        print(f"[X] 关键词匹配执行失败: {e}")

if __name__ == "__main__":
    pass
