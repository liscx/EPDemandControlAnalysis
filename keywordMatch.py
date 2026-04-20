import pandas as pd
import math
import os

def process_keyword_match(target_file, master_file, kw_file='映射表.xlsx'):
    print(f"[*] 正在进行关键词兜底匹配: {target_file}")
    
    # 1. 读管控表 (Master)
    try:
        df_control = pd.read_excel(master_file, sheet_name='专区管控表', engine='calamine')
        code_map = {}
        for _, row in df_control.iterrows():
            code = row.iloc[1]
            name = row.iloc[2]
            comp = row.iloc[5]
            if pd.notna(code):
                code_map[str(code).strip()] = (name, comp)
    except Exception as e:
        print(f"警告：无法读取管控表 {master_file}: {e}")
        code_map = {}

    # 2. 读关键字映射表
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
        kw_dict = {}

    # 3. 处理目标表
    try:
        xl = pd.ExcelFile(target_file, engine='calamine')
        sheets_data = {}
        for sheet in xl.sheet_names:
            df = pd.read_excel(target_file, sheet_name=sheet, engine='calamine')
            for idx, row in df.iterrows():
                # 只对分公司为空的行进行兜底
                if pd.isna(row.get('分公司')) or str(row.get('分公司')).strip() == "":
                    req_name = str(row.get('需求名称', ''))
                    matched = False
                    for kw, code_str in kw_dict.items():
                        if kw in req_name:
                            if code_str in code_map:
                                mapped_name, mapped_comp = code_map[code_str]
                                df.at[idx, '专区'] = mapped_name
                                df.at[idx, '分公司'] = mapped_comp
                                if '专区码' in df.columns:
                                    df.at[idx, '专区码'] = code_str
                                matched = True
                            break
                    
                    if not matched:
                        df.at[idx, '专区'] = None
                        df.at[idx, '分公司'] = None
            sheets_data[sheet] = df

        with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
            for sheet, data in sheets_data.items():
                data.to_excel(writer, sheet_name=sheet, index=False)
        print(f"[√] 关键词兜底匹配完成")
    except Exception as e:
        print(f"[X] 关键词匹配执行失败: {e}")

if __name__ == "__main__":
    pass
