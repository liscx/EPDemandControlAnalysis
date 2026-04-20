import pandas as pd

def process_person_match(target_file, master_file):
    try:
        print(f"[*] 正在从映射表读取负责人映射: {master_file}")
        # 1. 从负责人分配表读取映射关系
        # 我们使用：'省份'列 (对应分公司), '负责人'列
        df_control = pd.read_excel(master_file, engine='calamine')
        
        # 清洗主表列名
        df_control.columns = [str(c).strip() for c in df_control.columns]
        
        # 建立映射字典 (省份 -> 负责人)
        person_map = {}
        for _, row in df_control.iterrows():
            # 兼容处理列名，如果存在'省份'则使用'省份'，否则由于用户需求可能是'分公司'
            comp_key = '省份' if '省份' in df_control.columns else (df_control.columns[0] if len(df_control.columns) > 0 else None)
            person_key = '负责人' if '负责人' in df_control.columns else (df_control.columns[1] if len(df_control.columns) > 1 else None)
            
            if comp_key and person_key:
                comp = str(row[comp_key]).strip()
                person = str(row[person_key]).strip()
                if comp and person and person != 'nan':
                    person_map[comp] = person

        print(f"[*] 正在回填负责人数据: {target_file}")
        sheets_data = {}
        xl = pd.ExcelFile(target_file, engine='calamine')
        
        for sheet in xl.sheet_names:
            df = pd.read_excel(target_file, sheet_name=sheet, engine='calamine')
            
            # 清洗当前 Sheet 的列名
            df.columns = [str(c).strip() for c in df.columns]

            # 寻找目标列（分公司）
            target_col = None
            for possible_name in ['分公司', '条线分公司']:
                if possible_name in df.columns:
                    target_col = possible_name
                    break
            
            if target_col:
                print(f"  正在处理工作表 [{sheet}]，匹配列为: {target_col}")
                
                # 创建负责人列（如果不存在）
                if '运营负责人' not in df.columns:
                    df['运营负责人'] = None
                
                # 执行映射回填
                df['运营负责人'] = df[target_col].apply(lambda x: person_map.get(str(x).strip()) if pd.notnull(x) else None)
                
                # 调整列顺序，把负责人紧跟在分公司后面
                cols = list(df.columns)
                if '运营负责人' in cols:
                    idx = cols.index(target_col) + 1
                    # 移动负责人列
                    cols.insert(idx, cols.pop(cols.index('运营负责人')))
                    df = df[cols]
            else:
                print(f"  跳过工作表 [{sheet}]: 未找到‘分公司’相关列")

            sheets_data[sheet] = df

        # 写回结果文件
        with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
            for sheet, data in sheets_data.items():
                data.to_excel(writer, sheet_name=sheet, index=False)

        print(f"[√] 运营负责人回填成功")

    except Exception as e:
        print(f"[X] 负责人匹配出错: {e}")

if __name__ == "__main__":
    pass