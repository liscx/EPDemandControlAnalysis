import asyncio
import os
import sys
from datetime import datetime

# 导入子模块
from export_feishu import export_process as export_feishu_func
from export_tencent import export_tencent_process as export_tencent_func
from exportDataHandle import process_logic, finalize_hyperlinks
from companyMatch import process_company_match
from contractMatch import process_contract_match
from keywordMatch import process_keyword_match
from zoneBackfill import process_zone_backfill
from personMatch import process_person_match

# 获取当前脚本所在目录的绝对路径，解决定时任务路径引用问题
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

async def main():
    # 1. 路径和环境初始化
    default_resource_dir = os.path.join(BASE_DIR, "resource")
    default_result_dir = os.path.join(BASE_DIR, "result")
    
    # 命令行参数解析
    resource_dir = sys.argv[1] if len(sys.argv) > 1 else default_resource_dir
    result_dir = sys.argv[2] if len(sys.argv) > 2 else default_result_dir
    
    if not os.path.exists(resource_dir): os.makedirs(resource_dir)
    if not os.path.exists(result_dir): os.makedirs(result_dir)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    
    print(f"{'='*50}")
    print(f"🚀 开始执行自动化工作流 (Deterministic v2)")
    print(f"[*] 资源目录: {resource_dir}")
    print(f"[*] 结果目录: {result_dir}")
    print(f"[*] 时间戳: {timestamp}")
    print(f"{'='*50}\n")

    # 2. 导出阶段
    print("【步骤 1: 导出飞书原始数据】")
    feishu_raw_name = f"飞书原始数据_{timestamp}.xlsx"
    feishu_path = await export_feishu_func(save_dir=resource_dir, target_filename=feishu_raw_name)
    if not feishu_path:
        print("[X] 飞书导出失败，终止流程。")
        return

    print("\n【步骤 2: 导出腾讯文档(项目跟进表)】")
    tencent_master_name = f"项目跟进表_{timestamp}.xlsx"
    tencent_master_path = await export_tencent_func(save_dir=resource_dir, target_filename=tencent_master_name)
    if not tencent_master_path:
        print("[X] 腾讯文档导出失败，终止流程。")
        return

    # 3. 处理阶段 (生成唯一分析结果文件)
    print("\n【步骤 3: 数据清洗与分类】")
    # 生成 需求分析结果_timestamp.xlsx
    target_result_file = process_logic(source_file=feishu_path, output_dir=result_dir, timestamp=timestamp)
    if not target_result_file:
        print("[X] 数据处理失败，终止流程。")
        return

    # 4. 专区码识别阶段 (ID -> Fuzzy -> Keyword)
    print("\n【步骤 4.1: 合同编号对应专区码】")
    process_contract_match(target_file=target_result_file, master_file=tencent_master_path)

    print("\n【步骤 4.2: 关键词映射专区码(兜底)】")
    mapping_file = os.path.join(BASE_DIR, '映射表.xlsx') 
    process_keyword_match(target_file=target_result_file, master_file=tencent_master_path, kw_file=mapping_file)

    # 5. 专区信息回填阶段 (Code -> Name/Company)
    print("\n【步骤 5: 专区名及分公司回填】")
    process_zone_backfill(target_file=target_result_file, master_file=tencent_master_path)

    # 6. 运营负责人回填阶段
    print("\n【步骤 6: 运营负责人回填】")
    person_mapping_file = os.path.join(BASE_DIR, '省份负责人分配表.xlsx')
    process_person_match(target_file=target_result_file, master_file=person_mapping_file)

    # 7. 超链接回填阶段
    print("\n【步骤 7: 需求名称超链接回填】")
    finalize_hyperlinks(target_file=target_result_file)

    print(f"\n{'='*50}")
    print(f"🎉 全部任务顺利执行完毕！")
    print(f"最终报表: {target_result_file}")
    print(f"{'='*50}")

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n[!] 任务被用户取消")
    except Exception as e:
        print(f"\n[X] 运行出错: {e}")