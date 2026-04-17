---
name: demand-overdue-analysis
description: 自动化分析需求管控表，提取期望完成时间超期10天以上且未完成的需求。支持排除已上线、暂停、作废等特定状态。
triggers:
  - 统计需求超期情况
  - 分析需求管控表
  - 提取严重超期需求
  - 需求超期
requires:
  bins: [python]
---

# 需求超期自动化分析技能 (Demand Overdue Analysis)

此技能用于对《需求管控表》执行风险监控，识别那些“期望完成时间”已过期 10 天以上，但仍处于活跃/未完成状态的需求。

## 场景说明
适用于项目管理复盘或周报统计。技能会自动过滤掉已结束或已失效的状态（已上线、待更新正式、暂停、作废、待更新仿真），仅关注真正存在滞后风险的任务。

## 自动化执行流程

执行该技能时，OpenClaw 将尝试定位指定的 Excel 文件并运行分析脚本。输出结果将自动生成在脚本同级目录下。

```powershell
# 1. 设置输入文件路径（支持全量路径）
$inputFile = "<用户提供的路径>"

# 2. 检查文件是否存在
if (-not (Test-Path $inputFile)) { Write-Error "找不到指定的需求管控表文件，请检查路径。"; return }

# 3. 执行 Python 分析脚本
python main.py "$inputFile"