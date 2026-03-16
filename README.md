# Paper Workflow Suite

这是重新梳理后的统一项目目录，目标是把“下载论文 + 评分论文”放进同一套可执行流程。

## 目录结构

```text
E:\CodeX\paper-workflow-suite
├─ README.md
└─ pipeline
   ├─ pipeline.py
   ├─ pipeline.config.json
   ├─ run_pipeline.ps1
   └─ state/
```

## 依赖项目（保持原仓库不动）

- 下载端：`E:\CodeX\paperdownload`
- 评分端：`E:\CodeX\paper-grading-system`

统一编排层在这里调用两边现有脚本，不重写成熟逻辑。

## 一键流程（推荐）

```powershell
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 run-all --max-students 1 --stage initial_draft --visual-mode heuristic --limit 1
```

含义：

- 下载只跑 1 个学生
- 入队时自动改名和去重
- 评分只处理本次入队的 1 篇（默认优先跑“本次新文件”）

如果你要改成“评分队列模式”（扫描 `incoming_papers` 中所有未处理文件），可以加：

```powershell
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 run-all --max-students 1 --queue-grade --limit 1
```

## 分阶段运行

```powershell
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 download --max-students 1
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 ingest
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 grade --stage initial_draft --visual-mode heuristic --limit 1
```

## 结果位置

- 下载摘要：`E:\CodeX\paperdownload\longzhi_batch_output\state\latest_automation_summary.json`
- 入队目标：`E:\CodeX\paper-grading-system\assets\incoming_papers`
- 评分结果：`E:\CodeX\paper-grading-system\grading_runs`
- 编排状态与运行报告：`E:\CodeX\paper-workflow-suite\pipeline\state`
