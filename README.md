# Paper Download + Grade (Portable)

这是一个可迁移到新电脑的统一工作流仓库，负责：

1. Longzhi 批量下载论文
2. 新论文改名入队
3. 触发评分（支持“只评本次新文件”）

这版已去掉硬编码盘符路径，默认基于仓库相对路径运行。

## 新电脑能不能直接用？

可以，但前提是先跑一次初始化脚本安装依赖并拉取子项目。

## 前置要求

- Windows（评分链路依赖 Word COM）
- Microsoft Word 已安装
- Git
- Node.js 18+
- Python 3.11+

## 目录结构

```text
paper-download-grade/
├─ setup_windows.ps1                # 新电脑初始化（拉仓库+安装依赖）
├─ save_longzhi_credential.ps1      # 保存 Longzhi 凭据（一次）
├─ pipeline/
│  ├─ run_pipeline.ps1
│  ├─ pipeline.py
│  ├─ pipeline.config.json
│  └─ state/
└─ components/                       # setup 后自动拉取（默认不进 git）
   ├─ paperdownload/
   └─ essaygrade/
```

## 1. 新电脑初始化

在仓库根目录运行：

```powershell
PowerShell -ExecutionPolicy Bypass -File .\setup_windows.ps1
```

这个脚本会做三件事：

1. 拉取 `paperdownload` 与 `essaygrade` 到 `components/`
2. 安装下载端 Node 依赖并安装 Playwright Chromium
3. 初始化评分端 Python 环境

如果你要指定不同的依赖仓库地址（比如 fork 或私有库）：

```powershell
PowerShell -ExecutionPolicy Bypass -File .\setup_windows.ps1 `
  -PaperDownloadRepoUrl "https://github.com/<you>/paperdownload.git" `
  -EssayGradeRepoUrl "https://github.com/<you>/essaygrade.git"
```

## 2. 保存下载凭据（一次即可）

```powershell
PowerShell -ExecutionPolicy Bypass -File .\save_longzhi_credential.ps1 -Username "你的账号" -Password "你的密码"
```

## 3. 运行健康检查

```powershell
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 doctor
```

## 4. 跑 1 个学生做闭环验证

```powershell
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 run-all --max-students 1 --stage initial_draft --visual-mode heuristic --limit 1
```

含义：

- 下载只跑 1 人
- 入队自动改名并去重
- 默认只评分本次新入队文件（确保“下完就评”）

如果你要按评分队列（扫描全部未处理）运行：

```powershell
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 run-all --max-students 1 --queue-grade --limit 1
```

## 常用命令

```powershell
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 status
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 download --max-students 20
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 ingest
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 grade --stage initial_draft --visual-mode auto --limit 10
```

## 结果位置

- 下载摘要：`components/paperdownload/longzhi_batch_output/state/latest_automation_summary.json`
- 入队目录：`components/essaygrade/assets/incoming_papers`
- 评分结果：`components/essaygrade/grading_runs`
- 编排状态报告：`pipeline/state/reports/*.json`
