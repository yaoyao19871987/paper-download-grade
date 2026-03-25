# Paper Download + Grade

这个仓库现在按下面五类职责拆开了：

- `components/`: 第三方组件源码和依赖仓库
- `component_overrides/`: 对第三方组件的覆盖补丁
- `pipeline/`: 仓库自有的统一调度逻辑
- `scripts/`: 启动、初始化、凭据、批处理脚本
- `runtime/`: 下载结果、入队文件、评分产物、日志、导出包、凭据

## 当前目录结构

```text
paper-download-grade/
|-- config/
|   |-- env/
|   `-- pipeline/
|-- component_overrides/
|-- components/
|-- pipeline/
|-- runtime/
|   |-- downloads/
|   |-- grading/
|   |-- pipeline/
|   |-- tracking/
|   |-- exports/
|   `-- secrets/
|-- scripts/
|   |-- bootstrap/
|   |-- credentials/
|   |-- lib/
|   `-- run/
|-- tests/
|-- setup_windows.ps1
`-- run_teacher_batch.ps1
```

## 初始化

```powershell
PowerShell -ExecutionPolicy Bypass -File .\setup_windows.ps1
```

这个入口会：

1. 拉取 `components/paperdownload` 和 `components/essaygrade`
2. 将 `component_overrides/` 同步到组件目录
3. 安装 Node / Playwright 依赖
4. 初始化评分组件的 Python 环境

## 环境与路径

统一环境变量入口在：

- `config/env/project.env.ps1`
- `config/env/project.env.local.ps1`（可选，本机覆盖）

统一路径配置在：

- `config/pipeline/pipeline.config.json`

默认运行数据都写到 `runtime/`：

- 下载缓存：`runtime/downloads/longzhi_batch_output`
- 待评分论文：`runtime/grading/incoming_papers`
- 评分结果：`runtime/grading/runs`
- 流水线状态：`runtime/pipeline/state`
- 学生反馈与总表：`runtime/tracking/`
- 导出交付包：`runtime/exports/case_exports`
- 加密凭据：`runtime/secrets/credential_store`

## 保存凭据

```powershell
PowerShell -ExecutionPolicy Bypass -File .\scripts\credentials\save_longzhi_credential.ps1 -Username "你的账号" -Password "你的密码"
PowerShell -ExecutionPolicy Bypass -File .\scripts\credentials\save_kimi_credential.ps1 -ApiKey "你的 Kimi Key"
PowerShell -ExecutionPolicy Bypass -File .\scripts\credentials\save_siliconflow_credential.ps1 -ApiKey "你的 SiliconFlow Key"
```

凭据使用当前 Windows 用户的 DPAPI 加密，只能在当前机器当前用户下解密。

## 常用命令

```powershell
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 doctor
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 status
PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 run-all --max-students 1 --stage initial_draft --visual-mode auto --limit 1
PowerShell -ExecutionPolicy Bypass -File .\run_teacher_batch.ps1 -TeacherName "2B老师" -TargetPageUrl "<Longzhi链接>" -StageLabel "初稿"
```

## 说明

- 根目录只保留少量常用入口；详细脚本都放在 `scripts/`
- 运行数据和源码已经分开，后续清理 `runtime/` 不会影响项目代码
- 如果你要做本机定制，优先改 `config/env/project.env.local.ps1`
