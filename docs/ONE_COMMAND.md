# One-Command Batch Review

Use this command from any thread:

```powershell
PowerShell -ExecutionPolicy Bypass -File D:\CodeX\paper-download-grade\run_teacher_batch.ps1 -TeacherName "2B老师" -TargetPageUrl "http://longzhi.net.cn/examTest/subjective/queryGrade/15329-85215.mooc?schoolId=&teamId=&selectType=review-select&condition=2" -StageLabel "初稿"
```

What it does:

1. `set-source` (bind teacher + link)
2. `run-all` (download -> ingest -> grade)
3. `refresh-log`
4. `bundle-source` (final delivery folder)

For incremental daily rerun of current active source:

```powershell
PowerShell -ExecutionPolicy Bypass -File D:\CodeX\paper-download-grade\run_teacher_batch.ps1 -UseActiveSource
```

Useful optional params:

```powershell
-MaxStudents 20
-Limit 20
-Stage initial_draft
-VisualMode expert
-TextMode expert
-QueueGrade
-OverwriteBundle
```

Cross-thread prompt template:

```text
请到 D:\CodeX\paper-download-grade 执行：PowerShell -ExecutionPolicy Bypass -File D:\CodeX\paper-download-grade\run_teacher_batch.ps1 -TeacherName "<老师名>" -TargetPageUrl "<链接>" -StageLabel "初稿"
跑完后告诉我：新增下载数、完成评分数、最终打包目录路径。
```

Notes:

1. 当 `TargetPageUrl` 含有 `&` 时，优先使用 `.ps1` 入口，不要走 `.cmd`。
2. `paper_review.cmd` 只适合 `-UseActiveSource` 这种简单参数场景。
