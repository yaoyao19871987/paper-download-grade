# 2026-03-26 apply-audit 断点

## 当前状态
- `apply-audit` 已启动并执行到 `rerun_grading` 阶段。
- 这次执行在重批过程中被手动中断，最终的 `refresh-log` / 总汇总刷新没有完成。
- `runtime/tracking/student_progress_log.json`、`student_progress_log.md`、`audit_review_summary.json`、`audit_review_summary.md`、`audit_regrade_list.json` 仍然停留在中断前的旧版本。

## 已落盘的进展
- 今天下午 14:10 之后，`runtime/grading/runs/` 里有 `77` 个 run 目录被写动过。
- 当前可见的最新落点是 `2026-03-24_102200_20260320_24509056_____5dda54eb`，时间是 `2026-03-26 16:07:55`。
- 说明重批链路已经推进到 `24509056 / 李树芳` 附近，但没有走到最终汇总阶段。

## 已完成的代码修正
- `pipeline/pipeline_feedback.py`
  - 学生评语生成现在支持 `audit_review` 上下文。
  - `feedback_only / rescore_and_feedback / rerun_grading` 会进入新提示词。
  - `keep` 类学生会继续复用原缓存，不白烧 token。
- `pipeline/pipeline_tracking.py`
  - 现在会读取 `audit_results.json`，把复核结论写进学生总表。
  - 会重新生成 `audit_review_summary.json`、`audit_review_summary.md`、`audit_regrade_list.json`。
  - 学生总表新增了 `复核结论 / 复核建议` 两列。
- `pipeline/pipeline.py`
  - 新增 `apply-audit` 入口。
  - `rerun_grading` 现在会按学生分组重跑评分。
- `components/essaygrade/app/grade_paper.py`
  - 已补 UTF-8 stdout/stderr，避免中文输出触发 `gbk` 编码错误。

## 恢复命令
明天直接从这里续跑：

```powershell
python pipeline\pipeline.py apply-audit
```

## 注意点
- 如果 `apply-audit` 再次卡在单个学生的重批上，优先检查那个学生对应的原始下载文件是否还存在。
- 如果评分脚本再次报中文编码问题，先确认 `components/essaygrade/app/grade_paper.py` 的 UTF-8 输出补丁仍在。
- 这次中断前，最终汇总文件没有刷新，所以不要把旧的 `student_progress_log.md` 当成本轮结果。
