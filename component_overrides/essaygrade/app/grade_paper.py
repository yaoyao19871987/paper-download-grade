from __future__ import annotations

import argparse
from pathlib import Path

from paper_grader.grader import dump_json, grade_document, render_text_report


def main() -> None:
    parser = argparse.ArgumentParser(description="使用 Word COM 检查毕业论文格式并进行内容评分。")
    parser.add_argument("document", help="待评分的 .doc 或 .docx 文件路径")
    parser.add_argument(
        "--reference-doc",
        action="append",
        default=[],
        help="可选：作为相似度参考的范文路径，可重复传入多个",
    )
    parser.add_argument(
        "--stage",
        choices=["initial_draft", "final"],
        default="initial_draft",
        help="评阅阶段：initial_draft 更强调格式门槛，final 适合终稿复核",
    )
    parser.add_argument(
        "--visual-mode",
        choices=["auto", "openai", "moonshot", "siliconflow", "expert", "heuristic", "off"],
        default="auto",
        help="视觉审稿模式：auto 优先调用已配置的大模型，失败后回退到本地启发式；openai 只用 OpenAI；moonshot 只用 Moonshot/Kimi；siliconflow 只用 SiliconFlow；expert 使用 SiliconFlow 双专家；heuristic 只用本地启发式；off 关闭视觉审稿",
    )
    parser.add_argument(
        "--visual-model",
        default="gpt-5.4",
        help="视觉审稿使用的大模型名称，默认 gpt-5.4",
    )
    parser.add_argument(
        "--visual-output-dir",
        help="可选：视觉审稿导出的 PDF 与原始响应输出目录",
    )
    parser.add_argument("--json-out", help="可选：输出 JSON 评分结果")
    parser.add_argument("--text-out", help="可选：输出文本评分报告")
    args = parser.parse_args()

    result = grade_document(
        args.document,
        reference_docs=args.reference_doc,
        stage=args.stage,
        visual_mode=args.visual_mode,
        visual_model=args.visual_model,
        visual_output_dir=args.visual_output_dir,
    )
    report = render_text_report(result)
    print(report)

    if args.json_out:
        dump_json(result, args.json_out)
    if args.text_out:
        Path(args.text_out).write_text(report, encoding="utf-8")


if __name__ == "__main__":
    main()
