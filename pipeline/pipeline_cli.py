from __future__ import annotations

import argparse


def add_grading_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--stage", choices=["initial_draft", "final"], default="initial_draft")
    parser.add_argument(
        "--visual-mode",
        choices=["auto", "openai", "moonshot", "siliconflow", "expert", "heuristic", "off"],
        default="auto",
    )
    parser.add_argument("--visual-model", default="gpt-5.4")
    parser.add_argument(
        "--text-mode",
        choices=["off", "auto", "expert", "siliconflow", "moonshot"],
        default="expert",
    )
    parser.add_argument("--text-primary-model", default="deepseek-ai/DeepSeek-V3.2")
    parser.add_argument("--text-secondary-model", default="kimi-for-coding")
    parser.add_argument("--limit", type=int, default=0)


def build_parser(default_config: str) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Unified paper download and grading workflow."
    )
    parser.add_argument("--config", default=default_config, help="Path to pipeline.config.json")

    sub = parser.add_subparsers(dest="command", required=True)

    p_download = sub.add_parser("download", help="Run the download stage only.")
    p_download.add_argument("--page-size", type=int, default=100)
    p_download.add_argument("--start-page", type=int, default=1)
    p_download.add_argument("--max-students", type=int, default=0)

    sub.add_parser("ingest", help="Run the ingest stage only.")

    p_grade = sub.add_parser("grade", help="Run the grading stage only.")
    add_grading_args(p_grade)

    p_all = sub.add_parser("run-all", help="Run download, ingest, and grade.")
    p_all.add_argument("--page-size", type=int, default=100)
    p_all.add_argument("--start-page", type=int, default=1)
    p_all.add_argument("--max-students", type=int, default=0)
    add_grading_args(p_all)
    p_all.add_argument("--grade-even-if-no-new", action="store_true")
    p_all.add_argument(
        "--queue-grade",
        action="store_true",
        help="Use the grader queue instead of only grading newly ingested files.",
    )

    p_bundle = sub.add_parser("bundle-case", help="Export a delivery bundle.")
    p_bundle.add_argument("--case-name", required=True)
    p_bundle.add_argument("--student-ids", default="")
    p_bundle.add_argument("--latest-graded", type=int, default=0)
    p_bundle.add_argument("--all-graded", action="store_true")
    p_bundle.add_argument("--overwrite", action="store_true")

    p_source = sub.add_parser("set-source", help="Register a teacher source.")
    p_source.add_argument("--teacher-name", required=True)
    p_source.add_argument("--target-page-url", required=True)
    p_source.add_argument("--stage-label", default="初稿")
    p_source.add_argument("--no-set-active", action="store_true")
    p_source.add_argument("--bind-all-current", action="store_true")

    sub.add_parser("list-sources", help="List known sources.")

    p_bundle_source = sub.add_parser("bundle-source", help="Bundle all graded output for a source.")
    p_bundle_source.add_argument("--source-key", required=True)
    p_bundle_source.add_argument("--overwrite", action="store_true")

    sub.add_parser("refresh-log", help="Refresh repository-level tracking outputs.")
    sub.add_parser("status", help="Show pipeline status.")
    sub.add_parser("doctor", help="Check environment prerequisites.")
    return parser
