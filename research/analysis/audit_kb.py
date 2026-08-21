from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from kb_schema import REQUIRED_DATA_FILES, issues_to_frame, validate_data_dir

DEFAULT_DATA_DIR = PROJECT_ROOT / "data" / "processed"
DEFAULT_REPORT_DIR = PROJECT_ROOT / "research" / "drafts" / "reports"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="审计左联知识库标准数据目录。")
    parser.add_argument("--data-dir", type=Path, default=DEFAULT_DATA_DIR, help="标准数据目录")
    parser.add_argument("--report-dir", type=Path, default=DEFAULT_REPORT_DIR, help="报告输出目录")
    return parser.parse_args()


def _write_report(report_dir: Path, data_dir: Path, result) -> None:
    report_dir.mkdir(parents=True, exist_ok=True)

    issue_frame = issues_to_frame(result.issues)
    issues_path = report_dir / "audit_kb_issues.csv"
    issue_frame.to_csv(issues_path, index=False, encoding="utf-8-sig")

    existing_files = [filename for filename in REQUIRED_DATA_FILES if (data_dir / filename).exists()]
    lines = [
        "# Knowledge Base Audit Report",
        "",
        f"- 数据目录：`{data_dir}`",
        f"- 已发现文件：{len(existing_files)}/{len(REQUIRED_DATA_FILES)}",
        f"- 严重错误：{len(result.errors)}",
        f"- 警告：{len(result.warnings)}",
        "",
        "## Top Issues",
    ]
    if result.issues:
        for issue in result.issues[:30]:
            row_text = f"（{issue.row_ref}）" if issue.row_ref else ""
            lines.append(f"- `{issue.severity}` `{issue.code}` `{issue.table}` {row_text} {issue.message}")
    else:
        lines.append("- 未发现问题。")

    if not issue_frame.empty:
        lines.extend(
            [
                "",
                "## Issue Counts",
                "",
                "| severity | code | count |",
                "| --- | --- | ---: |",
            ]
        )
        grouped = (
            issue_frame.groupby(["severity", "code"], as_index=False)
            .size()
            .sort_values(["severity", "size", "code"], ascending=[True, False, True])
        )
        for _, row in grouped.iterrows():
            lines.append(f"| {row['severity']} | {row['code']} | {int(row['size'])} |")

    report_path = report_dir / "audit_kb_report.md"
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    args = parse_args()
    result = validate_data_dir(args.data_dir)
    _write_report(args.report_dir, args.data_dir, result)
    print(result.summary())
    return 1 if result.has_errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
