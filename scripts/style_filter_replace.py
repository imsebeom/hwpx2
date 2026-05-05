#!/usr/bin/env python3
"""스타일 필터 텍스트 치환 CLI.

`HwpxDocument.replace_text_in_runs(...)` 의 스타일 필터 (text_color,
underline_type, underline_color, char_pr_id_ref, limit) 를 활용해서
조건에 맞는 run 만 골라 치환한다.

활용 사례:
  - 빨간 글씨로 표시된 TODO 만 검정으로 일괄 변경
  - 밑줄 친 단어만 굵게 (1단계 치환 후 별도 처리)
  - 특정 charPr ID 사용한 라벨 텍스트 일괄 교체

요구사항: python-hwpx >= 2.6 (HwpxDocument API).

Usage:
  python3 style_filter_replace.py input.hwpx output.hwpx \\
      "TODO" "DONE" --color "#FF0000"

  python3 style_filter_replace.py student.hwpx graded.hwpx \\
      "잘못" "정정 필요" --color "#FF0000" --underline SOLID --limit 5
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def style_replace(
    input_path: str,
    output_path: str,
    search: str,
    replacement: str,
    *,
    text_color: str | None = None,
    underline_type: str | None = None,
    underline_color: str | None = None,
    char_pr_id_ref: str | int | None = None,
    limit: int | None = None,
) -> int:
    """스타일 필터 치환 본체. 치환된 run 개수 반환."""
    from hwpx import HwpxDocument

    with HwpxDocument.open(input_path) as doc:
        count = doc.replace_text_in_runs(
            search,
            replacement,
            text_color=text_color,
            underline_type=underline_type,
            underline_color=underline_color,
            char_pr_id_ref=char_pr_id_ref,
            limit=limit,
        )
        doc.save_to_path(output_path)
    return count


def list_styled_runs(
    input_path: str,
    *,
    text_color: str | None = None,
    underline_type: str | None = None,
    underline_color: str | None = None,
    char_pr_id_ref: str | int | None = None,
) -> list[str]:
    """필터에 매칭되는 run 의 텍스트 목록 (사전 점검용)."""
    from hwpx import HwpxDocument

    with HwpxDocument.open(input_path) as doc:
        runs = doc.find_runs_by_style(
            text_color=text_color,
            underline_type=underline_type,
            underline_color=underline_color,
            char_pr_id_ref=char_pr_id_ref,
        )
        return [r.text for r in runs]


def _parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="HWPX 문서에서 스타일 필터에 매칭되는 run 만 골라 텍스트 치환."
    )
    p.add_argument("input_hwpx", help="입력 .hwpx 경로")
    p.add_argument("output_hwpx", help="출력 .hwpx 경로")
    p.add_argument("search", help="찾을 문자열")
    p.add_argument("replacement", help="대체 문자열")
    p.add_argument(
        "--color",
        dest="text_color",
        default=None,
        help="글자 색 필터 (#RRGGBB). 미지정 시 모든 run 대상",
    )
    p.add_argument(
        "--underline",
        dest="underline_type",
        default=None,
        help="밑줄 타입 (SOLID/DOTTED/DASH 등). 미지정 시 미적용",
    )
    p.add_argument(
        "--underline-color",
        dest="underline_color",
        default=None,
        help="밑줄 색 (#RRGGBB)",
    )
    p.add_argument(
        "--char-pr",
        dest="char_pr_id_ref",
        default=None,
        help="charPrIDRef 값 (int 또는 str)",
    )
    p.add_argument(
        "--limit",
        type=int,
        default=None,
        help="최대 치환 횟수 (기본: 무제한)",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="실제 치환 없이 매칭되는 run 텍스트만 출력",
    )
    return p.parse_args(argv)


def main(argv: list[str]) -> int:
    args = _parse_args(argv)

    in_path = Path(args.input_hwpx).resolve()
    if not in_path.exists():
        print(f"[ERR] file not found: {in_path}", file=sys.stderr)
        return 2

    if args.dry_run:
        runs = list_styled_runs(
            str(in_path),
            text_color=args.text_color,
            underline_type=args.underline_type,
            underline_color=args.underline_color,
            char_pr_id_ref=args.char_pr_id_ref,
        )
        print(f"[DRY] 매칭 run {len(runs)}개:")
        for i, t in enumerate(runs[:50]):
            print(f"  {i}: {t!r}")
        if len(runs) > 50:
            print(f"  ... +{len(runs) - 50}개")
        return 0

    out_path = Path(args.output_hwpx).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    count = style_replace(
        str(in_path),
        str(out_path),
        args.search,
        args.replacement,
        text_color=args.text_color,
        underline_type=args.underline_type,
        underline_color=args.underline_color,
        char_pr_id_ref=args.char_pr_id_ref,
        limit=args.limit,
    )
    print(f"[OK] {count}개 run 치환 → {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
