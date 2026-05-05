#!/usr/bin/env python3
"""HWPX 자동 첨삭 메모 batch 삽입 CLI.

`HwpxDocument.add_memo_with_anchor(...)` 를 활용해서 학생 작품·보고서에
첨삭 메모를 일괄 추가한다.

JSON 입력 스키마 (배열):
[
  {
    "section": 0,                     # 섹션 인덱스 (기본 0)
    "paragraph": 3,                   # 문단 인덱스 (필수, 0-based)
    "text": "여기 보충 필요",          # 메모 본문 (필수)
    "author": "임세범"                 # 메모 작성자 (선택)
  },
  ...
]

또는 paragraph_text 매칭으로:
[
  {"paragraph_text": "학습 목표:", "text": "구체적 행동동사 사용 권장", "author": "교사"}
]

요구사항: python-hwpx >= 2.6.

Usage:
  python3 add_review_memo.py student.hwpx graded.hwpx --memos memos.json
  echo '[{"paragraph": 0, "text": "잘 썼어요"}]' | \\
      python3 add_review_memo.py student.hwpx graded.hwpx --memos -
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from hwpx_helpers import inject_dummy_linesegs  # noqa: E402


def _ensure_linesegs_in_zip(hwpx_path: str) -> int:
    """저장된 hwpx 의 section*.xml 에 더미 lineSegArray 주입 (in-place).

    `add_memo_with_anchor` 가 새 paragraph 를 만들면서 lineSegArray 를 비워두는
    경우가 있어 polaris-dvc strict (JID 11004) 가 깨진다. zip-level 후처리로
    무해하게 보정한다.
    """
    injected = 0
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".hwpx", dir=os.path.dirname(hwpx_path))
    os.close(tmp_fd)
    try:
        with zipfile.ZipFile(hwpx_path, "r") as zin:
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for info in zin.infolist():
                    data = zin.read(info.filename)
                    name_lower = info.filename.lower().replace("\\", "/")
                    if "/section" in name_lower and name_lower.endswith(".xml"):
                        try:
                            text = data.decode("utf-8")
                            new_text, n = inject_dummy_linesegs(text)
                            if n:
                                injected += n
                                data = new_text.encode("utf-8")
                        except UnicodeDecodeError:
                            pass
                    info_out = info
                    if info.filename == "mimetype":
                        info_out = zipfile.ZipInfo(info.filename)
                        info_out.compress_type = zipfile.ZIP_STORED
                    zout.writestr(info_out, data)
        shutil.move(tmp_path, hwpx_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise
    return injected


def add_memos_batch(
    input_path: str,
    output_path: str,
    memos: list[dict[str, Any]],
) -> int:
    """메모 batch 삽입. 추가된 메모 수 반환."""
    from hwpx import HwpxDocument

    added = 0
    with HwpxDocument.open(input_path) as doc:
        sections = list(doc.sections)
        for memo_spec in memos:
            text = memo_spec.get("text")
            if not text:
                print(
                    f"[WARN] memo without text, skipped: {memo_spec!r}",
                    file=sys.stderr,
                )
                continue

            kwargs: dict[str, Any] = {"text": text}
            if "author" in memo_spec and memo_spec["author"]:
                kwargs["author"] = memo_spec["author"]

            # 위치 지정: paragraph_text 가 우선, 없으면 section/paragraph 인덱스
            if "paragraph_text" in memo_spec and memo_spec["paragraph_text"]:
                kwargs["paragraph_text"] = memo_spec["paragraph_text"]
            else:
                section_idx = memo_spec.get("section", 0)
                para_idx = memo_spec.get("paragraph")
                if para_idx is None:
                    print(
                        f"[WARN] memo without paragraph index or paragraph_text, skipped: {memo_spec!r}",
                        file=sys.stderr,
                    )
                    continue
                if section_idx >= len(sections):
                    print(
                        f"[WARN] section_index {section_idx} 가 범위 밖 (총 {len(sections)}). skipped",
                        file=sys.stderr,
                    )
                    continue
                paragraphs = list(sections[section_idx].paragraphs)
                if para_idx >= len(paragraphs):
                    print(
                        f"[WARN] paragraph {para_idx} 범위 밖 (sec {section_idx} 총 {len(paragraphs)}). skipped",
                        file=sys.stderr,
                    )
                    continue
                kwargs["paragraph"] = paragraphs[para_idx]

            try:
                doc.add_memo_with_anchor(**kwargs)
                added += 1
            except Exception as exc:
                print(f"[WARN] add_memo failed for {memo_spec!r}: {exc}", file=sys.stderr)

        doc.save_to_path(output_path)

    # save 직후 lineSegArray 후처리 (polaris strict JID 11004 방지)
    injected = _ensure_linesegs_in_zip(output_path)
    if injected:
        print(f"[NS] lineseg dummy injected: {injected}", file=sys.stderr)
    return added


def _parse_memos_arg(value: str) -> list[dict[str, Any]]:
    if value == "-":
        raw = sys.stdin.read()
    else:
        raw = Path(value).read_text(encoding="utf-8")
    data = json.loads(raw)
    if not isinstance(data, list):
        raise ValueError("memos JSON must be a list of objects")
    return data


def _parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="HWPX 문서에 첨삭 메모를 batch 삽입한다 (학생 작품 평가용)."
    )
    p.add_argument("input_hwpx", help="입력 .hwpx 경로")
    p.add_argument("output_hwpx", help="출력 .hwpx 경로")
    p.add_argument(
        "--memos",
        required=True,
        help="메모 JSON 파일 경로 (또는 stdin: '-')",
    )
    p.add_argument(
        "--default-author",
        default=None,
        help="memo 객체에 author 가 없을 때 사용할 기본 작성자",
    )
    return p.parse_args(argv)


def main(argv: list[str]) -> int:
    args = _parse_args(argv)

    in_path = Path(args.input_hwpx).resolve()
    if not in_path.exists():
        print(f"[ERR] file not found: {in_path}", file=sys.stderr)
        return 2

    try:
        memos = _parse_memos_arg(args.memos)
    except Exception as exc:
        print(f"[ERR] memos JSON parse failed: {exc}", file=sys.stderr)
        return 2

    if args.default_author:
        for m in memos:
            m.setdefault("author", args.default_author)

    out_path = Path(args.output_hwpx).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    added = add_memos_batch(str(in_path), str(out_path), memos)
    print(f"[OK] {added}/{len(memos)}개 메모 추가 → {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
