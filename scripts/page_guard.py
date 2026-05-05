#!/usr/bin/env python3
"""HWPX 레퍼런스 대비 페이지 드리프트 위험 검사.

레퍼런스 문서의 구조·텍스트 길이 분포와 결과 문서를 비교해 "쪽수가 달라질
가능성이 높은" 변경을 사전에 차단한다. 실제 렌더러의 페이지 계산을 대체할
수는 없지만, 다음 신호로 위험을 잡아낸다:

검사 항목:
- 문단 수 / 표 수 / 표 구조(rowCnt, colCnt, width, height, repeatHeader, pageBreak) 동일성
- 명시적 ``pageBreak`` / ``columnBreak`` 속성 수 동일성
- 전체 텍스트 길이 편차 (기본 15%) 한도
- 문단별 텍스트 길이 급변 (기본 25%) 감지
- 모든 ``Contents/section*.xml`` 순회 (다중 섹션 대응)

알고리즘 출처: ``Canine89/hwpxskill`` (https://github.com/Canine89/hwpxskill,
라이선스 미지정 상태이므로 알고리즘만 참조하여 본 스킬 스타일 — `xpath_local()`,
다중 섹션, zip bomb 상한 — 으로 재구현). 자세한 고지는
``THIRD_PARTY_NOTICES.md`` 5번 항목 참조.

Usage:
    python3 page_guard.py --reference ref.hwpx --output result.hwpx
    python3 page_guard.py -r ref.hwpx -o result.hwpx --json
    python3 page_guard.py -r ref.hwpx -o result.hwpx \\
        --max-text-delta-ratio 0.10 --max-paragraph-delta-ratio 0.20
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from dataclasses import asdict, dataclass, field
from io import BytesIO
from pathlib import Path

from lxml import etree

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from hwpx_helpers import xpath_local  # noqa: E402

# verify_hwpx 와 동일한 zip bomb 상한 (32MB / 엔트리)
MAX_XML_BYTES = 32 * 1024 * 1024


@dataclass
class Metrics:
    """레퍼런스/결과 문서에서 측정한 구조·텍스트 메트릭."""

    paragraph_count: int = 0
    page_break_count: int = 0
    column_break_count: int = 0
    table_count: int = 0
    # (rowCnt, colCnt, width, height, repeatHeader, pageBreak)
    table_shapes: list[tuple[str, str, str, str, str, str]] = field(default_factory=list)
    text_char_total: int = 0
    text_char_total_nospace: int = 0
    paragraph_text_lengths: list[int] = field(default_factory=list)
    section_count: int = 0


def _list_section_entries(zf: zipfile.ZipFile) -> list[str]:
    """ZIP 내 ``Contents/section*.xml`` 항목명을 정렬해서 반환."""
    names = []
    for name in zf.namelist():
        n = name.lower().replace("\\", "/")
        if n.startswith("contents/section") and n.endswith(".xml"):
            names.append(name)
    return sorted(names)


def _read_section_bytes(zf: zipfile.ZipFile, name: str) -> bytes:
    info = zf.getinfo(name)
    if info.file_size > MAX_XML_BYTES:
        raise RuntimeError(
            f"section XML 크기가 상한을 초과: {name} = {info.file_size} bytes "
            f"(limit {MAX_XML_BYTES})"
        )
    return zf.read(name)


def _text_of(elem: etree._Element) -> str:
    """``<hp:t>`` 노드의 모든 inner text 를 이어 붙여 반환."""
    return "".join(elem.itertext())


def _paragraph_text_length(p: etree._Element) -> int:
    return sum(len(_text_of(t)) for t in xpath_local(p, "t"))


def collect_metrics(hwpx_path: Path) -> Metrics:
    """HWPX 파일에서 모든 section 의 메트릭을 합산해 반환."""
    metrics = Metrics()
    with zipfile.ZipFile(str(hwpx_path), "r") as zf:
        section_names = _list_section_entries(zf)
        if not section_names:
            raise RuntimeError(f"section XML 을 찾을 수 없습니다: {hwpx_path}")
        metrics.section_count = len(section_names)

        for name in section_names:
            data = _read_section_bytes(zf, name)
            try:
                root = etree.parse(BytesIO(data)).getroot()
            except etree.XMLSyntaxError as exc:
                raise RuntimeError(f"section XML parse 실패 ({name}): {exc}") from exc

            paragraphs = xpath_local(root, "p")
            metrics.paragraph_count += len(paragraphs)

            for p in paragraphs:
                if p.get("pageBreak") == "1":
                    metrics.page_break_count += 1
                if p.get("columnBreak") == "1":
                    metrics.column_break_count += 1

            tables = xpath_local(root, "tbl")
            metrics.table_count += len(tables)
            for t in tables:
                sz_list = xpath_local(t, "sz")
                width = sz_list[0].get("width", "") if sz_list else ""
                height = sz_list[0].get("height", "") if sz_list else ""
                metrics.table_shapes.append(
                    (
                        t.get("rowCnt", ""),
                        t.get("colCnt", ""),
                        width,
                        height,
                        t.get("repeatHeader", ""),
                        t.get("pageBreak", ""),
                    )
                )

            for t in xpath_local(root, "t"):
                s = _text_of(t)
                metrics.text_char_total += len(s)
                metrics.text_char_total_nospace += len("".join(s.split()))

            for p in paragraphs:
                metrics.paragraph_text_lengths.append(_paragraph_text_length(p))

    return metrics


def _ratio_delta(a: int, b: int) -> float:
    base = max(a, 1)
    return abs(b - a) / base


def compare_metrics(
    ref: Metrics,
    out: Metrics,
    *,
    max_text_delta_ratio: float = 0.15,
    max_paragraph_delta_ratio: float = 0.25,
) -> list[str]:
    """ref 와 out 메트릭을 비교해 위반 메시지 리스트를 반환 (빈 리스트 = PASS)."""
    errors: list[str] = []

    if ref.section_count != out.section_count:
        errors.append(
            f"섹션 수 불일치: ref={ref.section_count}, out={out.section_count}"
        )
    if ref.paragraph_count != out.paragraph_count:
        errors.append(
            f"문단 수 불일치: ref={ref.paragraph_count}, out={out.paragraph_count}"
        )
    if ref.page_break_count != out.page_break_count:
        errors.append(
            f"명시적 pageBreak 수 불일치: "
            f"ref={ref.page_break_count}, out={out.page_break_count}"
        )
    if ref.column_break_count != out.column_break_count:
        errors.append(
            f"명시적 columnBreak 수 불일치: "
            f"ref={ref.column_break_count}, out={out.column_break_count}"
        )
    if ref.table_count != out.table_count:
        errors.append(
            f"표 수 불일치: ref={ref.table_count}, out={out.table_count}"
        )
    if ref.table_shapes != out.table_shapes:
        errors.append("표 구조(rowCnt/colCnt/width/height/repeatHeader/pageBreak) 불일치")

    td = _ratio_delta(ref.text_char_total_nospace, out.text_char_total_nospace)
    if td > max_text_delta_ratio:
        errors.append(
            "전체 텍스트 길이 편차 초과: "
            f"ref={ref.text_char_total_nospace}, out={out.text_char_total_nospace}, "
            f"delta={td:.2%}, limit={max_text_delta_ratio:.2%}"
        )

    if len(ref.paragraph_text_lengths) == len(out.paragraph_text_lengths):
        for idx, (a, b) in enumerate(
            zip(ref.paragraph_text_lengths, out.paragraph_text_lengths), start=1
        ):
            if a == 0 and b == 0:
                continue
            pd = _ratio_delta(a, b)
            if pd > max_paragraph_delta_ratio:
                errors.append(
                    f"{idx}번째 문단 텍스트 길이 편차 초과: "
                    f"ref={a}, out={b}, delta={pd:.2%}, limit={max_paragraph_delta_ratio:.2%}"
                )

    return errors


def page_guard(
    reference: str | Path,
    output: str | Path,
    *,
    max_text_delta_ratio: float = 0.15,
    max_paragraph_delta_ratio: float = 0.25,
) -> tuple[Metrics, Metrics, list[str]]:
    """페이지 가드 진입점. (ref_metrics, out_metrics, errors) 반환."""
    ref = collect_metrics(Path(reference))
    out = collect_metrics(Path(output))
    errors = compare_metrics(
        ref,
        out,
        max_text_delta_ratio=max_text_delta_ratio,
        max_paragraph_delta_ratio=max_paragraph_delta_ratio,
    )
    return ref, out, errors


def _parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="HWPX 레퍼런스 대비 페이지 드리프트 위험 검사"
    )
    p.add_argument("--reference", "-r", required=True, help="기준 HWPX 경로")
    p.add_argument("--output", "-o", required=True, help="결과 HWPX 경로")
    p.add_argument(
        "--max-text-delta-ratio",
        type=float,
        default=0.15,
        help="전체 텍스트 길이 허용 편차 비율 (기본: 0.15)",
    )
    p.add_argument(
        "--max-paragraph-delta-ratio",
        type=float,
        default=0.25,
        help="문단별 텍스트 길이 허용 편차 비율 (기본: 0.25)",
    )
    p.add_argument("--json", action="store_true", help="metrics 를 JSON 으로 출력")
    return p.parse_args(argv)


def main(argv: list[str]) -> int:
    args = _parse_args(argv)
    ref_path = Path(args.reference)
    out_path = Path(args.output)

    if not ref_path.exists():
        print(f"Error: reference not found: {ref_path}", file=sys.stderr)
        return 2
    if not out_path.exists():
        print(f"Error: output not found: {out_path}", file=sys.stderr)
        return 2

    try:
        ref, out, errors = page_guard(
            ref_path,
            out_path,
            max_text_delta_ratio=args.max_text_delta_ratio,
            max_paragraph_delta_ratio=args.max_paragraph_delta_ratio,
        )
    except RuntimeError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2

    if args.json:
        print(
            json.dumps(
                {"reference": asdict(ref), "output": asdict(out)},
                ensure_ascii=False,
                indent=2,
            )
        )

    if errors:
        print("FAIL: page-guard")
        for e in errors:
            print(f" - {e}")
        return 1

    print("PASS: page-guard")
    print(
        f"  섹션 {ref.section_count} / 문단 {ref.paragraph_count} / 표 {ref.table_count}"
        f" 동일, 텍스트 길이 편차 허용 범위 내"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
