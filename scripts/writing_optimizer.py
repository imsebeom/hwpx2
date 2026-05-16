#!/usr/bin/env python3
"""writing_optimizer.py — 공공기관 보고서 글쓰기 규칙 자동 적용 + 검토 권장 리포트.

「적의를 보이는 것들」 4종(-적/의/것/들) + 12개 변환 패턴을 적용한다.
원칙·근거는 references/writing-principles.md, references/layout-rules.md 참조.

출처: 정규식 12개 규칙은 public-doc-to-hwpx v3.6.11 (Kminer2053, MIT) 의
references/layout-rules.md 표를 독립 재구현한 것이다. 원본 소스 코드는 포함하지 않음.
https://github.com/Kminer2053/public-doc-to-hwpx

한국어 조사 보강:
  - R6 (`~하는 것이 필요/중요합니다`): 받침 유무로 '이/가' 자동 선택
  - R8 (`여러 ~들이/들을/들에…`): 조사·구두점 다양 lookahead
  - R1 (`~와/과 관련된`): 양 조사 모두 매칭

CLI:
    python writing_optimizer.py input.md
    python writing_optimizer.py input.md --output cleaned.md --report report.md
    python writing_optimizer.py - < input.md           # stdin 입력

API:
    from writing_optimizer import optimize_text, format_report
    cleaned, suggestions = optimize_text(text)
    print(format_report(suggestions))
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple


@dataclass
class Suggestion:
    rule_id: str
    severity: str            # "auto" | "review"
    line: int                # 1-based
    column: int              # 0-based
    before: str
    after: Optional[str]     # review-only 은 None
    reason: str


# ──────────────────────────────────────────────────────────────────────
# 한국어 조사 헬퍼
# ──────────────────────────────────────────────────────────────────────

def _has_jongseong(ch: str) -> bool:
    """한 글자에 받침이 있는지. 한글 음절만 판정 (한자·영문은 False)."""
    if not ch:
        return False
    code = ord(ch) - 0xAC00
    return 0 <= code <= 11171 and (code % 28) != 0


def _josa_iga(word: str) -> str:
    """word 끝 받침 유무로 '이' 또는 '가' 선택. 한자/영문 끝은 '이' 기본."""
    if not word:
        return "이"
    return "이" if _has_jongseong(word[-1]) else "가"


def _r6_need(m: "re.Match[str]") -> str:
    return f"{m.group(1)}{_josa_iga(m.group(1))} 필요합니다"


def _r6_important(m: "re.Match[str]") -> str:
    return f"{m.group(1)}{_josa_iga(m.group(1))} 중요합니다"


# 받침 유무로 짝조사 재선택 (이/가, 은/는, 을/를 만 처리. 으로/로는 ㄹ 예외 복잡으로 제외)
_JOSA_PAIRS = {
    "이": ("이", "가"), "가": ("이", "가"),
    "은": ("은", "는"), "는": ("은", "는"),
    "을": ("을", "를"), "를": ("을", "를"),
}


def _josa_swap(word: str, particle: str) -> str:
    """word 받침 따라 짝조사 재선택. 매핑 외 조사는 원본 그대로."""
    pair = _JOSA_PAIRS.get(particle)
    if not pair or not word:
        return particle
    return pair[0] if _has_jongseong(word[-1]) else pair[1]


def _r8_replace(m: "re.Match[str]") -> str:
    """'(여러|많은|각|모든…) ~들[조사]' → '$prefix $word[받침-맞춤 조사]'."""
    prefix = m.group(1)
    word = m.group(2)
    particle = m.group(3) or ""
    new_particle = _josa_swap(word, particle)
    return f"{prefix} {word}{new_particle}"


# ──────────────────────────────────────────────────────────────────────
# 자동 적용 규칙 (severity="auto") — 신뢰도 높음
# ──────────────────────────────────────────────────────────────────────
#
# 조사 lookahead: 한국어 받침 뒤에 붙는 주격·목적격·부사격 조사를 광범위 포함
_JOSA_AHEAD = r"(?=\s|$|[이가은는을를도만에과와으로의도뿐까,.!?:;)\]])"

AUTO_RULES = [
    {
        "id": "R1",
        "pattern": re.compile(r"(\S+?)[와과] 관련된"),
        "replace": r"\1 관련",
        "reason": "~와/과 관련된 → ~ 관련",
    },
    {
        "id": "R3",
        "pattern": re.compile(
            r"([가-힣A-Za-z]+)(?:할|될) 것으로 보(?:입니다|인다)"
        ),
        "replace": r"\1 예상",
        "reason": "~할/될 것으로 보입니다/보인다 → ~ 예상",
    },
    {
        "id": "R4",
        "pattern": re.compile(
            r"([가-힣A-Za-z]+(?:한|된|진)) 것으로 판단(?:됩니다|된다)"
        ),
        "replace": r"\1 판단",
        "reason": "~한/된/진 것으로 판단됩니다/된다 → ~ 판단",
    },
    {
        "id": "R5",
        "pattern": re.compile(r"(\S+)할 예정이었으나 이를 유예하였습니다"),
        "replace": r"\1 예정 → 유예",
        "reason": "~할 예정이었으나 이를 유예하였습니다 → ~ 예정 → 유예",
    },
    {
        "id": "R6a",
        "pattern": re.compile(r"([가-힣A-Za-z]+)하는 것이 필요(?:합니다|하다)"),
        "replace": _r6_need,
        "reason": "~하는 것이 필요합니다 → ~이/가 필요합니다 (받침 자동)",
    },
    {
        "id": "R6b",
        "pattern": re.compile(r"([가-힣A-Za-z]+)하는 것이 중요(?:합니다|하다)"),
        "replace": _r6_important,
        "reason": "~하는 것이 중요합니다 → ~이/가 중요합니다 (받침 자동)",
    },
    {
        "id": "R8",
        "pattern": re.compile(
            r"(여러|많은|각|모든|수많은|대부분의|다양한)\s+([가-힣A-Za-z]+?)들"
            r"(?:(이|가|은|는|을|를)|(?=\s|$|[,.!?:;)\]]))"
        ),
        "replace": _r8_replace,
        "reason": "복수 표현(여러/많은/각/모든…) 뒤의 '들' 제거 + 짝조사 받침 재맞춤",
    },
    {
        "id": "R12a",
        "pattern": re.compile(r"([A-Za-z]{2,})([가-힣])"),
        "replace": r"\1 \2",
        "reason": "영문 약어 뒤 한글 사이 띄어쓰기 추가",
    },
    {
        "id": "R12b",
        "pattern": re.compile(r"([가-힣])([A-Za-z]{2,})"),
        "replace": r"\1 \2",
        "reason": "한글 뒤 영문 약어 사이 띄어쓰기 추가",
    },
]


# ──────────────────────────────────────────────────────────────────────
# 검토 권장 규칙 (severity="review") — 자동 치환은 의미 손실 위험
# ──────────────────────────────────────────────────────────────────────

REVIEW_RULES = [
    {
        "id": "R2", "pattern": re.compile(r"(\S+)에 대한"),
        "reason": "~에 대한 — 가능하면 생략 또는 동사형 (예: '검토에 대한 보고' → '검토 보고')",
    },
    {
        "id": "R7", "pattern": re.compile(r"(\S+) 중 하나인"),
        "reason": "~ 중 하나인 — 가능하면 '대표적인' 또는 생략",
    },
    {
        "id": "R9", "pattern": re.compile(
            r"(사회적|경제적|정치적|행정적|조직적|기술적|사업적|업무적|제도적)\s*([가-힣]+)"
        ),
        "reason": "'-적' 접미사 — 빼도 의미가 살면 제거 (예: 사회적 현상 → 사회 현상)",
    },
    {
        "id": "R10", "pattern": re.compile(r"(\S+)의\s+(\S+)의\s+(\S+)"),
        "reason": "'의' 연쇄 — 한 문장에 '의' 2개 이상은 분리·압축 권장",
    },
]


# 한 문장 46자 초과 검출 — 마침표·물음표·느낌표·줄바꿈으로 분리
SENTENCE_SPLIT = re.compile(r"[^.?!\n]+[.?!]?")
MAX_SENTENCE_LENGTH = 46
# 검출 제외할 라인 prefix (마크다운 글머리·헤딩·인용·코드펜스·표)
SKIP_LINE_PREFIXES = ("- ", "* ", "#", ">", "|", "```")


def _find_line_col(text: str, pos: int) -> Tuple[int, int]:
    """절대 오프셋 pos를 (1-based line, 0-based column)으로 변환."""
    line = text.count("\n", 0, pos) + 1
    last_nl = text.rfind("\n", 0, pos)
    col = pos - (last_nl + 1) if last_nl >= 0 else pos
    return line, col


def optimize_text(
    text: str, *, auto_apply: bool = True
) -> Tuple[str, List[Suggestion]]:
    """텍스트에 글쓰기 최적화 규칙을 적용한다.

    Args:
        text: 입력 텍스트.
        auto_apply: True 면 자동 적용 규칙으로 텍스트를 치환. False 면 변경 없이
            제안만 수집 (dry-run).

    Returns:
        (수정된 텍스트, 제안 목록). 제안은 자동 적용·검토 권장·문장 길이 초과
        모두 포함.
    """
    suggestions: List[Suggestion] = []
    new_text = text

    for rule in AUTO_RULES:
        rid: str = rule["id"]
        pat: re.Pattern = rule["pattern"]
        repl = rule["replace"]
        reason: str = rule["reason"]

        for m in pat.finditer(new_text):
            before = m.group(0)
            after = pat.sub(repl, before, count=1)
            line, col = _find_line_col(new_text, m.start())
            suggestions.append(
                Suggestion(
                    rule_id=rid, severity="auto",
                    line=line, column=col,
                    before=before, after=after,
                    reason=reason,
                )
            )

        if auto_apply:
            new_text = pat.sub(repl, new_text)

    for rule in REVIEW_RULES:
        rid = rule["id"]
        pat = rule["pattern"]
        reason = rule["reason"]
        for m in pat.finditer(new_text):
            line, col = _find_line_col(new_text, m.start())
            suggestions.append(
                Suggestion(
                    rule_id=rid, severity="review",
                    line=line, column=col,
                    before=m.group(0), after=None,
                    reason=reason,
                )
            )

    for li, line_text in enumerate(new_text.splitlines(), start=1):
        stripped = line_text.lstrip()
        if not stripped or stripped.startswith(SKIP_LINE_PREFIXES):
            continue
        for m in SENTENCE_SPLIT.finditer(line_text):
            sentence = m.group(0).strip()
            if len(sentence) > MAX_SENTENCE_LENGTH:
                suggestions.append(
                    Suggestion(
                        rule_id="R11", severity="review",
                        line=li, column=m.start(),
                        before=sentence, after=None,
                        reason=f"문장 {len(sentence)}자 (>{MAX_SENTENCE_LENGTH}자) — 분리 후보",
                    )
                )

    return new_text, suggestions


def format_report(suggestions: List[Suggestion]) -> str:
    """제안 목록을 마크다운 리포트로 직렬화."""
    if not suggestions:
        return "# 글쓰기 최적화 리포트\n\n변경·검토 항목 없음.\n"

    n_auto = sum(1 for s in suggestions if s.severity == "auto")
    n_review = sum(1 for s in suggestions if s.severity == "review")

    out = ["# 글쓰기 최적화 리포트", ""]
    out.append(f"- 자동 적용: **{n_auto}건**")
    out.append(f"- 검토 권장: **{n_review}건**")
    out.append("")

    if n_auto:
        out.append("## 자동 적용된 변환")
        out.append("")
        out.append("| 위치 | 규칙 | 원본 | 변경 | 근거 |")
        out.append("|---|---|---|---|---|")
        for s in suggestions:
            if s.severity != "auto":
                continue
            loc = f"L{s.line}:{s.column}"
            out.append(
                f"| {loc} | {s.rule_id} | `{s.before}` | `{s.after}` | {s.reason} |"
            )
        out.append("")

    if n_review:
        out.append("## 검토 권장")
        out.append("")
        out.append("| 위치 | 규칙 | 원본 | 권고 |")
        out.append("|---|---|---|---|")
        for s in suggestions:
            if s.severity != "review":
                continue
            loc = f"L{s.line}:{s.column}"
            before = s.before if len(s.before) < 60 else s.before[:57] + "..."
            out.append(f"| {loc} | {s.rule_id} | `{before}` | {s.reason} |")
        out.append("")

    return "\n".join(out)


def main():
    p = argparse.ArgumentParser(
        description="공공기관 보고서 글쓰기 규칙 자동 적용 + 검토 권장 리포트"
    )
    p.add_argument("input", help="입력 파일 (.md, .txt). '-' 면 stdin")
    p.add_argument("--output", "-o", help="수정된 텍스트 출력 (기본: stdout)")
    p.add_argument("--report", "-r", help="검토 리포트 마크다운 출력 (기본: stderr)")
    p.add_argument("--dry-run", action="store_true",
                   help="자동 적용 없이 제안만 수집")
    args = p.parse_args()

    if args.input == "-":
        text = sys.stdin.read()
    else:
        text = Path(args.input).read_text(encoding="utf-8")

    new_text, suggestions = optimize_text(text, auto_apply=not args.dry_run)
    report = format_report(suggestions)

    if args.output:
        Path(args.output).write_text(new_text, encoding="utf-8")
    else:
        sys.stdout.write(new_text)

    if args.report:
        Path(args.report).write_text(report, encoding="utf-8")
        print(f"[리포트] {args.report}", file=sys.stderr)
    else:
        sys.stderr.write("\n" + report)


if __name__ == "__main__":
    main()
