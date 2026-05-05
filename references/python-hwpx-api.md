# python-hwpx API 레퍼런스

`python-hwpx` 라이브러리 (PyPI: `python-hwpx`, import: `hwpx`) 의 핵심 API 시그니처와
주의사항을 빠르게 확인하기 위한 노트. 본 스킬은 라이브러리의 일부 기능만 의존하며,
대부분의 hwpx 조작은 자체 헬퍼 (`hwpx_modifier`, `hwpx_form_filler`, `md2hwpx`,
`hwpx_helpers`) 와 ZIP/lxml 직접 조작으로 처리한다. 이 문서는 `TextExtractor` /
`ObjectFinder` 등 라이브러리에 위임하는 부분과, 향후 메이저 업그레이드 시 활용
가능한 신규 API 를 함께 정리한다.

## 버전 매트릭스

| python-hwpx | 본 스킬 호환 | 비고 |
|---|---|---|
| 1.9 (현재 설치) | ✅ 완전 호환 | `TextExtractor`, `ObjectFinder`, `FoundElement` 만 사용 |
| 2.0 ~ 2.4 | ⚠️ 부분 호환 | `HwpxDocument` 등 신규 클래스 추가, 기존 API 시그니처 일부 변경 |
| 2.5 ~ 2.9.1 | ⚠️ 미검증 | airmang/hwpx-skill 기준 버전. 본 스킬 회귀 테스트 미완료 |
| 1.x 이전 | ❌ 비호환 | API 미존재 |

> 글로벌 업그레이드는 `.test/2026MMDD-airmang-스킬반영/` 류 별도 venv 에서 회귀
> (학급평가보고서·학급독서통계·사업계획서 strict PASS) 통과 후 결정한다.

## 1.9 (현재 설치) — 사용 중인 API

### 설치와 import

```bash
pip install -U python-hwpx lxml
```

```python
from hwpx import TextExtractor, ObjectFinder, FoundElement, ParagraphInfo, SectionInfo
```

`hwpx` top-level 노출 심볼: `__version__`, `DEFAULT_NAMESPACES`, `ParagraphInfo`,
`SectionInfo`, `TextExtractor`, `FoundElement`, `ObjectFinder`.

### TextExtractor

```python
with TextExtractor("input.hwpx") as ext:
    full = ext.extract_text(include_nested=True)        # 표 셀까지 포함
    for sec in ext.iter_sections():
        for para in ext.iter_paragraphs(sec, include_nested=True):
            print(para.path, para.text(object_behavior="nested"))
```

메서드 (1.9):

- `extract_text(*, include_nested=False, object_behavior="skip", skip_empty=False) -> str`
- `iter_sections() -> Iterable[SectionInfo]`
- `iter_paragraphs(section, *, include_nested=False) -> Iterable[ParagraphInfo]`
- `iter_document_paragraphs(*, include_nested=False) -> Iterable[ParagraphInfo]`
- `paragraph_text(paragraph, *, object_behavior="skip") -> str`
- `open()` / `close()` (with-block 권장)

`object_behavior` 값: `"skip"` (기본, 표 등 inline object 본문 무시), `"nested"`
(객체 내부 문단까지 직렬화), `"placeholder"` (객체 위치를 기호로 표시).

### ObjectFinder

```python
with TextExtractor("input.hwpx") as ext:
    finder = ObjectFinder(ext)
    tables = finder.find_all("hp:tbl")
    for t in tables[:3]:
        print(t.path, t.attrib)
```

메서드:

- `find_all(tag) -> list[FoundElement]`
- `find_first(tag) -> FoundElement | None`
- `iter(tag) -> Iterable[FoundElement]`
- `iter_annotations() -> Iterable[FoundElement]`

`tag` 는 namespace prefix 포함 형태 (`"hp:tbl"`, `"hp:p"`, `"hp:tcell"` 등).
`DEFAULT_NAMESPACES` 가 자동 매핑된다.

### 본 스킬에서 사용 위치

- `scripts/text_extract.py` — `TextExtractor` 만 wrap
- 그 외 모든 hwpx 조작 (`hwpx_modifier`, `hwpx_form_filler`, `md2hwpx`,
  `build_hwpx`, `table_calc`, `hwpx_helpers`) 은 자체 구현 + `lxml` 사용

## 2.x — 향후 활용 가능한 신규 API (참고)

> airmang/hwpx-skill (2.9.0 기준) 의 `references/api.md` 발췌. 본 스킬은 아직 미사용.

### HwpxDocument (2.0+ 신규 클래스)

```python
from hwpx import HwpxDocument

doc = HwpxDocument.new()                                # 빈 문서 생성
doc.add_paragraph("자동 생성 문서")
table = doc.add_table(2, 2)
table.set_cell_text(0, 0, "학년", logical=True)         # 병합 셀은 logical=True
table.set_cell_text(0, 1, "1학년")
doc.save_to_path("output.hwpx")                         # 기본 저장 API

with HwpxDocument.open("input.hwpx") as doc:
    doc.add_paragraph("추가 문단")
    doc.save_to_path("edited.hwpx")
```

- `HwpxDocument.new()` / `HwpxDocument.open(source)` / `to_bytes()`
- `save()` 는 deprecated wrapper, `save_to_path(path)` 사용
- `add_table(rows, cols, *, section=None, width=None, height=None, ...) -> HwpxOxmlTable`
- 표 객체의 `set_cell_text(row, col, text, *, logical=False, split_merged=False)`

### 메모 추가 (2.x)

```python
memo, anchor_paragraph, field_value = doc.add_memo_with_anchor(
    "표현을 한 번 더 확인하세요.",
    paragraph=paragraph,
    author="검토자",
)
```

`add_memo_with_anchor(text="", *, paragraph=None, section=None, paragraph_text=None,
author=None, ...) -> tuple[HwpxOxmlMemo, HwpxOxmlParagraph, str]`

→ 학생 작품 자동 첨삭 헬퍼로 활용 가치 있음.

### 스타일 필터 치환 (2.x)

```python
with HwpxDocument.open("input.hwpx") as doc:
    red_runs = doc.find_runs_by_style(text_color="#FF0000")
    replaced = doc.replace_text_in_runs(
        "TODO", "DONE",
        text_color="#FF0000", underline_type="SOLID", limit=3,
    )
    doc.save_to_path("output.hwpx")
```

→ "빨간 TODO만 검정으로" 같은 조건부 일괄 치환에 강점.

## 마이그레이션 노트 (1.9 → 2.x)

본 스킬을 2.x 로 올릴 때 영향받는 표면:

1. `scripts/text_extract.py` — `TextExtractor` 만 사용. 1.9 ↔ 2.x 시그니처 변화
   거의 없음 (호환성 높음).
2. 자체 hwpx 조작 코드 (`hwpx_modifier` 등) — 라이브러리 의존도 0, 영향 없음.
3. **신규 활용 후보**: `HwpxDocument` 의 `replace_text_in_runs(text_color=, ...)` 와
   `add_memo_with_anchor()` 는 본 스킬에 동등 기능이 없으므로 추가 헬퍼로 도입 가치.

권장 회귀 테스트 셋:

- `.test/20260418-rhwp-분석/학급평가보고서/2026_2학기_5학년3반_학급평가_종합보고서.hwpx`
- `.test/20260418-rhwp-분석/학교사업계획/2026_AI디지털교육혁신_사업계획서.hwpx`
- `.test/20260418-rhwp-분석/학급독서통계/2026_1학기_학급독서기록_통계표.hwpx`

각각에 대해 `verify_hwpx.py --strict --json` 로 polaris-dvc strict PASS 확인.

## 참고

- airmang/hwpx-skill (Apache-2.0): https://github.com/airmang/hwpx-skill
- python-hwpx: https://github.com/airmang/python-hwpx
