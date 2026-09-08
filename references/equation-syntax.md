# 한컴 수식 스크립트 문법 (`add_equation.py`)

`scripts/add_equation.py` 는 이미지가 아니라 **한컴 네이티브 수식 개체**
(`<hp:equation>`)를 넣는다. 한글에서 수식 편집기로 그대로 열리고 다시 편집할 수 있다.

문법 근거는 KS X 6101:2024 부속서 I (`ks-x-6101.md` §수식)이고, 아래 표의 토큰은
**한컴에서 실제로 열어 PDF 로 렌더해 확인한 것**만 적었다 (2026-09-08).

```bash
# 본문: 앵커 문구가 있는 문단 뒤에 새 문단으로
python scripts/add_equation.py in.hwpx -o out.hwpx --after "근의 공식" \
    --script "x = {-b +- sqrt{b^2 -4ac}} over {2a}"

# 표 셀 안에 (좌표는 cellAddr 격자 기준 — 규칙 33과 같다)
python scripts/add_equation.py in.hwpx -o out.hwpx --table 0 --row 1 --col 1 \
    --script "int _0 ^1 x^2 dx = 1 over 3"
```

- `--script` (필수) — 수식 스크립트. **대소문자를 구분한다** (`GAMMA` 는 Γ, `gamma` 는 γ).
- `--after` 또는 `--table/--row/--col` 중 하나로 위치를 정한다.
- `--size` — 글자 크기 pt (기본 10).
- 수식은 자기완결 개체라 header.xml·BinData·매니페스트 등록이 필요 없다.

## 검증한 토큰

| 종류 | 토큰 | 예시 → 결과 |
|---|---|---|
| 분수 | `{ } over { }` | `1 over 2` → ½ |
| 제곱근 | `sqrt { }` · `root n of x` | `sqrt {x+1}`, `root 3 of 8` |
| 위·아래 첨자 | `^` · `_` | `a_1 ^2`, `x^{n+1}` |
| 적분 | `int _아래 ^위` | `int _0 ^inf e^{-x} dx` |
| 합·곱 | `sum _{i=1} ^n` · `prod` | `sum _{i=1} ^n i` |
| 무한대 | `inf` | ∞ |
| 부등호 | `<` `>` `<=` `>=` | `a < b <= c` → a < b ≤ c |
| 연산 | `+-` `times` `div` `cdot` | `+-` → ±, `times` → × |
| 그리스 | `alpha` `beta` `delta` … / 대문자 `GAMMA` `SIGMA` | α β δ / Γ Σ |
| 괄호 (크기 자동) | `LEFT ( … RIGHT )` | 내용 높이에 맞춰 커진다 |
| 묶기 | `{ }` | 렌더에는 안 나오고 묶기만 한다 |
| 행 바꿈 | `#` | 여러 줄 수식 |
| 열 정렬 | `&` | 아래 함정 참조 |
| 배열 | `matrix{ … }` | 괄호 없는 배열 — 아래 함정 참조 |
| 공백 | `~` (넓게 `~~`) | 토큰 사이 간격 |

## 함정 셋

**1. `&` 는 글자가 아니라 열 구분자다.** `a & b` 를 쓰면 `&` 가 사라지고 두 항이
열로 갈린다. 논리곱 기호를 찍고 싶었다면 나오지 않는다.

```
"a < b <= c & d >= e > f"  →  a < b ≤ c  d ≥ e > f    (& 가 사라진다)
```

**2. `matrix` 는 괄호를 그리지 않는다.** 행렬 괄호가 필요하면 직접 감싼다.

```
matrix{1 & 0 # 0 & 1}                     →  1 0 / 0 1  (괄호 없음)
LEFT ( matrix{1 & 0 # 0 & 1} RIGHT )      →  괄호 있는 행렬
```

**3. XML 예약문자는 스크립트가 알아서 처리한다.** `<`, `>`, `&` 를 그대로 넘겨도
`escape_script()` 가 이스케이프하므로 XML 이 깨지지 않는다. 위 두 함정은 이스케이프
문제가 아니라 수식 문법 자체의 동작이다.

## 개체 봉투

실물 문항지(수학)에서 뽑은 구조를 따랐다.

- `version="Equation Version 60"`, `font="HYhwpEQ"` (한컴 수식 전용 글꼴)
- `baseUnit` = 글자 크기 × 100 (1000 = 10pt)
- `treatAsChar="1"` — **수식은 1이 맞다.** 글자 크기의 인라인 개체라 쪽을 넘길 일이
  없다. 쪽을 넘길 수 있는 표가 0이어야 하는 것과 반대다 (규칙 33, `table_page_fit.py`).
- `hp:sz` 의 width/height 는 스크립트 길이로 근사한다. 한글이 열 때 다시 렌더하므로
  정확할 필요는 없다.
