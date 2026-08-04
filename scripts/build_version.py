# -*- coding: utf-8 -*-
"""산출물 파일명의 날짜·순번을 스크립트가 매긴다: `<이름>_<YYMMDD>_<NN>.hwpx`.

"수정할 때마다 날짜와 순번을 붙인다"는 규칙을 손으로 지키면 곧 어긋난다.
채움AI 2차 출제에서 2026-07-30 하루에 아홉 번 수정하고도 전부 `_260730_01`에
덮어썼다. 순번을 사람이 기억하는 대신 디스크를 세어 매긴다.

- `new_build()`   : 빌드 시작 단계에서 다음 순번 경로를 받는다
- `latest_build()`: 후속 단계에서 방금 만들어진 최신본을 찾는다 — 파일명을
                    하드코딩하면 순번이 올라갈 때마다 깨진다

확장자는 `ext` 인자로 바꾼다(기본 `.hwpx`).

    from build_version import new_build, latest_build
    out = new_build(dirpath, '제출_문항1_전지직렬')          # → ..._260730_10.hwpx
    src = latest_build(dirpath, '제출_문항1_전지직렬')        # → 가장 큰 날짜·순번
"""

import datetime
import os
import re

EXT = '.hwpx'


def _scan(dirpath, prefix, ext=EXT):
    """[(날짜 YYMMDD, 순번, 경로), ...] — 오래된 것부터."""
    pat = re.compile(r'^' + re.escape(prefix) + r'_(\d{6})_(\d{2})' + re.escape(ext) + r'$')
    found = []
    if not os.path.isdir(dirpath):
        return found
    for name in os.listdir(dirpath):
        m = pat.match(name)
        if m:
            found.append((m.group(1), int(m.group(2)), os.path.join(dirpath, name)))
    return sorted(found)


def new_build(dirpath, prefix, today=None, ext=EXT):
    """오늘 날짜의 최대 순번 + 1 경로. 오늘 빌드가 없으면 _01."""
    day = today or datetime.date.today().strftime('%y%m%d')
    used = [n for d, n, _ in _scan(dirpath, prefix, ext) if d == day]
    return os.path.join(dirpath, '%s_%s_%02d%s' % (prefix, day, (max(used) + 1) if used else 1, ext))


def latest_build(dirpath, prefix, ext=EXT):
    """날짜·순번이 가장 큰 경로. 없으면 FileNotFoundError."""
    found = _scan(dirpath, prefix, ext)
    if not found:
        raise FileNotFoundError('%s 의 빌드 파일이 없다: %s*%s' % (dirpath, prefix, ext))
    return found[-1][2]
