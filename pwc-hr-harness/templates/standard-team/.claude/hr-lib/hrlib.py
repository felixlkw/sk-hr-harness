# -*- coding: utf-8 -*-
"""pwc-hr-harness 기본 인력계획 팀 공통 라이브러리.

Skill 스크립트가 함께 쓰는 기능:
  - 데이터 읽기: 열 대응표(①_열대응표.json)로 회사 파일의 열을 기본 열에 맞춰 읽는다
  - 데이터 사전 공통 규칙: 단위 · 기준일 · 집계 행 · 빈칸≠0 · 출처 · RPU 코드 이음매
  - 멈춤: CP(체크포인트)와 예외를 같은 4요소 형식으로 알린다 (무엇을 정해야 하는가 · 근거 · 선택지 · 답 형식)
  - 승인 읽기: CP 자료 파일의 「## 승인」 표
  - 결과 저장: xlsx(시트마다 출처 열) · md

판단은 담지 않는다. 판단 규칙은 각 Skill의 SKILL.md와, 사람이 승인한 파라미터 파일에 둔다.
"""
import os, re, sys, json, datetime, argparse

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:  # 설치 안내는 /hr-init 이 맡는다
    print('openpyxl 이 없습니다. `python -m pip install openpyxl` 로 설치한 뒤 다시 실행하세요.', file=sys.stderr)
    sys.exit(2)

HERE = os.path.dirname(os.path.abspath(__file__))
EXIT_STOP = 3          # 멈춤(CP 또는 예외)으로 끝날 때 종료 코드
ANSWER_FORMAT = '확정 · 추가 확인 · 보류 · 이관'
ANSWER_NOTE = '답은 넷 중 하나: 확정 (다음 단계로) · 추가 확인 (자료 · 근거를 더 받음) · 보류 (이번에 정하지 않음) · 이관 (다른 승인자 · 부서로 넘김) · 바꾼 내용은 「확정」 + 이유 칸에 적음'
GATE_ACCEPT = '관문 인정 (결정로그에 이유)'   # 관문 멈춤에서 사람이 미해결을 그대로 두고 가기로 할 때 (「예외」 = 데이터 이상과 다른 말)

# ---------------------------------------------------------------------------
# 설정 (기준 연도 등) · 자사 데이터 연도가 다르면 .claude/hr-lib/config.json 만 고친다
# ---------------------------------------------------------------------------
DEFAULT_CONFIG = {
    "base_year": 2025,            # 현원 기준 연도 (연말)
    "target_year": 2027,          # 계획 대상 연도
    "hist_years": [2023, 2024, 2025],
    "series_years": [2021, 2022, 2023, 2024, 2025],
    "fte_band": [0.8, 1.2],       # 업무량 검증: 2025 FTE ÷ 현원 허용 범위 (가설·예시)
    "small_rpu": 40,              # 현원 이 값 미만이면 「표본 작음」
    "recon_bands": [0.03, 0.07],  # 대사: 일치 · 설명 · 재검토 경계
    "approvers": {},              # 자사 승인자 직책 (자사 예 · 가설): {"CP2": "인사팀장", "CP5": "인사담당 임원"}
}


def config():
    p = os.path.join(HERE, 'config.json')
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(p):
        with open(p, encoding='utf-8') as f:
            cfg.update(json.load(f))
    return cfg


# ---------------------------------------------------------------------------
# 기본 표 정의 (K-Co. 양식 = 기본 열 이름) · 자사 파일은 열 대응표로 이 이름에 맞춘다
# ---------------------------------------------------------------------------
def _years(cfg, key):
    return [str(y) for y in cfg[key]]


def tables(cfg=None):
    cfg = cfg or config()
    by, ty = cfg['base_year'], cfg['target_year']
    sy = _years(cfg, 'series_years')
    return {
        'org': dict(no=2, title='조직·인원 현황', key='RPU 코드', point=True,
                    cols=['본부', 'RPU 코드', 'RPU(산정 단위)', 'G1', 'G2', 'G3', 'G4', 'G5', '계']),
        'trend': dict(no=2, title='연도별 추이', key='연도',
                      cols=['연도', '인원(명)', '매출', '영업이익', '총인건비']),
        'retire': dict(no=2, title='정년 도래 예정', key='RPU 코드',
                       cols=['RPU 코드', str(by + 1), str(by + 2)]),
        'age': dict(no=2, title='연령대 분포', key='RPU 코드', point=True, optional=True,
                    cols=['RPU 코드', '계']),
        'hist': dict(no=3, title='입퇴사 이력', key='RPU 코드',
                     cols=['연도', 'RPU 코드', '직급', '기초', '입사', '자발퇴직', '정년퇴직', '기말']),
        'histsum': dict(no=3, title='연도별 요약', key='연도', optional=True,
                        cols=['연도', '기초', '자발퇴직']),
        'work': dict(no=4, title='업무량 조사', key='RPU 코드',
                     cols=['RPU 코드', '업무 ID', '업무명', '연간 빈도(회)', '1회 소요시간(시간)',
                           '연간 소요시간(시간)', '태스크 판정']),
        'assume': dict(no=4, title='산정 가정', key='항목', cols=['항목', '값']),
        'skill': dict(no=5, title='스킬 인벤토리', key='RPU 코드', point=True,
                      cols=['RPU 코드', '도메인', '그룹', '스킬', '요구 레벨 2028', 'L0', 'L1', 'L2', 'L3', 'L4',
                            '인원 계', '2028 요구 충족 인원']),
        'skillg': dict(no=5, title='직급별 보유 분포', key='RPU 코드', point=True, optional=True,
                       cols=['RPU 코드', '직급', '스킬', '요구 레벨 2028', 'L0', 'L1', 'L2', 'L3', 'L4', '인원 계']),
        'hcs': dict(no=6, title='RPU별 연말 인원', key='RPU 코드', cols=['RPU 코드'] + sy),
        'drv': dict(no=6, title='조직별 인력동인 시계열', key='RPU 코드',
                    cols=['RPU 코드', '유형', '하위 유형', '동인 후보', '단위'] + sy + [f'{ty} 계획']),
    }


# ---------------------------------------------------------------------------
# 경로
# ---------------------------------------------------------------------------
def find_root(start=None):
    d = os.path.abspath(start or os.getcwd())
    while True:
        if os.path.exists(os.path.join(d, 'CLAUDE.md')) and os.path.isdir(os.path.join(d, '.claude')):
            return d
        nd = os.path.dirname(d)
        if nd == d:
            return os.path.abspath(start or os.getcwd())
        d = nd


class WS:
    """작업 폴더 하나 + 회사 하나 (data/<co>/ → outputs/<co>/)."""

    def __init__(self, root, co, data=None):
        self.root = root
        self.co = co
        self.data = os.path.join(root, 'data', data or co)
        self.out = os.path.join(root, 'outputs', co)
        os.makedirs(self.out, exist_ok=True)
        self.cfg = config()

    def o(self, name):
        return os.path.join(self.out, name)

    def rel(self, p):
        try:
            return os.path.relpath(p, self.root).replace('\\', '/')
        except ValueError:
            return p


def common_args(desc):
    ap = argparse.ArgumentParser(description=desc)
    ap.add_argument('--co', default='K-co', help='회사 폴더 이름 (data/<co>/ · outputs/<co>/)')
    ap.add_argument('--data', default=None, help='입력 폴더 이름을 바꿀 때 (예: K-co-검증). 결과는 outputs/<co>/')
    ap.add_argument('--root', default=None, help='작업 폴더 (기본: CLAUDE.md 가 있는 상위 폴더)')
    ap.add_argument('--allow', action='append', default=[],
                    help='사람이 답한 예외 처리. 형식 예외ID=선택 (결정로그에 남긴 뒤 씀)')
    return ap


def ws_from(args):
    return WS(find_root(args.root), args.co, args.data)


def allowed(args, key):
    for a in args.allow:
        if '=' in a:
            k, v = a.split('=', 1)
            if k.strip() == key:
                return v.strip()
    return None


# ---------------------------------------------------------------------------
# 멈춤 (CP · 예외 같은 형식)
# ---------------------------------------------------------------------------
class Stop(Exception):
    def __init__(self, kind, sid, what, evidence, options, note=''):
        super().__init__(what)
        self.kind, self.sid, self.what = kind, sid, what
        self.evidence, self.options, self.note = evidence, options, note


def write_stop(ws, st):
    by = os.path.basename(sys.argv[0])   # 같은 스크립트(어느 명령이든)가 성공하면 그 멈춤은 지난 것
    lines = [f'<!-- by: {by} -->', f'# 멈춤 · {st.kind} {st.sid}', '',
             f'- 시각: {datetime.datetime.now():%Y-%m-%d %H:%M}',
             f'- 회사: {ws.co} · 입력 폴더: {ws.rel(ws.data)}', '',
             '| 항목 | 내용 |', '|---|---|',
             f'| 무엇을 정해야 하는가 | {st.what} |',
             f'| 근거 | {"<br>".join(st.evidence)} |',
             f'| 선택지 | {"<br>".join(f"{i + 1}. {o}" for i, o in enumerate(st.options))} |',
             f'| 답 형식 | {ANSWER_FORMAT} |']
    lines += ['', f'- {ANSWER_NOTE}']
    if st.kind == '관문':
        lines += [f'- 관문은 넘긴 담당이 고친 뒤 다시 검증하는 것이 기본이다. 고치지 않고 가려면 「{GATE_ACCEPT}」 · 결정로그 체크포인트 칸 「관문 인정」']
    if st.note:
        lines += ['', st.note]
    lines += ['', '> 리드는 이 화면을 사람에게 보여 주고, 답을 `결정로그.md`에 한 행으로 남긴 뒤 다음으로 간다.',
              '> 예외 답을 반영할 때는 같은 명령에 `--allow ' + st.sid + '=<선택>` 을 붙여 다시 실행한다.' if st.kind == '예외' else
              '> CP 답은 해당 CP 자료 파일의 「## 승인」 표에 적은 뒤 다음 명령을 실행한다.']
    with open(ws.o('_멈춤.md'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')


def run_main(fn, ws_getter):
    """스크립트 진입점 공통 처리: Stop → _멈춤.md + 종료 코드 3."""
    args = ws_getter()
    ws = ws_from(args)
    try:
        fn(ws, args)
        p = ws.o('_멈춤.md')
        by = os.path.basename(sys.argv[0])   # 같은 스크립트(어느 명령이든)가 성공하면 그 멈춤은 지난 것
        if os.path.exists(p) and f'<!-- by: {by} -->' in read(p):   # 자기가 남긴 멈춤만 지운다
            os.remove(p)
    except Stop as st:
        write_stop(ws, st)
        print(json.dumps(dict(status='멈춤', kind=st.kind, id=st.sid, what=st.what,
                              evidence=st.evidence, options=st.options, answer=ANSWER_FORMAT,
                              file=ws.rel(ws.o('_멈춤.md'))), ensure_ascii=False, indent=1))
        sys.exit(EXIT_STOP)


# ---------------------------------------------------------------------------
# 마크다운 표 · 승인
# ---------------------------------------------------------------------------
def md_tables(text):
    """[(가장 가까운 제목, 머리 열 목록, 행 목록[dict])]"""
    out, head = [], ''
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        ln = lines[i].strip()
        if ln.startswith('#'):
            head = ln.lstrip('#').strip()
        if ln.startswith('|') and i + 1 < len(lines) and re.match(r'^\|\s*:?-{2,}', lines[i + 1].strip()):
            cols = [c.strip() for c in ln.strip('|').split('|')]
            rows = []
            j = i + 2
            while j < len(lines) and lines[j].strip().startswith('|'):
                cells = [c.strip() for c in lines[j].strip().strip('|').split('|')]
                cells += [''] * (len(cols) - len(cells))
                rows.append(dict(zip(cols, cells)))
                j += 1
            out.append((head, cols, rows))
            i = j
            continue
        i += 1
    return out


def approval_block(cp, approver=''):
    approver = config().get('approvers', {}).get(cp, approver)   # 자사 승인자 직책은 config.json 으로 바꾼다
    return '\n'.join(['## 승인', '',
                      f'- 답 형식: {ANSWER_FORMAT} · 승인자는 직책만 (이름 없이) · 「확정」일 때만 다음 단계로 감',
                      '', '| CP | 답 | 승인자 (직책) | 승인일 | 이유 한 줄 |', '|---|---|---|---|---|',
                      f'| {cp} |  | {approver} |  |  |', ''])


def approval(path, cp):
    """CP 자료 파일의 승인 상태. 반환 dict(ok, 답, 승인자, 승인일) · 파일이 없으면 ok=False."""
    if not os.path.exists(path):
        return dict(ok=False, 답='', 승인자='', 승인일='', missing=True)
    with open(path, encoding='utf-8') as f:
        text = f.read()
    for head, cols, rows in md_tables(text):
        if '승인' in head and 'CP' in cols and '답' in cols:
            for r in rows:
                if r.get('CP', '').startswith(cp):
                    ans = r.get('답', '')
                    who = r.get('승인자 (직책)', '')
                    return dict(ok=ans.startswith('확정') and bool(who.strip()), 답=ans, 승인자=who,
                                승인일=r.get('승인일', ''))
    return dict(ok=False, 답='', 승인자='', 승인일='')


def require_cp(ws, cp, fname, what_next):
    a = approval(ws.o(fname), cp)
    if not a['ok']:
        raise Stop('CP', cp, f'{cp} 승인 기록이 없어 {what_next}을(를) 시작할 수 없음',
                   [f'{fname} 「## 승인」 표 · 답={a["답"] or "(빈칸)"} · 승인자={a["승인자"] or "(빈칸)"}'],
                   [f'{fname} 를 사람에게 보여 주고 승인 받기', '앞 단계로 돌아가 자료 다시 만들기'])
    return a


# ---------------------------------------------------------------------------
# 데이터 읽기
# ---------------------------------------------------------------------------
def norm(v):
    if v is None:
        return ''
    s = str(v).strip()
    if isinstance(v, float) and v.is_integer() and re.fullmatch(r'\d{4}\.0', s):
        s = s[:-2]
    return re.sub(r'\s+', ' ', s)


BASIS_RE = [re.compile(r'(\d{4}-\d{2}-\d{2})'), re.compile(r'(\d{4})\s*(?:년\s*)?(자가진단|조사)')]


def sheet_basis(ws_):
    """시트 맨 위 3줄 안내문에서 시점 기준(날짜 또는 'YYYY 자가진단')을 찾는다. 기간(2021~2025)은 기준으로 보지 않는다."""
    for r in ws_.iter_rows(min_row=1, max_row=3, values_only=True):
        for v in r:
            if isinstance(v, str):
                t = re.sub(r'\d{4}\s*~\s*\d{4}', '', v)
                for rx in BASIS_RE:
                    m = rx.search(t)
                    if m:
                        return m.group(0).strip()
    return None


def find_header(ws_, want, max_scan=12):
    """기본 열(또는 대응 열) 이름이 가장 많이 들어 있는 행 = 표 머리 행."""
    best, best_n = None, 0
    for i, r in enumerate(ws_.iter_rows(min_row=1, max_row=max_scan, values_only=True), start=1):
        names = [norm(v) for v in r]
        n = sum(1 for w in want if w in names)
        if n > best_n:
            best, best_n = (i, names), n
    return best, best_n


def is_aggregate(first, keyval):
    f = norm(first)
    k = norm(keyval)
    return (not k) or f.endswith('소계') or f.endswith('합계') or f.startswith('주:') or k in ('합계', '소계', '계')


def load_map(ws):
    p = ws.o('①_열대응표.json')
    if not os.path.exists(p):
        raise Stop('CP', 'CP1', '열 대응표가 없어 데이터를 읽을 수 없음',
                   [ws.rel(p) + ' 없음'], ['wfp-step1-strategy 의 scan 을 먼저 실행', '자료 파일 위치 확인'])
    with open(p, encoding='utf-8') as f:
        return json.load(f)


def save_map(ws, m):
    with open(ws.o('①_열대응표.json'), 'w', encoding='utf-8') as f:
        json.dump(m, f, ensure_ascii=False, indent=1)


def load(ws, key, m=None, need=True):
    """기본 표 key 를 읽어 [dict(기본열 → 값, _ref)] 로 돌려준다. 집계 행은 빼고, 빈칸은 None 그대로 둔다."""
    m = m or load_map(ws)
    t = m['tables'].get(key)
    spec = tables(ws.cfg)[key]
    if not t or not t.get('file'):
        if need and not spec.get('optional'):
            raise Stop('예외', f'표없음-{key}', f'기본 표 「{spec["title"]}」(파일 {spec["no"]})에 대응하는 자사 표가 없음',
                       ['①_열대응표.json · tables.' + key], ['자료를 data 폴더에 넣고 scan 다시 실행', '이 표를 쓰는 방식을 생략 (보고서에 누락 명시)'])
        return []
    wb = openpyxl.load_workbook(os.path.join(ws.data, t['file']), data_only=True)
    sh = wb[t['sheet']]
    hrow = t['header_row']
    names = [norm(c.value) for c in sh[hrow]]
    colmap = t['cols']  # 기본열 → 자사열
    idx = {}
    for std, src in colmap.items():
        if src and src in names:
            idx[std] = names.index(src)
    keysrc = colmap.get(spec['key'])
    kpos = names.index(keysrc) if keysrc in names else None
    rows = []
    for rn, r in enumerate(sh.iter_rows(min_row=hrow + 1, values_only=True), start=hrow + 1):
        if all(v is None for v in r):
            continue
        keyval = r[kpos] if kpos is not None and kpos < len(r) else None
        if is_aggregate(r[0], keyval):
            continue
        d = {std: (r[i] if i < len(r) else None) for std, i in idx.items()}
        d['_ref'] = f'파일 {spec["no"]} · {t["sheet"]} · {rn}행'
        d['_row'] = rn
        rows.append(d)
    return rows


def aggregates(ws, key, m=None):
    """집계 행(소계 · 합계)과 RPU 행 합을 대조. [(행 이름, 파일 값, RPU 행 합, 셀)]"""
    m = m or load_map(ws)
    t = m['tables'][key]
    wb = openpyxl.load_workbook(os.path.join(ws.data, t['file']), data_only=True)
    sh = wb[t['sheet']]
    hrow = t['header_row']
    names = [norm(c.value) for c in sh[hrow]]
    tot_src = t['cols'].get('계')
    if tot_src not in names:
        return []
    ti = names.index(tot_src)
    out, group, total = [], 0, 0
    for rn, r in enumerate(sh.iter_rows(min_row=hrow + 1, values_only=True), start=hrow + 1):
        f = norm(r[0])
        v = r[ti] if ti < len(r) else None
        if f.endswith('소계'):
            out.append((f, v, group, f'{t["sheet"]}!{openpyxl.utils.get_column_letter(ti + 1)}{rn}'))
            group = 0
        elif f.endswith('합계'):
            out.append((f, v, total, f'{t["sheet"]}!{openpyxl.utils.get_column_letter(ti + 1)}{rn}'))
        elif isinstance(v, (int, float)) and norm(r[1] if len(r) > 1 else ''):
            group += v
            total += v
    return [x for x in out if x[1] is not None]


def num(v, what, ref):
    """숫자 칸 읽기. 빈칸은 0으로 읽지 않는다 (공통 규칙 4)."""
    if v is None or (isinstance(v, str) and not v.strip()):
        raise ValueError(f'빈칸: {what} ({ref})')
    if isinstance(v, str):
        s = v.strip().replace(',', '')
        if s.endswith('%'):
            return float(s[:-1]) / 100
        return float(s)
    return float(v)


def assume_value(rows, label, default=None):
    for r in rows:
        if label in norm(r.get('항목')):
            v = r.get('값')
            if isinstance(v, str) and v.strip().endswith('%'):
                return float(v.strip()[:-1]) / 100
            if isinstance(v, (int, float)):
                return float(v)
            m = re.search(r'[\d.]+', str(v))
            return float(m.group(0)) if m else default
    return default


# ---------------------------------------------------------------------------
# 계산 도구
# ---------------------------------------------------------------------------
def pearson(xs, ys):
    n = len(xs)
    mx, my = sum(xs) / n, sum(ys) / n
    sx = sum((x - mx) ** 2 for x in xs)
    sy = sum((y - my) ** 2 for y in ys)
    if sx == 0 or sy == 0:
        return None
    return sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / (sx * sy) ** 0.5


def regress(xs, ys):
    n = len(xs)
    mx, my = sum(xs) / n, sum(ys) / n
    b = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / sum((x - mx) ** 2 for x in xs)
    a = my - b * mx
    ss = sum((y - my) ** 2 for y in ys)
    r2 = 1 - sum((y - (a + b * x)) ** 2 for x, y in zip(xs, ys)) / ss if ss else None
    return a, b, r2


def r1(x, nd=1):
    return None if x is None else round(x, nd)


# ---------------------------------------------------------------------------
# 결과 저장
# ---------------------------------------------------------------------------
HEAD_FILL = PatternFill('solid', fgColor='FBE3D6')
TITLE_FONT = Font(bold=True, size=12)


def write_xlsx(path, sheets, script, inputs):
    """sheets = [dict(name, title, head[], rows[[]], notes[])] · 모든 시트에 스크립트 · 입력 파일 기록."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    stamp = f'만든 스크립트: {script} · 입력: {", ".join(inputs)} · {datetime.datetime.now():%Y-%m-%d %H:%M}'
    for s in sheets:
        sh = wb.create_sheet(s['name'][:31])
        sh.cell(1, 1, s.get('title', s['name'])).font = TITLE_FONT
        sh.cell(2, 1, stamp)
        r = 3
        for n in s.get('notes', []):
            sh.cell(r, 1, n)
            r += 1
        r += 1
        for j, h in enumerate(s['head'], start=1):
            c = sh.cell(r, j, h)
            c.font = Font(bold=True)
            c.fill = HEAD_FILL
            c.alignment = Alignment(wrap_text=True, vertical='center')
        for row in s['rows']:
            r += 1
            for j, v in enumerate(row, start=1):
                c = sh.cell(r, j, v)
                if isinstance(v, float):   # 값은 원 정밀도로 저장 · 보이는 자리만 줄임 (다음 단계 합계 오차 방지)
                    c.number_format = '0.0000' if abs(v) < 1 else '#,##0.0'
        for j, h in enumerate(s['head'], start=1):
            sh.column_dimensions[openpyxl.utils.get_column_letter(j)].width = max(10, min(46, len(str(h)) * 2 + 4))
    wb.save(path)
    # 보고서 모양 (2026-11-16): 「요약」 · 「○○표」 · 「산식·가정·출처」 를 앞뒤에 붙임 · 위 시트는 「첨부 데이터」로 그대로 (다음 단계 입력 계약)
    # 실패해도 결과 저장 · 계산은 끝난 상태 (경고만)
    try:
        import report_xlsx
        report_xlsx.decorate(path)
    except Exception as e:   # noqa: BLE001
        print(f'[알림] 보고서 시트를 붙이지 못함 ({os.path.basename(path)} · {type(e).__name__}: {e}) · 첨부 데이터는 그대로 저장됨', file=sys.stderr)


def side_outputs(ws, *kinds):
    """보고용 파일 (2026-11-16): gap = ④_Gap분석.xlsx · docx = 리포트_<회사>_<연도>.docx · 실패해도 단계는 그대로 (경고만)."""
    import glob
    try:
        import report_xlsx
        for k in kinds:
            if k == 'gap':
                report_xlsx.build_gap(ws.out)
            elif k == 'docx':
                if report_xlsx.build_docx(ws.out) is None and glob.glob(os.path.join(ws.out, '리포트_*.md')):
                    print('[알림] 워드 리포트를 만들지 못함 · python-docx 가 없으면 `python -m pip install python-docx`', file=sys.stderr)
    except Exception as e:   # noqa: BLE001
        print(f'[알림] 보고용 파일 ({", ".join(kinds)}) 을 만들지 못함 · {type(e).__name__}: {e} · 계산 결과는 그대로', file=sys.stderr)


def is_report_sheet(name):
    """보고서 시트 이름 규칙 (report_xlsx.is_report_sheet 와 같음): 요약 · 산식·가정·출처 · 「표」로 끝나는 시트"""
    return name in ('요약', '산식·가정·출처') or name.endswith('표')


def data_sheets(wb):
    """결과 엑셀의 첨부 데이터 시트 (2행이 「만든 스크립트」) · 보고서 시트 (요약 · ○○표 · 산식·가정·출처) 는 뺌."""
    return [s for s in wb.worksheets if str(s.cell(2, 1).value or '').startswith('만든 스크립트')]


def read_xlsx_table(path, sheet=None):
    """write_xlsx 로 만든 결과 파일을 다시 읽는다 → (head, rows[dict]) · sheet 를 안 주면 첫 첨부 데이터 시트."""
    wb = openpyxl.load_workbook(path, data_only=True)
    sh = wb[sheet] if sheet else (data_sheets(wb) or wb.worksheets)[0]
    rows = list(sh.iter_rows(values_only=True))
    hi = None
    for i, r in enumerate(rows):
        if r and r[0] is not None and i > 1 and all(v is None for v in rows[i - 1]) and any(v is not None for v in r[1:]):
            hi = i
            break
    if hi is None:
        return [], []
    head = [norm(v) for v in rows[hi]]
    out = []
    for r in rows[hi + 1:]:
        if all(v is None for v in r):
            continue
        out.append(dict(zip(head, r)))
    return head, out


def write_md(path, text):
    with open(path, 'w', encoding='utf-8') as f:
        f.write(text.rstrip() + '\n')


def md_row(cells):
    return '| ' + ' | '.join('' if c is None else str(c).replace('|', '/').replace('\n', ' ') for c in cells) + ' |'


def md_table(head, rows):
    return '\n'.join([md_row(head), '|' + '---|' * len(head)] + [md_row(r) for r in rows])


def append_log(path, row, head):
    """결정로그 · 검증로그에 한 행 추가 (지우지 않고 아래로)."""
    if not os.path.exists(path):
        write_md(path, f'# {os.path.splitext(os.path.basename(path))[0]}\n\n' + md_table(head, []))
    with open(path, encoding='utf-8') as f:
        text = f.read().rstrip('\n')
    n = 1 + sum(1 for ln in text.splitlines() if re.match(r'^\|\s*\d+\s*\|', ln))
    with open(path, 'w', encoding='utf-8') as f:
        f.write(text + '\n' + md_row([n] + list(row)) + '\n')
    return n


def is_blank_form(text):
    """양식 그대로인가 (스크립트 칸 · 사람 칸이 모두 비어 있음). step3 rates · 뷰어 · init 이 같은 기준을 쓴다.
    - auto 블록이 있으면 모두 비어 있어야 함
    - 표의 사람 답 칸 (「사람 답」 · 승인 표 「답」 · 로그 행의 번호 밖 칸)이 모두 비어 있어야 함"""
    for name, body in re.findall(r'<!-- auto:start (.+?) -->(.*?)<!-- auto:end \1 -->', text, re.S):
        if body.strip():
            return False
    for head, cols, rows in md_tables(text):
        for r in rows:
            for col in ('사람 답', '답'):
                if r.get(col, '').strip():
                    return False
            if '번호' in cols and '날짜' in cols and any(v.strip() for k, v in r.items() if k != '번호'):   # 로그 행
                return False
    return True


def round_shown(v, nd=1):
    """표시값(소수 nd 자리)을 반올림한 정수 (0.5 올림 · 재배치 인원 정수 규칙 · wfp-step4-gap SKILL)."""
    import decimal
    shown = decimal.Decimal(f'{v:.{nd}f}')
    return int(shown.quantize(decimal.Decimal('1'), rounding=decimal.ROUND_HALF_UP))


def row_answers(text, section='2.', col='승인'):
    """④ 2절 「승인」 칸 읽기. 빈칸 = CP4 답을 따름 (범위 전체).
    반환 (다른 답 행 [(RPU, 답)], 답 4종 밖 값 [(RPU, 값)])"""
    other, bad = [], []
    for head, cols, rows in md_tables(text):
        if not head.startswith(section) or col not in cols:
            continue
        for r in rows:
            v = r.get(col, '').strip()
            if not v or v.startswith('확정') or v.startswith('승인'):
                continue
            if any(v.startswith(k) for k in ('추가 확인', '보류', '이관')):
                other.append((r.get('RPU 코드', ''), v))
            else:
                bad.append((r.get('RPU 코드', ''), v))
    return other, bad


# ---------------------------------------------------------------------------
# 기본 대비 변경점 (customize.py lock · diff 와 뷰어가 같은 셈을 쓴다)
# ---------------------------------------------------------------------------
def _harness_files(base):
    import glob
    out = set()
    if os.path.exists(os.path.join(base, 'CLAUDE.md')):
        out.add('CLAUDE.md')
    for p in glob.glob(os.path.join(base, '.claude', '**', '*'), recursive=True):
        if os.path.isdir(p):
            continue
        rel = os.path.relpath(p, base).replace('\\', '/')
        if ('/_standard/' in rel or '/baselines/' in rel or '__pycache__' in rel or rel.endswith('.pyc')
                or rel == '.claude/settings.local.json'):
            continue
        if rel.startswith('.claude/hr-lib/') and rel != '.claude/hr-lib/config.json':
            continue   # 계산 엔진 · 뷰어 · 양식 · 도구 사본은 셈에서 뺀다 (엔진 변경은 lock 이 따로 봄)
        out.add(rel)
    return out


def harness_changes(root, context=1):
    """기본 사본(.claude/hr-lib/_standard)과 다른 파일 [dict(file, kind 추가 · 수정 · 삭제, plus, minus, diff)].
    대상 = CLAUDE.md + .claude/ 아래 파일 (hr-lib 은 config.json 만 · settings.local.json 제외)."""
    import difflib
    std = os.path.join(root, '.claude', 'hr-lib', '_standard')
    if not os.path.isdir(std):
        return None
    def lines(p):
        if not os.path.exists(p):
            return []
        with open(p, encoding='utf-8', errors='replace') as f:
            return f.read().splitlines()
    out = []
    for f in sorted(_harness_files(root) | _harness_files(std)):
        ta, tb = lines(os.path.join(std, f)), lines(os.path.join(root, f))
        if ta == tb:
            continue
        if f.endswith('.json') and ta and tb:   # 줄바꿈 · 들여쓰기만 다른 JSON 은 같은 것으로 본다
            try:
                if json.loads('\n'.join(ta)) == json.loads('\n'.join(tb)):
                    continue
            except ValueError:
                pass
        d = list(difflib.unified_diff(ta, tb, 'standard/' + f, 'mine/' + f, lineterm='', n=context))
        out.append(dict(file=f, kind='추가' if not os.path.exists(os.path.join(std, f)) else '삭제' if not os.path.exists(os.path.join(root, f)) else '수정',
                        plus=sum(1 for x in d if x.startswith('+') and not x.startswith('+++')),
                        minus=sum(1 for x in d if x.startswith('-') and not x.startswith('---')), diff='\n'.join(d[:400])))
    return out


# ---------------------------------------------------------------------------
# 스크립트가 맡는 칸 (auto 블록) · 사람이 적은 칸은 건드리지 않고 이 블록만 다시 쓴다
# ---------------------------------------------------------------------------
def replace_auto(text, name, content):
    a, b = f'<!-- auto:start {name} -->', f'<!-- auto:end {name} -->'
    block = f'{a}\n{content.rstrip()}\n{b}'
    if a in text and b in text:
        return text[:text.index(a)] + block + text[text.index(b) + len(b):]
    return text.rstrip() + '\n\n' + block + '\n'


def auto_block(text, name):
    """auto 블록 안 글 (없으면 None) · 다시 만들 때 숫자가 바뀌었는지 견줄 때 쓴다."""
    a, b = f'<!-- auto:start {name} -->', f'<!-- auto:end {name} -->'
    if a in text and b in text and text.index(a) < text.index(b):
        return text[text.index(a) + len(a):text.index(b)].strip()
    return None


def clear_approval(text, cp):
    """「## 승인」 표 cp 행의 답 · 승인일 · 이유 칸을 비운다 (승인자 직책은 그대로).
    숫자가 바뀐 자료를 다시 만들 때만 쓴다: 지난 답은 결정로그에 남아 있고, 새 숫자는 다시 승인받는다. 반환 (새 글, 지운 답)"""
    lines, old, in_appr = text.split('\n'), '', False
    for i, ln in enumerate(lines):
        s = ln.strip()
        if s.startswith('#'):
            in_appr = '승인' in s
            continue
        if in_appr and s.startswith('|') and s.strip('|').split('|')[0].strip().startswith(cp):
            c = [x.strip() for x in s.strip('|').split('|')]
            if len(c) >= 5 and c[1]:
                old = c[1]
                lines[i] = md_row([c[0], '', c[2], '', ''] + c[5:])
            break
    return '\n'.join(lines), old


def form(name):
    with open(os.path.join(HERE, 'forms', name), encoding='utf-8') as f:
        return f.read()


def read(path):
    with open(path, encoding='utf-8') as f:
        return f.read()


# ---------------------------------------------------------------------------
# 가상 데이터 표시 (pwc-hr-harness 가상 데이터 만들기 · virtual_data.py)
# 표시가 세 곳에 남는다: data/<회사>/_가상데이터.json · xlsx 마다 문서 속성 키워드와 2행 안내문 · 리포트 첫 줄.
# 표시 파일만 지워도 xlsx 표시가 남아 검증(verify)이 관문으로 멈춘다.
# ---------------------------------------------------------------------------
VIRTUAL_MARK = '_가상데이터.json'
VIRTUAL_TAG = 'pwc-hr-harness:virtual'
VIRTUAL_STAMP = '가상 데이터 · 실제 판단에 쓰지 않음'
VIRTUAL_LINE = '가상 데이터로 만든 결과 · 실제 판단에 쓰지 않음'


def virtual_info(data_dir):
    """data 폴더가 가상 데이터인가. 반환 dict(virtual, marker, stamped[파일], missing_marker, info)."""
    import glob
    info = None
    mark = os.path.join(data_dir, VIRTUAL_MARK)
    if os.path.exists(mark):
        try:
            with open(mark, encoding='utf-8') as f:
                info = json.load(f)
        except ValueError:
            info = {'표시': '(표시 파일을 읽을 수 없음)'}
    stamped = []
    for p in sorted(glob.glob(os.path.join(data_dir, '*.xlsx'))):
        if os.path.basename(p).startswith('~$'):
            continue
        try:
            wb = openpyxl.load_workbook(p, read_only=True)
            kw = wb.properties.keywords or ''
            a2 = next(wb.worksheets[0].iter_rows(min_row=2, max_row=2, max_col=1, values_only=True), (None,))[0]
            wb.close()
        except Exception:
            continue
        if VIRTUAL_TAG in kw or str(a2 or '').startswith(VIRTUAL_STAMP):
            stamped.append(os.path.basename(p))
    return dict(virtual=bool(info) or bool(stamped), marker=info is not None, stamped=stamped,
                missing_marker=bool(stamped) and info is None, info=info or {})


# ---------------------------------------------------------------------------
# 추가 자료 (data/<회사>/추가_* · 기본 표에 맞지 않는 파일) · 계산에 쓰지 않고 근거로만 인용
# ---------------------------------------------------------------------------
EXTRA_PREFIX = '추가_'
EXTRA_TITLE = '추가 자료 (계산에 안 씀 · 근거로 인용 가능)'


def guess_columns(ws_, max_scan=12, limit=12):
    """머리 행 짐작: 위 12행 가운데 글자 칸이 가장 많은 행 (같으면 위쪽). 반환 (행 번호, [열 이름])."""
    best, cols = None, []
    for i, r in enumerate(ws_.iter_rows(min_row=1, max_row=max_scan, values_only=True), start=1):
        names = [norm(v) for v in r if isinstance(v, str) and norm(v)]
        if len(names) >= 2 and len(names) > len(cols):
            best, cols = i, names
    return best, cols[:limit]


def ensure_logs(ws):
    """outputs/<회사>/ 에 결정로그 · 검증로그가 없으면 양식 (hr-lib/forms)에서 만든다 (새 회사 폴더 · /hr-run 첫 단계). 반환 만든 파일 이름."""
    made = []
    for f in ('결정로그.md', '검증로그.md'):
        dst, src = ws.o(f), os.path.join(HERE, 'forms', f)
        if not os.path.exists(dst) and os.path.exists(src):
            with open(src, encoding='utf-8') as a, open(dst, 'w', encoding='utf-8') as b:
                b.write(a.read())
            made.append(f)
    return made


def sheet_unit(ws_):
    """시트 맨 위 안내문에서 「단위: …」 문구를 찾는다 (없으면 None → CP1 에서 사람이 확인)."""
    for r in ws_.iter_rows(min_row=1, max_row=4, values_only=True):
        for v in r:
            if isinstance(v, str) and '단위' in v:
                m = re.search(r'단위\s*[:：]?\s*([^·]+)', v)
                return (m.group(1) if m else v).strip()[:60]
    return None
