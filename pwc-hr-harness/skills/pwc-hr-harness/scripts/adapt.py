# -*- coding: utf-8 -*-
"""pwc-hr-harness · 회사 적용 도우미 도구 (모드 F · /hr-customize 인자 없이 · Day 3).

사용 (작업 폴더에서 · /hr-init 이 .claude/hr-lib/adapt.py 에 사본을 둔다):
  python -X utf8 .claude/hr-lib/adapt.py status --co E-co             # 어디까지 했나 · 다음 질문
  python -X utf8 .claude/hr-lib/adapt.py record --co E-co --set "1 회사=가상 EPC 건설사 · 3,400명 · AI DC 전환"
  python -X utf8 .claude/hr-lib/adapt.py record --co E-co --add "7 자료=추가_현장배치기준.xlsx | 근거 인용 | 법정 배치 기준을 ④ 대안 근거로 | CP4 · CP5"
  python -X utf8 .claude/hr-lib/adapt.py design --co E-co             # outputs/E-co/자사버전_설계서.md 를 다시 씀
  python -X utf8 .claude/hr-lib/adapt.py harness --co E-co            # 도우미 답이 하네스 어느 파일 어느 줄로 들어갈지 미리 보기 (고치지 않음)
  python -X utf8 .claude/hr-lib/adapt.py harness --co E-co --apply    # 그대로 반영 (CLAUDE.md · 담당 · 스킬 「회사 규칙」 문단 · 변경 이력 · 잠금 검사)

하는 일
  - 도우미가 사람에게 물은 질문의 답을 data/<회사>/_도우미기록.json 에 남긴다 (중간에 멈춰도 이어 감)
  - 이미 있는 파일 (가상데이터_설정.json · _가상데이터.json · ①_열대응표.md · config.json · 결정로그 · 변경 이력 ·
    customize.py lock · _회귀테스트.md)을 읽어 어디까지 했는지 보여 주고, 설계서 한 장으로 모은다
  - 숫자를 만들거나 데이터 파일을 고치지 않는다. 하네스 파일은 harness --apply 일 때만, 「회사 규칙 · <회사>」 표시 문단 안만 고친다
    (기본 문장 · 잠금 칸은 그대로 · 같은 회사로 다시 돌리면 그 문단만 바뀜)
  - 설계서는 언제 다시 만들어도 같은 입력 = 같은 내용 (만든 시각 한 줄만 다름)
"""
import os, sys, json, re, glob, subprocess, datetime, argparse

HERE = os.path.dirname(os.path.abspath(__file__))
REC = '_도우미기록.json'
DESIGN = '자사버전_설계서.md'
SETTINGS = '가상데이터_설정.json'
MARK = '_가상데이터.json'
PRD = 'PRD.md'   # 회사 요구사항 (E-Co. 는 배포 자료 · 자사는 「우리 회사 PRD 만들어 줘」 로 만듦) · 질문 1~9 초안
PRD_TAG = '(자사 · /hr-customize 로 만듦)'
STEPS = [  # (번호, 이름, 기록 키) · 질문 대본 = references/adapt-interview.md
    ('0', '준비 (K-Co. 기준 · 회사 폴더)', None),
    ('1', '회사', '1 회사'),
    ('2', '데이터 길 (가상 · 자사)', '2 데이터 길'),
    ('3', '조직 (RPU)', None),
    ('4', '사업계획 (성장 · 신설 조직 · AI)', None),
    ('5', '동인', None),
    ('6', '만들기 · 열 대응 (CP1 전)', None),
    ('7', '우리 회사만의 자료', '7 자료'),
    ('8', 'CP 승인자', '8 승인자'),
    ('9', '하네스에 반영 (CLAUDE.md · 담당 · 스킬)', '9 하네스'),
    ('10', '결과 읽기 · 정리 (설계서)', '10 결과 읽기'),
]
# 자사 PRD 여섯 칸 (Day 3 v8 · E-Co. PRD 와 같은 칸 · 「/hr-customize 우리 회사 PRD 만들어 줘」 · 옛 이름 브리프 · 기록 키 「브리프 n」 = 「우리 회사 답 | 회사에서 구할 자료」 · 옛 키 「진단 n」 도 읽음)
# (번호, 브리프 칸, K-Co., E-Co., 하네스에서 바뀌는 곳 (도우미 질문))
DIAG = [
    ('1', '회사와 조직 (인력이 귀속되는 단위)', '공장 · 라인에 고정 · RPU 20개', '현장을 따라 이동 · 12개 조직 · 준공 시 해제', '1 회사 · 3 조직'),
    ('2', '사업 변화 (신설 · 확대 · 축소)', '3공장 신설 (2028-07 · 130명)', 'Fab 신규 현장 (2027-07 · 180명) · 플랜트 · 인프라 · 건축 축소', '4 사업계획'),
    ('3', '조직별 업무량 지표', '공장 생산량', '현장 = 공사 기성액 · PM = 수주 건수 · 안전 = 점검 건수', '5 지표 · CP2'),
    ('4', '계획 수치의 출처', '사업계획서', '사업부 추정치 대신 현장별 포트폴리오 합계', '4 사업계획 · 7 숫자 바꾸기'),
    ('5', '추가 자료', '없음', '법정 배치 기준 · 협력사 · 자격 · 조직 이동 · 준공 해제', '7 근거로 쓰기 · 숫자 바꾸기 · CP4'),
    ('6', '결재 구조', '인사기획팀장 · CHRO', 'Gap 대안에 해당 본부장 동석', '8 승인자 · 9 하네스에 반영'),
]


def root():
    d = os.getcwd()
    while d != os.path.dirname(d):
        if os.path.exists(os.path.join(d, 'CLAUDE.md')) and os.path.isdir(os.path.join(d, '.claude')):
            return d
        d = os.path.dirname(d)
    return os.getcwd()


def rd(p, default=''):
    try:
        with open(p, encoding='utf-8-sig') as f:
            return f.read()
    except OSError:
        return default


def rj(p):
    t = rd(p)
    if not t:
        return None
    try:
        return json.loads(t)
    except ValueError:
        return None


def blank(v):
    return v in (None, '', [], {}) or (isinstance(v, dict) and all(blank(x) for x in v.values())) or \
        (isinstance(v, list) and all(blank(x) for x in v))


def load_rec(R, co):
    j = rj(os.path.join(R, 'data', co, REC)) or {}
    j.setdefault('answers', {})
    return j


def save_rec(R, co, j):
    d = os.path.join(R, 'data', co)
    os.makedirs(d, exist_ok=True)
    j['updated'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    with open(os.path.join(d, REC), 'w', encoding='utf-8') as f:
        json.dump(j, f, ensure_ascii=False, indent=1)


def facts(R, co):
    """파일에서 읽은 사실 (도우미 진행 판단 · 설계서 공통)."""
    dd, od = os.path.join(R, 'data', co), os.path.join(R, 'outputs', co)
    st = rj(os.path.join(dd, SETTINGS)) or {}
    mk = rj(os.path.join(dd, MARK))
    names = sorted(os.path.basename(p) for p in glob.glob(os.path.join(dd, '*')) if os.path.isfile(p))
    skip = lambda x: x.startswith(('~$', '_', '.')) or x in (SETTINGS, REC, '열이름_대응.json', PRD) or x.lower().startswith('readme') or x.startswith('데이터사전')
    std_x = [x for x in names if not skip(x) and not x.startswith('추가_') and x.lower().endswith(('.xlsx', '.md', '.docx'))]
    extra = [x for x in names if x.startswith('추가_')]
    cfg = rj(os.path.join(R, '.claude', 'hr-lib', 'config.json')) or {}
    base = os.path.exists(os.path.join(R, '.claude', 'hr-lib', 'baselines', 'K-co_기준.json'))
    scan = os.path.exists(os.path.join(od, '①_열대응표.md'))
    log = rd(os.path.join(od, '결정로그.md'))
    return dict(dir=dd, out=od, settings=st, marker=mk, std_files=std_x, extra=extra, config=cfg, baseline=base,
                scan=scan, log=log, design=os.path.exists(os.path.join(od, DESIGN)))


def log_rows(text, kind=None):
    rows = []
    for ln in text.splitlines():
        if not ln.startswith('|') or ln.startswith('|---') or '번호' in ln[:8]:
            continue
        c = [x.strip() for x in ln.strip().strip('|').split('|')]
        if len(c) < 6 or not c[0].strip():
            continue
        if kind and not c[2].startswith(kind):
            continue
        rows.append(c)
    return rows


def progress(R, co):
    f, rec = facts(R, co), load_rec(R, co)
    a, st = rec['answers'], f['settings']
    virtual = (a.get('2 데이터 길', '') or '').startswith('가상') or bool(f['marker']) or bool(st)
    done = {
        '0': os.path.isdir(f['dir']),
        '1': not blank(a.get('1 회사')) or not blank(st.get('회사 이름')),
        '2': not blank(a.get('2 데이터 길')) or bool(f['marker']) or bool(f['std_files']) or bool(st),
        '3': (not virtual) or not blank(st.get('RPU')),
        '4': (not virtual) or any(not blank(st.get(k)) for k in ('성장 시나리오', '사업 이벤트', 'AI 범위', '인건비 인상률 (%)')) or a.get('4 사업계획') == 'K-Co. 와 같게',
        '5': (not virtual) or not blank(st.get('동인')) or not blank(st.get('추가 동인')) or a.get('5 동인') == 'K-Co. 원천 그대로',
        '6': bool(f['std_files']) and f['scan'],
        '7': '7 자료' in a,
        '8': '8 승인자' in a or not blank(f['config'].get('approvers')),
        '9': '9 하네스' in a,
        '10': False,
    }
    done['10'] = bool(f['design']) and '10 결과 읽기' in a and all(done[s] for s, _, _ in STEPS if s != '10')   # 첫 판 (6 뒤 「먼저 돌려 보기」)은 끝남이 아님 · 9 하네스 반영까지 해야 · 결과를 읽어야 끝남
    nxt = next((s for s, _, _ in STEPS if not done[s]), None)
    return f, rec, done, nxt, virtual


def status(R, args):
    f, rec, done, nxt, virtual = progress(R, args.co)
    out = dict(status='완료', co=args.co, 데이터=('가상' if virtual else '자사 · 아직 모름'),
               단계=[f'{"✓" if done[s] else "·"} {s} {n}' for s, n, _ in STEPS],
               다음=(f'{nxt} {dict((s, n) for s, n, _ in STEPS)[nxt]}' if nxt else '모두 끝남 · 「① 부터 돌릴까요?」'),
               기록=rec['answers'], 기록_시각=rec.get('updated', ''))
    if not f['baseline']:
        out['먼저'] = 'K-Co. 기준 스냅샷이 없음 · 「K-Co 결과를 기준으로 남길까요?」 (customize.py snapshot --co K-co --name 기준 · K-Co. 리포트까지 있어야 함)'
    if f['std_files'] and not f['marker'] and virtual:
        out['주의'] = '기본 파일은 있는데 가상 표시 파일이 없음 · 자사 파일이면 2 데이터 길을 「자사」로'
    if f['extra']:
        out['추가_자료'] = f['extra']
    dg = {d[0]: rec['answers'].get(f'브리프 {d[0]}') or rec['answers'].get(f'진단 {d[0]}') for d in DIAG}
    out['브리프'] = f'{sum(1 for v in dg.values() if not blank(v))}/{len(DIAG)} 칸 답함' + (' · 1~9 질문에서 이 답을 초안으로 보여 주고 확인만 받음' if any(dg.values()) else ' · 「/hr-customize 우리 회사 PRD 만들어 줘」')
    if os.path.exists(os.path.join(f['dir'], PRD)):
        sd = [x for x in f['extra'] if x == '추가_자사데이터.xlsx']
        out['PRD'] = f'data/{args.co}/{PRD} 있음' + (' · 자사 데이터 ' + sd[0] if sd else '') + ' → 질문 1~9 에서 PRD · 자사 데이터를 초안으로 보여 주고 확인 · 「판단 필요」 칸은 묻기만 함'
    if os.path.exists(os.path.join(f['out'], '_전후비교.md')):
        out['전후비교'] = f'outputs/{args.co}/_전후비교.md'
    print(json.dumps(out, ensure_ascii=False, indent=1))


def record(R, args):
    rec = load_rec(R, args.co)
    a = rec['answers']
    for kv in args.set or []:
        k, _, v = kv.partition('=')
        a[k.strip()] = v.strip()
    for kv in args.add or []:
        k, _, v = kv.partition('=')
        k = k.strip()
        cur = a.get(k)
        if not isinstance(cur, list):
            cur = [] if blank(cur) or cur == '없음' else [cur]
        if v.strip() not in cur:
            cur.append(v.strip())
        a[k] = cur
    save_rec(R, args.co, rec)
    print(json.dumps(dict(status='완료', file=f'data/{args.co}/{REC}', answers=a), ensure_ascii=False, indent=1))


# ---------------------------------------------------------------- 설계서
def tbl(head, rows):
    if not rows:
        return ['(없음)']
    return ['| ' + ' | '.join(head) + ' |', '|' + '---|' * len(head)] + ['| ' + ' | '.join(str(x if x not in (None, '') else '·') for x in r) + ' |' for r in rows]


def run_json(R, script, *a):
    p = os.path.join(R, '.claude', 'hr-lib', script)
    if not os.path.exists(p):
        p = os.path.join(HERE, script)
    try:
        r = subprocess.run([sys.executable, '-X', 'utf8', p, *a], cwd=R, capture_output=True, text=True, encoding='utf-8', timeout=120)
        return json.loads(r.stdout)
    except Exception:
        return None


def change_rows(R):
    hist = rd(os.path.join(R, 'CLAUDE.md'))
    i = hist.find('## 변경 이력')
    rows = []
    for ln in hist[i:].splitlines() if i >= 0 else []:
        if ln.startswith('| 20'):
            c = [x.strip() for x in ln.strip().strip('|').split('|')]
            if not any(w in (c[1] if len(c) > 1 else '') for w in ('기본 팀 설치', '표준 팀 설치')):   # 0.12.0 까지 설치 행 = 「표준 팀 설치」
                rows.append(c[:4])
    return rows


def waiting(R, co):
    """outputs/<회사>/_멈춤.md 의 멈춘 자리 (예: CP2) · 없으면 빈 문자열."""
    p = os.path.join(R, 'outputs', co, '_멈춤.md')
    if not os.path.exists(p):
        return ''
    m = re.search(r'^# 멈춤 · (?:CP )?(\S+)', open(p, encoding='utf-8').read(), re.M)
    return m.group(1) if m else ''


def run_from(R, co):
    w = waiting(R, co)
    return f'`/hr-run {co}` (' + (f'{w} 답부터' if w else 'CP1 부터') + ' · 결정로그에 사람 답)'


def nexts(R, co, nxt):
    od = os.path.join(R, 'outputs', co)
    if nxt in ('1', '2', '3', '4', '5', '6'):
        return f'도우미 {nxt} 단계부터 (`/hr-customize`)'
    gap = os.path.exists(os.path.join(od, '④_Gap대안.md'))
    if nxt == '10' and gap:
        return '도우미 10 결과 읽기 (`/hr-customize` 「결과 읽어 줘」 · K-Co. 와 비교 · 한 곳 고쳐 다시 계산)'
    if not glob.glob(os.path.join(od, '리포트_*.md')):
        return '①~④ 진행 · ' + run_from(R, co) + (' · 그 뒤 도우미 ' + nxt + ' 단계' if nxt in ('7', '8', '9', '10') else '')
    return f'도우미 {nxt} 단계부터 (`/hr-customize`)' if nxt in ('7', '8', '9') else '「오늘 작업 커밋하고 푸시해 줘」'


def design(R, args):
    co = args.co
    f, rec, done, nxt, virtual = progress(R, co)
    a, st, mk = rec['answers'], f['settings'], f['marker'] or {}
    name = st.get('회사 이름') or co
    L = [f'# 자사 버전 설계서 · {name} (`data/{co}/`)', '',
         f'> /hr-customize 가 묻고 답한 것과 바꾼 것을 모은 기록 · 만든 때 {datetime.datetime.now():%Y-%m-%d %H:%M} · `adapt.py design` 이 다시 만듦 (손으로 고치지 말고 도우미에게 말해 고침)']
    if virtual:
        L.append('> **가상 데이터로 만든 회사** · 숫자는 K-Co. 분포를 옮긴 연습용 · 실제 판단에 쓰지 않음')
    L += ['', '## 0. 한눈에', '']
    approvers = f['config'].get('approvers') or {}
    lk = run_json(R, 'customize.py', 'lock') or {}
    reg = rd(os.path.join(R, 'outputs', 'K-co', '_회귀테스트.md'))
    m = re.search(r'결과: \*\*(.+?)\*\*', reg)
    extra_n = len(a.get('7 자료') or []) if isinstance(a.get('7 자료'), list) else 0
    L += tbl(['항목', '내용'], [
        ['데이터', ('가상 (`_가상데이터.json` 표시)' if virtual else '자사 집계 파일') + f' · 기본 파일 {len(f["std_files"])}종 · 추가 자료 {len(f["extra"])}개'],
        ['전체 인원 · 조직', f'{(format(st["전체 인원"], ",") if isinstance(st.get("전체 인원"), (int, float)) else st.get("전체 인원", "·"))}명 · RPU {len(st.get("RPU") or [])}개' if virtual else '자사 파일 기준 (`①_열대응표.md`)'],
        ['우리 회사만의 자료', f'{extra_n}건 (6절)'],
        ['CP 승인자', '자사 직책' if approvers else '기본 그대로'],
        ['바꾼 하네스', f'{len(lk.get("changed") or [])}개 파일 · 변경 이력 {len(change_rows(R))}행'],
        ['잠금 검사', lk.get('status', '못 함')],
        ['K-Co. 회귀', m.group(1) if m else '아직 안 함'],
        ['다음 할 일', nexts(R, co, nxt)],
    ])
    # 0-1 자사화 진단 (K-Co. 와 무엇이 다른가 · 발표 첫 장)
    ex0 = a.get('7 자료')
    ex0 = ex0 if isinstance(ex0, list) else ([] if blank(ex0) else [ex0])
    ex0 = [x for x in ex0 if x != '없음']
    gal = lambda g: '; '.join(x.split('|')[0].strip() for x in ex0 if g in x) or ''
    dv0 = [k for k, v in (st.get('동인') or {}).items() if not blank(v)]
    ev0 = [e.get('이름') or e.get('코드') for e in st.get('사업 이벤트') or [] if not blank(e.get('코드'))]
    made = virtual or bool(f['std_files'])
    tagged = any(re.search(r'(?:진단|브리프)\s*[345]', x) for x in ex0)   # 7 자료 쓰임 칸에 「(진단 n)」 을 적었으면 그 줄로 나눔 · 없으면 갈래로
    how = lambda x: '숫자 바꾸기' if '입력 교체' in x else ('근거로 쓰기' if '근거 인용' in x else ('계산 바꾸기 (보류 · 이관)' if '계산 변경' in x else '자료'))
    def by(no, g):
        rows = [x for x in ex0 if re.search(rf'(?:진단|브리프)\s*{no}', x)] if tagged else [x for x in ex0 if g in x]
        return '; '.join(f'{how(x)}: {x.split("|")[0].strip()}' for x in rows)
    changed = {
        '1': ((f'조직 묶음 {len(st.get("RPU") or [])}개' + (f' · 새 조직 {", ".join(map(str, ev0))}' if ev0 else '')) if virtual else '자사 파일 조직') if made else '',
        '2': (('지표를 바꾼 묶음: ' + ', '.join(dv0)) if dv0 else 'K-Co. 지표 이름 그대로 (CP2 에서 확인)') if made else '',
        '3': by('3', '입력 교체'),
        '4': by('4', '근거 인용'),
        '5': by('5', '근거 인용') if tagged else '',
        '6': ('자사 직책' if approvers else '기본 그대로') + (f' · 규칙: {a["9 규칙"]}' if not blank(a.get('9 규칙')) and a.get('9 규칙') != '없음' else ''),
    }
    drows, need = [], []
    for no, q, kco, eco, where in DIAG:
        v = a.get(f'브리프 {no}') or a.get(f'진단 {no}') or ''
        mine, _, get = (str(v).partition('|'))
        drows.append([no, q, kco, mine.strip() or '(PRD 아직)', changed[no] or f'(아직 · 도우미 {where})'])
        if get.strip() and get.strip() not in ('없음', '-'):
            need.append(f'- {get.strip()} (PRD {no})')
    L += ['', '## 0-1. 자사 PRD (여섯 칸) · K-Co. 와 무엇이 다른가', '',
          '> 자사 PRD = 「/hr-customize 우리 회사 PRD 만들어 줘」 로 답한 것 (E-Co. PRD 와 같은 여섯 칸 · data/<회사>/PRD.md) · 바꾼 곳 = 질문 3~9 와 결정로그에서 읽음', '']
    L += tbl(['#', 'PRD 칸', 'K-Co.', '이 회사', '하네스에서 바꾼 곳'], drows)
    L += ['', '**회사에서 구할 자료 (실데이터로 다시 할 때)**', ''] + (need or ['- (PRD 에 적은 자료 없음)'])
    L += ['', '## 1. 회사', '', f'- 한 줄: {a.get("1 회사") or "(도우미 1 · 아직)"}']
    for x in st.get('사업계획 요점') or []:
        L.append(f'- 사업계획 요점: {x}')
    L += ['', '## 2. 데이터', '', f'- 길: {a.get("2 데이터 길") or ("가상" if virtual else "(도우미 2 · 아직)")}',
          f'- 폴더: `data/{co}/` · 기본 파일: {", ".join(f["std_files"]) or "(아직 없음)"}',
          f'- 열 대응 (CP1): {"`outputs/" + co + "/①_열대응표.md` 있음" if f["scan"] else "아직 (도우미 6)"}']
    if st.get('열이름_대응'):
        L.append('- 이름이 다른 열: ' + ' · '.join(f'{k} ← {v}' for k, v in st['열이름_대응'].items()))
    if virtual:
        L += ['', '## 3. 조직 (RPU)', '']
        L += tbl(['코드', '이름', '유형', '본부', '인원', 'K-Co. 원천'],
                 [[r.get('코드'), r.get('이름'), r.get('유형'), r.get('본부'), r.get('인원'), r.get('원천')] for r in st.get('RPU') or []])
        L += ['', '## 4. 사업계획', '']
        g = st.get('성장 시나리오') or {}
        gr = g.get('연평균 성장률 (%)') or {}
        L.append('- 성장: ' + (' · '.join(f'{k} {v}%' for k, v in gr.items() if v not in ('', None)) or 'K-Co. 성장률') +
                 (f' · 기준 매출 {g.get("기준 매출 (억 원)")}억 원' if g.get('기준 매출 (억 원)') else ''))
        ev = [e for e in st.get('사업 이벤트') or [] if not blank(e.get('코드'))]
        L += ['- 새로 생기는 조직:', ''] + (tbl(['코드', '이름', '가동', '확보', '인원', '근거'],
                                        [[e.get('코드'), e.get('이름'), e.get('가동'), e.get('확보'), e.get('인원'), e.get('근거')] for e in ev]) if ev else ['  (없음)'])
        L.append(f'- AI 먼저 쓸 곳: {", ".join(st.get("AI 범위") or []) or "K-Co. 범위 (관리 지원 · 영업)"} · 인건비 인상률: {st.get("인건비 인상률 (%)") or "K-Co. 3.5"}%')
        L += ['', '## 5. 동인', '']
        dv = {k: v for k, v in (st.get('동인') or {}).items() if not blank(v)}
        L += tbl(['RPU', '동인', '단위', '기준 값', '계획'], [[k, v.get('이름'), v.get('단위'), v.get('기준 값'),
                                                         v.get('계획 값') or (f'{v.get("계획 증가율 (%)")}%' if v.get('계획 증가율 (%)') else '')] for k, v in dv.items()])
        chk = (mk.get('동인 확인') or {}).get('RPU') if isinstance(mk.get('동인 확인'), dict) else None
        if chk:
            L.append(f'- K-Co. 원천 동인 이름 그대로인 RPU: {", ".join(map(str, chk)) if isinstance(chk, list) else chk} (CP2 에서 동인 이름부터 확인)')
    L += ['', '## 6. 우리 회사만의 자료 (근거 인용 · 입력 교체 · 계산 변경)', '']
    ex = a.get('7 자료')
    ex = ex if isinstance(ex, list) else ([] if blank(ex) else [ex])
    L += tbl(['자료', '갈래', '쓰임', '다시 받을 CP'], [([c.strip() for c in x.split('|')] + ['', '', '', ''])[:4] for x in ex if x != '없음'])
    dr = log_rows(f['log'], '데이터 반영')
    if dr:
        L += ['', '- 결정로그 「데이터 반영」 행:', ''] + tbl(['번호', '날짜', '정한 것', '답', '후속'], [[r[0], r[1], r[3], r[5], r[9] if len(r) > 9 else ''] for r in dr])
    L += ['', '## 7. CP 승인자 (CP 5곳은 잠금 · 직책만 바꿈)', '']
    std_ap = dict(CP1='인사기획 실무 담당', CP2='인사기획팀장', CP3='인사기획팀장', CP4='인사기획팀장 + R 행 소관 본부장', CP5='CHRO')
    L += tbl(['CP', '승인자 (직책)', '기본'], [[k, approvers.get(k) or '기본 그대로', v] for k, v in std_ap.items()])
    if a.get('8 승인자'):
        L.append(f'- 도우미 답: {a["8 승인자"]} (승인자는 작업 폴더 전체 `config.json` · 다른 회사 폴더에도 같이 보임)')
    L += ['', '## 8. 바꾼 하네스 (내 하네스)', '', f'- 도우미 9 하네스에 반영: {a.get("9 하네스") or "(아직 · `/hr-customize` 9 단계)"}']
    hr = rd(os.path.join(f['out'], '_하네스반영.md'))
    if '반영함' in hr:
        L += [''] + [x for x in hr.splitlines() if x.startswith('|')] + ['']
    L.append('- 기본 대비 바뀐 파일 (`customize.py diff`):')
    L += [f'  - {c}' for c in lk.get('changed') or []] or ['  - (없음 · 기본 팀 그대로)']
    L += ['- CLAUDE.md 변경 이력:', ''] + tbl(['날짜', '변경 내용', '대상', '사유'], change_rows(R))
    L += ['', '## 9. 잠금 · 회귀', '', f'- 잠금 검사 (`customize.py lock`): **{lk.get("status", "못 함")}**']
    L += [f'  - 위반: {v}' for v in lk.get('violations') or []] + [f'  - 확인: {w}' for w in lk.get('warnings') or []]
    L.append(f'- K-Co. 회귀 (`outputs/K-co/_회귀테스트.md`): **{m.group(1) if m else "아직 안 함 · 기준 스냅샷이 있어야 함"}**')
    L += ['', '## 10. 결과 읽기 · 고치기 전후', '',
          f'- 결과 읽기 (K-Co. 와 비교 · 우리 회사 감각과 맞나 · PRD 몇 번 칸 때문인가): {a.get("10 결과 읽기") or "(아직 · 계산 뒤 도우미 10)"}',
          f'- 고친 것 (한 곳): {a.get("10 고친 것") or "(아직)"}']
    cmpf = rd(os.path.join(f['out'], '_전후비교.md'))
    if cmpf:
        mm = re.search(r'결과: \*\*(.+?)\*\*', cmpf)
        L.append(f'- 전후 비교 (`outputs/{co}/_전후비교.md`): **{mm.group(1) if mm else "있음"}**')
        sec = cmpf.split('## 합계', 1)[1].split('## 달라진', 1)[0] if '## 합계' in cmpf else ''
        L += [''] + [x for x in sec.strip().splitlines() if x.startswith('|')]
        top = [x for x in cmpf.split('## 달라진', 1)[1].splitlines() if x.startswith('| ') and not x.startswith('| 항목')][:5] if '## 달라진' in cmpf else []
        if top:
            L += ['', '- 많이 달라진 조직 묶음 (위 5칸):', '', '| 항목 | RPU | 고치기 전 | 지금 | 차이 |', '|---|---|---|---|---|'] + top
    else:
        L.append('- 전후 비교: (아직 · 고치기 전에 결과를 저장하고 다시 계산한 뒤)')
    todo = [f'{s} {n}' for s, n, _ in STEPS if not done[s] and s != '10']
    L += ['', '## 11. 남은 것', ''] + ([f'- [ ] 도우미 {t}' for t in todo] or ['- 도우미 단계 모두 끝남 · 다음: ' + run_from(R, co)])
    L.append('')
    os.makedirs(f['out'], exist_ok=True)
    p = os.path.join(f['out'], DESIGN)
    with open(p, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(L))
    print(json.dumps(dict(status='완료', file=os.path.relpath(p, R).replace('\\', '/'), 남은=todo, 잠금=lk.get('status', '못 함')), ensure_ascii=False, indent=1))


# ---------------------------------------------------------------- 하네스에 반영 (도우미 9 · Day 3 v8)
# 회사 답 → 하네스 세 층 (CLAUDE.md = 회사 · 목표 / 담당 = 누가 무엇을 입력으로 / 스킬 = 어떻게 판단) 의 「회사 규칙」 문단.
# 기본 문장은 지우지 않고 회사 문단을 덧붙인다 → customize.py lock (잠금 칸) 통과 · 계산 스크립트 · 엔진 불변 → K-Co. 숫자 그대로.
LAYER = {'CLAUDE.md': 'CLAUDE.md', 'agents': '담당 (Agent)', 'skills': '스킬 (Skill)'}


def _mark(co):
    return f'<!-- 회사 규칙 · {co} · 회사 적용 도우미 (adapt.py harness) -->', f'<!-- /회사 규칙 · {co} -->'


def _v(x):
    if isinstance(x, list):
        return ' / '.join(str(i) for i in x if not blank(i))
    return '' if blank(x) else str(x)


def harness_plan(R, co):
    f, rec = facts(R, co), load_rec(R, co)
    a, st = rec['answers'], f['settings']
    name = st.get('회사 이름') or co
    ap = f['config'].get('approvers') or {}
    drv = st.get('동인') or {}
    no_k = [c for c, d in (drv.items() if isinstance(drv, dict) else []) if isinstance(d, dict) and str(d.get('다른 후보') or '').strip() in ('빼기', '뺌', '없음', '빼')]
    src = _v(a.get('4 계획 출처'))
    ex7 = [x for x in (a.get('7 자료') or []) if x != '없음'] if isinstance(a.get('7 자료'), list) else ([a['7 자료']] if not blank(a.get('7 자료')) and a.get('7 자료') != '없음' else [])
    B = []   # (파일, 브리프 칸, 왜, [줄])
    L = []
    if not blank(a.get('1 회사')):
        L.append(f'- 회사: {_v(a["1 회사"])}')
    if not blank(a.get('3 조직')):
        L.append(f'- 조직 (RPU): {_v(a["3 조직"])}')
    if not blank(a.get('4 사업계획')):
        L.append(f'- 사업 변화: {_v(a["4 사업계획"])}')
    if L:
        L += [f'- `{co}` 요청은 이 칸과 각 담당 · 스킬의 「회사 규칙 · {co}」 문단을 먼저 읽는다. 기본 규칙 · 업무 범위 · 금지는 그대로 따른다',
              f'- 요청 예: 「{co} 인력계획 처음부터 돌려 줘」 「{co} 어느 조직이 얼마나 부족해?」 「{co} 남는 인원은 어디로 옮길 수 있어?」']
        B.append(('CLAUDE.md', '1 · 2', '회사 · 조직 · 사업 변화 (하네스의 목표 칸)', [f'## 우리 회사 · {name} (`data/{co}/`)', ''] + L))
    L = []
    if not blank(a.get('3 조직')):
        L.append(f'- RPU 단위: {_v(a["3 조직"])} (K-Co. 의 공장 · 라인 묶음을 쓰지 않음)')
    ev = st.get('사업 이벤트')
    if not blank(ev):
        for e in (ev if isinstance(ev, list) else [ev]):
            L.append('- 사업 이벤트 (별도 행 · 과거 추이로 산정하지 않음): ' + (' · '.join(f'{k} {v}' for k, v in e.items() if not blank(v)) if isinstance(e, dict) else str(e)))
    if L:
        B.append(('.claude/skills/wfp-step1-strategy/SKILL.md', '1 · 2', 'RPU 와 신설 조직 (사업 이벤트) 을 판단하는 규칙', [f'## 회사 규칙 · {co}', ''] + L))
    L = []
    if not blank(a.get('5 동인')):
        L.append(f'- 인력 동인: {_v(a["5 동인"])}')
    if no_k:
        L.append(f'- K-Co. 에서 온 지표 이름은 이 조직의 동인 후보로 쓰지 않는다: {", ".join(no_k)} (업무량과 방향이 다를 수 있음)')
    if src:
        L.append(f'- 계획 수치의 출처: {src} · 다른 출처 (예: 사업부 추정치) 는 CP2 에서 비교로만 보여 준다')
    if L:
        B.append(('.claude/skills/wfp-step2-demand/SKILL.md', '3 · 4', '인력 동인과 계획 수치의 출처를 고르는 규칙', [f'## 회사 규칙 · {co}', ''] + L))
        B.append(('.claude/agents/wfp-demand.md', '3 · 4', '수요 담당의 기준 (동인 · 출처)',
                  [f'## 회사 규칙 · {co}', '', f'- `{co}` 수요는 `wfp-step2-demand` 「회사 규칙 · {co}」 의 동인 · 계획 수치 출처를 따른다. CP2 보고에 출처를 한 줄로 밝힌다']))
    if src:
        B.append(('.claude/agents/wfp-verifier.md', '4', '검증 담당의 대조 항목 (계획 값 출처)',
                  [f'## 회사 규칙 · {co}', '', f'- `{co}` ② 검증 때 계획 연도 동인 값이 「{src}」 와 같은지 원천 파일에서 대조한다. 다르면 「확인 필요」']))
    if f['extra'] or ex7:
        used = [x for x in f['extra'] if x in src or x.replace('.xlsx', '') in src]   # 계획 수치의 출처로 이미 계산에 들어간 자료
        cite = [x for x in f['extra'] if x not in used]
        L = [f'- 추가 자료 (계산에 넣지 않고 Gap 대안 근거로 인용): {", ".join(cite)}'] if cite else []
        L += [f'- 이미 계산에 들어간 자료 (계획 수치의 출처 · 도우미 4): {", ".join(used)}'] if used else []
        L += [f'- {x}' for x in ex7]
        B.append(('.claude/agents/wfp-gap.md', '5', 'Gap 담당의 입력 (추가 자료)', [f'## 회사 규칙 · {co}', ''] + L))
        L = ['- 여유 인력은 추가 자료 (예: 준공 해제 · 조직 이동) 의 시점과 함께 재배치 경로를 먼저 제시한다',
             '- 법 · 자격 기준 자료가 있으면 그 하한을 대안의 제약으로 적는다 (법령 원문 판단은 사람)']
        if ex7:
            L.append(f'- 자료별 쓰임은 `wfp-gap` 담당의 「회사 규칙 · {co}」 문단 (도우미 7) 을 따른다')
        B.append(('.claude/skills/wfp-step4-gap/SKILL.md', '5', 'Gap 대안 (4B · 재배치) 을 고르는 규칙', [f'## 회사 규칙 · {co}', ''] + L))
    L = []
    if not blank(a.get('8 승인자')) and str(a.get('8 승인자')).strip() not in ('그대로', '없음', '기본 그대로', '표준 그대로'):
        L.append(f'- CP 승인자 (직책): {_v(a["8 승인자"])}')
    elif ap:
        L.append('- CP 승인자 (직책): ' + ' · '.join(f'{k} {v}' for k, v in ap.items()))
    rule = a.get('9 규칙')
    if not blank(rule) and rule != '없음':
        L.append(f'- 회사 규칙 (도우미 9): {_v(rule)}')
    if L:
        L.append('- CP 5곳의 자리와 「승인 전에는 넘어가지 않음」 은 그대로')
        B.append(('.claude/skills/wfp-lead/SKILL.md', '6', 'CP 승인자 (결재 구조)', [f'## 회사 규칙 · {co}', ''] + L))
    base = co[:-len('_실데이터')] if co.endswith('_실데이터') else co
    scope = f'- 적용 폴더: `{base}` · `{base}_실데이터` (같은 회사 · 가상 데이터와 회사에서 바꾼 실데이터 모두)'
    B = [(p, k, w, lines[:2] + [scope] + lines[2:]) for p, k, w, lines in B]
    return name, B


def _rd_raw(fp):
    """원래 줄바꿈 (LF · CRLF) 을 기억하며 읽음 · Windows 에서 줄바꿈이 바뀌면 Claude Code 가 담당 머리말을 못 읽는 일이 있음 (2026-11-16 실측 · wfp-verifier 가 담당 목록에서 빠짐)."""
    with open(fp, 'rb') as fh:
        raw = fh.read().decode('utf-8-sig')
    return raw.replace('\r\n', '\n'), ('\r\n' if '\r\n' in raw else '\n')


def _wr_raw(fp, text, eol):
    with open(fp, 'w', encoding='utf-8', newline='') as fh:
        fh.write(text.replace('\r\n', '\n').replace('\n', eol))


def _put(text, block, co, before=None):
    s, e = _mark(co)
    text = re.sub(re.escape(s) + r'.*?' + re.escape(e) + r'\n*', '', text, flags=re.S)
    blk = s + '\n' + '\n'.join(block) + '\n' + e + '\n'
    if before and before in text:
        i = text.index(before)
        return text[:i].rstrip('\n') + '\n\n' + blk + '\n' + text[i:]
    return text.rstrip('\n') + '\n\n' + blk


def harness(R, args):
    co = args.co
    name, B = harness_plan(R, co)
    if not B:
        print(json.dumps(dict(status='멈춤', reason='반영할 답이 없음 · 도우미 1~5 를 먼저'), ensure_ascii=False))
        sys.exit(3)
    miss = [p for p, *_ in B if not os.path.exists(os.path.join(R, p))]
    B = [b for b in B if b[0] not in miss]
    layer = lambda p: LAYER['CLAUDE.md'] if p == 'CLAUDE.md' else LAYER[p.split('/')[1]]
    nm = lambda p: 'CLAUDE.md' if p == 'CLAUDE.md' else p.split('/')[2].replace('.md', '')
    od = os.path.join(R, 'outputs', co)
    os.makedirs(od, exist_ok=True)
    M = [f'# 하네스에 반영 · {name}', '',
         ('> **반영함** · ' if args.apply else '> **미리 보기 · 아직 고치지 않음** · ') +
         f'회사 답이 하네스 세 층 (CLAUDE.md · 담당 · 스킬) 의 「회사 규칙 · {co}」 문단으로 들어감 · 기본 문장 · 잠금 칸 · 계산 스크립트는 그대로', '',
         '| 층 | 파일 | PRD 칸 | 왜 |', '|---|---|---|---|']
    M += [f'| {layer(p)} | {nm(p)} | {k} | {w} |' for p, k, w, _ in B]
    for p, k, w, lines in B:
        M += ['', f'## {layer(p)} · `{p}`', ''] + [x for x in lines[2:]]
    out = dict(status='완료', mode='반영' if args.apply else '미리 보기', 파일=[f'{layer(p)} · {nm(p)}' for p, *_ in B])
    if miss:
        out['없는_파일'] = miss
    if args.apply:
        for p, k, w, lines in B:
            fp = os.path.join(R, p)
            t, eol = _rd_raw(fp)
            t = _put(t, lines, co, before='## 변경 이력' if p == 'CLAUDE.md' else None)
            _wr_raw(fp, t, eol)
        cp = os.path.join(R, 'CLAUDE.md')
        t, eol = _rd_raw(cp)
        tag = f'회사 규칙 반영 · {co}'
        t = '\n'.join(x for x in t.splitlines() if tag not in x) + '\n'
        row = (f'| {datetime.date.today():%Y-%m-%d} | {tag} (/hr-customize · PRD 칸 → 회사 규칙 문단 {len(B)}곳) | '
               + ' · '.join(nm(p) for p, *_ in B) + f' | {name} 사정 반영 (기본 문장 · 잠금 칸 그대로) |')
        lines = t.rstrip('\n').splitlines()
        lines.append(row)
        _wr_raw(cp, '\n'.join(lines) + '\n', eol)
        lk = run_json(R, 'customize.py', 'lock') or {}
        out['잠금'] = lk.get('status', '못 함')
        rg = run_json(R, 'customize.py', 'regress', '--co', 'K-co', '--name', '기준') if os.path.exists(os.path.join(R, '.claude', 'hr-lib', 'baselines', 'K-co_기준.json')) else None
        out['K-Co_회귀'] = (rg or {}).get('status', '기준 없음 · 못 함')
        M += ['', f'- 잠금 검사 (`customize.py lock`): **{out["잠금"]}** · K-Co. 다시 계산 (`customize.py regress`): **{out["K-Co_회귀"]}**']
        rec = load_rec(R, co)
        rec['answers']['9 하네스'] = f'{len(B)}곳 · ' + ' · '.join(nm(p) for p, *_ in B)
        save_rec(R, co, rec)
    with open(os.path.join(od, '_하네스반영.md'), 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(M) + '\n')
    out['file'] = f'outputs/{co}/_하네스반영.md'
    print(json.dumps(out, ensure_ascii=False, indent=1))


def prd_cmd(R, args):
    """자사 PRD (오전 끝 「/hr-customize 우리 회사 PRD 만들어 줘」 · 여섯 칸 답 「브리프 n」 → data/<회사>/PRD.md · 오후 질문 1~9 초안)."""
    co = args.co
    rec = load_rec(R, co)
    a = rec['answers']
    d = os.path.join(R, 'data', co)
    p = os.path.join(d, PRD)
    if os.path.exists(p) and PRD_TAG not in rd(p):
        print(json.dumps(dict(status='멈춤', reason=f'data/{co}/{PRD} 가 이미 있음 (직접 만든 PRD · 덮지 않음)'), ensure_ascii=False))
        sys.exit(3)
    L = [f'# {co} 인력계획 에이전트 PRD {PRD_TAG}', '',
         '> 실명 · 실제 내부 수치 · 개인 정보 없음 · 오후 `/hr-customize` 질문 1~9 의 초안 · 「모름」 칸은 그 질문에서 다시 묻거나 「판단 필요」', '']
    for n, k, kco, eco, where in DIAG:
        ans = a.get(f'브리프 {n}') or a.get(f'진단 {n}') or ''
        own, _, need = str(ans).partition('|')
        L += [f'## {n}. {k}', '', f'- 우리 회사: {own.strip() or "(아직 · 판단 필요)"}', f'- 회사에서 구할 자료: {need.strip() or "(모름)"}',
              f'- 참고 · K-Co.: {kco} · E-Co.: {eco}', f'- 들어가는 곳: /hr-customize 질문 {where}', '']
    os.makedirs(d, exist_ok=True)
    with open(p, 'w', encoding='utf-8') as fh:
        fh.write(chr(10).join(L) + chr(10))
    done = sum(1 for n, *_ in DIAG if not blank(a.get(f'브리프 {n}') or a.get(f'진단 {n}')))
    print(json.dumps(dict(status='완료', file=f'data/{co}/{PRD}', 칸=f'{done}/{len(DIAG)}'), ensure_ascii=False, indent=1))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('cmd', choices=['status', 'record', 'design', 'harness', 'prd'])
    ap.add_argument('--apply', action='store_true', help='harness: 미리 보기 대신 실제로 반영')
    ap.add_argument('--co', required=True)
    ap.add_argument('--set', action='append', help='「키=값」 · 키 예: 「1 회사」 「2 데이터 길」 「8 승인자」 「9 규칙」 (값 「없음」 = 건너뜀)')
    ap.add_argument('--add', action='append', help='목록 키에 한 줄 더함 · 「7 자료=파일 | 갈래 | 쓰임 | 다시 받을 CP」')
    a = ap.parse_args()
    R = root()
    if not os.path.isdir(os.path.join(R, '.claude')):
        print(json.dumps(dict(status='멈춤', reason='작업 폴더가 아님 (/hr-init 으로 설치한 폴더에서 실행)'), ensure_ascii=False))
        sys.exit(3)
    if re.search(r'[\\/]|\.\.', a.co):
        print(json.dumps(dict(status='멈춤', reason='회사 폴더 이름에 경로 기호를 쓸 수 없음'), ensure_ascii=False))
        sys.exit(3)
    dict(status=status, record=record, design=design, harness=harness, prd=prd_cmd)[a.cmd](R, a)


if __name__ == '__main__':
    main()
