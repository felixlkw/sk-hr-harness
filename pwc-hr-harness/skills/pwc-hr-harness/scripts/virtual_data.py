# -*- coding: utf-8 -*-
"""pwc-hr-harness · 가상 데이터 만들기 (/hr-customize B2 · 자사 파일이 없을 때).

사용 (작업 폴더에서 · /hr-init 이 .claude/hr-lib/virtual_data.py 에 사본을 둔다):
  python -X utf8 .claude/hr-lib/virtual_data.py --co 자사              # data/자사/가상데이터_설정.json 으로 만듦
  python -X utf8 .claude/hr-lib/virtual_data.py --co 자사 --dry-run    # 만들지 않고 RPU 대응 · 인원만 보여 줌
  python -X utf8 .claude/hr-lib/virtual_data.py --co 자사 --replace-virtual   # 전에 만든 가상 데이터만 다시 만듦

하는 일
  참가자가 적은 작은 설정 (회사 이름 · 기준일 · 전체 인원 · RPU 목록 · 열 이름 대응 · 사업계획 요점)으로
  data/<회사>/ 에 기본 6종 구조의 가상 데이터를 만든다. 숫자는 지어내지 않고 K-Co. 분포를 옮긴다.
  - RPU 마다 같은 유형의 K-Co. RPU 하나를 원천으로 고르고 (이름이 가장 비슷한 것), 인원 배율 = 자사 인원 ÷ 원천 인원
  - 직급 · 연령 · 입퇴사 흐름 (퇴직률) · 업무량 (FTE ÷ 현원) · 스킬 보유 분포 (충족률) · 동인 시계열 증가율을 그 배율로 옮긴다
  - 매출 · 이익 · 인건비 = K-Co. 값 × (자사 전체 인원 ÷ K-Co. 전체 인원) · 성장률은 K-Co. 사업계획과 같음
    (성장 시나리오를 적으면 그 값 · 원천에 있는 「전사 매출액」 같은 매출 동인 행 (억 원)도 파일 2 매출 추이 · 파일 1 기본안 계획 매출을 따름)
  - 동인 이름: 「동인」 칸을 비운 RPU 는 K-Co. 원천 동인 이름 (생산량 · 설비 수 · 천 톤 등)을 그대로 씀 → --dry-run · 데이터 사전 · 표시 파일에
    「동인 확인」 (동인 이름이 K-Co. 원천 이름 그대로 · 회사에 맞게 바꿀지 확인) 으로 알리고, 리드가 만들 때 사람에게 묻는다 (업종 동인 이름을 지어내지 않음)
  - 결정적: 시드 고정 · 같은 설정 = 같은 내용 (표시 파일 「내용 지문」으로 확인)
  - 개인 식별 열 (성명 · 사번 · 생년월일)을 만들지 않는다. 모든 행은 RPU · 직급 집계 단위다 (사람 단위 행 없음)
  - 가상 표시: data/<회사>/_가상데이터.json · xlsx 문서 속성과 2행 안내문 · 사업계획 문서 첫 줄. 리포트 · 뷰어 · 검증이 이 표시를 읽는다
  - 기존 자사 파일이 있으면 덮지 않고 멈춘다 (종료 코드 3)
  - 함께 둘 수 있는 파일: README · 설정 파일 · .gitkeep · 빈 데이터 사전 · `추가_*` (참가자가 넣은 추가 자료 · 다시 만들 때도 지우지 않음)
선택 설정 (0.10.0 · 비우면 지금처럼 K-Co. 값 · 같은 설정 = 같은 내용)
  - 성장 시나리오: 기준 매출 · 보수 / 기본 / 낙관 연평균 성장률 (%) 또는 계획 매출 · 계획 영업이익률 → 파일 1 · 1절과 파일 2 연도별 추이
  - 사업 이벤트: 신설 조직 [코드 · 이름 · 유형 · 본부 · 가동 · 확보 · 인원 · 구성 · 근거] → 파일 1 · 2절 (K-Co. 3공장 문장 형식)
  - 동인: RPU 별 주 동인 이름 · 단위 · 기준 값 · 계획 증가율 (%) 또는 계획 값 → 파일 6 행 이름 · 값 (원천 모양을 새 계획에 맞춤)
  - 추가 동인: RPU 별 동인 후보 행 추가 (인원 추이 모양 · 시드 고정) → CP2 후보
  - AI 범위 (유형 또는 RPU 코드) · 인건비 인상률 (%)
LLM 으로 숫자를 만들지 않는다.
"""
import os, sys, json, re, glob, hashlib, random, datetime, argparse, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
PLUGIN_LIB = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), 'templates', 'standard-team', '.claude', 'hr-lib')
CONFIG_NAME = '가상데이터_설정.json'
FORM_NAME = '가상데이터_설정_예.json'
EXIT_STOP = 3
BAN_COL = re.compile(r'성명|사번|생년월일|주민|휴대|전화|이메일|e-?mail', re.I)

# K-Co. 사업계획 (파일 1) 1절 · 6절의 성장률 · 인상률 (교육 고정 데이터 · 2025 실적 대비)
KCO_PLAN = dict(rev={'보수': 10600 / 9500, '기본': 11300 / 9500, '낙관': 11900 / 9500}, op=850 / 700, raise_=0.035)

TYPES = {   # 설정의 「유형」 → K-Co. 파일 6 하위 유형
    '운영': ('Profit Center', '운영 (실적 물량)', '운영본부'),
    '영업': ('Profit Center', '영업 (수주·고객)', '영업본부'),
    '기술·전문': ('Value Center', '기술·전문 (지원 대상 규모)', '기술본부'),
    '관리 지원': ('Value Center', '관리 지원 (전사 인원 대비 비율)', '경영지원본부'),
}
TYPE_ALIAS = {'운영': '운영', '영업': '영업', '기술·전문': '기술·전문', '기술': '기술·전문', '전문': '기술·전문', '기술전문': '기술·전문',
              '관리 지원': '관리 지원', '관리지원': '관리 지원', '관리': '관리 지원', '지원': '관리 지원'}
AI_TYPES = ('관리 지원', '영업')   # K-Co. 사업계획 4절 (경영지원 · 영업 반복 업무 우선)을 유형으로 옮김
NOSCALE_UNITS = {'일', '조', '시간', '%'}
REVENUE = re.compile(r'매출')                     # 매출 동인 행 (성장 시나리오를 적으면 그 매출을 따름)
WHOLE_CO = re.compile(r'전사|회사 전체|매출|영업이익')   # step2 동인 원칙: Value Center 에 회사 전체 실적을 쓰지 않음
KEEP_FILES = {'README.md', '.gitkeep', '_도우미기록.json', 'PRD.md'}   # PRD.md = 회사 요구사항 (2026-11-16 · /hr-customize 가 질문 초안으로 읽음)
#   # 가상 폴더에 함께 둘 수 있는 파일 (설정 파일 · 빈 데이터 사전 · 추가_* 은 따로 봄 · 도우미 기록은 adapt.py 가 씀 · 지우지 않음)
EXTRA_PREFIX = '추가_'                   # 참가자가 넣는 추가 자료 (기본 6종 밖 · step1 scan 이 「추가 자료」로 보고)
SKILL_MONTHS = 12                        # 사업 이벤트 확보 시점을 비웠을 때: K-Co. 3공장 「숙련에 12개월」 (파일 1 · 2절)


def _lib(root):
    for d in (HERE, os.path.join(root, '.claude', 'hr-lib'), PLUGIN_LIB):
        p = os.path.join(d, 'hrlib.py')
        if os.path.exists(p) and 'VIRTUAL_MARK' in open(p, encoding='utf-8').read():
            sys.path.insert(0, d)
            import hrlib
            return hrlib
    raise SystemExit('hrlib.py (가상 표시 포함 판) 를 찾지 못함 · 플러그인을 다시 설치하세요')


def find_root(start):
    d = os.path.abspath(start)
    while True:
        if os.path.exists(os.path.join(d, 'CLAUDE.md')) and os.path.isdir(os.path.join(d, '.claude')):
            return d
        nd = os.path.dirname(d)
        if nd == d:
            return os.path.abspath(start)
        d = nd


class Halt(Exception):
    def __init__(self, what, evidence, options):
        super().__init__(what)
        self.what, self.evidence, self.options = what, evidence, options


# ---------------------------------------------------------------------------
# 정수 도구 (결정적)
# ---------------------------------------------------------------------------
def rh(x):
    """0.5 올림 반올림 (음수는 0 쪽 대칭)."""
    return int(x + 0.5) if x >= 0 else -int(-x + 0.5)


def apportion(total, weights):
    """total 을 weights 비율로 나눈 정수 (최대 잔여 · 같으면 앞쪽) · 합 = total."""
    total = int(total)
    n = len(weights)
    s = float(sum(w for w in weights if w and w > 0))
    if total <= 0 or n == 0:
        return [0] * n
    if s == 0:
        weights, s = [1] * n, float(n)
    raw = [total * max(0.0, w or 0) / s for w in weights]
    out = [int(r) for r in raw]
    rest = total - sum(out)
    for i in sorted(range(n), key=lambda i: (-(raw[i] - out[i]), i))[:rest]:
        out[i] += 1
    return out


def sv(v, k):
    """동인 값 배율 (10 이상 정수 · 미만 소수 1자리 · 정수로 떨어지면 정수)."""
    x = v * k
    if abs(x) >= 10:
        return rh(x)
    x = round(x, 1)
    return int(x) if float(x).is_integer() else x


# ---------------------------------------------------------------------------
# K-Co. 원천 읽기
# ---------------------------------------------------------------------------
def _rows(sh, hrow):
    head = [str(c.value).strip() if c.value is not None else '' for c in sh[hrow]]
    out = []
    for r in sh.iter_rows(min_row=hrow + 1, values_only=True):
        if all(v is None for v in r):
            continue
        out.append(dict(zip(head, r)))
    return head, out


def _hrow(sh, key):
    for i, r in enumerate(sh.iter_rows(min_row=1, max_row=12, values_only=True), start=1):
        if key in [str(v).strip() if v is not None else '' for v in r]:
            return i
    raise Halt(f'K-Co. 원천 시트 「{sh.title}」 머리 행을 찾지 못함', [sh.title], ['K-Co. 교정판 data/K-co 를 다시 받기'])


def load_src(src_dir):
    import openpyxl
    def wb(pat):
        f = sorted(glob.glob(os.path.join(src_dir, pat)))
        if not f:
            raise Halt('K-Co. 원천 파일이 없어 분포를 옮길 수 없음', [os.path.join(src_dir, pat)],
                       ['작업 폴더 data/K-co 에 K-Co. 교정판 6종을 넣기', '--src 로 원천 폴더 지정'])
        return openpyxl.load_workbook(f[0], data_only=True)
    S = {}
    w2, w3, w4, w5, w6 = wb('*_2_*.xlsx'), wb('*_3_*.xlsx'), wb('*_4_*.xlsx'), wb('*_5_*.xlsx'), wb('*_6_*.xlsx')
    def sheet(w, prefix):
        return next(s for s in w.worksheets if s.title.startswith(prefix))
    org = sheet(w2, '조직'); _, rows = _rows(org, _hrow(org, 'RPU 코드'))
    S['org'] = {r['RPU 코드']: r for r in rows if r.get('RPU 코드')}
    tr = sheet(w2, '연도별'); _, rows = _rows(tr, _hrow(tr, '연도'))
    S['trend'] = {int(r['연도']): r for r in rows if isinstance(r.get('연도'), (int, float))}
    rt = sheet(w2, '정년'); hd, rows = _rows(rt, _hrow(rt, 'RPU 코드'))
    S['retire_years'] = [h for h in hd if re.fullmatch(r'\d{4}', h)]
    S['retire'] = {r['RPU 코드']: r for r in rows if r.get('RPU 코드') and r['RPU 코드'] != '합계'}
    ag = sheet(w2, '연령'); hd, rows = _rows(ag, _hrow(ag, 'RPU 코드'))
    S['age_cols'] = [h for h in hd if h.endswith('대') or h.endswith('이상')]
    S['age'] = {r['RPU 코드']: r for r in rows if r.get('RPU 코드')}
    hi = sheet(w3, '입퇴사'); _, rows = _rows(hi, _hrow(hi, 'RPU 코드'))
    S['hist'] = rows
    wk = sheet(w4, '업무량'); _, rows = _rows(wk, _hrow(wk, 'RPU 코드'))
    S['work'] = rows
    asm = sheet(w4, '산정'); _, rows = _rows(asm, _hrow(asm, '항목'))
    S['assume'] = rows
    sk = sheet(w5, '스킬 인벤토리'); _, rows = _rows(sk, _hrow(sk, 'RPU 코드'))
    S['skill'] = rows
    sg = sheet(w5, '직급별'); _, rows = _rows(sg, _hrow(sg, 'RPU 코드'))
    S['skillg'] = rows
    on = sheet(w5, '스킬 온톨로지'); _, rows = _rows(on, _hrow(on, '도메인'))
    S['onto'] = rows
    lv = sheet(w5, '레벨'); _, rows = _rows(lv, _hrow(lv, '레벨'))
    S['level'] = rows
    hc = sheet(w6, 'RPU별'); hd, rows = _rows(hc, _hrow(hc, 'RPU 코드'))
    S['series_cols'] = [h for h in hd if re.fullmatch(r'\d{4}', h)]
    S['hcs'] = {r['RPU 코드']: r for r in rows if r.get('RPU 코드')}
    dv = sheet(w6, '조직별'); hd, rows = _rows(dv, _hrow(dv, 'RPU 코드'))
    S['plan_col'] = next(h for h in hd if h.endswith('계획'))
    S['drv'] = [r for r in rows if r.get('RPU 코드') and not str(r['RPU 코드']).startswith('주:')]
    td = sheet(w6, '유형'); _, rows = _rows(td, _hrow(td, '하위 유형'))
    S['typedef'] = rows
    S['grades'] = [g for g in ('G1', 'G2', 'G3', 'G4', 'G5')]
    S['sub_of'] = {}
    for r in S['drv']:
        S['sub_of'].setdefault(r['RPU 코드'], str(r['하위 유형']))
    return S


# ---------------------------------------------------------------------------
# 설정 읽기 · 원천 RPU 고르기 · 인원 채우기
# ---------------------------------------------------------------------------
def _blank(v):
    return v is None or (isinstance(v, str) and not v.strip()) or v == [] or v == {}


def _bigrams(s):
    s = re.sub(r'[\s·()\-_/0-9]', '', str(s))
    return {s[i:i + 2] for i in range(len(s) - 1)} | ({s} if len(s) == 1 else set())


def plan_rpus(cfg, S):
    src_codes = list(S['org'])
    src_hc = {c: S['org'][c]['계'] for c in src_codes}
    by_sub = {}
    for c in src_codes:
        by_sub.setdefault(S['sub_of'].get(c, ''), []).append(c)
    rpus = cfg.get('RPU') or []
    notes = []
    if _blank(rpus):   # RPU 목록이 비면 K-Co. 20개 구조를 그대로 씀
        rpus = []
        for c in src_codes:
            sub = S['sub_of'].get(c, '')
            t = next(k for k, v in TYPES.items() if v[1] == sub)
            rpus.append({'코드': c, '이름': S['org'][c]['RPU(산정 단위)'], '유형': t, '본부': S['org'][c]['본부'], '인원': None, '원천': c})
        notes.append('RPU 목록이 비어 K-Co. 20개 RPU 구조를 그대로 씀')
    out, used, codes = [], {}, set()
    for i, r in enumerate(rpus):
        code = str(r.get('코드') or '').strip() or f'V{i + 1:02d}'
        if code in codes:
            raise Halt(f'RPU 코드 「{code}」 가 두 번 나옴', [CONFIG_NAME + ' · RPU'], ['코드를 하나씩 다르게 적기'])
        codes.add(code)
        name = str(r.get('이름') or '').strip() or code
        t = TYPE_ALIAS.get(str(r.get('유형') or '').strip())
        if not t:
            raise Halt(f'RPU {code} 「유형」 을 읽을 수 없음: {r.get("유형")!r}', [CONFIG_NAME + ' · RPU · 유형'],
                       ['운영 · 영업 · 기술·전문 · 관리 지원 중 하나로 적기'])
        sub = TYPES[t][1]
        cands = by_sub.get(sub, [])
        src = str(r.get('원천') or '').strip()
        if src and src not in S['org']:
            raise Halt(f'RPU {code} 「원천」 {src} 가 K-Co. 에 없음', [', '.join(src_codes)], ['원천 칸을 비우거나 K-Co. RPU 코드로 적기'])
        if not src:
            nb = _bigrams(name)
            src = sorted(cands, key=lambda c: (-len(nb & _bigrams(S['org'][c]['RPU(산정 단위)'])), used.get(c, 0), cands.index(c)))[0]
        used[src] = used.get(src, 0) + 1
        hc = r.get('인원')
        hc = None if _blank(hc) else int(round(float(hc)))
        out.append(dict(code=code, name=name, type=t, sub=sub, typ=TYPES[t][0], bu=str(r.get('본부') or '').strip() or TYPES[t][2],
                        hc=hc, src=src))
    if not any(x['type'] != '관리 지원' for x in out):
        raise Halt('관리 지원 밖 RPU 가 하나도 없음 (관리 지원 수요는 직접 RPU 인원 비율로 정해짐)', [CONFIG_NAME + ' · RPU'],
                   ['운영 · 영업 · 기술·전문 RPU 를 하나 이상 넣기'])
    total = cfg.get('전체 인원')
    total = None if _blank(total) else int(round(float(total)))
    known = [x for x in out if x['hc'] is not None]
    unknown = [x for x in out if x['hc'] is None]
    if unknown:
        if total is not None:
            rest = total - sum(x['hc'] for x in known)
            if rest < 3 * len(unknown):
                raise Halt('전체 인원에서 적은 RPU 인원을 빼면 남는 인원이 모자람', [f'전체 {total} · 적은 RPU 합 {sum(x["hc"] for x in known)} · 빈 RPU {len(unknown)}개'],
                           ['전체 인원을 고치기', 'RPU 인원을 모두 적기'])
            for x, n in zip(unknown, apportion(rest, [src_hc[x['src']] for x in unknown])):
                x['hc'] = n
        else:
            k = (sum(x['hc'] for x in known) / sum(src_hc[x['src']] for x in known)) if known else 1.0
            for x in unknown:
                x['hc'] = max(3, rh(src_hc[x['src']] * k))
            notes.append('전체 인원이 비어 빈 RPU 인원을 K-Co. 원천 인원 × 적은 RPU 배율로 채움' if known else '인원이 모두 비어 K-Co. 인원 그대로')
        notes.append(f'인원을 비운 RPU {len(unknown)}개는 K-Co. 원천 RPU 인원 비율로 채움: ' + ' · '.join(f'{x["code"]} {x["hc"]}' for x in unknown))
    s = sum(x['hc'] for x in out)
    if total is not None and s != total:
        notes.append(f'전체 인원 {total} 과 RPU 합 {s} 이 달라 RPU 합을 씀')
    small = [x['code'] for x in out if x['hc'] < 3]
    if small:
        raise Halt('RPU 인원이 3명보다 적음 (분포를 옮길 수 없음)', small, ['RPU 를 합치거나 인원을 고치기'])
    for x in out:
        x['k'] = x['hc'] / src_hc[x['src']]
    # 본부별로 묶음 (처음 나온 순서) · 소계 행이 RPU 행 합과 맞게
    order = list(dict.fromkeys(x['bu'] for x in out))
    out.sort(key=lambda x: order.index(x['bu']))
    return out, notes


# ---------------------------------------------------------------------------
# 선택 설정 (성장 시나리오 · 사업 이벤트 · 동인 · 추가 동인 · AI 범위 · 인건비 인상률)
# 모두 비우면 opts 가 비어 지금과 같은 내용을 만든다 (결정적 · rnd 순서도 그대로)
# ---------------------------------------------------------------------------
def _strip(v):
    """'_' 로 시작하는 설명 키를 안쪽까지 뺌."""
    if isinstance(v, dict):
        return {k: _strip(x) for k, x in v.items() if not str(k).startswith('_')}
    if isinstance(v, list):
        return [_strip(x) for x in v]
    return v


def _deep_blank(v):
    v = _strip(v)
    if isinstance(v, dict):
        return all(_deep_blank(x) for x in v.values())
    if isinstance(v, list):
        return all(_deep_blank(x) for x in v)
    return _blank(v)


def _num(v, what, where):
    if _blank(v):
        return None
    try:
        return float(str(v).replace(',', '').replace('%', '').strip())
    except ValueError:
        raise Halt(f'{what} 을 숫자로 읽을 수 없음: {v!r}', [f'{CONFIG_NAME} · {where}'], ['숫자만 적기 (단위 · 쉼표 없이)', '칸을 비우기 (K-Co. 값)'])


def _ym(v, what, where):
    s = str(v or '').strip()
    m_ = re.fullmatch(r'(\d{4})[-./](\d{1,2})', s)
    if not m_ or not 1 <= int(m_.group(2)) <= 12:
        raise Halt(f'{what} 형식이 다름: {v!r}', [f'{CONFIG_NAME} · {where}'], ['YYYY-MM 으로 적기 (예: 2028-07)'])
    return int(m_.group(1)), int(m_.group(2))


def _pick(d, *keys):
    for k in keys:
        if k in d and not _blank(d[k]):
            return d[k]
    return None


def plan_options(cfg, S, rpus, ty):
    """선택 설정을 읽어 검사한다 (만들기 전 · dry-run 에서도). 반환 opts dict (빈 칸은 키가 없음)."""
    opts, codes = {}, {x['code']: x for x in rpus}
    by = ty - 2
    # 성장 시나리오
    g = _strip(cfg.get('성장 시나리오') or {})
    if isinstance(g, dict) and not _deep_blank(g):
        w = '성장 시나리오'
        base = _num(_pick(g, '기준 매출 (억 원)', '기준 매출'), '기준 매출', w)
        cagr = {k: _num((_strip(g.get('연평균 성장률 (%)') or {}) or {}).get(k), f'{k} 연평균 성장률', w + ' · 연평균 성장률 (%)') for k in ('보수', '기본', '낙관')}
        plan = {k: _num((_strip(g.get('계획 매출 (억 원)') or {}) or {}).get(k), f'{k} 계획 매출', w + ' · 계획 매출 (억 원)') for k in ('보수', '기본', '낙관')}
        margin = _num(_pick(g, '계획 영업이익률 (%)', '계획 영업이익률'), '계획 영업이익률', w)
        for k in ('보수', '기본', '낙관'):
            if cagr[k] is not None and plan[k] is not None:
                raise Halt(f'성장 시나리오 {k}: 연평균 성장률과 계획 매출을 함께 적음 (둘 중 하나만)', [f'{CONFIG_NAME} · {w}'], ['한쪽 칸을 비우기'])
            if cagr[k] is not None and not -50 < cagr[k] < 100:
                raise Halt(f'성장 시나리오 {k} 연평균 성장률 {cagr[k]}% 가 범위 밖', [f'{CONFIG_NAME} · {w}'], ['-50 ~ 100 사이 % 로 적기 (예: 6 = 연 6%)'])
        if base is not None and base <= 0:
            raise Halt('기준 매출이 0 이하', [f'{CONFIG_NAME} · {w}'], ['억 원 단위 양수로 적기'])
        if margin is not None and not -50 < margin < 80:
            raise Halt(f'계획 영업이익률 {margin}% 가 범위 밖', [f'{CONFIG_NAME} · {w}'], ['% 로 적기 (예: 7.5)'])
        opts['growth'] = dict(base=base, cagr=cagr, plan=plan, margin=margin)
    # 인건비 인상률
    r_ = _num(cfg.get('인건비 인상률 (%)', cfg.get('인건비 인상률')), '인건비 인상률', '인건비 인상률 (%)')
    if r_ is not None:
        if not -10 < r_ < 30:
            raise Halt(f'인건비 인상률 {r_}% 가 범위 밖', [f'{CONFIG_NAME} · 인건비 인상률 (%)'], ['% 로 적기 (예: 3.5)'])
        opts['raise'] = r_
    # AI 범위
    ai = _strip(cfg.get('AI 범위'))
    if not _deep_blank(ai):
        items = ai if isinstance(ai, list) else [s for s in re.split(r'[,·/]', str(ai))]
        items = [str(s).strip() for s in items if str(s).strip()]
        types, picked = [], []
        for s in items:
            if s in codes:
                picked.append(s)
            elif TYPE_ALIAS.get(s):
                types.append(TYPE_ALIAS[s])
            else:
                raise Halt(f'AI 범위 「{s}」 를 읽을 수 없음', [f'{CONFIG_NAME} · AI 범위', 'RPU 코드: ' + ' · '.join(codes)],
                           ['RPU 유형 (운영 · 영업 · 기술·전문 · 관리 지원) 또는 RPU 코드로 적기', '칸을 비우기 (K-Co. 범위 = 관리 지원 · 영업)'])
        sel = [x['code'] for x in rpus if x['code'] in picked or x['type'] in types]
        opts['ai'] = dict(types=list(dict.fromkeys(types)), codes=sel, picked=picked)
    # 사업 이벤트 (신설 조직)
    ev = _strip(cfg.get('사업 이벤트') or [])
    if not _deep_blank(ev):
        out, seen = [], set()
        for i, e in enumerate([e for e in ev if not _deep_blank(e)]):
            w = f'사업 이벤트 {i + 1}'
            code = str(e.get('코드') or '').strip()
            if not code:
                raise Halt(f'{w} 「코드」 가 비어 있음 (신설 조직 하나 = RPU 코드 하나)', [f'{CONFIG_NAME} · {w}'], ['기존 RPU 와 겹치지 않는 새 코드를 적기 (예: A09)'])
            if code in codes or code in seen:
                raise Halt(f'{w} 코드 {code} 가 기존 RPU · 다른 이벤트와 겹침 (신설 조직은 새 RPU 코드)', [f'{CONFIG_NAME} · {w}'], ['같은 본부 코드 체계의 다음 번호로 적기'])
            seen.add(code)
            name = str(e.get('이름') or '').strip()
            if not name:
                raise Halt(f'{w} 「이름」 이 비어 있음', [f'{CONFIG_NAME} · {w}'], ['예: 「신규 라인 운영 (신설)」'])
            t = TYPE_ALIAS.get(str(e.get('유형') or '운영').strip())
            if not t:
                raise Halt(f'{w} 「유형」 을 읽을 수 없음: {e.get("유형")!r}', [f'{CONFIG_NAME} · {w}'], ['운영 · 영업 · 기술·전문 · 관리 지원 중 하나'])
            n = _num(e.get('인원'), f'{w} 인원', w)
            if n is None or n <= 0:
                raise Halt(f'{w} 「인원」 이 비어 있음 (신설 조직 운영 인력 · 결정 수요 별도 줄)', [f'{CONFIG_NAME} · {w}'], ['대략 인원을 적기 (가상 값)'])
            if _blank(e.get('가동')):
                raise Halt(f'{w} 「가동」 시점이 비어 있음', [f'{CONFIG_NAME} · {w}'], ['YYYY-MM 으로 적기 (예: 2028-07)'])
            gy, gm = _ym(e.get('가동'), f'{w} 가동', w)
            if _blank(e.get('확보')):
                k_ = gy * 12 + gm - 1 - SKILL_MONTHS
                hy_, hm_ = k_ // 12, k_ % 12 + 1
            else:
                hy_, hm_ = _ym(e.get('확보'), f'{w} 확보', w)
            months = (gy * 12 + gm) - (hy_ * 12 + hm_)
            if months < 0:
                raise Halt(f'{w} 확보 시점이 가동 시점보다 늦음', [f'{CONFIG_NAME} · {w}'], ['확보 ≤ 가동 으로 적기'])
            if hy_ > ty:
                raise Halt(f'{w} 확보 시점 {hy_}년이 계획연도 {ty} 뒤 (계획연도 결정 수요에 넣을 수 없음)', [f'{CONFIG_NAME} · {w}'],
                           ['확보 시점을 계획연도 안으로', '이 이벤트를 빼고 사업계획 요점에 문장으로만 적기'])
            nm = name if '신설' in name else f'{name} (신설)'
            out.append(dict(code=code, name=nm, base=re.sub(r'\s*\(신설\)\s*$', '', nm).strip(), short=re.sub(r'\s*운영$', '', re.sub(r'\s*\(신설\)\s*$', '', nm).strip()) or nm, type=t, bu=str(e.get('본부') or '').strip() or TYPES[t][2],
                            n=int(round(n)), go=(gy, gm), hire=(hy_, hm_), months=months, mix=str(e.get('구성') or '').strip(),
                            why=str(e.get('근거') or '').strip(), auto_hire=_blank(e.get('확보'))))
        opts['events'] = out
    # 동인 (RPU 별 주 동인 바꾸기)
    dv = _strip(cfg.get('동인') or {})
    if isinstance(dv, dict) and not _deep_blank(dv):
        ch = {}
        for c, d in dv.items():
            if _deep_blank(d):
                continue
            w = f'동인 · {c}'
            if c not in codes:
                raise Halt(f'동인 칸의 RPU 코드 {c} 가 RPU 목록에 없음', [f'{CONFIG_NAME} · {w}', ' · '.join(codes)], ['RPU 목록의 코드로 적기'])
            if not isinstance(d, dict):
                raise Halt(f'{w}: 한 RPU 에 동인 하나를 {{이름 · 단위 · 기준 값 · 계획 증가율 (%) 또는 계획 값}} 으로 적기', [f'{CONFIG_NAME} · {w}'], ['양식 예를 따라 적기'])
            ch[c] = _drv_spec(d, w, codes[c], ty, need_base=False)
            ch[c]['target'] = str(d.get('바꿀 후보') or '').strip()
            ch[c]['only'] = str(d.get('다른 후보') or '').strip() in ('빼기', '뺌', '없음', '빼')   # 「다른 후보: 빼기」 = 이 RPU 는 고른 동인 한 행만 (K-Co. 원천의 나머지 후보 이름을 남기지 않음)
        if ch:
            opts['drivers'] = ch
    # 추가 동인 (RPU 별 후보 행 추가)
    ad = _strip(cfg.get('추가 동인') or {})
    if isinstance(ad, dict) and not _deep_blank(ad):
        ex = {}
        for c, lst in ad.items():
            lst = [d for d in (lst if isinstance(lst, list) else [lst]) if not _deep_blank(d)]
            if not lst:
                continue
            if c not in codes:
                raise Halt(f'추가 동인 칸의 RPU 코드 {c} 가 RPU 목록에 없음', [f'{CONFIG_NAME} · 추가 동인 · {c}', ' · '.join(codes)], ['RPU 목록의 코드로 적기'])
            ex[c] = [_drv_spec(d, f'추가 동인 · {c} · {i + 1}', codes[c], ty, need_base=True) for i, d in enumerate(lst)]
        if ex:
            opts['extras'] = ex
    return opts


def _drv_spec(d, w, x, ty, need_base):
    name = str(d.get('이름') or d.get('동인 후보') or '').strip()
    if not name:
        raise Halt(f'{w} 「이름」 이 비어 있음', [f'{CONFIG_NAME} · {w}'], ['동인 이름을 적기 (예: 처리 물량)'])
    if x['typ'] == 'Value Center' and WHOLE_CO.search(name):
        raise Halt(f'{w} 「{name}」 은 Value Center 에 회사 전체 실적 (동인 원칙 위반)', [f'{CONFIG_NAME} · {w}', f'{x["code"]} 유형 {x["type"]} = Value Center'],
                   ['지원하는 대상의 규모 (설비 수 · 과제 수 · 의뢰 건수 등)로 바꾸기', '이 칸을 비우기 (K-Co. 원천 동인)'])
    if x['type'] == '관리 지원' and '직접' in name:
        raise Halt(f'{w}: 관리 지원 「직접 RPU 인원」 행은 직접 RPU 인원 합으로 정해져 바꾸지 않음', [f'{CONFIG_NAME} · {w}'], ['다른 이름의 동인으로 적기', '칸을 비우기'])
    base = _num(_pick(d, '기준 값'), f'{w} 기준 값', w)
    if need_base and base is None:
        raise Halt(f'{w} 「기준 값」 ({ty - 2}년 값) 이 비어 있음', [f'{CONFIG_NAME} · {w}'], ['기준 연도 값을 적기 (가상 값)'])
    if base is not None and base <= 0:
        raise Halt(f'{w} 기준 값이 0 이하', [f'{CONFIG_NAME} · {w}'], ['양수로 적기'])
    rate = _num(_pick(d, '계획 증가율 (%)', '계획 증가율'), f'{w} 계획 증가율', w)
    plan = _num(_pick(d, '계획 값', f'{ty} 계획'), f'{w} 계획 값', w)
    if rate is not None and plan is not None:
        raise Halt(f'{w}: 계획 증가율과 계획 값을 함께 적음 (둘 중 하나만)', [f'{CONFIG_NAME} · {w}'], ['한쪽 칸을 비우기'])
    if need_base and rate is None and plan is None:
        raise Halt(f'{w}: {ty} 계획이 비어 있음 (계획 값이 없는 동인은 고르지 않음 · 동인 원칙)', [f'{CONFIG_NAME} · {w}'],
                   ['계획 증가율 (%) 또는 계획 값을 적기'])
    return dict(name=name, unit=str(d.get('단위') or '').strip(), base=base, rate=rate, plan=plan)


DRV_CHECK = '동인 이름이 K-Co. 원천 이름 그대로 · 회사에 맞게 바꿀지 확인'


def inherited_drivers(S, rpus, dch):
    """설정 「동인」 을 비운 RPU 의 원천 동인 이름 (K-Co. 제조 이름 그대로 쓰이는 것) · 관리 지원 「직접 RPU 인원」 은 회사와 무관해 뺌."""
    out = []
    for x in rpus:
        if x['code'] in (dch or {}):
            continue
        nm = []
        for d in [d for d in S['drv'] if d['RPU 코드'] == x['src']]:
            name = re.sub(r'^\d·?\d?공장\s*', '', str(d['동인 후보'])).strip() or str(d['동인 후보'])
            if x['type'] == '관리 지원' and '직접' in name:
                continue
            nm.append(f'{name} ({d["단위"]})' if d.get('단위') else name)
        if nm:
            out.append(f'{x["code"]} {x["name"]} ← K-Co. {x["src"]}: ' + ' · '.join(nm))
    return out


def _seed_of(*parts):
    return int(hashlib.sha256('|'.join(str(p) for p in parts).encode('utf-8')).hexdigest()[:12], 16)



# ---------------------------------------------------------------------------
# 표 만들기
# ---------------------------------------------------------------------------
def build(cfg, S, rpus, seed, opts=None):
    opts = opts or {}
    rnd = random.Random(seed)
    by = int(str(cfg['기준일'])[:4])
    off = by - int(S['series_cols'][-1])
    sy = [str(int(y) + off) for y in S['series_cols']]           # 2021~2025 → 기준 연도 기준
    hy_src = sorted({int(r['연도']) for r in S['hist']})
    hy = [y + off for y in hy_src]
    ty = by + 2
    G = S['grades']
    T = {}
    hcs, grades, hist_rows = {}, {}, []
    for x in rpus:
        c, s, k = x['code'], x['src'], x['k']
        src_series = [S['hcs'][s][y] for y in S['series_cols']]
        ser = [rh(v * k) for v in src_series[:-1]] + [x['hc']]
        ser = [max(1, v) for v in ser]
        hcs[c] = dict(zip(sy, ser))
        H = [r for r in S['hist'] if r['RPU 코드'] == s]
        def gw(y, col):
            return [next((r[col] or 0 for r in H if int(r['연도']) == y and r['직급'] == g), 0) for g in G]
        beg = apportion(hcs[c][str(hy[0] - 1)], gw(hy_src[0], '기초'))
        for ys, yt in zip(hy_src, hy):
            fl = {}
            for col in ('입사', '자발퇴직', '정년퇴직'):
                w = gw(ys, col)
                fl[col] = apportion(rh(sum(w) * k), w)
            w = gw(ys, '승진이동')[:4]
            pout = apportion(rh(sum(w) * k), w) + [0]
            pin = [0] + pout[:4]
            end = [beg[i] + fl['입사'][i] + pin[i] - fl['자발퇴직'][i] - fl['정년퇴직'][i] - pout[i] for i in range(5)]
            for i in range(5):   # 음수 직급 바로잡기 (승진이동 → 자발 → 정년 순으로 줄임)
                while end[i] < 0:
                    if pout[i] > 0:
                        pout[i] -= 1; pin[i + 1] -= 1; end[i] += 1; end[i + 1] -= 1
                    elif fl['자발퇴직'][i] > 0:
                        fl['자발퇴직'][i] -= 1; end[i] += 1
                    elif fl['정년퇴직'][i] > 0:
                        fl['정년퇴직'][i] -= 1; end[i] += 1
                    else:
                        fl['입사'][i] += 1; end[i] += 1
            r_ = hcs[c][str(yt)] - sum(end)
            if r_ > 0:
                fl['입사'][0] += r_; end[0] += r_
            while r_ < 0:   # 남는 인원은 가장 큰 직급의 자발퇴직으로
                i = max(range(5), key=lambda j: (end[j], -j))
                fl['자발퇴직'][i] += 1; end[i] -= 1; r_ += 1
            for i, g in enumerate(G):
                hist_rows.append([yt, x['bu'], c, x['name'], g, beg[i], fl['입사'][i], pin[i], fl['자발퇴직'][i], fl['정년퇴직'][i], pout[i], end[i]])
            beg = end
        grades[c] = beg
    T['hcs'], T['grades'], T['hist'], T['sy'], T['hy'], T['ty'], T['by'] = hcs, grades, hist_rows, sy, hy, ty, by
    # 연령 · 정년 · 평균
    T['age'] = {x['code']: apportion(x['hc'], [S['age'][x['src']][a] for a in S['age_cols']]) for x in rpus}
    ry = [str(int(y) + off) for y in S['retire_years']]
    T['retire_years'] = ry
    T['retire'] = {x['code']: [rh((S['retire'][x['src']].get(y) or 0) * x['k']) for y in S['retire_years']] for x in rpus}
    T['avg'] = {x['code']: (round((S['org'][x['src']]['평균 연령'] or 40) + rnd.uniform(-0.8, 0.8), 1),
                            round(max(1.0, (S['org'][x['src']]['평균 근속'] or 10) + rnd.uniform(-0.8, 0.8)), 1)) for x in rpus}
    # 전사 추이 (매출 · 이익 · 인건비 = K-Co. × 회사 배율)
    tot = sum(x['hc'] for x in rpus)
    K25 = S['trend'][int(S['series_cols'][-1])]
    cs = tot / K25['인원(명)']
    T['cs'] = cs
    trend = []
    gr = opts.get('growth')
    nyr = ty - by
    if gr:   # 성장 시나리오: 기준 매출 · 기본 성장률로 과거 매출 모양을 기울임 (K-Co. 기본 연평균 대비) · 이익 = 매출 × K-Co. 이익률
        kb = S['trend'][int(S['series_cols'][-1])]['매출']
        base = gr['base'] if gr['base'] is not None else rh(kb * cs)
        gK = KCO_PLAN['rev']['기본'] ** (1 / nyr) - 1
        if gr['cagr']['기본'] is not None:
            gB = gr['cagr']['기본'] / 100
        elif gr['plan']['기본'] is not None:
            gB = (gr['plan']['기본'] / base) ** (1 / nyr) - 1
        else:
            gB = gK
        tilt = (1 + gK) / (1 + gB)
    for yk, yt in zip(S['series_cols'], sy):
        r = S['trend'][int(yk)]
        n = sum(hcs[x['code']][yt] for x in rpus)
        if gr:
            e = int(S['series_cols'][-1]) - int(yk)
            rev = rh(r['매출'] * base / kb * tilt ** e)
            op, lab = rh(rev * r['영업이익'] / r['매출']), rh(r['총인건비'] * cs)
        else:
            rev, op, lab = rh(r['매출'] * cs), rh(r['영업이익'] * cs), rh(r['총인건비'] * cs)
        trend.append([int(yt), n, rev, op, lab, round(rev / n, 2), round(lab / n, 3), round((op + lab) / lab, 2)])
    T['trend'] = trend
    r25, op25 = trend[-1][2], trend[-1][3]
    T['plan'] = {f'{ty} 매출 {k}': rh(r25 * v) for k, v in KCO_PLAN['rev'].items()}
    T['plan'][f'{ty} 영업이익'] = rh(op25 * KCO_PLAN['op'])
    if gr:
        for k in ('보수', '기본', '낙관'):
            if gr['plan'][k] is not None:
                T['plan'][f'{ty} 매출 {k}'] = rh(gr['plan'][k])
            elif gr['cagr'][k] is not None:
                T['plan'][f'{ty} 매출 {k}'] = rh(r25 * (1 + gr['cagr'][k] / 100) ** nyr)
        if gr['margin'] is not None:
            T['plan'][f'{ty} 영업이익'] = rh(T['plan'][f'{ty} 매출 기본'] * gr['margin'] / 100)
        T['growth'] = dict(기준매출=r25, 연평균기본=round(((T['plan'][f'{ty} 매출 기본'] / r25) ** (1 / nyr) - 1) * 100, 1))
    T['plan']['인건비 인상률 (연)'] = f'{(opts["raise"] if "raise" in opts else KCO_PLAN["raise_"] * 100):.1f}%'
    T['plan'][f'{by} 매출'] = r25
    T['plan'][f'{by} 영업이익'] = op25
    # 업무량 (FTE ÷ 현원 = 원천과 같게)
    work = []
    for x in rpus:
        src = [r for r in S['work'] if r['RPU 코드'] == x['src']]
        H_star = sum(float(r['연간 빈도(회)']) * float(r['1회 소요시간(시간)']) for r in src) * x['k']
        fs = [max(1, rh(float(r['연간 빈도(회)']) * x['k'])) for r in src]
        ts = [float(r['1회 소요시간(시간)']) for r in src]
        big = max(range(len(src)), key=lambda i: (fs[i] * ts[i], -i))
        fs[big] = max(1, fs[big] + rh((H_star - sum(f * t for f, t in zip(fs, ts))) / ts[big]))
        for i, r in enumerate(src):
            a = fs[i] * ts[i]
            a = int(a) if abs(a - round(a)) < 1e-9 else round(a, 2)
            t_ = int(ts[i]) if ts[i].is_integer() else ts[i]
            nm = str(r['업무명']).replace('3공장', '신규 라인').replace('1·2공장', '기존 라인')
            work.append([x['code'], x['name'], f'{x["code"]}-{i + 1:02d}', nm, fs[i], t_, a, r['태스크 판정'], r.get('판정 근거(요약)')])
    T['work'] = work
    # 스킬 (직급 × 레벨 → 인벤토리 = 직급 합)
    LV = ['L0', 'L1', 'L2', 'L3', 'L4']
    def med(levels):
        n, cum = sum(levels), 0
        for i, v in enumerate(levels):
            cum += v
            if n and cum * 2 >= n:
                return f'L{i}'
        return 'L0'
    sk_rows, sg_rows, used_sk = [], [], set()
    for x in rpus:
        for inv in [r for r in S['skill'] if r['RPU 코드'] == x['src']]:
            nm, req = inv['스킬'], inv['요구 레벨 2028']
            used_sk.add(nm)
            colsum = [0] * 5
            for gi, g in enumerate(G):
                srow = next((r for r in S['skillg'] if r['RPU 코드'] == x['src'] and r['직급'] == g and r['스킬'] == nm), None)
                w = [srow[l] for l in LV] if srow and sum(srow[l] or 0 for l in LV) else [inv[l] for l in LV]
                lv = apportion(grades[x['code']][gi], w)
                colsum = [a + b for a, b in zip(colsum, lv)]
                n = sum(lv)
                sg_rows.append([x['code'], x['name'], g, nm, req] + lv + [n, sum(lv[req:]), med(lv)])
            n = sum(colsum)
            sk_rows.append([x['code'], x['name'], inv['도메인'], inv['그룹'], nm, inv['요구 레벨 2025'], req] + colsum + [n, sum(colsum[req:]), med(colsum)])
    T['skill'], T['skillg'] = sk_rows, sg_rows
    # 동인 시계열
    direct = [x['code'] for x in rpus if x['type'] != '관리 지원']
    dsum = {y: sum(hcs[c][y] for c in direct) for y in sy}
    drv, plan_lines, changes = [], [], []
    dch, dex = opts.get('drivers', {}), opts.get('extras', {})
    def pline(x, name, vals, plan, unit):
        plan_lines.append(f'{x["code"]} {x["name"]} · {name}: {by}년 {vals[-1]:,} {unit} → {ty}년 계획 {plan:,} {unit}')
    for x in rpus:
        srows = [d for d in S['drv'] if d['RPU 코드'] == x['src']]
        tgt = None
        if x['code'] in dch:   # 바꿀 원천 행: 「바꿀 후보」 이름 · 없으면 계획 값이 있는 첫 행 (관리 지원 「직접」 행 제외)
            spec = dch[x['code']]
            nm_ = lambda d: re.sub(r'^\d·?\d?공장\s*', '', str(d['동인 후보'])).strip() or str(d['동인 후보'])
            ok_ = [i for i, d in enumerate(srows) if not (x['type'] == '관리 지원' and '직접' in nm_(d))]
            if spec['target']:
                hit = [i for i in ok_ if spec['target'] in (nm_(srows[i]), str(srows[i]['동인 후보']))]
                if not hit:
                    raise Halt(f'동인 · {x["code"]} 「바꿀 후보」 {spec["target"]} 가 원천 후보에 없음', [CONFIG_NAME + f' · 동인 · {x["code"]}', '후보: ' + ' · '.join(nm_(srows[i]) for i in ok_)],
                               ['후보 이름 그대로 적기', '칸을 비우기 (계획 값이 있는 첫 후보를 바꿈)'])
                tgt = hit[0]
            else:
                tgt = next((i for i in ok_ if srows[i].get(S['plan_col']) is not None), ok_[0] if ok_ else None)
        for ri, r in enumerate(srows):
            if tgt is not None and ri != tgt and dch.get(x['code'], {}).get('only'):
                continue
            name = re.sub(r'^\d·?\d?공장\s*', '', str(r['동인 후보'])).strip() or str(r['동인 후보'])
            unit = str(r['단위'] or '')
            vals = [r[y] for y in S['series_cols']]
            plan = r.get(S['plan_col'])
            if x['type'] == '관리 지원' and '직접' in name and unit == '명':
                vals, plan, note = [dsum[y] for y in sy], None, '관리 지원 밖 RPU 연말 인원 합 (가상) · 2027 = 직접 RPU 방식 B 결과 합'.replace('2027', str(ty))
            else:
                k = 1.0 if unit in NOSCALE_UNITS else (cs if unit in ('억 원', '명') else x['k'])
                const = len({v for v in vals if v is not None}) <= 1
                jit = [1.0 if (const or i == len(vals) - 1) else 1 + rnd.uniform(-0.02, 0.02) for i in range(len(vals))]
                vals = [None if v is None else sv(v * j, k) for v, j in zip(vals, jit)]
                plan = None if plan is None else sv(plan, k)
                note = f'가상 · K-Co. {x["src"]} 「{r["동인 후보"]}」 분포를 옮김' + (' · 계획 값 없음' if plan is None else '')
                if gr and unit == '억 원' and REVENUE.search(name) and ri != tgt:   # 매출 동인 = 성장 시나리오 (파일 2 매출 추이 · 파일 1 기본안) · 비우면 지금처럼 K-Co. 성장
                    rev_by = {str(row[0]): row[2] for row in trend}
                    vals = [None if v is None else rev_by.get(y, v) for v, y in zip(vals, sy)]
                    plan = None if plan is None else T['plan'][f'{ty} 매출 기본']
                    note = f'가상 · 설정 「성장 시나리오」 매출 (파일 2 연도별 추이 · 파일 1 {ty} 기본안) · K-Co. {x["src"]} 「{r["동인 후보"]}」 자리'
                if ri == tgt:   # 설정 「동인」: 이름 · 단위 · 기준 값 · 계획 (원천 모양을 기준 값에 맞춤 · rnd 순서는 그대로)
                    sp, old = dch[x['code']], name
                    last = vals[-1]
                    b = sp['base'] if sp['base'] is not None else last
                    f_ = (b / last) if last else None
                    vals = [None if v is None else (sv(v, f_) if f_ else sv(b, 1.0)) for v in vals]
                    vals[-1] = sv(b, 1.0)
                    if sp['plan'] is not None:
                        plan = sv(sp['plan'], 1.0)
                    elif sp['rate'] is not None:
                        plan = sv(b * (1 + sp['rate'] / 100), 1.0)
                    elif plan is not None and last:
                        plan = sv(b * plan / last, 1.0)
                    else:
                        raise Halt(f'동인 · {x["code"]} 「{sp["name"]}」 {ty} 계획이 비어 있음 (원천 후보에도 계획 값 없음 · 계획 값이 없는 동인은 고르지 않음)',
                                   [CONFIG_NAME + f' · 동인 · {x["code"]}'], ['계획 증가율 (%) 또는 계획 값을 적기'])
                    name, unit = sp['name'], sp['unit'] or unit
                    note = f'가상 · 설정 「동인」 (K-Co. {x["src"]} 「{r["동인 후보"]}」 모양을 기준 값 · 계획에 맞춤)'
                    changes.append(dict(구분='바꿈', RPU=x['code'], 원래=old, 이름=name, 단위=unit, 기준값=vals[-1], 계획=plan))
                if plan is not None and vals[-1]:
                    pline(x, name, vals, plan, unit)
            drv.append([x['code'], x['name'], x['bu'], x['typ'], x['sub'], name, unit] + vals + [plan, note])
        for sp in dex.get(x['code'], []):   # 설정 「추가 동인」: 그 RPU 인원 추이 모양 × 기준 값 (따로 둔 시드 · 결정적)
            rg = random.Random(_seed_of(seed, x['code'], sp['name']))
            hc_ = [hcs[x['code']][y] for y in sy]
            b = sp['base']
            vals = [sv(b * h / hc_[-1] * (1 + rg.uniform(-0.03, 0.03)), 1.0) for h in hc_[:-1]] + [sv(b, 1.0)]
            plan = sv(sp['plan'], 1.0) if sp['plan'] is not None else sv(b * (1 + sp['rate'] / 100), 1.0)
            note = '가상 · 설정 「추가 동인」 (그 RPU 인원 추이 모양 · 시드 고정)' + (' · 관리 지원은 직접 RPU 인원 비율이 원칙 (참고 후보)' if x['type'] == '관리 지원' else '')
            drv.append([x['code'], x['name'], x['bu'], x['typ'], x['sub'], sp['name'], sp['unit']] + vals + [plan, note])
            changes.append(dict(구분='추가', RPU=x['code'], 원래=None, 이름=sp['name'], 단위=sp['unit'], 기준값=vals[-1], 계획=plan))
            pline(x, sp['name'], vals, plan, sp['unit'])
    T['drv'], T['plan_lines'], T['drv_changes'] = drv, plan_lines, changes
    T['drv_check'] = inherited_drivers(S, rpus, dch)
    T['ai'] = opts['ai']['codes'] if 'ai' in opts else [x['code'] for x in rpus if x['type'] in AI_TYPES]
    T['events'] = opts.get('events', [])
    T['opts'] = opts
    return T


# ---------------------------------------------------------------------------
# 파일 쓰기
# ---------------------------------------------------------------------------
def write_all(H, cfg, S, rpus, T, dst, co, ren):
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    name = cfg['회사 이름']
    day = str(cfg['기준일'])
    by, ty, sy, hy = T['by'], T['ty'], T['sy'], T['hy']
    HF, HFILL, NOTE = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor='2D2926'), Font(size=9, color='7D7D7D')
    stamp_dt = datetime.datetime(by, 12, 31)
    files, rowsn = {}, {}

    def sheet(wb, title, lines, head, rows, first=False):
        sh = wb.active if first else wb.create_sheet()
        sh.title = title[:31]
        sh.cell(1, 1, lines[0]).font = Font(bold=True, size=11)
        sh.cell(2, 1, H.VIRTUAL_STAMP + ' · pwc-hr-harness 가상 데이터 만들기 (K-Co. 분포를 규모에 맞춰 옮김)').font = NOTE
        if len(lines) > 1 and lines[1]:
            sh.cell(3, 1, lines[1]).font = NOTE
        for j, h in enumerate(head, 1):
            c = sh.cell(5, j, ren.get(h, h))
            c.font, c.fill = HF, HFILL
        for i, r in enumerate(rows):
            for j, v in enumerate(r, 1):
                sh.cell(6 + i, j, v)
        rowsn[(fname, title)] = sum(1 for r in rows if r and r[0] is not None and not str(r[0]).startswith('주:'))
        return sh

    def save(wb, fn):
        wb.properties.keywords = H.VIRTUAL_TAG
        wb.properties.subject = H.VIRTUAL_STAMP
        wb.properties.title = f'{name} 가상 데이터'
        wb.properties.creator = 'pwc-hr-harness virtual_data.py'
        wb.properties.created = wb.properties.modified = stamp_dt
        wb.save(os.path.join(dst, fn))
        files[fn] = wb

    def grouped(make_row, ncol_num):
        """본부별 행 + 소계 · 전사 합계 (소계 = RPU 행 합)."""
        rows, acc_all = [], None
        for bu in dict.fromkeys(x['bu'] for x in rpus):
            acc = None
            for x in [r for r in rpus if r['bu'] == bu]:
                row = make_row(x)
                rows.append(row)
                nums = row[-ncol_num:]
                acc = nums if acc is None else [a + b for a, b in zip(acc, nums)]
            rows.append([f'{bu} 소계'] + [None] * (len(rows[-1]) - ncol_num - 1) + acc)
            acc_all = acc if acc_all is None else [a + b for a, b in zip(acc_all, acc)]
        return rows, acc_all

    # ---- 파일 2
    fname = f'{co}_2_조직인원현황.xlsx'
    wb = openpyxl.Workbook()
    rows, acc = grouped(lambda x: [x['bu'], x['code'], x['name']] + T['grades'][x['code']] + [sum(T['grades'][x['code']])], 6)
    for r in rows:   # 평균 연령 · 근속 (집계 행은 빈칸)
        c = r[1]
        r += list(T['avg'][c]) if c else [None, None]
    rows.append(['전사 합계', None, None] + acc + [None, None])
    sheet(wb, '조직·인원 현황', [f'{name} 조직·인원 현황 ({day} 기준)', '단위: 명 · 직급 G1 사원 / G2 대리 / G3 과장 / G4 차장 / G5 부장 · 개인 식별정보 없음 (집계 단위)'],
          ['본부', 'RPU 코드', 'RPU(산정 단위)', 'G1', 'G2', 'G3', 'G4', 'G5', '계', '평균 연령', '평균 근속'], rows, first=True)
    tr = [r + [] for r in T['trend']] + [[None] * 8, ['주: HCROI = (영업이익 + 총인건비) ÷ 총인건비 · 매출 · 이익 · 인건비 = K-Co. 값 × 회사 인원 배율'] + [None] * 7]
    sheet(wb, '연도별 추이', [f'{name} 연도별 경영·인원 추이 (연말 기준)', '매출·영업이익·인건비 단위: 억원 · 인건비는 급여·상여·퇴직급여·복리후생 포함'],
          ['연도', '인원(명)', '매출', '영업이익', '총인건비', '1인당 매출', '1인당 인건비', 'HCROI'], tr)
    rr = [[x['code'], x['name']] + T['retire'][x['code']] for x in rpus]
    rr.append(['합계', None] + [sum(T['retire'][x['code']][i] for x in rpus) for i in range(len(T['retire_years']))])
    sheet(wb, '정년 도래 예정', ['RPU별 정년 도래 예정 인원 (만 60세 · 연도별)', '단위: 명'], ['RPU 코드', 'RPU'] + T['retire_years'], rr)
    ar = [[x['code'], x['name']] + T['age'][x['code']] + [x['hc']] for x in rpus]
    sheet(wb, '연령대 분포', [f'RPU별 연령대 분포 ({day})', '단위: 명'], ['RPU 코드', 'RPU'] + S['age_cols'] + ['계'], ar)
    save(wb, fname)

    # ---- 파일 3
    fname = f'{co}_3_입퇴사이력_{hy[0]}-{hy[-1]}.xlsx'
    wb = openpyxl.Workbook()
    sheet(wb, '입퇴사 이력', [f'{name} 입퇴사·이동 이력 {hy[0]}~{hy[-1]} (RPU × 직급 × 연도)',
                             '단위: 명 · 기말 = 기초 + 입사 + 승진유입 − 자발퇴직 − 정년퇴직 − 승진이동 · 사람 단위 행 없음 (집계)'],
          ['연도', '본부', 'RPU 코드', 'RPU', '직급', '기초', '입사', '승진유입', '자발퇴직', '정년퇴직', '승진이동', '기말'], T['hist'], first=True)
    sm = []
    for y in hy:
        rs = [r for r in T['hist'] if r[0] == y]
        b, i_, v, re_, e = (sum(r[k] for r in rs) for k in (5, 6, 8, 9, 11))
        sm.append([y, b, i_, v, re_, e, round(v / b * 100, 1) if b else None])
    sheet(wb, '연도별 요약', ['전사 연도별 요약', '단위: 명 · 자발 퇴직률(%) = 자발퇴직 ÷ 기초 × 100'], ['연도', '기초', '입사', '자발퇴직', '정년퇴직', '기말', '자발 퇴직률(%)'], sm)
    save(wb, fname)

    # ---- 파일 4
    fname = f'{co}_4_업무량조사.xlsx'
    wb = openpyxl.Workbook()
    sheet(wb, '업무량 조사', [f'{name} 업무량 조사 결과 ({by}년 조사 · RPU별 업무 목록)',
                             '연간 소요시간 = 연간 빈도 × 1회 소요시간 · 1회 소요시간 단위: 시간 · 태스크 판정: 위임(AI 단독) / 함께(AI + 사람) / 유지(사람)'],
          ['RPU 코드', 'RPU', '업무 ID', '업무명', '연간 빈도(회)', '1회 소요시간(시간)', '연간 소요시간(시간)', '태스크 판정', '판정 근거(요약)'], T['work'], first=True)
    sheet(wb, '산정 가정', ['업무량 기반 인력 산정 가정 (K-Co. 실습 기본값과 같음)', ''], ['항목', '값', '설명'],
          [[r['항목'], r['값'], r.get('설명')] for r in S['assume']])
    save(wb, fname)

    # ---- 파일 5
    fname = f'{co}_5_스킬인벤토리.xlsx'
    wb = openpyxl.Workbook()
    sheet(wb, '스킬 인벤토리', [f'{name} 스킬 인벤토리 (직무 = RPU · {by} 자가진단 + 팀장 확인)',
                               '보유 분포 단위: 명 · L0 미보유 / L1 기초 / L2 숙련 / L3 전문 / L4 최고 · 2028 요구 레벨은 K-Co. 같은 유형 RPU 의 요구를 옮김'],
          ['RPU 코드', '직무(RPU)', '도메인', '그룹', '스킬', '요구 레벨 2025', '요구 레벨 2028', 'L0', 'L1', 'L2', 'L3', 'L4', '인원 계', '2028 요구 충족 인원', '현재 레벨'],
          T['skill'], first=True)
    sheet(wb, '직급별 보유 분포', ['RPU × 직급 × 스킬 보유 분포 (스킬 인벤토리 시트를 직급별로 나눈 것 · 직급 합 = RPU 행)', '현재 레벨 = 보유 분포의 중앙값 레벨'],
          ['RPU 코드', '직무(RPU)', '직급', '스킬', '요구 레벨 2028', 'L0', 'L1', 'L2', 'L3', 'L4', '인원 계', '2028 요구 충족 인원', '현재 레벨'], T['skillg'])
    sheet(wb, '스킬 온톨로지', [f'{name} 스킬 온톨로지 (도메인 → 그룹 → 스킬 → 레벨)', ''], ['도메인', '그룹', '스킬'],
          [[r['도메인'], r['그룹'], r['스킬']] for r in S['onto']])
    sheet(wb, '레벨 정의', ['스킬 레벨 정의 (공통)', ''], ['레벨', '명칭', '정의'], [[r['레벨'], r['명칭'], r['정의']] for r in S['level']])
    save(wb, fname)

    # ---- 파일 6
    fname = f'{co}_6_인력동인시계열_{sy[0]}-{sy[-1]}.xlsx'
    wb = openpyxl.Workbook()
    rows, acc = grouped(lambda x: [x['bu'], x['code'], x['name']] + [T['hcs'][x['code']][y] for y in sy], len(sy))
    d_acc = [sum(T['hcs'][x['code']][y] for x in rpus if x['type'] != '관리 지원') for y in sy]
    rows += [['직접 RPU 합계 (관리 지원 제외)', None, None] + d_acc, ['전사 합계', None, None] + acc]
    sheet(wb, f'RPU별 연말 인원 {sy[0]}~{sy[-1]}', [f'{name} RPU별 연말 인원 {sy[0]}~{sy[-1]}', '단위: 명 · 연말(12-31) 재직 기준 · 전사 합계 = 파일 2 연도별 추이 인원'],
          ['본부', 'RPU 코드', 'RPU'] + sy, rows, first=True)
    dr = T['drv'] + [[None] * 14, [f'주: 관리 지원 RPU 의 「직접 RPU 인원」 {ty} 값은 직접 RPU 산정 결과 합으로 정해지므로 비워 둠'] + [None] * 13]
    sheet(wb, f'조직별 인력동인 시계열 {sy[0]}~{sy[-1]}', [f'{name} 조직별 인력동인 시계열 {sy[0]}~{sy[-1]} (RPU마다 후보 동인 2~3개)',
          f'동인 = 조직의 업무량을 끌고 가는 사업 지표 · {ty} 계획 = 파일 1 사업계획 2절 · 빈칸은 {ty} 계획 값이 없음 · 유형은 Profit Center / Value Center 와 그 안의 하위 유형'],
          ['RPU 코드', 'RPU', '본부', '유형', '하위 유형', '동인 후보', '단위'] + sy + [f'{ty} 계획', '비고'], dr)
    tdr = []
    for r in S['typedef']:
        cs_ = [x['code'] for x in rpus if x['sub'] == r['하위 유형']]
        tdr.append([r['유형'], r['하위 유형'], r['동인 선정 원칙'], ' · '.join(cs_) or '-'])
    sheet(wb, '유형 정의', ['조직 유형 정의 (방식 B 동인 선정 원칙)', ''], ['유형', '하위 유형', '동인 선정 원칙', '해당 RPU'], tdr)
    save(wb, fname)

    # ---- 파일 1 (문서)
    fname = f'{co}_1_사업계획요약.md'
    P = T['plan']
    pts = [str(p).strip() for p in (cfg.get('사업계획 요점') or []) if str(p).strip()]
    EV = T.get('events') or []
    ym = lambda t: f'{t[0]}년 {t[1]}월'
    ev_lines = []
    for e in EV:   # K-Co. 파일 1 · 2절 「3공장 신설」 문장과 같은 꼴 (수요 담당이 ②_수요파라미터 4 · 5절 행으로 옮김)
        ev_lines.append(f'- {e["short"]} 신설 (신설 조직 코드 {e["code"]} · {e["bu"]} · {e["type"]} 유형): {ym(e["go"])} 가동 목표. '
                        f'운영 인력은 약 {e["n"]:,}명' + (f' ({e["mix"]})' if e['mix'] else '') + '으로 추정. '
                        f'숙련에 {e["months"]}개월이 걸려 {ym(e["hire"])}까지 채용·배치를 마쳐야 함' + (f' · {e["why"]}' if e['why'] else ''))
    if 'ai' in T['opts']:
        ai_ = T['opts']['ai']
        nm_of = {x['code']: x['name'] for x in rpus}
        desc = ' · '.join(([' · '.join(ai_['types']) + ' 유형 RPU'] if ai_['types'] else []) + [f'{c} {nm_of[c]}' for c in ai_['picked'] if c in nm_of])
    else:
        desc = '관리 지원 · 영업 유형 RPU'
    gr_ = T.get('growth')
    doc = [f'# {name} {ty}~{ty + 2} 중기 사업계획 요약 (가상)', '',
           f'> {H.VIRTUAL_STAMP} · K-Co. 사업계획의 성장률 · 비율을 {name} 규모에 맞춰 옮긴 가상 문서' + (' (성장 시나리오 · 사업 이벤트 등은 설정 파일 값)' if T['opts'] else ''), '',
           '## 1. 경영 목표', '',
           f'- 매출: {by}년 실적 {P[f"{by} 매출"]:,}억 원 → {ty}년 목표 {P[f"{ty} 매출 기본"]:,}억 원 (기본안' + (f' · 연평균 {gr_["연평균기본"]:.1f}% 성장)' if gr_ else ')'),
           f'- 시나리오: 보수안 {ty}년 매출 {P[f"{ty} 매출 보수"]:,}억 원 · 낙관안 {P[f"{ty} 매출 낙관"]:,}억 원',
           f'- 영업이익: {by}년 {P[f"{by} 영업이익"]:,}억 원 → {ty}년 {P[f"{ty} 영업이익"]:,}억 원', '',
           '## 2. 조직별 사업 지표 계획', '',
           f'- RPU별 {ty} 사업 지표 계획은 아래와 같다 (파일 6 「{ty} 계획」 열과 같은 값 · K-Co. 같은 유형 RPU 의 증가율을 옮김' + (' · 설정 「동인」 「추가 동인」 행은 설정 값)' if T.get('drv_changes') else ')'),
           *[f'- {ln}' for ln in T['plan_lines']],
           *(['- 사업 이벤트 (업무량 조사 · 파일 2~6 에 아직 없는 신설 조직 · 결정 수요에 별도 줄):'] + ev_lines if EV else []), '',
           '## 3. 디지털·AI 전환', '',
           f'- {ty}년 상반기 생성형 AI 플랫폼 도입. 우선 적용 범위는 {desc} 의 반복 업무 ({" · ".join(T["ai"]) or "해당 없음"})',
           f'- 그 밖의 RPU 적용은 {ty + 1}년 이후 검토', '',
           '## 4. 인력 운영 방향', '',
           ('- 자동화와 AI 로 생기는 여유 인력은 감원하지 않고 부족한 RPU 로 재배치를 먼저 검토' if not EV else
            f'- 자동화와 AI 로 생기는 여유 인력은 감원하지 않고 부족한 RPU · 신설 조직 ({" · ".join(e["code"] for e in EV)}) 으로 재배치를 먼저 검토 · 신설 조직은 선제적으로 확보'),
           f'- {ty}년 인건비 인상률은 연 {P["인건비 인상률 (연)"]} 가정', '',
           '## 5. 회사가 적은 요점 (설정 파일 그대로)', '',
           '- 아래 문장은 참가자가 설정 파일에 적은 그대로다. 숫자가 1~4절과 다르면 1~4절이 이 가상 데이터의 기준이다',
           *([f'- {p}' for p in pts] or ['- 적은 요점 없음']), '',
           H.VIRTUAL_STAMP]
    open(os.path.join(dst, fname), 'w', encoding='utf-8').write('\n'.join(doc) + '\n')
    files[fname] = None
    rowsn[(fname, '문서')] = len(T['plan_lines'])
    return files, rowsn


def opt_summary(opts):
    """선택 설정 요약 (dry-run · 표시 파일 · 데이터 사전). 쓰지 않은 칸은 「K-Co. 값」."""
    out = {}
    g = opts.get('growth')
    if g:
        parts = []
        for k in ('보수', '기본', '낙관'):
            parts.append(f'{k} 연 {g["cagr"][k]:g}%' if g['cagr'][k] is not None else f'{k} {g["plan"][k]:,.0f}억 원' if g['plan'][k] is not None else f'{k} K-Co. 성장률')
        if g['base'] is not None:
            parts.append(f'기준 매출 {g["base"]:,.0f}억 원')
        if g['margin'] is not None:
            parts.append(f'계획 영업이익률 {g["margin"]:g}%')
        out['성장 시나리오'] = '설정 값 · ' + ' · '.join(parts)
    else:
        out['성장 시나리오'] = 'K-Co. 성장률'
    out['사업 이벤트'] = [f'{e["code"]} {e["name"]} · {e["n"]}명 · 가동 {e["go"][0]}-{e["go"][1]:02d} · 확보 {e["hire"][0]}-{e["hire"][1]:02d}'
                       + (f' (확보 = 가동 {SKILL_MONTHS}개월 전 · K-Co. 숙련 기간)' if e['auto_hire'] else '') for e in opts.get('events', [])] or '없음'
    out['동인'] = [f'{c}: {d["name"]}' + (f' ({d["unit"]})' if d['unit'] else '') for c, d in opts.get('drivers', {}).items()] or 'K-Co. 원천 그대로'
    out['추가 동인'] = [f'{c}: {d["name"]}' + (f' ({d["unit"]})' if d['unit'] else '') for c, lst in opts.get('extras', {}).items() for d in lst] or '없음'
    out['AI 범위'] = (' · '.join(opts['ai']['types'] + opts['ai']['picked']) + ' → ' + (' · '.join(opts['ai']['codes']) or '해당 없음')) if 'ai' in opts else 'K-Co. 범위 (관리 지원 · 영업 유형)'
    out['인건비 인상률'] = f'{opts["raise"]:g}%' if 'raise' in opts else f'K-Co. {KCO_PLAN["raise_"] * 100:.1f}%'
    return out


def opt_lines(opts):
    return [f'{k}: ' + (' / '.join(v) if isinstance(v, list) else v) for k, v in opt_summary(opts).items()]


def event_rows(T):
    """사업 이벤트 → ②_수요파라미터.md 4절 · 5절 행 꼴 (수요 담당이 파일 1 · 2절을 근거로 옮길 값 · 가정 구분 = 사업계획)."""
    out = []
    for e in T.get('events') or []:
        hire = f'{e["hire"][0]}-{e["hire"][1]:02d}'
        out.append({'4절': [e['code'], e['name'], str(e['n']), f'{hire}까지', '파일 1 · 2절', '사업계획'],
                    '5절': [e['code'], f'{e["short"]} 운영 인력', hire, '파일 1 · 2절', '사업계획'],
                    '가동': f'{e["go"][0]}-{e["go"][1]:02d}', '구성': e['mix'], '유형': e['type'], '본부': e['bu']})
    return out


def content_hash(path):
    h = hashlib.sha256()
    if path.endswith('.xlsx'):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        for sh in wb.worksheets:
            h.update(sh.title.encode('utf-8'))
            for r in sh.iter_rows(values_only=True):
                h.update(repr(r).encode('utf-8'))
    else:
        h.update(open(path, 'rb').read())
    return h.hexdigest()[:16]


def blank_dict(path):
    """데이터 사전이 빈 양식인가 (표 데이터 행마다 채운 칸이 번호 하나 이하)."""
    if not os.path.exists(path):
        return True
    lines = open(path, encoding='utf-8').read().splitlines()
    for i, ln in enumerate(lines):
        if not ln.startswith('|') or re.match(r'^\|\s*:?-{2,}', ln):
            continue
        if i + 1 < len(lines) and re.match(r'^\|\s*:?-{2,}', lines[i + 1]):
            continue   # 표 머리 행
        cells = [c.strip() for c in ln.strip().strip('|').split('|')]
        if any(c and not re.fullmatch(r'\d+', c) for c in cells):
            return False
    return True


def main():
    ap = argparse.ArgumentParser(description='가상 데이터 만들기 (자사 파일이 없을 때)')
    ap.add_argument('--co', default='자사')
    ap.add_argument('--config', default=None, help=f'설정 파일 (기본 data/<co>/{CONFIG_NAME})')
    ap.add_argument('--src', default='K-co', help='K-Co. 원천 폴더 이름 (data/ 아래) 또는 경로')
    ap.add_argument('--root', default=None)
    ap.add_argument('--dry-run', action='store_true')
    ap.add_argument('--replace-virtual', action='store_true', help='전에 만든 가상 데이터만 지우고 다시 만듦 (실제 파일이 섞여 있으면 멈춤)')
    a = ap.parse_args()
    root = find_root(a.root or os.getcwd())
    H = _lib(root)
    dst = os.path.join(root, 'data', a.co)
    cfgp = a.config or os.path.join(dst, CONFIG_NAME)
    try:
        if not os.path.exists(cfgp):
            raise Halt('가상 데이터 설정 파일이 없음', [cfgp], [f'`양식/{FORM_NAME}` 을 `data/{a.co}/{CONFIG_NAME}` 으로 복사해 채우기', '리드가 묻는 말에 답해 설정 파일을 같이 채우기'])
        cfg = json.load(open(cfgp, encoding='utf-8'))
        cfg = {k: v for k, v in cfg.items() if not k.startswith('_')}
        if _blank(cfg.get('회사 이름')):
            cfg['회사 이름'] = a.co
        if _blank(cfg.get('기준일')):
            cfg['기준일'] = '2025-12-31'
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', str(cfg['기준일'])):
            raise Halt('기준일 형식이 다름', [str(cfg['기준일'])], ['YYYY-MM-DD 로 적기 (예: 2025-12-31)'])
        ren = {k: str(v).strip() for k, v in (cfg.get('열이름_대응') or {}).items() if not k.startswith('_') and str(v).strip()}
        bad = [f'{k} → {v}' for k, v in ren.items() if BAN_COL.search(v) or BAN_COL.search(k)]
        if bad:
            raise Halt('열 이름 대응에 개인 식별 열 이름이 있음 (가상 데이터에도 만들지 않음)', bad, ['그 대응을 지우기'])
        srcd = a.src if os.path.isdir(a.src) else os.path.join(root, 'data', a.src)
        S = load_src(srcd)
        rpus, notes = plan_rpus(cfg, S)
        seed = int(cfg.get('시드') or 1)
        cfgv = H.config()
        by = int(str(cfg['기준일'])[:4])
        if by != int(cfgv['base_year']):
            notes.append(f'기준 연도 {by} 가 .claude/hr-lib/config.json base_year {cfgv["base_year"]} 와 다름 → B3 에서 base_year {by} · target_year {by + 2} · hist_years · series_years 를 맞춤')
        plan = [dict(코드=x['code'], 이름=x['name'], 유형=x['type'], 본부=x['bu'], 인원=x['hc'], 원천=x['src'],
                     원천이름=S['org'][x['src']]['RPU(산정 단위)'], 배율=round(x['k'], 4)) for x in rpus]
        opts = plan_options(cfg, S, rpus, by + 2)
        if a.dry_run:
            chk = inherited_drivers(S, rpus, opts.get('drivers', {}))
            print(json.dumps(dict(status='미리 보기', co=a.co, 전체=sum(x['hc'] for x in rpus), rpu=plan, 선택설정=opt_summary(opts), notes=notes,
                                  **({'동인 확인': dict(뜻=DRV_CHECK, 묻기='RPU 마다 「이 이름 그대로 쓸까, 회사 지표 이름 · 단위 · 기준 값 · 계획을 줄까」 (지어 채우지 않음)', RPU=chk)} if chk else {})),
                             ensure_ascii=False, indent=1))
            return 0
        # 기존 자사 파일 확인 (덮지 않음 · README · 설정 · .gitkeep · 빈 데이터 사전 · 추가_* 은 그대로 둠)
        keep = KEEP_FILES | {os.path.basename(cfgp)}
        exist = [f for f in sorted(os.listdir(dst))] if os.path.isdir(dst) else []
        extra = [f for f in exist if f.startswith(EXTRA_PREFIX)]
        exist = [f for f in exist if f not in keep and not f.startswith(EXTRA_PREFIX) and not (f == '데이터사전.md' and blank_dict(os.path.join(dst, f)))]
        prev = H.virtual_info(dst)
        if exist:
            listed = set((prev['info'].get('파일') or {}).keys()) | {H.VIRTUAL_MARK}
            only_virtual = prev['marker'] and all(f in listed for f in exist)
            if not (a.replace_virtual and only_virtual):
                raise Halt(f'data/{a.co}/ 에 이미 파일이 있어 덮지 않음', exist[:10],
                           (['전에 만든 가상 데이터를 다시 만들기 (`--replace-virtual`)'] if only_virtual else ['실제 자사 파일이면 가상 데이터를 만들지 않고 그 파일로 B2 진행']) +
                           ['다른 회사 폴더 이름으로 만들기 (`--co 자사가상`)', '보류'])
            for f in exist:
                os.remove(os.path.join(dst, f))
        os.makedirs(dst, exist_ok=True)
        T = build(cfg, S, rpus, seed, opts)
        files, rowsn = write_all(H, cfg, S, rpus, T, dst, a.co, ren)
        written = list(files)
        if ren:
            json.dump(dict({'_설명': '가상 데이터 설정 「열이름_대응」 을 옮김 (virtual_data.py) · 왼쪽 = 기본 열 · 오른쪽 = 자사 열'}, **ren),
                      open(os.path.join(dst, '열이름_대응.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
            written.append('열이름_대응.json')
        dname = '데이터사전.md' if blank_dict(os.path.join(dst, '데이터사전.md')) else '데이터사전_가상.md'
        dd = [f'# {cfg["회사 이름"]} 데이터 사전 (가상)', '', f'- {H.VIRTUAL_STAMP} · 만든 도구 virtual_data.py · 표시 파일 `{H.VIRTUAL_MARK}`',
              '- 필드 뜻 · 단위 · 결측 규칙은 K-Co. 데이터 사전 (`data/K-co/데이터사전.md`)과 같다. 열 이름만 아래 대응표를 따른다',
              f'- 기준일 {cfg["기준일"]} (시점 시트) · 기간 값은 연도별 1~12월', '', '## 0. 파일 목록', '',
              '| 번호 | 자사 파일 이름 | K-Co. 대응 파일 | 기준일 | 민감도 | 비고 |', '|---|---|---|---|---|---|']
        for f in written:
            m_ = re.match(r'.+?_(\d)_', f)
            if m_:
                dd.append(f'| {m_.group(1)} | {f} | 파일 {m_.group(1)} | {cfg["기준일"]} | 사내 (가상) | 가상 데이터 |')
        dd += ['', '## 1. 열 이름 대응 (기본 열 → 자사 열)', '', '| 기본 열 | 자사 열 |', '|---|---|'] + [f'| {k} | {v} |' for k, v in ren.items()] + \
              (['| (없음) | 기본 열 이름 그대로 |'] if not ren else []) + ['', '## 2. 만든 방법', '',
              '- RPU 마다 같은 유형의 K-Co. RPU 하나를 원천으로 고르고, 인원 배율로 직급 · 연령 · 입퇴사 · 업무량 · 스킬 · 동인 분포를 옮김',
              '- 개인 식별 열 없음 · 모든 행은 RPU · 직급 집계 단위', '']
        if opts:   # 선택 설정을 쓴 경우만 (비우면 지금과 같은 사전)
            dd += ['## 3. 설정으로 바꾼 것 (가상 · 설정 파일 값)', ''] + [f'- {ln}' for ln in opt_lines(opts)] + ['']
            if T['drv_check']:
                dd += [f'### 동인 확인 ({DRV_CHECK})', '', '- 「동인」 칸을 비운 RPU 는 K-Co. 원천 동인 이름 · 단위를 그대로 씀. 회사 지표가 다르면 설정 「동인」 칸에 이름 · 단위 · 기준 값 · 계획을 적고 `--replace-virtual` 로 다시 만듦', '']
                dd += [f'- {ln}' for ln in T['drv_check']] + ['']
        if extra:
            dd += ['## 4. 추가 자료 (추가_ 파일 · 가상 데이터 만들기가 만들지 않음 · 숫자 계산에 쓰지 않음)', ''] + [f'- {f}' for f in extra] + ['']
        open(os.path.join(dst, dname), 'w', encoding='utf-8').write('\n'.join(dd) + '\n')
        written.append(dname)
        hashes = {f: dict(내용지문=content_hash(os.path.join(dst, f)),
                          행수={sh: n for (ff, sh), n in rowsn.items() if ff == f}) for f in written}
        marker = {'표시': H.VIRTUAL_STAMP, '리포트 첫 줄': H.VIRTUAL_LINE,
                  '설명': '이 폴더는 pwc-hr-harness 가상 데이터 만들기로 만든 것이다. 이 파일을 지우거나 옮기면 검증 (verify)이 관문으로 멈춘다',
                  '만든 도구': 'virtual_data.py', '만든 날': datetime.date.today().isoformat(), '시드': seed,
                  '원천': 'K-Co. 교정판 (data/' + os.path.basename(srcd) + ') 분포를 옮김', '설정': cfg,
                  'RPU': plan, '사업계획 숫자': T['plan'], 'AI 적용 RPU': T['ai'],
                  '선택 설정': opt_summary(opts), '사업 이벤트': event_rows(T), '동인 변경': T['drv_changes'],
                  '동인 확인': dict(뜻=DRV_CHECK, RPU=T['drv_check']) if T['drv_check'] else '없음',
                  '추가 자료': extra, '주의': notes, '파일': hashes}
        json.dump(marker, open(os.path.join(dst, H.VIRTUAL_MARK), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
        print(json.dumps(dict(status='완료', co=a.co, 전체=sum(x['hc'] for x in rpus), rpu=len(rpus),
                              files={f: v['행수'] for f, v in hashes.items()}, marker=f'data/{a.co}/{H.VIRTUAL_MARK}', notes=notes,
                              **({'동인 확인': dict(뜻=DRV_CHECK, RPU=T['drv_check'])} if T['drv_check'] else {}),
                              next=f'「{a.co} 열 대응표부터 만들어 줘」 (step1.py scan --co {a.co}) → CP1 → ① ~ ④ · 결과는 가상 표시가 붙음'),
                         ensure_ascii=False, indent=1))
        return 0
    except Halt as h:
        print(json.dumps(dict(status='멈춤', kind='가상 데이터', what=h.what, evidence=h.evidence, options=h.options,
                              answer='확정 · 추가 확인 · 보류 · 이관'), ensure_ascii=False, indent=1))
        return EXIT_STOP


if __name__ == '__main__':
    sys.exit(main())
