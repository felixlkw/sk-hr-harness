# -*- coding: utf-8 -*-
"""② 인력수요 예측 · 방식 A (매출·이익) · B (조직별 인력동인) · C (업무량) · 세 값 대사.

사용:
  python -X utf8 .claude/skills/wfp-step2-demand/scripts/step2.py params --co K-co   # CP2 자료 초안
  python -X utf8 .claude/skills/wfp-step2-demand/scripts/step2.py run    --co K-co   # CP2 승인 뒤 계산
  python -X utf8 .claude/skills/wfp-step2-demand/scripts/step2.py run --part A --co K-co   # CP2 전 방식 A 만 (1절이 채워졌을 때)

params : CP1 승인 확인 → ②_수요파라미터.md 초안 (동인 Pool 상관 순위 · 원칙 점검 · 제안 동인 · 동인 비율).
         사업계획(파일 1)에서 읽어야 하는 칸은 비워 두고, 수요 담당이 근거 문장과 함께 채운다. → CP2 에서 멈춤
         파일이 이미 있으면 덮지 않고 다시 읽어 CP2 멈춤을 낸다: 「가정 구분」이 「회사 가정」인 행 목록 (3 · 4 · 5절 · 파일 순서)을
         _멈춤.md 에 그대로 적는다 (담당이 빈칸을 채운 뒤 · 사람 답을 반영한 뒤 다시 돌림 · 리드는 이 목록을 다시 만들지 않고 그대로 보여 줌)
run    : CP2 승인 확인 → ②A_매출이익.xlsx · ②B_인력동인.xlsx · ②C_업무량.xlsx · ②_대사.xlsx
         업무량 원천 이상(빈도×1회 ≠ 연간 · FTE÷현원 범위 밖)은 「예외」로 멈춘다.
"""
import os, sys, json, re
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', '..', 'hr-lib'))
import hrlib as H

SCRIPT = 'wfp-step2-demand/scripts/step2.py'
PF = '②_수요파라미터.md'
WHOLE_CO = re.compile(r'전사|회사 전체|매출|영업이익')   # Value Center 에 쓰면 안 되는 회사 전체 실적 동인


def _hcy(ws, m):
    sy = [str(y) for y in ws.cfg['series_years']]
    out = {}
    for r in H.load(ws, 'hcs', m):
        out[r['RPU 코드']] = {y: r.get(y) for y in sy}
    return out


def _pool(ws, m):
    """RPU → [후보 dict] · 빈칸은 0 으로 읽지 않는다 (공통 규칙 4)."""
    sy = [str(y) for y in ws.cfg['series_years']]
    ty = f'{ws.cfg["target_year"]} 계획'
    pool, blanks = {}, []
    for r in H.load(ws, 'drv', m):
        vals = [r.get(y) for y in sy]
        miss = [y for y, v in zip(sy, vals) if v is None]
        if miss:
            blanks.append((r['RPU 코드'], H.norm(r['동인 후보']), miss, r['_ref']))
        pool.setdefault(r['RPU 코드'], []).append(dict(
            name=H.norm(r['동인 후보']), unit=H.norm(r.get('단위')), typ=H.norm(r.get('유형')), sub=H.norm(r.get('하위 유형')),
            vals=vals, v27=r.get(ty), ref=r['_ref']))
    return pool, blanks


def _principle(d):
    """동인 선정 원칙 점검 (법전 · 방법론: 동인 원칙 · Profit/Value Center)."""
    if '관리 지원' in d['sub']:
        return ('부합 (비율 방식)' if '직접' in d['name'] else '참고 (관리 지원은 직접 RPU 인원 비율)')
    if d['typ'].startswith('Value') and WHOLE_CO.search(d['name']):
        return '위반: Value Center 에 회사 전체 실적'
    if d['v27'] is None:
        return '위반: 2027 계획 값 없음'
    return '부합'


ASSUME_CO = '회사 가정'
_DEF3 = {'보정 계수': ('', '1', '1.0'), '가산 인원': ('', '0'), 'AI 적용': ('', '×', 'x', 'X'), 'C 계수': ('', '(자동)')}


def _cell(s, n=60):
    s = re.sub(r'\s+', ' ', (s or '').replace('|', '/')).strip()
    return s if len(s) <= n else s[:n - 1] + '…'


def assume_rows(text):
    """②_수요파라미터.md 3 · 4 · 5절에서 「가정 구분」이 「회사 가정」인 행 (파일 순서 · 결정적).
    반환 [(절, 코드, 값 요약, 근거)] · 회사마다 다른 행은 파라미터 파일이 정하고 스크립트는 고정 목록을 갖지 않는다."""
    T = {h.split('.')[0].strip(): rows for h, cols, rows in H.md_tables(text) if re.match(r'^\d\.', h)}
    out = []
    for r in T.get('3', []):
        if ASSUME_CO not in (r.get('가정 구분', '') or ''):
            continue
        vals = [f'{k} {r.get(k, "").strip()}' for k, dflt in _DEF3.items() if (r.get(k, '') or '').strip() not in dflt]
        out.append(('3', r.get('RPU 코드', '').strip(), ' · '.join(vals) or '(기본값 그대로)', _cell(r.get('근거 (파일 1 절 · 문장)', ''))))
    for r in T.get('4', []):
        if r.get('코드', '').strip() and ASSUME_CO in (r.get('가정 구분', '') or ''):
            n_ = next((v for k, v in r.items() if k.endswith('인원')), '')
            out.append(('4', r['코드'].strip(), f'사업 이벤트 {r.get("이름", "").strip()} · 인원 {n_.strip()} · 시점 {r.get("시점", "").strip()}', _cell(r.get('근거 (파일 1 절)', ''))))
    for r in T.get('5', []):
        if r.get('RPU 코드', '').strip() and ASSUME_CO in (r.get('가정 구분', '') or ''):
            out.append(('5', r['RPU 코드'].strip(), f'시점 이벤트 {r.get("내용", "").strip()} ({r.get("시점", "").strip()})', _cell(r.get('근거 (파일 1 절)', ''))))
    return out


def _cp2_stop(ws, p):
    """채운 파라미터 파일을 다시 읽어 CP2 멈춤 (회사 가정 행 목록을 스크립트가 적음 · 리드가 다시 고르지 않음)."""
    text = H.read(p)
    a = H.approval(p, 'CP2')
    rows = assume_rows(text)
    T = {h.split('.')[0].strip(): rows_ for h, cols, rows_ in H.md_tables(text) if re.match(r'^\d\.', h)}
    blank_kind = [f'3절 {r.get("RPU 코드")}' for r in T.get('3', [])
                  if any((r.get(k, '') or '').strip() not in d for k, d in _DEF3.items()) and not (r.get('가정 구분', '') or '').strip()]
    blank_kind += [f'{k}절 {r.get(c)}' for k, c in (('4', '코드'), ('5', 'RPU 코드')) for r in T.get(k, [])
                   if (r.get(c, '') or '').strip() and not (r.get('가정 구분', '') or '').strip()]
    a1 = {r.get('항목', ''): r.get('값', '') for r in T.get('1', [])}
    empty1 = [k for k, v in a1.items() if not (v or '').strip()]
    lst = ' · '.join(f'{s}절 {c}' for s, c, _, _ in rows) or '없음'
    tbl = H.md_table(['#', '절', 'RPU 코드', '값 (파라미터 파일 그대로)', '근거'], [[i + 1, s, c, v, b] for i, (s, c, v, b) in enumerate(rows)]) if rows else '(「회사 가정」 행 없음)'
    note = '\n'.join([f'### 「회사 가정」 행 {len(rows)}개 (`{PF}` 3 · 4 · 5절 「가정 구분」 = 회사 가정 · 스크립트가 파일 순서대로 옮김)', '',
                      '- 리드는 이 목록을 다시 고르거나 줄이지 않고 그대로 보여 준다. 사람이 행을 더하거나 값을 바꾸면 파라미터 파일에 적고 `params` 를 다시 돌려 이 목록을 새로 낸다', '',
                      tbl, '',
                      '- 「회사 가정」 행은 행마다 사람에게 묻는다. 결과 파일 ②B · ②C · ②_대사 의 「가정 구분」 열로 그대로 옮겨진다.'])
    if a['ok']:   # 이미 승인됨 → 멈추지 않고 목록만 알림 (지난 CP2 멈춤을 다시 띄우지 않음)
        print(json.dumps(dict(status='완료', note=f'CP2 이미 승인됨 ({a["답"]} · {a["승인자"]}) · 파라미터를 바꾸려면 결정로그 「가정 변경」 뒤 CP2 승인을 다시',
                              assume=[f'{s}절 {c} · {v}' for s, c, v, _ in rows], file=ws.rel(p)), ensure_ascii=False, indent=1))
        return
    ev = [ws.rel(p), f'CP2 답={a["답"] or "(빈칸)"}', f'회사 가정 {len(rows)}행: {lst}']
    if blank_kind:
        ev.append(f'가정 구분 빈칸 {len(blank_kind)}행 (바꾼 행 · 4 · 5절 행): {" · ".join(blank_kind[:12])} → 수요 담당이 채운 뒤 다시')
    if empty1:
        ev.append(f'1절 빈칸: {" · ".join(empty1)}')
    raise H.Stop('CP', 'CP2', '동인 선정 · 사업계획 보정 · 사업 이벤트 · 「회사 가정」 행 확인 (파라미터 파일을 덮지 않고 다시 읽음)', ev,
                 ['확정 (✓ 동인 · 보정 · 가정 구분 그대로 · 아래 「회사 가정」 행 포함)', '확정 · 바꾼 값을 이유 칸에 (RPU · 동인 · 보정 계수 · 회사 가정 값 지정)',
                  '추가 확인 · 보류 · 이관 (근거 문장을 받은 뒤 다시)', '새로 만들기 (`params --new`, 사람이 적은 값이 지워짐)'],
                 note=note)


def params(ws, args):
    H.require_cp(ws, 'CP1', '①_열대응표.md', '② 수요 파라미터 작성')
    p = ws.o(PF)
    if os.path.exists(p) and 'new' not in (args.allow_new or ''):
        return _cp2_stop(ws, p)
    m = H.load_map(ws)
    hcy = _hcy(ws, m)
    pool, blanks = _pool(ws, m)
    by = str(ws.cfg['base_year'])
    if blanks:
        yrs = sorted({y for _, _, ms, _ in blanks for y in ms})
        if not H.allowed(args, '동인빈칸'):
            raise H.Stop('예외', '동인빈칸', f'인력동인 시계열에 빈칸이 있음 ({", ".join(yrs)}) → 동인 비율 · 상관을 어떻게 낼지',
                         [f'{len(blanks)}행 · 예: {blanks[0][0]} {blanks[0][1]} ({blanks[0][3]})', '빈칸은 0으로 읽지 않음 (공통 규칙 4)'],
                         ['자료를 채워 다시 받기 (권장)', f'빈 연도를 빼고 남은 연도로 계산 (`--allow 동인빈칸=제외`)', '방식 B 생략 · A와 C로 대사 (`--allow 동인빈칸=B생략`)'],
                         note='방식 B 숫자를 내지 않았다.')
    skipB = H.allowed(args, '동인빈칸') == 'B생략'
    sy = [str(y) for y in ws.cfg['series_years']]

    rowsB, picks = [], {}
    for code in hcy:
        cands = pool.get(code, [])
        hc = [hcy[code][y] for y in sy]
        for d in cands:
            pairs = [(x, y) for x, y in zip(d['vals'], hc) if x is not None and y is not None]
            d['r'] = H.pearson([a for a, _ in pairs], [b for _, b in pairs]) if len(pairs) >= 3 else None
            d['chk'] = _principle(d)
        ranked = sorted([d for d in cands if d['r'] is not None], key=lambda d: -abs(d['r']))
        for i, d in enumerate(ranked):
            d['rank'] = i + 1
        ok = [d for d in cands if not d['chk'].startswith('위반') and not d['chk'].startswith('참고')]
        pick = (sorted([d for d in ok if d['r'] is not None], key=lambda d: -abs(d['r'])) or ok or cands[:1])
        pick = pick[0] if pick else None
        picks[code] = pick
        flat = all(d['r'] is None for d in cands)
        for d in cands:
            rowsB.append([code, d['typ'], d['sub'], d['name'], d['unit'],
                          '계산 불가' if d['r'] is None else f"{d['r']:.3f}", d.get('rank', '-'),
                          '-' if d['v27'] is None else d['v27'], d['chk'],
                          ('제안' + (' (상관 계산 불가 · 원칙과 파일 순서)' if flat else '')) if d is pick else '',
                          '✓' if d is pick and not skipB else '', '비율' if d is pick else ''])
    # 3. 사업계획 보정 표 (동인 비율은 제안 동인 기준 자동 · 사람이 바꾸면 run 이 다시 계산)
    rowsC = []
    org = {r['RPU 코드']: r for r in H.load(ws, 'org', m)}
    for code in hcy:
        pk = picks.get(code)
        rt = (f"{pk['v27'] / pk['vals'][-1]:.4f}" if pk and pk['v27'] is not None and pk['vals'][-1] else '(직접 RPU 비율)')
        rowsC.append([code, H.norm(org.get(code, {}).get('RPU(산정 단위)', '')), rt, 1.0, 0, '×', '', '', ''])
    tr = {H.norm(r['연도']): r for r in H.load(ws, 'trend', m)}
    t25 = tr.get(by, {})
    hcroi = (t25.get('영업이익', 0) + t25.get('총인건비', 0)) / t25['총인건비'] if t25.get('총인건비') else None
    ty = ws.cfg['target_year']
    text = [f'# ② 수요 파라미터 ({ws.co})', '',
            f'- 만든 스크립트: `{SCRIPT} params` · 수요 담당이 빈칸을 **파일 1(사업계획) 근거 문장과 함께** 채운다',
            '- **CP2 동인 선정**: 2절 「선정」 ✓ · 3절 보정 · 4절 사업 이벤트를 사람이 확인하고 「## 승인」 표를 채우기 전에는 계산하지 않는다',
            '- 사업계획에 없는 절감률 · 생산성 목표를 넣지 않는다. 숫자를 지어내지 않고, 문서에 없으면 빈칸으로 두고 묻는다', '',
            '## 1. 방식 A 입력 (매출·이익 연계)', '',
            H.md_table(['항목', '값', '근거 (파일 · 절)'], [
                [f'{ty} 매출 보수', '', '파일 1 · 1절'], [f'{ty} 매출 기본', '', '파일 1 · 1절'], [f'{ty} 매출 낙관', '', '파일 1 · 1절'],
                [f'{ty} 영업이익', '', '파일 1 · 1절'], ['인건비 인상률 (연)', '', '파일 1 · 6절'],
                ['목표 HCROI', '', f'사람이 정함 · {by} 실적 {hcroi:.2f} (파일 2 · 연도별 추이)' if hcroi else '사람이 정함']]), '',
            '## 2. 방식 B 동인 선정', '',
            '- 원칙 (법전 · 방법론): 동인은 인원과 함께 움직이고 매년 같은 정의로 모을 수 있어야 함 · Value Center 에 매출 같은 회사 전체 실적을 쓰지 않음 · '
            f'{ty} 목표가 없는 동인은 고르지 않음 · 상관이 높아도 원칙에 맞지 않으면 고르지 않음 (AI 순위 ≠ 결정)',
            '- 「선정」 칸 ✓ 를 RPU마다 하나만 둔다. 「방식」은 비율 (3개년 평균) 또는 회귀 (고정 인원이 뚜렷할 때)', '',
            H.md_table(['RPU 코드', '유형', '하위 유형', '동인 후보', '단위', '상관 r', '순위', f'{ty} 계획', '원칙 점검', '스크립트 제안', '선정', '방식'], rowsB), '',
            '## 3. 사업계획 보정 (방식 B · C 공통)', '',
            '- 동인 비율 = 선정 동인 2027 계획 ÷ 2025 값 (run 이 계산) · C 계수 = 동인 비율 × 보정 계수 (C 계수 칸에 숫자를 적으면 그 값을 씀)',
            '- 가산 인원 = 업무량에 없는 신규 업무 (법령 선임 · 신설 업무) · AI 적용 = 파일 1 디지털·AI 전환 절의 적용 범위 RPU 만 ○',
            '- 관리 지원 RPU 의 C 계수는 근거와 함께 직접 적는다 (적지 않으면 직접 RPU 동인 비율을 씀)',
            '- **가정 구분** (3 · 4 · 5절 마지막 칸 · CP2 전에 채움): 「사업계획」 = 파일 1 에 그 숫자 · 문장이 그대로 있음 · 「회사 가정」 = 파일 1 에 숫자가 없어 담당 · 사람이 정한 값 (산식 · 이유를 근거 칸에). '
            '보정 · 가산 · C 계수 · AI 적용을 바꾼 행과 4 · 5절 모든 행에 적는다. 「회사 가정」 행은 CP2 에서 사람이 답한다', '',
            H.md_table(['RPU 코드', 'RPU', '동인 비율', '보정 계수', '가산 인원', 'AI 적용', 'C 계수', '근거 (파일 1 절 · 문장)', '가정 구분'], rowsC), '',
            '## 4. 사업 이벤트 (별도 줄)', '',
            '- 업무량 조사에 아직 없는 신설 조직 (예: 신공장). 결정 수요에 별도 줄로 더하고, 공급에는 넣지 않는다', '',
            '- 신설 조직은 기존 RPU 에 넣지 않고 새 RPU 코드를 붙인다 (①_사업전략분석.md 3절 신설 줄과 같은 코드)',
            '- 신설 조직 하나 = 한 행 · RPU 코드 하나. 사업계획이 직무별 인원을 나눠 적어도 그 RPU 안의 구성이다 (근거 칸에 「구성: …」 · 행을 나누지 않음)',
            '- 행은 필요한 만큼 더한다 (4 · 5절 같음 · 가정 구분 칸 포함)', '',
            H.md_table(['코드', '이름', f'{ty} 인원', '시점', '근거 (파일 1 절)', '가정 구분'], [['', '', '', '', '', '']]), '',
            '## 5. 시점 이벤트 (④ Gap 유형 판정에 씀)', '',
            '- 정해진 시한까지 그 RPU 인력이 갖춰져야 하는 사건만 (법령 시행 · 신설 가동 · 특정 과제 착수) · 계획연도 전체의 증가 (증설 · 내재화 · AI 도입)는 3절에서 다룸 · 한 행에 RPU 코드 하나', '',
            H.md_table(['RPU 코드', '내용', '시점', '근거 (파일 1 절)', '가정 구분'], [['', '', '', '', '']]), '',
            '## 6. 대사 판정 바꾸기 (선택 · 사람이 정할 때만)', '',
            '- 기본은 원인 판정 규칙으로 C 채택. 다른 값을 채택하면 이유를 적는다 (평균은 내지 않는다)', '',
            H.md_table(['RPU 코드', '채택 (B · C · 숫자)', '이유'], [['', '', '']]), '',
            H.approval_block('CP2', '인사기획팀장')]
    H.write_md(p, '\n'.join(text))
    raise H.Stop('CP', 'CP2', '동인 선정 · 사업계획 보정 · 사업 이벤트 확인 (빈칸은 수요 담당이 파일 1 근거로 채운 뒤 보여 줌)',
                 [ws.rel(p), f'RPU {len(hcy)}개 · 동인 후보 {len(rowsB)}개 · 원칙 위반 후보 {sum(1 for r in rowsB if r[8].startswith("위반"))}개'],
                 ['확정 (✓ 동인 · 보정 · 가정 구분 그대로 · 「회사 가정」 행 포함)', '확정 · 바꾼 값을 이유 칸에 (RPU · 동인 · 보정 계수 지정)',
                  '추가 확인 · 보류 · 이관 (근거 문장을 받은 뒤 다시)'],
                 note='「회사 가정」 행 (파일 1 에 숫자가 없는 값)은 행마다 사람에게 묻는다. 결과 파일 ②B · ②C · ②_대사 의 「가정 구분」 열로 그대로 옮겨진다.\n\n'
                      '- 수요 담당이 빈칸 · 「가정 구분」을 채운 뒤 `step2.py params --co ' + ws.co + '` 를 다시 돌리면 파일을 덮지 않고 이 멈춤을 다시 내며, '
                      '「회사 가정」 행 목록을 스크립트가 적는다 (리드는 그 목록을 그대로 보여 줌)')


# ---------------------------------------------------------------------------
def _read_params(ws):
    p = ws.o(PF)
    with open(p, encoding='utf-8') as f:
        T = {h.split('.')[0].strip(): (cols, rows) for h, cols, rows in H.md_tables(f.read()) if re.match(r'^\d\.', h)}
    return T


def _f(s, what):
    s = (s or '').replace(',', '').replace('억', '').replace('원', '').replace('명', '').strip()
    if not s:
        raise H.Stop('예외', '파라미터빈칸', f'② 수요 파라미터 「{what}」 칸이 비어 있음', [PF], ['수요 담당이 파일 1 근거로 채움', '이 방식 생략'])
    if s.endswith('%'):
        return float(s[:-1]) / 100
    m = re.search(r'-?[\d.]+', s)
    if not m:
        raise H.Stop('예외', '파라미터형식', f'「{what}」 값을 숫자로 읽을 수 없음: {s}', [PF], ['숫자로 고쳐 적기'])
    return float(m.group(0))


def run(ws, args):
    H.require_cp(ws, 'CP1', '①_열대응표.md', '② 수요 계산')
    part_a = (getattr(args, 'part', '') or '').upper() == 'A'
    if not part_a:
        H.require_cp(ws, 'CP2', PF, '② 수요 계산')
    elif not os.path.exists(ws.o(PF)):
        raise H.Stop('예외', '파라미터없음', '방식 A 를 먼저 보려면 ②_수요파라미터.md 1절이 있어야 함', [PF], ['`step2.py params` 로 초안 만들기'])
    cfg = ws.cfg
    by, ty = str(cfg['base_year']), cfg['target_year']
    hy = [str(y) for y in cfg['hist_years']]
    m = H.load_map(ws)
    P = _read_params(ws)
    org = {r['RPU 코드']: r for r in H.load(ws, 'org', m)}
    codes = list(org)
    hcn = {c: org[c]['계'] for c in codes}
    inputs = ['①_열대응표.json', PF] + sorted({t['file'] for t in m['tables'].values() if t.get('file')})
    unit = None
    tt = m['tables'].get('trend') or {}
    if tt.get('file'):
        import openpyxl
        unit = H.sheet_unit(openpyxl.load_workbook(os.path.join(ws.data, tt['file']), data_only=True)[tt['sheet']])
    unit_line = (f'단위 (파일 2 연도별 추이 안내문): {unit} · 인원 = 명' if unit else '단위: 파일 2 연도별 추이 안내문에 없음 · CP1 에서 확인한 단위를 따름 · 인원 = 명')
    labor_line = '인건비 = 파일 2 연도별 추이 「총인건비」 열 (포함 항목은 안내문 · 데이터 사전 정의를 따름)'

    # ---------------- A
    A = {r['항목']: r for r in P['1'][1]}
    tr = sorted(H.load(ws, 'trend', m), key=lambda r: int(H.norm(r['연도'])))
    xs = [H.num(r['매출'], '매출', r['_ref']) for r in tr]
    ys = [H.num(r['인원(명)'], '인원', r['_ref']) for r in tr]
    a, b, r2 = H.regress(xs, ys)
    sc = {k: _f(A[f'{ty} 매출 {k}']['값'], f'{ty} 매출 {k}') for k in ('보수', '기본', '낙관')}
    reg = {k: a + b * v for k, v in sc.items()}
    t25 = tr[-1]
    op27 = _f(A[f'{ty} 영업이익']['값'], f'{ty} 영업이익')
    raise_ = _f(A['인건비 인상률 (연)']['값'], '인건비 인상률')
    raise_ = raise_ / 100 if raise_ > 1 else raise_
    tgt = _f(A['목표 HCROI']['값'], '목표 HCROI')
    lab27 = op27 / (tgt - 1)
    per27 = t25['총인건비'] / t25['인원(명)'] * (1 + raise_) ** (ty - int(by))
    hcroi_hc = lab27 / per27
    A_lo, A_hi = min(hcroi_hc, reg['기본']), max(hcroi_hc, reg['기본'])
    rowsA = [['회귀 (인원 = a + b × 매출)', k, v, round(reg[k], 1), f'{H.norm(tr[0]["연도"])}~{by} ({len(xs)}개)', round(r2, 3),
              f'a={a:.1f} · b={b:.4f} · 파일 2 연도별 추이'] for k, v in sc.items()]
    rowsA.append(['HCROI 역산', '목표 HCROI', tgt, round(hcroi_hc, 1), f'{by} 실적 기준', '-',
                  f'Target 인건비 = {op27:g} ÷ ({tgt} − 1) = {lab27:.1f} · {ty} 1인당 인건비 = {per27:.4f} (인상률 {raise_:.1%})'])
    # HR Index (연도별 · 법전 HR Index 4 + 인건비 비중)
    rowsI = []
    for r in tr:
        hc_, rev, op, lab = (H.num(r[k], k, r['_ref']) for k in ('인원(명)', '매출', '영업이익', '총인건비'))
        rowsI.append([H.norm(r['연도']), hc_, rev, op, lab, round(rev / hc_, 2), round(op / hc_, 3), round((op + lab) / hc_, 2),
                      round((op + lab) / lab, 2), round(lab / rev * 100, 1)])
    H.write_xlsx(ws.o('②A_매출이익.xlsx'), [dict(name='매출이익', title=f'②A 매출·이익 연계 ({ws.co}) · 전사 {ty} 적정 인원 범위',
        notes=[unit_line, labor_line, '표본이 5개라 회귀는 방향을 보는 정도 (R²가 높아도 한계 있음)', f'A 범위 = HCROI 역산 ~ 회귀 기본 = {A_lo:.0f} ~ {A_hi:.0f} (명)'] +
              (['CP2 승인 전 방식 A 만 먼저 계산함 (--part A) · B · C · 대사는 CP2 승인 뒤'] if part_a else []),
        head=['길', '시나리오', '입력 값', f'{ty} 인원', '표본', 'R²', '근거'], rows=rowsA),
        dict(name='HR Index', title=f'HR Index ({ws.co}) · 연도별',
        notes=[unit_line + ' · 인당 매출 · 인당 영업이익 · HCVA = 금액 단위 ÷ 명 · HCROI = 배 · 인건비 비중 = %', labor_line,
               '인당 매출 · 인당 영업이익 · HCVA = (영업이익 + 인건비) ÷ 인원 · HCROI = (영업이익 + 인건비) ÷ 인건비 · 인건비 비중 = 인건비 ÷ 매출',
               '지표 하나로 결론 내지 않음 · 비교 회사 값은 사람이 data/ 에 파일로 넣을 때만 씀'],
        head=['연도', '인원', '매출', '영업이익', '인건비', '인당 매출', '인당 영업이익', 'HCVA', 'HCROI', '인건비 비중 %'], rows=rowsI)], SCRIPT, inputs)
    if part_a:
        print(json.dumps(dict(status='완료 (방식 A 만)', A=dict(회귀={k: round(v, 1) for k, v in reg.items()}, R2=round(r2, 3), HCROI역산=round(hcroi_hc, 1),
                              범위=f'{A_lo:.0f}~{A_hi:.0f}'), HR_Index=rowsI[-1], 다음='CP2 승인 뒤 run (B · C · 대사)',
                              files=[ws.rel(ws.o('②A_매출이익.xlsx'))]), ensure_ascii=False, indent=1))
        return

    # ---------------- B
    pool, blanks = _pool(ws, m)
    if blanks and not H.allowed(args, '동인빈칸'):
        raise H.Stop('예외', '동인빈칸', '인력동인 시계열에 빈칸이 있어 방식 B 를 계산할 수 없음',
                     [f'{len(blanks)}행 · 예: {blanks[0][0]} {blanks[0][1]} ({blanks[0][3]})'],
                     ['자료 채우기', '`--allow 동인빈칸=제외`', '`--allow 동인빈칸=B생략`'])
    skipB = H.allowed(args, '동인빈칸') == 'B생략'
    hcy = _hcy(ws, m)
    sel = {}
    for r in P['2'][1]:
        if r.get('선정', '').strip() in ('✓', 'v', 'V', 'O', 'o', '○'):
            if r['RPU 코드'] in sel:
                raise H.Stop('예외', '동인중복', f'{r["RPU 코드"]} 에 선정 ✓ 가 둘 이상', [PF + ' · 2절'], ['하나만 남기기'])
            sel[r['RPU 코드']] = r
    if not skipB:
        miss = [c for c in codes if c not in sel]
        if miss:
            raise H.Stop('예외', '동인미선정', '선정 ✓ 가 없는 RPU 가 있음', [', '.join(miss)], ['동인 선정', '이 RPU 는 방식 B 생략'])
    adj = {r['RPU 코드']: r for r in P['3'][1]}
    kind = lambda c: (adj.get(c, {}).get('가정 구분', '') or '').strip() or '-'   # 파라미터 파일 「가정 구분」 → 결과 파일 열 (스크립트가 옮김 · 손으로 고치지 않음)
    aj = lambda c, k, d: _f(adj[c][k], f'{c} {k}') if c in adj and adj[c].get(k, '').strip() else d
    B, BHOW, ratio27 = {}, {}, {}
    direct = [c for c in codes if c in sel and '관리 지원' not in sel[c]['하위 유형']]
    mgmt = [c for c in codes if c in sel and '관리 지원' in sel[c]['하위 유형']]
    for c in direct:
        d = next(x for x in pool[c] if x['name'] == sel[c]['동인 후보'])
        hc = {y: hcy[c][y] for y in hcy[c]}
        if d['v27'] is None:
            raise H.Stop('예외', '동인계획없음', f'{c} 선정 동인 「{d["name"]}」 에 {ty} 계획 값이 없음', [d['ref']], ['다른 동인 선정', '계획 값을 받아 채움'])
        xv = dict(zip([str(y) for y in cfg['series_years']], d['vals']))
        ratio27[c] = d['v27'] / xv[by]
        if sel[c].get('방식', '').startswith('회귀'):
            ks = [y for y in xv if xv[y] is not None and hc.get(y) is not None]
            a_, b_, r2_ = H.regress([xv[y] for y in ks], [hc[y] for y in ks])
            B[c] = a_ + b_ * d['v27']
            BHOW[c] = f'회귀 a={a_:.2f} b={b_:.3f} R²={r2_:.3f} × {d["v27"]}'
        else:
            rs = [hc[y] / xv[y] for y in hy if xv.get(y)]
            rr = sum(rs) / len(rs)
            k, add = aj(c, '보정 계수', 1.0), aj(c, '가산 인원', 0.0)
            B[c] = rr * d['v27'] * k + add
            BHOW[c] = f'비율 {rr:.5f} ({"·".join(hy)} 평균) × {d["v27"]}' + (f' × 보정 {k}' if k != 1 else '') + (f' + 가산 {add:g}' if add else '')
    b_direct = sum(B[c] for c in direct)
    dir_hc = {y: sum(hcy[c][y] for c in codes if c not in mgmt) for y in hy}
    for c in mgmt:
        rs = [hcy[c][y] / dir_hc[y] for y in hy]
        rr = sum(rs) / len(rs)
        B[c] = rr * b_direct
        BHOW[c] = f'비율 {rr:.5f} (직접 RPU 인원 대비 {"·".join(hy)} 평균) × 직접 RPU B 합 {b_direct:.1f}'
    rowsB = []
    for c in codes:
        if c not in sel:
            rowsB.append([c, '', '(방식 B 생략)', '', '', '', '', '', '', hcn[c], kind(c)])
            continue
        s = sel[c]
        d = next((x for x in pool[c] if x['name'] == s['동인 후보']), None)
        rowsB.append([c, s['유형'], s['동인 후보'], s['단위'], d['vals'][-1] if d else '', s.get(f'{ty} 계획', ''), s['상관 r'] + f" ({s['순위']})",
                      BHOW.get(c, ''), B[c] if c in B else '', hcn[c], kind(c)])
    H.write_xlsx(ws.o('②B_인력동인.xlsx'), [dict(name='인력동인', title=f'②B 조직별 인력동인 ({ws.co}) · RPU별 {ty} 인원',
        notes=[f'직접 RPU B 합 {b_direct:.1f} · 관리 지원 포함 B 합 {sum(B.values()):.1f}', '동인 선정은 CP2 승인 (②_수요파라미터.md 2절)',
               '가정 구분 = ②_수요파라미터.md 3절 「가정 구분」 (사업계획 · 회사 가정) 을 옮김 · 이 파일을 손으로 고치지 않는다 (다시 돌리면 사라짐)'],
        head=['RPU 코드', '유형', '선정 동인', '단위', by, f'{ty} 계획', '상관 (순위)', '방식', f'B {ty}', '현원', '가정 구분'], rows=rowsB)], SCRIPT, inputs)

    # ---------------- C
    asm = H.load(ws, 'assume', m)
    STD = H.assume_value(asm, '표준근무시간', 1800)
    BUF = 1 + H.assume_value(asm, 'Buffer', 0.05)
    CO = H.assume_value(asm, '처리량 계수', 1.4)
    REV = H.assume_value(asm, '검토·관리 비율', 0.15)
    ADOPT = H.assume_value(asm, 'AI 적용률', 0.5)
    work = H.load(ws, 'work', m)
    bad, agg = [], {c: dict(D=0.0, T=0.0, K=0.0) for c in codes}
    JMAP = {'위임': 'D', '함께': 'T', '유지': 'K'}
    for w in work:
        try:
            f_ = H.num(w['연간 빈도(회)'], '연간 빈도', w['_ref'])
            o_ = H.num(w['1회 소요시간(시간)'], '1회 소요시간', w['_ref'])
        except ValueError as e:
            raise H.Stop('예외', '업무량빈칸', str(e), [w['_ref']], ['원천 행 확인 후 채움', '그 업무를 빼고 계산 (이유를 결정로그에)'])
        ann = f_ * o_
        if w.get('연간 소요시간(시간)') is not None and abs(ann - w['연간 소요시간(시간)']) > max(1, 0.01 * ann):
            bad.append(f"{w['_ref']} 빈도×1회={ann:g} · 파일 값={w['연간 소요시간(시간)']}")
        j = JMAP.get(H.norm(w['태스크 판정']))
        if j is None:
            raise H.Stop('예외', '판정값', f'태스크 판정 값을 읽을 수 없음: {w["태스크 판정"]}', [w['_ref']], ['위임 · 함께 · 유지 중 하나로 고침'])
        if w['RPU 코드'] in agg:
            agg[w['RPU 코드']][j] += ann
    if bad and not H.allowed(args, '업무량계산'):
        raise H.Stop('예외', '업무량계산', '업무량 조사의 연간 소요시간이 빈도 × 1회 소요시간과 다름', bad[:5],
                     ['빈도 × 1회 소요시간으로 다시 계산해서 씀 (`--allow 업무량계산=재계산`)', '원천 확인'])
    base = {c: sum(agg[c].values()) / STD * BUF for c in codes}
    lo, hi = cfg['fte_band']
    out_band = [(c, base[c] / hcn[c]) for c in codes if not lo <= base[c] / hcn[c] <= hi]
    if out_band and not H.allowed(args, 'FTE범위'):
        ev = []
        for c, rt in out_band:
            top = sorted([w for w in work if w['RPU 코드'] == c], key=lambda w: -(w['연간 빈도(회)'] or 0) * (w['1회 소요시간(시간)'] or 0))[:2]
            ev.append(f'{c} 2025 FTE {base[c]:.1f} ÷ 현원 {hcn[c]} = {rt:.2f}배 · 큰 업무: ' +
                      ' · '.join(f"{w['업무명']} 빈도 {w['연간 빈도(회)']} × 1회 {w['1회 소요시간(시간)']}시간 ({w['_ref']})" for w in top))
        raise H.Stop('예외', 'FTE범위', f'2025 FTE ÷ 현원이 {lo}~{hi}배 밖인 RPU 가 있음 → 원천 행을 확인할지',
                     ev, ['원천 행을 확인하고 고친 파일로 다시 받기 (권장)', '확인 결과 맞음 · 그대로 진행 (`--allow FTE범위=진행`)', '보류'],
                     note='단위 착오 (분 ↔ 시간) · 중복 기재 · 과대 보고 를 먼저 의심한다.')
    C, CD, CH = {}, {}, {}
    for c in codes:
        hD, hT, hK = agg[c]['D'], agg[c]['T'], agg[c]['K']
        k = aj(c, '보정 계수', 1.0)
        mm = adj.get(c, {}).get('C 계수', '').strip()
        if mm and mm != '(자동)':
            mult = _f(mm, f'{c} C 계수')
        elif c in ratio27:
            mult = ratio27[c] * k
        elif c in mgmt and direct:
            mult = b_direct / sum(hcy[x][by] for x in direct)
        else:
            mult = 1.0
        add = aj(c, '가산 인원', 0.0)
        ai = adj.get(c, {}).get('AI 적용', '').strip() in ('○', 'O', 'o', '예', 'Y')
        if ai:
            aD, aT = hD * ADOPT, hT * ADOPT
            human = (hK + (hT - aT) + (hD - aD) + aT / CO + REV * aD) / STD * BUF * mult
            dig = aD / STD * mult
        else:
            human, dig = base[c] * mult, 0.0
        C[c] = human + add
        CD[c] = dig
        CH[c] = dict(mult=mult, add=add, ai=ai, nonai=base[c] * mult + add)
    rowsC = [[c, hcn[c], base[c], base[c] / hcn[c], CH[c]['mult'], CH[c]['add'] or '',
              '○' if CH[c]['ai'] else '', CH[c]['nonai'], C[c], CD[c], kind(c)] for c in codes]
    tot = lambda d: sum(d.values())
    rowsC.append(['합계', sum(hcn.values()), round(tot(base), 1), round(tot(base) / sum(hcn.values()), 2), '', '', '',
                  round(sum(CH[c]['nonai'] for c in codes), 1), round(tot(C), 1), round(tot(CD), 1), ''])
    H.write_xlsx(ws.o('②C_업무량.xlsx'), [dict(name='업무량', title=f'②C 업무량 분석 ({ws.co}) · 사람 FTE + 디지털 FTE',
        notes=[f'2025 FTE = 연간 소요시간 합 ÷ {STD:g} × {BUF:g} (파일 4 산정 가정)',
               f'{ty} = 2025 FTE × C 계수 + 가산 · AI 적용 RPU: 위임·함께의 {ADOPT:.0%} 적용 (함께 ÷ {CO} · 위임 = 디지털 FTE + 검토 {REV:.0%} 사람)',
               '디지털 FTE 는 사람 FTE 에 더하지 않는다 (두 열) · 가정 구분 = ②_수요파라미터.md 3절 (스크립트가 옮김 · 손으로 고치지 않음)'],
        head=['RPU 코드', '현원', '2025 FTE', 'FTE ÷ 현원', 'C 계수', '가산', 'AI 적용', f'{ty} AI 전', f'{ty} 사람', f'{ty} 디지털', '가정 구분'], rows=rowsC)], SCRIPT, inputs)

    # ---------------- 대사
    b1, b2 = cfg['recon_bands']
    over = {r['RPU 코드']: r for r in P['6'][1] if r.get('RPU 코드', '').strip()}
    ratio_c = {c: base[c] / hcn[c] for c in codes}
    max_c = max(ratio_c, key=ratio_c.get)
    chk = {}
    for r in P['2'][1]:
        if r.get('선정', '').strip() in ('✓', 'v', 'V', 'O', 'o', '○'):
            chk[r['RPU 코드']] = r.get('원칙 점검', '')
    rowsR, DEC = [], {}
    for c in codes:
        if c not in B:
            DEC[c] = C[c]
            rowsR.append([c, '', C[c], '', '', 'B 없음', 'C 채택', C[c], CD[c], '방식 B 생략', kind(c)])
            continue
        diff = (B[c] - C[c]) / C[c]
        band = '일치 (3% 이하)' if abs(diff) <= b1 else '설명 (3~7%)' if abs(diff) <= b2 else '재검토 (7% 초과)'
        if chk.get(c, '').startswith('위반') and abs(diff) > b1:
            cause, why, verdict = '동인 선택', f'선정 동인이 원칙 점검 「{chk[c]}」 · 상관만 보고 고른 동인일 수 있음', '재검토 · 잠정 C'
        elif abs(diff) <= b1:
            cause, why, verdict = '해당 없음', '두 방식이 같은 범위', 'C 채택'
        elif abs(diff) > b2 and C[c] > B[c] and c == max_c:
            cause, why, verdict = '업무량 과대 보고', f'C 가 B 보다 크고 2025 FTE ÷ 현원이 전 RPU 중 가장 높음 ({ratio_c[c]:.2f}배) · 업무량 조사 원천 확인 필요', '추가 확인 · 잠정 C'
        elif CH[c]['ai'] and B[c] > C[c]:
            cause, why, verdict = '생산성 가정', f'C 는 {ty} AI 적용을 반영하고 B 는 과거 인원 비율을 그대로 씀', 'C 채택'
        elif C[c] > hcn[c] and B[c] < C[c]:
            cause, why, verdict = '생산성 가정', 'B 는 2025 현원을 적정으로 보고 늘려서 지금 부족분이 빠짐 (C 가 현원보다 큼)', 'C 채택'
        else:
            cause, why, verdict = '생산성 가정', 'B 는 과거 3개년 평균 인당 처리량을 그대로 씀 · 차이가 설명 가능한 범위', 'C 채택'
        dec = C[c]
        if c in over:
            v = over[c]['채택 (B · C · 숫자)'].strip()
            dec = B[c] if v.upper().startswith('B') else C[c] if v.upper().startswith('C') else _f(v, f'{c} 대사 채택')
            verdict = f'사람 판정: {v}'
            why += f' · 사람 이유: {over[c]["이유"]}'
        DEC[c] = dec
        rowsR.append([c, B[c], C[c], B[c] - C[c], f'{diff:+.1%}', band, cause, verdict,
                      dec, CD[c], why, kind(c)])
    ev = [r for r in P['4'][1] if r.get('코드', '').strip()]
    ev_sum = 0.0
    for r in ev:
        n = _f(r[f'{ty} 인원'], f'사업 이벤트 {r["코드"]} 인원')
        ev_sum += n
        rowsR.append([r['코드'], '', '', '', '', '사업 이벤트', '별도 줄', '사업 이벤트', n, 0, f'{r["이름"]} · {r["시점"]} · {r["근거 (파일 1 절)"]}',
                      (r.get('가정 구분', '') or '').strip() or '-'])
    dec_h = sum(DEC.values()) + ev_sum
    dec_d = sum(CD.values())
    where = '범위 안' if A_lo <= sum(DEC.values()) <= A_hi else '범위 밖'
    rowsR.append(['합계', round(sum(B.values()), 1), round(tot(C), 1), '', '', '', '', '', round(dec_h, 1), round(dec_d, 1),
                  f'RPU 합 {sum(DEC.values()):.1f} (A 범위 {A_lo:.0f}~{A_hi:.0f} {where}) + 사업 이벤트 {ev_sum:g}', ''])
    H.write_xlsx(ws.o('②_대사.xlsx'), [dict(name='대사', title=f'② 세 값 대사 · 결정 수요 ({ws.co} · {ty})',
        notes=[f'대사 규칙: 차이 {b1:.0%} 이하 일치 · {b1:.0%}~{b2:.0%} 설명 · {b2:.0%} 초과 재검토 · 평균 내지 않고 원인으로 채택',
               f'결정 수요 사람 = {dec_h:.1f} · 디지털 = {dec_d:.1f} (두 열 · 합산 금지)', f'A 범위 {A_lo:.0f}~{A_hi:.0f} (HCROI 역산 ~ 회귀 기본)',
               '결정 수요 사람 · 디지털 열 이름은 wfp-step4-gap 입력 계약과 같아야 한다',
               '판정 「추가 확인 · 잠정 C」 = 설계값 (C 로 잠정 채택 · ④ 대안표에서 그 RPU 를 「추가 확인」으로 다룸 · CP2 를 다시 열지 않음) · 가정 구분 = ②_수요파라미터.md 3 · 4절을 옮김'],
        head=['RPU 코드', 'B', 'C', 'B − C', '차이', '구간', '원인 판정', '판정', '결정 수요 사람', '결정 수요 디지털', '근거 한 줄', '가정 구분'], rows=rowsR)],
        SCRIPT, inputs)
    print(json.dumps(dict(status='완료', A=dict(회귀={k: round(v, 1) for k, v in reg.items()}, R2=round(r2, 3), HCROI역산=round(hcroi_hc, 1)),
                          B합=round(sum(B.values()), 1), B직접=round(b_direct, 1),
                          C=dict(FTE2025=round(tot(base), 1), AI전=round(sum(CH[c]['nonai'] for c in codes), 1), 사람=round(tot(C), 1), 디지털=round(tot(CD), 1)),
                          결정수요=dict(사람=round(dec_h, 1), 디지털=round(dec_d, 1)),
                          files=[ws.rel(ws.o(f)) for f in ('②A_매출이익.xlsx', '②B_인력동인.xlsx', '②C_업무량.xlsx', '②_대사.xlsx')]),
                     ensure_ascii=False, indent=1))


def main():
    ap = H.common_args('② 인력수요 예측 스크립트')
    ap.add_argument('cmd', choices=['params', 'run'])
    ap.add_argument('--new', dest='allow_new', action='store_const', const='new', default='')
    ap.add_argument('--part', default='', help='A = CP2 전에 방식 A (매출·이익 · HR Index) 만 계산')
    return ap.parse_args()


if __name__ == '__main__':
    args = main()
    H.run_main({'params': params, 'run': run}[args.cmd], lambda: args)
