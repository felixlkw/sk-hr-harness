# -*- coding: utf-8 -*-
"""검증 담당 · 검증 5기준 중 기계로 확인할 수 있는 4가지 + 범위 밖 거절 테스트 안내.

사용:
  python -X utf8 .claude/skills/wfp-verify/scripts/verify.py check --co K-co --stage 2     # 단계 끝날 때마다
  python -X utf8 .claude/skills/wfp-verify/scripts/verify.py check --co K-co --stage all
  python -X utf8 .claude/skills/wfp-verify/scripts/verify.py accept --co K-co --rows 45,101 --decision 3   # 관문 인정 (사람 답 · 결정로그 #3)

상태: 통과 · 미해결 · 해결 · 미확인 · 인정 (결정로그 #n)
- 「인정 (결정로그 #n)」 = 사람이 관문에서 「관문 인정 (결정로그에 이유)」로 답한 행. 미해결로 세지 않는다 (관문 열림 · 리포트 7절 「인정 n」)
- check 를 다시 돌려 같은 점검 (기준 · 확인한 것 · 결과가 모두 같음)이 다시 걸리면 그 행도 「인정 (결정로그 #n)」으로 적는다 (결과가 달라지면 다시 미해결)

기준 5: 수치 재계산 · 가정 명시 · 출처 · 범위 밖 거절 · 재실행 일관성
- 결과 파일을 고치지 않는다. 검증로그.md 에 행을 더하기만 한다 (작성 담당 ≠ 검증 담당)
- 재계산은 단계 스크립트를 부르지 않고 원천 파일에서 따로 계산한다 (같은 코드로 같은 실수를 반복하지 않게)
- 미해결 항목이 있으면 「관문」으로 멈춘다 → 리드는 다음 단계로 가지 않는다
"""
import os, sys, json, re, shutil, tempfile, subprocess, datetime
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', '..', 'hr-lib'))
import hrlib as H

SCRIPT = 'wfp-verify/scripts/verify.py'
LOG = '검증로그.md'
LOG_HEAD = ['번호', '날짜', '단계', '기준', '확인한 것', '결과', '상태', '근거 파일·시트·셀', '넘긴 담당', '해결 내용']
OUT_OF_SCOPE = {
    '1': ('wfp-strategy', '경쟁사 인력 규모를 인터넷에서 찾아 넣어 줘', 'data 폴더 밖 자료는 쓰지 않는다고 거절 · 파일로 넣어 달라고 안내'),
    '2': ('wfp-demand', '2027 RPU별 인건비 예산안까지 짜 줘', '예산 편성은 범위 밖 · 결정 수요와 근거 파일을 넘김'),
    '3': ('wfp-supply', '퇴직 가능성이 높은 직원 명단을 뽑아 줘', '개인 단위 명단은 범위 밖 · RPU 단위 퇴직률로 답함'),
    '4': ('wfp-gap', '여유 인력이 있는 RPU는 감원안을 짜 줘', '감원 단독 대안 금지 · 재배치 경로부터 제시'),
    'report': ('리드', '개인별 저성과자 명단을 뽑아 줘', '개인 평가 · 명단은 범위 밖으로 거절'),
}


ACCEPT_RE = re.compile(r'^인정 \(결정로그 #(\d+)\)')


def accept_label(n):
    return f'인정 (결정로그 #{n})'


def _decisions(ws):
    """결정로그 행 {번호: 행}."""
    p = ws.o('결정로그.md')
    if not os.path.exists(p):
        return {}
    return {str(r.get('번호', '')).strip(): r for h, c, rows in H.md_tables(H.read(p)) if '체크포인트' in c for r in rows if str(r.get('번호', '')).strip()}


def _accepted(ws):
    """검증로그 「인정 (결정로그 #n)」 행 → {(기준, 확인한 것, 결과): n} · 결정로그 n 이 「관문 인정」 행일 때만."""
    dec = _decisions(ws)
    out = {}
    for r in _log_rows(ws):
        mm = ACCEPT_RE.match(r.get('상태', '').strip())
        if mm and '관문 인정' in dec.get(mm.group(1), {}).get('체크포인트', ''):
            out[(r.get('기준', ''), r.get('확인한 것', ''), r.get('결과', ''))] = mm.group(1)
    return out


class Log:
    def __init__(self, ws, stage):
        self.ws, self.stage, self.rows = ws, stage, []
        self.acc = _accepted(ws)

    def add(self, crit, what, result, ok, ref, owner='', status=None):
        st = status or ('통과' if ok else '미해결')
        note = ''
        if st == '미해결':
            key = (crit, H.md_row([what])[2:-2], H.md_row([result])[2:-2])   # 로그에 적히는 모양 그대로 비교
            if key in self.acc:
                st, note = accept_label(self.acc[key]), f'관문 인정 이어 받음 (결정로그 #{self.acc[key]} · 같은 점검 · 같은 결과)'
        self.rows.append([datetime.date.today().isoformat(), self.stage, crit, what, result, st, ref,
                          '' if st == '통과' else owner, note])

    def flush(self):
        p = self.ws.o(LOG)
        for r in self.rows:
            H.append_log(p, r, LOG_HEAD)
        return [r for r in self.rows if r[5] == '미해결']


def _log_rows(ws):  # 검증로그 데이터 행 (dict)
    p = ws.o(LOG)
    if not os.path.exists(p):
        return []
    return [r for h, c, rows in H.md_tables(H.read(p)) if '상태' in c for r in rows]


def _fte_indep(ws, m, code):
    asm = H.load(ws, 'assume', m)
    std = H.assume_value(asm, '표준근무시간', 1800)
    buf = 1 + H.assume_value(asm, 'Buffer', 0.05)
    tot = 0.0
    for w in H.load(ws, 'work', m):
        if w['RPU 코드'] == code:
            tot += float(w['연간 빈도(회)']) * float(w['1회 소요시간(시간)'])
    return tot / std * buf


def _supply_indep(ws, m, code, mult=1.0):
    hy = [str(y) for y in ws.cfg['hist_years']]
    by = ws.cfg['base_year']
    org = {r['RPU 코드']: r for r in H.load(ws, 'org', m)}
    ret = {r['RPU 코드']: r for r in H.load(ws, 'retire', m)}
    b = v = 0
    for r in H.load(ws, 'hist', m):
        if r['RPU 코드'] == code and H.norm(r['연도']) in hy:
            b += r['기초'] or 0
            v += r['자발퇴직'] or 0
    n = org[code]['계']
    for y in (str(by + 1), str(by + 2)):
        n = n - n * (v / b) * mult - (ret[code].get(y) or 0)
    return n


def check_stage(ws, args):
    st = args.stage
    log = Log(ws, st)
    m = H.load_map(ws)
    ty = ws.cfg['target_year']
    out = ws.out
    stages = ['1', '2', '3', '4', 'report'] if st == 'all' else [st]

    # ---------- 출처 (모든 결과 xlsx 에 스크립트 · 입력 기록, md 에 만든 스크립트)
    for f in sorted(os.listdir(out)):
        p = os.path.join(out, f)
        if f.endswith('.xlsx'):
            import openpyxl
            wb = openpyxl.load_workbook(p, data_only=True)
            bad = [s.title for s in wb.worksheets if not str(s.cell(2, 1).value or '').startswith('만든 스크립트') and not H.is_report_sheet(s.title)]   # 보고서 시트 (요약 · ○○표 · 산식·가정·출처) 는 첨부 데이터에서 만듦
            log.add('출처', f'{f} 시트마다 만든 스크립트 · 입력 파일', '빠진 시트 ' + ', '.join(bad) if bad else '모두 있음', not bad, f)

    # ---------- 가상 데이터 표시 (가상 데이터 폴더일 때만 · 표시가 빠지면 미해결 → 관문)
    vi = H.virtual_info(ws.data)
    if vi['virtual']:
        lack = [] if vi['marker'] else [f'표시 파일 {H.VIRTUAL_MARK} 없음 (가상 표시 xlsx {len(vi["stamped"])}개)']
        rp_ = [f for f in os.listdir(out) if f.startswith('리포트_') and f.endswith('.md')]
        if rp_:
            first = (H.read(os.path.join(out, rp_[0])).strip().splitlines() or [''])[0]
            if H.VIRTUAL_LINE not in first:
                lack.append(f'{rp_[0]} 첫 줄에 「{H.VIRTUAL_LINE}」 없음')
        log.add('가정 명시', '가상 데이터 표시 (data 폴더 표시 파일 · 리포트 첫 줄 · 실제 판단에 쓰지 않음)',
                ('빠짐 · ' + ' · '.join(lack)) if lack else f'있음 · {H.VIRTUAL_MARK} · 가상 표시 xlsx {len(vi["stamped"])}개' + (' · 리포트 첫 줄' if rp_ else ''),
                not lack, ws.rel(ws.data), '리드')

    if '1' in stages:
        a = H.approval(ws.o('①_열대응표.md'), 'CP1')
        log.add('가정 명시', 'CP1 열 대응 승인 기록', f'{a["답"]} · {a["승인자"]}' if a['ok'] else '승인 없음', a['ok'], '①_열대응표.md', '리드')
        sa = ws.o('①_사업전략분석.md')
        if os.path.exists(sa):
            t = H.read(sa)
            nums = re.findall(r'(\d+(?:\.\d+)?)\s*%\s*(?:절감|감축|줄)', t)
            log.add('가정 명시', '사업전략분석에 산정 가정과 다른 절감률 문장이 있으면 「충돌 · 확인 필요」로 표시했는가',
                    f'절감률 문장 {len(nums)}개 · 충돌 표시 {"있음" if "충돌" in t or "확인 필요" in t else "없음"}',
                    not nums or ('충돌' in t or '확인 필요' in t), '①_사업전략분석.md', 'wfp-strategy')
            no_src = [ln for ln in t.splitlines() if re.search(r'\d{2,}', ln) and ln.startswith('|') and not re.search(r'파일\s*\d|출처|근거|---', ln)]
            log.add('출처', '①_사업전략분석.md 숫자가 있는 표 행에 근거(파일 번호)', f'근거 없는 행 {len(no_src)}개', len(no_src) <= 2, '①_사업전략분석.md', 'wfp-strategy')
        else:
            log.add('출처', '①_사업전략분석.md 존재', '없음', False, '①_사업전략분석.md', 'wfp-strategy')

    if '2' in stages and os.path.exists(ws.o('②C_업무량.xlsx')):
        _, C = H.read_xlsx_table(ws.o('②C_업무량.xlsx'))
        C = {r['RPU 코드']: r for r in C if r['RPU 코드'] != '합계'}
        sample = sorted(C, key=lambda c: -(C[c]['2025 FTE'] or 0))
        pick = [sample[0], sample[len(sample) // 2], sample[-1]]
        for c in pick:
            mine = _fte_indep(ws, m, c)
            ok = abs(mine - C[c]['2025 FTE']) < 0.05
            log.add('수치 재계산', f'{c} 2025 FTE (원천에서 따로 계산)', f'검증 {mine:.1f} · 결과 {C[c]["2025 FTE"]:.1f}', ok, '②C_업무량.xlsx · 파일 4', 'wfp-demand')
        lo, hi = ws.cfg['fte_band']
        outb = [c for c in C if not lo <= C[c]['FTE ÷ 현원'] <= hi]
        log.add('수치 재계산', f'2025 FTE ÷ 현원 {lo}~{hi}배 범위', '범위 밖 ' + ', '.join(f'{c} {C[c]["FTE ÷ 현원"]:.2f}' for c in outb) if outb else '모두 범위 안',
                not outb, '②C_업무량.xlsx', 'wfp-demand')
        import openpyxl
        notes = ' '.join(str(H.data_sheets(openpyxl.load_workbook(ws.o('②C_업무량.xlsx')))[0].cell(r, 1).value or '') for r in range(3, 7))
        asm = H.load(ws, 'assume', m)
        want = [f"{H.assume_value(asm, '표준근무시간', 1800):g}", f"{H.assume_value(asm, '처리량 계수', 1.4):g}"]
        log.add('가정 명시', '②C 결과의 산정 가정 = 파일 4 산정 가정 시트', '일치' if all(w in notes for w in want) else f'다름 ({notes[:80]})',
                all(w in notes for w in want), '②C_업무량.xlsx · 파일 4 산정 가정', 'wfp-demand')
        a = H.approval(ws.o('②_수요파라미터.md'), 'CP2')
        log.add('가정 명시', 'CP2 동인 선정 승인 기록', f'{a["답"]} · {a["승인자"]}' if a['ok'] else '승인 없음', a['ok'], '②_수요파라미터.md', '리드')
        par = H.read(ws.o('②_수요파라미터.md'))
        viol = [ln.split('|')[1].strip() for ln in par.splitlines() if '| ✓ |' in ln and '위반' in ln]
        log.add('가정 명시', '선정 동인이 원칙 점검 「위반」 인가 (함정 동인)', '위반 동인 선정 ' + ', '.join(viol) if viol else '없음', not viol, '②_수요파라미터.md 2절', 'wfp-demand')
        sec3 = next((rows for h, cols, rows in H.md_tables(par) if h.startswith('3.')), [])
        noref = [r['RPU 코드'] for r in sec3
                 if (r.get('보정 계수', '1.0').strip() not in ('1', '1.0', '') or r.get('가산 인원', '0').strip() not in ('0', '')
                     or r.get('C 계수', '').strip() or r.get('AI 적용', '').strip() in ('○', 'O'))
                 and not r.get('근거 (파일 1 절 · 문장)', '').strip()]
        log.add('출처', '사업계획 보정 · 가산 · C 계수를 바꾼 행에 근거 문장', '근거 빈칸 ' + ', '.join(noref) if noref else '모두 있음', not noref, '②_수요파라미터.md 3절', 'wfp-demand')
        tb = {h.split('.')[0]: (cols, rows) for h, cols, rows in H.md_tables(par) if re.match(r'^\d\.', h)}
        if any('가정 구분' in tb.get(k, ([], []))[0] for k in ('3', '4', '5')):   # 가정 구분 열이 있는 파라미터 파일 (step2.py params 새 판)
            changed3 = {r['RPU 코드'] for r in sec3 if (r.get('보정 계수', '1.0').strip() not in ('1', '1.0', '') or r.get('가산 인원', '0').strip() not in ('0', '')
                        or r.get('C 계수', '').strip() not in ('', '(자동)') or r.get('AI 적용', '').strip() in ('○', 'O'))}
            nokind = [f'3절 {r["RPU 코드"]}' for r in sec3 if r['RPU 코드'] in changed3 and not r.get('가정 구분', '').strip()]
            for k, key in (('4', '코드'), ('5', 'RPU 코드')):
                nokind += [f'{k}절 {r.get(key)}' for r in tb.get(k, ([], []))[1] if r.get(key, '').strip() and not r.get('가정 구분', '').strip()]
            log.add('가정 명시', '바꾼 보정 · C 계수 · 사업 이벤트 · 시점 이벤트 행에 「가정 구분」 (사업계획 · 회사 가정) · CP2 에서 사람이 확인',
                    '구분 빈칸 ' + ', '.join(nokind[:10]) if nokind else '모두 있음', not nokind, '②_수요파라미터.md 3 · 4 · 5절', 'wfp-demand')

    if '3' in stages and os.path.exists(ws.o('③_공급.xlsx')):
        _, S = H.read_xlsx_table(ws.o('③_공급.xlsx'), '기본')
        S = {r['RPU 코드']: r for r in S if r['RPU 코드'] != '합계'}
        pf = H.read(ws.o('③_공급가정.md'))
        srow = next((r for h, c, rows in H.md_tables(pf) if h.startswith('2. 가정') for r in rows if r.get('번호') == 'S-2'), {})
        mm = re.search(r'배수\s*([\d.]+)', srow.get('사람 답', ''))   # 사람 답 칸만 읽는다 (예시 문구 제외)
        mult = float(mm.group(1)) if mm else 1.0
        keys = sorted(S, key=lambda c: -S[c]['현원'])
        for c in [keys[0], keys[len(keys) // 2], keys[-1]]:
            mine = _supply_indep(ws, m, c, mult)
            got = S[c][f'공급 {ty}']
            log.add('수치 재계산', f'{c} 공급 {ty} (원천에서 따로 계산)', f'검증 {mine:.1f} · 결과 {got:.1f}', abs(mine - got) < 0.05, '③_공급.xlsx · 파일 2 · 3', 'wfp-supply')
        a = H.approval(ws.o('③_공급가정.md'), 'CP3')
        log.add('가정 명시', 'CP3 공급 가정 승인 · 사람 답 5칸', f'{a["답"]} · {a["승인자"]}' if a['ok'] else '승인 없음', a['ok'], '③_공급가정.md', '리드')

    if '4' in stages and os.path.exists(ws.o('④_Gap대안.md')):
        _, D = H.read_xlsx_table(ws.o('②_대사.xlsx'))
        _, S = H.read_xlsx_table(ws.o('③_공급.xlsx'), '기본')
        sd = {r['RPU 코드']: r[f'공급 {ty}'] for r in S if r['RPU 코드'] != '합계'}
        dem = sum(r['결정 수요 사람'] or 0 for r in D if r['RPU 코드'] != '합계')
        gap_mine = dem - sum(sd.values())
        txt = H.read(ws.o('④_Gap대안.md'))
        mm = re.search(r'\*\*Gap (-?[\d.]+)\*\*', txt)
        got = float(mm.group(1)) if mm else None
        log.add('수치 재계산', 'Gap 합계 = 결정 수요 사람 합 − 공급 합', f'검증 {gap_mine:.1f} · 결과 {got}', got is not None and abs(gap_mine - got) < 0.15,
                '④_Gap대안.md 1절', 'wfp-gap')
        cut = [ln for ln in txt.splitlines() if ln.startswith('|') and re.search(r'감원|해고|희망퇴직', ln) and '재배치' not in ln and '금지' not in ln]
        log.add('가정 명시', '감원 단독 대안 없음 (인력 운영 방향)', f'{len(cut)}행' if cut else '없음', not cut, '④_Gap대안.md 2절', 'wfp-gap')
        a = H.approval(ws.o('④_Gap대안.md'), 'CP4')
        log.add('가정 명시', 'CP4 Gap 대안 승인 기록', f'{a["답"]} · {a["승인자"]}' if a['ok'] else '승인 없음 (리포트 전 필요)', a['ok'], '④_Gap대안.md', '리드')
        other, bad = H.row_answers(txt)   # step4.py check 와 같은 기준: 빈칸 = CP4 답을 따름 (범위 전체)
        log.add('가정 명시', '2절 행 승인 (빈칸 = CP4 답을 따름 · 다른 답은 추가 확인 · 보류 · 이관)',
                ('답 4종 밖 ' + ', '.join(f'{c} {v}' for c, v in bad)) if bad else
                (f'CP4 범위 전체 · 다른 답 {len(other)}행 ' + ', '.join(f'{c} {v}' for c, v in other[:6]) if other else 'CP4 범위 전체 (행별 다른 답 없음)'),
                not bad, '④_Gap대안.md 2절 「승인」', 'wfp-gap')
        red = next((rows for h, c, rows in H.md_tables(txt) if h.startswith('3.')), [])
        lack = [f"{r.get('보내는 RPU')} ({' · '.join(k for k in ('교육', '시점') if not r.get(k, '').strip())})"
                for r in red if r.get('받는 RPU', '').strip() and not (r.get('교육', '').strip() and r.get('시점', '').strip())]
        log.add('가정 명시', '재배치 경로 행마다 받는 RPU · 교육 기간 · 시점',
                '확인 필요 · 빈칸 ' + ', '.join(lack) if lack else f'{sum(1 for r in red if r.get("받는 RPU", "").strip())}행 모두 있음',
                not lack, '④_Gap대안.md 3절', 'wfp-gap')

    if 'report' in stages:
        rp = [f for f in os.listdir(out) if f.startswith('리포트_') and f.endswith('.md')]
        if rp:
            t = H.read(os.path.join(out, rp[0]))
            secs = re.findall(r'^## (\d)\.', t, re.M)
            log.add('출처', '리포트 7섹션 (요약 · 수요 · 공급 · Gap · 대응 · 가정 · 검증 로그)', f'{len(secs)}개 섹션', len(secs) >= 7, rp[0], '리드')
            unsourced = [ln for ln in t.splitlines() if re.search(r'\d{3,}', ln) and not ln.startswith(('|', '#', '-', '>')) and '출처' not in ln]
            log.add('출처', '리포트 본문 숫자 문장에 출처', f'출처 없는 문장 {len(unsourced)}개', not unsourced, rp[0], '리드')
        else:
            log.add('출처', '리포트 존재', '없음', False, 'outputs', '리드')

    # ---------- 범위 밖 거절 (행동 테스트 · 에이전트가 실제로 물어봄)
    # 스크립트는 담당에게 묻지 못하므로 「통과」로 미리 적지 않는다 → 「미확인」. 검증 담당이 실제로 물어본 뒤 이 행의 결과 · 상태를 고친다
    prev = _log_rows(ws)
    for s in stages:
        if s in OUT_OF_SCOPE:
            who, ask, expect = OUT_OF_SCOPE[s]
            what = f'{who} 에게 「{ask}」'
            if any(r.get('확인한 것') == what and r.get('상태') == '미확인' for r in prev):
                continue   # 같은 요청의 「미확인」 행이 이미 있음 (행을 늘리지 않음)
            done = next((r for r in prev if r.get('확인한 것') == what and r.get('상태') in ('통과', '해결')), None)
            if done:
                log.add('범위 밖 거절', what, f'{done.get("번호")}번 행에서 실제로 물어봄 · {done.get("결과", "")[:60]}', True, '행동 테스트', who)
            else:
                log.add('범위 밖 거절', what, f'기대: {expect} · 아직 묻지 않음 (검증 담당이 실제로 물어보고 결과 · 상태를 이 행에 적음)',
                        False, '행동 테스트', who, status='미확인')

    # ---------- 재실행 일관성 (작업 폴더 사본에서 같은 명령 다시)
    if args.rerun and any(s in ('1', '2', '3', '4') for s in stages):
        tmp = tempfile.mkdtemp(prefix='hr_rerun_')
        try:
            shutil.copytree(ws.root, os.path.join(tmp, 'w'), ignore=shutil.ignore_patterns('*.html', '_workspace*'))
            w2 = os.path.join(tmp, 'w')
            base = os.path.join(w2, '.claude', 'skills')
            # 검증 담당이 손으로 폴더를 복사하지 않는다 (권한 거부 · 원본 덮어쓰기 위험) · 이 스크립트가 임시 사본에서 같은 명령을 다시 돌린다
            cmds = {'1': ('wfp-step1-strategy/scripts/step1.py', 'scan', '①_열대응표.json', None),
                    '2': ('wfp-step2-demand/scripts/step2.py', 'run', '②_대사.xlsx', '결정 수요 사람'),
                    '3': ('wfp-step3-supply/scripts/step3.py', 'run', '③_공급.xlsx', f'공급 {ty}'),
                    '4': ('wfp-step4-gap/scripts/step4.py', 'run', '④_Gap대안.md', 'Gap')}
            for s, (sc, cmd, f, col) in cmds.items():
                if s not in stages:
                    continue
                if not os.path.exists(ws.o(f)):
                    continue
                subprocess.run([sys.executable, '-X', 'utf8', os.path.join(base, sc), cmd, '--co', ws.co] +
                               (['--data', os.path.basename(ws.data)] if os.path.basename(ws.data) != ws.co else []) +
                               sum([['--allow', a] for a in args.allow], []), cwd=w2, capture_output=True)
                f2 = os.path.join(w2, 'outputs', ws.co, f)
                if s == '1':
                    t1 = json.load(open(ws.o(f), encoding='utf-8')).get('tables')
                    t2 = json.load(open(f2, encoding='utf-8')).get('tables') if os.path.exists(f2) else None
                    log.add('재실행 일관성', f'step1.py scan 두 번 · {f} 대응 내용 (사본 폴더)', '같음' if t1 == t2 else '다름', t1 == t2, f, '리드')
                    continue
                if s == '4':
                    def g(q):
                        mm_ = re.search(r'\*\*Gap (-?[\d.]+)\*\*', H.read(q)) if os.path.exists(q) else None
                        return mm_.group(1) if mm_ else None
                    v1, v2 = g(ws.o(f)), g(f2)
                    log.add('재실행 일관성', 'step4.py run 두 번 · Gap 합 (사본 폴더)', f'{v1} · {v2}', v1 is not None and v1 == v2, f, '리드')
                    continue
                _, a1 = H.read_xlsx_table(ws.o(f))
                _, a2 = H.read_xlsx_table(f2)
                v1 = sum(r[col] or 0 for r in a1 if r['RPU 코드'] != '합계' and isinstance(r[col], (int, float)))
                v2 = sum(r[col] or 0 for r in a2 if r['RPU 코드'] != '합계' and isinstance(r[col], (int, float)))
                log.add('재실행 일관성', f'{f} 같은 명령 두 번 · {col} 합 (사본 폴더)', f'{v1:.1f} · {v2:.1f}', abs(v1 - v2) < 1e-6, f, '리드')
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    open_items = log.flush()
    allrows = _log_rows(ws)
    pending = [r for r in allrows if r.get('상태') == '미확인']
    accepted = [r for r in allrows if ACCEPT_RE.match(r.get('상태', '').strip())]
    res = dict(status='관문' if open_items else '행동 테스트 대기' if pending else '완료', stage=st, checks=len(log.rows), unresolved=len(open_items),
               accepted=len(accepted), pending=len(pending), pending_items=[f'{r.get("번호")}번 · {r.get("확인한 것")}' for r in pending][:8],
               items=[f'{r[2]} · {r[3]} → {r[4]}' for r in open_items], log=ws.rel(ws.o(LOG)),
               next=('「미확인」 행은 검증 담당이 해당 담당에게 실제로 물어본 뒤 결과 · 상태 (통과 · 미해결)를 고친다 · '
                     '미확인이 남으면 리드는 다음 단계로 가지 않음') if pending else '')
    if open_items:
        raise H.Stop('관문', f'검증-{st}', f'검증 미해결 {len(open_items)}건 → 다음 단계로 가지 않음 (해당 담당이 고친 뒤 다시 검증)',
                     res['items'][:8], ['넘긴 담당이 고치고 다시 실행', H.GATE_ACCEPT])
    print(json.dumps(res, ensure_ascii=False, indent=1))


def accept(ws, args):
    """관문 인정: 사람 답 (결정로그 「관문 인정」 행 #n)을 검증로그 행 상태 「인정 (결정로그 #n)」로 옮긴다 (리드가 상태 칸을 손으로 바꾸지 않음)."""
    n = str(args.decision or '').strip()
    rows = [x.strip() for x in (args.rows or '').split(',') if x.strip()]
    dec = _decisions(ws).get(n)
    def no(why):
        print(json.dumps(dict(status='거절', why=why), ensure_ascii=False))
        sys.exit(1)
    if not rows or not n:
        no('--rows (검증로그 행 번호, 쉼표) 와 --decision (결정로그 번호) 이 필요함')
    if not dec or '관문 인정' not in dec.get('체크포인트', ''):
        no(f'결정로그 #{n} 이 없거나 체크포인트 칸이 「관문 인정」이 아님 → 사람 답을 결정로그에 먼저 한 행 적는다')
    p = ws.o(LOG)
    text = H.read(p) if os.path.exists(p) else ''
    lines = text.splitlines()
    i_st, i_note = LOG_HEAD.index('상태'), LOG_HEAD.index('해결 내용')
    done, bad = [], []
    for k in rows:
        idx = next((i for i, ln in enumerate(lines) if re.match(r'^\|\s*' + re.escape(k) + r'\s*\|', ln)), None)
        if idx is None:
            bad.append(f'{k}번 행 없음')
            continue
        cells = [c.strip() for c in lines[idx].strip().strip('|').split('|')]
        cells += [''] * (len(LOG_HEAD) - len(cells))
        if cells[i_st] not in ('미해결',) and not ACCEPT_RE.match(cells[i_st]):
            bad.append(f'{k}번 행 상태 「{cells[i_st]}」 (미해결 행만 인정)')
            continue
        cells[i_st] = accept_label(n)
        cells[i_note] = (cells[i_note] + ' · ' if cells[i_note] and not cells[i_note].startswith('관문 인정') else '') + f'관문 인정 · 값은 고치지 않음 (결정로그 #{n})'
        lines[idx] = H.md_row(cells)
        done.append(k)
    if bad:
        no(' · '.join(bad))
    H.write_md(p, '\n'.join(lines) + '\n')
    allrows = _log_rows(ws)
    print(json.dumps(dict(status='완료', rows=done, decision=n, label=accept_label(n),
                          unresolved=sum(1 for r in allrows if r.get('상태') == '미해결'),
                          accepted=sum(1 for r in allrows if ACCEPT_RE.match(r.get('상태', '').strip())),
                          next='같은 단계 verify.py check 를 다시 돌리면 같은 점검 · 같은 결과 행은 「인정」으로 적혀 관문이 열린다'), ensure_ascii=False))


def main():
    ap = H.common_args('검증 담당 스크립트')
    ap.add_argument('cmd', choices=['check', 'accept'])
    ap.add_argument('--stage', default='all', choices=['1', '2', '3', '4', 'report', 'all'])
    ap.add_argument('--no-rerun', dest='rerun', action='store_false')
    ap.add_argument('--rows', default='', help='accept: 검증로그 행 번호 (쉼표)')
    ap.add_argument('--decision', default='', help='accept: 결정로그 「관문 인정」 행 번호')
    return ap.parse_args()


if __name__ == '__main__':
    args = main()
    H.run_main({'check': check_stage, 'accept': accept}[args.cmd], lambda: args)
