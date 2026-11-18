# -*- coding: utf-8 -*-
"""pwc-hr-harness · 자사 버전 도구 (/hr-customize · Day 3).

사용 (작업 폴더에서):
  python -X utf8 <플러그인>/skills/pwc-hr-harness/scripts/customize.py lock     # 잠금 규칙 검사
  python -X utf8 <플러그인>/skills/pwc-hr-harness/scripts/customize.py diff     # 기본 대비 바뀐 파일
  python -X utf8 <플러그인>/skills/pwc-hr-harness/scripts/customize.py snapshot --co K-co [--name 기준]
  python -X utf8 <플러그인>/skills/pwc-hr-harness/scripts/customize.py regress  --co K-co [--name 기준]

lock     : 자사 버전에서도 바꾸면 안 되는 칸을 기본 사본(.claude/hr-lib/_standard)과 대조한다
           (CP1~CP5 존재 · 업무 범위 · 금지 · 감원 단독 금지 · 빈칸≠0 등 공통 규칙 · 계산 엔진)
           바꿔도 되는 것: 승인자 직책 · 멈출 때 보여 줄 기준 · CP 추가 · 변형 한 칸 · 리포트 양식 · 동인 · 계획 단위 · 4B 우선순위 기준
diff     : 기본 사본과 다른 파일 목록 (+ / − 줄 수) · lock 「changed」 · 뷰어 「기본 대비 변경점」과 같은 셈 (hrlib.harness_changes)
           셈 대상 = CLAUDE.md + .claude/ 아래 파일 (hr-lib 은 config.json 만 · settings.local.json · baselines 제외)
snapshot : outputs/<co>/ 결과의 핵심 숫자를 .claude/hr-lib/baselines/<name>.json 에 남긴다 (Day 2 끝 K-Co. 결과 = 회귀 기준)
           RPU 행 숫자 (합계 행 제외) + 합계 값 (결정 수요 · 공급 · Gap) + 방식 B · 대사 판정. CP5 승인 여부는 보지 않는다 (리포트 공유와 무관)
compare  : 지금 outputs/<co>/ 결과를 저장해 둔 결과 (snapshot --name 고치기전 등) 와 견줘 outputs/<co>/_전후비교.md (다시 계산하지 않음 · Day 3 결과 읽기)
regress  : 작업 폴더 사본에서 승인된 파라미터 · 가정 그대로 ② ③ ④ 를 다시 돌려 기준과 비교 → outputs/<co>/_회귀테스트.md
           차이가 있으면 「의도한 변화인가」를 사람이 확인한다 (변형이 K-Co. 에도 걸리는 규칙이면 차이가 나는 것이 맞을 수 있음)
"""
import os, sys, json, re, glob, shutil, tempfile, subprocess, difflib, datetime, argparse

LOCKED_CLAUDE = ['## 업무 범위 · 금지']
LOCKED_PHRASES = {
    'CLAUDE.md': ['감원 대안을 단독으로 내지 않는다', '폴더 밖 파일을 읽지 않는다', '자율성 L1', 'CP5'],
    '.claude/skills/wfp-lead/SKILL.md': ['| CP1 |', '| CP2 |', '| CP3 |', '| CP4 |', '| CP5 |', '직접 계산하지 않는다', '검증 미해결이 있으면 다음 단계로 가지 않는다'],
    '.claude/agents/wfp-gap.md': ['감원 대안을 단독으로 내지 않는다'],
    '.claude/agents/wfp-verifier.md': ['결과 파일을 고치지 않는다'],
    '.claude/skills/wfp-step2-demand/SKILL.md': ['평균 내지 않', '디지털 FTE 를 사람 FTE 에 더하지 않는다'],
    '.claude/skills/wfp-step3-supply/SKILL.md': ['0으로 읽지 않는다'],
}
LOCKED_CODE = {  # 스크립트 안 CP 강제 (구조적 멈춤) · 지워지면 안 됨
    '.claude/skills/wfp-step2-demand/scripts/step2.py': ["require_cp(ws, 'CP1'", "require_cp(ws, 'CP2'"],
    '.claude/skills/wfp-step3-supply/scripts/step3.py': ["require_cp(ws, 'CP3'"],
    '.claude/skills/wfp-step4-gap/scripts/step4.py': ["('CP2', '②_수요파라미터.md')", "('CP3', '③_공급가정.md')", "'CP4'", 'CUT'],
    '.claude/skills/wfp-report/scripts/report.py': ["require_cp(ws, 'CP4'", "'CP5'"],
}
ENGINE = ['.claude/hr-lib/hrlib.py']   # 공통 규칙 엔진 (자사 연도 등은 config.json 으로)
TEMPLATE_LIB = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))),
                            'templates', 'standard-team', '.claude', 'hr-lib')


def _lib(R):
    """작업 폴더의 hrlib (config.json 을 그 폴더에서 읽음)."""
    sys.path.insert(0, os.path.join(R, '.claude', 'hr-lib'))
    import hrlib as H
    return H


def _changes(R):
    """기본 대비 바뀐 파일 (뷰어와 같은 셈). 작업 폴더 hrlib 이 옛 판이면 플러그인 판 함수를 빌려 쓴다."""
    H = _lib(R)
    if not hasattr(H, 'harness_changes'):
        import importlib.util
        for d in (os.path.dirname(os.path.abspath(__file__)), TEMPLATE_LIB):
            q = os.path.join(d, 'hrlib.py')
            if os.path.exists(q):
                spec = importlib.util.spec_from_file_location('hrlib_plugin', q)
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                if hasattr(mod, 'harness_changes'):
                    return mod.harness_changes(R)
        return []
    return H.harness_changes(R) or []


def root():
    d = os.getcwd()
    while d != os.path.dirname(d):
        if os.path.exists(os.path.join(d, 'CLAUDE.md')) and os.path.isdir(os.path.join(d, '.claude')):
            return d
        d = os.path.dirname(d)
    return os.getcwd()


def rd(p):
    with open(p, encoding='utf-8') as f:
        return f.read()


def section(text, head):
    if head not in text:
        return None
    s = text[text.index(head):]
    m = re.search(r'\n## ', s[len(head):])
    return s[:len(head) + m.start()] if m else s


def lock(R, args):
    std = os.path.join(R, '.claude', 'hr-lib', '_standard')
    res = []
    if not os.path.isdir(std):
        print(json.dumps(dict(status='멈춤', reason='기본 사본이 없음 · /hr-init 으로 설치한 폴더에서 실행'), ensure_ascii=False))
        sys.exit(3)
    for f, phrases in LOCKED_PHRASES.items():
        p = os.path.join(R, f)
        if not os.path.exists(p):
            res.append(('삭제됨', f, '잠긴 파일이 없음'))
            continue
        t = rd(p)
        for ph in phrases:
            if ph not in t:
                res.append(('잠금 위반', f, f'「{ph}」 문장 · 규칙이 사라짐'))
    cs, cm = rd(os.path.join(std, 'CLAUDE.md')), rd(os.path.join(R, 'CLAUDE.md'))
    for h in LOCKED_CLAUDE:
        a, b = section(cs, h), section(cm, h)
        if b is None:
            res.append(('잠금 위반', 'CLAUDE.md', f'{h} 절이 없음'))
        elif a is not None:
            removed = [ln for ln in a.splitlines() if ln.strip() and ln not in b.splitlines()]
            if removed:
                res.append(('잠금 위반', 'CLAUDE.md', f'{h} 절에서 지워지거나 바뀐 줄 {len(removed)}개 · 예: {removed[0][:60]}'))
    for f, needles in LOCKED_CODE.items():
        p = os.path.join(R, f)
        t = rd(p) if os.path.exists(p) else ''
        for n in needles:
            if n not in t:
                res.append(('잠금 위반', f, f'CP 강제 코드 「{n}」 가 없음 (구조적 멈춤이 풀림)'))
    for f in ENGINE:
        a, b = os.path.join(std, f), os.path.join(R, f)
        if os.path.exists(a) and os.path.exists(b) and rd(a) != rd(b):
            res.append(('확인 필요', f, '공통 규칙 엔진이 바뀜 · 자사 연도 · 범위는 config.json 으로 바꾸고 엔진은 그대로 두는 것이 원칙'))
    # 허용된 변경 (보고만) · 뷰어 「기본 대비 변경점」 · diff 와 같은 셈
    allowed = [(c['kind'], c['file']) for c in _changes(R)]
    viol = [r for r in res if r[0] in ('잠금 위반', '삭제됨')]
    hist = section(cm, '## 변경 이력') or ''
    rows = [ln for ln in hist.splitlines() if ln.startswith('| 20')]
    if any(f != 'CLAUDE.md' for _, f in allowed) and len(rows) < 2:
        res.append(('확인 필요', 'CLAUDE.md', '바뀐 파일이 있는데 변경 이력에 행이 추가되지 않음'))
    out = dict(status='통과' if not viol else '위반', violations=[f'{a} · {b} · {c}' for a, b, c in res if a != '확인 필요'],
               warnings=[f'{b} · {c}' for a, b, c in res if a == '확인 필요'], changed=[f'{a} {b}' for a, b in allowed])
    print(json.dumps(out, ensure_ascii=False, indent=1))
    sys.exit(3 if viol else 0)


def diff(R, args):
    if not os.path.isdir(os.path.join(R, '.claude', 'hr-lib', '_standard')):
        print(json.dumps(dict(status='멈춤', reason='기본 사본이 없음 · /hr-init 으로 설치한 폴더에서 실행'), ensure_ascii=False))
        sys.exit(3)
    out = [dict(file=c['file'], kind=c['kind'], plus=c['plus'], minus=c['minus']) for c in _changes(R)]
    print(json.dumps(dict(status='완료', count=len(out), changed=out), ensure_ascii=False, indent=1))


TOTAL = '합계'


def _numbers(R, co):
    """(RPU 행 숫자 · 판정 {항목: {RPU: 값}} · 합계 값 {항목: 값}). 합계 행은 RPU 행에 넣지 않는다."""
    H = _lib(R)
    o = os.path.join(R, 'outputs', co)
    ty = H.config()['target_year']
    n, tot = {}, {}
    rows_of = lambda rs, col: {r['RPU 코드']: r[col] for r in rs if r.get('RPU 코드') != TOTAL}
    total_of = lambda rs, col: next((r[col] for r in rs if r.get('RPU 코드') == TOTAL), None)
    if os.path.exists(os.path.join(o, '②_대사.xlsx')):
        _, D = H.read_xlsx_table(os.path.join(o, '②_대사.xlsx'))
        n['결정 수요 사람'] = rows_of(D, '결정 수요 사람')
        n['결정 수요 디지털'] = rows_of(D, '결정 수요 디지털')
        n['방식 B'] = {k: v for k, v in rows_of(D, 'B').items() if isinstance(v, (int, float))}
        n['대사 판정'] = {r['RPU 코드']: f"{r.get('구간') or ''} · {r.get('판정') or ''}" for r in D
                         if r.get('RPU 코드') != TOTAL and r.get('구간') != '사업 이벤트'}
        tot['결정 수요 사람'] = total_of(D, '결정 수요 사람')
        tot['결정 수요 디지털'] = total_of(D, '결정 수요 디지털')
    if os.path.exists(os.path.join(o, '②C_업무량.xlsx')):
        _, C = H.read_xlsx_table(os.path.join(o, '②C_업무량.xlsx'))
        n['2025 FTE'] = rows_of(C, '2025 FTE')
        tot['2025 FTE'] = total_of(C, '2025 FTE')
    if os.path.exists(os.path.join(o, '③_공급.xlsx')):
        _, S = H.read_xlsx_table(os.path.join(o, '③_공급.xlsx'), '기본')
        _, U = H.read_xlsx_table(os.path.join(o, '③_공급.xlsx'), '상승')
        n[f'공급 {ty}'] = rows_of(S, f'공급 {ty}')
        tot[f'공급 {ty}'] = total_of(S, f'공급 {ty}')
        tot[f'공급 {ty} 상승'] = total_of(U, f'공급 {ty}')
    g = os.path.join(o, '④_Gap대안.md')
    if os.path.exists(g):
        rows = next((rows for h, c, rows in H.md_tables(H.read(g)) if h.startswith('1. Gap')), [])
        n['Gap'] = {r['RPU 코드']: float(r['Gap']) for r in rows if r.get('Gap') and r['RPU 코드'] != TOTAL}
        t = next((r for r in rows if r['RPU 코드'] == TOTAL), None)
        if t:
            tot['Gap'] = float(t['Gap'])
            if t.get('Gap (공급 상승)'):
                tot['Gap 상승'] = float(t['Gap (공급 상승)'])
    return n, {k: v for k, v in tot.items() if v is not None}


def snapshot(R, args):
    n, tot = _numbers(R, args.co)
    d = os.path.join(R, '.claude', 'hr-lib', 'baselines')
    os.makedirs(d, exist_ok=True)
    p = os.path.join(d, f'{args.co}_{args.name}.json')
    rp = glob.glob(os.path.join(R, 'outputs', args.co, '리포트_*.md'))
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(dict(co=args.co, name=args.name, at=datetime.datetime.now().isoformat(timespec='minutes'), numbers=n, totals=tot),
                  f, ensure_ascii=False, indent=1)
    r1 = lambda v: round(v, 1) if isinstance(v, (int, float)) else v
    print(json.dumps(dict(status='완료', file=os.path.relpath(p, R).replace('\\', '/'),
                          rpu={k: len(v) for k, v in n.items()}, totals={k: r1(v) for k, v in tot.items()},
                          note='RPU 수는 합계 행을 뺀 수 (사업 이벤트 줄 포함) · CP5 승인 여부와 무관하게 남김'
                               + ('' if rp else ' · 리포트 없음 (Day 2 끝 리포트까지 마친 뒤 남기는 것이 원칙)')),
                     ensure_ascii=False))


def regress(R, args):
    p = os.path.join(R, '.claude', 'hr-lib', 'baselines', f'{args.co}_{args.name}.json')
    if not os.path.exists(p):
        print(json.dumps(dict(status='멈춤', reason=f'기준 스냅샷이 없음: snapshot --co {args.co} --name {args.name} 먼저 (Day 2 끝 결과)'), ensure_ascii=False))
        sys.exit(3)
    bj = json.load(open(p, encoding='utf-8'))
    base, base_tot = bj['numbers'], bj.get('totals', {})
    tmp = tempfile.mkdtemp(prefix='hr_regress_')
    try:
        w = os.path.join(tmp, 'w')
        shutil.copytree(R, w, ignore=shutil.ignore_patterns('*.html', '_workspace*', '.git'))
        sk = os.path.join(w, '.claude', 'skills')
        log = []
        for sc, cmd in (('wfp-step2-demand/scripts/step2.py', 'run'), ('wfp-step3-supply/scripts/step3.py', 'run'), ('wfp-step4-gap/scripts/step4.py', 'run')):
            force = ['--force'] if sc.startswith('wfp-step4') and "'--force'" in rd(os.path.join(sk, sc)) else []   # 임시 사본이라 ④ 4절 손 칸 보호를 넘김 (작업 폴더는 그대로)
            r = subprocess.run([sys.executable, '-X', 'utf8', os.path.join(sk, sc), cmd, '--co', args.co] + force + sum([['--allow', x] for x in args.allow], []),
                               cwd=w, capture_output=True, text=True, encoding='utf-8')
            log.append(f'{sc} {cmd} → 종료 코드 {r.returncode}')
            if r.returncode != 0:
                log.append(r.stdout[-600:])
                break
        now, now_tot = _numbers(w, args.co)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    rows, nd = [], 0
    show = lambda v: v if not isinstance(v, (int, float)) else round(v, 1)
    for k, b in base.items():
        a = now.get(k, {})
        for code in sorted((set(b) | set(a)) - {TOTAL}):   # 옛 기준 파일의 합계 행은 아래 합계 비교로 봄
            x, y = b.get(code), a.get(code)
            if isinstance(x, (int, float)) and isinstance(y, (int, float)) and abs(x - y) <= 0.05:
                continue
            if x == y:
                continue
            nd += 1
            rows.append(f'| {k} | {code} | {show(x)} | {show(y)} |')
    for k in sorted(set(base_tot) | {k for k, v in base.items() if TOTAL in v}):
        x = base_tot.get(k, base.get(k, {}).get(TOTAL))
        y = now_tot.get(k)
        if isinstance(x, (int, float)) and isinstance(y, (int, float)) and abs(x - y) <= 0.05:
            continue
        nd += 1
        rows.append(f'| {k} | 합계 | {show(x)} | {show(y)} |')
    text = '\n'.join([f'# 회귀 테스트 ({args.co} · 기준 「{args.name}」)', '',
                      f'- 시각: {datetime.datetime.now():%Y-%m-%d %H:%M} · 방법: 작업 폴더 사본에서 승인된 파라미터 · 가정 그대로 ② ③ ④ 를 다시 돌림',
                      '- 비교 항목: 결정 수요 사람 · 디지털 · 방식 B · 대사 판정 (구간 · 판정) · 2025 FTE · 공급 · Gap (RPU 행 + 합계) · 기준 파일에 없는 항목은 비교하지 않음',
                      '- 기준 합계: ' + (' · '.join(f'{k} {show(v)}' for k, v in base_tot.items()) or '(옛 기준 파일 · 합계 없음)'),
                      f'- 결과: **{"차이 없음 · 변형이 기본 계산을 깨지 않음" if nd == 0 else f"차이 {nd}칸 · 의도한 변화인지 사람이 확인"}**', '',
                      '## 실행 기록', '', *[f'- {x}' for x in log], '',
                      '## 차이', '', '| 항목 | RPU | 기준 | 지금 |', '|---|---|---|---|', *rows])
    o = os.path.join(R, 'outputs', args.co, '_회귀테스트.md')
    with open(o, 'w', encoding='utf-8') as f:
        f.write(text + '\n')
    print(json.dumps(dict(status='통과' if nd == 0 else '차이', diffs=nd, file=os.path.relpath(o, R).replace('\\', '/')), ensure_ascii=False))
    sys.exit(0 if nd == 0 else 3)


def compare(R, args):
    """지금 outputs/<co>/ 결과를 저장해 둔 결과 (snapshot --name) 와 견준다 (다시 계산하지 않음) → outputs/<co>/_전후비교.md.
    Day 3 결과 읽기: 내 회사 결과를 「고치기전」으로 저장 → 진단표 한 줄을 고쳐 다시 계산 → compare 로 무엇이 얼마나 달라졌는지 본다."""
    p = os.path.join(R, '.claude', 'hr-lib', 'baselines', f'{args.co}_{args.name}.json')
    if not os.path.exists(p):
        print(json.dumps(dict(status='멈춤', reason=f'비교할 저장 결과가 없음: snapshot --co {args.co} --name {args.name} 먼저 (고치기 전에 저장)'), ensure_ascii=False))
        sys.exit(3)
    bj = json.load(open(p, encoding='utf-8'))
    base, base_tot = bj['numbers'], bj.get('totals', {})
    now, now_tot = _numbers(R, args.co)
    H = _lib(R)
    ty = H.config()['target_year']
    show = lambda v: '·' if v is None else (f'{v:,.1f}' if isinstance(v, (int, float)) else str(v))
    dlt = lambda x, y: (y - x) if isinstance(x, (int, float)) and isinstance(y, (int, float)) else None
    keys = ['결정 수요 사람', '결정 수요 디지털', f'공급 {ty}', f'공급 {ty} 상승', 'Gap', 'Gap 상승']
    trows, changed = [], 0
    for k in keys:
        x, y = base_tot.get(k), now_tot.get(k)
        if x is None and y is None:
            continue
        d = dlt(x, y)
        if d is not None and abs(d) > 0.05:
            changed += 1
        trows.append(f'| {k} | {show(x)} | {show(y)} | {"·" if d is None or abs(d) <= 0.05 else f"{d:+,.1f}"} |')
    rrows = []
    for k in ('결정 수요 사람', f'공급 {ty}', 'Gap'):
        b, a = base.get(k, {}), now.get(k, {})
        for code in sorted(set(b) | set(a)):
            d = dlt(b.get(code), a.get(code))
            if d is None and b.get(code) == a.get(code):
                continue
            if d is not None and abs(d) <= 0.05:
                continue
            rrows.append((abs(d) if d is not None else 0, f'| {k} | {code} | {show(b.get(code))} | {show(a.get(code))} | {"새로 생김 · 없어짐" if d is None else f"{d:+,.1f}"} |'))
    rrows = [r for _, r in sorted(rrows, key=lambda x: -x[0])]
    text = '\n'.join([f'# 고치기 전후 비교 ({args.co} · 「{args.name}」 → 지금)', '',
                      f'- 시각: {datetime.datetime.now():%Y-%m-%d %H:%M} · 저장해 둔 결과 ({bj.get("at", "")}) 와 지금 outputs/{args.co}/ 결과를 견줌 (다시 계산하지 않음)',
                      f'- 결과: **{"달라진 합계 없음" if changed == 0 else f"합계 {changed}개가 달라짐"} · 조직 묶음 {len(rrows)}칸이 달라짐**', '',
                      '## 합계', '', '| 항목 | 고치기 전 | 지금 | 차이 |', '|---|---|---|---|', *trows, '',
                      '## 달라진 조직 묶음 (차이가 큰 순서)', '', '| 항목 | RPU | 고치기 전 | 지금 | 차이 |', '|---|---|---|---|---|', *(rrows or ['| · | · | · | · | 달라진 칸 없음 |'])])
    o = os.path.join(R, 'outputs', args.co, '_전후비교.md')
    with open(o, 'w', encoding='utf-8') as f:
        f.write(text + '\n')
    print(json.dumps(dict(status='완료', totals_changed=changed, rpu_changed=len(rrows), file=os.path.relpath(o, R).replace('\\', '/')), ensure_ascii=False))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('cmd', choices=['lock', 'diff', 'snapshot', 'regress', 'compare'])
    ap.add_argument('--co', default='K-co')
    ap.add_argument('--name', default='기준')
    ap.add_argument('--allow', action='append', default=[])
    a = ap.parse_args()
    R = root()
    dict(lock=lock, diff=diff, snapshot=snapshot, regress=regress, compare=compare)[a.cmd](R, a)


if __name__ == '__main__':
    main()
