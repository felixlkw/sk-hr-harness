# -*- coding: utf-8 -*-
"""③ 인력공급 예측 · 채용 없는 공급 (기본 · 상승).

사용:
  python -X utf8 .claude/skills/wfp-step3-supply/scripts/step3.py rates --co K-co   # 퇴직률 산출 + ③_공급가정.md → CP3
  python -X utf8 .claude/skills/wfp-step3-supply/scripts/step3.py run   --co K-co   # CP3 승인 뒤 ③_공급.xlsx

rates : CP1 승인 확인 → RPU별 자발 퇴직률 (3개년 합 비율 · 최고 연도) · 정년 도래 → ③_공급가정.md 부록 갱신 → CP3 멈춤
        자발퇴직 칸이 비면 0 으로 읽지 않고 「예외」로 멈춘다 (연도별 요약으로 대체할지 물음)
run   : CP3 승인 + 사람 답 5칸 확인 → 기말 = 기초 − 기초 × 퇴직률 × 배수 − 정년 도래 (2년)
"""
import os, sys, json, re
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', '..', 'hr-lib'))
import hrlib as H

SCRIPT = 'wfp-step3-supply/scripts/step3.py'
PF = '③_공급가정.md'


def _rates(ws, args, m):
    hy = [str(y) for y in ws.cfg['hist_years']]
    hist = H.load(ws, 'hist', m)
    blank = [r for r in hist if r.get('자발퇴직') is None]
    mode = H.allowed(args, '퇴직빈칸')
    if blank and not mode:
        raise H.Stop('예외', '퇴직빈칸', f'입퇴사 이력 자발퇴직 칸이 {len(blank)}행 비어 있음 → 퇴직률을 어떻게 낼지',
                     [f'예: {blank[0]["_ref"]}', '빈칸은 0 으로 읽지 않음 (공통 규칙 4 · 퇴직률 0% 로 계산하면 공급이 과대)'],
                     ['자료를 채워 다시 받기 (권장)', '연도별 요약 시트의 전사 퇴직률로 모든 RPU 대체 (`--allow 퇴직빈칸=요약` · 가정 표에 적음)', '보류'],
                     note='공급 숫자를 내지 않았다.')
    beg, vol, by_y = {}, {}, {}
    for r in hist:
        y = H.norm(r['연도'])
        if y not in hy:
            continue
        c = r['RPU 코드']
        b = r.get('기초') or 0
        beg[c] = beg.get(c, 0) + b
        by_y.setdefault(c, {}).setdefault(y, [0, 0])
        by_y[c][y][0] += b
        if r.get('자발퇴직') is not None:
            vol[c] = vol.get(c, 0) + r['자발퇴직']
            by_y[c][y][1] += r['자발퇴직']
    rate, up, note = {}, {}, {}
    if blank and mode == '요약':
        hs = {H.norm(r['연도']): r for r in H.load(ws, 'histsum', m)}
        tb = sum((hs[y]['기초'] or 0) for y in hy if y in hs)
        tv = sum((hs[y]['자발퇴직'] or 0) for y in hy if y in hs)
        ys = [hs[y]['자발퇴직'] / hs[y]['기초'] for y in hy if y in hs and hs[y]['기초']]
        for c in beg:
            rate[c], up[c], note[c] = tv / tb, max(ys), '연도별 요약 전사 퇴직률로 대체 (예외 답)'
    else:
        for c in beg:
            rate[c] = vol.get(c, 0) / beg[c] if beg[c] else 0
            ys = [v[1] / v[0] for v in by_y[c].values() if v[0]]
            up[c] = max(ys) if ys else rate[c]
            note[c] = ''
    return rate, up, note


def rates(ws, args):
    H.require_cp(ws, 'CP1', '①_열대응표.md', '③ 공급 가정 작성')
    m = H.load_map(ws)
    cfg = ws.cfg
    by = cfg['base_year']
    y1, y2 = str(by + 1), str(by + 2)
    org = {r['RPU 코드']: r for r in H.load(ws, 'org', m)}
    ret = {r['RPU 코드']: r for r in H.load(ws, 'retire', m)}
    rate, up, note = _rates(ws, args, m)
    r1 = sum((ret.get(c, {}).get(y1) or 0) for c in org)
    r2 = sum((ret.get(c, {}).get(y2) or 0) for c in org)
    p = ws.o(PF)
    replaced = False
    if os.path.exists(p) and not H.is_blank_form(H.read(p)):   # 사람 답 · 승인이 적힌 파일만 보존한다
        text = H.read(p)
        if not any('승인' in h and 'CP' in c for h, c, _ in H.md_tables(text)):   # 승인 표가 없는 옛 양식 → 승인 표만 덧붙임 (사람 칸은 그대로)
            text = text.rstrip() + '\n\n' + H.approval_block('CP3', '인사기획팀장')
    else:   # 없거나 양식 그대로 (사람 칸이 모두 빔) → 플러그인 양식으로 새로 만든다
        replaced = os.path.exists(p)
        text = H.form(PF).format(co=ws.co, by=by, y1=y1, y2=y2, r1=int(r1), r2=int(r2),
                                 hist='~'.join([str(cfg['hist_years'][0]), str(cfg['hist_years'][-1])]))
    rows = [[c, org[c]['계'], f'{rate[c]:.2%}', f'{up[c]:.2%}', ret.get(c, {}).get(y1), ret.get(c, {}).get(y2),
             '표본 작음' if org[c]['계'] < cfg['small_rpu'] else '', note[c] or '파일 3 · 입퇴사 이력'] for c in org]
    tb = sum(org[c]['계'] for c in org)
    block = '\n'.join(['## 부록 A. RPU별 퇴직률 · 정년 (스크립트가 씀 · 사람이 고치지 않음)', '',
                       f'- 만든 스크립트: `{SCRIPT} rates` · 현원 합 {tb:g}명 · 정년 {y1} {r1:g}명 · {y2} {r2:g}명',
                       '', H.md_table(['RPU 코드', '현원', '자발 퇴직률 (기본)', '최고 연도 (상승)', f'정년 {y1}', f'정년 {y2}', '주의', '출처'], rows)])
    H.write_md(p, H.replace_auto(text, '퇴직률', block))
    _status(p)
    a = H.approval(p, 'CP3')
    if not a['ok']:
        raise H.Stop('CP', 'CP3', '공급 가정 S-1 ~ S-5 확인 (사람 답 5칸 + 승인)',
                     [ws.rel(p), f'RPU {len(org)}개 · 퇴직률 범위 {min(rate.values()):.1%} ~ {max(rate.values()):.1%}']
                     + (['양식 그대로였던 파일을 플러그인 양식으로 새로 만듦'] if replaced else []),
                     ['확정 (기본값 그대로)', '확정 · 바꾼 값을 그 행 사람 답 칸에 (예: 배수 1.2)', '추가 확인 · 보류 · 이관'])
    print(json.dumps(dict(status='완료', note='CP3 이미 승인됨 · run 실행 가능', file=ws.rel(p)), ensure_ascii=False))


def _status(p):
    """머리 「상태」 줄을 CP3 승인 칸에 맞춤 (2026-11-16 · 리허설: 확정 뒤에도 「승인 전」으로 남던 것)"""
    t = H.read(p)
    a = H.approval(p, 'CP3')
    new = f'- 상태: **CP3 승인됨** ({a["승인자"] or "승인자 칸"} · {a["승인일"] or "날짜 칸"})' if a['ok'] else '- 상태: **승인 전**'
    t2 = re.sub(r'^- 상태: \*\*[^*\n]+\*\*.*$', lambda _m: new, t, count=1, flags=re.M)
    if t2 != t:
        H.write_md(p, t2)


def run(ws, args):
    H.require_cp(ws, 'CP1', '①_열대응표.md', '③ 공급 계산')
    p = ws.o(PF)
    H.require_cp(ws, 'CP3', PF, '③ 공급 계산')
    S = {}
    for head, cols, rows in H.md_tables(H.read(p)):
        if head.startswith('2. 가정'):
            S = {r['번호']: r for r in rows}
    empty = [k for k, r in S.items() if not r.get('사람 답', '').strip()]
    if empty:
        raise H.Stop('CP', 'CP3', '가정 표 「사람 답」 칸이 비어 있음 (빈 행으로는 계산하지 않음)', [PF + ' · ' + ', '.join(empty)],
                     ['사람 답 채우기 (확정 · 추가 확인 · 보류 · 이관)', '보류'])
    not_ok = [k for k, r in S.items() if not r['사람 답'].strip().startswith('확정')]
    if not_ok:
        raise H.Stop('CP', 'CP3', '「확정」이 아닌 가정이 있음', [PF + ' · ' + ', '.join(f'{k}: {S[k]["사람 답"]}' for k in not_ok)],
                     ['확정으로 정리', '보류 유지 (계산하지 않음)'])
    _status(p)
    mult = 1.0
    mm = re.search(r'배수\s*([\d.]+)', S.get('S-2', {}).get('사람 답', ''))
    if mm:
        mult = float(mm.group(1))
    recent = '최근 1년' in S.get('S-1', {}).get('사람 답', '')
    m = H.load_map(ws)
    cfg = ws.cfg
    by = cfg['base_year']
    y1, y2 = str(by + 1), str(by + 2)
    org = {r['RPU 코드']: r for r in H.load(ws, 'org', m)}
    ret = {r['RPU 코드']: r for r in H.load(ws, 'retire', m)}
    rate, up, note = _rates(ws, args, m)
    if recent:
        hist = H.load(ws, 'hist', m)
        ly = str(cfg['hist_years'][-1])
        for c in org:
            rs = [r for r in hist if r['RPU 코드'] == c and H.norm(r['연도']) == ly]
            b = sum(r['기초'] or 0 for r in rs)
            rate[c] = sum(r['자발퇴직'] or 0 for r in rs) / b if b else rate[c]

    def path(rt, mu):
        out = {}
        for c in org:
            n = org[c]['계']
            row = {}
            for y in (y1, y2):
                n = n - n * rt[c] * mu - (ret.get(c, {}).get(y) or 0)
                row[y] = n
            out[c] = row
        return out

    base, upp = path(rate, mult), path(up, 1.0)
    sheets = []
    for nm, P, rt in (('기본', base, rate), ('상승', upp, up)):
        rows = [[c, org[c]['계'], round(rt[c] * (mult if nm == '기본' else 1), 4), ret.get(c, {}).get(y1), ret.get(c, {}).get(y2),
                 P[c][y1], P[c][y2], P[c][y2] / org[c]['계'],
                 '표본 작음' if org[c]['계'] < cfg['small_rpu'] else '', note[c] or f'파일 2 · 파일 3 · {org[c]["_ref"]}'] for c in org]
        rows.append(['합계', sum(org[c]['계'] for c in org), '', sum((ret.get(c, {}).get(y1) or 0) for c in org),
                     sum((ret.get(c, {}).get(y2) or 0) for c in org), round(sum(P[c][y1] for c in org), 1),
                     round(sum(P[c][y2] for c in org), 1), '', '', ''])
        sheets.append(dict(name=nm, title=f'③ 공급 {y2} ({ws.co}) · {nm} 시나리오 · 채용 없음',
                           notes=[f'기말 = 기초 − 기초 × 퇴직률{" × 배수 " + str(mult) if nm == "기본" and mult != 1 else ""} − 정년 도래 · {y1} · {y2}',
                                  '기본 = 3개년 합 비율' + (' (S-1 최근 1년)' if recent and nm == '기본' else '') if nm == '기본' else '상승 = RPU별 가장 높은 연도 퇴직률',
                                  '신규 채용 · 신설 조직 · 디지털 FTE 는 넣지 않음 · 「공급 ' + y2 + '」 열 이름은 wfp-step4-gap 입력 계약과 같아야 한다'],
                           head=['RPU 코드', '현원', '퇴직률', f'정년 {y1}', f'정년 {y2}', f'공급 {y1}', f'공급 {y2}', '현원 대비', '주의', '출처'], rows=rows))
    H.write_xlsx(ws.o('③_공급.xlsx'), sheets, SCRIPT, ['①_열대응표.json', PF, m['tables']['hist']['file'], m['tables']['org']['file']])
    print(json.dumps(dict(status='완료', 공급=dict(기본=round(sum(base[c][y2] for c in org), 1), 상승=round(sum(upp[c][y2] for c in org), 1)),
                          배수=mult, 정년=sum((ret.get(c, {}).get(y1) or 0) + (ret.get(c, {}).get(y2) or 0) for c in org),
                          file=ws.rel(ws.o('③_공급.xlsx'))), ensure_ascii=False, indent=1))


def main():
    ap = H.common_args('③ 인력공급 예측 스크립트')
    ap.add_argument('cmd', choices=['rates', 'run'])
    return ap.parse_args()


if __name__ == '__main__':
    args = main()
    H.run_main({'rates': rates, 'run': run}[args.cmd], lambda: args)
