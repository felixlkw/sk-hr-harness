# -*- coding: utf-8 -*-
"""④ 수급 Gap 분석 · 확보·육성계획.

사용:
  python -X utf8 .claude/skills/wfp-step4-gap/scripts/step4.py run   --co K-co   # Gap 표 · 유형 · 스킬 근거 · 대안표 틀
  python -X utf8 .claude/skills/wfp-step4-gap/scripts/step4.py check --co K-co   # Gap 담당이 채운 대안표 점검 → CP4

run   : ②_대사.xlsx 결정 수요 − ③_공급.xlsx 공급 (기본 · 상승) → ④_Gap대안.md 의 스크립트 칸(auto 블록)을 쓴다.
        대안표 · 재배치 경로는 Gap 담당이 채우는 칸이라 이미 적힌 내용을 지우지 않는다.
check : 대안표 규칙 점검 (4B 값 · 근거 · 감원 단독 금지 · 여유 RPU 재배치 경로) → 통과하면 CP4 에서 멈춘다.
"""
import os, sys, json, re
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', '..', 'hr-lib'))
import hrlib as H
import openpyxl

SCRIPT = 'wfp-step4-gap/scripts/step4.py'
OF = '④_Gap대안.md'
FB = ('Buy', 'Build', 'Borrow', 'Bot')
CUT = re.compile(r'감원|해고|구조조정|희망퇴직|정리해고|Bump', re.I)


CP4_WHO = '인사기획팀장 + R 1순위 행 소관 본부장'   # 기본 승인자 (config.json approvers 가 있으면 그것)
HAND4 = ('Staffing 변화', '다시 돌릴 트리거 지표')   # 4절 시나리오 표에서 Gap 담당이 채우는 손 칸

def _params(ws):
    T = {}
    for h, cols, rows in H.md_tables(H.read(ws.o('②_수요파라미터.md'))):
        if re.match(r'^\d\.', h):
            T[h.split('.')[0]] = rows
    ai = {r['RPU 코드'] for r in T.get('3', []) if r.get('AI 적용', '').strip() in ('○', 'O', 'o', '예', 'Y')}
    timing = {r['RPU 코드']: f"{r['내용']} ({r['시점']})" for r in T.get('5', []) if r.get('RPU 코드', '').strip()}
    events = {r['코드']: r for r in T.get('4', []) if r.get('코드', '').strip()}
    return ai, timing, events


def _skill_basis(ws, m):
    """파일 5 두 시트의 기준 연도가 다르면 섞지 않는다 (공통 규칙 2 · 경계 입력 4)."""
    out = {}
    for key in ('skill', 'skillg'):
        t = m['tables'].get(key) or {}
        if t.get('file'):
            sh = openpyxl.load_workbook(os.path.join(ws.data, t['file']), data_only=True)[t['sheet']]
            out[key] = H.sheet_basis(sh)
    return out


def _skills(ws, m, gap, surplus, args):
    b = _skill_basis(ws, m)
    if b.get('skill') and b.get('skillg') and b['skill'] != b['skillg']:
        use = H.allowed(args, '스킬기준일')
        if not use:
            raise H.Stop('예외', '스킬기준일', '스킬 인벤토리 두 시트의 기준 연도가 다름 → Build 가능 여부를 어느 시트로 판단할지',
                         [f'스킬 인벤토리 시트: {b["skill"]}', f'직급별 보유 분포 시트: {b["skillg"]}', '두 연도를 한 표로 합치지 않음 (공통 규칙 2)'],
                         ['스킬 인벤토리 시트만 씀 (`--allow 스킬기준일=스킬인벤토리`)', '자료를 같은 기준 연도로 다시 받기', '보류 · Build 판정 「추가 확인」'])
    sk = H.load(ws, 'skill', m, need=False)
    basis = b.get('skill') or '(기준 표시 없음)'
    by_rpu = {}
    for s in sk:
        by_rpu.setdefault(s['RPU 코드'], []).append(s)
    rows, adj = [], []
    for c, g in sorted(gap.items(), key=lambda x: -x[1]):
        if g <= 0.5 or c not in by_rpu:
            continue
        req = [s for s in by_rpu[c] if (s.get('요구 레벨 2028') or 0) >= 2]
        need_ftc = sum((s['인원 계'] or 0) * s['요구 레벨 2028'] for s in req)
        have_ftc = sum(sum((s.get(f'L{i}') or 0) * i for i in range(5)) for s in req)
        low = [f"{s['스킬']} L{s['요구 레벨 2028']} 충족 {(s['2028 요구 충족 인원'] or 0) / (s['인원 계'] or 1):.0%}" for s in req
               if (s['2028 요구 충족 인원'] or 0) / (s['인원 계'] or 1) < 0.5]
        rows.append([c, round(g, 1), len(req), need_ftc, have_ftc, f'{have_ftc / need_ftc:.0%}' if need_ftc else '-',
                     ' · '.join(low[:3]) or '-', f'파일 5 · 스킬 인벤토리 · {basis}'])
        tgt = {s['스킬'] for s in req}
        tgrp = {s['그룹'] for s in req}
        for src in surplus:
            mine = by_rpu.get(src, [])
            same = {s['스킬'] for s in mine if s['스킬'] in tgt and sum(s.get(f'L{i}') or 0 for i in range(2, 5)) > 0}
            grp = {s['그룹'] for s in mine} & tgrp
            if tgt and (same or grp):
                adj.append([c, src, round(surplus[src], 1), f'{len(same) / len(tgt):.0%}', ' · '.join(sorted(same)) or '-',
                            ' · '.join(sorted(grp)) or '-'])
    return rows, adj, basis


def _types(c, g, hc, ai, timing, events):
    if c in events:
        return ['수량', '시점']
    t = []
    if g > 0.5:
        t.append('수량')
    if c in timing:
        t.append('시점')
    if g > 0 and hc and g / hc >= 0.30:   # 시점과 겹쳐도 함께 적음 (1절 머리 「겹치면 함께 적음」)
        t.append('구조')
    if c in ai:
        t.append('AI 활용')
    return t or ['균형']


def run(ws, args):
    for cp, f in (('CP2', '②_수요파라미터.md'), ('CP3', '③_공급가정.md')):
        H.require_cp(ws, cp, f, '④ Gap 분석')
    for f in ('②_대사.xlsx', '③_공급.xlsx', '②A_매출이익.xlsx'):
        if not os.path.exists(ws.o(f)):
            raise H.Stop('예외', '입력없음', f'{f} 가 없어 Gap 을 계산할 수 없음', [ws.rel(ws.o(f))], ['앞 단계 실행', '보류'])
    ty = ws.cfg['target_year']
    _, dem = H.read_xlsx_table(ws.o('②_대사.xlsx'))
    _, sup = H.read_xlsx_table(ws.o('③_공급.xlsx'), '기본')
    _, supu = H.read_xlsx_table(ws.o('③_공급.xlsx'), '상승')
    _, mac = H.read_xlsx_table(ws.o('②A_매출이익.xlsx'))
    need = ['RPU 코드', '결정 수요 사람', '결정 수요 디지털']
    if not dem or any(k not in dem[0] for k in need):
        raise H.Stop('예외', '이음매', '②_대사.xlsx 열 이름이 입력 계약과 다름', [str(list(dem[0].keys()) if dem else '빈 표')], ['② 다시 실행'])
    sk_ = f'공급 {ty}'
    S = {r['RPU 코드']: r[sk_] for r in sup if r['RPU 코드'] != '합계'}
    SU = {r['RPU 코드']: r[sk_] for r in supu if r['RPU 코드'] != '합계'}
    HC = {r['RPU 코드']: r['현원'] for r in sup if r['RPU 코드'] != '합계'}
    ai, timing, events = _params(ws)
    D = {r['RPU 코드']: r for r in dem if r['RPU 코드'] not in ('합계',)}
    gap, gapu, rowsG = {}, {}, []
    for c, r in D.items():
        dh = r['결정 수요 사람'] or 0
        dd = r['결정 수요 디지털'] or 0
        s = S.get(c, 0.0)
        g = dh - s
        gap[c], gapu[c] = g, dh - SU.get(c, 0.0)
        t = _types(c, g, HC.get(c), ai, timing, events)
        direction = '부족' if g > 0.5 else '여유' if g < -0.5 else '균형'
        rowsG.append([c, HC.get(c, '-'), round(dh, 1), round(dd, 1) if dd else '', round(s, 1), round(g, 1), round(gapu[c], 1),
                      ' · '.join(t), direction, timing.get(c, events.get(c, {}).get('이름', '') if c in events else '')])
    tot = lambda d: sum(d.values())
    rowsG.append(['합계', sum(v for v in HC.values() if isinstance(v, (int, float))), round(sum((r['결정 수요 사람'] or 0) for r in D.values()), 1),
                  round(sum((r['결정 수요 디지털'] or 0) for r in D.values()), 1), round(tot(S), 1), round(tot(gap), 1), round(tot(gapu), 1), '', '', ''])
    surplus = {c: -g for c, g in gap.items() if g < -0.5 and c not in events}
    redep_n = sum(H.round_shown(v) for v in surplus.values())   # 정수 규칙: 행마다 여유 표시값(소수 1자리) 반올림
    sk_rows, adj_rows, basis = _skills(ws, M := H.load_map(ws), gap, surplus, args)
    reg = [r for r in mac if str(r.get('길', '')).startswith('회귀')]
    rowsS = [[r['시나리오'], r[f'{ty} 인원'], '', ''] for r in reg]
    auto1 = '\n'.join(['## 1. Gap 표 (스크립트)', '',
                       f'- Gap = 결정 수요 사람 FTE − 공급 {ty} (기본) · 상승 = 공급 상승 시나리오 · 디지털 FTE 는 「도입 계획」 칸 (Gap 에 넣지 않음)',
                       '- 유형 4: 수량 (Gap > 0.5) · 시점 (시점 이벤트) · 구조 (Gap ÷ 현원 ≥ 30%) · AI 활용 (AI 적용 RPU) · 겹치면 함께 적음',
                       f'- 합계: 결정 수요 사람 {rowsG[-1][2]} · 공급 {rowsG[-1][4]} · **Gap {rowsG[-1][5]}** (상승 {rowsG[-1][6]})',
                       f'- 여유 RPU {len(surplus)}개 · 사람 여유 합 {sum(surplus.values()):.1f} · 정수 규칙 (행마다 여유 표시값 반올림) 적용 시 합 {redep_n}명 · 차이는 3절 「정수 처리」 줄에 근거로 적음', '',
                       H.md_table(['RPU 코드', '현원', '결정 수요 사람', '디지털 (도입 계획)', f'공급 {ty}', 'Gap', 'Gap (공급 상승)', '유형', '방향', '시점 이벤트'], rowsG)])
    auto5 = '\n'.join(['## 5. 스킬 근거 (Build 가능 여부 · 법전 Skill 장)', '',
                       f'- FTC = Σ 인원 × 요구 레벨 (요구 FTC) · 보유 FTC = Σ 레벨별 인원 × 레벨 · 기준: {basis}',
                       '- 충족률이 낮고 육성 기간이 Gap 시점보다 길면 Build 만으로는 메우기 어렵다 → Buy · Borrow 병행 검토', '',
                       H.md_table(['RPU 코드', 'Gap', '요구 L2 이상 스킬 수', '요구 FTC', '보유 FTC', '보유 ÷ 요구', '충족 50% 미만 스킬', '출처'], sk_rows) if sk_rows else '- 스킬 인벤토리 없음 · Build 판정 「추가 확인」', '',
                       '### 인접 스킬 재배치 후보 (여유 RPU → 부족 RPU)', '',
                       '- 겹침률 = 받는 RPU 요구 스킬 중 보내는 RPU 가 L2 이상 보유한 스킬 비율 · 같은 그룹 = 짧은 교육으로 옮길 수 있는 후보', '',
                       H.md_table(['받는 RPU', '보내는 RPU', '보내는 RPU 여유', '겹침률', '같은 스킬', '같은 그룹'], adj_rows) if adj_rows else '- 후보 없음',
                       *(['', f'- 신설 조직 (사업 이벤트 · 스킬 인벤토리 행 없음 · 위 표에 나오지 않음): {" · ".join(sorted(events))} → 받는 RPU 로 쓸 수 있음 · 근거는 파일 1 인력 운영 방향 문장 · 겹침률 대신 「신설 · 스킬 행 없음」'] if events else [])])
    auto4 = '\n'.join(['## 4. 시나리오 (스크립트 숫자 + Gap 담당이 트리거 지표를 적음)', '',
                       H.md_table(['시나리오', f'A 회귀 인원 ({ty})', 'Staffing 변화', '다시 돌릴 트리거 지표'], rowsS)])
    p = ws.o(OF)
    reset = ''
    if os.path.exists(p):
        text = H.read(p)
        hand = [f'{r.get("시나리오", "")}: ' + ' · '.join(f'{k} 「{r.get(k, "").strip()}」' for k in HAND4 if r.get(k, '').strip())
                for h, cols, rows in H.md_tables(text) if h.startswith('4.') and any(k in cols for k in HAND4)
                for r in rows if any(r.get(k, '').strip() for k in HAND4)]
        if hand and not getattr(args, 'force', False):   # 4절 손 칸은 auto 블록 안이라 다시 쓰면 지워짐 → 덮지 않고 멈춤
            raise H.Stop('예외', '손칸보호', f'{OF} 4절 손 칸 (Staffing 변화 · 트리거 지표) {len(hand)}행이 채워져 있어 run 이 덮지 않음',
                         [ws.rel(p) + ' 「4. 시나리오」'] + hand[:6],
                         ['대안 · 경로만 고쳤으면 run 대신 `step4.py check` 만 다시', '앞 단계 숫자가 바뀌어 다시 계산해야 하면 사람 확인 뒤 `--force` 로 다시 run (4절 손 칸이 비워짐 · Gap 담당이 다시 채움)', '보류'])
        old = (H.auto_block(text, 'gap'), H.auto_block(text, 'skill'))
        if old[0] is not None and old != (auto1.strip(), auto5.strip()):   # 숫자가 바뀌면 지난 CP4 답은 새 숫자의 승인이 아님 → 답 칸을 비워 다시 CP4
            text, reset = H.clear_approval(text, 'CP4')
    else:
        alt = [[c, round(gap[c], 1), next(r[7] for r in rowsG if r[0] == c), '', '', '', '', '', '', '']
               for c in sorted(gap, key=lambda k: -gap[k]) if abs(gap[c]) > 0.5]
        red = [[c, round(v, 1), '', '', '', '', ''] for c, v in sorted(surplus.items(), key=lambda x: -x[1])]
        text = '\n'.join([f'# ④ Gap · 대안 ({ws.co} · {ty})', '',
                          f'- 만든 스크립트: `{SCRIPT} run` (1 · 4 · 5절 숫자) · 2 · 3절은 **Gap 담당이 채운다** (스크립트가 지우지 않음)',
                          '- **CP4 Gap 대안 승인**: 「## 승인」 표가 「확정」이면 2절 모든 행이 승인된 것으로 본다 (범위 전체). 행마다 다르게 정할 때만 2절 「승인」 칸에 그 행 답 (추가 확인 · 보류 · 이관)과 이유를 적는다 · 빈칸 = CP4 답을 따름 · CP4 확정 전에는 리포트를 쓰지 않는다',
                          '- 감원 대안을 단독으로 내지 않는다. 여유 RPU 는 재배치 경로를 먼저 붙인다 (파일 1 인력 운영 방향)', '',
                          '<!-- auto:start gap -->', '<!-- auto:end gap -->', '',
                          '## 2. 대안표 (Gap 담당)', '',
                          '- 4B = Buy (채용) · Build (육성 · 이동 · 승진) · Borrow (고용형태 변경 · 외부 활용) · Bot (AI · 자동화) + 구조 대안 (직무 · 조직 재설계)',
                          '- 조합 규칙 · 우선순위 기준: `wfp-step4-gap` SKILL.md 「조합 규칙」 「우선순위」 줄을 따른다 (자사 변형은 SKILL 한 곳만 고침) · 근거 = ①_사업전략분석.md 요인 번호 또는 파일 1 절', '',
                          H.md_table(['RPU 코드', 'Gap', '유형', '4B 대안', '시점', '근거 (요인 번호)', '경로', '우선순위', '2×2', '승인'], alt), '',
                          '## 3. 재배치 경로 (Gap 담당)', '',
                          '- 5절 「인접 스킬 재배치 후보」 또는 신설 조직에서 받는 RPU 를 고르고, 인원 · 교육 기간 · 시점을 적는다 (교육 · 시점 칸이 비면 검증에서 「확인 필요」)',
                          '- 인원 = 행마다 사람 여유 표시값 (소수 1자리)을 반올림한 정수 · 표 아래 「정수 처리」 줄에 행 합과 여유 합의 차이 · 근거를 적는다', '',
                          H.md_table(['보내는 RPU', '사람 여유', '인원', '받는 RPU', '교육', '전환', '시점'], red), '',
                          '<!-- auto:start scenario -->', '<!-- auto:end scenario -->', '',
                          '<!-- auto:start skill -->', '<!-- auto:end skill -->', '',
                          H.approval_block('CP4', '인사기획팀장 + R 행 소관 본부장')])
    text = H.replace_auto(text, 'gap', auto1)
    text = H.replace_auto(text, 'scenario', auto4)
    text = H.replace_auto(text, 'skill', auto5)
    H.write_md(p, text)
    H.side_outputs(ws, 'gap')   # ④_Gap분석.xlsx (보고용 · 승인 원본은 md)
    print(json.dumps(dict(status='완료', Gap=round(tot(gap), 1), Gap상승=round(tot(gapu), 1), 여유RPU=list(surplus),
                          여유합=round(sum(surplus.values()), 1), 재배치정수합=redep_n, 재배치후보=len(adj_rows), 신설조직=sorted(events), file=ws.rel(p),
                          next='Gap 담당이 2·3절을 채운 뒤 check 실행',
                          **({'CP4': f'숫자가 바뀌어 CP4 승인 칸을 비움 (지난 답 「{reset}」 은 결정로그에 · check 뒤 CP4 다시)'} if reset else {})), ensure_ascii=False, indent=1))


def check(ws, args):
    p = ws.o(OF)
    if not os.path.exists(p):
        raise H.Stop('예외', '입력없음', f'{OF} 가 없음', [ws.rel(p)], ['run 먼저'])
    H.side_outputs(ws, 'gap')   # Gap 담당이 채운 2 · 3절을 보고용 엑셀에 반영
    T = {h.split('.')[0]: rows for h, cols, rows in H.md_tables(H.read(p)) if re.match(r'^\d\.', h)}
    alt, red, gaps = T.get('2', []), T.get('3', []), T.get('1', [])
    probs = []
    for r in alt:
        v = r.get('4B 대안', '')
        if not v.strip():
            probs.append(f'{r["RPU 코드"]}: 4B 대안 빈칸')
            continue
        if CUT.search(v) and not any(k in v for k in FB + ('재배치',)):
            probs.append(f'{r["RPU 코드"]}: 감원 대안 단독 ({v}) · 금지')
        if not any(k in v for k in FB + ('구조', '재배치', '추가 확인', '유지')):
            probs.append(f'{r["RPU 코드"]}: 4B 값이 아님 ({v})')
        if not r.get('근거 (요인 번호)', '').strip():
            probs.append(f'{r["RPU 코드"]}: 근거 빈칸')
    surplus = [r['RPU 코드'] for r in gaps if r.get('방향') == '여유']
    sent = {r.get('보내는 RPU', '').strip(): r for r in red if r.get('받는 RPU', '').strip()}
    for c in surplus:
        if c not in sent:
            probs.append(f'{c}: 여유 RPU 인데 재배치 경로 (받는 RPU) 가 없음')
    other, bad = H.row_answers(H.read(p))
    for c, v in bad:
        probs.append(f'{c}: 「승인」 칸 값이 답 4종 밖 ({v}) · 빈칸 (CP4 답을 따름) 또는 추가 확인 · 보류 · 이관 + 이유')
    if probs:
        raise H.Stop('예외', '대안점검', f'대안표 점검에서 {len(probs)}건이 걸림 → Gap 담당이 고친 뒤 check 다시', probs[:12],
                     ['Gap 담당이 고침', '사람이 이 행을 그대로 두기로 정함 (결정로그 「예외 · 대안점검」에 이유)'])
    a = H.approval(p, 'CP4')
    if not a['ok']:
        who = str((H.config().get('approvers') or {}).get('CP4') or '').strip() or CP4_WHO
        raise H.Stop('CP', 'CP4', 'Gap 대안 승인 (범위 전체 · 감원 단독 대안 없음 확인 · 행마다 다른 답은 2절 「승인」 칸)',
                     [ws.rel(p), f'대안 {len(alt)}행 · 재배치 {len(sent)}행 · 점검 통과',
                      f'승인자: {who}' + (' (받는 RPU 의 본부) · 나머지 행은 Y · G 계획 승인으로 이유 칸에 기록 (기본값 · 자사 예 · 가설 · config.json approvers 로 바꿈)' if who == CP4_WHO else ' (config.json approvers · 자사 승인자)')]
                     + ([f'2절 「승인」 칸에 다른 답 {len(other)}행: ' + ' · '.join(f'{c} {v}' for c, v in other[:6])] if other else []),
                     ['확정 (범위 전체 · 고친 행은 이유 칸에 내용)', '추가 확인 (행 지정 · 2절 「승인」 칸)', '보류 · 이관'])
    print(json.dumps(dict(status='완료', note='점검 통과 · CP4 승인됨 (범위 전체' + (f' · 다른 답 {len(other)}행' if other else '') + ') · 리포트 단계로',
                          file=ws.rel(p)), ensure_ascii=False))


def main():
    ap = H.common_args('④ Gap 분석 스크립트')
    ap.add_argument('cmd', choices=['run', 'check'])
    ap.add_argument('--force', action='store_true', help='run: 4절 손 칸 (Staffing 변화 · 트리거 지표)이 채워져 있어도 다시 씀 (사람 확인 뒤만)')
    return ap.parse_args()


if __name__ == '__main__':
    args = main()
    H.run_main({'run': run, 'check': check}[args.cmd], lambda: args)
