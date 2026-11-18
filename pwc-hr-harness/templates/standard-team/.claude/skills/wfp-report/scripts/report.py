# -*- coding: utf-8 -*-
"""인력계획 리포트 7섹션 조립 · 요약 · 수요 · 공급 · Gap · 대응 · 가정 · 검증 로그.

사용:
  python -X utf8 .claude/skills/wfp-report/scripts/report.py build --co K-co
  python -X utf8 .claude/skills/wfp-report/scripts/report.py confirm --co K-co   # CP5 「확정」을 적은 뒤 (멈춤 파일을 닫음 · 리포트는 다시 쓰지 않음)

- CP4 승인 확인 → 리포트_<co>_<연도>.md 의 숫자 칸(auto 블록)을 결과 파일에서 옮겨 쓴다 (계산하지 않음)
- 「요약」 핵심 메시지 문장은 리드가 쓰는 칸이라 지우지 않는다
- 끝에 CP5 리포트 공유 승인 표 → 승인 전에는 리포트를 폴더 밖으로 내보내지 않는다
- 다시 build 할 때 숫자 칸 (kpi · demand · supply · gap · resp)이 바뀌면 CP5 답 칸을 비운다 (지난 답은 결정로그 · 새 숫자는 다시 CP5)
"""
import os, sys, json, re
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', '..', 'hr-lib'))
import hrlib as H

SCRIPT = 'wfp-report/scripts/report.py'
NUM_BLOCKS = ('kpi', 'demand', 'supply', 'gap', 'resp')   # 숫자 칸 (바뀌면 CP5 다시)


def a_range_row(hcroi, reg_base, f1):
    """A 범위 행: 낮은 값 ~ 높은 값 순서 · 라벨도 같은 순서 (회사마다 HCROI 역산과 회귀 기본 가운데 작은 쪽이 다름)."""
    if hcroi is not None and reg_base is not None and hcroi > reg_base:
        return ['A 범위 (회귀 기본 ~ HCROI 역산)', f'{f1(reg_base)} ~ {f1(hcroi)}', '②A_매출이익.xlsx']
    return ['A 범위 (HCROI 역산 ~ 회귀 기본)', f'{f1(hcroi)} ~ {f1(reg_base)}', '②A_매출이익.xlsx']


def _md_section(path, start, stop=None):
    t = H.read(path)
    if start not in t:
        return ''
    s = t[t.index(start):]
    if stop and stop in s[len(start):]:
        s = s[:len(start) + s[len(start):].index(stop)]
    return s.strip()


def build(ws, args):
    H.require_cp(ws, 'CP4', '④_Gap대안.md', '리포트 조립')
    ty = ws.cfg['target_year']
    need = ['②A_매출이익.xlsx', '②_대사.xlsx', '③_공급.xlsx', '④_Gap대안.md']
    miss = [f for f in need if not os.path.exists(ws.o(f))]
    if miss:
        raise H.Stop('예외', '입력없음', '리포트에 옮길 결과 파일이 없음', miss, ['앞 단계 실행'])
    _, A = H.read_xlsx_table(ws.o('②A_매출이익.xlsx'))
    _, D = H.read_xlsx_table(ws.o('②_대사.xlsx'))
    _, S = H.read_xlsx_table(ws.o('③_공급.xlsx'), '기본')
    _, SU = H.read_xlsx_table(ws.o('③_공급.xlsx'), '상승')
    tot = next(r for r in D if r['RPU 코드'] == '합계')
    stot = next(r for r in S if r['RPU 코드'] == '합계')
    sutot = next(r for r in SU if r['RPU 코드'] == '합계')
    gtxt = H.read(ws.o('④_Gap대안.md'))
    g = re.search(r'\*\*Gap (-?[\d.]+)\*\*', gtxt)
    gu = re.search(r'\(상승 (-?[\d.]+)\)', gtxt)
    gap = float(g.group(1)) if g else None
    hc = stot['현원']
    reg = {r['시나리오']: r[f'{ty} 인원'] for r in A if str(r['길']).startswith('회귀')}
    hcroi = next((r[f'{ty} 인원'] for r in A if str(r['길']).startswith('HCROI')), None)
    f1 = lambda x: '-' if x is None else f'{x:,.1f}'
    kpi = H.md_table(['지표', '값', '출처'], [
        ['현원 (기준일)', f'{hc:,.0f}명', '③_공급.xlsx 기본 · 합계 행 (파일 2 RPU 행 합)'],
        [f'결정 수요 사람 FTE ({ty})', f1(tot['결정 수요 사람']), '②_대사.xlsx · 합계 행'],
        [f'디지털 FTE ({ty} · 별도 줄)', f1(tot['결정 수요 디지털']), '②_대사.xlsx · 합계 행'],
        [f'공급 {ty} (채용 없음 · 기본 / 상승)', f'{f1(stot[f"공급 {ty}"])} / {f1(sutot[f"공급 {ty}"])}', '③_공급.xlsx'],
        [f'Gap {ty} (기본 / 공급 상승)', f'{f1(gap)} / {gu.group(1) if gu else "-"}', '④_Gap대안.md 1절'],
        a_range_row(hcroi, reg.get('기본'), f1),
    ])
    rec_rows = [[r['RPU 코드'], f1(r['B']) if isinstance(r['B'], (int, float)) else r['B'], f1(r['C']) if isinstance(r['C'], (int, float)) else r['C'],
                 r['차이'], r['원인 판정'], r['판정'], f1(r['결정 수요 사람']), f1(r['결정 수요 디지털'])] for r in D if r['RPU 코드'] != '합계']
    sup_rows = [[r['RPU 코드'], r['현원'], f'{r["퇴직률"]:.1%}' if isinstance(r['퇴직률'], float) else r['퇴직률'], f1(r[f'공급 {ty}']), r['주의'] or ''] for r in S if r['RPU 코드'] != '합계']
    alt = _md_section(ws.o('④_Gap대안.md'), '## 2. 대안표', '## 3.')
    red = _md_section(ws.o('④_Gap대안.md'), '## 3. 재배치 경로', '<!--')
    gap_tbl = _md_section(ws.o('④_Gap대안.md'), '## 1. Gap 표', '<!--')
    assume = []
    for f, cp in (('①_열대응표.md', 'CP1'), ('②_수요파라미터.md', 'CP2'), ('③_공급가정.md', 'CP3'), ('④_Gap대안.md', 'CP4')):
        a = H.approval(ws.o(f), cp)
        assume.append([cp, f, a['답'], a['승인자'], a['승인일']])
    pf = H.read(ws.o('③_공급가정.md'))
    s_rows = next((rows for h, c, rows in H.md_tables(pf) if h.startswith('2. 가정')), [])
    dlog = ws.o('결정로그.md')
    vlog = ws.o('검증로그.md')
    vt = H.read(vlog) if os.path.exists(vlog) else ''
    v_open = len(re.findall(r'\| 미해결 \|', vt))
    v_wait = len(re.findall(r'\| 미확인 \|', vt))   # 범위 밖 거절 행동 테스트를 아직 묻지 않은 행
    v_acc = len(re.findall(r'\| 인정 \(결정로그 #\d+\) \|', vt))   # 관문 인정 (사람 답 · 결정로그) · 미해결로 세지 않음
    v_all = len(re.findall(r'^\| \d+ \|', vt, re.M))

    auto = {
        'kpi': '\n'.join(['### 핵심 숫자 (결과 파일에서 옮김)', '', kpi]),
        'demand': '\n'.join(['## 2. 수요 · 결정 수요 (방식 A · B · C 대사)', '',
                             f'- 방식 A (매출·이익): 회귀 보수 {f1(reg.get("보수"))} · 기본 {f1(reg.get("기본"))} · 낙관 {f1(reg.get("낙관"))} · HCROI 역산 {f1(hcroi)} (출처 ②A_매출이익.xlsx)',
                             '- 결정 수요는 세 값을 평균 내지 않고 원인 판정으로 채택했다 · 디지털 FTE 는 사람 FTE 와 별도 열', '',
                             H.md_table(['RPU', 'B', 'C', '차이', '원인 판정', '판정', '결정 수요 사람', '디지털'], rec_rows)]),
        'supply': '\n'.join(['## 3. 공급 (채용 없음)', '',
                             f'- 기본 {f1(stot[f"공급 {ty}"])} · 상승 {f1(sutot[f"공급 {ty}"])} (출처 ③_공급.xlsx) · 신규 채용 · 신설 조직 · 디지털 FTE 는 넣지 않음', '',
                             H.md_table(['RPU', '현원', '퇴직률', f'공급 {ty}', '주의'], sup_rows)]),
        'gap': '\n'.join(['## 4. Gap', '', gap_tbl.replace('## 1. Gap 표 (스크립트)', '').strip()]),
        'resp': '\n'.join(['## 5. 대응 (4B · 재배치 · CP4 승인 내용)', '', alt.replace('## 2. 대안표 (Gap 담당)', '### 대안표').strip(), '',
                           red.replace('## 3. 재배치 경로 (Gap 담당)', '### 재배치 경로').strip()]),
        'assume': '\n'.join(['## 6. 가정 · 결정 기록', '', '### 체크포인트 승인', '', H.md_table(['CP', '자료', '답', '승인자 (직책)', '승인일'], assume), '',
                             '### 공급 가정 (③_공급가정.md)', '', H.md_table(['번호', '가정', '사람 답'], [[r['번호'], r['가정'], r['사람 답']] for r in s_rows]), '',
                             f'- 결정로그: `{ws.rel(dlog)}` · 산정 가정: 파일 4 산정 가정 시트 · 수요 파라미터: `②_수요파라미터.md`']),
        'verify': '\n'.join(['## 7. 검증 로그', '', f'- 검증 행 {v_all}개 · 미해결 {v_open}개 · 미확인 {v_wait}개 · 인정 {v_acc}개 (출처 `{ws.rel(vlog)}`)',
                             '- 미해결 · 미확인이 0이 아니면 CP5 에서 공유를 승인하지 않는다' if (v_open or v_wait) else f'- 미해결 0건 · 미확인 0건 · 인정 {v_acc}건',
                             *([f'- 인정 {v_acc}건 = 사람이 관문에서 「관문 인정 (결정로그에 이유)」로 답한 행 (값은 고치지 않음 · 이유는 결정로그 · 미해결로 세지 않음 · CP5 공유 승인을 막지 않음)'] if v_acc else [])]),
    }
    name = f'리포트_{ws.co}_{ty}.md'
    p = ws.o(name)
    reset = ''
    if os.path.exists(p):
        text = H.read(p)
        old = [H.auto_block(text, k) for k in NUM_BLOCKS]
        if any(o is not None for o in old) and old != [auto[k].strip() if k in auto else None for k in NUM_BLOCKS]:
            text, reset = H.clear_approval(text, 'CP5')   # 숫자가 바뀐 리포트는 다시 공유 승인
    else:
        text = '\n'.join([f'# {ws.co} {ty} 인력계획 리포트', '',
                          f'- 만든 스크립트: `{SCRIPT} build` (숫자 칸) · 요약 문장은 리드가 씀 · 숫자는 결과 파일에서 옮기기만 하고 다시 계산하지 않는다',
                          '- **CP5 리포트 공유**: 승인 전에는 이 리포트를 폴더 밖으로 내보내지 않는다 (메일 · 메신저 · 공유 폴더 복사 포함)', '',
                          '## 1. 요약', '',
                          '### 핵심 메시지 (리드가 씀 · 10줄 이내 · 문장마다 출처)', '',
                          '- (목적 「왜」에 맞춘 첫 문장 · 예: 핵심 Gap 과 확보 경로)', '',
                          '<!-- auto:start kpi -->', '<!-- auto:end kpi -->', '',
                          *[f'<!-- auto:start {k} -->\n<!-- auto:end {k} -->\n' for k in ('demand', 'supply', 'gap', 'resp', 'assume', 'verify')],
                          H.approval_block('CP5', 'CHRO')])
    for k, v in auto.items():
        text = H.replace_auto(text, k, v)
    vi = H.virtual_info(ws.data)
    if vi['virtual']:   # 가상 데이터면 첫 줄에 늘 표시 (지워져 있어도 다시 붙임 · 검증이 첫 줄을 확인)
        body = [ln for ln in text.splitlines() if H.VIRTUAL_LINE not in ln]
        text = '\n'.join([f'> **{H.VIRTUAL_LINE}** · 근거 `{ws.rel(os.path.join(ws.data, H.VIRTUAL_MARK))}`'] + body)
    H.write_md(p, text)
    H.side_outputs(ws, 'gap', 'docx')   # 리포트 워드판 · ④_Gap분석.xlsx (보고용 · CP5 전에는 「초안」 표시)
    a = H.approval(p, 'CP5')
    if not a['ok']:
        raise H.Stop('CP', 'CP5', '리포트 공유 승인 (7섹션 · 결정로그 · 검증로그 미해결 확인)',
                     [ws.rel(p), f'검증 미해결 {v_open}건 · 미확인 {v_wait}건 · 인정 {v_acc}건 (관문 인정 · 결정로그)'] + ([f'{H.VIRTUAL_LINE} (공유 승인 = 교육 안에서만)'] if vi['virtual'] else []) + ([f'숫자 칸이 바뀌어 CP5 답 칸을 비움 (지난 답 「{reset}」 · 결정로그에 남음)'] if reset else []), ['확정 (공유 허용)', '추가 확인 (요약 문장 · 숫자 출처를 고친 뒤 다시)', '보류 · 이관'])
    print(json.dumps(dict(status='완료', note='CP5 승인됨 · 공유 가능', file=ws.rel(p)), ensure_ascii=False))


def confirm(ws, args):
    """CP5 「확정」을 리포트 승인 칸에 적은 뒤: 승인을 확인하고 지난 CP5 멈춤 파일을 닫는다 (리포트 내용은 다시 쓰지 않음)."""
    p = ws.o(f'리포트_{ws.co}_{ws.cfg["target_year"]}.md')
    if not os.path.exists(p):
        raise H.Stop('예외', '입력없음', '리포트 파일이 없음', [ws.rel(p)], ['report.py build 먼저'])
    a = H.approval(p, 'CP5')
    H.side_outputs(ws, 'docx')   # 승인 상태를 워드 표지 · 꼬리말에 반영
    if not a['ok']:
        raise H.Stop('CP', 'CP5', '리포트 공유 승인 (7섹션 · 결정로그 · 검증로그 미해결 확인)',
                     [ws.rel(p), f'답={a["답"] or "(빈칸)"} · 승인자={a["승인자"] or "(빈칸)"}'],
                     ['확정 (공유 허용)', '추가 확인 (요약 문장 · 숫자 출처를 고친 뒤 다시)', '보류 · 이관'])
    print(json.dumps(dict(status='완료', note=f'CP5 승인 확인 ({a["승인자"]} · {a["승인일"]}) · 멈춤 파일 닫음 · 공유 가능', file=ws.rel(p)), ensure_ascii=False))


def main():
    ap = H.common_args('리포트 조립 스크립트')
    ap.add_argument('cmd', choices=['build', 'confirm'])
    return ap.parse_args()


if __name__ == '__main__':
    args = main()
    H.run_main({'build': build, 'confirm': confirm}[args.cmd], lambda: args)
