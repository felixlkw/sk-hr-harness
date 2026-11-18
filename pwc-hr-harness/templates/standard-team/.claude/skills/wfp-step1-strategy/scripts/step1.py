# -*- coding: utf-8 -*-
"""① 사업전략 · 요구사항 · 열 대응표 · 데이터 점검 · RPU 목록 · Pivotal Role 보조표.

사용:
  python -X utf8 .claude/skills/wfp-step1-strategy/scripts/step1.py scan --co K-co
  python -X utf8 .claude/skills/wfp-step1-strategy/scripts/step1.py rpu  --co K-co

scan : data/<co>/ 의 모든 xlsx 시트를 기본 표(K-Co. 양식)에 대응시켜 ①_열대응표.md/.json 을 만들고 CP1 에서 멈춘다.
       기준일이 서로 다른 시점 시트 · 소계/합계 불일치는 「예외」로 먼저 멈춘다.
rpu  : CP1 승인 뒤, ①_사업전략분석.md 에 붙일 RPU 목록 · 유형 · 스킬 요구 수준(Pivotal Role 판정 참고 자료)을 표로 낸다.
"""
import os, sys, glob, json, re
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', '..', 'hr-lib'))
import hrlib as H
import openpyxl

SCRIPT = 'wfp-step1-strategy/scripts/step1.py'


def _sheets(ws):
    for f in sorted(glob.glob(os.path.join(ws.data, '*.xlsx'))):
        if os.path.basename(f).startswith('~$'):
            continue
        wb = openpyxl.load_workbook(f, data_only=True)
        for sh in wb.worksheets:
            yield os.path.basename(f), sh


def _synonyms(ws):
    """자사 열 이름 → 기본 열 이름 사전 (선택). data/<co>/열이름_대응.json 에 사람이 적어 둘 수 있다."""
    p = os.path.join(ws.data, '열이름_대응.json')
    if os.path.exists(p):
        with open(p, encoding='utf-8') as f:
            return json.load(f)
    return {}


def scan(ws, args):
    T = H.tables(ws.cfg)
    pj = ws.o('①_열대응표.json')
    if os.path.exists(pj):
        with open(pj, encoding='utf-8') as f:
            ws._prev_map = json.load(f)
    syn = _synonyms(ws)
    made = H.ensure_logs(ws)   # 새 회사 폴더: 결정로그 · 검증로그를 양식에서 만듦 (있으면 그대로)
    cands, seen = [], []
    for fname, sh in _sheets(ws):
        seen.append((fname, sh))
        if fname.startswith(H.EXTRA_PREFIX):   # 추가_ 파일은 기본 표에 대응시키지 않는다 (추가 자료로만 보고)
            continue
        for key, spec in T.items():
            want = spec['cols']
            best, n = H.find_header(sh, list(dict.fromkeys([syn.get(w, w) for w in want] + want)))
            if not best:
                continue
            hdr, names = best
            nonempty = [x for x in names if x]
            title = sh.title + ' ' + H.norm(sh.cell(1, 1).value)
            tsim = 1 if spec['title'].replace(' ', '') in title.replace(' ', '') else 0
            cands.append(((n, tsim, n / max(1, len(nonempty))), key, fname, sh, hdr, names))
    chosen, used = {}, set()
    for score, key, fname, sh, hdr, names in sorted(cands, key=lambda c: c[0], reverse=True):
        if key in chosen or (fname, sh.title) in used or score[0] < 2:
            continue
        chosen[key] = (fname, sh, hdr, names)
        used.add((fname, sh.title))
    extra = _extras(ws, seen, used)

    m = dict(co=ws.co, data=ws.rel(ws.data), tables={}, basis={})
    if extra:   # 추가 자료가 있을 때만 (K-Co. 처럼 없으면 키도 없음)
        m['extra'] = extra
    rows_md, missing_req = [], []
    for key, spec in T.items():
        if key not in chosen:
            m['tables'][key] = dict(file=None)
            mark = '없음 (선택 표)' if spec.get('optional') else '**없음 · 대체 필요**'
            if not spec.get('optional'):
                missing_req.append(key)
            rows_md.append([spec['no'], spec['title'], '-', '-', '-', '-', 0, mark, '없을 때 대체 표 참조 (customization-guide)'])
            continue
        fname, sh, hdr, names = chosen[key]
        basis = H.sheet_basis(sh)
        cols = {}
        for c in spec['cols']:
            src = syn.get(c, c)
            cols[c] = src if src in names else (c if c in names else None)
        m['tables'][key] = dict(file=fname, sheet=sh.title, header_row=hdr, cols=cols, basis=basis)
        same_n = sum(1 for c, src in cols.items() if src == c)
        diff = [f'{c} ← {src}' for c, src in cols.items() if src and src != c]
        miss = [c for c, src in cols.items() if not src]
        st = '일치' if not diff and not miss else ('**없는 열**' if miss else '대응 (이름 다름)')
        note = ' · '.join(diff + [f'**{c} 없음** (대체 · 추정 · 제외 중 사람이 정함)' for c in miss]) or '-'
        unit = H.sheet_unit(sh)
        rows_md.append([spec['no'], spec['title'], f'{fname} · {sh.title}', hdr, basis or '(안내문에 없음)', unit or '**확인 필요**',
                        f'{same_n}/{len(cols)}', st, note])

    # --- 예외 1: 같은 파일 안 시점 시트의 기준일이 서로 다름 (공통 규칙 2)
    by_file = {}
    for key, t in m['tables'].items():
        if t.get('file') and T[key].get('point') and t.get('basis'):
            by_file.setdefault(t['file'], []).append((key, t['sheet'], t['basis']))
    for fname, lst in by_file.items():
        bases = {b for _, _, b in lst}
        if len(bases) > 1 and not H.allowed(args, '기준일'):
            ev = [f'{fname} · {s} · 안내문 기준 「{b}」' for _, s, b in lst]
            raise H.Stop('예외', '기준일', f'{fname} 안 시점 시트의 기준일이 서로 다름 → 어느 기준일로 RPU 현원·보유 분포를 확정할지',
                         ev, ['기준일이 맞는 파일로 교체 받기 (권장)', '시트별 기준일을 열로 두고 따로 씀 (한 표로 합치지 않음)',
                              '보류 · 해당 단계 생략'],
                         note='공통 규칙 2: 기준일이 다른 값은 한 표로 합치지 않는다. RPU 목록은 확정하지 않았다.')

    # --- 예외 2: 집계 행(소계·합계)과 RPU 행 합이 다름 (공통 규칙 3 · 파일 사이 같아야 할 값)
    H.save_map(ws, m)
    if m['tables']['org'].get('file'):
        bad = [(n, v, s, cell) for n, v, s, cell in H.aggregates(ws, 'org', m) if abs((v or 0) - s) > 0.5]
        if bad and not H.allowed(args, '합계'):
            raise H.Stop('예외', '합계', '조직·인원 현황의 소계·합계가 RPU 행 합과 다름 → 어느 값을 현원으로 쓸지',
                         [f'{n}: 파일 값 {v} · RPU 행 합 {s:g} · 셀 {cell}' for n, v, s, cell in bad],
                         ['RPU 행 합을 쓰고 집계 행 오류를 자료 담당에게 알림 (권장)', '자료를 고쳐 다시 받음', '보류'],
                         note='공통 규칙 3: 집계 행은 데이터 행이 아니다. 모든 계산은 RPU 행에서 한다.')

    # --- 교차 점검 (멈추지 않고 표에 적음 · CP1 에서 사람이 봄)
    checks = _cross_checks(ws, m)
    vi = H.virtual_info(ws.data)

    text = [f'# ① 열 대응표 ({ws.co})', '',
            f'- 만든 스크립트: `{SCRIPT} scan` · 입력 폴더: `{ws.rel(ws.data)}`',
            '- 기본 열 = K-Co. 양식 열 이름 (데이터 사전 3절). 자사 파일은 이 표로 기본 열에 맞춰 읽는다',
            '- **CP1 열 대응 확인**: 아래 표를 확인하고 「## 승인」 표를 채우기 전에는 수요 · 공급을 계산하지 않는다',
            '- 이름이 다른 열은 `data/<회사>/열이름_대응.json` 에 {"기본 열": "자사 열"} 로 적고 scan 을 다시 돌린다',
            '- 「단위 · 기준일」이 「확인 필요 · 안내문에 없음」인 표는 CP1 에서 사람이 단위 · 기준일을 확인해 이유 칸에 적는다 (공통 규칙 1 · 2)', '',
            *([f'> **{H.VIRTUAL_LINE}** · 이 회사 폴더는 가상 데이터다 (`{ws.rel(os.path.join(ws.data, H.VIRTUAL_MARK))}` · 가상 표시 xlsx {len(vi["stamped"])}개)', ''] if vi['virtual'] else []),
            '## 1. 기본 열 대응', '',
            H.md_table(['파일', '기본 표', '자사 파일 · 시트', '머리 행', '기준일', '단위 (안내문)', '같은 이름 열', '상태', '다른 이름 · 없는 열'], rows_md), '',
            '- 기본 열 목록: ' + ' / '.join(f"{v['title']}: " + ' · '.join(v['cols']) for v in T.values()), '',
            '## 2. 교차 점검 (파일 사이 같아야 하는 값)', '',
            H.md_table(['점검', '결과', '근거'], checks) if checks else '- 점검할 표가 없음', '']
    if missing_req:
        text += ['## 3. 없는 필수 표', '', *[f'- {T[k]["title"]} (파일 {T[k]["no"]}): 이 표를 쓰는 방식은 생략하거나 대체 자료를 받는다 (`references` 없을 때 대체 표)' for k in missing_req], '']
    if extra:
        text += [f'## {4 if missing_req else 3}. {H.EXTRA_TITLE}', '',
                 '- 기본 6종 표에 대응시키지 않은 자료다 (`추가_` 로 시작하는 파일은 늘 여기 · 기본 표에 맞는 시트가 하나도 없는 파일도 여기). 수요 · 공급 · Gap 숫자 계산에 쓰지 않는다',
                 '- 담당은 판단 칸 (Pivotal 근거 · 4B 대안 근거 등)에 「회사 기준 · <파일명>」으로 표기해 인용할 수 있다. 계산에 넣으려면 하네스 개정 (새 계산 스크립트 · 회귀 시험)이 필요하다',
                 '- CP1 에서 사람이 「이 자료를 근거로 써도 되는가」 (기준일 · 단위 · 개인 식별 열 없음)를 함께 확인한다', '',
                 H.md_table(['파일', '시트', '머리 행', '열 (최대 12)', '단위 (안내문)', '왜 추가 자료'],
                            [[e['file'], e['sheet'] or '-', e['header_row'] or '-', ' · '.join(e['cols']) or '(열 없음)', e['unit'] or '-', e['why']] for e in extra]), '']
    # 이미 승인된 대응표와 같으면 승인 기록을 그대로 둔다 (재실행 일관성)
    old = H.approval(ws.o('①_열대응표.md'), 'CP1')
    prev = getattr(ws, '_prev_map', None)
    same = old['ok'] and prev is not None and prev.get('tables') == m['tables']
    block = H.approval_block('CP1', '인사기획 실무 담당')
    if same:
        block = block.replace('| CP1 |  | 인사기획 실무 담당 |  |  |',
                              H.md_row(['CP1', old['답'], old['승인자'], old['승인일'], '대응표 변경 없음 · 기존 승인 유지']))
    text += [block]
    H.write_md(ws.o('①_열대응표.md'), '\n'.join(text))
    if same:
        print(json.dumps(dict(status='완료', note='대응표 변경 없음 · CP1 승인 유지', file=ws.rel(ws.o('①_열대응표.md')),
                              **({'extra': [f"{e['file']} · {e['sheet']}" for e in extra]} if extra else {})), ensure_ascii=False))
        return
    raise H.Stop('CP', 'CP1', '열 대응표 확인 (기본 열 ↔ 자사 열 · 단위 · 기준일 · 없는 열 처리)',
                 [ws.rel(ws.o('①_열대응표.md')), f'교차 점검 {len(checks)}건 · 없는 필수 표 {len(missing_req)}개']
                 + ([f'{H.EXTRA_TITLE} {len(extra)}건: ' + ' · '.join(f"{e['file']}" + (f" · {e['sheet']}" if e['sheet'] else '') for e in extra)] if extra else [])
                 + ([f'새 회사 폴더 · 양식에서 만든 로그: {" · ".join(made)}'] if made else []),
                 ['확정 (이대로 진행)', '추가 확인 (열 이름 대응을 고친 뒤 scan 다시)', '보류'])


def _extras(ws, seen, used):
    """추가 자료: 추가_ 로 시작하는 파일 (xlsx 시트 · csv · 그 밖의 파일) + 기본 표에 맞는 시트가 하나도 없는 xlsx 의 시트.
    기본 파일 안의 보조 시트 (예: 레벨 정의 · 유형 정의)는 넣지 않는다 (K-Co. 결과를 바꾸지 않음)."""
    out = []
    used_files = {f for f, _ in used}
    for fname, sh in seen:
        if fname.startswith(H.EXTRA_PREFIX) or fname not in used_files:
            hr, cols = H.guess_columns(sh)
            out.append(dict(file=fname, sheet=sh.title, header_row=hr, cols=cols, unit=H.sheet_unit(sh),
                            why='추가_ 파일 (기본 표에 대응시키지 않음)' if fname.startswith(H.EXTRA_PREFIX) else '기본 표에 맞는 시트 없음'))
    xl = {f for f, _ in seen}
    for p in sorted(glob.glob(os.path.join(ws.data, H.EXTRA_PREFIX + '*'))):
        fn = os.path.basename(p)
        if fn in xl or fn.startswith('~$') or os.path.isdir(p):
            continue
        cols = []
        if fn.lower().endswith('.csv'):
            import csv
            for enc in ('utf-8-sig', 'cp949'):
                try:
                    with open(p, encoding=enc, newline='') as f:
                        cols = [H.norm(c) for c in next(csv.reader(f), []) if H.norm(c)][:12]
                    break
                except (UnicodeDecodeError, StopIteration):
                    continue
        out.append(dict(file=fn, sheet=None, header_row=1 if cols else None, cols=cols, unit=None,
                        why='추가_ 파일 (' + (fn.rsplit('.', 1)[-1] if '.' in fn else '확장자 없음') + ')'))
    return out


def _cross_checks(ws, m):
    out = []
    try:
        org = {r['RPU 코드']: r for r in H.load(ws, 'org', m)}
        by = str(ws.cfg['base_year'])
        if m['tables']['hcs'].get('file'):
            hcs = {r['RPU 코드']: r for r in H.load(ws, 'hcs', m)}
            diff = [c for c in org if c in hcs and hcs[c].get(by) is not None and abs(hcs[c][by] - (org[c]['계'] or 0)) > 0.5]
            out.append(['현원(파일 2 계) = RPU별 연말 인원(파일 6) ' + by, '일치' if not diff else '불일치 ' + ' · '.join(diff), '파일 2 · 파일 6'])
        if m['tables']['hist'].get('file'):
            h = H.load(ws, 'hist', m)
            end = {}
            for r in h:
                if H.norm(r['연도']) == by and r.get('기말') is not None:
                    end[r['RPU 코드']] = end.get(r['RPU 코드'], 0) + r['기말']
            diff = [c for c in org if c in end and abs(end[c] - (org[c]['계'] or 0)) > 0.5]
            if end:
                out.append([f'현원 = 입퇴사 이력 {by} 기말 합', '일치' if not diff else '불일치 ' + ' · '.join(diff), '파일 2 · 파일 3'])
        if m['tables']['trend'].get('file'):
            tr = {H.norm(r['연도']): r for r in H.load(ws, 'trend', m)}
            if by in tr:
                s = sum((r['계'] or 0) for r in org.values())
                out.append([f'RPU 현원 합 = 연도별 추이 {by} 인원', '일치' if abs(s - tr[by]['인원(명)']) < 0.5 else f'불일치 ({s:g} · {tr[by]["인원(명)"]})', '파일 2 두 시트'])
        if m['tables']['age'].get('file'):
            age = {r['RPU 코드']: r for r in H.load(ws, 'age', m)}
            diff = [c for c in org if c in age and age[c].get('계') is not None and abs(age[c]['계'] - (org[c]['계'] or 0)) > 0.5]
            out.append(['현원 = 연령대 분포 계', '일치' if not diff else '불일치 ' + ' · '.join(diff), '파일 2 두 시트'])
    except H.Stop:
        pass
    except (KeyError, TypeError, ValueError) as e:   # 대응이 안 된 열이 있으면 점검을 건너뛰고 CP1 에서 보이게 함
        out.append(['교차 점검 일부 생략', f'**열 대응 확인 필요** (읽지 못한 열: {e})', '열이름_대응.json 을 채운 뒤 scan 다시'])
    return out


def rpu(ws, args):
    H.require_cp(ws, 'CP1', '①_열대응표.md', 'RPU 목록 정리')
    m = H.load_map(ws)
    org = H.load(ws, 'org', m)
    typ = {}
    if m['tables']['drv'].get('file'):
        for r in H.load(ws, 'drv', m):
            typ.setdefault(r['RPU 코드'], (H.norm(r.get('유형')), H.norm(r.get('하위 유형'))))
    sk = H.load(ws, 'skill', m, need=False)
    rows = []
    for r in org:
        c = r['RPU 코드']
        mine = [s for s in sk if s['RPU 코드'] == c]
        hi = [s for s in mine if (s.get('요구 레벨 2028') or 0) >= 3]
        gaps = []
        for s in mine:
            n = s.get('인원 계') or 0
            ok = s.get('2028 요구 충족 인원') or 0
            if n and (s.get('요구 레벨 2028') or 0) >= 2 and ok / n < 0.5:
                gaps.append(f"{s['스킬']} L{s['요구 레벨 2028']} 충족 {ok / n:.0%}")
        new_req = [s['스킬'] for s in mine if (s.get('요구 레벨 2028') or 0) - (s.get('요구 레벨 2025') or 0) >= 2] if mine and '요구 레벨 2025' in mine[0] else []
        t = typ.get(c, ('', ''))
        rows.append([H.norm(r['본부']), c, H.norm(r['RPU(산정 단위)']), r['계'], t[0], t[1],
                     len(hi), ' · '.join(gaps[:3]) or '-', r['_ref']])
    total = sum(r[3] for r in rows)
    text = [f'# ① RPU 목록 보조표 ({ws.co})', '',
            f'- 만든 스크립트: `{SCRIPT} rpu` · ①_사업전략분석.md 3절 · 4절에 옮겨 쓴다 (숫자는 그대로 옮기기만 함)',
            f'- RPU {len(rows)}개 · 현원 합 {total:g}명 · 출처 파일 2 조직·인원 현황 RPU 행 합', '',
            '## RPU 목록 · 유형 · 스킬 요구 (Pivotal Role 판정 참고 자료 · 기준 정의는 wfp-step1-strategy SKILL 「Pivotal Role 판정」)', '',
            H.md_table(['본부', 'RPU 코드', 'RPU', '현원', '유형', '하위 유형', '2028 요구 L3 이상 스킬 수',
                        '요구 충족 50% 미만 스킬 (최대 3)', '출처'], rows), '',
            '## ①_사업전략분석.md 5절 머리 (담당이 그대로 옮겨 쓰고 표를 채움)', '',
            '```', PIV_HEAD, '',
            PIV_NOTE, '',
            H.md_table(PIV_COLS, [['', '', '', '', '', '', '', '', '', '']]), '```']
    H.write_md(ws.o('①_RPU보조표.md'), '\n'.join(text))
    print(json.dumps(dict(status='완료', file=ws.rel(ws.o('①_RPU보조표.md')), rpu=len(rows), total=total,
                          next=f'①_사업전략분석.md 5절은 「{PIV_HEAD[3:]}」 머리 그대로 (보조표 끝 절) · 끝나면 리드가 ① 끝 보고에 이 표를 그대로 보여 줌'), ensure_ascii=False))


PIV_HEAD = '## 5. Pivotal Role 후보 (담당 판정 · 풀이 예)'
PIV_NOTE = ('- **담당 판정 (풀이 예)**: 기준 6개 (wfp-step1-strategy SKILL 「Pivotal Role 판정」)마다 자료가 있을 때만 ● · 기준마다 근거 (파일 · 시트 · 행 또는 절) · '
            '**문턱 ● 4개 이상만 후보** · 사람 확정 값이 아니다. 사람이 다르게 보면 그 기준의 근거 자료를 받아 다시 판정한다')
PIV_COLS = ['RPU', '1 전략 실행', '2 역량 변화', '3 충원 난이도', '4 학습곡선', '5 새 스킬', '6 경쟁 스킬', '기준별 근거', '● 수', '후보 (문턱 ● 4)']


def main():
    ap = H.common_args('① 사업전략 · 열 대응 스크립트')
    ap.add_argument('cmd', choices=['scan', 'rpu'])
    a = ap.parse_args()
    return a


if __name__ == '__main__':
    args = main()
    H.run_main({'scan': scan, 'rpu': rpu}[args.cmd], lambda: args)
