# -*- coding: utf-8 -*-
"""결과 파일을 사람이 읽는 보고서 모양으로 (2026-11-16 · 사용자 결정 · 숫자는 다시 계산하지 않음).

엑셀 (단계 스크립트가 write_xlsx 로 저장한 직후 decorate 가 붙임)
  「요약」 (보고서 한 장: 핵심 지표 · 핵심 해설 · 결과 읽는 법 · 판단할 것) → 「○○표」 (본부 소계 · 합계 · 강조)
  → 첨부 데이터 시트 (기존 이름 · 모양 그대로 · 다음 단계가 읽는 입력 계약 · 회색 탭) → 「산식·가정·출처」
  · 보고서 시트 = 「요약」 · 「산식·가정·출처」 · 이름이 「표」로 끝나는 시트 (is_report_sheet) · 첨부 데이터 시트는 2행이 「만든 스크립트」
  · 해설 문장은 첨부 데이터의 숫자로 짓는다 (같은 결과 = 같은 문장 · LLM 문장 아님)
④_Gap분석.xlsx (step4 run · check · report build 가 부름) · 리포트_<회사>_<연도>.docx (report build · confirm 이 부름)
내보내기: python report_xlsx.py export <outputs/회사> [파일 이름 | all] → outputs/<회사>/내보내기/ (읽기용 사본)

이 모듈이 실패해도 계산 · 결과 저장은 그대로 끝난다 (부르는 쪽이 경고만 냄).
"""
import datetime, glob, os, re, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.page import PageMargins

FONT = '맑은 고딕'
INK, SUB, LINE, ACC = '2D2926', '6B645D', 'CFC8BF', 'D04A02'
BAND, HEAD, GROUP, TOTAL, NOTE = '2D2926', 'F3EEE8', 'FBF7F2', 'FBE3D6', 'FFFCF8'
HL = dict(red='FCE4DF', amber='FFF2CC', green='E8F3E8', gray='EFEFEF', blue='E3EEF9')
thin = Side(style='thin', color=LINE)
mid = Side(style='medium', color=INK)
FMT = dict(int='#,##0', n1='#,##0.0', n2='#,##0.00', n3='0.000', n4='0.0000', pct='0.0%', spct='+0.0%;-0.0%;0.0%', s1='+#,##0.0;-#,##0.0;0.0')
BASIS = '산식·가정·출처'
SEC = '■ '      # 요약 시트 절 머리 (뷰어가 이 표시로 절을 나눔)


def is_report_sheet(name):
    return name in ('요약', BASIS) or name.endswith('표')


# ---------------------------------------------------------------------------
# 모양
# ---------------------------------------------------------------------------
def f(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(c):
    return PatternFill('solid', fgColor=c)


def _page(sh, landscape=True):
    sh.page_setup.orientation = 'landscape' if landscape else 'portrait'
    sh.page_setup.paperSize = sh.PAPERSIZE_A4
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
    sh.sheet_properties.pageSetUpPr.fitToPage = True
    sh.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    sh.oddFooter.center.text = '&P / &N'
    sh.oddFooter.right.text = '&A'
    sh.sheet_view.showGridLines = False
    sh.sheet_properties.tabColor = ACC


def _band(sh, title, sub, ncol):
    sh.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    c = sh.cell(1, 1, title); c.font = f(15, True, 'FFFFFF'); c.fill = fill(BAND)
    c.alignment = Alignment(vertical='center', indent=1)
    sh.row_dimensions[1].height = 30
    for j in range(2, ncol + 1):
        sh.cell(1, j).fill = fill(BAND)
    sh.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)
    c = sh.cell(2, 1, sub); c.font = f(9, color=SUB); c.alignment = Alignment(vertical='center', indent=1, wrap_text=True)
    for j in range(1, ncol + 1):
        sh.cell(2, j).border = Border(bottom=Side(style='medium', color=ACC))
    sh.row_dimensions[2].height = 18 if len(sub) < 110 else 30


def _sec(sh, r, text, ncol):
    sh.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
    c = sh.cell(r, 1, SEC + text); c.font = f(11, True, INK)
    for j in range(1, ncol + 1):
        sh.cell(r, j).border = Border(bottom=thin)
    sh.row_dimensions[r].height = 20
    return r + 1


def _para(sh, r, text, ncol, bullet='·', size=10, color=INK, bold=False, bg=None):
    sh.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
    c = sh.cell(r, 1, f'{bullet} {text}' if bullet else text)
    c.font = f(size, bold, color); c.alignment = Alignment(wrap_text=True, vertical='top', indent=1)
    if bg:
        for j in range(1, ncol + 1):
            sh.cell(r, j).fill = fill(bg)
    sh.row_dimensions[r].height = 16 * (len(text) // 62 + 1) + 2
    return r + 1


def summary_sheet(wb, s):
    """s = dict(title, sub, status, kpis[(이름, 값, 단위 · 한 줄)], messages[], read[], decide[], sheets[(이름, 설명)], refs[])"""
    sh = wb.create_sheet('요약')
    ncol = 8
    for j in range(1, ncol + 1):
        sh.column_dimensions[L(j)].width = 15.5
    _band(sh, s['title'], s['sub'], ncol)
    r = 4
    if s.get('status'):
        sh.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        c = sh.cell(r, 1, s['status']); c.font = f(10, True, ACC); c.alignment = Alignment(indent=1, vertical='center', wrap_text=True)
        for j in range(1, ncol + 1):
            sh.cell(r, j).fill = fill(TOTAL)
        sh.row_dimensions[r].height = 20 if len(s['status']) < 80 else 32
        r += 2
    for i, (name, val, note) in enumerate(s['kpis'][:4]):
        c0 = 1 + i * 2
        for rr in (r, r + 1, r + 2):
            sh.merge_cells(start_row=rr, start_column=c0, end_row=rr, end_column=c0 + 1)
        sh.cell(r, c0, name).font = f(9, True, SUB)
        b = sh.cell(r + 1, c0, val); b.font = f(19, True, ACC if i == 0 else INK)
        if isinstance(val, float):
            b.number_format = FMT['n1']
        elif isinstance(val, int):
            b.number_format = FMT['int']
        n = sh.cell(r + 2, c0, note); n.font = f(8.5, color=SUB)
        for rr in (r, r + 1, r + 2):
            for cc in (c0, c0 + 1):
                x = sh.cell(rr, cc); x.fill = fill(GROUP)
                x.border = Border(left=thin if cc == c0 else None, right=thin if cc == c0 + 1 else None,
                                  top=Side(style='medium', color=ACC if i == 0 else INK) if rr == r else None,
                                  bottom=thin if rr == r + 2 else None)
            sh.cell(rr, c0).alignment = Alignment(indent=1, vertical='center' if rr < r + 2 else 'top', wrap_text=rr == r + 2,
                                                  horizontal='left')
    sh.row_dimensions[r].height = 18; sh.row_dimensions[r + 1].height = 32; sh.row_dimensions[r + 2].height = 30
    r += 4
    r = _sec(sh, r, '핵심 해설', ncol)
    for i, m in enumerate(s['messages'], 1):
        r = _para(sh, r, m, ncol, bullet=f'{i}.')
    r += 1
    for key, title, bullet, bg in (('read', '결과 읽는 법', '·', None), ('decide', '사람이 판단할 것 · 확인할 것', '□', NOTE)):
        if s.get(key):
            r = _sec(sh, r, title, ncol)
            for m in s[key]:
                r = _para(sh, r, m, ncol, bullet=bullet, bg=bg)
            r += 1
    if s.get('sheets'):
        r = _sec(sh, r, '이 파일의 시트', ncol)
        for name, d in s['sheets']:
            r = _para(sh, r, f'「{name}」 {d}', ncol, size=9.5, color=SUB)
        r += 1
    if s.get('refs'):
        r = _sec(sh, r, '근거 · 다시 만드는 법', ncol)
        for m in s['refs']:
            r = _para(sh, r, m, ncol, size=9, color=SUB)
    _page(sh, landscape=False)
    return sh


def table_sheet(wb, name, title, sub, cols, groups, total=None, notes=None, freeze_cols=2):
    """cols = [dict(h, w, fmt, align, wrap)] · groups = [(그룹 이름 | None, rows, 소계 row | None)]
    row 끝에 dict(_hl=색, _hlcols=[열 번호]) 를 붙이면 그 칸만 강조"""
    sh = wb.create_sheet(name)
    n = len(cols)
    for j, c in enumerate(cols, 1):
        sh.column_dimensions[L(j)].width = c.get('w', 12)
    _band(sh, title, sub, n)
    r = 4
    for t in notes or []:
        r = _para(sh, r, t, n, size=9, color=SUB)
    if notes:
        r += 1
    hr = r
    for j, c in enumerate(cols, 1):
        x = sh.cell(r, j, c['h']); x.font = f(9.5, True); x.fill = fill(HEAD)
        x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        x.border = Border(top=mid, bottom=mid)
    sh.row_dimensions[r].height = 32
    r += 1

    def put(row, style=None):
        nonlocal r
        meta = row[-1] if row and isinstance(row[-1], dict) else {}
        vals = row[:-1] if meta else row
        for j, (c, v) in enumerate(zip(cols, vals), 1):
            x = sh.cell(r, j, v)
            x.font = f(9.5, bold=style in ('sub', 'total'))
            if isinstance(v, (int, float)) and not isinstance(v, bool) and c.get('fmt'):
                x.number_format = FMT[c['fmt']]
            ha = c.get('align', 'right' if isinstance(v, (int, float)) else 'left')
            x.alignment = Alignment(horizontal=ha, vertical='center', wrap_text=c.get('wrap', False), indent=1 if ha == 'left' else 0)
            x.border = Border(bottom=thin)
            if style == 'sub':
                x.fill = fill(GROUP); x.border = Border(bottom=thin, top=thin)
            if style == 'total':
                x.fill = fill(TOTAL); x.border = Border(top=mid, bottom=mid)
            if meta.get('_hl') and j in meta.get('_hlcols', []):
                x.fill = fill(HL[meta['_hl']])
        lines = max((len(str(v)) // max(6, int(c.get('w', 12) / 1.9)) + 1 for c, v in zip(cols, vals) if c.get('wrap') and v), default=1)
        sh.row_dimensions[r].height = max(18, 14 * lines + 4)
        r += 1

    for g, rows, subt in groups:
        if g:
            sh.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n)
            x = sh.cell(r, 1, g); x.font = f(9.5, True, ACC); x.alignment = Alignment(indent=1)
            r += 1
        for row in rows:
            put(row)
        if subt:
            put(subt, 'sub')
    if total:
        put(total, 'total')
    sh.freeze_panes = f'{L(freeze_cols + 1)}{hr + 1}'
    sh.auto_filter.ref = f'A{hr}:{L(n)}{hr}'
    _page(sh)
    return sh


def basis_sheet(wb, rows, stamp):
    sh = wb.create_sheet(BASIS)
    for j, w in enumerate([14, 26, 70, 34], 1):
        sh.column_dimensions[L(j)].width = w
    _band(sh, '산식 · 가정 · 출처', '출처 · ' + stamp, 4)
    r = 4
    for j, h in enumerate(['구분', '항목', '내용', '출처'], 1):
        x = sh.cell(r, j, h); x.font = f(9.5, True); x.fill = fill(HEAD); x.border = Border(top=mid, bottom=mid)
    for row in rows:
        r += 1
        for j, v in enumerate(row, 1):
            x = sh.cell(r, j, v); x.font = f(9.5, bold=j == 1, color=INK if j < 4 else SUB)
            x.alignment = Alignment(wrap_text=True, vertical='top', indent=1); x.border = Border(bottom=thin)
        sh.row_dimensions[r].height = max(18, 15 * (len(str(row[2])) // 38 + 1))
    _page(sh)
    return sh


# ---------------------------------------------------------------------------
# 읽기 도우미
# ---------------------------------------------------------------------------
def num(v):
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(',', '').replace('명', '').strip()
    try:
        return float(s)
    except ValueError:
        return None


def nz(v):
    x = num(v)
    return 0.0 if x is None else x


def sheet_rows(sh):
    """write_xlsx 첨부 데이터 시트 → (notes, head, rows[dict])  (hrlib.read_xlsx_table 과 같은 규칙)"""
    rows = list(sh.iter_rows(values_only=True))
    notes = [str(r[0]) for r in rows[2:8] if r and r[0] is not None and all(v is None for v in r[1:])]
    hi = None
    for i, r in enumerate(rows):
        if r and r[0] is not None and i > 1 and all(v is None for v in rows[i - 1]) and any(v is not None for v in r[1:]):
            hi = i
            break
    if hi is None:
        return notes, [], []
    head = [re.sub(r'\s+', ' ', str(v)).strip() if v is not None else '' for v in rows[hi]]
    return notes, head, [dict(zip(head, r)) for r in rows[hi + 1:] if not all(v is None for v in r)]


def org(outdir):
    """①_RPU보조표.md → {코드: dict(div, name, now)} · 본부 순서 목록"""
    p = os.path.join(outdir, '①_RPU보조표.md')
    out, divs = {}, []
    if os.path.exists(p):
        t = open(p, encoding='utf-8').read()
        for m in re.finditer(r'^\| ([^|]+?) \| ([A-Za-z]+\d+) \| ([^|]+?) \| (\d+(?:\.\d+)?) \|', t, re.M):
            d = m.group(1).strip()
            out[m.group(2)] = dict(div=d, name=m.group(3).strip(), now=float(m.group(4)))
            if d not in divs:
                divs.append(d)
    return out, divs


def grouped(codes, O, divs):
    """코드 목록 → [(본부, [코드])] (보조표에 없는 코드는 「그 밖」)"""
    g = [(d, [c for c in codes if O.get(c, {}).get('div') == d]) for d in divs]
    rest = [c for c in codes if c not in O]
    if rest:
        g.append(('그 밖 (보조표에 없음)', rest))
    return [x for x in g if x[1]]


def nm(c, O):
    return O.get(c, {}).get('name', '')


def stamp_of(sh):
    return str(sh.cell(2, 1).value or '')


def fmt1(x):
    return f'{x:,.1f}'


def ty_of(head, prefix):
    for h in head:
        m = re.match(prefix + r'\s*(\d{4})$', h)
        if m:
            return m.group(1)
    return ''


# ---------------------------------------------------------------------------
# 파일별 보고서
# ---------------------------------------------------------------------------
def rep_daesa(wb, outdir, O, divs):
    sh = wb['대사']
    notes, head, rows = sheet_rows(sh)
    data = [r for r in rows if str(r.get('RPU 코드') or '') not in ('', '합계')]
    rpu = [r for r in data if str(r.get('구간') or '') != '사업 이벤트']
    ev = [r for r in data if str(r.get('구간') or '') == '사업 이벤트']
    m = re.search(r'A 범위 ([\d.]+)~([\d.]+)', ' '.join(notes))
    a_lo, a_hi = (float(m.group(1)), float(m.group(2))) if m else (None, None)
    m = re.search(r'\((?:[^)]*?)(\d{4})\)', str(sh.cell(1, 1).value or ''))
    ty = m.group(1) if m else ''
    dec_rpu = sum(nz(r['결정 수요 사람']) for r in rpu)
    evsum = sum(nz(r['결정 수요 사람']) for r in ev)
    dec = dec_rpu + evsum
    dig = sum(nz(r['결정 수요 디지털']) for r in data)
    now = sum(v['now'] for v in O.values()) or None
    zone = lambda r: str(r.get('구간') or '').split(' ')[0]
    rev = [r for r in rpu if zone(r) == '재검토']
    rev_ai = [r for r in rev if nz(r['결정 수요 디지털']) > 0]
    rev_other = [r for r in rev if r not in rev_ai]
    extra = [r for r in rpu if str(r.get('판정') or '').startswith(('추가 확인', '재검토 · 잠정'))]
    human = [r for r in rpu if str(r.get('판정') or '').startswith('사람 판정')]
    msgs = [f'{ty}년 결정 수요는 사람 {fmt1(dec)}명입니다. RPU 채택 값 합 {fmt1(dec_rpu)}명'
            + (f'에 사업 이벤트 {fmt1(evsum)}명 ({", ".join(str(r["RPU 코드"]) for r in ev)}) 을 별도 줄로 더했습니다' if ev else '입니다')
            + (f'. 현원 {now:,.0f}명보다 {dec - now:+,.1f}명 ({(dec - now) / now:+.1%}) 입니다.' if now else '.')]
    if a_lo is not None:
        inside = a_lo <= dec_rpu <= a_hi
        t = f'RPU 합 {fmt1(dec_rpu)}명은 방식 A 범위 {a_lo:,.0f}~{a_hi:,.0f}명 {"안" if inside else "밖"}에 있습니다.'
        if dec > a_hi and ev:
            t += f' 결정 수요가 A 상한을 {fmt1(dec - a_hi)}명 넘는 것은 사업 이벤트를 별도 줄로 더했기 때문입니다 (매출 회귀 · HCROI 역산에는 아직 잡히지 않는 인원).'
        elif not inside:
            t += ' 범위 밖인 이유를 대사표 근거 열과 결정로그에 적어야 합니다.'
        msgs.append(t)
    if rev:
        t = f'B 와 C 차이가 7%를 넘는 재검토 RPU 는 {len(rev)}개입니다.'
        if rev_ai:
            t += f' 그중 {len(rev_ai)}개 ({", ".join(r["RPU 코드"] for r in rev_ai)}) 는 AI 적용 RPU 로, C 가 AI 적용을 반영해 B 보다 작습니다.'
        if rev_other:
            t += f' 나머지 {len(rev_other)}개 ({", ".join(r["RPU 코드"] for r in rev_other)}) 는 근거 열의 원인을 확인합니다.'
        msgs.append(t)
    if extra:
        msgs.append(f'{", ".join(str(r["RPU 코드"]) + " " + nm(r["RPU 코드"], O) for r in extra)} 는 「{extra[0]["판정"]}」 입니다. C 로 잠정 채택하고 ④ 대안표에서 확인 대상으로 다룹니다.')
    if human:
        msgs.append(f'사람이 대사 판정을 바꾼 RPU: {", ".join(r["RPU 코드"] for r in human)} (②_수요파라미터.md 6절 · 결정로그).')
    if dig:
        msgs.append(f'디지털 FTE {fmt1(dig)} 는 사람 결정 수요에 더하지 않고 ④ Gap 의 「도입 계획」 칸으로 따로 넘깁니다.')
    summary_sheet(wb, dict(
        title=f'② 세 값 대사 · {ty} 결정 수요',
        sub='방식 A 매출·이익 · B 인력동인 · C 업무량을 RPU 마다 견주고, 평균 내지 않고 원인으로 한 값을 채택함',
        status='CP2 승인 뒤 계산 · 이 결정 수요가 ④ Gap 분석의 수요 기준선',
        kpis=[('결정 수요 · 사람', round(dec, 1), f'명 · RPU {fmt1(dec_rpu)}' + (f' + 사업 이벤트 {fmt1(evsum)}' if ev else '')),
              ('디지털 FTE', round(dig, 1), 'FTE · 사람 수요에 더하지 않음'),
              ('현원 대비', f'{dec - now:+,.1f}' if now else '-', f'명 · 현원 {now:,.0f}명' if now else '현원 자료 없음'),
              ('방식 A 범위', f'{a_lo:,.0f}~{a_hi:,.0f}' if a_lo else '-', '명 · 참고 범위 (매출 회귀 ~ HCROI 역산)')],
        messages=msgs,
        read=['결정 수요 = RPU 마다 원인 판정으로 채택한 값 (보통 방식 C 사람 FTE) + 사업 이벤트 (별도 줄). 세 값을 평균 내지 않음',
              '차이 구간: 3% 이하 일치 · 3~7% 설명 · 7% 초과 재검토. 재검토는 「틀렸다」가 아니라 원인을 적어야 하는 RPU',
              '원인 판정: 생산성 가정 · 동인 선택 · 업무량 과대 보고 · 사업 이벤트. 원인이 B 쪽 가정이면 C 채택',
              '디지털 FTE 는 AI 가 맡는 업무량. 사람 수요와 열을 나눠 두고 ④에서 따로 다룸'],
        decide=['재검토 RPU 의 원인 판정이 우리 회사 사정과 맞는가 (「대사표」 원인 · 근거 열)',
                '추가 확인 RPU 의 업무량 조사 원천을 누가 언제 확인할지',
                '결정 규칙을 바꾸려면 ②_수요파라미터.md 6절을 고치고 CP2 를 다시 거침 (이 파일을 손으로 고치지 않음)'],
        sheets=[('대사표', '본부별 RPU · B · C · 차이 · 구간 · 원인 · 결정 수요 (강조: 빨강 재검토 · 노랑 추가 확인 · 회색 일치)'),
                ('대사', '첨부 데이터 · 원 정밀도 · ④ Gap 스크립트가 읽는 열 이름 그대로'), (BASIS, '대사 규칙 · 결정 규칙 · 입력')],
        refs=['교안 Day 2 세 값 대사 · 대사 결과 읽기 장 · 뷰어 「대시보드」 · 「산출물」 › ②_대사.xlsx',
              '다시 만들기: CP2 를 다시 승인하면 리드가 ② 스크립트로 다시 계산 (Claude에게 「② 다시 계산해 줘」)', stamp_of(sh)]))
    cols = [dict(h='RPU', w=7, align='center'), dict(h='조직', w=20), dict(h='현원', w=8, fmt='int'),
            dict(h='B 인력동인', w=10, fmt='n1'), dict(h='C 업무량', w=10, fmt='n1'), dict(h='차이\n(B−C)÷C', w=10, fmt='spct'),
            dict(h='구간', w=9, align='center'), dict(h='원인 판정', w=14, wrap=True), dict(h='채택', w=14, wrap=True),
            dict(h='결정 수요\n사람', w=10, fmt='n1'), dict(h='디지털\nFTE', w=8, fmt='n1'), dict(h='근거', w=50, wrap=True), dict(h='가정 구분', w=11, align='center')]
    groups = []
    for d, cs in grouped([r['RPU 코드'] for r in rpu], O, divs):
        rs = [r for r in rpu if r['RPU 코드'] in cs]
        out = []
        for r in rs:
            c, z, b, cc = r['RPU 코드'], zone(r), num(r['B']), num(r['C'])
            diff = (b - cc) / cc if b is not None and cc else None
            v = str(r.get('판정') or '')
            hl = 'amber' if v.startswith(('추가', '재검토 · 잠정', '사람')) else 'red' if z == '재검토' else 'gray' if z == '일치' else None
            row = [c, nm(c, O), O.get(c, {}).get('now'), b, cc, diff, z, r['원인 판정'], v, num(r['결정 수요 사람']), nz(r['결정 수요 디지털']),
                   r.get('근거 한 줄'), r.get('가정 구분')]
            if hl:
                row.append(dict(_hl=hl, _hlcols=[6, 7, 8, 9]))
            out.append(row)
        sb = sum(nz(r['B']) for r in rs); sc = sum(nz(r['C']) for r in rs)
        groups.append((d, out, ['', f'{d} 소계', sum(O.get(r['RPU 코드'], {}).get('now') or 0 for r in rs), sb, sc, (sb - sc) / sc if sc else None,
                                '', '', '', sum(nz(r['결정 수요 사람']) for r in rs), sum(nz(r['결정 수요 디지털']) for r in rs), '', '']))
    if ev:
        groups.append(('사업 이벤트 (별도 줄)', [[r['RPU 코드'], nm(r['RPU 코드'], O) or '신설 · 사업 이벤트', None, None, None, None, '별도 줄', '사업 이벤트',
                                            '사업 이벤트', num(r['결정 수요 사람']), nz(r['결정 수요 디지털']), r.get('근거 한 줄'), r.get('가정 구분')] for r in ev], None))
    sb = sum(nz(r['B']) for r in rpu); sc = sum(nz(r['C']) for r in rpu)
    table_sheet(wb, '대사표', f'대사표 · RPU별 세 값 비교와 {ty} 결정 수요', '단위 명 (FTE) · 강조: 빨강 = 재검토 (7% 초과) · 노랑 = 추가 확인 · 사람 판정 · 회색 = 일치',
                cols, groups, ['', '전사 합계', now, sb, sc, (sb - sc) / sc if sc else None, '', '', '', dec, dig,
                               f'RPU 합 {fmt1(dec_rpu)}' + (f' + 사업 이벤트 {fmt1(evsum)}' if ev else ''), ''],
                notes=['차이 = (B − C) ÷ C · 양수 = B 가 큼 (C 가 AI 적용 · 생산성 반영으로 작음) · 음수 = C 가 큼 (지금 부족분이 드러남)'])
    basis_sheet(wb, [('규칙', '차이 구간', notes[0] if notes else '', 'wfp-step2-demand SKILL 대사 규칙'),
                     ('규칙', '결정 규칙', '세 값을 평균 내지 않음 · 원인 판정으로 한 방식을 채택 · 사업 이벤트는 별도 줄로 더함', '②_수요파라미터.md 6절'),
                     ('규칙', '추가 확인 · 잠정 C', 'C 로 잠정 채택하고 ④ 대안표에서 「추가 확인」으로 다룸 · CP2 를 다시 열지 않음', 'wfp-step2-demand SKILL'),
                     ('가정', '가정 구분', '사업계획 = 사업계획 숫자 그대로 · 회사 가정 = CP2 에서 사람이 정한 값', '②_수요파라미터.md 3 · 4절'),
                     ('출처', '입력', 'B = ②B_인력동인.xlsx · C = ②C_업무량.xlsx · A 범위 = ②A_매출이익.xlsx', '같은 폴더')]
                + [('메모', '첨부 데이터 안내', n, '첨부 「대사」') for n in notes[1:]], stamp_of(sh))
    return ['대사']


def dict_def(outdir, term):
    """data/<회사>/데이터사전.md 에서 용어 정의 칸 (표 2번째 칸) · 없으면 빈칸"""
    p = os.path.join(os.path.dirname(os.path.dirname(outdir)), 'data', os.path.basename(outdir), '데이터사전.md')
    if not os.path.exists(p):
        return ''
    for ln in open(p, encoding='utf-8'):
        c = [x.strip() for x in ln.strip().strip('|').split('|')]
        if len(c) > 2 and c[0] == term:
            return c[1]
    return ''


def rep_A(wb, outdir, O, divs):
    sh, shi = wb['매출이익'], wb['HR Index'] if 'HR Index' in wb.sheetnames else None
    notes, head, rows = sheet_rows(sh)
    ty = ty_of(head, '') or ''
    yk = next((h for h in head if re.fullmatch(r'\d{4} 인원', h)), None)
    ty = yk.split()[0] if yk else ''
    reg = {str(r['시나리오']): nz(r[yk]) for r in rows if str(r.get('길') or '').startswith('회귀')}
    hc = next((r for r in rows if str(r.get('길') or '').startswith('HCROI')), None)
    hcv = nz(hc[yk]) if hc else None
    lo, hi = (min(hcv, reg.get('기본', hcv)), max(hcv, reg.get('기본', hcv))) if hcv is not None and reg else (None, None)
    ni, _, idx = sheet_rows(shi) if shi else ([], [], [])
    msgs = []
    now = None
    if idx:
        a, b = idx[0], idx[-1]
        now = nz(b['인원'])
        g_hc, g_rev = nz(b['인원']) / nz(a['인원']) - 1, nz(b['매출']) / nz(a['매출']) - 1
        msgs.append(f'{a["연도"]}~{b["연도"]} 인원이 {g_hc:+.0%} 변할 때 매출은 {g_rev:+.0%} 변했습니다. 인당 매출 {nz(a["인당 매출"]):,.2f} → {nz(b["인당 매출"]):,.2f}.')
        dips = [str(r['연도']) for p, r in zip(idx, idx[1:]) if nz(r['HCROI']) < nz(p['HCROI'])]
        msgs.append(f'HCROI {nz(a["HCROI"]):.2f} → {nz(b["HCROI"]):.2f}배 · 인건비 비중 {nz(a["인건비 비중 %"]):.1f}% → {nz(b["인건비 비중 %"]):.1f}%'
                    + (f' · HCROI 가 꺾인 해 {", ".join(dips)}' if dips else '') + '. 지표 하나로 결론 내지 않고 2~3개를 겹쳐 봅니다.')
    if lo is not None:
        msgs.append(f'{ty}년 전사 적정 인원 범위는 {lo:,.0f}~{hi:,.0f}명입니다 (회귀 기본 {reg.get("기본", 0):,.1f} · 목표 HCROI {hc["입력 값"]} 역산 {hcv:,.1f})'
                    + (f'. 현원 {now:,.0f}명보다 {lo - now:+,.0f}~{hi - now:+,.0f}명.' if now else '.'))
    if reg:
        msgs.append(f'회귀 시나리오 보수 {reg.get("보수", 0):,.1f} · 기본 {reg.get("기본", 0):,.1f} · 낙관 {reg.get("낙관", 0):,.1f}. 이 범위는 RPU 결정 값이 아니라 B · C 결과를 견줄 기준선입니다.')
    part = any('방식 A 만 먼저' in n for n in notes)
    summary_sheet(wb, dict(
        title=f'②A 매출·이익 연계 · {ty} 전사 적정 인원 범위', sub='매출 회귀와 목표 HCROI 역산, 두 길로 전사 인원 범위를 내고 HR Index 추이로 생산성을 읽음',
        status='CP2 승인 전에 먼저 계산한 값 (방식 A 만) · B · C · 대사는 CP2 승인 뒤' if part else '전사 참고 범위 · ② 세 값 대사의 기준선',
        kpis=[('적정 인원 범위', f'{lo:,.0f}~{hi:,.0f}' if lo is not None else '-', '명 · HCROI 역산 ~ 회귀 기본'),
              ('회귀 기본', round(reg.get('기본', 0), 1), '명 · 기본 시나리오 매출'),
              ('HCROI 역산', round(hcv, 1) if hcv is not None else '-', f'명 · 목표 HCROI {hc["입력 값"] if hc else "-"}'),
              ('현원', now if now is not None else '-', '명 · HR Index 마지막 해')],
        messages=msgs,
        read=([f'인건비 = 총인건비 ({dict_def(outdir, "총인건비")} · 데이터 사전)'] if dict_def(outdir, '총인건비') else []) + ['길 ① 매출 회귀: 지난 인원과 매출의 직선 관계에 계획 매출을 넣음 · 표본이 적으면 방향만 봄',
              '길 ② HCROI 역산: 계획 영업이익과 목표 HCROI 로 감당할 인건비를 내고 1인당 인건비로 나눔',
              'HCVA = (영업이익 + 인건비) ÷ 인원 · HCROI = (영업이익 + 인건비) ÷ 인건비',
              '방식 A 는 전사 범위만 냄 · RPU 별 숫자는 방식 B · C 와 대사에서 정함'],
        decide=['목표 HCROI 수준과 그 근거 (②_수요파라미터.md 1절)', '기준 시나리오 (보수 · 기본 · 낙관)', '단위 · 인건비 정의 · 회귀 표본 기간이 맞는가 (아래 산식 시트)'],
        sheets=[('범위표', '두 길 결과'), ('HR Index표', '연도별 생산성 지표'), ('매출이익 · HR Index', '첨부 데이터 (원 정밀도)'), (BASIS, '단위 · 인건비 정의 · 표본')],
        refs=['교안 Day 2 방식 A 매출·이익 연계 장 · 뷰어 「산출물」 › ②A_매출이익.xlsx', stamp_of(sh)]))
    cols = [dict(h='길', w=24), dict(h='시나리오', w=12), dict(h='입력 값', w=12, fmt='n2'), dict(h=f'{ty} 인원', w=11, fmt='n1'),
            dict(h='표본', w=16), dict(h='R²', w=8, fmt='n3'), dict(h='근거', w=60, wrap=True)]
    table_sheet(wb, '범위표', f'{ty} 전사 적정 인원 범위 · 두 길', '단위: 인원 명 · 매출 · 이익은 회사 파일 단위',
                cols, [(None, [[r['길'], r['시나리오'], num(r['입력 값']) if num(r['입력 값']) is not None else r['입력 값'], num(r[yk]), r['표본'],
                                num(r['R²']) if num(r['R²']) is not None else r['R²'], r['근거']] for r in rows], None)], notes=notes)
    if idx:
        cols = [dict(h='연도', w=8, align='center'), dict(h='인원', w=9, fmt='int'), dict(h='매출', w=11, fmt='int'), dict(h='영업이익', w=10, fmt='int'),
                dict(h='인건비', w=10, fmt='int'), dict(h='인당 매출', w=10, fmt='n2'), dict(h='인당\n영업이익', w=10, fmt='n3'), dict(h='HCVA', w=9, fmt='n2'),
                dict(h='HCROI', w=9, fmt='n2'), dict(h='인건비\n비중 %', w=9, fmt='n1')]
        table_sheet(wb, 'HR Index표', 'HR Index · 연도별 생산성', ni[0] if ni else '',
                    cols, [(None, [[str(r['연도'])] + [num(r[h]) for h in ['인원', '매출', '영업이익', '인건비', '인당 매출', '인당 영업이익', 'HCVA', 'HCROI', '인건비 비중 %']] for r in idx], None)],
                    notes=ni[1:], freeze_cols=1)
    dd = dict_def(outdir, '총인건비')
    basis_sheet(wb, ([('정의', '총인건비', dd, '데이터사전.md')] if dd else []) + [('메모', '안내', n, '첨부 「매출이익」') for n in notes]
                + [('메모', 'HR Index', n, '첨부 「HR Index」') for n in ni], stamp_of(sh))
    return ['매출이익', 'HR Index']


def rep_B(wb, outdir, O, divs):
    sh = wb['인력동인']
    notes, head, rows = sheet_rows(sh)
    bk = next((h for h in head if re.fullmatch(r'B \d{4}', h)), None)
    ty = bk.split()[1] if bk else ''
    pk = next((h for h in head if re.fullmatch(r'\d{4} 계획', h)), None)
    by = next((h for h in head if re.fullmatch(r'\d{4}', h)), None)
    rs = [r for r in rows if str(r.get('RPU 코드') or '') not in ('', '합계')]
    have = [r for r in rs if num(r.get(bk)) is not None]
    skip = [r for r in rs if num(r.get(bk)) is None]
    btot = sum(nz(r[bk]) for r in have)
    now = sum(nz(r['현원']) for r in rs)
    inc = sorted(have, key=lambda r: nz(r[bk]) - nz(r['현원']), reverse=True)
    msgs = [f'방식 B {ty}년 인원 합은 {fmt1(btot)}명으로 현원 {now:,.0f}명보다 {btot - now:+,.1f}명입니다' + (f' (방식 B 를 쓰지 않은 RPU {len(skip)}개 제외).' if skip else '.')]
    if inc:
        msgs.append('가장 많이 늘어나는 RPU: ' + ' · '.join(f'{r["RPU 코드"]} {nm(r["RPU 코드"], O)} {nz(r[bk]) - nz(r["현원"]):+.1f} ({r["선정 동인"]})' for r in inc[:3]) + '.')
    mg = [r for r in have if '관리' in str(r.get('유형') or '') or '직접 RPU' in str(r.get('방식') or '')]
    if mg:
        msgs.append(f'관리 지원 RPU {len(mg)}개는 동인 대신 직접 RPU 인원 대비 비율로 계산했습니다.')
    ca = [r for r in rs if '회사 가정' in str(r.get('가정 구분') or '')]
    if ca:
        msgs.append(f'가정 구분이 「회사 가정」인 RPU {len(ca)}개: {", ".join(r["RPU 코드"] for r in ca)} (사업계획에 숫자가 없어 CP2 에서 사람이 정한 값).')
    summary_sheet(wb, dict(
        title=f'②B 조직별 인력동인 · {ty} RPU별 인원', sub='RPU 마다 인원을 움직이는 동인을 고르고 (CP2), 과거 인원 ÷ 동인 비율에 계획 동인을 넣어 인원을 냄',
        status='CP2 에서 승인한 동인으로 계산 · ② 세 값 대사의 B 값',
        kpis=[('B 합', round(btot, 1), '명 · 관리 지원 포함'), ('현원', now, '명'), ('현원 대비', round(btot - now, 1), '명'),
              ('동인 RPU', len(have), f'개 · 방식 B 생략 {len(skip)}개')],
        messages=msgs,
        read=['선정 동인은 상관 순위가 아니라 조직 유형 원칙으로 고름 (운영 = 실적 물량 · 영업 = 수주 · 고객 · 기술·전문 = 지원 대상 규모 · 관리 지원 = 직접 RPU 인원 비율)',
              f'{ty} 인원 = 과거 3개년 인원 ÷ 동인 평균 비율 × {ty} 계획 동인 × 보정 (보정 · 가산은 ②_수요파라미터.md 3절)',
              '상관 (순위) 는 참고 · 2027 계획이 없는 동인은 쓰지 않음'],
        decide=['고른 동인을 조직 원칙으로 설명할 수 있는가', f'동인마다 {ty} 계획 값이 사업계획에 있는가', 'B 합이 방식 A 범위 밖이면 왜인가 (대사표에서 확인)'],
        sheets=[('동인표', '본부별 RPU · 선정 동인 · 계획 · 상관 · B · 증감'), ('인력동인', '첨부 데이터 (원 정밀도)'), (BASIS, '방식 · 출처')],
        refs=['교안 Day 2 방식 B 조직별 인력동인 장 · 뷰어 「산출물」 › ②B_인력동인.xlsx', stamp_of(sh)]))
    cols = [dict(h='RPU', w=7, align='center'), dict(h='조직', w=18), dict(h='유형', w=14), dict(h='선정 동인', w=22, wrap=True), dict(h='단위', w=8, align='center'),
            dict(h=f'{by or "기준"} 값', w=10, fmt='n1'), dict(h=f'{ty} 계획', w=10, fmt='n1'), dict(h='상관 (순위)', w=11, align='center'),
            dict(h='현원', w=8, fmt='int'), dict(h=f'B {ty}', w=9, fmt='n1'), dict(h='증감', w=9, fmt='s1'), dict(h='방식', w=46, wrap=True), dict(h='가정 구분', w=11, align='center')]
    groups = []
    for d, cs in grouped([r['RPU 코드'] for r in rs], O, divs):
        g = [r for r in rs if r['RPU 코드'] in cs]
        out = []
        for r in g:
            b = num(r.get(bk))
            row = [r['RPU 코드'], nm(r['RPU 코드'], O), r.get('유형'), r.get('선정 동인'), r.get('단위'), num(r.get(by)) if by else None,
                   num(r.get(pk)) if num(r.get(pk)) is not None else r.get(pk), nz(r['현원']), b, (b - nz(r['현원'])) if b is not None else None, r.get('방식'), r.get('가정 구분')]
            if b is not None and nz(r['현원']) and abs(b / nz(r['현원']) - 1) > 0.15:
                row.append(dict(_hl='amber', _hlcols=[10, 11]))
            out.append(row)
        sb = sum(nz(r.get(bk)) for r in g); sn = sum(nz(r['현원']) for r in g)
        groups.append((d, out, ['', f'{d} 소계', '', '', '', None, None, '', sn, sb, sb - sn, '', '']))
    table_sheet(wb, '동인표', f'동인표 · RPU별 선정 동인과 B {ty}', '단위: 인원 명 · 강조 (노랑) = 현원 대비 ±15% 넘게 변하는 RPU',
                cols, groups, ['', '전사 합계', '', '', '', None, None, '', now, btot, btot - now, '', ''])
    basis_sheet(wb, [('메모', '안내', n, '첨부 「인력동인」') for n in notes], stamp_of(sh))
    return ['인력동인']


def rep_C(wb, outdir, O, divs):
    sh = wb['업무량']
    notes, head, rows = sheet_rows(sh)
    pk = next((h for h in head if re.fullmatch(r'\d{4} 사람', h)), None)
    ty = pk.split()[0] if pk else ''
    ak, dk = f'{ty} AI 전', f'{ty} 디지털'
    rs = [r for r in rows if str(r.get('RPU 코드') or '') not in ('', '합계')]
    now = sum(nz(r['현원']) for r in rs)
    f25 = sum(nz(r['2025 FTE']) for r in rs) if '2025 FTE' in head else 0
    pre = sum(nz(r.get(ak)) for r in rs); hum = sum(nz(r.get(pk)) for r in rs); dig = sum(nz(r.get(dk)) for r in rs)
    ai = [r for r in rs if str(r.get('AI 적용') or '').strip()]
    ratio = sorted(rs, key=lambda r: nz(r.get('FTE ÷ 현원')))
    msgs = [f'2025 업무량을 FTE 로 바꾸면 {fmt1(f25)} 로 현원 {now:,.0f}명의 {f25 / now:.0%} 입니다. 조사 결과가 현원과 크게 어긋나지 않아 {ty} 계산의 출발점으로 씁니다.' if now else '',
            f'{ty} 사업 동인 · 보정을 반영하면 {fmt1(pre)} (AI 전), AI 적용 뒤 사람 {fmt1(hum)} + 디지털 {fmt1(dig)} 로 갈립니다.']
    if ai:
        red = sum(nz(r.get(ak)) - nz(r.get(pk)) for r in ai)
        msgs.append(f'AI 적용 RPU {len(ai)}개 ({", ".join(r["RPU 코드"] for r in ai)}) 에서 사람 FTE 가 {fmt1(red)} 줄고 디지털 FTE {fmt1(dig)} 가 생깁니다. 줄어든 사람 FTE 는 감원 대상이 아니라 ④ 재배치 검토 대상입니다.')
    if ratio:
        lo, hi = ratio[0], ratio[-1]
        msgs.append(f'FTE ÷ 현원 양 끝: 가장 높은 {hi["RPU 코드"]} {nm(hi["RPU 코드"], O)} {nz(hi["FTE ÷ 현원"]):.2f}배 · 가장 낮은 {lo["RPU 코드"]} {nm(lo["RPU 코드"], O)} {nz(lo["FTE ÷ 현원"]):.2f}배. 양 끝은 입력 행부터 확인합니다.')
    summary_sheet(wb, dict(
        title=f'②C 업무량 분석 · {ty} 사람 FTE + 디지털 FTE', sub='업무량 조사 시간을 FTE 로 바꾸고, 사업 동인 · 보정 · AI 적용을 반영해 사람과 디지털로 나눔',
        status='CP2 승인 뒤 계산 · ② 세 값 대사의 C 값 (대개 결정 수요로 채택)',
        kpis=[(f'{ty} 사람 FTE', round(hum, 1), 'AI 적용 뒤 · 사업 이벤트 제외'), ('디지털 FTE', round(dig, 1), '사람 FTE 에 더하지 않음'),
              ('2025 FTE', round(f25, 1), f'현원 {now:,.0f}명 대비 {f25 / now:.0%}' if now else ''), (f'{ty} AI 전', round(pre, 1), '사업 동인 · 보정 반영')],
        messages=[m for m in msgs if m],
        read=['2025 FTE = 연간 소요시간 합 ÷ 1인 가용시간 × (1 + Buffer)', f'{ty} = 2025 FTE × 업무량 계수 + 가산 · AI 적용 RPU 는 위임 · 함께 판정 업무에 적용률을 반영',
              '디지털 FTE 는 AI 가 맡는 업무량 · 사람 수요와 열을 나눔', 'FTE ÷ 현원이 크게 벗어난 RPU 는 업무량 조사 원천부터 확인'],
        decide=['규제 · 책임 업무에 「위임」 판정이 붙지 않았는가', '연간 가용시간을 넘는 개인이 있는가 (업무량 조사)', 'FTE ÷ 현원 양 끝 RPU 의 입력 행'],
        sheets=[('업무량표', '본부별 RPU · 현원 · 2025 FTE · 계수 · AI 전 · 사람 · 디지털'), ('업무량', '첨부 데이터 (원 정밀도)'), (BASIS, '산정 가정 · 출처')],
        refs=['교안 Day 2 방식 C 업무량 분석 장 · 뷰어 「산출물」 › ②C_업무량.xlsx', stamp_of(sh)]))
    cols = [dict(h='RPU', w=7, align='center'), dict(h='조직', w=18), dict(h='현원', w=8, fmt='int'), dict(h='2025 FTE', w=10, fmt='n1'),
            dict(h='FTE ÷ 현원', w=10, fmt='n2'), dict(h='업무량 계수', w=10, fmt='n3'), dict(h='가산', w=7, fmt='n1'), dict(h='AI 적용', w=8, align='center'),
            dict(h=f'{ty} AI 전', w=10, fmt='n1'), dict(h=f'{ty} 사람', w=10, fmt='n1'), dict(h=f'{ty} 디지털', w=10, fmt='n1'),
            dict(h='현원 대비\n사람', w=10, fmt='s1'), dict(h='가정 구분', w=11, align='center')]
    groups = []
    for d, cs in grouped([r['RPU 코드'] for r in rs], O, divs):
        g = [r for r in rs if r['RPU 코드'] in cs]
        out = []
        for r in g:
            q = nz(r.get('FTE ÷ 현원'))
            row = [r['RPU 코드'], nm(r['RPU 코드'], O), nz(r['현원']), num(r.get('2025 FTE')), q, num(r.get('C 계수')), num(r.get('가산')),
                   r.get('AI 적용'), num(r.get(ak)), num(r.get(pk)), nz(r.get(dk)), nz(r.get(pk)) - nz(r['현원']), r.get('가정 구분')]
            if q and (q > 1.1 or q < 0.9):
                row.append(dict(_hl='amber', _hlcols=[5]))
            out.append(row)
        sn = sum(nz(r['현원']) for r in g)
        groups.append((d, out, ['', f'{d} 소계', sn, sum(nz(r.get('2025 FTE')) for r in g), None, None, None, '',
                                sum(nz(r.get(ak)) for r in g), sum(nz(r.get(pk)) for r in g), sum(nz(r.get(dk)) for r in g),
                                sum(nz(r.get(pk)) for r in g) - sn, '']))
    table_sheet(wb, '업무량표', f'업무량표 · RPU별 FTE 와 {ty} 사람 · 디지털', '단위 FTE · 강조 (노랑) = FTE ÷ 현원 0.9 미만 · 1.1 초과',
                cols, groups, ['', '전사 합계', now, f25, f25 / now if now else None, None, None, '', pre, hum, dig, hum - now, ''])
    basis_sheet(wb, [('메모', '산정 가정', n, '첨부 「업무량」 · 파일 4 산정 가정') for n in notes], stamp_of(sh))
    return ['업무량']


def rep_S(wb, outdir, O, divs):
    sb, su = wb['기본'], wb['상승'] if '상승' in wb.sheetnames else None
    nb, hb, rb = sheet_rows(sb)
    nu, hu, ru = sheet_rows(su) if su else ([], [], [])
    sk = [h for h in hb if re.fullmatch(r'공급 \d{4}', h)]
    rk = [h for h in hb if re.fullmatch(r'정년 \d{4}', h)]
    k2 = sk[-1] if sk else None
    ty = k2.split()[1] if k2 else ''
    B = {r['RPU 코드']: r for r in rb if str(r.get('RPU 코드') or '') not in ('', '합계')}
    U = {r['RPU 코드']: r for r in ru if str(r.get('RPU 코드') or '') not in ('', '합계')}
    now = sum(nz(r['현원']) for r in B.values())
    sb_, su_ = sum(nz(r[k2]) for r in B.values()), sum(nz(U[c][k2]) for c in U)
    ret = sum(nz(r[k]) for r in B.values() for k in rk)
    loss = now - sb_
    rate = sorted(B.values(), key=lambda r: nz(r['현원 대비']))
    small = [c for c, r in B.items() if '표본' in str(r.get('주의') or '')]
    msgs = [f'채용하지 않으면 {ty}년 말 인원은 {fmt1(sb_)}명 (기본) 으로 현원 {now:,.0f}명보다 {fmt1(loss)}명 ({-loss / now:+.1%}) 줄어듭니다. 그중 정년 도래가 {ret:,.0f}명입니다.' if now else '']
    if U:
        msgs.append(f'RPU 마다 가장 높았던 해의 퇴직률을 쓰는 상승 시나리오는 {fmt1(su_)}명으로, 기본보다 {fmt1(sb_ - su_)}명 더 줄어듭니다.')
    if rate:
        msgs.append('감소율이 큰 RPU: ' + ' · '.join(f'{r["RPU 코드"]} {nm(r["RPU 코드"], O)} {nz(r["현원 대비"]) - 1:+.0%}' for r in rate[:3]) + '.')
    if small:
        msgs.append(f'현원이 적어 퇴직률이 흔들리는 「표본 작음」 RPU {len(small)}개: {", ".join(small)} (S-1 에서 본부 평균으로 바꿀지 사람이 정함).')
    msgs.append('신규 채용 · 신설 조직 · 디지털 FTE 는 넣지 않은 값이며, 이 공급이 ④ Gap 의 공급 기준선입니다.')
    summary_sheet(wb, dict(
        title=f'③ 인력공급 예측 · {ty} 채용 없음', sub='기초 인원에서 퇴직률과 정년 도래를 빼 해마다 남는 인원을 냄 · 기본 · 상승 두 시나리오',
        status='CP3 공급 가정 (S-1~S-5) 에 사람이 답한 뒤 계산 · ④ Gap 의 공급 기준선',
        kpis=[('공급 · 기본', round(sb_, 1), f'명 · {ty}년 말 · 채용 없음'), ('공급 · 상승', round(su_, 1) if U else '-', '명 · 퇴직률 상승 시나리오'),
              ('현원 대비 감소', round(-loss, 1), f'명 · 현원 {now:,.0f}명'), ('정년 도래', int(ret), f'명 · {" · ".join(k.split()[1] for k in rk)}')],
        messages=[m for m in msgs if m],
        read=['기말 = 기초 − 기초 × 퇴직률 (× 배수) − 정년 도래 · 해마다 반복', '퇴직률은 자발 퇴직만 (정년퇴직은 따로 빼서 섞지 않음)',
              '기본 = 3개년 합 비율 · 상승 = RPU 별 가장 높은 해 비율', '현원 대비 = 공급 ÷ 현원 (1 보다 작을수록 많이 줄어듦)'],
        decide=['S-1 퇴직률 기준 (3개년 평균 · 최근 1년 · 본부 평균)', 'S-2 배수 (기본 1.0 · 시험으로 바꾼 뒤 되돌렸는가)', 'S-3 정년 · S-4 승진 · S-5 이동 가정',
                '기초 합이 조직인원현황 합과 같은가'],
        sheets=[('공급표', '본부별 RPU · 현원 · 퇴직률 · 정년 · 공급 기본 · 상승 · 현원 대비'), ('기본 · 상승', '첨부 데이터 (원 정밀도 · ④ 가 읽음)'), (BASIS, '산식 · 출처')],
        refs=['교안 Day 2 ③ 인력공급 예측 장 · 뷰어 「산출물」 › ③_공급.xlsx · 「대시보드」 공급', stamp_of(sb)]))
    cols = [dict(h='RPU', w=7, align='center'), dict(h='조직', w=18), dict(h='현원', w=8, fmt='int'), dict(h='퇴직률\n기본', w=9, fmt='pct'),
            dict(h='퇴직률\n상승', w=9, fmt='pct'), dict(h='정년 도래', w=9, fmt='int'), dict(h=f'공급 {ty}\n기본', w=10, fmt='n1'),
            dict(h=f'공급 {ty}\n상승', w=10, fmt='n1'), dict(h='현원 대비\n(기본)', w=10, fmt='pct'), dict(h='감소 (기본)', w=10, fmt='s1'), dict(h='주의', w=12), dict(h='출처', w=40, wrap=True)]
    groups = []
    for d, cs in grouped(list(B), O, divs):
        out = []
        for c in cs:
            r, u = B[c], U.get(c, {})
            q = nz(r['현원 대비'])
            row = [c, nm(c, O), nz(r['현원']), num(r['퇴직률']), num(u.get('퇴직률')), sum(nz(r[k]) for k in rk), num(r[k2]), num(u.get(k2)),
                   q, nz(r[k2]) - nz(r['현원']), r.get('주의'), r.get('출처')]
            if q and q < 0.85:
                row.append(dict(_hl='red', _hlcols=[9, 10]))
            out.append(row)
        g = [B[c] for c in cs]
        sn, s1 = sum(nz(r['현원']) for r in g), sum(nz(r[k2]) for r in g)
        groups.append((d, out, ['', f'{d} 소계', sn, None, None, sum(nz(r[k]) for r in g for k in rk), s1, sum(nz(U.get(c, {}).get(k2)) for c in cs),
                                s1 / sn if sn else None, s1 - sn, '', '']))
    table_sheet(wb, '공급표', f'공급표 · RPU별 {ty}년 말 공급 (채용 없음)', '단위 명 · 강조 (빨강) = 현원 대비 85% 미만',
                cols, groups, ['', '전사 합계', now, None, None, ret, sb_, su_, sb_ / now if now else None, sb_ - now, '', ''])
    basis_sheet(wb, [('메모', '기본', n, '첨부 「기본」') for n in nb] + [('메모', '상승', n, '첨부 「상승」') for n in nu[:2]], stamp_of(sb))
    return ['기본', '상승']


REPORTS = {'②_대사.xlsx': rep_daesa, '②A_매출이익.xlsx': rep_A, '②B_인력동인.xlsx': rep_B, '②C_업무량.xlsx': rep_C, '③_공급.xlsx': rep_S}


def decorate(path):
    """write_xlsx 로 막 저장한 결과 파일에 보고서 시트를 붙여 다시 저장 · 첨부 데이터 시트는 손대지 않고 뒤로 보냄"""
    fn = REPORTS.get(os.path.basename(path))
    if not fn:
        return False
    outdir = os.path.dirname(path)
    O, divs = org(outdir)
    wb = openpyxl.load_workbook(path)
    data = [s.title for s in wb.worksheets]
    fn(wb, outdir, O, divs)
    for t in data:
        s = wb[t]
        s.sheet_properties.tabColor = '9E9E9E'
        a1 = str(s.cell(1, 1).value or '')
        if not a1.startswith('첨부 데이터'):
            s.cell(1, 1).value = '첨부 데이터 · ' + a1
    rep = [s for s in wb.worksheets if is_report_sheet(s.title) and s.title != BASIS]
    order = rep + [wb[t] for t in data] + ([wb[BASIS]] if BASIS in wb.sheetnames else [])
    wb._sheets = order
    wb.active = 0
    for s in wb.worksheets:
        s.sheet_view.tabSelected = s is order[0]
    wb.save(path)
    return True


# ---------------------------------------------------------------------------
# md 표 읽기 (④ · 리포트 · 내보내기)
# ---------------------------------------------------------------------------
def md_blocks(text):
    text = re.sub(r'<!--.*?-->', '', text, flags=re.S)
    text = re.sub(r'```.*?```', '', text, flags=re.S)
    out, lines, i = [], text.splitlines(), 0
    while i < len(lines):
        ln = lines[i].rstrip()
        if ln.startswith('|') and i + 1 < len(lines) and re.match(r'^\|[\s\-:|]+\|$', lines[i + 1].strip()):
            head = [c.strip() for c in ln.strip().strip('|').split('|')]
            rows = []
            i += 2
            while i < len(lines) and lines[i].startswith('|'):
                rows.append([c.strip() for c in lines[i].strip().strip('|').split('|')])
                i += 1
            out.append(('table', (head, rows)))
            continue
        m = re.match(r'^(#{1,4}) (.+)', ln)
        if m:
            out.append((f'h{len(m.group(1))}', m.group(2).strip()))
        elif re.match(r'^\s*[-*] ', ln):
            out.append(('bullet', re.sub(r'^\s*[-*] ', '', ln)))
        elif re.match(r'^\s*\d+\. ', ln):
            out.append(('bullet', ln.strip()))
        elif ln.strip():
            out.append(('para', ln.strip()))
        i += 1
    return out


def md_table_after(blocks, heading_prefix):
    on = False
    for k, v in blocks:
        if k.startswith('h'):
            on = v.startswith(heading_prefix)
        elif on and k == 'table':
            return v
    return None


def plain(s):
    return re.sub(r'\*\*|`', '', str(s))


def build_gap(outdir, script='wfp-step4-gap/scripts/step4.py'):
    """④_Gap대안.md (1 Gap 표 · 2 대안표 · 3 재배치 경로) → ④_Gap분석.xlsx"""
    p = os.path.join(outdir, '④_Gap대안.md')
    if not os.path.exists(p):
        return None
    B = md_blocks(open(p, encoding='utf-8').read())
    gt = md_table_after(B, '1. Gap 표')
    if not gt:
        return None
    O, divs = org(outdir)
    head, rows = gt
    ty = ty_of(head, '공급')
    rr = [dict(zip(head, r)) for r in rows]
    tot = next((r for r in rr if r.get('RPU 코드') == '합계'), None)
    rs = [r for r in rr if r.get('RPU 코드') not in ('합계', '')]
    G = lambda r: num(r.get('Gap')) or 0.0
    short = sorted([r for r in rs if G(r) > 0.5], key=G, reverse=True)
    spare = sorted([r for r in rs if G(r) < -0.5], key=G)
    gsum = num(tot.get('Gap')) if tot else sum(G(r) for r in rs)
    gup = num(tot.get('Gap (공급 상승)')) if tot else None
    dem = num(tot.get('결정 수요 사람')) if tot else None
    sup = num(tot.get(f'공급 {ty}')) if tot else None
    at = md_table_after(B, '2. 대안표')
    alt = [dict(zip(at[0], r)) for r in at[1]] if at else []
    alt = [a for a in alt if a.get('RPU 코드') and a.get('RPU 코드') != '합계' and any(v for k, v in a.items() if k != 'RPU 코드')]
    pt = md_table_after(B, '3. 재배치 경로')
    path = [dict(zip(pt[0], r)) for r in pt[1]] if pt else []
    path = [x for x in path if x.get('보내는 RPU') and any(v for k, v in x.items() if k != '보내는 RPU')]
    at4 = md_table_after(B, '승인')   # 맨 아래 「## 승인」 표 (CP4 답)
    cp4 = next((dict(zip(at4[0], r)) for r in at4[1] if r and r[0].strip() == 'CP4'), {}) if at4 else {}
    cp4_ok = str(cp4.get('답', '')).strip().startswith('확정') and bool(str(cp4.get('승인자 (직책)', '')).strip())
    other = [a for a in alt if str(a.get('승인') or '').strip() and not str(a.get('승인')).strip().startswith(('승인', '확정'))]
    # 행 「승인」 칸: 빈칸 = CP4 답을 따름 · 「승인 · 확정」 = 승인 · 그 밖 (추가 확인 · 보류 · 이관) = 다른 답
    ok = [a for a in alt if a not in other and (cp4_ok or str(a.get('승인') or '').strip().startswith(('승인', '확정')))]
    mix = {}
    for a in alt:
        for k in ('Buy', 'Build', 'Borrow', 'Bot'):
            if k in str(a.get('4B 대안') or ''):
                mix[k] = mix.get(k, 0) + 1
    top = short[:3]
    msgs = [f'{ty}년 Gap 은 결정 수요 {fmt1(dem)} − 공급 {fmt1(sup)} = {fmt1(gsum)}명입니다 (채용 없음 기준).' if dem is not None and sup is not None else f'Gap 합 {fmt1(gsum)}명.']
    if top:
        share = sum(G(r) for r in top) / sum(G(r) for r in short) if short else 0
        msgs.append(f'부족 RPU {len(short)}개 (합 {fmt1(sum(G(r) for r in short))}) 중 상위 3개가 {share:.0%} 를 차지합니다: '
                    + ' · '.join(f'{r["RPU 코드"]} {nm(r["RPU 코드"], O) or "(신설)"} {G(r):+.1f}' for r in top) + '.')
    if spare:
        msgs.append(f'여유 RPU {len(spare)}개 (합 {fmt1(-sum(G(r) for r in spare))}): {", ".join(r["RPU 코드"] for r in spare)}. 감원이 아니라 재배치 경로로 다룹니다'
                    + (f' (경로 {len(path)}건).' if path else ' (재배치 경로가 아직 비어 있음).'))
    if gup is not None:
        msgs.append(f'퇴직률이 오르면 Gap 은 {fmt1(gup)}명으로 {fmt1(gup - gsum)}명 커집니다.')
    if alt:
        msgs.append(f'대안 {len(alt)}행 · 승인 {len(ok)}행' + (f' (CP4 확정 · 범위 전체)' if cp4_ok else ' (CP4 승인 전)') + (f' · 다른 답 {len(other)}행' if other else '')
                    + ' · 4B 구성 ' + ' · '.join(f'{k} {v}' for k, v in mix.items()) + '.')
    else:
        msgs.append('대안표가 아직 비어 있습니다 (Gap 담당이 채우고 점검을 통과하면 CP4).')
    summary_sheet_wb = openpyxl.Workbook()
    wb = summary_sheet_wb
    wb.remove(wb.active)
    st = f'만든 스크립트: {script} · 입력: ④_Gap대안.md · ②_대사.xlsx · ③_공급.xlsx · {datetime.datetime.now():%Y-%m-%d %H:%M}'
    summary_sheet(wb, dict(
        title=f'④ 수급 Gap 분석 · {ty} 확보 · 육성 대안', sub='Gap = 결정 수요 (사람) − 공급 (채용 없음) · 유형 판정 · 4B 대안 · 재배치 경로 · 우선순위',
        status='CP4 승인 상태: ' + ((f'CP4 확정 ({cp4.get("승인자 (직책)", "").strip()}) · 대안 {len(ok)}/{len(alt)}행' + (f' · 다른 답 {len(other)}행' if other else '')) if cp4_ok
                                  else (f'CP4 승인 전 · 대안 {len(alt)}행 작성됨' if alt else '대안 작성 전')) + ' · 승인 원본 = ④_Gap대안.md',
        kpis=[('Gap · 기본', round(gsum, 1), '명 · 수요 − 공급'), ('Gap · 공급 상승', round(gup, 1) if gup is not None else '-', '명 · 퇴직률 상승'),
              ('부족 RPU', len(short), f'개 · 합 {fmt1(sum(G(r) for r in short))}명'), ('여유 RPU', len(spare), f'개 · 합 {fmt1(-sum(G(r) for r in spare))}명')],
        messages=msgs,
        read=['Gap > 0 부족 · < 0 여유 · 디지털 FTE 는 Gap 에 넣지 않고 「도입 계획」 칸', '유형: 수량 · 시점 · 구조 · AI 활용 (겹치면 함께 적음)',
              '4B = Buy 채용 · Build 육성 · 이동 · Borrow 외부 활용 · Bot AI · 자동화', '우선순위 = 영향 × 시급성 (R 1순위 · Y 2순위 · G 3순위)',
              '여유 RPU 는 「감원」이 아니라 받는 RPU · 교육 기간 · 시점이 있는 재배치 경로로 적음'],
        decide=['R (1순위) 대안의 시점 · 방법이 사업계획 일정과 맞는가', '여유 인원 재배치 경로 (받는 RPU · 교육 · 시점) 가 현실적인가',
                'R 행은 소관 본부장 승인이 필요한가 (CP4 답에 범위를 밝힘)'],
        sheets=[('Gap표', '본부별 RPU · 수요 · 공급 · Gap · 유형 · 방향'), ('대안표', '4B 대안 · 시점 · 경로 · 우선순위 · 승인'), ('경로표', '여유 → 부족 재배치 경로'),
                ('Gap', '첨부 데이터 (④_Gap대안.md 1절 그대로)'), (BASIS, '규칙 · 출처')],
        refs=['교안 Day 2 ④ 수급 Gap 분석 장 · 뷰어 「산출물」 › ④_Gap분석.xlsx · 승인 원본 ④_Gap대안.md', '이 파일은 ④_Gap대안.md 에서 다시 만들어짐 (손으로 고치지 않음)', st]))
    ok_hd = [h for h in head if h not in ('RPU 코드',)]
    cols = [dict(h='RPU', w=7, align='center'), dict(h='조직', w=18)] + [dict(h=h, w=13 if len(h) > 6 else 10, fmt='n1' if h not in ('유형', '방향', '시점 이벤트') else None,
                                                                         wrap=h in ('유형', '시점 이벤트')) for h in ok_hd]
    groups = []
    for d, cs in grouped([r['RPU 코드'] for r in rs], O, divs):
        out = []
        for c in cs:
            r = next(x for x in rs if x['RPU 코드'] == c)
            row = [c, nm(c, O) or '신설'] + [num(r[h]) if num(r[h]) is not None else r[h] for h in ok_hd]
            g = G(r)
            row.append(dict(_hl='red' if g >= 10 else 'amber' if g > 0.5 else 'green' if g < -0.5 else None, _hlcols=[2 + ok_hd.index('Gap') + 1]) if 'Gap' in ok_hd and (g >= 10 or g > 0.5 or g < -0.5) else {})
            if not row[-1]:
                row.pop()
            out.append(row)
        g = [x for x in rs if x['RPU 코드'] in cs]
        sub = ['', f'{d} 소계'] + [sum(nz(x[h]) for x in g) if h not in ('유형', '방향', '시점 이벤트') else '' for h in ok_hd]
        groups.append((d, out, sub))
    table_sheet(wb, 'Gap표', f'Gap표 · RPU별 {ty} 수급 Gap', '단위 명 · 강조: 빨강 = 부족 10명 이상 · 노랑 = 부족 · 초록 = 여유',
                cols, groups, (['', '전사 합계'] + [num(tot[h]) if num(tot[h]) is not None else tot[h] for h in ok_hd]) if tot else None)
    if at:
        ah = at[0]
        cols = [dict(h=h, w={'RPU 코드': 8, 'Gap': 8, '경로': 40, '근거 (요인 번호)': 34}.get(h, 12), align='center' if h in ('RPU 코드', '우선순위', '2×2', '승인') else None,
                     fmt='n1' if h == 'Gap' else None, wrap=h in ('유형', '4B 대안', '경로', '근거 (요인 번호)', '시점')) for h in ah]
        rows_a = []
        for a in alt:
            row = [num(a[h]) if h == 'Gap' and num(a[h]) is not None else plain(a.get(h, '')) for h in ah]
            pr = str(a.get('우선순위') or '')
            if pr in ('R', 'Y', 'G'):
                row.append(dict(_hl={'R': 'red', 'Y': 'amber', 'G': 'green'}[pr], _hlcols=[ah.index('우선순위') + 1]))
            rows_a.append(row)
        table_sheet(wb, '대안표', '대안표 · 4B 대안 · 우선순위 · 승인', '승인 원본 = ④_Gap대안.md 2절 · 우선순위 R 1순위 · Y 2순위 · G 3순위', cols, [(None, rows_a, None)], freeze_cols=1)
    if pt:
        ph = pt[0]
        cols = [dict(h=h, w=14 if h not in ('교육', '전환') else 22, wrap=h in ('교육', '전환'), fmt='n1' if h == '사람 여유' else 'int' if h == '인원' else None) for h in ph]
        table_sheet(wb, '경로표', '재배치 경로 · 여유 RPU → 받는 RPU', '승인 원본 = ④_Gap대안.md 3절', cols,
                    [(None, [[num(x[h]) if h in ('사람 여유', '인원') and num(x[h]) is not None else plain(x.get(h, '')) for h in ph] for x in path], None)], freeze_cols=1)
    sh = wb.create_sheet('Gap')
    sh.cell(1, 1, '첨부 데이터 · ④ Gap 표 (④_Gap대안.md 1절 그대로)')
    sh.cell(2, 1, st)
    for j, h in enumerate(head, 1):
        sh.cell(4, j, h).font = f(9.5, True)
    for i, r in enumerate(rows, 5):
        for j, v in enumerate(r, 1):
            sh.cell(i, j, num(v) if num(v) is not None else v)
    sh.sheet_properties.tabColor = '9E9E9E'
    basis_sheet(wb, [('규칙', 'Gap', 'Gap = 결정 수요 사람 − 공급 (채용 없음 · 기본) · 상승 = 공급 상승 시나리오', 'wfp-step4-gap SKILL'),
                     ('규칙', '유형', '수량 (Gap > 0.5) · 시점 (시점 이벤트) · 구조 (Gap ÷ 현원 ≥ 30%) · AI 활용 (AI 적용 RPU)', 'wfp-step4-gap SKILL'),
                     ('규칙', '대안', '핵심직무 = Build 우선 · 시점 Gap = Buy · Borrow 병행 · 여유 RPU = 재배치 경로 먼저 · 감원 단독 금지', 'wfp-step4-gap SKILL'),
                     ('출처', '입력', '②_대사.xlsx 결정 수요 · ③_공급.xlsx 공급 · ④_Gap대안.md 2 · 3절 (Gap 담당 · CP4)', '같은 폴더')], st)
    wb._sheets = [wb['요약']] + [wb[n] for n in ('Gap표', '대안표', '경로표') if n in wb.sheetnames] + [wb['Gap'], wb[BASIS]]
    wb.active = 0
    dst = os.path.join(outdir, '④_Gap분석.xlsx')
    wb.save(dst)
    return dst


# ---------------------------------------------------------------------------
# 워드 (리포트 · 내보내기)
# ---------------------------------------------------------------------------
def _docx():
    try:
        import docx  # noqa
        return True
    except ImportError:
        return False


class W:
    """python-docx 얇은 감싸기 (맑은 고딕 · 표 머리 음영 · 합계 행 · 넓은 표는 가로 쪽)"""

    def __init__(self, footer):
        from docx import Document
        from docx.shared import Pt, Cm, RGBColor
        from docx.oxml.ns import qn
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        self.Pt, self.Cm, self.RGB, self.qn, self.AL = Pt, Cm, RGBColor, qn, WD_ALIGN_PARAGRAPH
        self.d = Document()
        s = self.d.sections[0]
        s.page_height, s.page_width = Cm(29.7), Cm(21.0)
        s.left_margin = s.right_margin = Cm(2.0)
        s.top_margin, s.bottom_margin = Cm(2.0), Cm(1.8)
        st = self.d.styles['Normal']; st.font.name = FONT; st.font.size = Pt(10)
        st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
        self.footer = footer
        self._foot(s)
        self.ink, self.sub, self.acc = RGBColor(0x2D, 0x29, 0x26), RGBColor(0x6B, 0x64, 0x5D), RGBColor(0xD0, 0x4A, 0x02)

    def _foot(self, s):
        p = s.footer.paragraphs[0]; p.alignment = self.AL.CENTER
        p.text = ''
        self.run(p, self.footer, 8, color=getattr(self, 'sub', None))

    def run(self, p, text, size=10, bold=False, color=None):
        for t in re.split(r'(\*\*[^*]+\*\*|`[^`]+`)', str(text)):
            if not t:
                continue
            b = bold
            if t.startswith('**'):
                t, b = t[2:-2], True
            elif t.startswith('`'):
                t = t[1:-1]
            r = p.add_run(t); r.font.size = self.Pt(size); r.font.bold = b; r.font.name = FONT
            r._element.rPr.rFonts.set(self.qn('w:eastAsia'), FONT)
            if color is not None:
                r.font.color.rgb = color
        return p

    def para(self, text='', size=10, bold=False, color=None, after=4):
        p = self.d.add_paragraph(); p.paragraph_format.space_after = self.Pt(after); p.paragraph_format.line_spacing = 1.25
        return self.run(p, text, size, bold, color if color is not None else self.ink)

    def heading(self, text, level=1):
        from docx.oxml import OxmlElement
        p = self.d.add_paragraph(); p.paragraph_format.space_before = self.Pt(14 if level == 1 else 8)
        p.paragraph_format.space_after = self.Pt(4); p.paragraph_format.keep_with_next = True
        self.run(p, text, {1: 15, 2: 12, 3: 10.5}.get(level, 10.5), True, self.acc if level == 1 else self.ink)
        if level == 1:
            pPr = p._p.get_or_add_pPr(); bd = OxmlElement('w:pBdr'); bt = OxmlElement('w:bottom')
            bt.set(self.qn('w:val'), 'single'); bt.set(self.qn('w:sz'), '6'); bt.set(self.qn('w:color'), ACC); bd.append(bt); pPr.append(bd)

    def page_break(self):
        from docx.enum.text import WD_BREAK
        self.d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    def orient(self, landscape):
        from docx.enum.section import WD_ORIENT, WD_SECTION
        cur = self.d.sections[-1]
        if (cur.orientation == WD_ORIENT.LANDSCAPE) == landscape:
            return
        s = self.d.add_section(WD_SECTION.NEW_PAGE)
        s.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
        w, h = self.Cm(29.7), self.Cm(21.0)
        s.page_width, s.page_height = (w, h) if landscape else (h, w)

    def table(self, head, rows, size=8.5):
        from docx.oxml import OxmlElement
        from docx.enum.table import WD_TABLE_ALIGNMENT
        t = self.d.add_table(rows=1, cols=len(head)); t.alignment = WD_TABLE_ALIGNMENT.CENTER
        b = OxmlElement('w:tblBorders')
        for e in ('top', 'bottom', 'insideH'):
            x = OxmlElement(f'w:{e}'); x.set(self.qn('w:val'), 'single'); x.set(self.qn('w:sz'), '4' if e == 'insideH' else '8')
            x.set(self.qn('w:color'), LINE if e == 'insideH' else INK); b.append(x)
        for e in ('left', 'right', 'insideV'):
            x = OxmlElement(f'w:{e}'); x.set(self.qn('w:val'), 'nil'); b.append(x)
        t._tbl.tblPr.append(b)

        def shade(c, hx):
            s = OxmlElement('w:shd'); s.set(self.qn('w:val'), 'clear'); s.set(self.qn('w:color'), 'auto'); s.set(self.qn('w:fill'), hx)
            c._tc.get_or_add_tcPr().append(s)
        for j, h in enumerate(head):
            c = t.rows[0].cells[j]; c.text = ''; shade(c, HEAD)
            self.run(c.paragraphs[0], h, size, True, self.ink); c.paragraphs[0].alignment = self.AL.CENTER
        for r in rows:
            cells = t.add_row().cells
            last = str(r[0]).strip() in ('합계', '전사 합계') if r else False
            for j in range(len(head)):
                v = str(r[j]).strip() if j < len(r) else ''
                c = cells[j]; c.text = ''
                self.run(c.paragraphs[0], v, size, last, self.ink)
                if re.fullmatch(r'[+\-−]?[\d,]+(\.\d+)?%?( / [\d,.]+)?( ~ [\d,.]+)?명?', v):
                    c.paragraphs[0].alignment = self.AL.RIGHT
                if last:
                    shade(c, TOTAL)
        self.d.add_paragraph().paragraph_format.space_after = self.Pt(2)

    def blocks(self, blocks, skip_h1=True):
        for k, v in blocks:
            if k == 'h1' and skip_h1:
                continue
            if k in ('h1', 'h2'):
                self.orient(False)
                self.heading(v, 1)
            elif k in ('h3', 'h4'):
                self.heading(v, 2)
            elif k == 'bullet':
                self.para('• ' + v, 9.5, after=2)
            elif k == 'para':
                self.para(v, 9.5)
            elif k == 'table':
                h, rows = v
                wide = len(h) > 7 or sum(len(x) for x in h) > 70
                if wide:
                    self.orient(True)
                self.table(h, rows, 7.5 if len(h) > 9 else 8.5)
                if wide:
                    self.orient(False)

    def save(self, path, title):
        self.d.core_properties.title = title
        self.d.core_properties.keywords = 'pwc-hr-harness'
        self.d.save(path)


def _xlsx_read(outdir, fn, sheet):
    p = os.path.join(outdir, fn)
    if not os.path.exists(p):
        return [], [], []
    wb = openpyxl.load_workbook(p, data_only=True)
    return sheet_rows(wb[sheet]) if sheet in wb.sheetnames else ([], [], [])


def exec_summary(outdir, ty):
    """경영진 요약 문장 (결과 파일 숫자로 · 계산 없이 옮김)"""
    O, _ = org(outdir)
    out = []
    _, _, D = _xlsx_read(outdir, '②_대사.xlsx', '대사')
    tot = next((r for r in D if r.get('RPU 코드') == '합계'), None)
    ev = [r for r in D if str(r.get('구간') or '') == '사업 이벤트']
    now = sum(v['now'] for v in O.values())
    if tot:
        dec = nz(tot['결정 수요 사람'])
        out.append(f'{ty}년 필요 인원 (결정 수요) 은 사람 {fmt1(dec)}명' + (f' · 현원 {now:,.0f}명 대비 {dec - now:+,.1f}명' if now else '')
                   + (f' · 이 중 {", ".join(nm(r["RPU 코드"], O) or r["RPU 코드"] for r in ev)} {fmt1(sum(nz(r["결정 수요 사람"]) for r in ev))}명은 사업 이벤트 별도 줄' if ev else '')
                   + f' · AI 가 맡는 디지털 FTE {fmt1(nz(tot["결정 수요 디지털"]))} 는 따로 봄.')
    _, hb, SB = _xlsx_read(outdir, '③_공급.xlsx', '기본')
    _, _, SU = _xlsx_read(outdir, '③_공급.xlsx', '상승')
    k2 = next((h for h in reversed(hb) if re.fullmatch(r'공급 \d{4}', h)), None)
    sb = next((r for r in SB if r.get('RPU 코드') == '합계'), None)
    su = next((r for r in SU if r.get('RPU 코드') == '합계'), None)
    if sb and k2:
        out.append(f'채용하지 않으면 {ty}년 말 인원은 {fmt1(nz(sb[k2]))}명 (퇴직률 상승 시 {fmt1(nz(su[k2])) if su else "-"}명) 으로 줄어듦.')
    gp = os.path.join(outdir, '④_Gap대안.md')
    if os.path.exists(gp):
        B = md_blocks(open(gp, encoding='utf-8').read())
        gt = md_table_after(B, '1. Gap 표')
        if gt:
            rr = [dict(zip(gt[0], r)) for r in gt[1]]
            t = next((r for r in rr if r.get('RPU 코드') == '합계'), {})
            rs = [r for r in rr if r.get('RPU 코드') != '합계']
            short = sorted([r for r in rs if (num(r.get('Gap')) or 0) > 0.5], key=lambda r: -(num(r['Gap']) or 0))
            spare = [r for r in rs if (num(r.get('Gap')) or 0) < -0.5]
            out.append(f'수급 Gap 은 {fmt1(num(t.get("Gap")) or 0)}명 부족 · 부족이 큰 곳: '
                       + ' · '.join(f'{nm(r["RPU 코드"], O) or r["RPU 코드"]} {num(r["Gap"]):+.1f}' for r in short[:3])
                       + (f' · 여유 {len(spare)}개 조직은 감원 없이 재배치 경로로 다룸.' if spare else '.'))
            at = md_table_after(B, '2. 대안표')
            if at:
                alt = [dict(zip(at[0], r)) for r in at[1]]
                R = [a for a in alt if str(a.get('우선순위') or '').strip() == 'R']
                if R:
                    out.append('1순위 (R) 대안: ' + ' · '.join(f'{nm(a["RPU 코드"], O) or a["RPU 코드"]} {plain(a.get("4B 대안", ""))} ({plain(a.get("시점", ""))})' for a in R[:4]) + '.')
    return out


def build_docx(outdir):
    """리포트_<회사>_<연도>.md → 같은 이름 .docx (표지 · 경영진 요약 · 본문 · 승인) · python-docx 없으면 건너뜀"""
    if not _docx():
        return None
    md = sorted(glob.glob(os.path.join(outdir, '리포트_*.md')))
    if not md:
        return None
    src = md[-1]
    m = re.match(r'리포트_(.+)_(\d{4})\.md$', os.path.basename(src))
    co, ty = (m.group(1), m.group(2)) if m else ('', '')
    text = open(src, encoding='utf-8').read()
    approved = bool(re.search(r'\|\s*CP5\s*\|\s*확정\s*\|', text))
    state = 'CP5 리포트 공유 승인 완료' if approved else '초안 · CHRO 공유 승인 (CP5) 전 · 폴더 밖으로 내보내지 않음'
    w = W(f'{co} {ty} 인력계획 리포트 · {state}')
    for _ in range(7):
        w.para('')
    w.para(co, 14, True, w.sub)
    w.para(f'{ty} 전략적 인력계획 리포트', 26, True, after=6)
    w.para('사업전략 → 인력수요 → 인력공급 → 수급 Gap 및 확보 · 육성 대안', 12, color=w.sub, after=30)
    w.para('작성: 인사기획팀 · pwc-hr-harness 기본 팀 결과 (숫자는 결과 파일에서 옮김)', 10, color=w.sub)
    w.para(f'만든 날: {datetime.date.today():%Y-%m-%d} · 원본: {os.path.basename(src)}', 10, color=w.sub)
    w.para(f'상태: {state}', 10, True, w.ink if approved else w.acc)
    w.page_break()
    B = md_blocks(text)
    w.heading('경영진 요약', 1)
    msgs, on = [], False
    for k, v in B:
        if k.startswith('h'):
            on = v.startswith('1.') or v.startswith('핵심 메시지')
            continue
        if on and k == 'bullet' and not v.startswith('('):
            msgs.append(v)
    auto = exec_summary(outdir, ty)
    if auto:
        w.para('결과 한눈에 (결과 파일 숫자)', 11, True)
        for x in auto:
            w.para('• ' + x, 10, after=3)
    if msgs:
        w.para('핵심 메시지 (리드 작성 · 리포트 1절)', 11, True)
        for x in msgs:
            w.para('• ' + x, 10, after=3)
    kpi = next((v for k, v in B if k == 'table' and v[0][:2] == ['지표', '값']), None)
    if kpi:
        w.para('핵심 숫자', 11, True)
        w.table(*kpi, size=9)
    w.page_break()
    body, started = [], False
    for k, v in B:
        if k == 'h2' and re.match(r'\d+\.', v):
            started = not v.startswith('1.')
        if k == 'h2' and v.startswith('승인'):
            started = True
        if started:
            body.append((k, v))
    w.blocks(body)
    dst = src[:-3] + '.docx'
    w.save(dst, f'{co} {ty} 인력계획 리포트')
    return dst


# ---------------------------------------------------------------------------
# 내보내기 (md → 읽기용 워드 · 엑셀 사본)
# ---------------------------------------------------------------------------
EXPORT_XLSX = ('①_RPU보조표', '①_열대응표', '결정로그', '검증로그')
BANNER = '읽기용 사본 · 여기 고친 내용은 계산에 반영되지 않음 · 바꾸려면 Claude 에게 말하거나 CP 답으로'


def export_one(outdir, name, buf=None):
    """buf 가 io.BytesIO 면 파일로 쓰지 않고 거기에 담아 (이름, 바이트) 를 돌려줌 (뷰어 내려받기 · 2026-11-16)"""
    src = os.path.join(outdir, name if name.endswith('.md') else name + '.md')
    if not os.path.exists(src):
        return None
    dst_dir = os.path.join(outdir, '내보내기')
    if buf is None:
        os.makedirs(dst_dir, exist_ok=True)
    base = os.path.basename(src)[:-3]
    B = md_blocks(open(src, encoding='utf-8').read())
    title = next((v for k, v in B if k == 'h1'), base)
    when = f'{datetime.datetime.now():%Y-%m-%d %H:%M}'
    if base.startswith(EXPORT_XLSX):
        wb = openpyxl.Workbook(); wb.remove(wb.active)
        tabs = [v for k, v in B if k == 'table']
        heads = []
        cur = title
        for k, v in B:
            if k.startswith('h'):
                cur = v
            elif k == 'table':
                heads.append(cur)
        used = set()
        for (h, rows), hd in zip(tabs, heads):
            nm_ = re.sub(r'[\\/*?:\[\]]', '', re.sub(r'^\d+\.\s*', '', hd))[:28].strip() or '표'
            while nm_ in used:
                nm_ = nm_[:26] + str(len(used))
            used.add(nm_)
            cols = [dict(h=x, w=max(8, min(46, max([len(x)] + [len(str(r[i])) if i < len(r) else 0 for r in rows]) * 1.6 + 2)), wrap=True) for i, x in enumerate(h)]
            table_sheet(wb, nm_, hd, f'{BANNER} · 원본 {base}.md · {when}', cols,
                        [(None, [[num(c) if num(c) is not None and not re.search(r'[A-Za-z가-힣]', c) else plain(c) for c in r] for r in rows], None)], freeze_cols=1)
        if not wb.sheetnames:
            return None
        if buf is not None:
            wb.save(buf)
            return base + '.xlsx', buf.getvalue()
        dst = os.path.join(dst_dir, base + '.xlsx')
        wb.save(dst)
        return dst
    if not _docx():
        return None
    w = W(f'{base} · {BANNER}')
    w.para(title, 18, True, after=4)
    w.para(f'{BANNER} · 원본 {base}.md · {when}', 9, color=w.acc, after=10)
    w.blocks([b for b in B if b[0] != 'h1'])
    if buf is not None:
        w.d.core_properties.title = title
        w.d.save(buf)
        return base + '.docx', buf.getvalue()
    dst = os.path.join(dst_dir, base + '.docx')
    w.save(dst, title)
    return dst


def export(outdir, which='all'):
    names = [os.path.basename(p) for p in sorted(glob.glob(os.path.join(outdir, '*.md'))) if not os.path.basename(p).startswith(('_', '리포트_'))] if which == 'all' else [which]
    return [x for x in (export_one(outdir, n) for n in names) if x]


MIME = {'.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.md': 'text/markdown;charset=utf-8', '.zip': 'application/zip'}


def downloads(outdir, co):
    """뷰어에 넣을 내려받기 (이 시점 사본) → {파일 이름: [dict(fn, href, label)]} · 'ZIP' = 전체 묶음
    엑셀 · 워드는 그대로 · md 는 워드 (글) 또는 엑셀 (표) 로 바꾼 것 + md 원본 · 실패한 파일은 건너뜀"""
    import base64, io, zipfile
    uri = lambda fn, b: f'data:{MIME[os.path.splitext(fn)[1]]};base64,' + base64.b64encode(b).decode('ascii')
    out, zbuf = {}, io.BytesIO()
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    with zipfile.ZipFile(zbuf, 'w', zipfile.ZIP_DEFLATED) as z:
        for p in sorted(glob.glob(os.path.join(outdir, '*'))):
            fn = os.path.basename(p)
            ext = os.path.splitext(fn)[1]
            if fn.startswith(('_', '~$')) or ext not in ('.xlsx', '.docx', '.md') or not os.path.isfile(p):
                continue
            try:
                b = open(p, 'rb').read()
            except OSError:
                continue
            z.writestr(fn, b)
            if ext in ('.xlsx', '.docx'):
                out[fn] = [dict(fn=fn, href=uri(fn, b), label='엑셀 받기' if ext == '.xlsx' else '워드 받기')]
                continue
            items = []
            if not fn.startswith('리포트_'):
                try:
                    r = export_one(outdir, fn, buf=io.BytesIO())
                    if r:
                        items.append(dict(fn=r[0], href=uri(r[0], r[1]), label=('엑셀' if r[0].endswith('.xlsx') else '워드') + '로 받기'))
                        z.writestr('읽기용_워드엑셀/' + r[0], r[1])
                except Exception:   # noqa: BLE001
                    pass
            items.append(dict(fn=fn, href=uri(fn, b), label='md 원본'))
            out[fn] = items
    return out   # 여러 파일 묶음 (zip) 은 뷰어가 고른 파일로 브라우저에서 만듦 (HTML 크기를 줄임)


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('사용: report_xlsx.py decorate <xlsx> | gap <outputs/회사> | docx <outputs/회사> | export <outputs/회사> [파일 | all]')
        sys.exit(1)
    cmd, arg = sys.argv[1], sys.argv[2]
    if cmd == 'decorate':
        print('보고서 시트 붙임' if decorate(arg) else '대상 아님')
    elif cmd == 'gap':
        print(build_gap(arg) or '④_Gap대안.md 1절 Gap 표가 없음')
    elif cmd == 'docx':
        print(build_docx(arg) or '리포트 md 가 없거나 python-docx 없음 (python -m pip install python-docx)')
    elif cmd == 'export':
        r = export(arg, sys.argv[3] if len(sys.argv) > 3 else 'all')
        print('\n'.join(r) if r else '내보낼 파일 없음 (워드는 python-docx 필요)')
