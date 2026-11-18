# -*- coding: utf-8 -*-
"""hr-view · 하네스 구성과 산출물을 한 장의 HTML 뷰어로 만든다 (인터넷 · 계정 없이 브라우저에서 열림).

사용:
  python -X utf8 .claude/hr-lib/render.py            # 작업 폴더 전체 → hr-viewer.html
  python -X utf8 .claude/hr-lib/render.py --no-open  # 만들기만 (브라우저를 열지 않음)

만들 때마다 기본 브라우저로 연다 (Windows · Mac). 이미 열린 뷰어 탭은 2초마다 새 판인지 보고 스스로 다시 읽고,
새 탭이 열리면 지난 탭은 스스로 닫힌다 (탭이 쌓이지 않게). 시험 · 시연 녹화처럼 자동으로 돌릴 때는 HR_VIEWER_OPEN=0 (또는 CI) 이면 열지 않는다.

화면 (왼쪽 메뉴): 작업 = 개요 · 진행 · CP · 대시보드 / 결과 = 산출물 · 문서 (outputs/<회사>/*.md 읽기) · 리포트 · 검증 / 하네스 관리 = 하네스 · 변경점 · 회귀 · 설계서
(outputs/<회사>/자사버전_설계서.md · /hr-customize 가 남김 · 아래 참고 문서 한 줄 = 양식/ 의 실데이터 전환 점검표 · GitHub 가져가기 이름 · 경로만) · 부록 화면은 2026-11-16 에 뺌
첫 화면 「개요」 = 「지금 정할 것」 4요소 (정할 것 · 근거 파일 · 선택지 · 답 형식) + 숫자 타일 + 단계 진행. 다른 화면에서는 위쪽 띠로 보인다.
근거 파일을 누르면 「문서」 화면에서 그 절로 간다. 회사가 여럿이면 (outputs/K-co · outputs/자사 …) 위쪽 회사 단추로 바꿔 본다.
결과 파일을 읽기만 하고 고치지 않는다.
「하네스」 화면 = CLAUDE.md · 에이전트 · Skill · 훅 · MCP · 권한 (절마다 🔒 잠금 · ✏️ 변형 메뉴). 「수정」은 브라우저 안 초안일 뿐이고 파일을 쓰지 않는다.
초안의 차이를 「/hr-customize 하네스 수정 요청」 문장으로 복사하면 Claude 가 적용한다 (잠금 검사 · K-Co. 회귀 · 변경 이력 · 뷰어 갱신).
"""
import os, sys, re, json, html, glob, difflib, fnmatch, datetime, argparse, webbrowser, subprocess, pathlib
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hrlib as H
import openpyxl

E = html.escape
CP_FILES = [('CP1', '열 대응 확인', '①_열대응표.md', '인사기획 실무 담당'), ('CP2', '동인 선정', '②_수요파라미터.md', '인사기획팀장'),
            ('CP3', '공급 가정', '③_공급가정.md', '인사기획팀장'), ('CP4', 'Gap 대안 승인', '④_Gap대안.md', '팀장 + R 행 소관 본부장'),
            ('CP5', '리포트 공유', '리포트_*.md', 'CHRO')]
FILE_DESC = {   # 산출물 표 「무엇」 칸 (2026-11-16)
    '①_열대응표.md': '회사 엑셀 열 ↔ 기본 열 대응 · CP1 승인 원본', '①_사업전략분석.md': '사업계획 → 인력 영향 요인 · RPU · Pivotal 후보',
    '①_RPU보조표.md': 'RPU 목록 · 유형 · 스킬 요구', '②_수요파라미터.md': '수요 판단 값 1~6절 · CP2 승인 원본',
    '②A_매출이익.xlsx': '방식 A · 전사 적정 인원 범위 · HR Index', '②B_인력동인.xlsx': '방식 B · RPU별 동인 · 인원',
    '②C_업무량.xlsx': '방식 C · 업무량 FTE · 사람 · 디지털', '②_대사.xlsx': '세 값 대사 · 결정 수요 (④ 기준선)',
    '③_공급가정.md': '공급 가정 S-1~S-5 · CP3 승인 원본', '③_공급.xlsx': '채용 없는 공급 · 기본 · 상승',
    '④_Gap대안.md': 'Gap · 4B 대안 · 재배치 경로 · CP4 승인 원본', '④_Gap분석.xlsx': 'Gap 표 · 대안표 · 재배치 경로 (보고용)',
    '리포트_*.md': '리포트 7섹션 · CP5 승인 원본', '리포트_*.docx': '경영진 보고용 워드 (CP5 전 「초안」)',
    '결정로그.md': 'CP 답 · 가정 변경 기록', '검증로그.md': '검증 5기준 · 미해결 · 인정', '_멈춤.md': '지금 멈춘 자리 (4요소)'}


def file_desc(n):
    if n in FILE_DESC:
        return FILE_DESC[n]
    for k, v in FILE_DESC.items():
        if '*' in k and fnmatch.fnmatch(n, k):
            return v
    return ''


STAGES = [('①', '사업전략 · 요구사항 분석', ['①_열대응표.md', '①_사업전략분석.md', '①_RPU보조표.md']),
          ('②', '인력수요 예측 (방식 A 매출 · 이익 · B 인력동인 · C 업무량 · 대사)', ['②_수요파라미터.md', '②A_매출이익.xlsx', '②B_인력동인.xlsx', '②C_업무량.xlsx', '②_대사.xlsx']),
          ('③', '인력공급 예측 (채용 없음 · 기본 · 상승 시나리오)', ['③_공급가정.md', '③_공급.xlsx']),
          ('④', '수급 Gap 분석 · 확보 · 육성 (Gap 4유형 · 4B 대안)', ['④_Gap대안.md', '④_Gap분석.xlsx']),
          ('리포트', '7섹션', ['리포트_*.md', '리포트_*.docx']),
          ('회귀', 'K-Co. 회귀 테스트 (customize.py regress)', ['_회귀테스트.md'])]

# ---------------------------------------------------------------------------
# 용어 풀이 (화면에 코드만 두지 않음 · 잠금 용어 그대로 · 뜻을 지어내지 않음)
# 회사마다 다른 것 (RPU 이름 · 본부 · 공급 가정 S-1~S-5) 은 company() 가 결과 파일에서 읽어 더한다
# ---------------------------------------------------------------------------
GLOSS = dict(
    cp={'CP1': dict(n='열 대응', who='실무 담당', d='회사 파일 열을 기본 열에 맞춘 열 대응표를 사람이 확인'),
        'CP2': dict(n='동인 선정', who='팀장', d='수요 방식 B 의 인력동인 선정 · 사업계획 보정을 사람이 승인'),
        'CP3': dict(n='공급 가정', who='팀장', d='공급 가정 S-1~S-5 에 사람이 답한 뒤에만 공급을 계산'),
        'CP4': dict(n='Gap 대안', who='팀장 + 본부장', d='Gap 4B 대안 · 재배치 경로 · 우선순위를 사람이 승인'),
        'CP5': dict(n='리포트 공유', who='CHRO', d='리포트를 폴더 밖으로 공유해도 되는지 승인 (승인 전에는 내보내지 않음)')},
    cp_def='체크포인트(CP) = 에이전트가 멈추고 사람 승인을 받는 자리 · 답은 확정 · 추가 확인 · 보류 · 이관',
    step={'①': '사업전략 · 요구사항 분석', '②': '인력수요 예측', '③': '인력공급 예측', '④': '수급 Gap 분석 · 확보 · 육성'},
    step_short={'1': '사업전략', '2': '수요', '3': '공급', '4': 'Gap', '①': '사업전략', '②': '수요', '③': '공급', '④': 'Gap', 'all': '전 단계', 'report': '리포트'},
    method={'A': '매출 · 이익 연계 (HCROI 역산)', 'B': '인력동인', 'C': '업무량 (사람 FTE + 디지털 FTE)'},
    method_short={'A': '매출 · 이익', 'B': '인력동인', 'C': '업무량'},
    b4={'Buy': '채용', 'Build': '육성 (재배치 포함)', 'Borrow': '외부 활용', 'Bot': '자동화'},
    gap4=['수량', '구조', '시점', 'AI 활용'],
    crit=['수치 재계산', '가정 명시', '출처', '범위 밖 거절', '재실행 일관성'],
    term={'대사': '방식 B · C 값을 맞대어 차이 원인을 판정하는 것',
          '결정 수요': '대사로 정한 RPU별 사람 수요 (디지털 FTE 는 별도 줄)',
          '예외': '데이터 이상으로 멈춤 (CP 아님)',
          '관문': '검증 담당의 5기준 점검 · 미해결이 남으면 다음 단계로 넘어가지 않음',
          '인정': '관문 항목을 사람이 결정로그로 인정 (미해결로 세지 않음)',
          '미확인': '범위 밖 거절 행동 테스트를 아직 묻지 않음',
          'RPU': '인력 산정 단위 (수요 · 공급 · Gap 을 세는 조직 묶음)',
          'Pivotal Role': '핵심 직무',
          'G1~G5': '직급 단계',
          '기본 · 상승': '공급 시나리오 · 기본 = 과거 퇴직률 그대로 · 상승 = 퇴직률 상승 가정',
          '자율성 L1': '제안만 하고 확정은 사람이 함'})
GL = dict(rpu={}, org={}, sup={}, rx=None)   # 지금 회사의 풀이 (company() 가 채움 · 하네스 화면은 회사 없이 공통 풀이만)


def gloss_title(tok):
    """코드 한 개 → 풀이 한 줄 (모르면 '')."""
    g = GLOSS
    if tok in g['cp']:
        c = g['cp'][tok]
        return f'체크포인트(CP) {tok} · {c["n"]} · 승인 {c["who"]} · {c["d"]}'
    if tok in GL['sup']:
        return f'공급 가정 {tok} · {GL["sup"][tok]} (CP3 에서 사람이 답함)'
    if tok in GL['rpu']:
        o = GL['org'].get(tok)
        return f'RPU {tok} · {GL["rpu"][tok]}' + (f' · {o}' if o else '') + ' (RPU = 인력 산정 단위)'
    if tok in g['b4']:
        return f'4B 대안 {tok} · {g["b4"][tok]}'
    if tok in g['step']:
        return f'{tok} {g["step"][tok]}'
    if tok in g['method']:
        return f'수요 방식 {tok} · {g["method"][tok]}'
    if re.fullmatch(r'G[1-5]', tok):
        return f'{tok} · 직급 단계'
    if tok.startswith('Pivotal'):
        return 'Pivotal Role · 핵심 직무'
    if tok == '자율성 L1':
        return '자율성 L1 · 제안만 하고 확정은 사람이 함'
    if re.fullmatch(r'S-[1-5]', tok):
        return f'공급 가정 {tok} (CP3 에서 사람이 답함)'
    return ''


def set_gloss(rpu=None, org=None, sup=None):
    GL.update(rpu=dict(rpu or {}), org=dict(org or {}), sup=dict(sup or {}))
    codes = '|'.join(re.escape(c) for c in sorted(GL['rpu'], key=len, reverse=True) if re.fullmatch(r'[A-Z]{1,3}\d{1,3}', c))
    alts = [r'(?<![A-Za-z0-9])CP[1-5](?![0-9])', r'(?<![A-Za-z0-9])S-[1-5](?![0-9])']
    if codes:
        alts.append(r'(?<![A-Za-z0-9_])(?:' + codes + r')(?![A-Za-z0-9_])')
    alts += [r'(?<![A-Za-z])(?:Buy|Build|Borrow|Bot)(?![A-Za-z])', r'Pivotal(?: Role)?', r'자율성 L1',
             r'[①②③④](?![_A-Za-z0-9])', r'(?<![A-Za-z0-9_./\-])G[1-5](?![0-9A-Za-z])',
             r'(?<![A-Za-z0-9_./\-①②③④가-힣])[ABC](?![A-Za-z0-9_\-가-힣])']
    GL['rx'] = re.compile('|'.join(alts))


def gloss(s):
    """이스케이프된 HTML 의 글자 부분에서 코드를 찾아 풀이 말풍선 (abbr) 으로 감쌈 · 태그 · <code> 안은 그대로."""
    rx = GL['rx']
    if not rx or not s:
        return s
    out, skip = [], 0
    for part in re.split(r'(<[^>]+>)', s):
        if part.startswith('<'):
            t = part[:6].lower()
            if t.startswith(('<code', '<abbr', '<pre')):
                skip += 1
            elif t.startswith(('</code', '</abbr', '</pre')):
                skip = max(0, skip - 1)
            out.append(part)
            continue
        if skip:
            out.append(part)
            continue
        def sub(m):
            tok = m.group(0)
            t = gloss_title(tok)
            if not t:
                return tok
            extra = ''
            if tok in GLOSS['b4'] and not re.match(r'\s*\(', m.string[m.end():]):   # 4B 영어 낱말 옆에 한국어 (이미 「Buy (채용)」 이면 그대로)
                extra = f'<span class="gl">{E(GLOSS["b4"][tok].split(" (")[0])}</span>'
            rp = f' data-rpu="{E(tok)}"' if t.startswith('RPU ') else ''   # 뷰어 JS 가 첫 등장에 이름을 작게 붙임 (코드만 두지 않음)
            return f'<abbr class="g"{rp} title="{E(t)}">{tok}</abbr>{extra}'
        out.append(rx.sub(sub, part))
    return ''.join(out)


def head_cell(c):
    """표 머리 칸이 코드 하나 (B · C · 결정 …) 이면 뜻을 작게 붙임."""
    k = c.strip()
    if k in GLOSS['method_short']:
        return f'{E(k)}<span class="gl">{E(GLOSS["method_short"][k])}</span>'
    if k == '결정':
        return '결정<span class="gl">결정 수요 (사람)</span>'
    if re.fullmatch(r'G[1-5]', k):
        return f'<abbr class="g" title="{E(k)} · 직급 단계">{E(k)}</abbr>'
    return None


RPU_CODE_COLS = ('RPU', 'RPU 코드', '코드')
RPU_NAME_COLS = ('이름', 'RPU 이름', 'RPU(산정 단위)', 'RPU (산정 단위)', '직무군', 'RPU명')


def table_ctx(head):
    """표 머리 → (RPU 코드 칸 번호 · 이름을 붙일지, 단계 칸 번호)."""
    h = [plain(x) for x in head]
    ci = next((i for i, x in enumerate(h) if x in RPU_CODE_COLS), -1)
    has_name = any(x in RPU_NAME_COLS for x in h) or (ci >= 0 and h[ci] == 'RPU 코드' and 'RPU' in h)
    si = next((i for i, x in enumerate(h) if x == '단계'), -1)
    return (ci if not has_name else -1), si


def cell_extra(i, c, ctx):
    ci, si = ctx
    k = c.strip()
    if i == ci and k in GL['rpu']:
        return f'<span class="gl">{E(GL["rpu"][k])}</span>'
    if i == si and k in GLOSS['step_short']:
        return f'<span class="gl">{E(GLOSS["step_short"][k])}</span>'
    return ''


# ---------------------------------------------------------------------------
# 마크다운 → HTML (표 · 목록 · 제목 · 굵게 · 코드 · 인용 · 주석 제거)
# ---------------------------------------------------------------------------
NUM = re.compile(r'^[-−+]?[\d,]+(\.\d+)?\s?(%|명|배)?$')   # 숫자 칸 (오른쪽 정렬)
DATE = re.compile(r'^\d{4}-\d{2}-\d{2}(\s\d{1,2}:\d{2})?$')   # 날짜 칸 (줄바꿈 없음)
DOCS = []   # 지금 회사 폴더의 .md 이름들 (본문에 이름이 나오면 「문서」 탭 링크로 바꿈) · company() 가 채움


def doc_link(name, label, sec=''):
    return f'<a href="#" class="dl" data-doc="{E(name)}" data-sec="{E(sec)}">{label}</a>'


def autolink(s):
    """이스케이프된 문장 안의 산출물 이름 (outputs/<회사>/ 앞머리 포함) → 문서 링크."""
    if not DOCS:
        return s
    pat = r'(?:outputs/[^\s/<>&]+/)?(' + '|'.join(re.escape(E(n)) for n in sorted(DOCS, key=len, reverse=True)) + r')'
    return re.sub(pat, lambda m: doc_link(html.unescape(m.group(1)), m.group(0)), s)


def inline(s):
    links = []
    def keep(m):   # [글](주소) 는 따로 빼 두었다가 마지막에 돌려 넣음 (주소 안 파일 이름이 두 번 바뀌지 않게)
        label, url = m.group(1), m.group(2).strip()
        base, _, frag = url.partition('#')
        name = os.path.basename(base.replace('\\', '/'))
        if name in DOCS:
            links.append(doc_link(name, E(label), frag.replace('-', ' ')))
        elif re.match(r'^https?://', url):
            links.append(f'<a href="{E(url)}" target="_blank" rel="noopener">{E(label)}</a>')
        else:
            links.append(f'{E(label)} <code>{E(url)}</code>')
        return f'\x00{len(links) - 1}\x00'
    s = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', keep, s)
    s = E(s)
    s = re.sub(r'`([^`]+)`', r'<code>\1</code>', s)
    s = re.sub(r'\*\*([^*]+)\*\*', r'<b>\1</b>', s)
    s = re.sub(r'(?<![*\w])\*([^*\s][^*]*?)\*(?![*\w])', r'<i>\1</i>', s)
    s = s.replace('&lt;br&gt;', '<br>')
    s = re.sub(r'(✓)', r'<span class="ok">\1</span>', s)
    s = autolink(s)
    return gloss(re.sub(r'\x00(\d+)\x00', lambda m: links[int(m.group(1))], s))


def plain(s):
    """제목 글자만 (굵게 · 코드 · 링크 표시를 걷어 냄) · 절 찾기 기준."""
    s = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', s)
    return re.sub(r'[*`]', '', s).strip()


FOCUS_RE = re.compile(r'확인 필요|물을 것|물어 볼|정해 주실|짚어 둘|정할 것|결정할 것|확인할 것')


def sections(text):
    """[(제목, 그 절 본문)] · 제목 앞 머리말은 제목 ''."""
    text = re.sub(r'<!--.*?-->', '', text, flags=re.S)
    out, head, buf, fence = [], '', [], False
    for ln in text.splitlines():
        if ln.strip().startswith('```'):
            fence = not fence
        m = None if fence else re.match(r'^\s*#{1,4}\s+(.*)', ln)
        if m:
            out.append((head, '\n'.join(buf)))
            head, buf = plain(m.group(1)), []
        else:
            buf.append(ln)
    out.append((head, '\n'.join(buf)))
    return out


def focus(text):
    """사람이 볼 절: 「확인 필요 · 사람에게 물을 것 …」 제목 → 없으면 「사람 답」 칸이 있는 표의 절 · 그리고 「승인」 절."""
    secs = sections(text)
    f = [dict(sec=h, label=re.sub(r'\s*\([^)]*\)\s*$', '', h)) for h, b in secs if h and FOCUS_RE.search(h)]
    if not f:
        f = [dict(sec=h, label=f'{h} · 사람 답 칸') for h, b in secs if h and re.search(r'^\|.*\|\s*사람 답\s*\|', b, re.M)]
    f += [dict(sec=h, label='승인 칸') for h, b in secs if h == '승인']
    return f


def md(text):
    text = re.sub(r'<!--.*?-->', '', text, flags=re.S)
    lines = text.splitlines()
    out, i, in_list = [], 0, []
    def close_lists(level=0):
        while len(in_list) > level:
            out.append('</ul>')
            in_list.pop()
    while i < len(lines):
        ln = lines[i]
        s = ln.strip()
        if s.startswith('|') and i + 1 < len(lines) and re.match(r'^\|\s*:?-{2,}', lines[i + 1].strip()):
            close_lists()
            head = [c.strip() for c in s.strip('|').split('|')]
            ctx = table_ctx(head)
            out.append('<div class="tw"><table><thead><tr>' + ''.join(f'<th>{head_cell(c) or inline(c)}</th>' for c in head) + '</tr></thead><tbody>')
            i += 2
            held = []   # 옛 양식: 표 아래 안내 줄 뒤에 같은 열 수의 행이 이어 붙음 → 행은 표에 잇고 안내 줄은 표 뒤에 보임 (화면 표시만)
            while i < len(lines):
                if not lines[i].strip().startswith('|'):
                    k, skip = i, []
                    while k < len(lines) and len(skip) < 4 and (not lines[k].strip() or re.match(r'^\s*[-*]\s+', lines[k])):
                        skip.append(lines[k])
                        k += 1
                    nxt = lines[k].strip() if k < len(lines) else ''
                    if (skip and nxt.startswith('|') and len(nxt.strip('|').split('|')) == len(head)
                            and not (k + 1 < len(lines) and re.match(r'^\|\s*:?-{2,}', lines[k + 1].strip()))):
                        held += skip
                        i = k
                        continue
                    break
                cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                out.append('<tr>' + ''.join((f'<td class="bad">{inline(c)}</td>' if c == '미해결' or c.startswith('위반') or c.startswith('**없')
                                             else f'<td class="ok">{inline(c)}</td>' if c in ('통과', '승인', '일치', '해결') or c.startswith('인정 (결정로그') else f'<td class="warn">{inline(c)}</td>' if c == '미확인'
                                             else f'<td>{E(c)}{cell_extra(j, c, ctx)}</td>' if cell_extra(j, c, ctx) else f'<td class="n">{inline(c)}</td>' if NUM.match(c) else f'<td class="dt">{E(c)}</td>' if DATE.match(c) else f'<td>{inline(c)}</td>')
                                            for j, c in enumerate(cells)) + '</tr>')
                i += 1
            out.append('</tbody></table></div>')
            lines[i:i] = [x for x in held if x.strip()]
            continue
        m = re.match(r'^(#{1,4})\s+(.*)', s)
        if m:
            close_lists()
            lv = min(len(m.group(1)) + 1, 5)
            out.append(f'<h{lv} data-h="{E(plain(m.group(2)))}">{inline(m.group(2))}</h{lv}>')
        elif re.match(r'^-{3,}$', s) or re.match(r'^\*{3,}$', s):
            close_lists()
            out.append('<hr>')
        elif re.match(r'^(\s*)[-*]\s+', ln) or re.match(r'^(\s*)\d+\.\s+', ln):
            ind = len(ln) - len(ln.lstrip())
            level = ind // 2 + 1
            if len(in_list) < level:
                while len(in_list) < level:
                    out.append('<ul>')
                    in_list.append(1)
            else:
                close_lists(level)
            out.append('<li>' + inline(re.sub(r'^\s*([-*]|\d+\.)\s+', '', ln)) + '</li>')
        elif s.startswith('>'):
            close_lists()
            out.append(f'<blockquote>{inline(s.lstrip("> "))}</blockquote>')
        elif s.startswith('```'):
            close_lists()
            j = i + 1
            buf = []
            while j < len(lines) and not lines[j].strip().startswith('```'):
                buf.append(lines[j])
                j += 1
            out.append('<pre>' + E('\n'.join(buf)) + '</pre>')
            i = j
        elif s:
            close_lists()
            out.append(f'<p>{inline(s)}</p>')
        else:
            close_lists()
        i += 1
    close_lists()
    return '\n'.join(out)


def frontmatter(text):
    m = re.match(r'^---\s*\n(.*?)\n---\s*\n(.*)$', text, re.S)
    if not m:
        return {}, text
    fm = {}
    for ln in m.group(1).splitlines():
        if ':' in ln:
            k, v = ln.split(':', 1)
            fm[k.strip()] = v.strip().strip('"')
    return fm, m.group(2)


REP_HL = {'FCE4DF': 'hl-r', 'FFF2CC': 'hl-y', 'E8F3E8': 'hl-g', 'EFEFEF': 'hl-n', 'FBE3D6': '', 'FBF7F2': ''}


def _cell_txt(c):
    """보고서 시트 칸 → 엑셀에 보이는 모양 그대로 (number_format 을 따름)"""
    v = c.value
    if v is None:
        return ''
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return str(v)
    f = c.number_format or ''
    if '%' in f:
        s = f'{v * 100:+.1f}%' if f.startswith('+') else f'{v * 100:.1f}%'
    elif '0.0000' in f:
        s = f'{v:.4f}'
    elif '0.000' in f:
        s = f'{v:.3f}'
    elif '0.00' in f:
        s = f'{v:,.2f}'
    elif '0.0' in f:
        s = f'{v:+,.1f}' if f.startswith('+') else f'{v:,.1f}'
    elif isinstance(v, float) and v != int(v):
        s = f'{v:,.1f}'
    else:
        s = f'{int(round(v)):,}'
    return s


def _rep_summary(sh):
    """「요약」 시트 → 카드 (제목 · 상태 · 핵심 지표 · 절)"""
    rows = list(sh.iter_rows())
    val = lambda r, j: (rows[r][j].value if j < len(rows[r]) else None)
    out = [f'<div class="rsum"><div class="rt">{E(str(val(0, 0) or ""))}</div><div class="rsub">{E(str(val(1, 0) or ""))}</div>']
    i, kp = 2, []
    while i < len(rows) and not str(val(i, 0) or '').startswith('■ '):
        vs = [c for c in rows[i] if c.value is not None]
        if len(vs) == 1 and not kp:
            out.append(f'<div class="rstat">{E(str(vs[0].value))}</div>')
        elif len(vs) >= 2:
            kp.append(rows[i])
        i += 1
    if len(kp) >= 3:
        out.append('<div class="kpi">')
        for j in range(0, 8, 2):
            if j < len(kp[0]) and kp[0][j].value is not None:
                out.append(f'<div class="{"hl" if j == 0 else ""}"><span>{E(str(kp[0][j].value))}</span><b>{E(_cell_txt(kp[1][j]))}</b><span>{E(str(kp[2][j].value or ""))}</span></div>')
        out.append('</div>')
    cur = None
    for r in rows[i:]:
        t = r[0].value
        if t is None:
            continue
        t = str(t)
        if t.startswith('■ '):
            if cur is not None:
                out.append('</ul>')
            out.append(f'<h4 class="rh">{E(t[2:])}</h4><ul class="rl">')
            cur = t
        else:
            out.append(f'<li>{inline(t)}</li>')
    if cur is not None:
        out.append('</ul>')
    out.append('</div>')
    return '\n'.join(out)


def _rep_table(sh):
    """「○○표」 시트 → 표 (본부 머리 · 소계 · 합계 · 엑셀 강조 색 그대로)"""
    rows = list(sh.iter_rows())
    out = [f'<p class="muted" style="margin:2px 0 4px">{E(str(rows[1][0].value or ""))}</p>'] if len(rows) > 1 else []
    hi = next((i for i, r in enumerate(rows) if i > 1 and sum(c.value is not None for c in r) >= 3), None)
    if hi is None:
        return ''
    notes = [str(r[0].value) for r in rows[2:hi] if r[0].value is not None]
    if notes:
        out.append('<ul class="notes">' + ''.join(f'<li>{E(n)}</li>' for n in notes) + '</ul>')
    hd = [str(c.value or '').replace('\n', ' ') for c in rows[hi]]
    n = len([h for h in hd if h]) or len(hd)
    hd = hd[:n]
    out.append('<div class="tw"><table><thead><tr>' + ''.join(f'<th>{E(h)}</th>' for h in hd) + '</tr></thead><tbody>')
    for r in rows[hi + 1:]:
        vs = [c for c in r[:n] if c.value is not None]
        if not vs:
            continue
        if len(vs) == 1 and r[0].value is not None and r[1].value is None:
            out.append(f'<tr class="grp"><td colspan="{n}">{E(str(r[0].value))}</td></tr>')
            continue
        lab = ' '.join(str(c.value or '') for c in r[:2])
        cls = 'tot' if '전사 합계' in lab or str(r[0].value or '') == '합계' else 'sub' if '소계' in lab else ''
        tds = []
        for c in r[:n]:
            num_ = isinstance(c.value, (int, float)) and not isinstance(c.value, bool)
            rgb = str(getattr(c.fill.fgColor, 'rgb', '') or '')[-6:] if c.fill and c.fill.fill_type == 'solid' else ''
            k = ' '.join(x for x in ('n' if num_ else '', 'long' if not num_ and len(str(c.value or '')) > 16 else '', REP_HL.get(rgb.upper(), '')) if x)
            tds.append(f'<td{" class=" + chr(34) + k + chr(34) if k else ""}>{E(_cell_txt(c)) if num_ else inline(_cell_txt(c))}</td>')
        out.append(f'<tr{" class=" + chr(34) + cls + chr(34) if cls else ""}>' + ''.join(tds) + '</tr>')
    out.append('</tbody></table></div>')
    return '\n'.join(out)


def xlsx_html(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    parts = []
    rep = [s for s in wb.worksheets if H.is_report_sheet(s.title)]
    if rep and '요약' in wb.sheetnames:   # 보고서 모양 (2026-11-16) · 요약 → ○○표 → (접힘) 첨부 데이터 · 산식
        for sh in wb.worksheets:
            if sh.title == '요약':
                parts.append(f'<section class="xsh" data-t="요약">{_rep_summary(sh)}</section>')
            elif H.is_report_sheet(sh.title) and sh.title != '산식·가정·출처':
                parts.append(f'<section class="xsh" data-t="{E(sh.title)}">' + _rep_table(sh) + '</section>')
        rest = [s for s in wb.worksheets if not H.is_report_sheet(s.title) or s.title == '산식·가정·출처']
        if rest:
            parts.append('<section class="xsh" data-t="첨부 데이터 · 산식">'
                         + f'<p class="muted" style="margin:2px 0 8px">다음 단계 스크립트가 읽는 원 값 ({E(" · ".join(s.title for s in rest))})</p>'
                         + _plain_sheets([s for s in rest if s.title != '산식·가정·출처'])
                         + ('<h4>산식·가정·출처</h4>' + _rep_table(wb['산식·가정·출처']) if '산식·가정·출처' in wb.sheetnames else '') + '</section>')
        return '<div class="xl">' + '\n'.join(parts) + '</div>'
    return '<div class="xl">' + ''.join(f'<section class="xsh" data-t="{E(sh.title)}">' + _plain_sheets([sh]) + '</section>' for sh in wb.worksheets) + '</div>'


def _plain_sheets(sheets):
    parts = []
    for sh in sheets:
        exact = sh.title == 'HR Index'   # 인당 · HCVA · HCROI 는 스크립트가 2~3자리로 반올림해 저장 · 그 자리 그대로 보임
        rows = [r for r in sh.iter_rows(values_only=True)]
        parts.append(f'<h4>{E(sh.title)}</h4>')
        notes = [str(r[0]) for r in rows[:6] if r and r[0] is not None and all(v is None for v in r[1:])]
        if notes:
            parts.append('<ul class="notes">' + ''.join(f'<li>{E(n)}</li>' for n in notes) + '</ul>')
        body = [r for r in rows if r and sum(v is not None for v in r) >= 2]
        if not body:
            continue
        hd = [H.norm(v) for v in body[0]]
        ctx = table_ctx(hd)
        parts.append('<div class="tw"><table><thead><tr>' + ''.join(f'<th>{head_cell(v) or E(v)}</th>' for v in hd) + '</tr></thead><tbody>')
        for r in body[1:]:
            cells = []
            for j, v in enumerate(r):
                if isinstance(v, float):
                    t = (f'{v:,.3f}'.rstrip('0').rstrip('.') if exact else f'{v:.4f}' if abs(v) < 1 else f'{v:,.1f}')
                    cells.append(f'<td class="n">{t}</td>')
                elif isinstance(v, int):
                    cells.append(f'<td class="n">{v:,}</td>')
                elif v is not None and cell_extra(j, str(v), ctx):
                    cells.append(f'<td>{E(str(v))}{cell_extra(j, str(v), ctx)}</td>')
                else:
                    cells.append(f'<td>{inline("" if v is None else str(v))}</td>')
            cls = ' class="tot"' if r[0] == '합계' else ''
            parts.append(f'<tr{cls}>' + ''.join(cells) + '</tr>')
        parts.append('</tbody></table></div>')
    return '\n'.join(parts)


def find(out, pat):
    if '*' in pat:
        g = sorted(glob.glob(os.path.join(out, pat)))
        return g[0] if g else None
    p = os.path.join(out, pat)
    return p if os.path.exists(p) else None


# ---------------------------------------------------------------------------
# 하네스 탭
# ---------------------------------------------------------------------------
# 잠금 규칙 (🔒) · 변형 메뉴 자리 (✏️)
# 정본은 customize.py (LOCKED_CLAUDE · LOCKED_PHRASES · LOCKED_CODE · ENGINE) 다. 작업 폴더 사본 (.claude/hr-lib/customize.py · /hr-init 이 둠)
# → 플러그인 원본 순서로 읽고, 둘 다 없을 때만 아래 거울 사본을 쓴다 (customize.py 를 고치면 거울 사본도 같이 고칠 것).
_LOCK_MIRROR = dict(
    claude=['## 업무 범위 · 금지'],
    phrases={
        'CLAUDE.md': ['감원 대안을 단독으로 내지 않는다', '폴더 밖 파일을 읽지 않는다', '자율성 L1', 'CP5'],
        '.claude/skills/wfp-lead/SKILL.md': ['| CP1 |', '| CP2 |', '| CP3 |', '| CP4 |', '| CP5 |', '직접 계산하지 않는다', '검증 미해결이 있으면 다음 단계로 가지 않는다'],
        '.claude/agents/wfp-gap.md': ['감원 대안을 단독으로 내지 않는다'],
        '.claude/agents/wfp-verifier.md': ['결과 파일을 고치지 않는다'],
        '.claude/skills/wfp-step2-demand/SKILL.md': ['평균 내지 않', '디지털 FTE 를 사람 FTE 에 더하지 않는다'],
        '.claude/skills/wfp-step3-supply/SKILL.md': ['0으로 읽지 않는다'],
    },
    code={
        '.claude/skills/wfp-step2-demand/scripts/step2.py': ["require_cp(ws, 'CP1'", "require_cp(ws, 'CP2'"],
        '.claude/skills/wfp-step3-supply/scripts/step3.py': ["require_cp(ws, 'CP3'"],
        '.claude/skills/wfp-step4-gap/scripts/step4.py': ["('CP2', '②_수요파라미터.md')", "('CP3', '③_공급가정.md')", "'CP4'", 'CUT'],
        '.claude/skills/wfp-report/scripts/report.py': ["require_cp(ws, 'CP4'", "'CP5'"],
    },
    engine=['.claude/hr-lib/hrlib.py'], src='render.py 거울 사본 (customize.py 를 찾지 못함)')
# 절 제목으로 보는 잠금 (customize.py lock 이 직접 잡지는 않지만 규칙상 바꾸지 않는 칸 · 보수적으로 넓게)
LOCK_HEAD = re.compile(r'업무 범위|금지|범위 밖|하지 않는 일|멈춤|체크포인트|법전|계산식|5기준|^원칙$|^에러$|^이음매$')
# customization-guide 3절 「변형 메뉴」 자리 (파일 → 절 제목) · 잠금 제목보다 먼저 본다 · 그 안의 잠금 문장은 여전히 🔒
EDIT_HEAD = {
    'CLAUDE.md': r'변경 이력',
    '.claude/skills/wfp-step1-strategy/SKILL.md': r'사업전략분석\.md 구성',
    '.claude/skills/wfp-step2-demand/SKILL.md': r'^판단 규칙',
    '.claude/skills/wfp-step4-gap/SKILL.md': r'4B 해법별 질문|조합 규칙|우선순위',
    '.claude/skills/wfp-report/SKILL.md': r'리포트 7섹션|자사 버전에서 바꿀',
}
# 잠금 절 안이라도 이 줄들은 ✏️ (wfp-lead CP 표 행: 기본 승인자 바꾸기 · 자사 CP 행 추가 · 「| CP1 |」~「| CP5 |」 문장 자체는 그대로 잠금 검사)
EDIT_LINES = {'.claude/skills/wfp-lead/SKILL.md': (r'^\|\s*CP\d*\s*\|', 'CP 표 행 · 기본 승인자 바꾸기 · 자사 CP 행 추가 (변형 메뉴 「승인자」 · 「자사 CP 추가」)')}


def lock_rules():
    here = os.path.dirname(os.path.abspath(__file__))
    cands = [os.path.join(here, 'customize.py'),
             os.path.normpath(os.path.join(here, '..', '..', '..', '..', 'skills', 'pwc-hr-harness', 'scripts', 'customize.py'))]
    if os.environ.get('CLAUDE_PLUGIN_ROOT'):
        cands.append(os.path.join(os.environ['CLAUDE_PLUGIN_ROOT'], 'skills', 'pwc-hr-harness', 'scripts', 'customize.py'))
    for p in cands:
        if not os.path.exists(p):
            continue
        try:
            import importlib.util
            spec = importlib.util.spec_from_file_location('hr_customize_rules', p)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return dict(claude=list(mod.LOCKED_CLAUDE), phrases=dict(mod.LOCKED_PHRASES), code=dict(mod.LOCKED_CODE),
                        engine=list(getattr(mod, 'ENGINE', [])), src='customize.py')
        except Exception:
            continue
    return dict(_LOCK_MIRROR)


def _fwd(p):
    return p.replace('\\', '/')


def _vscode(p):
    import urllib.parse
    return 'vscode://file/' + urllib.parse.quote(_fwd(os.path.abspath(p)), safe='/:')


def _secs_raw(lines):
    """원문 줄 → [(시작 줄, 끝 줄, 제목, 수준)] (1부터 · 머리말 frontmatter 제외 · 코드 울타리 안 # 은 제목 아님)."""
    start = 0
    if lines and lines[0].strip() == '---':
        for i in range(1, len(lines)):
            if lines[i].strip() == '---':
                start = i + 1
                break
    heads, fence = [], False
    for i in range(start, len(lines)):
        s = lines[i].strip()
        if s.startswith('```'):
            fence = not fence
        m = None if fence else re.match(r'^(#{1,4})\s+(.*)', s)
        if m:
            heads.append((i + 1, plain(m.group(2)), len(m.group(1))))
    out = []
    if heads and heads[0][0] > start + 1 and any(x.strip() for x in lines[start:heads[0][0] - 1]):
        out.append((start + 1, heads[0][0] - 1, '', 0))
    if not heads and start < len(lines):
        out.append((start + 1, len(lines), '', 0))
    for k, (n, h, lv) in enumerate(heads):
        out.append((n, (heads[k + 1][0] - 1) if k + 1 < len(heads) else len(lines), h, lv))
    return out, start


# 회사 규칙 문단 (/hr-customize 9 · adapt.py harness 가 넣음): <!-- 회사 규칙 · <회사> · … --> … <!-- /회사 규칙 · <회사> -->
CO_START = re.compile(r'^\s*<!--\s*회사 규칙\s*·\s*(.+?)\s*·.*-->\s*$')
CO_END = re.compile(r'^\s*<!--\s*/회사 규칙\s*·\s*(.+?)\s*-->\s*$')


def co_blocks(lines):
    """원문 줄 → [(시작 줄, 끝 줄, 회사)] (1부터 · 마커 줄 포함)."""
    out, open_ = [], None
    for i, ln in enumerate(lines):
        m = CO_START.match(ln)
        if m and not CO_END.match(ln):
            open_ = (i + 1, m.group(1).strip())
            continue
        m = CO_END.match(ln)
        if m and open_:
            out.append((open_[0], i + 1, open_[1]))
            open_ = None
    if open_:
        out.append((open_[0], len(lines), open_[1]))
    return out


def diff_cos(diff):
    """unified diff 글 → (회사 규칙 문단 회사 목록, 회사 문단 밖에서 바뀐 줄 수)."""
    cos, inside, other = [], None, 0
    for ln in diff.splitlines():
        if ln.startswith(('+++', '---', '@@')):
            continue
        body = ln[1:] if ln[:1] in '+- ' else ln
        m = CO_START.match(body)
        if m and not CO_END.match(body):
            inside = m.group(1).strip()
            if ln[:1] in '+-' and inside not in cos:
                cos.append(inside)
            continue
        if CO_END.match(body):
            inside = None
            continue
        if ln[:1] in '+-':
            h = re.search(r'회사 규칙 반영\s*·\s*([^\s|(]+)', body)   # CLAUDE.md 변경 이력 행 「회사 규칙 반영 · <회사>」 도 그 회사 몫
            if h and not inside:
                if h.group(1) not in cos:
                    cos.append(h.group(1))
                continue
            if inside:
                if inside not in cos:
                    cos.append(inside)
            elif body.strip():
                other += 1
    return cos, other


def lock_map(rel, text, rules):
    """한 파일의 절마다 🔒 (hard = customize.py lock 이 잡음 · soft = 규칙상 잠금) · ✏️ (변형 메뉴) 표시와 잠금 문장 줄."""
    lines = text.splitlines()
    secs, fm_end = _secs_raw(lines)
    hard = []   # CLAUDE.md 잠금 절 (customize.section 과 같은 범위: 그 제목 줄부터 다음 「## 」 앞까지)
    if rel == 'CLAUDE.md':
        for h in rules['claude']:
            for i, ln in enumerate(lines):
                if ln.startswith(h):
                    j = next((k for k in range(i + 1, len(lines)) if lines[k].startswith('## ')), len(lines))
                    hard.append((i + 1, j, h.lstrip('# ').strip()))
                    break
    edit_re = EDIT_HEAD.get(rel)
    blocks = co_blocks(lines)
    out = []
    for s, e, h, lv in secs:
        lock, why, edit = '', '', False
        cb = next((b for b in blocks if b[0] <= s <= b[1]), None)
        hh = next((x for x in hard if x[0] <= s <= x[1]), None)
        if cb:
            out.append(dict(s=s, e=e, h=h, lv=lv, lock='', edit=False, co=cb[2], why=f'회사 규칙 · {cb[2]} (/hr-customize 9 「하네스에 반영」이 넣은 문단)'))
            continue
        if hh:
            lock, why = 'hard', f'CLAUDE.md 「{hh[2]}」 절 (customize.py lock 이 지운 줄 · 바뀐 줄을 잡음)'
        elif edit_re and h and re.search(edit_re, h):
            edit, why = True, '변형 메뉴 자리 (customization-guide 3절)'
        elif h and LOCK_HEAD.search(h):
            lock, why = 'soft', '규칙상 잠금 칸 (금지 · 범위 밖 · 멈춤 · 법전 · CP · 검증 원칙 · 이음매)'
        out.append(dict(s=s, e=e, h=h, lv=lv, lock=lock, edit=edit, why=why))
    phr = []
    for ph in rules['phrases'].get(rel, []):
        ns = [i + 1 for i, ln in enumerate(lines) if ph in ln]
        if ns:
            phr.append(dict(ph=ph, lines=ns))
    soft_lines = []
    if fm_end:
        for i in range(fm_end):
            if re.match(r'^\s*name\s*:', lines[i]):
                soft_lines.append(dict(n=i + 1, why='frontmatter name (이름을 바꾸면 CLAUDE.md · 리드 · 에이전트 연결이 끊김)'))
    el = []
    if rel in EDIT_LINES:
        pat, why = EDIT_LINES[rel]
        ns = [i + 1 for i, ln in enumerate(lines) if re.match(pat, ln)]
        if ns:
            el.append(dict(lines=ns, why=why))
    return out, phr, soft_lines, el


def sec_html(text, secs):
    """절마다 렌더 + 앞에 🔒 · ✏️ 표시 (원문 줄 범위 그대로)."""
    lines = text.splitlines()
    parts = []
    for x in secs:
        co = x.get('co')
        badge = (f'<span class="lk co" title="{E(x["why"])}">🏢 회사 규칙 · {E(co)}</span>' if co else
                 f'<span class="lk {x["lock"]}" title="{E(x["why"])}">🔒 잠금</span>' if x['lock'] else '')
        body = md('\n'.join(ln for ln in lines[x['s'] - 1:x['e']] if not (CO_START.match(ln) or CO_END.match(ln))))   # 마커 주석은 화면에 안 보임
        parts.append(f'<div class="hsec{" company" if co else " locked" if x["lock"] else " editable" if x["edit"] else ""}">{badge}{body}</div>')
    return '\n'.join(parts)


def hfile(root, p, rules, kind, **extra):
    """하네스 파일 한 장 (원문 · 렌더 · 잠금 표시 · 열기 주소). 읽지 못하면 빈 칸."""
    rel = _fwd(os.path.relpath(p, root))
    try:
        raw = H.read(p) if os.path.exists(p) else ''
    except Exception:
        raw = ''
    secs, phr, soft, el = lock_map(rel, raw, rules) if raw and not rel.endswith('.json') else ([], [], [], [])
    d = dict(kind=kind, file=rel, abs=_fwd(os.path.abspath(p)), vscode=_vscode(p), raw=raw, lines=len(raw.splitlines()),
             secs=secs, phr=phr, soft=soft, el=el, exists=os.path.exists(p),
             cos=list(dict.fromkeys(b[2] for b in co_blocks(raw.splitlines()))) if raw else [],
             locked=any(x['lock'] for x in secs) or bool(phr), editable=any(x['edit'] for x in secs) or bool(el))
    d.update(extra)
    return d


def _json(p):
    if not os.path.exists(p):
        return None, ''
    try:
        with open(p, encoding='utf-8-sig') as f:
            v = json.load(f)
        return (v if isinstance(v, dict) else {}), ('' if isinstance(v, dict) else '최상위가 객체가 아님')
    except Exception as e:
        return {}, f'JSON 을 읽지 못함: {e}'


def settings_info(root, rules):
    """훅 (settings.json · settings.local.json 의 hooks) · 권한 (permissions) · MCP (.mcp.json · settings 의 mcpServers)."""
    hooks, perms, mcp, files = [], [], [], []
    srcs = [('.claude/settings.json', os.path.join(root, '.claude', 'settings.json')),
            ('.claude/settings.local.json', os.path.join(root, '.claude', 'settings.local.json'))]
    def add_mcp(src, servers):
        if not isinstance(servers, dict):
            return
        for name, cfg in servers.items():
            cfg = cfg if isinstance(cfg, dict) else {}
            mcp.append(dict(src=src, name=str(name), type=str(cfg.get('type') or ('http' if cfg.get('url') else 'stdio')),
                            target=str(cfg.get('url') or ' '.join(str(x) for x in [cfg.get('command', '')] + list(cfg.get('args') or []) if x))))
    for rel, p in srcs:
        data, err = _json(p)
        if data is None:
            continue
        files.append(hfile(root, p, rules, 'settings', err=err))
        hk = data.get('hooks') if isinstance(data.get('hooks'), dict) else {}
        for ev, arr in hk.items():
            for m in (arr if isinstance(arr, list) else []):
                m = m if isinstance(m, dict) else {}
                for h in (m.get('hooks') if isinstance(m.get('hooks'), list) else []):
                    h = h if isinstance(h, dict) else {}
                    hooks.append(dict(src=rel, event=str(ev), matcher=str(m.get('matcher', '') or '*'), type=str(h.get('type', '')),
                                      command=str(h.get('command') or h.get('prompt') or json.dumps(h, ensure_ascii=False)),
                                      timeout=h.get('timeout')))
        pm = data.get('permissions') if isinstance(data.get('permissions'), dict) else {}
        lst = lambda k: [str(x) for x in pm.get(k, [])] if isinstance(pm.get(k), list) else []
        if pm:
            perms.append(dict(src=rel, allow=lst('allow'), deny=lst('deny'), ask=lst('ask'), mode=str(pm.get('defaultMode', '') or ''),
                              dirs=lst('additionalDirectories')))
        add_mcp(rel, data.get('mcpServers'))
    p = os.path.join(root, '.mcp.json')
    data, err = _json(p)
    if data is not None:
        files.append(hfile(root, p, rules, 'settings', err=err))
        add_mcp('.mcp.json', data.get('mcpServers'))
    return dict(hooks=hooks, perms=perms, mcp=mcp, files=files)


def harness(root):
    rules = lock_rules()
    agents, skills = [], []
    for p in sorted(glob.glob(os.path.join(root, '.claude', 'agents', '*.md'))):
        raw = H.read(p)
        fm, body = frontmatter(raw)
        a = hfile(root, p, rules, 'agent', name=fm.get('name', os.path.basename(p)[:-3]), desc=fm.get('description', ''),
                  skills=fm.get('skills', ''), tools=fm.get('tools', ''), fm={k: v for k, v in fm.items()})
        a['body'] = sec_html(raw, a['secs'])
        agents.append(a)
    for p in sorted(glob.glob(os.path.join(root, '.claude', 'skills', '*', 'SKILL.md'))):
        raw = H.read(p)
        fm, body = frontmatter(raw)
        d = os.path.dirname(p)
        scripts = []
        for x in sorted(glob.glob(os.path.join(d, 'scripts', '*.py'))):
            xr = _fwd(os.path.relpath(x, root))
            try:
                xt = H.read(x)
            except Exception:
                xt = ''
            need = rules['code'].get(xr, [])
            scripts.append(dict(name=os.path.basename(x), file=xr, vscode=_vscode(x), lines=len(xt.splitlines()),
                                locked=bool(need), lock_ok=all(n in xt for n in need)))
        s = hfile(root, p, rules, 'skill', name=fm.get('name', os.path.basename(d)), desc=fm.get('description', ''),
                  scripts=[x['name'] for x in scripts], script_info=scripts, fm={k: v for k, v in fm.items()})
        s['body'] = sec_html(raw, s['secs'])
        skills.append(s)
    cm = os.path.join(root, 'CLAUDE.md')
    claude = H.read(cm) if os.path.exists(cm) else ''
    hist = ''
    m = re.search(r'(##+\s*변경 이력.*)', claude, re.S)
    if m:
        hist = md(m.group(1))
    cfile = hfile(root, cm, rules, 'claude', name='CLAUDE.md', desc='리드 (메인 세션)가 먼저 읽는 목표 · 트리거 · 법전 · 업무 범위 · 금지 · 변경 이력')
    cfile['body'] = sec_html(claude, cfile['secs']) if claude else ''
    # 기본 대비 변경점 (hr-init 이 남긴 사본 .claude/hr-lib/_standard/)
    std = os.path.join(root, '.claude', 'hr-lib', '_standard')
    diffs = H.harness_changes(root) or []   # customize.py lock · diff 와 같은 셈 (CLAUDE.md + .claude/ · hr-lib 은 config.json 만)
    for d in diffs:   # 어느 회사 규칙 문단인지 (회사 선택에 맞게 거르는 데 씀)
        d['cos'], d['other'] = diff_cos(d.get('diff', ''))
    st = settings_info(root, rules)
    return dict(agents=agents, skills=skills, history=hist, diffs=diffs, has_std=os.path.isdir(std), claude=cfile,
                settings=st['files'], hooks=st['hooks'], perms=st['perms'], mcp=st['mcp'], lock_src=rules['src'],
                engine=rules['engine'], root_abs=_fwd(os.path.abspath(root)))


# ---------------------------------------------------------------------------
# 회사 탭 묶음
# ---------------------------------------------------------------------------
def mtime(f):
    return datetime.datetime.fromtimestamp(os.path.getmtime(f)).strftime('%m-%d %H:%M')


def stop_table(text):
    """_멈춤.md 4요소 표 → dict (무엇을 정해야 하는가 · 근거 · 선택지 · 답 형식)."""
    for head, cols, rows in H.md_tables(text):
        if '항목' in cols and '내용' in cols:
            return {r['항목']: r['내용'] for r in rows}
    return {}


def evidence(items, docs):
    """근거 문장 목록 → [dict(text, links[dict(doc, sec, label)])] · 문장 안 파일 이름 · 「## 절」 을 문서 링크로."""
    out = []
    for t in items:
        t = t.strip()
        if not t:
            continue
        links = []
        for n in sorted(docs, key=len, reverse=True):
            if n in t and not any(n in x['doc'] for x in links):
                m = re.search(re.escape(n) + r'\s*「#*\s*([^」]+)」', t)
                sec = m.group(1).strip() if m else ''
                links.append(dict(doc=n, sec=sec, label=f'{n} › {sec}' if sec else n))
                for fz in docs[n]['focus']:
                    if not m or fz['sec'] != m.group(1).strip():
                        links.append(dict(doc=n, sec=fz['sec'], label=f'{n} › {fz["label"]}'))
        out.append(dict(text=t, links=links))
    return out


NO_CP = re.compile(r'(?:CP[1-5]\s*(?:[·,/~\-]|및|와|과)\s*)*CP[1-5]\s*(?:은|는|도)?\s*(?:영향\s*없|그대로|다시\s*받지\s*않|안\s*받|변화\s*없|해당\s*없)')


def data_reflect(dt):
    """결정로그 「데이터 반영」 행 → 다시 받을 CP (확정 행만). 반환 [dict(n, date, cps)] · 같은 CP 를 그 뒤에 다시 승인한 CP 행 [dict(n, cp, date)]."""
    refl, appr = [], []
    for h, cols, rows in H.md_tables(dt or ''):
        if '체크포인트' not in cols or '번호' not in cols:
            continue
        for r in rows:
            try:
                n = int(re.sub(r'\D', '', r.get('번호', '')) or 0)
            except ValueError:
                continue
            kp, ans = r.get('체크포인트', '').strip(), r.get('답', '').strip()
            if kp.startswith('데이터 반영') and ans.startswith('확정'):
                txt = NO_CP.sub(' ', r.get('정한 것', '') + ' ' + r.get('후속', ''))   # 「CP1 · CP3 영향 없음」처럼 다시 받지 않는다고 적은 CP 는 뺌 (2026-11-15 실측)
                refl.append(dict(n=n, date=r.get('날짜', '').strip(), cps=sorted(set(re.findall(r'CP[1-5]', txt)))))
            m = re.match(r'^(CP[1-5])\b', kp)
            if m and ans.startswith('확정'):
                appr.append(dict(n=n, cp=m.group(1), date=r.get('날짜', '').strip()))
    return refl, appr


def reapproval(cp, a, refl, appr):
    """이 CP 가 데이터 반영 뒤 다시 승인받아야 하는가 → 「데이터 반영 #n」 (아니면 '').
    ① 승인 칸 답이 「추가 확인 (데이터 반영 #n)」 (CP1~CP3 를 리드가 되돌림) ② 결정로그 「데이터 반영」 확정 행이 이 CP 를 다시 받을 CP 로 적었고,
    그 행 뒤에 이 CP 확정 행이 없음 (승인 칸 승인일이 그 행 날짜보다 뒤면 다시 승인한 것으로 봄)."""
    m = re.search(r'데이터 반영\s*#?\s*(\d+)', a.get('답', '') or '')
    if m and not a.get('ok'):
        return f'데이터 반영 #{m.group(1)}'
    for r in sorted(refl, key=lambda x: -x['n']):
        if cp not in r['cps']:
            continue
        if any(x['cp'] == cp and x['n'] > r['n'] for x in appr):
            return ''
        if a.get('ok') and r['date'] and (a.get('승인일') or '') > r['date']:
            return ''
        return f'데이터 반영 #{r["n"]}'
    return ''


def decision(out, cps, docs, v_open, v_wait, skip_stop=False):
    """지금 정할 것 (4요소). 순서: 멈춤 파일 (CP · 예외 · 관문) → 관문 대기 (검증 미해결 · 미확인) → 승인 대기 CP.
    skip_stop = 멈춤 파일이 가리키는 CP 가 이미 승인됨 (지난 멈춤 · 예: CP5 확정 뒤 남은 파일) → 그 파일은 보지 않음."""
    stop = os.path.join(out, '_멈춤.md')
    if os.path.exists(stop) and not skip_stop:
        t = H.read(stop)
        m = re.search(r'^#\s*멈춤\s*·\s*(\S+)\s*(.*)$', t, re.M)
        kind, sid = (m.group(1), m.group(2).strip()) if m else ('멈춤', '')
        tb = stop_table(t)
        ev = re.split(r'<br>', tb.get('근거', ''))
        cpf = next((pat for c, nm, pat, who in CP_FILES if c == sid), None)
        cpd = next((n for n in docs if cpf and (n == cpf or ('*' in cpf and n.startswith(cpf.split('*')[0])))), None)
        if cpd and not any(cpd in x for x in ev):
            ev.append(cpd)
        return dict(kind=kind, sid=sid, src='_멈춤.md', time=mtime(stop),
                    what=tb.get('무엇을 정해야 하는가', ''), evidence=evidence(ev, docs),
                    options=[re.sub(r'^\d+\.\s*', '', o).strip() for o in tb.get('선택지', '').split('<br>') if o.strip()],
                    answer=tb.get('답 형식', H.ANSWER_FORMAT))
    if v_open or v_wait:
        return dict(kind='관문', sid='검증', src='검증로그.md', time=mtime(os.path.join(out, '검증로그.md')),
                    what=f'검증로그에 닫히지 않은 행이 있음 · 미해결 {v_open} · 미확인 {v_wait} (미확인 = 범위 밖 거절 행동 테스트를 아직 묻지 않음)',
                    evidence=evidence(['검증로그.md 「검증 로그」 · 「미해결」 · 「미확인」 행'], docs),
                    options=['넘긴 담당이 고친 뒤 다시 검증', H.GATE_ACCEPT, '추가 확인 · 보류'], answer=H.ANSWER_FORMAT)
    p = next((c for c in cps if c['state'] not in ('승인', '자료 없음', '앞 CP 먼저')), None)
    if p and p.get('re'):
        return dict(kind='CP', sid=p['cp'], src=p['file'], time='', re=p['re'],
                    what=f'{p["cp"]} · {p["name"]} 재승인 필요 ({p["re"]} · 결정로그) · 바뀐 자료가 닿는 CP 라 그 단계를 다시 돌린 뒤 {p["file"]} 「## 승인」 표에 다시 답을 받음 · 승인자 {p["who"]}',
                    evidence=evidence([p['file'], '결정로그.md 「데이터 반영」 행'], docs),
                    options=['영향받는 가장 앞 단계부터 다시 돌리고 이 CP 에서 다시 확정', '고칠 곳을 말하고 고친 뒤 확정 (바꾼 내용은 이유 칸에)', '추가 확인 · 보류 · 이관'],
                    answer=H.ANSWER_FORMAT)
    if p:
        return dict(kind='CP', sid=p['cp'], src=p['file'], time='',
                    what=f'{p["cp"]} · {p["name"]} · {p["file"]} 의 「승인」 표가 비어 있음 (지금 답: {p["state"]}) · 승인자 {p["who"]}',
                    evidence=evidence([p['file']], docs),
                    options=['자료 그대로 확정', '고칠 곳을 말하고 고친 뒤 확정 (바꾼 내용은 이유 칸에)', '추가 확인 · 보류 · 이관'],
                    answer=H.ANSWER_FORMAT)
    return None


def blank(text):
    """양식 그대로인가. 사람 칸 · 스크립트 칸이 있는 양식 모양 파일만 본다 (칸이 없는 문서 · 예: ①_사업전략분석 은 내용이 있으면 결과)."""
    slots = re.search(r'<!-- auto:start ', text) or any(('사람 답' in cols or '답' in cols or ('번호' in cols and '날짜' in cols))
                                                         for h, cols, rows in H.md_tables(text))
    return bool(slots) and H.is_blank_form(text)


def rpu_names(out):
    """RPU 코드 → 이름 · 본부 (코드만 두면 사람이 못 읽음): ① RPU 보조표 + ② 사업 이벤트 (새 RPU)."""
    names, orgs = {}, {}
    try:
        for fn, code_col, name_col in (('①_RPU보조표.md', 'RPU 코드', 'RPU'), ('②_수요파라미터.md', '코드', '이름')):
            p = os.path.join(out, fn)
            if not os.path.exists(p):
                continue
            for h, cols, rows in H.md_tables(H.read(p)):
                if code_col in cols and name_col in cols:
                    for r in rows:
                        c, n = (r.get(code_col) or '').strip(), (r.get(name_col) or '').strip()
                        if c and n and c not in names:
                            names[c] = n
                            if r.get('본부'):
                                orgs[c] = r['본부'].strip()
    except Exception:
        pass
    return names, orgs


def supply_terms(out):
    """공급 가정 S-1~S-5 → 가정 이름 (③_공급가정.md 「가정 표」 · 없으면 양식)."""
    texts = []
    p = os.path.join(out, '③_공급가정.md')
    if os.path.exists(p):
        texts.append(H.read(p))
    try:
        texts.append(H.form('③_공급가정.md'))
    except Exception:
        pass
    for t in texts:
        try:
            for h, cols, rows in H.md_tables(t):
                if '번호' in cols and '가정' in cols:
                    m = {(r.get('번호') or '').strip(): (r.get('가정') or '').strip() for r in rows}
                    m = {k: v for k, v in m.items() if re.fullmatch(r'S-\d', k) and v and '{' not in v}
                    if m:
                        return m
        except Exception:
            continue
    return {}


def harness_reflect(text):
    """outputs/<회사>/_하네스반영.md (adapt.py harness 요약) → 표 행 · 층별 개수 · 잠금 · 회귀 결과. 표가 없으면 None."""
    rows, hdr = [], None
    for ln in text.splitlines():
        s = ln.strip()
        if not s.startswith('|'):
            if hdr and rows:
                break
            continue
        cells = [plain(c.strip()) for c in s.strip('|').split('|')]
        if hdr is None:
            if cells and cells[0] == '층':
                hdr = cells
            continue
        if all(re.fullmatch(r':?-+:?', c) for c in cells if c):
            continue
        cells += [''] * (4 - len(cells))
        rows.append(dict(layer=cells[0], file=cells[1], brief=cells[2], why=cells[3]))
    if not rows:
        return None
    def grp(l):
        return '담당' if l.startswith('담당') else '스킬' if l.startswith('스킬') else l
    counts = {}
    for r in rows:
        r['grp'] = grp(r['layer'])
        counts[r['grp']] = counts.get(r['grp'], 0) + 1
    order = ['CLAUDE.md', '담당', '스킬']
    cnt = [dict(k=k, n=counts[k]) for k in order if k in counts] + [dict(k=k, n=v) for k, v in counts.items() if k not in order]
    m = re.search(r'잠금 검사[^*\n]*\*\*([^*]+)\*\*', text)
    g = re.search(r'(?:회귀|다시 계산)[^*\n]*\*\*([^*]+)\*\*', text)
    st = re.search(r'^>\s*\*\*([^*]+)\*\*', text, re.M)
    return dict(rows=rows, counts=cnt, lock=m.group(1).strip() if m else '', regress=g.group(1).strip() if g else '',
                status=st.group(1).strip() if st else '')


def company(root, co):
    out = os.path.join(root, 'outputs', co)
    names = sorted(os.path.basename(p) for p in glob.glob(os.path.join(out, '*.md')))
    DOCS[:] = names
    rn, ro = rpu_names(out)
    sup = supply_terms(out)
    set_gloss(rn, ro, sup)
    docs = {}
    for n in names:
        t = H.read(os.path.join(out, n))
        docs[n] = dict(html=md(t), mtime=mtime(os.path.join(out, n)), blank=blank(t) and n not in ('결정로그.md', '검증로그.md', '_멈춤.md'),
                       focus=focus(t), heads=[h for h, b in sections(t) if h])
    cps = []
    stop_cp = ''   # 멈춤 파일이 가리키는 CP (그 CP 자료가 아직 양식 모양이어도 「자료 없음」이 아니라 「승인 대기」)
    sp_ = os.path.join(out, '_멈춤.md')
    if os.path.exists(sp_):
        m_ = re.search(r'^#\s*멈춤\s*·\s*CP\s+(CP[1-5])\b', H.read(sp_), re.M)
        stop_cp = m_.group(1) if m_ else ''
    before_ok = True   # 앞 CP 가 모두 승인됐는가 (아니면 뒤 CP 를 「승인 대기」로 띄우지 않음)
    dt = H.read(find(out, '결정로그.md')) if find(out, '결정로그.md') else ''
    refl, appr = data_reflect(dt)
    for cp, nm, pat, who in CP_FILES:
        f = find(out, pat)
        if f and blank(H.read(f)) and cp != stop_cp:
            f = None   # 양식 그대로 (사람 칸 · 스크립트 칸이 모두 빔) = 자료 없음 · 멈춤 파일이 이 CP 를 가리키면 그대로 둠
        a = H.approval(f, cp) if f else dict(ok=False, 답='', 승인자='')
        state = '승인' if a['ok'] else ('승인 대기' if cp == stop_cp else '자료 없음' if not f else ('승인 대기' if not a.get('답') else a['답']))
        if state == '승인 대기' and not before_ok and cp != stop_cp:
            state = '앞 CP 먼저'
        re_ = reapproval(cp, a, refl, appr) if f else ''
        if re_:   # 데이터 반영 뒤 다시 받을 CP (승인 대기 · 추가 확인 · 지난 승인 모두 「재승인 필요」 · 승인 대기 셈은 그대로)
            state = '재승인 필요'
        before_ok = before_ok and state == '승인'
        fn = os.path.basename(f) if f else None
        cps.append(dict(cp=cp, name=nm, file=fn or pat, state=state, who=a.get('승인자') or who, date=a.get('승인일', ''), re=re_,
                        doc=fn if fn in docs else None, focus=docs[fn]['focus'] if fn in docs else []))
    stages, staged = [], set()
    for st, nm, files in STAGES:
        items = []
        for pat in files:
            f = find(out, pat)
            if not f or (f.endswith('.md') and docs.get(os.path.basename(f), {}).get('blank')):
                items.append(dict(name=pat, html='', exists=False))
                continue
            n = os.path.basename(f)
            staged.add(n)
            items.append(dict(name=n, html=xlsx_html(f) if f.endswith('.xlsx') else '', doc=n if n in docs else None, exists=True, mtime=mtime(f), desc=file_desc(n)))
        stages.append(dict(st=st, name=nm, items=items, done=sum(1 for x in items if x['exists'])))
    extra = [n for n in names if n not in staged]   # 결정로그 · 검증로그 · 멈춤 · 그 밖에 담당이 남긴 .md
    vt = H.read(find(out, '검증로그.md')) if find(out, '검증로그.md') else ''
    exc = [ln for ln in dt.splitlines() if ln.startswith('|') and '예외' in ln]
    data = dash(out)
    rep = find(out, '리포트_*.md')
    pending = next((c for c in cps if c['state'] not in ('승인', '자료 없음', '앞 CP 먼저')), None)
    v_open, v_wait = len(re.findall(r'\| 미해결 \|', vt)), len(re.findall(r'\| 미확인 \|', vt))
    v_acc = len(re.findall(r'\| 인정 \(결정로그 #\d+\) \|', vt))   # 관문 인정 (사람 답 · 결정로그) · 미해결로 세지 않음 (관문 열림)
    vi = H.virtual_info(os.path.join(root, 'data', co))
    checks = [dict(doc=n, sec=fz['sec'], label=fz['label']) for n in names if not docs[n]['blank'] for fz in docs[n]['focus'] if fz['label'] != '승인 칸']
    stale = bool(stop_cp) and any(c['cp'] == stop_cp and c['state'] == '승인' for c in cps)   # 지난 CP 멈춤 (그 CP 가 이미 승인됨)
    dec = decision(out, cps, docs, v_open, v_wait, skip_stop=stale)
    if dec:   # 화면에 보일 글 (코드에 풀이 말풍선) · 원문 글은 그대로 둠
        dec['what_h'] = gloss(E(dec['what']))
        dec['options_h'] = [gloss(E(o)) for o in dec['options']]
        for e in dec['evidence']:
            e['text_h'] = gloss(E(e['text'].replace('<br>', ' ')))
    hr_ = find(out, '_하네스반영.md')
    harn = harness_reflect(H.read(hr_)) if hr_ else None
    try:   # 내려받기 (2026-11-16 · 이 시점 사본을 HTML 안에 넣음 · 뷰어를 다시 만들 때마다 새로)
        import report_xlsx
        files = report_xlsx.downloads(out, co)
    except Exception as e:   # noqa: BLE001
        print(f'[알림] 내려받기 묶음을 만들지 못함 ({type(e).__name__}: {e}) · 화면은 그대로', file=sys.stderr)
        files = {}
    res = dict(co=co, files=files, descs={n: file_desc(n) for n in extra}, virtual=dict(on=vi['virtual'], marker=vi['marker'], stamped=len(vi['stamped'])) if vi['virtual'] else None,
               cps=cps, pending=pending, stop='_멈춤.md' if '_멈춤.md' in docs and not stale else '', stale_stop=stale,
               reapprove=[c['cp'] for c in cps if c.get('re')], stages=stages, extra=extra, docs=docs, dash=data,
               report=os.path.basename(rep) if rep else '', v_open=v_open, v_wait=v_wait, v_acc=v_acc, checks=checks,
               decision=dec, gl=dict(rpu=rn, org=ro, sup=sup), harn=harn,
               v_all=len(re.findall(r'^\| \d+ \|', vt, re.M)), exceptions=len(exc))
    DOCS[:] = []
    set_gloss()
    return res


def dash(out):
    d = {}
    ty = H.config()['target_year']
    d['names'], d['orgs'] = rpu_names(out)   # 차트 · 표에 코드만 두면 사람이 못 읽음
    try:
        if os.path.exists(os.path.join(out, '②_대사.xlsx')):
            _, R = H.read_xlsx_table(os.path.join(out, '②_대사.xlsx'))
            d['recon'] = [dict(c=r['RPU 코드'], b=r['B'] if isinstance(r['B'], (int, float)) else None, cc=r['C'] if isinstance(r['C'], (int, float)) else None,
                               dec=r['결정 수요 사람'], dig=r['결정 수요 디지털'], cause=r['원인 판정'] or '') for r in R if r['RPU 코드'] != '합계']
            t = next((r for r in R if r['RPU 코드'] == '합계'), None)
            if t:
                d['dec_h'], d['dec_d'] = t['결정 수요 사람'], t['결정 수요 디지털']
        if os.path.exists(os.path.join(out, '③_공급.xlsx')):
            _, S = H.read_xlsx_table(os.path.join(out, '③_공급.xlsx'), '기본')
            _, U = H.read_xlsx_table(os.path.join(out, '③_공급.xlsx'), '상승')
            t = next(r for r in S if r['RPU 코드'] == '합계')
            d['hc'], d['sup'] = t['현원'], t[f'공급 {ty}']
            d['sup_up'] = next(r for r in U if r['RPU 코드'] == '합계')[f'공급 {ty}']
            d['supply'] = {r['RPU 코드']: r[f'공급 {ty}'] for r in S if r['RPU 코드'] != '합계'}
        g = os.path.join(out, '④_Gap대안.md')
        if os.path.exists(g):
            rows = next((rows for h, c, rows in H.md_tables(H.read(g)) if h.startswith('1. Gap')), [])
            d['gap'] = [dict(c=r['RPU 코드'], g=float(r['Gap']), t=r['유형'], hc=r['현원']) for r in rows if r['RPU 코드'] != '합계' and r['Gap']]
            tot = next((r for r in rows if r['RPU 코드'] == '합계'), None)
            if tot:
                d['gap_total'] = float(tot['Gap'])
    except Exception as e:  # 뷰어는 결과가 덜 된 상태에서도 열려야 한다
        d['error'] = str(e)
    d['ty'] = ty
    return d


# ---------------------------------------------------------------------------
CSS = r"""
:root{--bg:#f7f7f8;--surface:#ffffff;--surface-2:#f9fafb;--hover:#fafafa;--ink:#2d2926;--ink-2:#4b4642;--muted:#6b6560;--line:#e5e7eb;--line-2:#eef0f2;
--accent:#d04a02;--accent-ink:#b03f02;--accent-soft:#fdf0e8;--a2:#eb8c00;--a3:#ffb600;
--ok:#1f7a45;--ok-bg:#e8f5ee;--warn:#a85a00;--warn-bg:#fff4e0;--bad:#b42318;--bad-bg:#fdecea;--gray:#55504b;--gray-bg:#f1f1f2;
--short:#c62828;--surplus:#2b6cb0;--zebra:#f6f6f8;--s1:#d04a02;--s2:#e0a000;--shadow:0 1px 2px rgba(16,24,40,.04),0 1px 3px rgba(16,24,40,.05);--side:236px;color-scheme:light}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#121212;--surface:#1c1b1a;--surface-2:#222120;--hover:#252321;--ink:#f2f0ee;--ink-2:#cfcac5;--muted:#9c958f;--line:#34322f;--line-2:#2a2826;
--accent:#ff7a33;--accent-ink:#ff8f52;--accent-soft:#3a2418;--a2:#f5a030;--a3:#ffc633;--ok:#5cc28a;--ok-bg:#16301f;--warn:#f0ac45;--warn-bg:#3a2a12;--bad:#ff7b6e;--bad-bg:#3a1a17;--gray:#bdb7b1;--gray-bg:#2a2826;
--short:#ff6b6b;--surplus:#6aaaf2;--zebra:#201f1e;--s1:#ff7a33;--s2:#ffc633;--shadow:0 1px 2px rgba(0,0,0,.3);color-scheme:dark}}
:root[data-theme="dark"]{--bg:#121212;--surface:#1c1b1a;--surface-2:#222120;--hover:#252321;--ink:#f2f0ee;--ink-2:#cfcac5;--muted:#9c958f;--line:#34322f;--line-2:#2a2826;
--accent:#ff7a33;--accent-ink:#ff8f52;--accent-soft:#3a2418;--a2:#f5a030;--a3:#ffc633;--ok:#5cc28a;--ok-bg:#16301f;--warn:#f0ac45;--warn-bg:#3a2a12;--bad:#ff7b6e;--bad-bg:#3a1a17;--gray:#bdb7b1;--gray-bg:#2a2826;
--short:#ff6b6b;--surplus:#6aaaf2;--zebra:#201f1e;--s1:#ff7a33;--s2:#ffc633;--shadow:0 1px 2px rgba(0,0,0,.3);color-scheme:dark}
*{box-sizing:border-box}html,body{margin:0}body{background:var(--bg);color:var(--ink);font:14px/1.6 "Pretendard",system-ui,-apple-system,"Segoe UI","Malgun Gothic","Apple SD Gothic Neo",sans-serif;-webkit-font-smoothing:antialiased}
svg.i{width:18px;height:18px;flex:none;fill:none;stroke:currentColor;stroke-width:1.7;stroke-linecap:round;stroke-linejoin:round}
.app{display:flex;min-height:100vh}
.side{position:fixed;inset:0 auto 0 0;width:var(--side);background:var(--surface);border-right:1px solid var(--line);display:flex;flex-direction:column;z-index:20}
.brand{display:flex;gap:10px;align-items:center;padding:18px 18px 14px;border-bottom:1px solid var(--line-2)}.brand .mark{width:28px;height:28px;border-radius:7px;background:var(--accent);flex:none}
.brand b{display:block;font-size:14.5px;font-weight:800;letter-spacing:-.2px}.brand small{color:var(--muted);font-size:11.5px}
#nav{flex:1;overflow:auto;padding:10px 10px}#nav h6{margin:14px 10px 4px;font-size:11px;font-weight:700;color:var(--muted);letter-spacing:.3px}
#nav button{display:flex;align-items:center;gap:10px;width:100%;border:0;background:none;color:var(--ink-2);font:inherit;font-size:13.5px;padding:7px 10px;border-radius:8px;cursor:pointer;text-align:left}
#nav button:hover{background:var(--surface-2);color:var(--ink)}#nav button.on{background:var(--accent-soft);color:var(--accent-ink);font-weight:700}
#nav .badge{margin-left:auto;font-size:11px;font-weight:800;border-radius:999px;padding:0 7px;line-height:18px}
.sidefoot{padding:12px 18px;border-top:1px solid var(--line-2);color:var(--muted);font-size:11.5px;word-break:break-all}
.col{flex:1;min-width:0;margin-left:var(--side)}
.top{position:sticky;top:0;z-index:10;height:56px;background:var(--surface);border-bottom:1px solid var(--line);display:flex;align-items:center;gap:12px;padding:0 24px}
.top h1{margin:0;font-size:16px;font-weight:800;letter-spacing:-.2px;white-space:nowrap}#ttl{min-width:0;overflow:hidden;text-overflow:ellipsis}#stat{min-width:0;display:flex}#stat .pill{overflow:hidden;text-overflow:ellipsis}.sp{flex:1}.built{color:var(--muted);font-size:12px;white-space:nowrap}
.seg{display:flex;gap:2px;background:var(--surface-2);border:1px solid var(--line);border-radius:999px;padding:2px}
.seg button{border:0;background:none;color:var(--ink-2);border-radius:999px;padding:4px 12px;font:inherit;font-size:12.5px;cursor:pointer}.seg button.on{background:var(--surface);color:var(--ink);font-weight:700;box-shadow:var(--shadow)}
.icon-btn{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border:1px solid var(--line);border-radius:8px;background:var(--surface);color:var(--ink-2);cursor:pointer}.icon-btn:hover{background:var(--surface-2)}
#menu{display:none}.scrim{display:none}
main{max-width:1240px;margin:0 auto;padding:24px 24px 64px}section.pane{display:none}section.pane.on{display:block}
.pagehd{display:flex;align-items:flex-end;gap:12px;margin:0 0 16px;flex-wrap:wrap}.pagehd h2{margin:0;font-size:22px;font-weight:800;letter-spacing:-.4px}.pagehd p{margin:0;color:var(--muted)}
.card{background:var(--surface);border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow);padding:20px 24px;margin:0 0 16px}
.card>h3{margin:0 0 12px;font-size:15px;font-weight:800}.card .sub{color:var(--muted);font-size:12.5px;font-weight:400;margin-left:6px}
.grid{display:grid;gap:16px;grid-template-columns:repeat(auto-fill,minmax(280px,1fr))}.two{display:grid;gap:16px;grid-template-columns:1fr 1fr}
.kpi{display:grid;gap:16px;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));margin:0 0 16px}
.kpi>div{background:var(--surface);border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow);padding:16px 18px}
.kpi span{display:block;color:var(--muted);font-size:12px}.kpi b{display:block;font-size:26px;font-weight:800;letter-spacing:-.6px;font-variant-numeric:tabular-nums;margin:2px 0}
.kpi span.pill{display:inline-flex;color:var(--gray)}.kpi span.pill.ok{color:var(--ok)}.kpi span.pill.warn{color:var(--warn)}.kpi span.pill.bad{color:var(--bad)}.kpi b.none{font-size:15px;font-weight:600;color:var(--muted);padding:7px 0 6px}.kpi .hl{border-top:3px solid var(--accent)}
.pill{display:inline-flex;align-items:center;gap:4px;border-radius:999px;padding:2px 10px;font-size:12px;font-weight:700;background:var(--gray-bg);color:var(--gray);white-space:nowrap}
.pill.ok,.badge.ok{background:var(--ok-bg);color:var(--ok)}.pill.warn,.badge.warn{background:var(--warn-bg);color:var(--warn)}.pill.bad,.badge.bad{background:var(--bad-bg);color:var(--bad)}.pill.gray{background:var(--gray-bg);color:var(--gray)}.pill.acc{background:var(--accent-soft);color:var(--accent-ink)}
.st-승인{color:var(--ok)}.st-대기{color:var(--warn)}.st-자료{color:var(--muted)}.st-기타{color:var(--bad)}
.tw{overflow:auto;max-height:72vh;margin:8px 0 14px;border:1px solid var(--line);border-radius:10px}
table{word-break:keep-all;border-collapse:separate;border-spacing:0;width:100%;font-size:12.5px}th,td{border-bottom:1px solid var(--line-2);padding:7px 10px;text-align:left;vertical-align:top}
th{position:sticky;top:0;z-index:1;background:var(--surface-2);color:var(--ink-2);font-weight:700;white-space:nowrap;border-bottom:1px solid var(--line)}
tbody tr:hover td{background:var(--hover)}tbody tr:last-child td{border-bottom:0}td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}tr.tot td{font-weight:800;border-top:2px solid var(--ink-2)}
h2{font-size:17px;margin:24px 0 10px;font-weight:800;letter-spacing:-.2px}h3{font-size:15px}h4{font-size:14px;margin:14px 0 6px}
code{background:var(--surface-2);border:1px solid var(--line);border-radius:4px;padding:0 4px;font-size:12px}hr{border:0;border-top:1px solid var(--line);margin:16px 0}
pre{background:var(--surface-2);border:1px solid var(--line);border-radius:8px;padding:12px;overflow:auto;font-size:12px}blockquote{border-left:3px solid var(--accent);margin:8px 0;padding:4px 12px;color:var(--ink-2);background:var(--surface-2);border-radius:0 6px 6px 0}
.rsum{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:16px 18px;margin:6px 0 14px}.rsum .rt{font-size:17px;font-weight:800}.rsum .rsub{color:var(--muted);font-size:12.5px;margin:2px 0 10px;border-bottom:2px solid var(--accent);padding-bottom:8px}.rsum .rstat{background:var(--accent-soft,#FBE3D6);color:var(--accent-ink);font-weight:700;font-size:13px;border-radius:8px;padding:6px 10px;margin:0 0 12px}.rsum h4.rh{margin:12px 0 4px;font-size:14px}.rsum ul.rl{margin:0 0 4px;padding-left:0;list-style:none;font-size:13.5px;line-height:1.65}.rsum ul.rl li{margin:2px 0}tr.grp td{color:var(--accent-ink);font-weight:800;background:var(--surface-2)}tr.sub td{font-weight:700;background:var(--surface-2)}td.hl-r{background:rgba(220,60,30,.16)}td.hl-y{background:rgba(235,170,0,.2)}td.hl-g{background:rgba(40,150,60,.15)}td.hl-n{background:rgba(120,120,120,.12)}.otab{padding:0;overflow:visible}.otab .obar{display:flex;align-items:center;gap:14px;padding:12px 16px;border-bottom:1px solid var(--line);flex-wrap:wrap}.otab .obar label{display:inline-flex;align-items:center;gap:6px;font-size:13px;color:var(--ink-2);cursor:pointer}.otab .obar .sp{flex:1}.otab .tw{margin:0;border:0;border-radius:0;max-height:none;overflow:visible}.otab .otable>thead th{position:static}.otable tr.pv .xl .tw{overflow-x:auto}.otable td{vertical-align:middle}.otable{table-layout:fixed}.otable td.oc,.otable th.oc{width:44px;text-align:center}.otable th:nth-child(2){width:64px}.otable th:nth-child(4){width:110px}.otable th:nth-child(5){width:104px}.otable th:nth-child(6){width:64px}.otable tr.pv td>.in{max-width:100%;min-width:0}.otable>tbody>tr:not(.pv)>td{overflow-wrap:anywhere}.rsum .kpi b{white-space:nowrap;font-size:24px}.otable tr.off td{color:var(--muted)}.otable tr.open td{background:var(--accent-soft)}.otable .stb{display:inline-block;min-width:26px;text-align:center;font-weight:800;color:var(--accent-ink);background:var(--accent-soft);border-radius:6px;padding:1px 6px;font-size:12px}.otable button.fname{all:unset;cursor:pointer;font-weight:700;color:var(--ink)}.otable button.fname:hover{color:var(--accent-ink);text-decoration:underline}.otable button.fname[disabled]{cursor:default;color:var(--muted);text-decoration:none}.otable .fdesc{color:var(--muted);font-size:12px;margin-top:1px}.otable td.fk,.otable td.fm{white-space:nowrap;color:var(--ink-2)}.otable td.fv{white-space:nowrap;text-align:right}.otable tr.pv td{background:var(--surface-2);padding:14px 18px}.otable tr.pv:hover td{background:var(--surface-2)}.onote{margin:0;padding:10px 16px;font-size:12px;border-top:1px solid var(--line)}.otable input[type=checkbox],.otab input[type=checkbox]{width:15px;height:15px;accent-color:var(--accent)}.pagehd.wact{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;flex-wrap:wrap}.pact,.dact{display:inline-flex;gap:6px;flex-wrap:wrap}.dact{float:right}.btn.quiet svg.i{width:14px;height:14px}.itip{display:inline-flex;align-items:center;justify-content:center;width:17px;height:17px;border-radius:50%;border:1px solid var(--line);color:var(--muted);font-size:11px;font-weight:700;cursor:help;margin:0 4px;position:relative;vertical-align:middle;font-style:normal}.decide .dh{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.decide .dh .sp{flex:1}.decide .sm{font-size:12px}.decide .dwhat{font-size:15.5px;font-weight:700;margin:10px 0 12px;line-height:1.5}.dtab{width:100%;border-collapse:collapse;font-size:13px}.dtab th{width:88px;text-align:left;vertical-align:top;color:var(--muted);font-weight:700;padding:8px 10px 8px 0;background:none;position:static;border-bottom:1px solid var(--line-2)}.dtab td{padding:8px 0;border-bottom:1px solid var(--line-2)}.dtab tr:last-child th,.dtab tr:last-child td{border-bottom:0}.dtab ul,.dtab ol{margin:0;padding-left:18px}.dtab li{margin:2px 0}.achips{display:flex;gap:6px;flex-wrap:wrap;align-items:center}.achip{border:1px solid var(--accent);color:var(--accent-ink);border-radius:999px;padding:2px 10px;font-weight:700;font-size:12.5px;cursor:help}details.drawer{margin:8px 0 0}details.drawer>summary{cursor:pointer;color:var(--muted);font-size:12.5px;font-weight:600}details.drawer>ul{margin:6px 0 0;padding-left:18px;font-size:12.5px;color:var(--ink-2);line-height:1.6}.mtoc.tabbed{position:sticky;top:var(--stick,56px);z-index:4;background:var(--surface);padding:10px 0;margin:0 0 12px;border-bottom:1px solid var(--line)}.mtoc.tabbed a.on{background:var(--accent);border-color:var(--accent);color:#fff}.xl>.mtoc.tabbed{background:var(--surface-2)}td.long{min-width:18em}details.att{margin:8px 0 14px}details.att>summary{color:var(--muted);font-size:12.5px;cursor:pointer}.fopen{font-size:12px;color:var(--muted);margin:4px 0 8px}.fopen code{font-size:11.5px}.fopen b{color:var(--accent-ink)}
.warn{color:var(--warn);font-weight:700}.ok{color:var(--ok);font-weight:700}.bad{color:var(--bad);font-weight:700}.muted{color:var(--muted)}ul.notes{color:var(--muted);font-size:12px;margin:2px 0 6px}
details{border:1px solid var(--line);border-radius:12px;background:var(--surface);box-shadow:var(--shadow);margin:0 0 10px}details>summary{cursor:pointer;padding:12px 16px;font-weight:700;list-style:none}
details>summary::-webkit-details-marker{display:none}details>summary:before{content:"";display:inline-block;width:6px;height:6px;border-right:2px solid var(--muted);border-bottom:2px solid var(--muted);transform:rotate(-45deg);margin:0 10px 2px 0;transition:transform .15s}
details[open]>summary:before{transform:rotate(45deg)}details>.in{padding:0 18px 14px}
.vban{padding:10px 24px;background:var(--warn-bg);color:var(--ink);border-bottom:1px solid var(--line);font-size:13.5px}.tag{font-size:11px;color:var(--muted);border:1px solid var(--line);border-radius:6px;padding:1px 6px;margin-right:4px}
.flow{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin:6px 0 4px}.flow .node{border:1px solid var(--line);border-radius:10px;padding:8px 12px;background:var(--surface-2);font-size:12.5px;min-width:110px}
.flow .node b{display:block;font-size:13px}.flow .cp{font-size:11px;font-weight:800;border-radius:999px;padding:2px 8px;background:var(--accent-soft);color:var(--accent-ink)}
.chart{width:100%;overflow-x:auto}.legend{display:flex;flex-wrap:wrap;gap:6px 16px;font-size:12px;color:var(--ink-2);margin:4px 0 8px}.legend i{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:5px;vertical-align:-1px}
#tip{position:fixed;pointer-events:none;background:var(--surface);color:var(--ink);border:1px solid var(--line);border-radius:8px;padding:6px 9px;font-size:12px;box-shadow:0 4px 14px rgba(0,0,0,.15);display:none;z-index:40;max-width:280px}
.diff pre{max-height:360px}.add{color:var(--ok)}.del{color:var(--bad)}
.decide{background:var(--surface);border:1px solid var(--line);border-top:4px solid var(--accent);border-radius:12px;box-shadow:0 4px 16px rgba(208,74,2,.08);padding:20px 24px;margin:0 0 16px}
.decide .dh{display:flex;flex-wrap:wrap;align-items:center;gap:8px 10px;margin-bottom:14px}.decide .dh b{font-size:18px;font-weight:800;letter-spacing:-.3px}.decide .dh .muted{font-size:12.5px}
.dg{display:grid;gap:16px 24px;grid-template-columns:minmax(200px,1.2fr) minmax(240px,1.6fr) minmax(200px,1.2fr) minmax(160px,.9fr)}
.dg>div{border-left:2px solid var(--line);padding-left:14px}.dg>div>span{display:block;color:var(--muted);font-size:11.5px;font-weight:800;letter-spacing:.3px;margin-bottom:4px}
.dg p{margin:0 0 4px}.dg ul,.dg ol{margin:0;padding-left:18px}.dg li{margin:0 0 6px}.dg .ans{font-size:15px}
#dc{position:sticky;top:56px;z-index:9}.strip{display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:9px 24px;background:var(--accent-soft);border-bottom:1px solid var(--line);font-size:13px}.strip .what{color:var(--ink-2);min-width:0;flex:1 1 280px}
.btn{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--accent);background:var(--accent);color:#fff;border-radius:8px;padding:5px 12px;font:inherit;font-size:12.5px;font-weight:700;cursor:pointer;text-decoration:none}
.btn.ghost{background:var(--surface);color:var(--accent-ink)}
a.dl{color:var(--accent-ink);text-decoration:none;border-bottom:1px dotted currentColor;cursor:pointer}a.dl:hover{border-bottom-style:solid}.jump{font-size:12px;margin-right:6px;white-space:nowrap}.kpi .jump{white-space:normal;display:inline-block}
.steps{display:flex;align-items:flex-start;gap:0;overflow-x:auto;padding:4px 0 2px}.steps .st{flex:1 1 0;min-width:92px;display:flex;flex-direction:column;align-items:center;text-align:center;position:relative;padding:0 4px}
.steps .st:before{content:"";position:absolute;top:15px;left:-50%;width:100%;height:2px;background:var(--line);z-index:0}.steps .st:first-child:before{display:none}.steps .st.ok:before{background:var(--ok)}
.steps .dw{position:relative;z-index:1;height:32px;display:flex;align-items:center;justify-content:center}
.steps .sg .dot{display:block;width:12px;height:12px;border-radius:50%;background:var(--gray-bg);border:2px solid var(--line);box-shadow:0 0 0 3px var(--surface)}
.steps .sg.ok .dot{background:var(--ok);border-color:var(--ok)}.steps .sg.warn .dot{background:var(--warn-bg);border-color:var(--warn)}
.steps .cp .dot{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:12px;background:var(--surface);color:var(--gray);border:2px solid var(--line);box-shadow:0 0 0 3px var(--surface)}
.steps .cp.ok .dot{background:var(--ok);border-color:var(--ok);color:var(--surface);font-size:16px}.steps .cp.warn .dot{background:var(--warn-bg);border-color:var(--warn);color:var(--warn)}.steps .cp.bad .dot{background:var(--bad-bg);border-color:var(--bad);color:var(--bad)}
.steps .lbl{font-size:12px;font-weight:600;color:var(--ink-2);margin-top:6px;line-height:1.3;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;word-break:keep-all}.steps .cp .lbl{font-weight:800;color:var(--ink)}
.steps .meta{font-size:11px;color:var(--muted);line-height:1.35;word-break:keep-all}.steps .meta.re{color:var(--warn);font-weight:700}.steps .meta.who{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.empty{text-align:center;padding:48px 16px;color:var(--muted)}.empty svg.i{width:36px;height:36px;color:var(--line);stroke-width:1.4}.empty p{margin:8px 0 0}.empty .say{color:var(--ink-2);font-size:13px}.empty .say b{color:var(--accent-ink)}
.docwrap{display:grid;grid-template-columns:240px 1fr;gap:16px;align-items:start}.doclist{position:sticky;top:112px;max-height:calc(100vh - 136px);overflow:auto;font-size:12.5px;background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:8px}
.doclist h4{margin:8px 8px 4px;color:var(--muted);font-size:11px;letter-spacing:.3px}.doclist a.dl{display:block;border:0;padding:6px 8px;border-radius:8px;color:var(--ink-2);word-break:break-all}
.doclist a.dl:hover{background:var(--surface-2)}.doclist a.dl.on{background:var(--accent-soft);color:var(--accent-ink);font-weight:700}.dochd{border-bottom:1px solid var(--line);padding-bottom:10px;margin-bottom:8px}
.dochd .toc{margin:8px 0 0;border:0;background:none;box-shadow:none}.dochd .toc>summary{padding:2px 0;font-weight:400;font-size:12.5px;color:var(--muted)}.dochd .toc .in{padding:4px 0;display:flex;flex-wrap:wrap;gap:4px 12px;font-size:12.5px}
.mdbody [data-h]{scroll-margin-top:calc(var(--stick,56px) + 16px);border-radius:6px;transition:background .6s}.mdbody [data-h].hit{background:var(--accent-soft);box-shadow:0 0 0 6px var(--accent-soft)}
.totop{float:right;color:var(--accent-ink);text-decoration:none;font-weight:700}.closed{position:sticky;top:0;z-index:50;background:var(--warn);color:#fff;font-weight:800;padding:10px 24px}
.hnav{display:flex;flex-wrap:wrap;gap:6px;margin:0 0 14px}.hnav button{border:1px solid var(--line);background:var(--surface);color:var(--ink-2);border-radius:999px;padding:4px 12px;font:inherit;font-size:12.5px;cursor:pointer}.hnav button:hover{border-color:var(--accent);color:var(--accent-ink)}
.hsech{scroll-margin-top:72px;display:flex;align-items:baseline;gap:10px;flex-wrap:wrap}.hsech .muted{font-size:12.5px;font-weight:400}
.legend2{display:flex;flex-wrap:wrap;gap:6px 14px;font-size:12px;color:var(--ink-2);margin:0 0 12px}
.hgrid{display:grid;gap:12px;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));margin:0 0 12px}
.hc{background:var(--surface);border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow);padding:14px 16px;display:flex;flex-direction:column;gap:6px;min-width:0}
.hc .t{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.hc .t b{font-size:14px}.hc .d{color:var(--ink-2);font-size:12.5px;margin:0;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.hc .p{color:var(--muted);font-size:11.5px;word-break:break-all;font-family:ui-monospace,Consolas,monospace}.hc .acts{display:flex;gap:6px;flex-wrap:wrap;margin-top:2px}
.hc details{box-shadow:none;margin:4px 0 0;border-radius:8px}.hc details>summary{padding:6px 10px;font-size:12.5px}.hc details>.in{padding:0 12px 10px;max-height:60vh;overflow:auto}
.hc.wide{grid-column:1/-1}.hc.idea{border-style:dashed;background:var(--surface-2)}
.btn.sm{padding:3px 10px;font-size:12px}.btn[disabled]{opacity:.45;cursor:not-allowed}.btn.quiet{border-color:var(--line);background:var(--surface);color:var(--ink-2)}
.lk{display:inline-block;font-size:11px;font-weight:800;border-radius:6px;padding:1px 7px;margin:8px 0 0}.lk.hard,.lk.soft{background:var(--bad-bg);color:var(--bad)}.lk.edit{background:var(--ok-bg);color:var(--ok)}
.hsec.locked{border-left:3px solid var(--bad);padding-left:10px;margin-left:-13px}.hsec.company{border-left:3px solid var(--accent);padding-left:10px;margin-left:-13px;background:var(--accent-soft);border-radius:0 8px 8px 0}.lk.co{background:var(--accent-soft);color:var(--accent-ink);border:1px solid var(--accent)}.hc.hasco{border-color:var(--accent)}.flow .node .lk.co{margin:4px 4px 0 0;font-size:10.5px}.flow.leadrow{margin-bottom:10px;padding-bottom:10px;border-bottom:1px dashed var(--line)}.flow .lbl{font-size:12px;color:var(--muted)}.corow{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:4px 0 10px}.cohd{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.diff .lk.co{margin:0 0 0 6px}details.drawer.adv>summary{font-size:14px;color:var(--ink-2);margin:18px 0 6px}.hsec.editable{border-left:3px solid var(--ok);padding-left:10px;margin-left:-13px}
.hsec>h2:first-of-type,.hsec>h3:first-of-type,.hsec>h4:first-of-type,.hsec>h5:first-of-type{margin-top:4px}
.kv{display:grid;grid-template-columns:auto 1fr;gap:4px 12px;font-size:12.5px;margin:4px 0}.kv span{color:var(--muted)}.kv code{word-break:break-all}
.nohit{color:var(--muted);font-size:12.5px;margin:0}
.bkbar{position:sticky;top:64px;z-index:8;display:flex;align-items:center;gap:10px;flex-wrap:wrap;background:var(--surface);border:1px solid var(--accent);border-radius:12px;padding:8px 14px;margin:0 0 14px;box-shadow:var(--shadow);font-size:13px}
.modal{position:fixed;inset:0;z-index:60;display:none;align-items:center;justify-content:center;background:rgba(0,0,0,.4);padding:16px}.modal.on{display:flex}
.mbox{background:var(--surface);color:var(--ink);border:1px solid var(--line);border-radius:14px;box-shadow:0 12px 40px rgba(0,0,0,.25);width:min(1100px,100%);max-height:calc(100vh - 32px);display:flex;flex-direction:column;overflow:hidden}
.mhd{display:flex;align-items:center;gap:10px;padding:12px 16px;border-bottom:1px solid var(--line);flex-wrap:wrap}.mhd b{font-size:15px}.mhd .p{font-family:ui-monospace,Consolas,monospace;font-size:12px;color:var(--muted);word-break:break-all;flex:1 1 200px}
.mbd{padding:12px 16px;overflow:auto;display:grid;gap:10px}.mft{display:flex;gap:8px;flex-wrap:wrap;align-items:center;padding:10px 16px;border-top:1px solid var(--line);background:var(--surface-2)}
.mbd textarea,.mbd input{width:100%;font:inherit;color:var(--ink);background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:8px 10px}
.mbd textarea.code{font:12.5px/1.55 ui-monospace,Consolas,"D2Coding",monospace;min-height:300px;resize:vertical;tab-size:2;white-space:pre;overflow:auto}
.mbd textarea:focus,.mbd input:focus{outline:2px solid var(--accent);outline-offset:-1px}.mbd label{font-size:12px;font-weight:700;color:var(--muted)}
.lockban{border-radius:10px;padding:10px 12px;background:var(--bad-bg);color:var(--bad);font-weight:700;font-size:13px}.lockban ul{margin:4px 0 0;padding-left:18px;font-weight:400;color:var(--ink)}
.lockban.soft{background:var(--warn-bg);color:var(--warn)}.okban{border-radius:10px;padding:8px 12px;background:var(--ok-bg);color:var(--ok);font-size:12.5px;font-weight:700}
.dv{font:12px/1.5 ui-monospace,Consolas,monospace;border:1px solid var(--line);border-radius:8px;max-height:280px;overflow:auto;background:var(--surface-2)}
.dv div{white-space:pre-wrap;word-break:break-all;padding:0 10px}.dv .a{background:var(--ok-bg);color:var(--ok)}.dv .r{background:var(--bad-bg);color:var(--bad);text-decoration:line-through;text-decoration-color:rgba(180,35,24,.35)}.dv .g{color:var(--muted);background:var(--surface)}.dv .x{font-weight:800}
.guide{font-size:12px;color:var(--ink-2)}.guide summary{font-weight:700;cursor:pointer}.guide ul{margin:4px 0 0;padding-left:18px}
.cpout{font:12px/1.5 ui-monospace,Consolas,monospace;min-height:120px}
abbr.g{text-decoration:none;border-bottom:1px dotted var(--muted);cursor:help}.gl{color:var(--muted);font-size:11px;font-weight:400;margin-left:4px;white-space:nowrap}
.cpm{display:block;color:var(--muted);font-size:12px;font-weight:400;line-height:1.45;margin:2px 0 4px}.kpi .cpm{font-size:11.5px}
.gloss{display:grid;gap:6px 18px;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));font-size:12.5px;margin:0}.gloss div{min-width:0}.gloss b{color:var(--ink);margin-right:6px}.gloss span{color:var(--ink-2)}
.crit{display:flex;flex-wrap:wrap;gap:6px;margin:6px 0 0}
@media (max-width:1100px){.dg{grid-template-columns:1fr 1fr}}
@media (max-width:900px){.side{transform:translateX(-100%);transition:transform .2s;box-shadow:0 0 30px rgba(0,0,0,.2)}body.nav-open .side{transform:none}
 body.nav-open .scrim{display:block;position:fixed;inset:0;background:rgba(0,0,0,.3);z-index:15}.col{margin-left:0}#menu{display:inline-flex}.top{padding:0 12px;gap:8px}.built{display:none}
 main{padding:16px 16px 48px}.docwrap{grid-template-columns:minmax(0,1fr)}.doclist{position:static;max-height:none}.two{grid-template-columns:1fr}.strip,.vban{padding-left:16px;padding-right:16px}}
@media (max-width:640px){.dg{grid-template-columns:1fr}.hgrid{grid-template-columns:1fr}.modal{padding:0}.mbox{max-height:100vh;height:100%;border-radius:0}.bkbar{top:60px}.decide,.card{padding:16px}.top h1{font-size:14.5px}#stat .pill{max-width:120px;overflow:hidden;text-overflow:ellipsis}.kpi{grid-template-columns:1fr 1fr;gap:10px}.kpi b{font-size:21px}}
@media print{.side,.top,.strip,.scrim,#menu,.modal,.bkbar,.acts{display:none!important}.col{margin-left:0}section.pane{display:block;break-before:page}section.pane:first-child{break-before:auto}
 .tw{max-height:none;overflow:visible}th{position:static}.card,.decide,details,.kpi>div{box-shadow:none;break-inside:avoid}body{background:#fff}}
"""


JS = r"""
const D = JSON.parse(document.getElementById('data').textContent);
const TABS = ['home','harness','flow','dash','out','doc','report','verify','changes','design','regress'];   // 저장된 탭이 이 목록 밖 (예전 부록 등) 이면 개요
let CO = D.companies.length ? D.companies[0].co : null, TAB = 'home', DOC = '', SEC = '', CLOSED = false, RESTORED = false, DETAIL = false;   // DETAIL = 자세히 보기 (기본 = 간단히)
try { const s = JSON.parse(localStorage.getItem('hrv') || '{}'); if (TABS.includes(s.tab)) TAB = s.tab; if (s.co && D.companies.some(c => c.co === s.co)) CO = s.co; if (s.doc) DOC = s.doc; DETAIL = s.detail === true; } catch (e) {}
// 교안 · 안내문에서 특정 화면으로 바로 열기: hr-viewer.html#tab=dash&co=K-co · 문서 한 절: #co=K-co&doc=②_수요파라미터.md&sec=7. 확인 필요
(() => { const h = new URLSearchParams(location.hash.slice(1)); const t = h.get('tab'), c = h.get('co');
  if (t && TABS.includes(t)) TAB = t; if (c && D.companies.some(x => x.co === c)) CO = c;
  if (h.get('doc')) { DOC = h.get('doc'); SEC = h.get('sec') || ''; if (!t) TAB = 'doc'; }
  const th = h.get('theme'); if (th === 'light' || th === 'dark') document.documentElement.dataset.theme = th;
  if (location.hash && !th) try { history.replaceState(null, '', location.pathname + location.search); } catch (e) {} })();   // 새 판 다시 읽기 때 주소의 절로 끌려가지 않게
const $ = (s, el = document) => el.querySelector(s), $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const f1 = v => v == null || v === '' || isNaN(v) ? '-' : Number(v).toLocaleString('ko-KR', {maximumFractionDigits: 1, minimumFractionDigits: 1});
const tip = $('#tip');
function showTip(e, h) { tip.style.whiteSpace = 'normal'; tip.innerHTML = h; tip.style.display = 'block'; const x = Math.min(e.clientX + 14, innerWidth - 300); tip.style.left = x + 'px'; tip.style.top = (e.clientY + 14) + 'px'; }
function hideTip() { tip.style.display = 'none'; }
function save() { try { localStorage.setItem('hrv', JSON.stringify({tab: TAB, co: CO, doc: DOC, detail: DETAIL})); } catch (e) {} }
const coName = co => /^k-co\.?$/i.test(String(co ?? '')) ? 'K-Co.' : String(co ?? '');   // 화면 표시만 (폴더 이름 · 자료 열쇠는 그대로)
const txt = n => ((n && n.textContent) || '').replace(/\s+/g, ' ').trim();

// ── 간단히 보기 (기본): 출처 · 근거 열은 첫 칸 ⓘ 로 접고, 빈 열은 숨기고, 긴 칸은 두 줄만 · 문서 머리 설명 · 괄호 속 참고도 ⓘ ──
// 「자세히」를 누르면 모두 펼침. 렌더할 때마다 .mdbody (결과 md 문서) 에만 다시 적용 · 하네스 · 차트 · 지금 정할 것 카드는 건드리지 않음
const REF = /\.(md|xlsx|json|py)\b|파일 ?\d|결정로그|검증로그|\d+ ?행\b|\d+ ?절\b|시트|출처|근거|스크립트|auto 블록/;
const SRC_COL = /^(출처|근거)|근거 파일|출처$/;
const META = /^(만든 스크립트|만든 곳|출처|근거 =|근거:|이 파일|상태:|기준일|자료 파일)/;
const SKIP = 'svg,code,pre,th,summary,nav,.toc,.mtoc,abbr,.gl,.nm,.ri,a,button,input,textarea,select,label';
function icon(t) { const s = document.createElement('span'); s.className = 'ri'; s.tabIndex = 0; s.setAttribute('role', 'note'); s.setAttribute('aria-label', '참고: ' + t); s.dataset.tip = t; return s; }
function declutter(root) {
  if (DETAIL || !root) return;
  $$('table', root).forEach(t => {
    if (t.dataset.dc || t.classList.contains('inp') || t.closest('svg')) return; t.dataset.dc = '1';
    const ths = $$('thead th', t), trs = $$('tbody tr', t);
    const blank = tr => [...tr.children].slice(1).every(td => !txt(td));   // 번호 칸만 있는 빈 양식 행
    trs.forEach(tr => { if (blank(tr)) tr.classList.add('dc-hide'); });
    if (trs.length && trs.every(blank)) { (t.closest('.tw') || t).classList.add('dc-hide'); return; }
    const live = trs.filter(tr => !blank(tr)), move = [], drop = [];
    ths.forEach((th, i) => { const h = txt(th);
      if (i > 0 && SRC_COL.test(h)) move.push(i);
      else if (i > 0 && live.every(tr => !(tr.children[i] && txt(tr.children[i])))) drop.push(i); });
    const hide = new Set([...move, ...drop]);
    trs.forEach(tr => { const cs = [...tr.children];
      const notes = move.map(i => cs[i] && txt(cs[i]) ? `${txt(ths[i])}: ${txt(cs[i])}` : '').filter(Boolean);
      cs.forEach((td, i) => { if (hide.has(i)) td.classList.add('dc-hide'); });
      if (notes.length && cs[0]) cs[0].append(icon(notes.join('\n'))); });
    ths.forEach((th, i) => { if (hide.has(i)) th.classList.add('dc-hide'); });
    $$('tbody td', t).forEach(td => { if (!td.classList.contains('dc-hide') && txt(td).length > 70 && !td.querySelector('table,.clamp')) td.innerHTML = `<div class="clamp" title="눌러서 펼치기 · 접기">${td.innerHTML}</div>`; });
  });
  $$('h2,h3,h4,h5', root).forEach(h => { if (h.dataset.dc) return; h.dataset.dc = '1'; const ul = h.nextElementSibling; if (!ul || ul.tagName !== 'UL') return;
    const meta = [...ul.children].filter(li => META.test(txt(li))); if (!meta.length) return;
    h.append(icon(meta.map(li => '· ' + txt(li)).join('\n'))); meta.forEach(li => li.classList.add('dc-hide'));
    if (meta.length === ul.children.length) ul.classList.add('dc-hide'); });
  const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {acceptNode: n => n.parentElement.closest(SKIP) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT});
  const ns = []; while (w.nextNode()) if (/\([^()]*\)/.test(w.currentNode.nodeValue)) ns.push(w.currentNode);
  ns.forEach(t => { const v = t.nodeValue, re = /\s*\(([^()]{2,220})\)/g; let m, last = 0, hit = false; const f = document.createDocumentFragment();
    while ((m = re.exec(v))) { if (!REF.test(m[1])) continue; hit = true; f.append(v.slice(last, m.index)); f.append(icon(m[1].trim())); last = m.index + m[0].length; }
    if (hit) { f.append(v.slice(last)); t.replaceWith(f); } });
  // 괄호 안에 문서 링크 · 코드가 섞여 글 조각이 나뉜 경우 「(②_대사.xlsx · ④_Gap대안.md 1절)」 도 ⓘ 하나로 (링크는 자세히 보기에서)
  $$('p,li,td,blockquote,.clamp', root).forEach(el => { if (el.closest(SKIP)) return;
    const kids = [...el.childNodes];
    for (let i = 0; i < kids.length; i++) {
      const n = kids[i]; if (n.nodeType !== 3) continue;
      const v = n.nodeValue, o = v.lastIndexOf('('); if (o < 0 || v.indexOf(')', o) >= 0) continue;
      let inner = v.slice(o + 1), end = null, ei = -1, j = i + 1;
      for (; j < kids.length; j++) { const k = kids[j];
        if (k.nodeType === 3) { const c = k.nodeValue.indexOf(')'), q = k.nodeValue.indexOf('('); if (q >= 0 && (c < 0 || q < c)) break; if (c >= 0) { inner += k.nodeValue.slice(0, c); end = k; ei = c; break; } inner += k.nodeValue; }
        else { if (k.nodeType !== 1 || k.matches('.ri,br,ul,ol,table,div,pre')) break; inner += k.textContent; } }
      inner = inner.replace(/\s+/g, ' ').trim();
      if (!end || inner.length < 2 || inner.length > 220 || !REF.test(inner)) continue;
      n.nodeValue = v.slice(0, o).replace(/\s+$/, '');
      for (let k = i + 1; k < j; k++) kids[k].remove();
      end.nodeValue = end.nodeValue.slice(ei + 1); end.before(icon(inner)); i = j - 1;
    } });
}
// RPU 코드: 점선 풀이 대신 이름을 작게 (덩이마다 첫 등장 · 이미 이름이 보이거나 코드가 5개 넘게 늘어선 목록이면 점선 풀이 그대로)
function nameRpu(root) {
  const N = ((cur() || {}).gl || {}).rpu || {}; if (!root || !Object.keys(N).length) return;
  const blocks = new Map();
  $$('abbr.g[data-rpu]', root).forEach(a => { if (a.dataset.nm || a.closest('th,svg,summary,nav,.mtoc,.toc,a,button,.rtoc')) return;
    const b = a.closest('td,li,p,h2,h3,h4,h5,blockquote,.what,div') || root; if (!blocks.has(b)) blocks.set(b, []); blocks.get(b).push(a); });
  blocks.forEach((list, b) => { list.forEach(a => { a.dataset.nm = '1'; }); if (new Set(list.map(a => a.dataset.rpu)).size > 4) return;
    const tr = b.tagName === 'TD' ? b.closest('tr') : null, seen = new Set();
    list.forEach(a => { const k = a.dataset.rpu, n = N[k]; if (!n || seen.has(k)) return; seen.add(k);
      const after = ((a.nextSibling && a.nextSibling.textContent) || '').trimStart();
      if (after.startsWith(n) || (tr && [...tr.children].some(td => td !== b && txt(td) === n))) return;
      a.classList.add('rp'); const s = document.createElement('small'); s.className = 'nm'; s.textContent = n; a.after(s); }); });
}
// 절 바로가기 (긴 문서 위쪽 알약 줄 · 제목 4개 이상) · 리포트 오른쪽 목차
let UID = 0, RIO = null;
const hid = h => h.id || (h.id = 'h' + (++UID));
const htext = h => { const c = h.cloneNode(true); $$('.gl,.nm,.ri', c).forEach(x => x.remove()); return txt(c); };
function heads(mb) { let hs = $$(':scope > h2, :scope > h3, :scope > h4', mb); if (hs[0] && hs[0].tagName === 'H2' && hs.filter(h => h.tagName === 'H2').length === 1) hs = hs.slice(1); return hs; }
function miniToc(root) {   // 2026-11-16: 절 바로가기 = 화면 위에 고정되는 탭 · 누르면 그 절만 보여 줌 (「전체」 = 모두)
  $$('.mdbody', root).forEach(mb => {
    const pv = mb.previousElementSibling; if (pv && pv.classList.contains('mtoc')) return;
    let hs = heads(mb); if (hs.length < 4) return;
    const top = ['H2', 'H3', 'H4'].find(t => hs.some(h => h.tagName === t)), l1 = hs.filter(h => h.tagName === top);
    if (l1.length >= 4 || hs.length > 18) hs = l1.length >= 2 ? l1 : hs;
    const doc = mb.closest('.doc'), dt = doc && $('.dochd .toc', doc); if (dt) dt.remove();
    const direct = hs.every(h => h.parentElement === mb);
    if (!direct) {   // 머리가 깊이 들어 있으면 예전처럼 바로가기만
      mb.insertAdjacentHTML('beforebegin', `<nav class="mtoc" aria-label="절 바로가기">${hs.map(h => `<a href="#${hid(h)}" data-go="${h.id}" title="${esc(htext(h))}">${esc(htext(h).replace(/\s*\(.*$/, '').slice(0, 28))}</a>`).join('')}</nav>`);
      return;
    }
    const nodes = [...mb.childNodes]; let cur = null;
    nodes.forEach(n => { const i = hs.indexOf(n); if (i >= 0) { cur = document.createElement('div'); cur.className = 'msec'; cur.dataset.k = i; mb.insertBefore(cur, n); } if (cur) cur.appendChild(n); });
    mb.insertAdjacentHTML('beforebegin', `<nav class="mtoc tabbed" aria-label="절 탭"><a href="#" data-k="all" class="on">전체</a>${hs.map((h, i) => `<a href="#" data-k="${i}" title="${esc(htext(h))}">${esc(htext(h).replace(/\s*\(.*$/, '').slice(0, 28))}</a>`).join('')}</nav>`);
  });
}
const stickY = el => el.getBoundingClientRect().top + scrollY - (parseInt(getComputedStyle(document.documentElement).getPropertyValue('--stick')) || 56) - 4;
function showSec(nav, k) {
  const mb = nav.nextElementSibling; if (!mb) return;
  $$('a', nav).forEach(a => a.classList.toggle('on', a.dataset.k === String(k)));
  $$(':scope > .msec', mb).forEach(x => x.hidden = k !== 'all' && x.dataset.k !== String(k));
  $$(':scope > :not(.msec)', mb).forEach(x => x.hidden = k !== 'all');
  const y = stickY(nav); if (scrollY > y) scrollTo({top: y});
}
function sheetTabs(root) {   // 엑셀 미리보기: 시트 = 고정 탭 (요약 · ○○표 · 첨부 데이터)
  $$('.xl', root).forEach(xl => { if (xl.dataset.tabbed) return; const ss = $$(':scope > .xsh', xl); if (ss.length < 2) return; xl.dataset.tabbed = '1';
    xl.insertAdjacentHTML('afterbegin', `<nav class="mtoc tabbed xtabs" aria-label="시트">${ss.map((x, i) => `<a href="#" data-x="${i}" class="${i ? '' : 'on'}">${esc(x.dataset.t)}</a>`).join('')}</nav>`);
    ss.forEach((x, i) => x.hidden = i > 0); });
}
document.addEventListener('click', e => {
  const a = e.target.closest('.mtoc.tabbed a'); if (!a) return; e.preventDefault(); e.stopPropagation();
  const nav = a.parentElement;
  if (a.dataset.x != null) { $$('a', nav).forEach(x => x.classList.toggle('on', x === a)); $$(':scope > .xsh', nav.parentElement).forEach((x, i) => x.hidden = String(i) !== a.dataset.x);
    const y = stickY(nav); if (scrollY > y) scrollTo({top: y}); return; }
  showSec(nav, a.dataset.k);
}, true);
function setStick() { const dc = $('#dc'); document.documentElement.style.setProperty('--stick', (56 + (dc ? dc.offsetHeight : 0)) + 'px'); }
const dl = (doc, sec, label) => `<a href="#" class="dl" data-doc="${esc(doc)}" data-sec="${esc(sec || '')}">${esc(label || doc)}</a>`;
function go(doc, sec) { DOC = doc; SEC = sec || ''; TAB = 'doc'; render(); }
const KIND = {CP: '체크포인트(CP)', '예외': '예외 · 데이터 이상', '관문': '관문 · 검증'};
// 용어 풀이: 화면에 코드만 두지 않는다 (GLOSS = render.py 의 잠금 용어 · 회사별 RPU 이름 · 공급 가정은 c.gl)
const G = D.gloss || {cp: {}, step: {}, method: {}, method_short: {}, b4: {}, term: {}, gap4: [], crit: []};
const KMEAN = {CP: () => G.cp_def, '예외': () => G.term['예외'], '관문': () => G.term['관문']};
const cpLbl = k => G.cp[k] ? `${k} · ${G.cp[k].n}` : String(k ?? '');
const stLbl = p => p.re ? `재승인 필요 · ${p.re}` : p.state;   // CP 상태 글 (데이터 반영 뒤 다시 받을 CP 는 근거 번호까지)
const reCp = (c, k) => c && (c.cps.find(p => p.cp === k) || {}).re;
const cpTip = k => G.cp[k] ? `체크포인트(CP) ${k} · ${G.cp[k].n} · 승인 ${G.cp[k].who} · ${G.cp[k].d}` : '';
const stepLbl = k => G.step[k] ? `${k} ${G.step[k]}` : String(k ?? '');
function lbl(code, c) {   // 코드 → 「코드 · 뜻」 (모르는 코드는 그대로)
  const k = String(code ?? ''), gl = (c || cur() || {}).gl || {};
  if (G.cp[k]) return cpLbl(k);
  if (G.step[k]) return stepLbl(k);
  if (G.method[k]) return `${k} ${G.method_short[k]}`;
  if (G.b4[k]) return `${k} (${G.b4[k]})`;
  if (gl.sup && gl.sup[k]) return `${k} · ${gl.sup[k]}`;
  if (gl.rpu && gl.rpu[k]) return `${k} · ${gl.rpu[k]}`;
  return k;
}
const gloss = (code, c) => { const k = String(code ?? ''), t = G.cp[k] ? cpTip(k) : G.method[k] ? `수요 방식 ${k} · ${G.method[k]}` : ''; const l = lbl(k, c);
  return t ? `<abbr class="g" title="${esc(t)}">${esc(l)}</abbr>` : esc(l); };
function cur() { return D.companies.find(c => c.co === CO); }
function stClass(s) { return s === '승인' ? 'st-승인' : (s === '승인 대기' || s === '대기' || s === '재승인 필요') ? 'st-대기' : (s === '자료 없음' || s === '앞 CP 먼저') ? 'st-자료' : 'st-기타'; }
function pillCls(s) { return s === '승인' ? 'ok' : (s === '승인 대기' || s === '대기' || s === '재승인 필요') ? 'warn' : (s === '자료 없음' || s === '앞 CP 먼저') ? 'gray' : 'bad'; }
const IC = {download: '<path d="M12 4v11M7 10l5 5 5-5M5 20h14"/>', home: '<path d="M3 10.5 12 3l9 7.5V21h-6v-6H9v6H3z"/>', harness: '<circle cx="9" cy="8" r="3"/><path d="M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6M16 4.6a3 3 0 0 1 0 6M21 20c0-2.6-1.6-4.8-4-5.6"/>',
  flow: '<path d="M5 21V4h11l-2 4 2 4H5"/>', dash: '<path d="M4 20V11M10 20V4M16 20v-8M2 20h20"/>', out: '<path d="M3 6h6l2 2h10v11H3z"/>',
  doc: '<path d="M6 3h9l4 4v14H6z"/><path d="M14 3v5h5M9 13h7M9 17h7"/>', report: '<path d="M5 3h14v18H5z"/><path d="M9 7h6M9 11h6M9 15h4"/>',
  verify: '<path d="M12 3l8 3v6c0 4.5-3.4 8.3-8 9-4.6-.7-8-4.5-8-9V6z"/><path d="m8.5 12 2.5 2.5 4.5-5"/>', changes: '<circle cx="6" cy="6" r="2.5"/><circle cx="18" cy="18" r="2.5"/><path d="M6 8.5V21M18 15.5V3"/>',
  regress: '<path d="M20 11a8 8 0 1 0-2.3 5.7M20 4v7h-7"/>', pause: '<path d="M9 5v14M15 5v14"/>', menu: '<path d="M4 7h16M4 12h16M4 17h16"/>',
  theme: '<circle cx="12" cy="12" r="8"/><path d="M12 4a8 8 0 0 1 0 16z" fill="currentColor"/>', inbox: '<path d="M3 13l3-8h12l3 8v6H3z"/><path d="M3 13h5l1 3h6l1-3h5"/>',
  design: '<path d="M6 3h9l4 4v14H6z"/><path d="M14 3v5h5"/><path d="m9 15 1.6 1.6L15 12"/>'};
const ic = k => `<svg class="i" viewBox="0 0 24 24" aria-hidden="true">${IC[k] || ''}</svg>`;
const NAV = [['작업', [['home', '개요'], ['flow', '진행 · 체크포인트'], ['dash', '대시보드']]], ['결과', [['out', '산출물'], ['doc', '문서'], ['report', '리포트'], ['verify', '검증']]],
  ['하네스 관리', [['harness', '하네스'], ['changes', '변경점'], ['regress', '회귀'], ['design', '설계서']]]];
const TITLE = Object.fromEntries(NAV.flatMap(g => g[1]));
const empty = (text, say) => `<div class="card"><div class="empty">${ic('inbox')}<p>${esc(text)}</p>${say ? `<p class="say">다음에 할 말: <b>「${esc(say)}」</b></p>` : ''}</div></div>`;
const subShort = sub => { if (!sub) return ''; const plain = sub.replace(/<[^>]+>/g, '').replace(/&[a-z]+;/g, ' ');
  if (plain.length <= 56 || /<a |<abbr /.test(sub)) return sub; const cut = sub.indexOf(' · '); return (cut > 0 && cut < 60 ? sub.slice(0, cut) : '') + itip(plain); };
const head = (t, sub, act) => `<div class="pagehd${act ? ' wact' : ''}"><div><h2>${esc(t)}</h2>${sub ? `<p>${subShort(sub)}</p>` : ''}</div>${act ? `<div class="pact">${act}</div>` : ''}</div>`;
const getBtns = (c, names) => names.flatMap(n => ((c && c.files) || {})[n] || []).map(f => `<a class="btn sm quiet" href="${f.href}" download="${esc(f.fn)}">${ic('download')}${esc(f.label.replace('로 받기', '').replace(' 받기', ''))}</a>`).join('');

function renderNav() {
  const c = cur();
  const badge = k => { if (k === 'harness' && ED.basket.length) return `<span class="badge warn" title="요청 바구니">${ED.basket.length}</span>`; if (!c) return ''; if (k === 'flow' && c.decision) { const d = c.decision; return `<span class="badge ${d.kind === 'CP' ? 'warn' : 'bad'}" title="${esc(d.kind === 'CP' ? cpTip(d.sid) : (KMEAN[d.kind] ? KMEAN[d.kind]() : ''))}">${esc(d.kind === 'CP' ? d.sid + (reCp(c, d.sid) ? ' 재승인' : ' 대기') : d.kind)}</span>`; }
    if (k === 'verify' && c.v_open) return `<span class="badge bad" title="미해결 ${c.v_open} · ${esc(G.term['관문'])}">${c.v_open}</span>`;
    if (k === 'verify' && c.v_wait) return `<span class="badge warn" title="미확인 ${c.v_wait} · ${esc(G.term['미확인'])}">${c.v_wait}</span>`;
    if (k === 'changes' && D.harness.diffs.length) return `<span class="badge acc pill">${D.harness.diffs.length}</span>`; return ''; };
  $('#nav').innerHTML = NAV.map(([g, items]) => `<h6>${g}</h6>` + items.map(([k, t]) => `<button data-tab="${k}" title="${esc(t)}" class="${k === TAB ? 'on' : ''}">${ic(k)}<span>${t}</span>${badge(k)}</button>`).join('')).join('');
  $('#ttl').textContent = TITLE[TAB] || '';
  let st = '<span class="pill gray">결과 없음</span>';
  if (c) { const d = c.decision, all = c.cps.every(p => p.state === '승인');
    st = d ? `<span class="pill ${d.kind === 'CP' ? 'warn' : 'bad'}" title="${esc(d.kind === 'CP' ? cpTip(d.sid) : (KMEAN[d.kind] ? KMEAN[d.kind]() : ''))}">${d.kind === 'CP' ? esc(cpLbl(d.sid)) + (reCp(c, d.sid) ? ' 재승인 필요' : ' 승인 대기') : esc(d.kind) + ' · ' + esc(d.sid)}</span>`
      : all ? `<span class="pill ok" title="${esc('체크포인트(CP) 5곳 모두 사람 승인 · ' + cpTip('CP5'))}">CP5 · 리포트 공유 승인 · 마침</span>` : '<span class="pill ok" title="승인 대기 · 관문 대기가 없음">멈춘 곳 없음</span>'; }
  $('#stat').innerHTML = st;
}

const itip = (t, lab) => t ? `<span class="itip" tabindex="0" data-tip="${esc(t)}">${lab || 'ⓘ'}</span>` : '';
const ANS = {'확정': '이대로 다음 단계로 감', '추가 확인': '자료 · 근거를 더 받고 다시 정함', '보류': '이번에는 정하지 않음 (다음 단계로 안 감)', '이관': '다른 승인자 · 부서로 넘김'};
function decideCard(c, d, noDef) {   // 2026-11-16 정리: 정할 것 한 줄 → 근거 · 선택지 · 답 형식 표 · 정의는 툴팁
  const tx = e => e.text_h != null ? e.text_h : esc(e.text.replace(/<br>/g, ' '));
  const ev = d.evidence.map(e => e.links.length ? `<li>${e.links.map((l, i) => i === 0 ? `<b>${dl(l.doc, l.sec, l.label)}</b>` : `<span class="jump">${dl(l.doc, l.sec, '→ ' + l.label.split(' › ').pop())}</span>`).join(' ')}${e.text !== e.links[0].doc ? `<div class="muted sm">${tx(e)}</div>` : ''}</li>` : `<li>${tx(e)}</li>`).join('');
  const isCp = d.kind === 'CP' && G.cp[d.sid], mean = isCp ? `${G.cp[d.sid].d} · 승인 ${G.cp[d.sid].who}` : '';
  const def = (KMEAN[d.kind] ? KMEAN[d.kind]() : '') + (mean ? ' · ' + d.sid + ': ' + mean : '');
  const chips = String(d.answer || '').split(/\s*·\s*/).filter(Boolean).map(a => `<span class="achip"${ANS[a] ? ` data-tip="${esc(ANS[a])}" tabindex="0"` : ''}>${esc(a)}</span>`).join('');
  return `<div class="decide"><div class="dh">${ic('pause')}<b>지금 정할 것</b><span class="pill ${d.kind === 'CP' ? 'warn' : 'bad'}">${esc(KIND[d.kind] || d.kind)} ${esc(isCp ? cpLbl(d.sid) : d.sid)}</span>${itip(def)}<span class="sp"></span><span class="muted sm">${esc(coName(c.co))}${d.time ? ' · ' + esc(d.time) : ''} · ${dl(d.src, '', '원문')}</span></div>
   <p class="dwhat">${d.what_h != null ? d.what_h : esc(d.what)}</p>
   <table class="dtab"><tbody><tr><th>근거</th><td><ul>${ev}</ul></td></tr>
   <tr><th>선택지</th><td><ol>${d.options.map((o, i) => `<li>${d.options_h ? d.options_h[i] : esc(o)}</li>`).join('')}</ol></td></tr>
   <tr><th>답하는 법</th><td><div class="achips">${chips}${itip('대화창에 「답 · 승인자 직책 · 이유」로 답하면 리드가 승인 칸과 결정로그에 적음 (예: 확정, 인사기획팀장, 동인 원칙 점검 통과)')}</div></td></tr></tbody></table></div>`;
}
function renderDecide() {
  const c = cur(), d = c && c.decision, el = $('#dc');
  if (!d || TAB === 'home' || TAB === 'flow') { el.innerHTML = ''; return; }
  const l = (d.evidence[0] && d.evidence[0].links[0]) || null;
  el.innerHTML = `<div class="strip">${ic('pause')}<b>지금 정할 것</b><span class="pill ${d.kind === 'CP' ? 'warn' : 'bad'}" title="${esc(d.kind === 'CP' ? cpTip(d.sid) : (KMEAN[d.kind] ? KMEAN[d.kind]() : ''))}">${esc(d.kind === 'CP' ? '체크포인트 ' + cpLbl(d.sid) : d.kind + ' · ' + d.sid)}</span><span class="what">${d.what_h != null ? d.what_h : esc(d.what)}</span>${l ? dl(l.doc, l.sec, '근거 ' + l.doc) : ''}<button class="btn ghost" data-tab="home">4요소 보기</button></div>`;
}

function stepper(c) {   // CP = 큰 원 (승인 ✓ · 승인자 · 날짜) · 단계 = 작은 점 (파일 수)
  const S = c.stages.slice(0, 5), P = Object.fromEntries(c.cps.map(p => [p.cp, p]));
  const seq = [['s', S[0]], ['c', P.CP1], ['s', S[1]], ['c', P.CP2], ['c', P.CP3], ['s', S[2]], ['s', S[3]], ['c', P.CP4], ['s', S[4]], ['c', P.CP5]];
  const items = seq.map(([k, x]) => { if (!x) return '';
    if (k === 's') { const cls = x.done === x.items.length ? 'ok' : x.done ? 'warn' : '', nm = x.st.length > 2 ? x.st : `${x.st} ${G.step[x.st] || x.name.split(' (')[0]}`;
      return `<div class="st sg ${cls}" title="${esc(x.st + ' ' + x.name)}"><div class="dw"><i class="dot"></i></div><div class="lbl">${esc(nm)}</div><div class="meta">파일 ${x.done}/${x.items.length}</div></div>`; }
    const cls = pillCls(x.state), ok = x.state === '승인';
    return `<div class="st cp ${ok ? 'ok' : cls === 'gray' ? '' : cls}" title="${esc(cpTip(x.cp))}"><div class="dw"><span class="dot">${ok ? '✓' : esc(x.cp.slice(2))}</span></div><div class="lbl">${esc(x.cp)} · ${esc(x.name)}</div><div class="meta${x.re ? ' re' : ''}">${esc(x.re ? '재승인 필요' : x.state)}</div>${x.re ? `<div class="meta who" title="${esc(x.re + ' · 결정로그')}">${esc(x.re)}</div>` : ''}${ok && (x.who || x.date) ? `<div class="meta who" title="${esc((x.who || '') + (x.date ? ' · ' + x.date : ''))}">${esc(x.who || '')}${x.date ? ' · ' + esc(x.date) : ''}</div>` : ''}</div>`; }).join('');
  const ok = c.cps.filter(p => p.state === '승인').length, re = c.cps.filter(p => p.re);
  return `<div class="card"><h3>단계 진행<span class="sub"><abbr class="g" title="${esc(G.cp_def)}">체크포인트(CP)</abbr> ${ok}/5 승인 · 큰 원 = CP · 작은 점 = 단계</span>${re.length ? ` <span class="pill warn" title="결정로그 「데이터 반영」 뒤 다시 승인받을 CP">재승인 필요 ${re.map(p => esc(p.cp)).join(' · ')}</span>` : ''}</h3><div class="steps">${items}</div></div>`;
}
function kpis(c) {   // 흐름 순서: 현원 → 결정 수요 → (디지털 별도 줄) → 공급 기본 → 공급 상승 → Gap
  const d = c.dash, ty = d.ty;
  const k = [['현원', d.hc, '', '기준일 RPU 인원 합'], [`결정 수요 사람 (${ty})`, d.dec_h, '', '방식 B · C 대사로 정한 사람 수요 합'], ['디지털 FTE (별도 줄)', d.dec_d, 'aside', '사람 FTE 에 더하지 않음'],
    [`공급 ${ty} · 기본`, d.sup, '', '채용 없이 남는 인원 · 과거 퇴직률'], [`공급 ${ty} · 상승`, d.sup_up, '', '퇴직률 상승 가정'], [`Gap ${ty}`, d.gap_total, 'hl', '결정 수요 사람 − 공급 기본']];
  return `<div class="kpi">${k.map(([a, v, cls, m]) => `<div class="${cls}"><span>${a}</span>${v == null ? '<b class="none">아직 없음</b>' : `<b>${f1(v)}</b>`}<span class="cpm">${esc(m)}</span></div>`).join('')}</div>`;
}
function vpills(c) {
  const T = G.term;
  return `<span class="pill ${c.v_open ? 'bad' : 'ok'}" title="관문 항목이 아직 풀리지 않음">미해결 ${c.v_open}</span> <span class="pill ${c.v_wait ? 'warn' : 'ok'}" title="${esc(T['미확인'])}">미확인 ${c.v_wait}</span> <span class="pill gray" title="${esc(T['인정'])}">인정 ${c.v_acc || 0}</span> <span class="pill gray">검증 행 ${c.v_all}</span> <span class="pill gray" title="${esc(T['예외'])}">예외 (데이터 이상 멈춤) ${c.exceptions}건</span>`;
}
function glossCard(c) {   // 개요 맨 아래 · 작은 용어 풀이 (코드가 뜻하는 것)
  const gl = (c && c.gl) || {}, T = G.term, row = (k, v) => `<div><b>${esc(k)}</b><span>${esc(v)}</span></div>`;
  const sup = Object.keys(gl.sup || {}).sort();
  return `<details class="card" style="padding:0"><summary>용어 풀이 <span class="muted" style="font-weight:400">· 화면에 나오는 코드가 뜻하는 것 (점선 밑줄 글자에 마우스를 올리거나 누르면 풀이)</span></summary><div class="in">
    <h4>체크포인트(CP) · 에이전트가 멈추고 사람 승인을 받는 자리</h4><div class="gloss">${Object.keys(G.cp).map(k => row(cpLbl(k), `승인 ${G.cp[k].who} · ${G.cp[k].d}`)).join('')}</div>
    <h4>단계</h4><div class="gloss">${Object.keys(G.step).map(k => row(k, G.step[k])).join('')}</div>
    <h4>수요 방식 · 대사</h4><div class="gloss">${Object.keys(G.method).map(k => row('방식 ' + k, G.method[k])).join('')}${row('대사', T['대사'])}${row('결정 수요', T['결정 수요'])}</div>
    <h4>공급</h4><div class="gloss">${row('기본 · 상승', T['기본 · 상승'])}${sup.map(k => row(k, gl.sup[k])).join('')}</div>
    <h4>Gap · 대안</h4><div class="gloss">${row('Gap 4유형', G.gap4.join(' · '))}${Object.keys(G.b4).map(k => row(k, G.b4[k])).join('')}</div>
    <h4>멈춤 · 검증</h4><div class="gloss">${['예외', '관문', '인정', '미확인'].map(k => row(k, T[k])).join('')}${row('검증 5기준', G.crit.join(' · '))}</div>
    <h4>그 밖</h4><div class="gloss">${['RPU', 'Pivotal Role', 'G1~G5', '자율성 L1'].map(k => row(k, T[k])).join('')}</div></div></details>`;
}
function renderHome() {
  const c = cur(), el = $('#p-home');
  if (!c) { el.innerHTML = head('개요') + empty('outputs/ 아래 회사 결과가 아직 없음', 'K-Co 사업계획 분석부터 시작해 줘'); return; }
  const next = !c.decision ? (c.cps.every(p => p.state === '승인') ? '' : c.cps[0].state === '자료 없음' ? `${c.co} 사업계획 분석부터 시작해 줘` : '다음 단계 진행해 줘') : '';
  el.innerHTML = head('개요', `${esc(coName(c.co))} · 계획연도 ${esc(c.dash.ty)} · ${esc(D.built)} 기준`) +
    (c.decision ? decideCard(c, c.decision) : `<div class="card"><span class="pill ok">멈춘 곳 없음</span> <span class="muted">승인 대기 · 관문 대기가 없음${next ? ` · 다음에 할 말: 「${esc(next)}」` : ''}</span></div>`) +
    kpis(c) + stepper(c) +
    `<div class="two"><div class="card"><h3>담당이 짚어 둔 것<span class="sub">파일 안 「확인 필요」 · 「사람에게 물을 것」</span></h3>${c.checks.length ? `<ul>${c.checks.map(k => `<li>${dl(k.doc, k.sec, k.doc + ' › ' + k.label)}</li>`).join('')}</ul>` : '<p class="muted">없음</p>'}</div>
     <div class="card"><h3>검증<span class="sub">관문 = 검증 담당의 5기준 점검 · 통과해야 다음 단계</span></h3><p>${vpills(c)}</p>
     <p class="muted">미해결 · 미확인이 있으면 다음 단계로 가지 않음 · <a href="#" class="dl" data-doc="검증로그.md">검증로그 보기</a></p></div></div>` + glossCard(c);
}

// ---------------------------------------------------------------------------
// 하네스 관리 (보기 · 수정 초안 · 요청 복사) · 이 화면은 파일을 쓰지 않는다. 적용은 Claude 가 /hr-customize 로 (잠금 검사 · K-Co. 회귀 · 변경 이력)
// ---------------------------------------------------------------------------
const HF = {};
[D.harness.claude, ...D.harness.agents, ...D.harness.skills, ...(D.harness.settings || [])].forEach(f => { if (f && f.file) HF[f.file] = f; });
const EK = 'hrv_edit';
let ED = {drafts: {}, basket: []}, CURF = null, CURIDEA = null, EDT = 0, RELOAD_WAIT = false, LAST = null;
try { const v = JSON.parse(localStorage.getItem(EK) || '{}'); if (v && typeof v === 'object') ED = {drafts: (v.drafts && typeof v.drafts === 'object') ? v.drafts : {}, basket: Array.isArray(v.basket) ? v.basket : []}; } catch (e) {}
function edSave() { try { localStorage.setItem(EK, JSON.stringify(ED)); } catch (e) {} }
const hsh = s => { let h = 5381; for (let i = 0; i < s.length; i++) h = (h * 33 + s.charCodeAt(i)) | 0; return (h >>> 0).toString(36) + ':' + s.length; };
const lines = s => String(s ?? '').replace(/\r\n?/g, '\n').split('\n');

function ldiff(a, b) {   // 줄 단위 차이 · ' ' 같음 · '-' 지움 (o = 원본 줄) · '+' 더함 (at = 그 앞까지 원본 줄 수)
  let p = 0; while (p < a.length && p < b.length && a[p] === b[p]) p++;
  let q = 0; while (q < a.length - p && q < b.length - p && a[a.length - 1 - q] === b[b.length - 1 - q]) q++;
  const A = a.slice(p, a.length - q), B = b.slice(p, b.length - q), n = A.length, m = B.length, ops = [];
  for (let i = 0; i < p; i++) ops.push({t: ' ', s: a[i], o: i + 1});
  if (n * m > 4e6) { A.forEach((s, i) => ops.push({t: '-', s, o: p + i + 1})); B.forEach(s => ops.push({t: '+', s, at: p + n})); }
  else {
    const W = m + 1, L = new Uint32Array((n + 1) * W);
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i * W + j] = A[i] === B[j] ? L[(i + 1) * W + j + 1] + 1 : Math.max(L[(i + 1) * W + j], L[i * W + j + 1]);
    let i = 0, j = 0;
    while (i < n || j < m) {
      if (i < n && j < m && A[i] === B[j]) { ops.push({t: ' ', s: A[i], o: p + i + 1}); i++; j++; }
      else if (i < n && (j >= m || L[(i + 1) * W + j] >= L[i * W + j + 1])) { ops.push({t: '-', s: A[i], o: p + i + 1}); i++; }
      else { ops.push({t: '+', s: B[j], at: p + i}); j++; }
    }
  }
  for (let k = 0; k < q; k++) ops.push({t: ' ', s: a[a.length - q + k], o: a.length - q + k + 1});
  return ops;
}
const secAt = (f, n) => (f.secs || []).find(x => x.s <= n && n <= x.e);
const secName = s => s ? (s.h || '머리말') : 'frontmatter';
function opLine(o) { return o.t === '+' ? o.at : o.o; }
function lockEval(f, orig, now, ops) {   // customize.py lock 과 같은 기준 (hard) + 규칙상 잠금 절 (soft)
  const hits = [], seen = new Set(), add = (lvl, msg) => { if (!seen.has(msg)) { seen.add(msg); hits.push({lvl, msg}); } };
  (f.phr || []).forEach(x => { if (orig.includes(x.ph) && !now.includes(x.ph)) add('hard', `잠금 문장 「${x.ph}」 이 사라짐`); });
  const soft = new Map((f.soft || []).map(x => [x.n, x.why])), phrN = new Map(), okN = new Set((f.el || []).flatMap(x => x.lines));
  (f.phr || []).forEach(x => x.lines.forEach(n => phrN.set(n, x.ph)));
  ops.forEach(o => {
    if (o.t === '-') { const s = okN.has(o.o) ? null : secAt(f, o.o);
      if (s && s.lock === 'hard') add('hard', `「${secName(s)}」 절의 줄을 지우거나 바꿈`);
      else if (s && s.lock) add('soft', `「${secName(s)}」 절 (규칙상 잠금)`);
      if (soft.has(o.o)) add('soft', soft.get(o.o));
      if (phrN.has(o.o) && !okN.has(o.o) && now.includes(phrN.get(o.o))) add('soft', `잠금 문장 「${phrN.get(o.o)}」 이 있는 줄이 바뀜 (문장은 남음)`);
    } else if (o.t === '+') { const s = (okN.has(o.at) || okN.has(o.at + 1)) ? null : secAt(f, Math.max(o.at, 1)); if (s && s.lock) add('soft', `「${secName(s)}」 잠금 절 안에 줄 추가`); }
  });
  return hits;
}
function hunks(ops, ctx) {   // 바뀐 줄 앞뒤 ctx 줄씩 묶음
  const ch = ops.map((o, i) => o.t !== ' ' ? i : -1).filter(i => i >= 0), out = [];
  ch.forEach(i => { const a = Math.max(0, i - ctx), b = Math.min(ops.length - 1, i + ctx), g = out[out.length - 1]; if (g && a <= g[1] + 1) g[1] = Math.max(g[1], b); else out.push([a, b]); });
  return out;
}
function diffText(f, ops) {
  return hunks(ops, 1).map(([a, b]) => { const first = ops.slice(a, b + 1).find(o => o.t !== ' ') || ops[a], ln = Math.max(1, opLine(first) + (first.t === '+' ? 1 : 0));
    return [`@@ ${ln}줄 근처 · ${secName(secAt(f, Math.max(opLine(first), 1)))}`, ...ops.slice(a, b + 1).map(o => (o.t === ' ' ? '  ' : o.t + ' ') + o.s)].join('\n'); }).join('\n');
}
function touched(f, ops) { return [...new Set(ops.filter(o => o.t !== ' ').map(o => secName(secAt(f, Math.max(opLine(o), 1)))))]; }
const lockLine = hits => hits.length ? `있음 (${hits.map(h => (h.lvl === 'hard' ? '잠금 검사에서 막힘 · ' : '규칙상 잠금 · ') + h.msg).join(' · ')})` : '없음';
const REQ_HEAD = '/hr-customize 하네스 수정 요청';
const REQ_TAIL = '처리: 작업 폴더 사본에 적용 → 잠금 검사 (customize.py lock) → K-Co. 회귀 (customize.py regress) → CLAUDE.md 변경 이력 한 행 → 뷰어 다시 만들기 · 잠금 칸 변경이 있으면 적용 전에 멈추고 4요소로 묻기 · 옛 줄이 지금 파일과 다르면 멈추고 알리기';
function itemText(it) {
  return [`대상: ${it.target}`, `사유: ${it.reason || '(적지 않음 · 사유를 물어 주세요)'}`, it.kind === 'idea' ? '바꿀 내용 (제안):' : '바꿀 내용 (diff):', it.body, `잠금 칸 변경: ${it.lock}`].join('\n');
}
function reqText(items) {
  if (items.length === 1) return [REQ_HEAD, itemText(items[0]), `원본 기준: 뷰어 ${D.built} 판`, REQ_TAIL].join('\n');
  return [`${REQ_HEAD} (${items.length}건 · 한 건씩 차례로)`, `원본 기준: 뷰어 ${D.built} 판`, ...items.map((it, i) => `\n[${i + 1}/${items.length}]\n` + itemText(it)), '', REQ_TAIL].join('\n');
}
async function copyText(t, fb, msg, okHtml) {
  let ok = false;
  try { if (navigator.clipboard && navigator.clipboard.writeText) { await navigator.clipboard.writeText(t); ok = true; } } catch (e) {}
  if (!ok) { try { fb.value = t; fb.hidden = false; fb.focus(); fb.select(); ok = !!(document.execCommand && document.execCommand('copy')); } catch (e) {} }
  if (ok) { fb.hidden = true; msg.innerHTML = okHtml || '<span class="ok">복사함</span> · Claude Code 대화창에 붙여 넣으면 Claude 가 /hr-customize 로 적용합니다'; }
  else { fb.value = t; fb.hidden = false; fb.select(); msg.innerHTML = '<span class="warn">자동 복사가 막힘</span> · 아래 글을 골라 두었으니 Ctrl+C (Mac ⌘C)'; }
  return ok;
}

// 편집 창 (원문 그대로 · 고정폭) ------------------------------------------------
function edHtml() {
  return `<div class="mbox" role="dialog" aria-modal="true" aria-labelledby="ed-t"><div class="mhd"><b id="ed-t">수정 초안</b><span class="p" id="ed-p"></span><button class="btn quiet sm" id="ed-x">닫기</button></div>
   <div class="mbd"><div id="ed-new" class="lockban soft" hidden>새 판 뷰어가 만들어졌습니다 · 창을 닫으면 다시 읽습니다 (초안은 이 브라우저에 남음)</div>
    <div id="ed-note"></div><details class="guide" id="ed-g"><summary>이 파일의 잠금 칸 · 변형 메뉴 칸</summary><div id="ed-gl"></div></details>
    <label for="ed-a" id="ed-al">원문 (마크다운 그대로 · 고친 것은 이 브라우저에만 초안으로 남고 파일은 바뀌지 않음)</label><textarea id="ed-a" class="code" spellcheck="false"></textarea>
    <label for="ed-r">사유 (왜 바꾸는가 · 변경 이력 「사유」 칸에 들어감)</label><input id="ed-r" placeholder="예: 자사 영업은 수주 기준으로 관리 · 인사팀장 확인">
    <div id="ed-b"></div><label id="ed-nl">바뀐 줄</label><div class="dv" id="ed-d"></div>
    <textarea id="ed-c" class="cpout" readonly hidden aria-label="복사할 글"></textarea></div>
   <div class="mft"><button class="btn quiet sm" id="ed-0">원본으로 되돌리기</button><span class="sp"></span><span id="ed-m" class="muted"></span>
    <button class="btn ghost sm" id="ed-k">요청 바구니에 담기</button><button class="btn sm" id="ed-cp">Claude 에게 맡길 문장 복사</button></div></div>`;
}
function guideList(f) {
  const L = (f.secs || []).filter(x => x.lock || x.edit).map(x => `<li>${x.lock ? '🔒' : '✏️'} ${esc(secName(x))} <span class="muted">${x.s}~${x.e}줄 · ${esc(x.why)}</span></li>`);
  (f.phr || []).forEach(x => L.push(`<li>🔒 잠금 문장 「${esc(x.ph)}」 <span class="muted">${x.lines.join(', ')}줄 · 지우면 customize.py lock 위반</span></li>`));
  (f.soft || []).forEach(x => L.push(`<li>🔒 ${x.n}줄 <span class="muted">${esc(x.why)}</span></li>`));
  (f.el || []).forEach(x => L.push(`<li>✏️ ${x.lines.join(', ')}줄 <span class="muted">${esc(x.why)}</span></li>`));
  return L.length ? `<ul>${L.join('')}</ul>` : '<p class="nohit">표시된 잠금 칸 · 변형 메뉴 칸 없음 (바꿀 수 있으나 변경 이력 · 잠금 검사 · K-Co. 회귀는 그대로 거침)</p>';
}
function edOpen() { return $('#ed').classList.contains('on'); }
function edShow(on) { $('#ed').classList.toggle('on', on); document.body.style.overflow = on ? 'hidden' : ''; }
function openEditor(file) {
  const f = HF[file]; if (!f) return; CURF = f; CURIDEA = null;
  const dr = ED.drafts[file]; let text = f.raw, reason = '', note = '';
  if (dr) { if (dr.base === hsh(f.raw)) { text = dr.text; reason = dr.reason || ''; note = '<div class="okban">이 브라우저에 남은 초안을 이어서 엽니다</div>'; }
    else { note = `<div class="lockban soft">원본이 바뀌어 지난 초안을 쓰지 않았습니다 (Claude 가 이미 고쳤을 수 있음)<details class="guide"><summary>지난 초안 보기</summary><textarea class="cpout" readonly>${esc(dr.text)}</textarea></details></div>`; delete ED.drafts[file]; edSave(); } }
  $('#ed-t').textContent = '수정 초안'; $('#ed-p').textContent = f.file; $('#ed-note').innerHTML = note + (f.err ? `<div class="lockban soft">${esc(f.err)}</div>` : '');
  $('#ed-g').hidden = false; $('#ed-gl').innerHTML = guideList(f); $('#ed-al').hidden = false;
  $('#ed-a').value = text; $('#ed-r').value = reason; $('#ed-c').hidden = true; $('#ed-m').textContent = ''; $('#ed-0').hidden = false; $('#ed-nl').hidden = false; $('#ed-d').hidden = false;
  edShow(true); try { sessionStorage.setItem('hrv_ed_open', file); } catch (e) {}
  edUpdate(); $('#ed-a').focus();
}
function openIdea(id) {
  const x = IDEAS.find(i => i.id === id); if (!x) return; CURF = null; CURIDEA = x;
  $('#ed-t').textContent = x.lock ? '🔒 요청 만들기' : '요청 만들기'; $('#ed-p').textContent = x.target;
  $('#ed-note').innerHTML = `<div class="${x.lock ? 'lockban' : 'okban'}">${esc(x.t)} · ${esc(x.note)}</div>`;
  $('#ed-g').hidden = true; $('#ed-al').hidden = false; $('#ed-a').value = x.body.join('\n'); $('#ed-r').value = '';
  $('#ed-c').hidden = true; $('#ed-m').textContent = ''; $('#ed-0').hidden = true; $('#ed-nl').hidden = true; $('#ed-d').hidden = true;
  edShow(true); edUpdate(); $('#ed-r').focus();
}
function edClose() {
  edShow(false); CURF = null; CURIDEA = null; try { sessionStorage.removeItem('hrv_ed_open'); } catch (e) {}
  if (RELOAD_WAIT) { CLOSED = true; try { sessionStorage.setItem('hrv_y', String(scrollY)); } catch (e) {} location.reload(); }
  else renderHarness();
}
function edItem() {   // 지금 창의 요청 한 건 (바뀐 것이 없으면 null)
  const reason = $('#ed-r').value.trim();
  if (CURIDEA) { const x = CURIDEA, body = lines($('#ed-a').value).filter(s => s.trim()).map(s => '+ ' + s).join('\n');
    return body ? {id: x.id + ':' + Date.now(), kind: 'idea', file: x.target, target: x.target, reason, body, lock: x.lock ? `있음 (${x.lock})` : `없음${x.lockNote ? ' (' + x.lockNote + ')' : ''}`, hard: !!x.lock} : null; }
  const f = CURF, now = $('#ed-a').value, ops = ldiff(lines(f.raw), lines(now));
  if (!ops.some(o => o.t !== ' ')) return null;
  const hits = lockEval(f, f.raw, now, ops), t = touched(f, ops);
  return {id: f.file + ':' + Date.now(), kind: 'diff', file: f.file, target: `${f.file} (${t.slice(0, 4).join(' · ')}${t.length > 4 ? ' 외 ' + (t.length - 4) : ''})`, reason, body: diffText(f, ops), lock: lockLine(hits), hard: hits.some(h => h.lvl === 'hard'), base: hsh(f.raw)};
}
function edUpdate() {
  let hits = [];
  if (CURIDEA) { if (CURIDEA.lock) hits = [{lvl: 'hard', msg: CURIDEA.lock}]; }
  else if (CURF) {
    const f = CURF, now = $('#ed-a').value, ops = ldiff(lines(f.raw), lines(now)), ch = ops.filter(o => o.t !== ' ');
    hits = lockEval(f, f.raw, now, ops);
    const add = ch.filter(o => o.t === '+').length, del = ch.length - add;
    $('#ed-nl').innerHTML = ch.length ? `바뀐 줄 <span class="add">+${add}</span> <span class="del">−${del}</span> · ${esc(touched(f, ops).join(' · '))}` : '바뀐 줄 · 아직 없음';
    const okN = new Set((f.el || []).flatMap(x => x.lines)), bad = new Set(); ops.forEach((o, i) => { if (o.t === '-') { const s = okN.has(o.o) ? null : secAt(f, o.o); if ((s && s.lock) || (f.soft || []).some(x => x.n === o.o) || (!okN.has(o.o) && (f.phr || []).some(x => x.lines.includes(o.o)))) bad.add(i); }
      else if (o.t === '+') { const s = (okN.has(o.at) || okN.has(o.at + 1)) ? null : secAt(f, Math.max(o.at, 1)); if (s && s.lock) bad.add(i); } });
    const HK = hunks(ops, 2).slice(0, 60);
    $('#ed-d').innerHTML = ch.length ? HK.map(([a, b], k) => (k ? '<div class="g">…</div>' : '') + ops.slice(a, b + 1).map((o, j) => {
      const i = a + j, ln = o.t === '+' ? '    ' : String(o.o).padStart(4);
      return `<div class="${o.t === '+' ? 'a' : o.t === '-' ? 'r' : ''}${bad.has(i) ? ' x' : ''}">${ln} ${o.t === ' ' ? ' ' : o.t} ${bad.has(i) ? '🔒 ' : ''}${esc(o.s) || ' '}</div>`; }).join('')).join('') : '<div class="g">원문과 같음</div>';
    clearTimeout(EDT); EDT = setTimeout(() => { const r = $('#ed-r').value;
      if (now === f.raw && !r.trim()) delete ED.drafts[f.file]; else ED.drafts[f.file] = {text: now, reason: r, base: hsh(f.raw), at: Date.now()}; edSave(); }, 200);
  }
  const hard = hits.some(h => h.lvl === 'hard');
  $('#ed-b').innerHTML = hits.length ? `<div class="lockban${hard ? '' : ' soft'}">${hard ? '이 변경은 잠금 검사에서 막힙니다 · 그래도 요청하면 Claude 가 이유를 묻고 멈춥니다' : '잠금 칸을 건드립니다 · 요청하면 Claude 가 적용 전에 이유를 묻고 멈춥니다'}<ul>${hits.map(h => `<li>${h.lvl === 'hard' ? '🔒 잠금 검사' : '🔒 규칙상 잠금'} · ${esc(h.msg)}</li>`).join('')}</ul></div>` :
    (CURF ? '<div class="okban">잠금 칸 변경 없음 · 적용 때 잠금 검사 · K-Co. 회귀를 그대로 거칩니다</div>' : '');
  const it = edItem(); $('#ed-cp').disabled = !it; $('#ed-k').disabled = !it;
}

// 요청 바구니 ---------------------------------------------------------------
function bkHtml() {
  const B = ED.basket;
  return `<div class="mbox" role="dialog" aria-modal="true" aria-labelledby="bk-t"><div class="mhd"><b id="bk-t">요청 바구니 ${B.length}건</b><span class="p">모아 둔 수정 요청을 한 번에 Claude 에게 맡김</span><button class="btn quiet sm" id="bk-x">닫기</button></div>
   <div class="mbd">${B.length ? `<div class="hgrid">${B.map((it, i) => `<div class="hc"><div class="t"><b>${i + 1}. ${esc(it.file)}</b>${!it.lock.startsWith('없음') ? `<span class="pill ${it.hard ? 'bad' : 'warn'}">🔒 잠금 칸</span>` : ''}</div>
     <p class="d">${esc(it.reason || '사유 없음')}</p><div class="dv" style="max-height:140px">${lines(it.body).map(s => `<div class="${s.startsWith('+') ? 'a' : s.startsWith('-') ? 'r' : s.startsWith('@@') ? 'g' : ''}">${esc(s) || ' '}</div>`).join('')}</div>
     <div class="acts"><button class="btn quiet sm" data-bkdel="${i}">빼기</button></div></div>`).join('')}</div>` : '<p class="muted">비어 있음 · 하네스 화면에서 「수정」 → 「요청 바구니에 담기」</p>'}
    <textarea id="bk-c" class="cpout" readonly hidden aria-label="복사할 글"></textarea></div>
   <div class="mft"><button class="btn quiet sm" id="bk-0" ${B.length ? '' : 'disabled'}>바구니 비우기</button><span class="sp"></span><span id="bk-m" class="muted"></span><button class="btn sm" id="bk-cp" ${B.length ? '' : 'disabled'}>모두 복사 (${B.length}건)</button></div></div>`;
}
function openBasket() { $('#bk').innerHTML = bkHtml(); $('#bk').classList.add('on'); document.body.style.overflow = 'hidden'; }
function closeBasket() { $('#bk').classList.remove('on'); document.body.style.overflow = ''; renderHarness(); renderNav(); }

// 훅 · MCP · 권한 후보 (요청으로 만들 수 있는 카드) ---------------------------------
const HOOK_NOTE = '훅은 플러그인이 아니라 작업 폴더 .claude/settings.json 에 둔다 (회사 · 폴더마다 다름 · 플러그인 업데이트에 덮이지 않음)';
const IDEAS = [
  {id: 'hook-cp5', grp: 'hook', t: 'CP5 (리포트 공유) 승인 전 반출 차단', ev: 'PreToolUse', m: 'Bash · Write', target: '.claude/settings.json (hooks · PreToolUse · 새로 추가)', note: HOOK_NOTE, lockNote: '잠금 규칙 「CP5 전 반출 금지」를 장치로 받침',
   d: '리포트 공유 승인 전에 리포트 · 뷰어를 작업 폴더 밖으로 내보내는 명령을 막는다. 지금은 CLAUDE.md 규칙 (지시문)으로만 지킨다',
   body: ['훅 추가: CP5 (리포트 공유) 승인 전 반출 차단', '이벤트: PreToolUse · matcher: Bash|Write', '하는 일: 리포트_*.md · hr-viewer.html 을 작업 폴더 밖으로 복사 · 이동 · 업로드 · 메일 보내는 Bash 명령, 작업 폴더 밖 경로에 쓰는 Write 를 막는다 (종료 코드 2 · 이유 안내)', '통과 조건: outputs/<회사>/리포트_*.md 「## 승인」 표 CP5 답이 「확정」', '놓는 곳: 작업 폴더 .claude/settings.json + 훅 스크립트 .claude/hooks/ (플러그인에는 넣지 않음)', '시험: 승인 전 복사 시도 1회 막힘 · 승인 뒤 통과 확인']},
  {id: 'hook-xlsx', grp: 'hook', t: '결과 엑셀 손 수정 차단', ev: 'PreToolUse', m: 'Write · Edit', target: '.claude/settings.json (hooks · PreToolUse · 새로 추가)', note: HOOK_NOTE, lockNote: '「숫자는 스크립트가 만든다」를 장치로 받침',
   d: 'outputs/**/*.xlsx 를 담당이 Write · Edit 로 직접 고치지 못하게 한다. 숫자는 단계 스크립트만 만든다 (같은 지시 = 같은 숫자)',
   body: ['훅 추가: 결과 엑셀 손 수정 차단', '이벤트: PreToolUse · matcher: Write|Edit', '하는 일: 대상 경로가 outputs/**/*.xlsx 이면 막고 「파라미터 · 가정 파일을 고친 뒤 단계 스크립트로 다시 계산」을 안내 (종료 코드 2)', '놓는 곳: 작업 폴더 .claude/settings.json + .claude/hooks/ (플러그인에는 넣지 않음)', '시험: xlsx 에 Write 시도 1회 막힘 · md 쓰기는 통과']},
  {id: 'hook-view', grp: 'hook', t: '단계 끝 · 멈춤마다 뷰어 자동 갱신', ev: 'SubagentStop · Stop', m: '(전체)', target: '.claude/settings.json (hooks · SubagentStop · Stop · 새로 추가)', note: HOOK_NOTE, lockNote: '',
   d: '담당이 끝나거나 세션이 멈출 때 render.py 를 돌린다. 리드가 잊어도 열린 뷰어 탭이 새 판을 스스로 읽는다',
   body: ['훅 추가: 단계 끝 · 멈춤마다 뷰어 자동 갱신', '이벤트: SubagentStop · Stop', '하는 일: python -X utf8 .claude/hr-lib/render.py --no-open (열린 탭이 2초마다 새 판을 읽으므로 브라우저를 새로 띄우지 않음)', '실패해도 작업을 막지 않는다 (종료 코드 0 · 오류는 화면에 한 줄)', '놓는 곳: 작업 폴더 .claude/settings.json (플러그인에는 넣지 않음)']},
  {id: 'hook-lockfile', grp: 'hook', t: '엑셀 잠금 파일 알림', ev: 'PreToolUse', m: 'Bash', target: '.claude/settings.json (hooks · PreToolUse · 새로 추가)', note: HOOK_NOTE, lockNote: '',
   d: 'data/ 에 엑셀이 열려 있으면 생기는 ~$*.xlsx 를 단계 스크립트 실행 전에 알려 준다 (엑셀이 파일을 잡고 있으면 읽기 · 저장이 실패함)',
   body: ['훅 추가: 엑셀 잠금 파일 알림', '이벤트: PreToolUse · matcher: Bash', '하는 일: 명령이 .claude/skills/*/scripts/*.py 를 실행하고 data/<회사>/ 또는 outputs/<회사>/ 에 ~$*.xlsx 가 있으면 「엑셀에서 그 파일을 닫고 다시」를 알린다', '막을지 알림만 할지: 알림만 (사람이 고름)', '놓는 곳: 작업 폴더 .claude/settings.json + .claude/hooks/ (플러그인에는 넣지 않음)']},
  {id: 'mcp-add', grp: 'mcp', t: 'MCP 서버 추가 (회사 정책 결정 필요)', target: '.mcp.json (새로 추가) + CLAUDE.md 「업무 범위 · 금지」', note: 'MCP 를 붙이면 data/<회사>/ 밖 자료를 읽게 되어 CLAUDE.md 잠금 칸을 바꿔야 한다. 회사 정책 · 승인자 결정이 먼저다',
   lock: 'CLAUDE.md 「업무 범위 · 금지」 · data/<회사>/ 폴더 밖 파일을 읽지 않는다 · 외부 자료 검색 금지',
   d: '사내 HR 시스템 · 문서 저장소를 MCP 로 붙이는 요청. 잠금 칸을 바꾸는 일이므로 Claude 가 적용 전에 멈추고 묻는다',
   body: ['MCP 서버 추가 요청', '서버 이름 · 종류 (stdio · http): ', '읽을 자료와 범위 (개인 식별정보 제외 · 집계 단위만): ', '회사 정책 근거 · 승인자 (직책): ', '바꿔야 하는 잠금 칸: CLAUDE.md 「업무 범위 · 금지」 data/<회사>/ 밖 읽기 금지 문장', '놓는 곳: 작업 폴더 .mcp.json (플러그인에는 넣지 않음)']},
  {id: 'perm-web', grp: 'perm', t: '인터넷 검색 · 가져오기 막기', target: '.claude/settings.json (permissions.deny · 새로 추가)', note: '권한 설정도 작업 폴더 .claude/settings.json 에 둔다 (플러그인에는 넣지 않음)', lockNote: '잠금 규칙 「외부 자료 검색 금지」를 권한으로 받침',
   d: 'CLAUDE.md 의 「외부 자료 검색 금지」를 권한 설정으로도 막는다. 지금은 지시문 규칙만 있다',
   body: ['권한 추가: permissions.deny 에 WebFetch · WebSearch', '하는 일: 담당이 인터넷 검색 · 웹 페이지 가져오기를 시도하면 Claude Code 가 막음', '놓는 곳: 작업 폴더 .claude/settings.json (플러그인에는 넣지 않음)', '시험: 「인터넷에서 업계 평균 찾아 줘」 → 거절 확인 (검증 「범위 밖 거절」 행동 테스트와 같이)']},
];

// 하네스 화면 ---------------------------------------------------------------
function lockBadges(f) {
  const nl = (f.secs || []).filter(x => x.lock).length + (f.phr || []).length, ne = (f.secs || []).filter(x => x.edit).length + (f.el || []).length;
  return coChips(f.cos) + (nl ? `<span class="lk hard" style="margin:0" title="잠금 절 · 잠금 문장">🔒 ${nl}</span>` : '') + (ED.drafts[f.file] ? '<span class="pill warn">초안 있음</span>' : '');
}
// 회사 규칙 칩 (/hr-customize 9 「하네스에 반영」이 넣은 「회사 규칙 · 〈회사〉」 문단)
const coChips = (cos, st) => (cos || []).map(k => `<span class="lk co" style="${st || 'margin:0'}" title="${esc('회사 규칙 · ' + k + ' · /hr-customize 9 「하네스에 반영」이 넣은 문단 (기본 문장 · 잠금 칸은 그대로)')}">🏢 회사 규칙 · ${esc(k)}</span>`).join('');
const sameCo = (a, b) => String(a ?? '').toLowerCase().replace(/\.$/, '') === String(b ?? '').toLowerCase().replace(/\.$/, '');
function acts(f) {
  return `<div class="acts"><button class="btn sm" data-edit="${esc(f.file)}">수정</button><a class="btn ghost sm" href="${esc(f.vscode)}" title="VS Code 가 설치되어 있어야 열림 · ${esc(f.abs)}">원본 열기 (VS Code)</a></div>`;
}
function hcard(f, tags, wide) {
  return `<div class="hc${wide ? ' wide' : ''}${(f.cos || []).length ? ' hasco' : ''}"><div class="t"><b>${esc(f.name || f.file)}</b>${lockBadges(f)}<span class="tag">${f.lines}줄</span></div>
   ${f.desc ? `<p class="d" title="${esc(f.desc)}">${esc(stepNames(f.desc))}</p>` : ''}<div class="p">${esc(f.file)}</div>${tags || ''}
   ${f.body ? `<details><summary>본문 펼쳐 보기 <span class="muted">🔒 · 🏢 표시는 절마다</span></summary><div class="in">${f.body}</div></details>` : ''}${acts(f)}</div>`;
}
function ideaCard(x) {
  return `<div class="hc idea"><div class="t"><b>${x.lock ? '🔒 ' : ''}${esc(x.t)}</b>${x.ev ? `<span class="tag">${esc(x.ev)}</span><span class="tag">${esc(x.m)}</span>` : ''}</div><p class="d">${esc(x.d)}</p>
   <div class="p">${esc(x.target)}</div><div class="acts"><button class="btn ${x.lock ? 'quiet' : 'ghost'} sm" data-idea="${esc(x.id)}">요청 만들기</button></div></div>`;
}
const stepNames = t => String(t ?? '').replace(/([①②③④]) 단계/g, (m, k) => `${k} ${G.step[k] || ''} 단계`);   // 설명 속 「② 단계」 → 이름 붙임
const fcp = k => `<span class="cp" title="${esc(cpTip(k))}">${esc(cpLbl(k))}</span>`;
const hsec = (id, t, sub) => `<h2 class="hsech" id="hs-${id}">${esc(t)}${sub ? ` <span class="muted">${sub}</span>` : ''}</h2>`;
function renderHarness() {
  const h = D.harness, el = $('#p-harness');
  const ncos = names => { const L = []; names.forEach(n => { const f = n === 'CLAUDE.md' ? h.claude : [...h.agents, ...h.skills].find(x => x.name === n); ((f && f.cos) || []).forEach(k => { if (!L.includes(k)) L.push(k); }); }); return L; };
  const nd = (t, sub, names) => `<div class="node"><b>${t}</b>${sub}${(() => { const L = ncos(names); return L.length ? `<div>${coChips(L, ' ')}</div>` : ''; })()}</div>`;
  const flow = `<div class="flow leadrow">
   ${nd('리드 (메인 세션)', 'CLAUDE.md · wfp-lead · 순서 · CP · 결정로그 (계산하지 않음)', ['CLAUDE.md', 'wfp-lead'])}<span class="lbl">단계를 담당에게 맡김 · CP 에서 멈춤 ⇄</span>
   ${nd('검증 담당', 'wfp-verifier · wfp-verify · 단계마다 5기준 관문', ['wfp-verifier', 'wfp-verify'])}</div>
   <div class="flow">
   <div class="node"><b>data/〈회사〉</b>파일 6종 + 데이터 사전</div>→
   ${nd('① 사업전략 · 요구사항 분석', 'wfp-strategy', ['wfp-strategy', 'wfp-step1-strategy'])}${fcp('CP1')}→
   ${nd('② 인력수요 예측', 'wfp-demand · 방식 A 매출 · 이익 · B 인력동인 · C 업무량 · 대사', ['wfp-demand', 'wfp-step2-demand'])}${fcp('CP2')}→
   ${fcp('CP3')}${nd('③ 인력공급 예측', 'wfp-supply · 기본 · 상승', ['wfp-supply', 'wfp-step3-supply'])}→
   ${nd('④ 수급 Gap 분석 · 확보 · 육성', 'wfp-gap · 4B 대안 (Buy 채용 · Build 육성 · Borrow 외부 활용 · Bot 자동화)', ['wfp-gap', 'wfp-step4-gap'])}${fcp('CP4')}→
   ${nd('리포트', '리드 · wfp-report', ['wfp-report'])}${fcp('CP5')}</div>
   <details class="drawer"><summary>이 그림 읽는 법</summary><ul><li>체크포인트(CP) = 에이전트가 멈추고 사람 승인을 받는 자리</li><li>리드 = CLAUDE.md 를 읽은 메인 세션 (계산하지 않음 · 순서 · CP · 결정로그)</li><li>검증 담당 wfp-verifier = 단계마다 검증 5기준 (${esc(G.crit.join(' · '))}) 점검 · 통과해야 다음 단계 (관문)</li><li>자율성 L1 = 제안만 하고 확정은 사람이 함</li></ul></details>`;
  const tagRow = a => `<p style="margin:0"><span class="tag">Skill</span>${esc(a.skills || '-')} <span class="tag">도구</span>${esc(a.tools || '-')}</p>`;
  const scr = s => (s.script_info || []).length ? `<p style="margin:0">${s.script_info.map(x => `<a class="tag" href="${esc(x.vscode)}" title="VS Code 로 열기 · ${x.lines}줄${x.locked ? ' · CP 강제 코드 (잠금)' : ''}" style="text-decoration:none">${x.locked ? (x.lock_ok ? '🔒 ' : '⚠ ') : ''}${esc(x.name)}</a>`).join('')}${s.script_info.some(x => x.locked && !x.lock_ok) ? ' <span class="bad">CP 강제 코드가 빠짐 · customize.py lock 위반</span>' : ''}</p>` : '';
  const N = ED.basket.length, ideas = g => IDEAS.filter(x => x.grp === g).map(ideaCard).join('');
  const hooks = h.hooks || [], mcp = h.mcp || [], perms = h.perms || [], sets = h.settings || [];
  el.innerHTML = head('하네스', '보기 · 수정 초안 · Claude 에게 맡길 요청 (이 화면은 파일을 쓰지 않음)') +
   `<div class="bkbar"><b>요청 바구니 ${N}건</b><span class="muted" style="flex:1 1 220px">회사 규칙은 <code>/hr-customize</code> 9 (PRD 칸 답을 고친 뒤 다시) · 바구니는 그 밖의 수정${itip('수정 → 바구니에 담기 → 모두 복사 → Claude Code 대화창에 붙여 넣기 → Claude 가 /hr-customize 로 적용 (잠금 검사 · K-Co. 회귀 · 변경 이력 · 뷰어 갱신)')}</span><button class="btn ghost sm" data-bk="1">바구니 보기</button></div>
   <div class="hnav">${[['team', '팀 구성도'], ['claude', 'CLAUDE.md'], ['agents', `에이전트 ${h.agents.length}`], ['skills', `Skill ${h.skills.length}`], ['adv', `심화 · 훅 ${hooks.length} · MCP ${mcp.length} · 권한`]].map(([k, t]) => `<button data-jump="hs-${k}">${esc(t)}</button>`).join('')}</div>
   <div class="legend2"><span><span class="lk hard" style="margin:0">🔒 잠금</span> 바꾸지 않는 칸${itip('CP 5곳 · 업무 범위 · 금지 · 법전 규칙 · 승인 확인 · 검증 분리')}</span><span><span class="lk co" style="margin:0">🏢 회사 규칙</span> <code>/hr-customize</code> 9 「하네스에 반영」${itip('회사 답이 CLAUDE.md · 담당 · 스킬에 「회사 규칙 · 〈회사〉」 문단으로 들어간 자리 · 기본 문장 · 잠금 칸 · 계산 스크립트는 그대로 · 고치려면 PRD 칸 답을 고친 뒤 /hr-customize 9 를 다시')}</span><span class="muted">표시 없음 = 기본${itip('기본 팀 문장 · 바꿀 수 있으나 변경 이력 · 잠금 검사 · K-Co. 회귀 필수 · 잠금 기준: ' + h.lock_src)}</span></div>
   ${hsec('team', '팀 구성도')}<div class="card">${flow}</div>
   ${hsec('claude', 'CLAUDE.md', '리드가 먼저 읽는 파일')}<div class="hgrid">${h.claude && h.claude.exists ? hcard(h.claude, '', true) : '<div class="card muted">CLAUDE.md 가 없음 · /hr-init 먼저</div>'}</div>
   ${hsec('agents', `에이전트 ${h.agents.length}`, '.claude/agents/ · 누가')}<div class="hgrid">${h.agents.map(a => hcard(a, tagRow(a))).join('') || '<p class="nohit">없음</p>'}</div>
   ${hsec('skills', `Skill ${h.skills.length}`, '.claude/skills/ · 어떻게 · 스크립트 🔒 = CP 강제 코드')}<div class="hgrid">${h.skills.map(s => hcard(s, scr(s))).join('') || '<p class="nohit">없음</p>'}</div>
   <details class="drawer adv" id="hs-adv"><summary>심화 · 훅 ${hooks.length} · MCP ${mcp.length} · 권한 <span class="muted">(정해진 순간에 도는 명령 · 외부 연결 · 도구 사용 규칙 · 기본 팀은 두지 않음)</span></summary>
   ${hsec('hooks', `훅 ${hooks.length}`, '정해진 순간에 Claude Code 가 스스로 돌리는 명령 · 규칙을 장치로')}
   ${hooks.length ? `<div class="hgrid">${hooks.map(k => { const f = HF[k.src]; return `<div class="hc"><div class="t"><b>${esc(k.event)}</b><span class="tag">matcher ${esc(k.matcher)}</span><span class="tag">${esc(k.type || '-')}</span></div><div class="p">${esc(k.command)}</div><div class="p">${esc(k.src)}${k.timeout ? ' · 제한 ' + esc(k.timeout) + '초' : ''}</div>${f ? acts(f) : ''}</div>`; }).join('')}</div>`
     : `<div class="card"><p style="margin:0"><span class="pill gray">없음</span> <span class="muted">기본 팀은 훅을 두지 않는다. CP · 금지는 스크립트 (require_cp) 와 CLAUDE.md 규칙으로 지키고, 외부 전송 · 결과 파일 손 수정은 막지 않는다. 필요한 훅을 아래 후보에서 골라 요청으로 만든다 · ${esc(HOOK_NOTE)}</span></p></div>`}
   <h4>훅 후보</h4><div class="hgrid">${ideas('hook')}</div>
   ${hsec('mcp', `MCP ${mcp.length}`, '외부 도구 · 자료 연결 (.mcp.json · settings 의 mcpServers)')}
   ${mcp.length ? `<div class="hgrid">${mcp.map(x => { const f = HF[x.src]; return `<div class="hc"><div class="t"><b>${esc(x.name)}</b><span class="tag">${esc(x.type)}</span></div><div class="p">${esc(x.target || '-')}</div><div class="p">${esc(x.src)}</div><p class="d bad">CLAUDE.md 「업무 범위 · 금지」 (data/〈회사〉/ 밖 읽기 금지)와 맞는지 확인 필요</p>${f ? acts(f) : ''}</div>`; }).join('')}</div>`
     : `<div class="card"><p style="margin:0"><span class="pill gray">없음</span> <span class="muted">일부러 두지 않았다. CLAUDE.md 가 data/〈회사〉/ 폴더 밖 파일 읽기 · 외부 자료 검색을 잠그고 있어서, MCP 를 붙이려면 회사 정책으로 그 잠금을 바꾸는 결정이 먼저다 (요청하면 Claude 가 멈추고 묻는다)</span></p></div>`}
   <div class="hgrid">${ideas('mcp')}</div>
   ${hsec('perms', '권한', 'permissions · Claude Code 가 도구를 쓰기 전에 묻거나 막는 규칙')}
   ${perms.length ? `<div class="hgrid">${perms.map(p => { const f = HF[p.src], li = (k, a) => `<span>${k}</span><div>${a.length ? a.map(x => `<code>${esc(x)}</code>`).join(' ') : '<span class="muted">없음</span>'}</div>`;
       return `<div class="hc"><div class="t"><b>${esc(p.src)}</b>${p.mode ? `<span class="tag">기본 ${esc(p.mode)}</span>` : ''}</div><div class="kv">${li('허용 allow', p.allow)}${li('막음 deny', p.deny)}${li('묻기 ask', p.ask)}${p.dirs.length ? li('추가 폴더', p.dirs) : ''}</div>${f ? acts(f) : ''}</div>`; }).join('')}</div>`
     : `<div class="card"><p style="margin:0"><span class="pill gray">없음</span> <span class="muted">작업 폴더 권한 설정이 없다 · Claude Code 기본값대로 파일 쓰기 · 명령 실행 때마다 묻는다 (개인 설정 ~/.claude 는 여기 보이지 않음)</span></p></div>`}
   <div class="hgrid">${ideas('perm')}${sets.filter(f => !perms.some(p => p.src === f.file) && !hooks.some(k => k.src === f.file) && !mcp.some(m => m.src === f.file)).map(f => hcard(f, f.err ? `<p class="bad" style="margin:0">${esc(f.err)}</p>` : '')).join('')}</div></details>`;
}
function harnCard(c) {
  const r = c && c.harn; if (!r) return '';
  const pl = (lab, v) => v ? `<span class="pill ${/통과/.test(v) ? 'ok' : /실패|위반/.test(v) ? 'bad' : 'warn'}">${esc(lab)} ${esc(v)}</span>` : '';
  return `<div class="card"><div class="cohd"><h3 style="margin:0">회사 규칙 반영 요약</h3>${coChips([c.co])}<span class="sp" style="flex:1"></span><span class="jump">${dl('_하네스반영.md', '', '요약 문서 보기')}</span></div>
   <div class="corow">${r.status ? `<span class="pill ok">${esc(r.status)}</span>` : ''}${pl('잠금 검사', r.lock)}${pl('K-Co. 회귀', r.regress)}<span class="muted">층별 ${esc(r.counts.map(x => x.k + ' ' + x.n).join(' · '))} · 합 ${r.rows.length}곳</span>${itip('/hr-customize 9 「하네스에 반영」이 회사 답 (PRD 칸) 을 CLAUDE.md · 담당 · 스킬의 「회사 규칙 · ' + c.co + '」 문단으로 넣은 결과 · 기본 문장 · 잠금 칸 · 계산 스크립트는 그대로')}</div>
   <div class="mdbody"><table><thead><tr><th>층</th><th>파일</th><th>PRD 칸</th><th>왜</th></tr></thead><tbody>${r.rows.map(x => `<tr><td>${esc(x.grp)}</td><td><b>${esc(x.file)}</b></td><td>${esc(x.brief)}</td><td>${esc(x.why)}</td></tr>`).join('')}</tbody></table></div></div>`;
}
function renderChanges() {
  const h = D.harness, c = cur(), co = c ? c.co : '';
  const dcard = d => `<details class="diff"><summary>${esc(d.kind)} · ${esc(d.file)} <span class="add">+${d.plus}</span> <span class="del">−${d.minus}</span>${coChips(d.cos, 'margin:0 0 0 6px')}${(d.cos || []).length && d.other ? ' <span class="tag">기본 문장 변경도 있음</span>' : ''}</summary><div class="in"><pre>${esc(d.diff)}</pre></div></details>`;
  const mine = h.diffs.filter(d => !(d.cos || []).length || d.other || d.cos.some(k => sameCo(k, co)));
  const others = h.diffs.filter(d => !mine.includes(d)), ocos = [...new Set(others.flatMap(d => d.cos))];
  $('#p-changes').innerHTML = head('변경점', '기본 팀 대비 바뀐 파일 · CLAUDE.md 변경 이력') + harnCard(c) + `<h2>기본 대비 변경점${co ? ` <span class="muted" style="font-size:12.5px;font-weight:400">${esc(coName(co))} 기준 · 다른 회사 문단만 바뀐 파일은 아래에 접어 둠</span>` : ''}</h2>${!h.has_std ? empty('기본 사본(.claude/hr-lib/_standard)이 없음 · /hr-init 으로 설치한 폴더에서 보임', '') :
     h.diffs.length === 0 ? '<div class="card"><span class="pill ok">변경 없음</span> <span class="muted">기본 팀 그대로</span></div>' :
     (mine.length ? mine.map(dcard).join('') : `<div class="card"><span class="pill ok">${esc(coName(co))} 변경 없음</span> <span class="muted">이 회사의 회사 규칙 문단 · 기본 문장 변경이 없음</span></div>`) +
     (others.length ? `<details class="drawer"><summary>다른 회사의 회사 규칙 문단만 바뀐 파일 ${others.length}개 ${coChips(ocos, 'margin:0 0 0 4px')}</summary>${others.map(dcard).join('')}</details>` : '')}
   <h2>변경 이력 (CLAUDE.md)</h2><div class="card">${h.history || '<p class="muted">변경 이력 표 없음</p>'}</div>`;
}

function renderFlow() {
  const c = cur(), el = $('#p-flow'); if (!c) { el.innerHTML = head('진행 · 체크포인트') + empty('outputs/ 아래 회사 폴더가 아직 없음', 'K-Co 사업계획 분석부터 시작해 줘'); return; }
  el.innerHTML = head('진행 · 체크포인트', esc(coName(c.co)) + ' · 체크포인트(CP) = 에이전트가 멈추고 사람 승인을 받는 자리 · 답은 확정 · 추가 확인 · 보류 · 이관') + (c.decision ? decideCard(c, c.decision, true) : '<div class="card"><span class="pill ok">멈춘 곳 없음 · 승인 대기 없음</span></div>') + stepper(c) +
   `<h2>체크포인트(CP) 5곳</h2><div class="kpi">${c.cps.map(p => `<div title="${esc(cpTip(p.cp))}"><span><strong style="color:var(--ink);font-size:13.5px">${esc(p.cp)} · ${esc(p.name)}</strong>${itip(G.cp[p.cp] ? G.cp[p.cp].d : '')}</span><p style="margin:6px 0"><span class="pill ${pillCls(p.state)}"${p.re ? ` title="${esc('결정로그 「데이터 반영」 뒤 다시 승인 · ' + p.re)}"` : ''}>${esc(p.re ? '재승인 필요' : p.state)}</span>${p.re ? ` <span class="cpm" style="white-space:normal">${esc(p.re)}</span>` : ''}</p><span>승인자 ${esc(p.who)} ${esc(p.date || '')}</span>
     <span>${p.doc ? dl(p.doc, '', '근거 ' + p.file) : esc(p.file)}</span>${p.focus.length ? p.focus.map(f => `<span class="jump">${dl(p.doc, f.sec, '→ ' + f.label)}</span>`).join(' ') : ''}</div>`).join('')}</div>
   ${c.checks.length ? `<div class="card"><h3>담당이 짚어 둔 것<span class="sub">파일 안 「확인 필요」 · 「사람에게 물을 것」 같은 절</span></h3><ul>${c.checks.map(k => `<li>${dl(k.doc, k.sec, k.doc + ' › ' + k.label)}</li>`).join('')}</ul></div>` : ''}
   <h2>결정로그 <span class="pill gray" title="${esc(G.term['예외'])}">예외 (데이터 이상 멈춤) ${c.exceptions}건</span></h2><p class="muted">CP 답 · 예외 · 가정 변경 · 관문 인정을 한 행씩 남긴 기록</p><div class="card"><div class="mdbody">${docHtml(c, '결정로그.md') || '<p class="muted">없음</p>'}</div></div>`;
}
function docHtml(c, n) { return c.docs[n] ? c.docs[n].html : ''; }

function bars(rows, ty) {
  const W = 960, rowH = 24, pad = 230, H = rows.length * rowH + 40;
  const mx = Math.max(1, ...rows.map(r => Math.abs(r.g)));
  const x0 = pad + (W - pad - 60) * (mx / (2 * mx)), sc = (W - pad - 60) / (2 * mx);
  let s = `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="RPU별 Gap"><line x1="${x0}" x2="${x0}" y1="10" y2="${H - 24}" stroke="var(--muted)" stroke-width="1"/>`;
  rows.forEach((r, i) => {
    const y = 14 + i * rowH, w = Math.max(1, Math.abs(r.g) * sc), x = r.g >= 0 ? x0 : x0 - w, col = r.g >= 0 ? 'var(--short)' : 'var(--surplus)';
    s += `<text x="${pad - 8}" y="${y + 12}" text-anchor="end" font-size="12" fill="var(--ink-2)"><tspan font-weight="700">${esc(r.c)}</tspan> ${esc(r.n || '')}</text>
      <rect x="${x}" y="${y + 3}" width="${w}" height="${rowH - 8}" rx="3" fill="${col}" data-t="${esc(`<b>${r.c} ${r.n || ''}</b>${r.o ? ' · ' + r.o : ''}<br>Gap ${f1(r.g)} · 유형 ${r.t}<br>현원 ${r.hc}`)}"/>
      <text x="${r.g >= 0 ? x + w + 4 : x - 4}" y="${y + 12}" text-anchor="${r.g >= 0 ? 'start' : 'end'}" font-size="11" fill="var(--ink-2)">${f1(r.g)}</text>`;
  });
  return s + `<text x="${x0}" y="${H - 6}" text-anchor="middle" font-size="11" fill="var(--muted)">0 · 오른쪽 부족 · 왼쪽 여유</text></svg>`;
}

function niceStep(mx) { const raw = mx / 5, p = Math.pow(10, Math.floor(Math.log10(raw))), n = raw / p; return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10) * p; }
function dots(rows, types) {
  const W = 960, rowH = 24, pad = 230, top = 6, bot = top + rows.length * rowH, H = bot + 26;
  const vals = rows.flatMap(r => [r.b, r.cc, r.dec]).filter(v => v != null);
  const mx = Math.max(1, ...vals) * 1.05, sc = (W - pad - 30) / mx, st = niceStep(mx), near = (a, b) => a != null && b != null && Math.abs(a - b) < 0.05;
  let s = `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="RPU별 방식 B 인력동인 · C 업무량 · 결정 수요">`;
  rows.forEach((r, i) => { if (i % 2 === 0) s += `<rect x="0" y="${top + i * rowH}" width="${W}" height="${rowH}" fill="var(--zebra)"/>`; });
  for (let v = 0; v <= mx + 1e-9; v += st) { const x = pad + v * sc;
    s += `<line x1="${x}" x2="${x}" y1="${top}" y2="${bot}" stroke="var(--line)" stroke-width="1"${v ? ' stroke-dasharray="2 3"' : ''}/><text x="${x}" y="${bot + 16}" text-anchor="middle" font-size="11" fill="var(--muted)">${Math.round(v).toLocaleString('ko-KR')}</text>`; }
  s += `<text x="${pad - 18}" y="${bot + 16}" text-anchor="end" font-size="11" fill="var(--muted)">명 · FTE</text>`;
  rows.forEach((r, i) => {
    const y = top + i * rowH + rowH / 2, xb = r.b != null ? pad + r.b * sc : null, xc = r.cc != null ? pad + r.cc * sc : null, xd = pad + (r.dec || 0) * sc;
    const pick = near(r.dec, r.b) && near(r.dec, r.cc) ? '방식 B · C 값과 같음' : near(r.dec, r.b) ? '방식 B 값을 고름' : near(r.dec, r.cc) ? '방식 C 값을 고름' : '대사로 조정한 값', ty = (types || {})[r.c];
    const t = esc(`<b>${r.c} ${r.n || ''}</b><br>B 인력동인 ${f1(r.b)} · C 업무량 ${f1(r.cc)}<br>결정 수요 (사람) ${f1(r.dec)} · ${pick}<br>원인 판정 ${r.cause || '-'}${ty ? '<br>Gap 유형 ' + ty : ''}`);
    s += `<text x="${pad - 8}" y="${y + 4}" text-anchor="end" font-size="12" fill="var(--ink-2)"><tspan font-weight="700">${esc(r.c)}</tspan> ${esc(r.n || '')}</text>`;
    s += `<rect x="0" y="${y - rowH / 2}" width="${W}" height="${rowH}" fill="transparent" data-t="${t}"/>`;
    if (xb != null && xc != null) s += `<line x1="${xb}" x2="${xc}" y1="${y}" y2="${y}" stroke="var(--muted)" stroke-opacity=".45" stroke-width="2"/>`;
    if (xb != null) s += `<circle cx="${xb}" cy="${y}" r="5" fill="var(--s1)" stroke="var(--surface)" stroke-width="2" data-t="${t}"/>`;
    if (xc != null) s += `<circle cx="${xc}" cy="${y}" r="5" fill="var(--s2)" stroke="var(--surface)" stroke-width="2" data-t="${t}"/>`;
    if (r.dec != null && (near(r.dec, r.b) || near(r.dec, r.cc))) s += `<circle cx="${xd}" cy="${y}" r="10" fill="none" stroke="var(--accent)" stroke-width="1.5" data-t="${t}"/>`;
    s += `<path d="M${xd} ${y - 7} L${xd + 5} ${y} L${xd} ${y + 7} L${xd - 5} ${y} Z" fill="none" stroke="var(--ink)" stroke-width="2" data-t="${t}"/>`;
  });
  return s + '</svg>';
}

function renderDash() {
  const c = cur(), el = $('#p-dash'); if (!c) { el.innerHTML = head('대시보드') + empty('결과가 아직 없음', 'K-Co 사업계획 분석부터 시작해 줘'); return; }
  const d = c.dash, ty = d.ty;
  const nm = d.names || {}, og = d.orgs || {};
  const g = (d.gap || []).map(x => ({ ...x, n: nm[x.c] || '', o: og[x.c] || '' })).sort((a, b) => b.g - a.g);
  const r = (d.recon || []).filter(x => x.b != null || x.cc != null).map(x => ({ ...x, n: nm[x.c] || '' }));
  el.innerHTML = head('대시보드', `${esc(coName(c.co))} · 계획연도 ${esc(ty)}`) + kpis(c) + `
   ${g.length ? `<div class="card"><h3>RPU별 Gap (결정 수요 사람 − 공급 기본)<span class="sub">RPU = 인력 산정 단위 · 유형 = Gap 4유형 (${esc(G.gap4.join(' · '))})</span></h3><div class="legend"><span><i style="background:var(--short)"></i>부족 (수요가 공급보다 많음)</span><span><i style="background:var(--surplus)"></i>여유 (공급이 수요보다 많음)</span></div><div class="chart">${bars(g, ty)}</div>
     <details><summary>표로 보기</summary><div class="in"><div class="tw"><table><thead><tr><th>RPU 코드</th><th>이름</th><th>본부</th><th class="n">Gap</th><th>Gap 유형</th><th class="n">현원</th></tr></thead><tbody>${g.map(x => `<tr><td>${esc(x.c)}</td><td>${esc(x.n)}</td><td>${esc(x.o)}</td><td class="n">${f1(x.g)}</td><td>${esc(x.t)}</td><td class="n">${esc(x.hc)}</td></tr>`).join('')}</tbody></table></div></div></details></div>` : empty('④ 수급 Gap 분석 결과가 아직 없음', r.length ? '④ Gap 분석 진행해 줘' : '② 수요 예측 진행해 줘')}
   ${r.length ? `<div class="card"><h3>수요 대사 · 방식 B 인력동인 · C 업무량 · 결정 수요 (사람)<span class="sub">대사 = 방식 B · C 값을 맞대어 차이 원인을 판정하는 것 · 평균 내지 않음</span></h3><div class="legend"><span><i style="background:var(--s1);border-radius:50%"></i>B 인력동인</span><span><i style="background:var(--s2);border-radius:50%"></i>C 업무량 (사람 FTE)</span><span>◇ 결정 수요 (사람) · 대사로 정한 값</span><span><i style="border:1.5px solid var(--accent);background:none;border-radius:50%"></i>◇ 가 B 또는 C 와 같으면 고른 방식</span><span class="muted">줄에 마우스를 올리면 값 · 유형</span></div><div class="chart">${dots(r, Object.fromEntries((d.gap || []).map(x => [x.c, x.t])))}</div>
     <details><summary>표로 보기</summary><div class="in"><div class="tw"><table><thead><tr><th>RPU 코드</th><th>이름</th><th class="n" title="${esc('수요 방식 B · ' + G.method.B)}">B 인력동인</th><th class="n" title="${esc('수요 방식 C · ' + G.method.C)}">C 업무량</th><th class="n">결정 수요 (사람)</th><th>원인 판정</th></tr></thead><tbody>${r.map(x => `<tr><td>${esc(x.c)}</td><td>${esc(x.n)}</td><td class="n">${f1(x.b)}</td><td class="n">${f1(x.cc)}</td><td class="n">${f1(x.dec)}</td><td>${esc(x.cause)}</td></tr>`).join('')}</tbody></table></div></div></details></div>` : ''}
   ${d.error ? `<p class="bad">일부 결과를 읽지 못함: ${esc(d.error)}</p>` : ''}`;
  $$('[data-t]', el).forEach(n => { n.addEventListener('mousemove', e => showTip(e, n.dataset.t.replace(/&lt;/g, '<').replace(/&gt;/g, '>'))); n.addEventListener('mouseleave', hideTip); });
}

function renderOut() {
  const c = cur(), el = $('#p-out'); if (!c) { el.innerHTML = head('산출물') + empty('결과가 아직 없음', 'K-Co 사업계획 분석부터 시작해 줘'); return; }
  // 표 + 체크 + 「선택한 파일 받기」 (2026-11-16) · 파일은 뷰어를 다시 만든 시점의 사본 (HTML 안에 들어 있음 · 고쳐도 계산에 반영 안 됨)
  const F = c.files || {}, kind = n => /\.xlsx$/.test(n) ? '엑셀' : /\.docx$/.test(n) ? '워드' : '문서';
  const rows = [];
  c.stages.forEach(s => s.items.forEach(it => rows.push({st: s.st, ...it})));
  c.extra.forEach(n => rows.push({st: '기록', name: n, doc: n, exists: true, mtime: c.docs[n].mtime, desc: (c.descs || {})[n] || ''}));
  OUTROWS = rows;
  const tr = (r, i) => {
    const fs = F[r.name] || [], main = fs.find(f => f.label !== 'md 원본');
    const fmt = !r.exists ? '' : kind(r.name) === '문서' ? (main ? (/\.xlsx$/.test(main.fn) ? '엑셀로' : '워드로') + ' · md' : 'md') : kind(r.name);
    return `<tr class="${r.exists ? '' : 'off'}" data-i="${i}"><td class="oc"><input type="checkbox" ${r.exists && fs.length ? '' : 'disabled'} aria-label="${esc(r.name)} 선택"></td>` +
      `<td class="stc"><span class="stb">${esc(r.st)}</span></td>` +
      `<td><button class="fname" type="button" ${r.exists ? '' : 'disabled'}>${esc(r.name)}</button>${r.desc ? `<div class="fdesc">${esc(r.desc)}</div>` : ''}</td>` +
      `<td class="fk">${r.exists ? esc(fmt) : '<span class="muted">아직 없음</span>'}</td><td class="fm">${esc(r.mtime || '')}</td>` +
      `<td class="fv">${r.exists ? (r.doc ? dl(r.doc, '', '문서') : '') : ''}</td></tr>`;
  };
  el.innerHTML = head('산출물', `outputs/${esc(c.co)}/ · 이 시점 사본 (뷰어를 다시 만들 때마다 새로 담김)`) +
    `<div class="card otab"><div class="obar"><label class="ckall"><input type="checkbox" id="o-all"> 전체 선택</label><span class="muted" id="o-n">선택 0개</span>` +
    `<label class="omd"><input type="checkbox" id="o-md"> md 원본도 함께</label><span class="sp"></span><button class="btn" id="o-dl" disabled>${ic('download')}선택한 파일 받기</button></div>` +
    `<div class="tw"><table class="otable"><thead><tr><th class="oc"></th><th>단계</th><th>파일 · 무엇</th><th>받는 형식</th><th>만든 시각</th><th></th></tr></thead><tbody>` +
    rows.map(tr).join('') + `</tbody></table></div>` +
    `<p class="muted onote">파일 이름을 누르면 내용을 바로 봄 · 받은 파일은 이 시점 사본이라 고쳐도 계산에 반영되지 않음 (바꾸려면 Claude 에게 말하거나 CP 답으로) · 여러 개를 고르면 zip 하나로 받음</p></div>`;
  const tb = $('.otable tbody', el), sel = () => $$('tbody input[type=checkbox]:checked', el).map(x => +x.closest('tr').dataset.i);
  const upd = () => { const n = sel().length; $('#o-n', el).textContent = `선택 ${n}개`; $('#o-dl', el).disabled = !n;
    const all = $$('tbody input[type=checkbox]:not([disabled])', el); $('#o-all', el).checked = all.length && all.every(x => x.checked); };
  $('#o-all', el).onchange = e => { $$('tbody input[type=checkbox]:not([disabled])', el).forEach(x => x.checked = e.target.checked); upd(); };
  tb.addEventListener('change', upd);
  tb.addEventListener('click', e => { const b = e.target.closest('button.fname'); if (!b) return;
    const row = b.closest('tr'), nx = row.nextElementSibling;
    if (nx && nx.classList.contains('pv')) { nx.remove(); row.classList.remove('open'); return; }
    const r = OUTROWS[+row.dataset.i];
    const body = r.doc ? `<div class="mdbody">${docHtml(c, r.doc)}</div>` : r.html || (/\.docx$/.test(r.name) ? '<p class="muted">워드 리포트 (표지 · 경영진 요약 · 본문 · 승인) · 내용은 「리포트」 화면과 같음 · CP5 승인 전에는 「초안」 표시</p>' : '');
    row.classList.add('open'); row.insertAdjacentHTML('afterend', `<tr class="pv"><td colspan="6"><div class="in">${body}</div></td></tr>`); post(); });
  $('#o-dl', el).onclick = () => {
    const md = $('#o-md', el).checked, out = [];
    sel().forEach(i => (F[OUTROWS[i].name] || []).forEach(f => { if (f.label !== 'md 원본' || md || (F[OUTROWS[i].name] || []).length === 1) out.push(f); }));
    if (out.length === 1) { const a = document.createElement('a'); a.href = out[0].href; a.download = out[0].fn; a.click(); return; }
    const blob = zipBlob(out.map(f => ({name: f.fn, data: b64bytes(f.href)})));
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `${c.co}_결과_${new Date().toISOString().slice(0, 16).replace(/[^0-9]/g, '')}.zip`; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  };
}
let OUTROWS = [];
const b64bytes = href => { const s = atob(href.slice(href.indexOf(',') + 1)), u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; };
const CRCT = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = u => { let c = 0xFFFFFFFF; for (let i = 0; i < u.length; i++) c = CRCT[(c ^ u[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
function zipBlob(files) {   // 저장 방식 (압축 없음) zip · 파일 이름 UTF-8 (플래그 0x0800)
  const enc = new TextEncoder(), parts = [], cen = []; let off = 0;
  const d = new Date(), dt = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(), tm = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
  files.forEach(f => { const nm = enc.encode(f.name), crc = crc32(f.data), sz = f.data.length;
    const h = new DataView(new ArrayBuffer(30)); h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
    h.setUint16(10, tm, true); h.setUint16(12, dt, true); h.setUint32(14, crc, true); h.setUint32(18, sz, true); h.setUint32(22, sz, true); h.setUint16(26, nm.length, true); h.setUint16(28, 0, true);
    parts.push(h.buffer, nm, f.data);
    const cd = new DataView(new ArrayBuffer(46)); cd.setUint32(0, 0x02014b50, true); cd.setUint16(4, 20, true); cd.setUint16(6, 20, true); cd.setUint16(8, 0x0800, true); cd.setUint16(10, 0, true);
    cd.setUint16(12, tm, true); cd.setUint16(14, dt, true); cd.setUint32(16, crc, true); cd.setUint32(20, sz, true); cd.setUint32(24, sz, true); cd.setUint16(28, nm.length, true);
    cd.setUint16(30, 0, true); cd.setUint16(32, 0, true); cd.setUint16(34, 0, true); cd.setUint16(36, 0, true); cd.setUint32(38, 0, true); cd.setUint32(42, off, true);
    cen.push(cd.buffer, nm); off += 30 + nm.length + sz; });
  const csz = cen.reduce((a, x) => a + (x.byteLength || x.length), 0), e = new DataView(new ArrayBuffer(22));
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, csz, true); e.setUint32(16, off, true);
  return new Blob([...parts, ...cen, e.buffer], {type: 'application/zip'});
}
function renderDoc() {
  const c = cur(), el = $('#p-doc'); if (!c) { el.innerHTML = head('문서') + empty('결과가 아직 없음', 'K-Co 사업계획 분석부터 시작해 줘'); return; }
  const names = Object.keys(c.docs);
  if (!names.length) { el.innerHTML = head('문서') + empty('outputs/' + c.co + '/ 에 .md 파일이 아직 없음', 'K-Co 사업계획 분석부터 시작해 줘'); return; }
  if (!c.docs[DOC]) DOC = (c.decision && c.decision.evidence[0] && c.decision.evidence[0].links[0] || {}).doc || names[0];
  const d = c.docs[DOC], dec = c.decision ? new Set(c.decision.evidence.flatMap(e => e.links.map(l => l.doc))) : new Set();
  const inStage = new Set(c.stages.flatMap(s => s.items.map(i => i.doc).filter(Boolean)));
  const group = (t, list) => list.length ? `<h4>${t}</h4>` + list.map(n => `<a href="#" class="dl${n === DOC ? ' on' : ''}" data-doc="${esc(n)}">${dec.has(n) ? '⏸ ' : ''}${esc(n)}${c.docs[n].blank ? ' <span class="muted">(양식)</span>' : ''}</a>`).join('') : '';
  el.innerHTML = `<div class="docwrap"><aside class="doclist">${group('단계 산출물', names.filter(n => inStage.has(n)))}${group('자사 버전 설계서', names.filter(n => n === DSN && !inStage.has(n)))}${group('기록 · 멈춤', names.filter(n => !inStage.has(n) && n !== DSN))}</aside>
   <article class="card doc"><div class="dochd"><b>outputs/${esc(c.co)}/${esc(DOC)}</b> <span class="muted">${esc(d.mtime)}</span><span class="dact">${getBtns(c, [DOC])}</span>
     ${d.focus.length ? `<div>${d.focus.map(f => `<span class="jump">${dl(DOC, f.sec, '→ ' + f.label)}</span>`).join(' ')}</div>` : ''}
     ${d.heads.length > 2 ? `<details class="toc"><summary>목차 ${d.heads.length}</summary><div class="in">${d.heads.map(h => dl(DOC, h, h)).join('')}</div></details>` : ''}</div>
     <div class="mdbody">${d.html}</div></article></div>`;
  if (TAB !== 'doc') return;
  if (SEC) { const want = SEC; SEC = '';
    requestAnimationFrame(() => { const hs = $$('.mdbody [data-h]', el); const h = hs.find(x => x.dataset.h === want) || hs.find(x => x.dataset.h.startsWith(want)) || hs.find(x => x.dataset.h.includes(want));
      if (h) { const sec = h.closest('.msec'), nav = sec && sec.parentElement.previousElementSibling; if (nav && nav.classList.contains('tabbed')) showSec(nav, sec.dataset.k); h.scrollIntoView({block: 'start'}); h.classList.add('hit'); setTimeout(() => h.classList.remove('hit'), 2600); } else scrollTo(0, 0); }); }
  else if (!RESTORED) requestAnimationFrame(() => scrollTo(0, 0));
}
function renderReport() {
  const c = cur(), el = $('#p-report'); if (RIO) { RIO.disconnect(); RIO = null; }
  el.innerHTML = head('리포트', c && c.report ? `${dl(c.report, '', c.report)} · <abbr class="g" title="${esc(cpTip('CP5'))}">CP5 · 리포트 공유</abbr> 승인 전에는 밖으로 보내지 않음` : '', c && c.report ? getBtns(c, [c.report.replace(/\.md$/, '.docx'), c.report]) : '') +
    (c && c.report ? `<div class="withtoc"><div class="card"><div class="mdbody" id="rep">${docHtml(c, c.report)}</div></div><nav class="rtoc" id="rtoc" aria-label="리포트 목차"></nav></div>` : empty('리포트가 아직 없음 (CP4 · Gap 대안 승인 뒤 wfp-report 가 7섹션으로 만듦)', '보고서로 정리해 줘'));
  if (!(c && c.report)) return;
  const rep = $('#rep', el), nav = $('#rtoc', el); declutter(rep);
  const hs = heads(rep);
  nav.innerHTML = `<b>목차</b>` + hs.map(h => `<a href="#${hid(h)}" data-go="${h.id}" class="l${h.tagName[1]}">${esc(htext(h))}</a>`).join('');
  if (TAB !== 'report' || !hs.length || !('IntersectionObserver' in window)) return;
  const mark = id => $$('a', nav).forEach(a => a.classList.toggle('on', a.dataset.go === id)); mark(hs[0].id);
  RIO = new IntersectionObserver(es => es.forEach(e => { if (e.isIntersecting) mark(e.target.id); }), {rootMargin: `-${(parseInt(getComputedStyle(document.documentElement).getPropertyValue('--stick')) || 56) + 20}px 0px -65% 0px`});
  hs.forEach(h => RIO.observe(h));
}
// 검증: 찾기 · 상태 · 단계로 행 거르기 (기본 = 미해결이 있으면 미해결)
let VF = {co: null, s: '', g: '', q: ''};
const VORD = ['미해결', '미확인', '인정', '통과', '해결'];
function renderVerify() {
  const c = cur(), el = $('#p-verify');
  if (!c) { el.innerHTML = head('검증', '관문 = 검증 담당의 5기준 점검 · 미해결이 남으면 다음 단계로 넘어가지 않음') + empty('결과가 아직 없음', ''); return; }
  const vk = (k, lbl) => ` class="vk" role="button" tabindex="0" data-vf="${k}" title="눌러서 ${esc(lbl)} 행만 보기"`;
  el.innerHTML = head('검증', '관문 = 검증 담당의 5기준 점검 · 미해결이 남으면 다음 단계로 넘어가지 않음') + `<div class="card" style="padding:14px 20px"><b>검증 5기준</b><div class="crit">${G.crit.map((x, i) => `<span class="pill gray">${i + 1}. ${esc(x)}</span>`).join('')}</div><p class="muted" style="margin:8px 0 0">단계 칸 1 · 2 · 3 · 4 = ${Object.keys(G.step).map(k => esc(k + ' ' + G.step[k])).join(' · ')}</p></div>
    <div class="kpi"><div${vk('', '전체')}><span>검증 행</span><b>${c.v_all}</b><span class="cpm">단계마다 원천에서 따로 계산해 대조한 행</span></div><div${vk('미해결', '미해결')}><span>미해결</span><b class="${c.v_open ? 'st-기타' : 'st-승인'}">${c.v_open}</b><span class="cpm">넘긴 담당이 고친 뒤 다시 검증</span></div><div${vk('미확인', '미확인')}><span>미확인 (행동 테스트 대기)</span><b class="${c.v_wait ? 'st-대기' : 'st-승인'}">${c.v_wait}</b><span class="cpm">${esc(G.term['미확인'])}</span></div><div${vk('인정', '인정')}><span>인정 (관문 인정 · 결정로그)</span><b>${c.v_acc || 0}</b><span class="cpm">${esc(G.term['인정'])}</span></div></div>
    <div class="vbar" id="vbar" hidden></div><div class="card vlog"><div class="mdbody" id="vlog">${docHtml(c, '검증로그.md') || '<p class="muted">없음</p>'}</div><p class="nohit" id="vnone" hidden>조건에 맞는 검증 행이 없음 · 「전체」를 누르거나 찾는 낱말을 지우세요</p></div>`;
  const log = $('#vlog', el); declutter(log);
  const rows = [];
  $$('table', log).forEach(t => { const hs = $$('thead th', t).map(txt), si = hs.indexOf('상태'), gi = hs.indexOf('단계'); if (si < 0) return;
    $$('tbody tr', t).forEach(tr => { const td = [...tr.children]; if (!td[si] || td.slice(1).every(x => !txt(x))) return;
      const s0 = htext(td[si]); rows.push({tr, s: s0.startsWith('인정') ? '인정' : s0, g: gi >= 0 ? htext(td[gi]) : '', t: (txt(tr) + ' ' + $$('[data-tip]', tr).map(x => x.dataset.tip).join(' ')).toLowerCase()}); }); });
  if (!rows.length) return;
  const cnt = k => rows.filter(r => r.s === k).length, sts = [...VORD.filter(k => cnt(k)), ...[...new Set(rows.map(r => r.s))].filter(k => k && !VORD.includes(k))];
  const go_ = g => { const i = ['1', '2', '3', '4', 'report'].indexOf(g); return i < 0 ? 9 : i; };
  const stages = [...new Set(rows.map(r => r.g).filter(Boolean))].sort((a, b) => go_(a) - go_(b) || a.localeCompare(b));
  if (VF.co !== c.co) VF = {co: c.co, s: cnt('미해결') ? '미해결' : '', g: '', q: ''};
  if (VF.s && !cnt(VF.s)) VF.s = ''; if (VF.g && !stages.includes(VF.g)) VF.g = '';
  const gName = g => (g === 'all' || g === 'report') ? (G.step_short[g] || g) : G.step_short[g] ? `${g} ${G.step_short[g]}` : g;
  const bar = $('#vbar', el); bar.hidden = false;
  bar.innerHTML = `<input type="search" id="vq" placeholder="검증 행 찾기 (예: 재계산 · 리포트 · 파일 이름)" aria-label="검증 행 찾기" value="${esc(VF.q)}">
    <div class="chips" role="group" aria-label="상태로 거르기">${[['', '전체', rows.length], ...sts.map(k => [k, k, cnt(k)])].map(([k, n, v]) => `<button class="chip" data-vs="${esc(k)}" title="${esc(G.term[k] || '')}">${esc(n)}<em>${v}</em></button>`).join('')}</div>
    ${stages.length > 1 ? `<div class="chips" role="group" aria-label="단계로 거르기"><span class="lbl">단계</span>${['', ...stages].map(g => `<button class="chip" data-vg="${esc(g)}">${esc(g ? gName(g) : '전체')}</button>`).join('')}</div>` : ''}`;
  const apply = () => { let n = 0;
    rows.forEach(r => { const on = (!VF.s || r.s === VF.s) && (!VF.g || r.g === VF.g) && (!VF.q || r.t.includes(VF.q)); r.tr.classList.toggle('vf-hide', !on); if (on) n++; });
    $$('[data-vs]', bar).forEach(b => { const on = b.dataset.vs === VF.s; b.classList.toggle('on', on); b.setAttribute('aria-pressed', on); });
    $$('[data-vg]', bar).forEach(b => { const on = b.dataset.vg === VF.g; b.classList.toggle('on', on); b.setAttribute('aria-pressed', on); });
    $$('[data-vf]', el).forEach(k => k.classList.toggle('on', k.dataset.vf === VF.s));
    $('#vnone', el).hidden = n > 0; fitVlog(); };
  bar.addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; if ('vs' in b.dataset) VF.s = b.dataset.vs; if ('vg' in b.dataset) VF.g = b.dataset.vg; apply(); });
  $('#vq', el).addEventListener('input', e => { VF.q = e.target.value.trim().toLowerCase(); apply(); });
  const pick = k => { VF.s = !k || cnt(k) ? k : ''; apply(); bar.scrollIntoView({block: 'nearest', behavior: 'smooth'}); };
  $$('[data-vf]', el).forEach(k => { k.addEventListener('click', () => pick(k.dataset.vf)); k.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick(k.dataset.vf); } }); });
  apply();
}
function renderRegress() { const n = '_회귀테스트.md', c0 = cur(), c = (c0 && c0.docs[n]) ? c0 : (D.companies.find(x => x.docs[n] && /^k-co\.?$/i.test(x.co)) || c0); $('#p-regress').innerHTML = head('회귀', 'K-Co. 기준 결과와 비교 · 하네스를 고친 뒤에도 K-Co. 숫자가 그대로인지') + (c && c.docs[n] ? `<div class="card"><div class="mdbody">${docHtml(c, n)}</div></div>` : empty('회귀 테스트 결과가 아직 없음', 'K-Co 로 회귀 테스트 해 줘')); }

// 설계서: 회사 적용 도우미 (/hr-customize · 인자 없이) 가 남기는 outputs/<회사>/자사버전_설계서.md (묻고 답한 것 · 바꾼 것 · 잠금 · 회귀)
const DSN = '자사버전_설계서.md';
function renderDesign() {
  const c = cur(), el = $('#p-design'), d = c && c.docs[DSN];
  const sub = `${c ? 'outputs/' + esc(c.co) + '/' : 'outputs/〈회사〉/'}${esc(DSN)} · <code>/hr-customize</code> 가 남기는 회사별 자산 · 묻고 답한 것 · 바꾼 것 · 잠금 · 회귀`;
  const refs = (D.refs || []).length ? `<p class="muted" style="margin:12px 2px 0;font-size:12.5px">참고 문서 <span class="pill gray">읽기 전용</span> ${D.refs.map(r => `${esc(r.t)} <code>${esc(r.path)}</code>${r.exists ? '' : ' <span class="warn">(없음)</span>'}`).join(' · ')}</p>` : '';
  if (!d) { el.innerHTML = head('설계서', sub) + empty(`${c ? coName(c.co) + ' · ' : ''}/hr-customize 를 끝내면 설계서가 여기 보입니다${c && /^k-co\.?$/i.test(c.co) ? ' (K-Co. 는 교육 예제 회사라 보통 설계서가 없음)' : ''}`, '/hr-customize') + refs; return; }
  el.innerHTML = head('설계서', sub) + `<article class="card doc"><div class="dochd"><b>outputs/${esc(c.co)}/${esc(DSN)}</b> <span class="muted">${esc(d.mtime)}</span> <span class="jump">${dl(DSN, '', '문서 화면에서 보기')}</span>
     ${d.heads.length > 2 ? `<details class="toc"><summary>목차 ${d.heads.length}</summary><div class="in">${d.heads.map(h => dl(DSN, h, h)).join('')}</div></details>` : ''}</div>
     <div class="mdbody">${d.html}</div></article>` + refs;
}
function renderVirtual() {
  const c = cur(), v = c && c.virtual;   // 가상 데이터 배너 (모든 화면 위쪽)
  $('#vb').innerHTML = v ? `<div class="vban"><b>${esc(D.virtual_line)}</b> · ${esc(c.co)} · ${v.marker ? `표시 파일 data/${esc(c.co)}/${esc(D.virtual_mark)}` : `<span class="bad">표시 파일 ${esc(D.virtual_mark)} 이 빠짐 · 검증 관문</span>`}</div>` : '';
}
function render() {
  $$('section.pane').forEach(p => p.classList.toggle('on', p.id === 'p-' + TAB));
  $$('#co button').forEach(b => b.classList.toggle('on', b.dataset.co === CO));
  document.body.classList.remove('nav-open');
  renderNav(); renderVirtual(); renderDecide(); renderHome(); renderHarness(); renderChanges(); renderDesign(); renderFlow(); renderDash(); renderOut(); renderDoc(); renderReport(); renderVerify(); renderRegress(); post(); save();
}
function post() {   // 렌더 뒤: 지금 화면의 md 문서에 간단히 보기 · RPU 이름 · 절 바로가기 · 보기 단추 · 고정 줄 높이
  const pane = $('#p-' + TAB);
  if (pane) { $$('.mdbody', pane).forEach(declutter); if (TAB === 'doc' || TAB === 'out' || TAB === 'report') miniToc(pane); sheetTabs(pane); nameRpu(pane); }
  nameRpu($('#dc'));
  $$('#detail button').forEach(b => { const on = (b.dataset.mode === 'detail') === DETAIL; b.classList.toggle('on', on); b.setAttribute('aria-pressed', on); });
  setStick(); fitVlog();
}
document.addEventListener('click', e => { const t = e.target.closest('[data-tab]'); if (t) { e.preventDefault(); TAB = t.dataset.tab; render(); scrollTo(0, 0); return; }
  const a = e.target.closest('a.dl'); if (!a) return; e.preventDefault(); go(a.dataset.doc, a.dataset.sec); });
// 풀이 말풍선: 마우스는 title 로 보이고, 손가락으로 누르면 잠깐 띄움 (휴대폰)
let GT = 0;
document.addEventListener('click', e => { const g = e.target.closest('abbr.g'); if (!g || e.target.closest('a')) return; const r = g.getBoundingClientRect();
  showTip({clientX: Math.max(4, r.left - 14), clientY: r.bottom - 8}, esc(g.title)); clearTimeout(GT); GT = setTimeout(hideTip, 3000); });

// ⓘ 말풍선 (마우스 · 키보드 · 손가락) · 같은 #tip 을 씀 · 긴 칸 누르면 펼침 · 절 바로가기
function tipAt(n) { tip.style.whiteSpace = 'pre-line'; tip.textContent = n.dataset.tip; tip.style.display = 'block'; const r = n.getBoundingClientRect(), tw = tip.offsetWidth, th = tip.offsetHeight;
  tip.style.left = Math.max(8, Math.min(r.left, innerWidth - tw - 8)) + 'px'; tip.style.top = (r.bottom + th + 12 > innerHeight ? Math.max(8, r.top - th - 6) : r.bottom + 6) + 'px'; }
const tipOf = e => e.target && e.target.closest ? e.target.closest('[data-tip]') : null;
document.addEventListener('mouseover', e => { const n = tipOf(e); if (n) tipAt(n); });
document.addEventListener('mouseout', e => { if (tipOf(e)) hideTip(); });
document.addEventListener('focusin', e => { const n = tipOf(e); if (n) tipAt(n); });
document.addEventListener('focusout', e => { if (tipOf(e)) hideTip(); });
document.addEventListener('click', e => { const n = tipOf(e); if (n) { tipAt(n); clearTimeout(GT); GT = setTimeout(hideTip, 4000); return; }
  const c = e.target.closest('.clamp'); if (c && !e.target.closest('a,abbr,button') && !String(getSelection())) c.classList.toggle('open'); });
document.addEventListener('click', e => { const a = e.target.closest('[data-go]'); if (!a) return; e.preventDefault(); const t = document.getElementById(a.dataset.go); if (!t) return;
  for (let d = t.closest('details'); d; d = d.parentElement && d.parentElement.closest('details')) d.open = true; t.scrollIntoView({behavior: 'smooth', block: 'start'}); });
addEventListener('scroll', hideTip, {passive: true});
addEventListener('resize', () => { setStick(); fitVlog(); });
function fitVlog() {   // 검증 표: 넓은 화면에서는 머리 줄이 화면 위에 붙게 (넘치면 표 안 가로 스크롤로 되돌림 · 페이지는 넘치지 않음)
  const tw = $('#vlog .tw'); if (!tw || !tw.offsetParent) return; tw.classList.remove('wide'); if (tw.scrollWidth > tw.clientWidth + 1) tw.classList.add('wide'); }
let PRINT_OPEN = [];
addEventListener('beforeprint', () => { PRINT_OPEN = $$('section.pane.on details:not([open])'); PRINT_OPEN.forEach(d => { d.open = true; }); });
addEventListener('afterprint', () => { PRINT_OPEN.forEach(d => { d.open = false; }); PRINT_OPEN = []; });
const TT = $('#totop'); TT.innerHTML = '<svg class="i" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 19V5M5 12l7-7 7 7"/></svg>';
addEventListener('scroll', () => TT.classList.toggle('on', scrollY > 600), {passive: true}); TT.addEventListener('click', () => scrollTo({top: 0, behavior: 'smooth'}));
$('#detail').addEventListener('click', e => { const b = e.target.closest('[data-mode]'); if (!b) return; const v = b.dataset.mode === 'detail'; if (v === DETAIL) return; DETAIL = v; const y = scrollY; render(); scrollTo(0, y); });

// 새 판 저절로 읽기: render.py 가 다시 만들 때마다 stamp 파일을 바꾼다 · 열린 탭이 2초마다 읽어 보고 다르면 다시 불러옴 (file:// 에서도 script 읽기는 됨)
window.__hrvStamp = s => { if (s && s !== D.stamp && !CLOSED && edOpen()) { RELOAD_WAIT = true; $('#ed-new').hidden = false; return; }   // 고치는 중에는 다시 읽지 않고 닫을 때 읽음
  if (s && s !== D.stamp && !CLOSED) { CLOSED = true; try { sessionStorage.setItem('hrv_y', String(scrollY)); } catch (e) {} location.reload(); } };
if (D.stamp_src && location.protocol === 'file:') setInterval(() => { if (CLOSED) return; const sc = document.createElement('script'); sc.src = D.stamp_src + '?t=' + Date.now(); sc.onload = sc.onerror = () => sc.remove(); document.head.appendChild(sc); }, 2000);
// 탭이 쌓이지 않게: 새로 연 탭 (다시 불러오기가 아닌 열기)이 알리면 같은 파일을 연 지난 탭은 스스로 닫는다 · 닫지 못하는 브라우저면 안내만
(() => { const key = 'hrv_new:' + location.pathname, me = Math.random().toString(36).slice(2);
  const nav = (performance.getEntriesByType && performance.getEntriesByType('navigation')[0]) || {};
  try { if (nav.type !== 'reload' && nav.type !== 'back_forward') localStorage.setItem(key, JSON.stringify({id: me, t: Date.now()})); } catch (e) {}
  addEventListener('storage', e => { if (e.key !== key || !e.newValue) return; let v = {}; try { v = JSON.parse(e.newValue); } catch (x) {} if (v.id === me) return;
    CLOSED = true; window.close(); setTimeout(() => { if (!$('.closed')) document.body.insertAdjacentHTML('afterbegin', '<div class="closed">이 뷰어는 새 탭에서 다시 열렸습니다 · 이 탭은 닫아도 됩니다</div>'); }, 400); }); })();
$('#co').innerHTML = D.companies.map(c => `<button data-co="${esc(c.co)}" title="outputs/${esc(c.co)}/">${esc(coName(c.co))}</button>`).join('');
$('#co').addEventListener('click', e => { if (e.target.dataset.co) { CO = e.target.dataset.co; render(); } });
$('#menu').addEventListener('click', () => document.body.classList.toggle('nav-open'));
$('#scrim').addEventListener('click', () => document.body.classList.remove('nav-open'));
$('#menu').innerHTML = ic('menu'); $('#theme').innerHTML = ic('theme');
$('#theme').addEventListener('click', () => { const r = document.documentElement, dark = r.dataset.theme ? r.dataset.theme === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches; r.dataset.theme = dark ? 'light' : 'dark'; });
$('#ed').innerHTML = edHtml();
$('#ed-a').addEventListener('input', edUpdate); $('#ed-r').addEventListener('input', edUpdate);
$('#ed-a').addEventListener('keydown', e => { if (e.key === 'Tab' && !e.shiftKey && !e.ctrlKey && !e.altKey) { e.preventDefault(); const a = e.target, s = a.selectionStart; a.setRangeText('  ', s, a.selectionEnd, 'end'); edUpdate(); } });
$('#ed-x').addEventListener('click', edClose);
$('#ed-0').addEventListener('click', () => { if (CURF && confirm('초안을 버리고 원본으로 되돌릴까요?')) { $('#ed-a').value = CURF.raw; edUpdate(); } });
$('#ed-k').addEventListener('click', () => { const it = edItem(); if (!it) return; ED.basket.push(it); if (CURF) delete ED.drafts[CURF.file]; edSave(); $('#ed-m').innerHTML = `<span class="ok">바구니에 담음</span> · 지금 ${ED.basket.length}건`; renderNav(); setTimeout(edClose, 700); });
$('#ed-cp').addEventListener('click', () => { const it = edItem(); if (it) copyText(reqText([it]), $('#ed-c'), $('#ed-m')); });
$('#ed').addEventListener('click', e => { if (e.target.id === 'ed') edClose(); });
$('#bk').addEventListener('click', e => {
  if (e.target.id === 'bk' || e.target.id === 'bk-x') return closeBasket();
  const d = e.target.closest('[data-bkdel]'); if (d) { ED.basket.splice(+d.dataset.bkdel, 1); edSave(); openBasket(); return; }
  if (e.target.id === 'bk-cp') { copyText(reqText(ED.basket), $('#bk-c'), $('#bk-m')).then(() => $('#bk-m').insertAdjacentHTML('beforeend', ' · 붙여 넣은 뒤 「바구니 비우기」')); return; }
  if (e.target.id === 'bk-0' && confirm('바구니를 비울까요? (붙여 넣은 뒤에 비우세요)')) { ED.basket = []; edSave(); openBasket(); renderNav(); } });
document.addEventListener('click', e => {
  const x = e.target.closest('[data-edit],[data-idea],[data-jump],[data-bk]'); if (!x) return; e.preventDefault();
  if (x.dataset.edit) openEditor(x.dataset.edit); else if (x.dataset.idea) openIdea(x.dataset.idea); else if (x.dataset.bk) openBasket();
  else { const t = document.getElementById(x.dataset.jump); if (t) { const dd = t.closest('details') || (t.tagName === 'DETAILS' ? t : null); if (dd) dd.open = true; t.scrollIntoView({block: 'start', behavior: 'smooth'}); } } });
addEventListener('keydown', e => { if (e.key !== 'Escape') return; if ($('#bk').classList.contains('on')) closeBasket(); else if (edOpen()) edClose(); });
try { if (sessionStorage.getItem('hrv_y') != null) RESTORED = true; } catch (e) {}
render();
try { const f = sessionStorage.getItem('hrv_ed_open'); if (f && HF[f]) openEditor(f); } catch (e) {}
try { const y = sessionStorage.getItem('hrv_y'); if (y != null) { sessionStorage.removeItem('hrv_y'); requestAnimationFrame(() => scrollTo(0, +y)); } } catch (e) {}
"""


PORT_CSS = r"""
[hidden]{display:none!important}
.dc-hide,.vf-hide{display:none!important}
.ri{display:inline-grid;place-items:center;width:16px;height:16px;border-radius:50%;font:italic 700 10.5px/1 Georgia,serif;color:var(--muted);border:1px solid var(--line);margin:0 3px;cursor:help;vertical-align:1px;background:var(--surface)}
.ri:before{content:"i"}.ri:hover,.ri:focus{color:#fff;background:var(--accent);border-color:var(--accent);outline:0}h2 .ri,h3 .ri,h4 .ri,h5 .ri{vertical-align:3px}
.clamp{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;cursor:pointer;min-width:12em}.clamp.open{display:block}.clamp:not(.open):hover{color:var(--accent-ink)}
#tip{max-width:min(380px,calc(100vw - 16px));line-height:1.55}
abbr.g.rp{border-bottom:0}.nm{font-size:.84em;color:var(--muted);margin-left:3px;font-weight:500;white-space:nowrap}
td.dt{white-space:nowrap;font-variant-numeric:tabular-nums}
.mtoc{display:flex;flex-wrap:wrap;gap:6px;margin:4px 0 12px;padding:0 0 12px;border-bottom:1px dashed var(--line)}
.mtoc a{font-size:12.5px;color:var(--ink-2);text-decoration:none;border:1px solid var(--line);border-radius:999px;padding:2px 10px;background:var(--surface-2);max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.mtoc a:hover{color:var(--accent-ink);border-color:var(--accent)}
.docwrap>.doc{min-width:0}.mtoc a{min-width:0}
.mdbody h2,.mdbody h3,.mdbody h4,.mdbody h5{scroll-margin-top:calc(var(--stick,56px) + 16px)}
.withtoc{display:grid;grid-template-columns:minmax(0,1fr) 220px;gap:20px;align-items:start}.withtoc>.card{min-width:0;margin:0}
.rtoc{position:sticky;top:calc(var(--stick,56px) + 16px);max-height:calc(100vh - var(--stick,56px) - 40px);overflow:auto;font-size:12.5px;border-left:2px solid var(--line);padding:2px 0}
.rtoc b{display:block;font-size:11.5px;color:var(--muted);padding:0 12px 6px;letter-spacing:.3px}
.rtoc a{display:block;color:var(--ink-2);text-decoration:none;padding:4px 12px;border-left:2px solid transparent;margin-left:-2px;line-height:1.4}.rtoc a:hover{color:var(--accent-ink)}
.rtoc a.on{color:var(--accent-ink);border-left-color:var(--accent);font-weight:700}.rtoc a.l4,.rtoc a.l5{padding-left:22px;font-size:12px}
@media (min-width:1101px){#p-report .mtoc{display:none}}
@media (max-width:1100px){.withtoc{grid-template-columns:minmax(0,1fr)}.rtoc{display:none}}
.vbar{display:flex;flex-wrap:wrap;gap:8px 14px;align-items:center;margin:0 0 12px}
.vbar input{font:inherit;font-size:13.5px;padding:6px 12px;border:1px solid var(--line);border-radius:8px;background:var(--surface);color:var(--ink);flex:1 1 220px;max-width:360px;min-width:0}
.vbar input:focus{outline:2px solid var(--accent);outline-offset:-1px}
.chips{display:flex;flex-wrap:wrap;gap:6px;align-items:center}.chips .lbl{font-size:12px;color:var(--muted);margin-right:2px}
.chip{border:1px solid var(--line);background:var(--surface);color:var(--ink-2);border-radius:999px;padding:3px 11px;font:inherit;font-size:12.5px;cursor:pointer;white-space:nowrap}
.chip:hover{border-color:var(--accent)}.chip.on{background:var(--ink);border-color:var(--ink);color:var(--surface);font-weight:700}.chip em{font-style:normal;opacity:.7;margin-left:5px}
.kpi .vk{cursor:pointer}.kpi .vk:hover{border-color:var(--accent)}.kpi .vk.on{outline:2px solid var(--accent);outline-offset:-2px}
.vlog .tw{max-height:none}.vlog td{overflow-wrap:break-word}.vlog td.ok,.vlog td.bad,.vlog td.warn{white-space:nowrap}
@media (min-width:901px){.vlog .tw{overflow:visible}.vlog th{top:var(--stick,56px);z-index:2}.vlog .tw.wide{overflow-x:auto}.vlog .tw.wide th{top:0}}.vlog .clamp{min-width:8em}
.kpi{word-break:keep-all}.kpi>div.aside{border-style:dashed;background:var(--surface-2);box-shadow:none}
#totop{position:fixed;right:20px;bottom:20px;width:42px;height:42px;border-radius:50%;border:1px solid var(--line);background:var(--surface);color:var(--ink);cursor:pointer;box-shadow:0 4px 14px rgba(0,0,0,.12);opacity:0;pointer-events:none;transition:opacity .2s;z-index:30;display:flex;align-items:center;justify-content:center}
#totop.on{opacity:1;pointer-events:auto}#totop:hover{border-color:var(--accent);color:var(--accent-ink)}
#detail button,#co button{white-space:nowrap}#nav .badge{white-space:nowrap;flex:none}#nav button>span:not(.badge){min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
@media (max-width:640px){#stat{display:none}.top .seg button{padding:4px 8px}#totop{right:14px;bottom:14px}}
@media print{#totop,.rtoc,.mtoc,.vbar,#detail,#co,#theme,#menu,.btn,.icon-btn,#tip{display:none!important}section.pane{break-before:auto}section.pane:not(.on){display:none!important}
 .withtoc,.docwrap{display:block}.doclist{display:none!important}.tw,.vlog .tw,.hc details>.in,.diff pre,.dv{max-height:none!important;overflow:visible!important}th,.bkbar,.rtoc{position:static!important}.clamp{display:block!important;overflow:visible}}
"""

STAMP = os.path.join('.claude', 'hr-lib', '_viewer_stamp.js')   # 열린 뷰어 탭이 2초마다 읽어 새 판인지 봄 (작업 폴더 안 · 결과 파일 아님)

# 화면 칸 · 설계서 화면 아래 「참고 문서」 (작업 폴더 양식/ · 이름 · 경로만 · 읽기 전용)
# ---------------------------------------------------------------------------
PANES = ('home', 'harness', 'flow', 'dash', 'out', 'doc', 'report', 'verify', 'changes', 'design', 'regress')
REF_DOCS = (('실데이터 전환 점검표', '실데이터_전환_점검표.md'), ('GitHub 로 가져가기', 'GitHub_가져가기.md'))


def ref_docs(root):
    return [dict(t=t, path='양식/' + n, exists=os.path.isfile(os.path.join(root, '양식', n))) for t, n in REF_DOCS]


def build(root, out_file):
    stamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S%f')
    sp = os.path.join(root, STAMP)
    try:
        src = os.path.relpath(sp, os.path.dirname(os.path.abspath(out_file))).replace('\\', '/')
    except ValueError:   # 다른 드라이브
        src = ''
    cos = sorted(d for d in os.listdir(os.path.join(root, 'outputs')) if os.path.isdir(os.path.join(root, 'outputs', d))) if os.path.isdir(os.path.join(root, 'outputs')) else []
    set_gloss()   # 하네스 화면 = 회사 없이 공통 풀이 (CP · 단계 · 방식 · 4B)
    data = dict(harness=harness(root), companies=[company(root, c) for c in cos], gloss={k: v for k, v in GLOSS.items()}, refs=ref_docs(root),
                built=datetime.datetime.now().strftime('%Y-%m-%d %H:%M'), root=os.path.basename(root),
                virtual_line=H.VIRTUAL_LINE, virtual_mark=H.VIRTUAL_MARK, stamp=stamp, stamp_src=src)
    payload = json.dumps(data, ensure_ascii=False).replace('</', '<\\/')
    panes = ''.join(f'<section class="pane" id="p-{k}"></section>' for k in PANES)
    page = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>인력계획 하네스 뷰어</title><style>{CSS}{PORT_CSS}</style></head><body>
<div class="app"><aside class="side" id="side"><div class="brand"><span class="mark"></span><div><b>인력계획 뷰어</b><small>pwc-hr-harness</small></div></div>
<nav id="nav" class="tabs" aria-label="화면"></nav><div class="sidefoot">{E(data['root'])} · 결과 파일을 읽기만 함</div></aside><div class="scrim" id="scrim"></div>
<div class="col"><header class="top"><button id="menu" class="icon-btn" aria-label="메뉴"></button><h1 id="ttl"></h1><span id="stat"></span><span class="sp"></span>
<div class="seg" id="co" aria-label="회사"></div><span class="built">{data['built']} 생성</span><div class="seg" id="detail" role="group" aria-label="보기 방식"><button data-mode="simple" title="출처 · 근거 · 참고 설명은 ⓘ 로 접고 빈 열은 숨기고 긴 칸은 두 줄만">간단히</button><button data-mode="detail" title="출처 · 근거 · 참고 설명을 모두 펼침">자세히</button></div><button id="theme" class="icon-btn" title="밝게 · 어둡게" aria-label="밝게 · 어둡게"></button></header>
<div id="vb"></div><div id="dc"></div><main>{panes}</main></div></div>
<div id="tip"></div><button id="totop" title="맨 위로" aria-label="맨 위로"></button><div class="modal" id="ed"></div><div class="modal" id="bk"></div><script type="application/json" id="data">{payload}</script><script>{JS}</script></body></html>"""
    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(page)
    if src and os.path.isdir(os.path.dirname(sp)):   # 뷰어를 다 쓴 다음에 바꾼다 (열린 탭이 반쯤 쓴 파일을 읽지 않게)
        with open(sp, 'w', encoding='utf-8') as f:
            f.write(f'window.__hrvStamp && window.__hrvStamp("{stamp}");\n')
    return out_file, data


OFF = ('0', 'false', 'no', 'off', '끔')


def auto_open_off():
    """자동 실행 환경에서는 브라우저를 띄우지 않는다: HR_VIEWER_OPEN=0 (시험 · 동영상 시연 녹화 claude -p) · CI."""
    v = os.environ.get('HR_VIEWER_OPEN', '').strip().lower()
    if v in OFF:
        return 'HR_VIEWER_OPEN=' + v
    if os.environ.get('CI') and v not in ('1', 'true', 'yes', 'on'):
        return 'CI'
    return ''


def open_browser(path):
    """기본 브라우저로 연다 (Windows · Mac · 그 밖). 열린 지난 탭은 새 탭이 열리면 스스로 닫힌다 (JS)."""
    path = os.path.abspath(path)
    try:
        if sys.platform == 'win32':
            os.startfile(path)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            webbrowser.open(pathlib.Path(path).as_uri())
        return True
    except Exception:
        try:
            return webbrowser.open(pathlib.Path(path).as_uri())
        except Exception:
            return False


def main():
    ap = argparse.ArgumentParser(description='hr-view 뷰어 생성')
    ap.add_argument('--root', default=None)
    ap.add_argument('--out', default='hr-viewer.html')
    ap.add_argument('--open', action='store_true', help='(예전 이름 · 지금은 늘 엶) 만든 뒤 브라우저로 열기')
    ap.add_argument('--no-open', action='store_true', help='만들기만 하고 브라우저는 열지 않음')
    a = ap.parse_args()
    root = H.find_root(a.root)
    path, data = build(root, os.path.join(root, a.out))
    off = 'no-open' if a.no_open else auto_open_off()
    opened = False if off else open_browser(path)
    dec = [dict(co=c['co'], kind=c['decision']['kind'], id=c['decision']['sid']) for c in data['companies'] if c.get('decision')]
    print(json.dumps(dict(status='완료', file=path.replace('\\', '/'), companies=[c['co'] for c in data['companies']],
                          agents=len(data['harness']['agents']), skills=len(data['harness']['skills']),
                          opened=opened, open_off=off or None, decide=dec), ensure_ascii=False))


if __name__ == '__main__':
    main()
