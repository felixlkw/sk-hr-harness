# -*- coding: utf-8 -*-
"""한 장 정리 (모드 G · New3강 마무리 · 0.13.4) · 「지금까지 한 내용을 한 장으로 정리해 줘」.

python -X utf8 onepage.py collect [--root C:\\PwC] [--dir <하네스 폴더> ...] [--since YYYY-MM-DD] [--name 이름]
    → 오늘 (또는 --since 이후) 만들거나 바꾼 하네스 폴더를 찾아 사실만 JSON 으로 출력 (계산 · 판단 없음 · 결정적)
python -X utf8 onepage.py render --facts <facts.json> [--summary <요약.json>] [--out <폴더>]
    → 한장정리_<이름>_<날짜>.html (한 장 · 인쇄 가능) + 하네스 사본 zip (자료 · 결과 폴더 제외) · 경로를 JSON 으로 출력

- 하네스 폴더 = CLAUDE.md 또는 .claude/agents · .claude/skills 가 있는 폴더 (플러그인 폴더 · 숨김 · data · outputs · _workspace 안은 찾지 않음)
- 이름: --name > git config user.name > 「(이름 없음)」
- 요약.json (Claude 가 facts 안에서만 씀): {"한줄": "...", "하네스": {"<폴더 이름>": {"주제": "...", "선정기준": ["반복성", ...], "다음": "..."}}, "다음에 할 일": ["..."]}
- 개인 식별 정보 · 자료 내용은 담지 않음 (파일 이름 · 개수 · 문서 머리말 · 변경 이력만)
"""
import argparse, datetime, glob, html, json, os, re, subprocess, sys, zipfile

SKIP_DIRS = {'.git', 'data', 'outputs', '_workspace', '__pycache__', 'node_modules', '입력', '출력', '결과보고', '제출'}
ZIP_SKIP = SKIP_DIRS | {'양식'}
PLUGIN_MARK = os.path.join('.claude-plugin', 'plugin.json')


def today():
    return datetime.date.today().isoformat()


def mtime_date(p):
    return datetime.date.fromtimestamp(os.path.getmtime(p)).isoformat()


def is_harness(d):
    if os.path.exists(os.path.join(d, PLUGIN_MARK)):
        return False
    return (os.path.exists(os.path.join(d, 'CLAUDE.md')) or os.path.isdir(os.path.join(d, '.claude', 'skills'))
            or os.path.isdir(os.path.join(d, '.claude', 'agents')))


def harness_files(d):
    out = []
    for base, dirs, files in os.walk(d):
        dirs[:] = [x for x in dirs if x not in SKIP_DIRS and not (x.startswith('.') and x != '.claude')]
        for f in files:
            out.append(os.path.join(base, f))
    return out


def find(root, since, depth=3):
    found = []
    root = os.path.abspath(root)
    for base, dirs, _ in os.walk(root):
        lvl = os.path.relpath(base, root).count(os.sep) + (0 if base == root else 1)
        dirs[:] = [x for x in dirs if x not in SKIP_DIRS and not x.startswith('.')]
        if is_harness(base):
            if any(mtime_date(f) >= since for f in harness_files(base) if os.path.exists(f)):
                found.append(base)
            dirs[:] = []   # 하네스 안은 더 찾지 않음
        if lvl >= depth:
            dirs[:] = []
    return found


def front(p):
    """SKILL.md · agent.md 머리말 name · description (첫 문장)."""
    try:
        t = open(p, encoding='utf-8').read()
    except Exception:
        return None, ''
    m = re.match(r'---\s*\n(.*?)\n---', t, re.S)
    if not m:
        return None, ''
    head = m.group(1)
    nm = re.search(r'^name:\s*(.+)$', head, re.M)
    ds = re.search(r'^description:\s*"?(.+?)"?\s*$', head, re.M)
    d = (ds.group(1) if ds else '').strip()
    d = re.split(r'(?<=[.다음함])\s', d)[0][:140]
    return (nm.group(1).strip() if nm else os.path.basename(os.path.dirname(p))), d


def change_log(claude_md):
    """CLAUDE.md 「변경 이력」 표 행 [(날짜, 변경, 대상, 사유)]."""
    if not os.path.exists(claude_md):
        return []
    rows = []
    for line in open(claude_md, encoding='utf-8'):
        c = [x.strip() for x in line.strip().strip('|').split('|')]
        if len(c) >= 3 and re.match(r'\d{4}-\d{2}-\d{2}', c[0]):
            rows.append(dict(날짜=c[0], 변경=c[1][:160], 대상=(c[2] if len(c) > 3 else '')[:80], 사유=c[-1][:120]))
    return rows


RULE_RX = re.compile(r'(금지|하지 않(는다|음|고)|않는다|지어내지|확인 필요|승인|CP\d|멈춘|멈추|멈춤|사람이|개인 식별|개인정보|스크립트만|스크립트가)')
RULE_STRONG = re.compile(r'(금지|지어내지|확인 필요|승인|멈춘|멈추|멈춤|개인 식별|개인정보|원본|사람이|사용자에게|스크립트만|스크립트가)')
LEAD_RX = re.compile(r'^(그래서|그리고|또한|따라서|즉|단|다만|이때|여기서)[,\s]+')


def rules(paths, n=3):
    seen, out = set(), []
    for p in paths:
        try:
            lines = open(p, encoding='utf-8').read().splitlines()
        except Exception:
            continue
        for ln in lines:
            if ln.count('|') >= 2 or ln.lstrip().startswith(('Agent(', 'python', '```')):
                continue
            s = LEAD_RX.sub('', re.sub(r'^[\s\-*#>|0-9.=:·]+', '', ln).strip())
            if 12 <= len(s) <= 200 and RULE_RX.search(s) and s not in seen and not s.startswith(('name:', 'description:')):
                seen.add(s)
                s = s.replace('**', '').replace('`', '')
                out.append((0 if RULE_STRONG.search(s) else 1, len(out), s if len(s) <= 100 else s[:98] + '…'))
    return [x for _, _, x in sorted(out)[:n]]


def runs(d, since):
    """오늘 만든 결과 파일 (출력 · outputs · 결과 · _workspace) 이름과 개수."""
    hits = []
    for sub in ('출력', 'outputs', '결과보고', '결과', '_workspace'):
        for base, _, files in os.walk(os.path.join(d, sub)):
            for f in files:
                p = os.path.join(base, f)
                if mtime_date(p) >= since:
                    hits.append(os.path.relpath(p, d).replace('\\', '/'))
    for sub in glob.glob(os.path.join(d, '*', '')):   # 하위 예시 폴더 (하네스 루트 아래 업무 폴더)
        nm = os.path.basename(os.path.dirname(sub))
        if nm in SKIP_DIRS or nm.startswith('.'):
            continue
        for s2 in ('출력', 'outputs', '결과보고', '_workspace'):
            for base, _, files in os.walk(os.path.join(sub, s2)):
                for f in files:
                    p = os.path.join(base, f)
                    if mtime_date(p) >= since:
                        hits.append(os.path.relpath(p, d).replace('\\', '/'))
    return sorted(set(hits))


def wfp_check(d):
    """표준 인력계획 팀 (hr-agent) 이면 3부 자사화 흔적: 정리 Agent · 대응표 · 6종 · CP1."""
    if not os.path.exists(os.path.join(d, '.claude', 'skills', 'wfp-lead', 'SKILL.md')):
        return None
    r = dict(정리_Agent=os.path.exists(os.path.join(d, '.claude', 'agents', 'wfp-intake.md')),
             정리_스킬=os.path.isdir(os.path.join(d, '.claude', 'skills', 'wfp-intake-format')), 회사=[])
    for co in sorted(glob.glob(os.path.join(d, 'data', '*', ''))):
        nm = os.path.basename(os.path.dirname(co))
        if nm in ('K-co',) or nm.endswith('_원본'):
            continue
        six = sorted(os.path.basename(p) for p in glob.glob(os.path.join(co, f'{nm}_[1-6]_*')))
        cp1 = os.path.join(d, 'outputs', nm, '①_열대응표.md')
        cp1s = '없음'
        if os.path.exists(cp1):
            t = open(cp1, encoding='utf-8').read()
            cp1s = '확정' if re.search(r'\|\s*확정\s*\|', t) else '대기'
        r['회사'].append(dict(폴더=nm, 원본=os.path.isdir(os.path.join(d, 'data', nm + '_원본')),
                             대응표=os.path.exists(os.path.join(co, '_정리_대응표.md')), 기본6종=len(six), CP1=cp1s))
    return r


def facts_one(d, since):
    cm = os.path.join(d, 'CLAUDE.md')
    agents = [front(p) for p in sorted(glob.glob(os.path.join(d, '.claude', 'agents', '*.md')))]
    skills = []
    for p in sorted(glob.glob(os.path.join(d, '.claude', 'skills', '*', 'SKILL.md'))):
        nm, ds = front(p)
        skills.append(dict(이름=nm, 설명=ds, 스크립트=len(glob.glob(os.path.join(os.path.dirname(p), 'scripts', '*.py')))))
    scripts = sorted(os.path.relpath(p, d).replace('\\', '/') for p in harness_files(d) if p.endswith('.py'))
    title = goal = ''
    if os.path.exists(cm):
        t = open(cm, encoding='utf-8').read()
        m = re.search(r'^##\s*하네스\s*[:：]\s*(.+)$', t, re.M) or re.search(r'^#\s+(.+)$', t, re.M)
        title = m.group(1).strip() if m else ''
        g = re.search(r'\*\*목표[:：]?\*\*\s*[:：]?\s*(.+)', t)
        goal = g.group(1).strip()[:200] if g else ''
    elif os.path.exists(os.path.join(d, 'README.md')):
        m = re.search(r'^#\s+(.+)$', open(os.path.join(d, 'README.md'), encoding='utf-8').read(), re.M)
        title = m.group(1).strip() if m else ''
    rp = [cm] + [p for p in glob.glob(os.path.join(d, '.claude', 'skills', '*', 'SKILL.md'))]
    return dict(폴더=os.path.basename(d), 경로=d, 이름=title or os.path.basename(d), 목표=goal, CLAUDE_md=os.path.exists(cm),
                Agent=[dict(이름=a, 설명=b) for a, b in agents], 스킬=skills, 스크립트=scripts,
                규칙=rules(rp), 변경이력=change_log(cm), 오늘_결과=runs(d, since)[:12], 오늘_결과_수=len(runs(d, since)),
                표준팀=wfp_check(d))


def git_name():
    try:
        return subprocess.run(['git', 'config', '--global', 'user.name'], capture_output=True, text=True, timeout=10).stdout.strip()
    except Exception:
        return ''


def collect(a):
    since = a.since or today()
    dirs = [os.path.abspath(x) for x in (a.dir or [])]
    if not dirs:
        root = a.root or (r'C:\PwC' if os.path.isdir(r'C:\PwC') else os.path.dirname(os.getcwd()))
        dirs = find(root, since)
        cwd = os.getcwd()
        if is_harness(cwd) and cwd not in dirs:
            dirs.insert(0, cwd)
    hs = [facts_one(d, since) for d in dirs]
    out = dict(이름=a.name or git_name() or '(이름 없음)', 날짜=today(), 기준일=since, 하네스=hs,
               안내='facts 안의 사실만으로 요약.json 을 쓰고 render 를 돌린다 · 숫자 · 결과를 지어내지 않는다')
    print(json.dumps(out, ensure_ascii=False, indent=1))


def esc(x):
    return html.escape(str(x or ''))


_END = r'(?=$|[\s.,;:)」·…])'
_EUM_FIX = [('묻는다', '물음'), ('만든다', '만듦'), ('든다', '듦'), ('아니다', '아님'), ('이다', '임')]


def _jong(ch, j):
    c = ord(ch) - 0xAC00
    return chr(0xAC00 + c - c % 28 + j) if 0 <= c < 11172 else ch


def eum(t):
    """「~한다 · ~이다 · ~있다」 로 끝나는 문장을 음슴체 (~함 · ~임 · ~있음) 로."""
    t = str(t or '').replace('**', '').replace('`', '').strip()
    for a, b in _EUM_FIX:
        t = re.sub(a + _END, b, t)
    t = re.sub(r'([가-힣])는다' + _END, r'\1음', t)                      # 맡는다 · 않는다 · 넣는다
    t = re.sub(r'([있없같았었였했겠않])다' + _END, r'\1음', t)            # 있다 · 했다 · 않다
    t = re.sub(r'([가-힣])다' + _END, lambda m: (_jong(m.group(1), 16) if (ord(m.group(1)) - 0xAC00) % 28 == 4 else m.group(0)), t)  # 한다 · 된다 · 쓴다 (받침 ㄴ → ㅁ)
    return re.sub(r'\.$', '', t)


CSS = """
@page{size:A4;margin:9mm 12mm}
*{box-sizing:border-box}
body{font-family:'Malgun Gothic','맑은 고딕','Pretendard','Apple SD Gothic Neo',sans-serif;color:#111;margin:0;background:#E9EBEE;
     font-size:14px;line-height:1.65;letter-spacing:-0.03em;word-break:keep-all;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.pg{width:794px;max-width:100%;margin:24px auto;background:#fff;padding:48px 56px 40px;box-shadow:0 1px 6px rgba(0,0,0,.12)}
.top{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;margin-bottom:12px}
.org{font-size:12.5px;color:#444;letter-spacing:-0.02em;line-height:1.5}
.sign{border-collapse:collapse;font-size:12px;text-align:center}
.sign th,.sign td{border:1px solid #333;padding:3px 0;width:68px;letter-spacing:0}
.sign th{background:#EEF0F3;font-weight:700} .sign td{height:40px}
.ttl{border-top:3px solid #111;border-bottom:1px solid #111;padding:14px 0 12px;text-align:center;margin-bottom:12px}
.ttl h1{font-size:23px;font-weight:800;letter-spacing:-0.05em;margin:0}
.ttl .s{font-size:13px;color:#444;margin-top:3px;letter-spacing:-0.02em}
.meta{width:100%;border-collapse:collapse;font-size:13px;margin:10px 0 14px}
.meta th,.meta td{border:1px solid #9AA0A8;padding:5px 10px;text-align:left}
.meta th{background:#EEF0F3;font-weight:700;width:92px;white-space:nowrap}
h2{font-size:16.5px;font-weight:800;letter-spacing:-0.04em;margin:16px 0 6px;padding-bottom:3px;border-bottom:1.5px solid #1F3864;color:#1F3864}
h3{font-size:15px;font-weight:800;letter-spacing:-0.04em;margin:10px 0 4px} h3 small{font-size:12.5px;font-weight:400;color:#555;margin-left:6px;letter-spacing:-0.02em}
table.t tr,p{page-break-inside:avoid} h2,h3,p.b1{page-break-after:avoid}
p.b1{margin:6px 0 2px;padding-left:1.25em;text-indent:-1.25em;font-weight:700}
p.b2{margin:2px 0;padding-left:2.5em;text-indent:-1.25em}
p.b3{margin:1px 0;padding-left:3.6em;text-indent:-1em;color:#333;font-size:13.5px}
p.lead{border:1px solid #9AA0A8;background:#F6F7F9;padding:10px 14px;margin:6px 0 4px;font-size:14.5px;font-weight:700;letter-spacing:-0.035em}
table.t{width:100%;border-collapse:collapse;font-size:13px;line-height:1.45;margin:3px 0 4px}
table.t th,table.t td{border:1px solid #9AA0A8;padding:4px 8px;vertical-align:top;text-align:left}
table.t thead th{background:#EEF0F3;font-weight:700;text-align:center;border-top:2px solid #333}
table.t td.c{text-align:center;white-space:nowrap} table.t th.r{background:#F6F7F9;font-weight:700;width:118px;white-space:nowrap}
.mk{letter-spacing:0;margin-right:10px;white-space:nowrap}
.nil{color:#777}
.ok{font-weight:700}
.foot{border-top:1px solid #111;margin-top:26px;padding-top:6px;font-size:11.5px;color:#555;line-height:1.55;letter-spacing:-0.02em}
.end{margin-left:14px;font-weight:700;letter-spacing:0}
@media print{body{background:#fff;font-size:12.5px;line-height:1.42}.pg{width:auto;margin:0;padding:0;box-shadow:none}.ttl{padding:10px 0 8px}.ttl h1{font-size:20px}.ttl .s{font-size:12px}.meta,table.t{font-size:12px}table.t th,table.t td,.meta th,.meta td{padding:3px 7px}h2{font-size:15px;margin:10px 0 4px}h3{font-size:14px;margin:8px 0 3px}p.b1{margin:5px 0 1px}p.b3{font-size:12px}p.lead{font-size:13px;padding:7px 12px}.org{font-size:11.5px}.top{margin-bottom:8px}.sign td{height:26px}.meta{margin:8px 0 10px}p.b2{margin:1px 0}table.t{line-height:1.38}.foot{font-size:10.5px;margin-top:4px}}
"""

CRIT = ['반복성', '표준화', '정확성', '공유']
ROMAN = 'ⅠⅡⅢⅣⅤⅥ'


def render(a):
    F = json.load(open(a.facts, encoding='utf-8'))
    S = json.load(open(a.summary, encoding='utf-8')) if a.summary and os.path.exists(a.summary) else {}
    nm = F['이름']
    out_dir = a.out or (os.path.join(r'C:\PwC', '제출') if os.path.isdir(r'C:\PwC') else os.path.join(os.getcwd(), '제출'))
    os.makedirs(out_dir, exist_ok=True)
    safe = re.sub(r'[\\/:*?"<>| ]', '_', nm)
    hp = os.path.join(out_dir, f'한장정리_{safe}_{F["날짜"]}.html')
    zp = os.path.join(out_dir, f'하네스사본_{safe}_{F["날짜"]}.zip')
    H = F['하네스']
    nil = lambda t: f'<span class="nil">{t}</span>'
    body = ['<div class="pg">'
            '<div class="top"><div class="org">mySUNI 「AI Powered Workforce Planning」<br>Day 3 · 내가 필요한 HR-Harness 만들기</div>'
            '<table class="sign"><tr><th>작성</th><th>확인</th></tr><tr><td></td><td></td></tr></table></div>'
            '<div class="ttl"><h1>내 업무 하네스 구축 현황 리포트</h1><div class="s">New3강 실습 · 한 장 정리</div></div>']
    sec = 0

    def h2(t):
        nonlocal sec
        s = f'<h2>{ROMAN[sec]}. {t}</h2>'
        sec += 1
        return s
    body.append(h2('요약'))
    if S.get('한줄'):
        body.append(f'<p class="lead">{esc(eum(S["한줄"]))}</p>')
    for h in H:
        sh = (S.get('하네스') or {}).get(h['폴더'], {})
        body.append(f'<p class="b2">○ {esc(h["이름"])} : Agent {len(h["Agent"])}개 · 스킬 {len(h["스킬"])}개 · 스크립트 {len(h["스크립트"])}개로 구성함</p>'
                    f'<p class="b3">- 시험 실행 결과 파일 {h["오늘_결과_수"]}개 · 개선 이력 {len(h["변경이력"])}건</p>')
    body.append(h2('하네스별 구축 현황'))
    for i, h in enumerate(H, 1):
        sh = (S.get('하네스') or {}).get(h['폴더'], {})
        got = set(sh.get('선정기준', []))
        crit = ''.join(f'<span class="mk">{"■" if c in got else "□"} {c}</span>' for c in CRIT)
        b = [f'<div class="blk"><h3>{i}. {esc(h["이름"])}<small>(폴더 {esc(h["폴더"])})</small></h3>',
             '<p class="b1">□ 개요</p><table class="t">',
             f'<tr><th class="r">업무 주제</th><td>{esc(eum(sh.get("주제") or h["이름"]))}</td></tr>']
        if h['목표'] and not sh.get('주제'):
            b.append(f'<tr><th class="r">목표</th><td>{esc(eum(h["목표"]))}</td></tr>')
        b.append(f'<tr><th class="r">선정 기준</th><td>{crit} <span class="nil">(해당 {len(got & set(CRIT))} / 4)</span></td></tr>'
                 '</table></div>')
        ag = ', '.join(esc(x['이름']) for x in h['Agent'])
        sk = ', '.join(esc(x['이름']) + (f' (스크립트 {x["스크립트"]})' if x['스크립트'] else '') for x in h['스킬'])
        sc = ', '.join(esc(os.path.basename(x)) for x in h['스크립트'])
        b.append('<div class="blk"><p class="b1">□ 구성</p><table class="t"><thead><tr><th style="width:118px">구분</th><th style="width:60px">수</th><th>명칭</th></tr></thead>'
                 f'<tr><th class="r">Agent (역할)</th><td class="c">{len(h["Agent"])}</td><td>{ag or nil("없음 (스킬 중심 구성)")}</td></tr>'
                 f'<tr><th class="r">스킬 (방법)</th><td class="c">{len(h["스킬"])}</td><td>{sk or nil("없음")}</td></tr>'
                 f'<tr><th class="r">스크립트 (계산)</th><td class="c">{len(h["스크립트"])}</td><td>{sc or nil("없음 (AI 가 직접 처리)")}</td></tr></table></div>')
        b.append('<div class="blk"><p class="b1">□ 운영 통제 (사람 확인 자리 · 하지 않는 일)</p>')
        if h.get('표준팀'):
            b.append('<p class="b2">○ 표준 팀 규칙을 그대로 적용함 (CP 5곳 · 감원 단독 대안 금지 · 폴더 밖 자료 금지 · 검증 관문)</p>')
        elif h['규칙']:
            b += [f'<p class="b2">○ {esc(eum(x))}</p>' for x in h['규칙']]
        else:
            b.append(f'<p class="b2">○ {nil("문서에서 확인되지 않음")}</p>')
        b.append('</div><div class="blk"><p class="b1">□ 시험 실행</p>')
        if h['오늘_결과_수']:
            b.append(f'<p class="b2">○ 시험 실행으로 결과 파일 {h["오늘_결과_수"]}개가 생성됨</p>'
                     f'<p class="b3">- {", ".join(esc(x) for x in h["오늘_결과"][:4])}'
                     + (f' 외 {h["오늘_결과_수"] - 4}개' if h['오늘_결과_수'] > 4 else '') + '</p>')
        else:
            b.append(f'<p class="b2">○ {nil("결과 파일 없음")}</p>')
        b.append('</div><div class="blk"><p class="b1">□ 개선 이력 (CLAUDE.md 변경 이력)</p>')
        if h['변경이력']:
            b.append('<table class="t"><thead><tr><th style="width:92px">일자</th><th>변경 내용</th><th style="width:34%">사유</th></tr></thead>'
                     + ''.join(f'<tr><td class="c">{esc(r["날짜"])}</td><td>{esc(eum(r["변경"]))}</td><td>{esc(eum(r["사유"])) if r["사유"] and r["사유"] != "-" else nil("-")}</td></tr>' for r in h['변경이력'])
                     + '</table>')
        else:
            b.append(f'<p class="b2">○ {nil("기록 없음")}</p>')
        b.append('</div>')
        if h.get('표준팀'):
            t = h['표준팀']
            b.append('<div class="blk"><p class="b1">□ 표준 인력계획 팀 자사화</p>'
                     f'<p class="b2">○ 정리 Agent {"있음" if t["정리_Agent"] else "없음"} · 정리 스킬 {"있음" if t["정리_스킬"] else "없음"}</p>')
            if t['회사']:
                b.append('<table class="t"><thead><tr><th>회사 폴더</th><th>원본</th><th>대응표</th><th>기본 6종</th><th>CP1</th></tr></thead>'
                         + ''.join(f'<tr><td>{esc(c["폴더"])}</td><td class="c">{"있음" if c["원본"] else "없음"}</td><td class="c">{"있음" if c["대응표"] else "없음"}</td>'
                                   f'<td class="c">{c["기본6종"]}개</td><td class="c">{esc(c["CP1"])}</td></tr>' for c in t['회사']) + '</table>')
            else:
                b.append(f'<p class="b2">○ {nil("자사 회사 폴더 없음")}</p>')
            b.append('</div>')
        body += b
    nxt = [(f'{h["이름"]} : ' if len(H) > 1 else '') + (S.get('하네스') or {}).get(h['폴더'], {}).get('다음', '') for h in H]
    nxt = [x for x in nxt if x and not x.endswith(' : ')] + list(S.get('다음에 할 일') or [])
    if nxt:
        body.append(h2('향후 계획') + ''.join(f'<p class="b2">○ {esc(eum(x))}' + ('<span class="end">끝.</span>' if k == len(nxt) - 1 else '') + '</p>' for k, x in enumerate(nxt)))
    else:
        body.append('<p class="b2" style="text-align:right"><span class="end">끝.</span></p>')
    body.append('<div class="foot">※ 구성 · 규칙 · 개선 이력 · 결과 파일 수는 폴더에서 자동으로 읽은 값임 · 자료 내용과 개인 정보는 담지 않음 (함께 저장한 zip 도 하네스 파일만 포함)</div></div>')
    open(hp, 'w', encoding='utf-8').write(f'<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>하네스 구축 현황 리포트 · {esc(nm)}</title><style>{CSS}</style></head><body>' + ''.join(body) + '</body></html>')
    with zipfile.ZipFile(zp, 'w', zipfile.ZIP_DEFLATED) as z:
        for h in F['하네스']:   # 하네스 파일만: CLAUDE.md · README.md · .claude/ (settings.local 제외) · scripts/ · config*
            d = h['경로']
            pick = [os.path.join(d, f) for f in ('CLAUDE.md', 'README.md') if os.path.exists(os.path.join(d, f))]
            pick += [p for p in glob.glob(os.path.join(d, 'config*')) if os.path.isfile(p)]
            for sub in ('.claude', 'scripts', 'config'):
                for base, dirs, files in os.walk(os.path.join(d, sub)):
                    dirs[:] = [x for x in dirs if x not in ZIP_SKIP and x != 'baselines']
                    pick += [os.path.join(base, f) for f in files if f != 'settings.local.json' and not f.endswith(('.pyc', '.xlsx', '.docx'))]
            for p in sorted(set(pick)):
                z.write(p, os.path.join(h['폴더'], os.path.relpath(p, d)))
    print(json.dumps(dict(status='완료', html=hp, zip=zp, 하네스=[h['폴더'] for h in F['하네스']],
                     제출=f'html 하나만 kangouk.lee@pwc.com 으로 메일 첨부 · 제목 [New3강] 한장정리_{F["이름"]} · zip 은 보내지 않음'), ensure_ascii=False, indent=1))


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('cmd', choices=['collect', 'render'])
    ap.add_argument('--root'); ap.add_argument('--dir', action='append'); ap.add_argument('--since'); ap.add_argument('--name')
    ap.add_argument('--facts'); ap.add_argument('--summary'); ap.add_argument('--out')
    a = ap.parse_args()
    collect(a) if a.cmd == 'collect' else render(a)
