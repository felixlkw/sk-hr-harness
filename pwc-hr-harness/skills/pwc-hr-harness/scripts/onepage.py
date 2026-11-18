# -*- coding: utf-8 -*-
"""한 장 정리 (모드 G · New3강 마무리 · 0.13.3) · 「지금까지 한 내용을 한 장으로 정리해 줘」.

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


RULE_RX = re.compile(r'(금지|하지 않|지어내지|확인 필요|승인|CP\d|멈춘|멈추|사람이|개인 식별|개인정보|스크립트만|스크립트가)')


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
            s = re.sub(r'^[\s\-*#>|0-9.]+', '', ln).strip()
            if 12 <= len(s) <= 200 and RULE_RX.search(s) and s not in seen and not s.startswith(('name:', 'description:')):
                seen.add(s)
                s = s.replace('**', '').replace('`', '')
                out.append(s if len(s) <= 80 else s[:78] + '…')
            if len(out) >= n:
                return out
    return out


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


CSS = """
body{font-family:'Pretendard','Malgun Gothic',sans-serif;color:#2D2926;margin:0;background:#fff}
.pg{max-width:980px;margin:0 auto;padding:28px 34px}
h1{font-size:24px;margin:0} .sub{color:#6B645D;font-size:13px;margin:4px 0 14px}
.one{background:#2D2926;color:#fff;border-radius:10px;padding:10px 16px;font-size:15px;margin-bottom:14px}
.h{border:1.5px solid #E2DDD7;border-radius:12px;padding:14px 18px;margin-bottom:14px;page-break-inside:avoid}
.h h2{font-size:18px;margin:0 0 2px;color:#D04A02} .h .g{font-size:13px;color:#6B645D;margin-bottom:10px}
table{border-collapse:collapse;width:100%;font-size:12.5px} td,th{border-top:1px solid #ECE8E3;padding:5px 8px;vertical-align:top;text-align:left}
th{width:150px;color:#6B645D;font-weight:700;background:#F7F5F2}
.chip{display:inline-block;border:1px solid #D7D2CC;border-radius:10px;padding:0 8px;margin:1px 3px 1px 0;font-size:12px}
.ok{color:#1F7A3D;font-weight:700} .no{color:#9A948D}
.foot{font-size:11.5px;color:#9A948D;margin-top:8px}
@media print{.pg{padding:10mm}}
"""


def render(a):
    F = json.load(open(a.facts, encoding='utf-8'))
    S = json.load(open(a.summary, encoding='utf-8')) if a.summary and os.path.exists(a.summary) else {}
    nm = F['이름']
    out_dir = a.out or (os.path.join(r'C:\PwC', '제출') if os.path.isdir(r'C:\PwC') else os.path.join(os.getcwd(), '제출'))
    os.makedirs(out_dir, exist_ok=True)
    safe = re.sub(r'[\\/:*?"<>| ]', '_', nm)
    hp = os.path.join(out_dir, f'한장정리_{safe}_{F["날짜"]}.html')
    zp = os.path.join(out_dir, f'하네스사본_{safe}_{F["날짜"]}.zip')
    body = [f'<div class="pg"><h1>New3강 · 내가 만든 하네스 한 장 정리</h1>'
            f'<div class="sub">{esc(nm)} · {esc(F["날짜"])} · 하네스 {len(F["하네스"])}개 · 자료 내용 · 개인 정보 없음 (폴더 구성과 기록만)</div>']
    if S.get('한줄'):
        body.append(f'<div class="one">{esc(S["한줄"])}</div>')
    for h in F['하네스']:
        sh = (S.get('하네스') or {}).get(h['폴더'], {})
        crit = ''.join(f'<span class="chip">{esc(c)}</span>' for c in sh.get('선정기준', []))
        ag = ''.join(f'<span class="chip">{esc(x["이름"])}</span>' for x in h['Agent']) or '<span class="no">없음 (스킬 중심)</span>'
        sk = ''.join(f'<span class="chip">{esc(x["이름"])}{" · 스크립트 " + str(x["스크립트"]) if x["스크립트"] else ""}</span>' for x in h['스킬']) or '<span class="no">없음</span>'
        rl = ('표준 팀 규칙 그대로 (CP 5곳 · 감원 단독 금지 · 폴더 밖 자료 금지 · 검증 관문)' if h.get('표준팀') else
              '<br>'.join(esc(x) for x in h['규칙']) or '<span class="no">문서에서 찾지 못함</span>')
        log = '<br>'.join(f'{esc(r["날짜"])} · {esc(r["변경"][:90])}' + (f' <span class="no">(사유: {esc(r["사유"][:60])})</span>' if r['사유'] and r['사유'] != '-' else '') for r in h['변경이력']) or '<span class="no">기록 없음</span>'
        rows = [('주제', esc(sh.get('주제') or h['이름'])), ('선정 기준', crit or '<span class="no">-</span>'),
                ('구성', f'Agent {len(h["Agent"])} · 스킬 {len(h["스킬"])} · 스크립트 {len(h["스크립트"])}<br>{ag}<br>{sk}'),
                ('일관성 (계산)', f'<span class="ok">스크립트 {len(h["스크립트"])}개로 계산</span>' if h['스크립트'] else '<span class="no">스크립트 없음 (AI 가 직접 처리)</span>'),
                ('확인 자리 · 하지 않는 일', rl),
                ('시험 실행', f'기간 안에 만든 결과 파일 {h["오늘_결과_수"]}개' if h['오늘_결과_수'] else '<span class="no">결과 파일 없음</span>'),
                ('개선 · 받은 피드백 반영', log)]
        if h.get('표준팀'):
            t = h['표준팀']
            cos = '<br>'.join(f'{esc(c["폴더"])}: 원본 {"있음" if c["원본"] else "없음"} · 대응표 {"있음" if c["대응표"] else "없음"} · 기본 6종 {c["기본6종"]}개 · CP1 {esc(c["CP1"])}' for c in t['회사']) or '<span class="no">자사 회사 폴더 없음</span>'
            rows.append(('표준 인력계획 팀 자사화', f'정리 Agent {"<span class=ok>있음</span>" if t["정리_Agent"] else "<span class=no>없음</span>"} · 정리 스킬 {"<span class=ok>있음</span>" if t["정리_스킬"] else "<span class=no>없음</span>"}<br>{cos}'))
        if sh.get('다음'):
            rows.append(('다음에 할 일', esc(sh['다음'])))
        body.append(f'<div class="h"><h2>{esc(h["이름"])}</h2><div class="g">폴더 {esc(h["폴더"])}' + (f' · 목표: {esc(h["목표"])}' if h['목표'] else '') + '</div><table>'
                    + ''.join(f'<tr><th>{k}</th><td>{v}</td></tr>' for k, v in rows) + '</table></div>')
    if S.get('다음에 할 일'):
        body.append('<div class="h"><h2>다음에 할 일</h2>' + '<br>'.join(esc(x) for x in S['다음에 할 일']) + '</div>')
    body.append(f'<div class="foot">만든 도구: pwc-hr-harness onepage.py · 구성 · 규칙 · 변경 이력 · 결과 파일 수는 폴더에서 자동으로 읽은 값 · 주제 · 요약 문장은 그 값 안에서 작성 · 하네스 사본: {esc(os.path.basename(zp))}</div></div>')
    open(hp, 'w', encoding='utf-8').write(f'<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>한 장 정리 · {esc(nm)}</title><style>{CSS}</style></head><body>' + ''.join(body) + '</body></html>')
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
    print(json.dumps(dict(status='완료', html=hp, zip=zp, 하네스=[h['폴더'] for h in F['하네스']]), ensure_ascii=False, indent=1))


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('cmd', choices=['collect', 'render'])
    ap.add_argument('--root'); ap.add_argument('--dir', action='append'); ap.add_argument('--since'); ap.add_argument('--name')
    ap.add_argument('--facts'); ap.add_argument('--summary'); ap.add_argument('--out')
    a = ap.parse_args()
    collect(a) if a.cmd == 'collect' else render(a)
