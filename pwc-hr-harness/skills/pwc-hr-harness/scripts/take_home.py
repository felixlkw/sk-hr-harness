# -*- coding: utf-8 -*-
"""pwc-hr-harness · 내 hr-agent 저장소 (작업 폴더 = 내 HR-Agent · GitHub 비공개 저장소 · 막히면 zip).

흐름 (2026-11-15): Day 1 /hr-init 직후 첫 푸시 (hr-agent 비공개 저장소 생성) → Day 1 · 2 · 3 마감마다 커밋 · 푸시 → Day 3 마지막에 플러그인 사본을 넣고 푸시 (회사 동료 공유)

사용 (작업 폴더에서 · /hr-init 이 .claude/hr-lib/take_home.py 에 사본을 둔다):
  python -X utf8 .claude/hr-lib/take_home.py check                     # 올리기 전 점검 (git · gh · 실데이터 · 개인 식별 열)
  python -X utf8 .claude/hr-lib/take_home.py message                   # 커밋 메시지 (첫 커밋 · 이후 저장 · CLAUDE.md 변경 이력에서)
  python -X utf8 .claude/hr-lib/take_home.py remote                    # push 전: git remote -v · 원격이 있으면 그 저장소가 PRIVATE 인지
  python -X utf8 .claude/hr-lib/take_home.py visibility [--repo 소유자/이름]   # push 뒤: gh repo view --json visibility → PRIVATE 이 아니면 멈춤
  python -X utf8 .claude/hr-lib/take_home.py copy-plugin --plugin <플러그인 폴더>   # Day 3 가져가기: 폴더 안 pwc-hr-harness/ 사본 (동료가 편집자가 될 때 이 사본으로 설치)
  python -X utf8 .claude/hr-lib/take_home.py zip [--out 파일]           # GitHub 가 막혔을 때: .gitignore 로 뺀 파일 없이 zip

하는 일 (결정적 · 아무것도 올리지 않음 · push 는 사람이 확인한 뒤 Claude 가 git · gh 로 한다)
  - .gitignore (없으면 플러그인 견본 규칙)로 올라갈 파일 목록을 셈한다
  - 막음 (종료 코드 3): 올라갈 표 파일에 개인 식별 열 (성명 · 사번 · 생년월일 · 주민 · 연락처 · 이메일 · 모든 시트의 모든 머리 행) · 가상 표시 없는 회사 데이터 폴더 (실데이터일 수 있음)
  - 가상 표시로 인정하는 것은 정해진 표시뿐: virtual_data.py 표시 (`_가상데이터.json` · 문서 속성 pwc-hr-harness:virtual · 2행 「가상 데이터 · 실제 판단에 쓰지 않음」) ·
    E-Co. 키트 표시 「가상 데이터 · 실제 회사 자료 아님」 · K-Co. 키트 표시 「K-Co. 합성 데이터 · 교육 실습 전용」. 「가상」 글자만 있는 것은 인정하지 않는다
  - 확인 필요 (막지 않음 · 사람에게 물음): 가상 폴더 안 추가_ 파일 중 정해진 표시가 없는 것 · 표 파일 없이 문서 (md · docx · pdf 등) 만 있는 회사 폴더 (실자료일 수 있음)
  - push 는 하지 않는다. push 전 `remote` · push 뒤 `visibility` 로 저장소가 비공개 (PRIVATE) 인지 확인하고, 아니면 멈춘다 (4요소)
  - 플러그인은 강사가 준 로컬 폴더에서 설치한다 (GitHub 에 따로 올리지 않음). Day 3 가져가기 때 copy-plugin 으로 hr-agent 안에 사본을 넣는다 (그 전에는 참고 칸에만 적음)
  - 회사 동료: hr-agent 를 clone 하면 플러그인 없이 인력계획을 실행 (사용자) · 전파교육 뒤 사본으로 플러그인을 설치하면 /hr-customize 로 함께 고침 (편집자) · 운영자 = 교육 이수자
  - 실데이터는 회사 폴더 이름 끝을 「_실데이터」로 두면 (예: data/자사_실데이터/) .gitignore 견본이 뺀다
"""
import os, sys, re, json, glob, fnmatch, shutil, zipfile, datetime, argparse

HERE = os.path.dirname(os.path.abspath(__file__))
EXIT_STOP = 3
BAN_COL = re.compile(r'성명|사번|생년월일|주민|휴대|전화|이메일|e-?mail', re.I)   # virtual_data.py 와 같은 목록
KCO_STAMP = 'K-Co. 합성 데이터 · 교육 실습 전용'                                  # K-Co. 키트 안내문 (가상 기업 K-Co. 합성 데이터 · 교육 실습 전용 · 실존 기업과 무관)
VIRTUAL_MARK = '_가상데이터.json'
VIRTUAL_STAMP = '가상 데이터 · 실제 판단에 쓰지 않음'                              # virtual_data.py (hrlib VIRTUAL_STAMP)
KIT_STAMP = '가상 데이터 · 실제 회사 자료 아님'                                    # 키트 gen_eco.py (E-Co. · 데이터 카드 파일 2행)
MARKS = (VIRTUAL_STAMP, KIT_STAMP, KCO_STAMP)                                      # 가상 · 교육 데이터로 인정하는 표시 (이것뿐 · 「가상」 글자만으로는 인정하지 않음)
DOC_EXT = ('.md', '.docx', '.doc', '.pdf', '.txt', '.pptx', '.ppt', '.hwp', '.hwpx', '.rtf')
SKIP_DOCS = {'README.md', '.gitkeep', '가상데이터_설정.json', '열이름_대응.json'}
VIRTUAL_TAG = 'pwc-hr-harness:virtual'
EXTRA_PREFIX = '추가_'
PLUGIN_DIR = 'pwc-hr-harness'
DATA_EXT = ('.xlsx', '.xlsm', '.xls', '.csv')


def find_root(start):
    d = os.path.abspath(start)
    while True:
        if os.path.exists(os.path.join(d, 'CLAUDE.md')) and os.path.isdir(os.path.join(d, '.claude')):
            return d
        nd = os.path.dirname(d)
        if nd == d:
            return os.path.abspath(start)
        d = nd


def _template_ignore():
    for p in (os.path.join(HERE, 'gitignore_견본'),
              os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), 'templates', 'forms', 'gitignore_견본')):
        if os.path.exists(p):
            return p
    return None


def ignore_rules(root):
    p = os.path.join(root, '.gitignore')
    src = p if os.path.exists(p) else (os.path.join(root, '양식', 'gitignore_견본') if os.path.exists(os.path.join(root, '양식', 'gitignore_견본')) else _template_ignore())
    rules = []
    if src:
        for ln in open(src, encoding='utf-8').read().splitlines():
            ln = ln.strip()
            if not ln or ln.startswith('#'):
                continue
            neg = ln.startswith('!')
            ln = ln[1:] if neg else ln
            rules.append(dict(pat=ln.rstrip('/'), dir=ln.endswith('/'), anchored='/' in ln.rstrip('/'), neg=neg))
    return rules, (os.path.relpath(src, root).replace('\\', '/') if src and src.startswith(root) else src)


def _seg_match(pat, rel):
    a, b = pat.strip('/').split('/'), rel.split('/')
    return len(a) == len(b) and all(fnmatch.fnmatchcase(y, x) for x, y in zip(a, b))


def ignored(rel, rules, is_dir=False):
    """.gitignore 의 쓰는 부분만 (주석 · ! · 끝 / = 폴더 · / 가 있으면 앞에서부터)."""
    parts = rel.split('/')
    out = False
    for r in rules:
        hit = False
        for i in range(len(parts)):
            sub = '/'.join(parts[:i + 1])
            last = i == len(parts) - 1
            if r['dir'] and last and not is_dir:
                continue   # 폴더 규칙은 폴더 (경로 앞부분)에만
            if r['anchored']:
                hit = _seg_match(r['pat'], sub)
            else:
                hit = fnmatch.fnmatchcase(parts[i], r['pat'])
            if hit:
                break
        if hit:
            out = not r['neg']
    return out


def included(root, rules, extra_skip=()):
    out = []
    for dp, dn, fn in os.walk(root):
        rel_d = os.path.relpath(dp, root).replace('\\', '/')
        rel_d = '' if rel_d == '.' else rel_d
        keep = []
        for d in sorted(dn):
            rd = f'{rel_d}/{d}' if rel_d else d
            if d == '.git' or ignored(rd, rules, is_dir=True):
                continue
            keep.append(d)
        dn[:] = keep
        for f in sorted(fn):
            rf = f'{rel_d}/{f}' if rel_d else f
            if rf in extra_skip or ignored(rf, rules):
                continue
            out.append(rf)
    return sorted(out)


HEAD_TOP = 12   # 위 12행은 글자 칸을 모두 머리 칸 후보로 봄 (안내문 · 머리 행이 있는 자리)


def _head_row(vals):
    """머리 행처럼 보이는 행 (글자 칸 2개 이상 · 숫자 칸 없음). 위 12행 아래에 놓인 표 머리 행도 잡는다."""
    xs = [v for v in vals if v is not None and str(v).strip() != '']
    return len(xs) >= 2 and all(isinstance(v, str) for v in xs)


def _xlsx_head(path, rows=HEAD_TOP):
    """(머리 칸 글자들, 위 3행 안내문, 문서 속성 키워드). 모든 시트 · 모든 행을 본다: 위 rows 행은 글자 칸 모두, 그 아래는 머리 행 모양인 행만."""
    import openpyxl
    cells, notes, kw = [], [], ''
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        kw = wb.properties.keywords or ''
        for i, sh in enumerate(wb.worksheets):
            for j, r in enumerate(sh.iter_rows(values_only=True)):
                if j >= rows and not _head_row(r):
                    continue
                for v in r:
                    if isinstance(v, str) and v.strip():
                        cells.append(v.strip())
                        if j < 3:
                            notes.append(v.strip())
        wb.close()
    except Exception:
        pass
    return cells, notes, kw


def _csv_head(path, top=3):
    """(머리 칸 글자들, 위 3행 안내문). 위 3행은 모두 · 그 아래는 머리 행 모양 (숫자 칸 없음 · 글자 칸 2개 이상)인 행만."""
    import csv
    num = re.compile(r'^-?[\d,.]+%?$')
    for enc in ('utf-8-sig', 'cp949'):
        try:
            with open(path, encoding=enc, newline='') as f:
                rows = list(csv.reader(f))
            notes = [c.strip() for r in rows[:top] for c in r if c.strip()]
            cells = list(notes)
            for r in rows[top:]:
                xs = [c.strip() for c in r if c.strip()]
                if len(xs) >= 2 and not any(num.match(x) for x in xs):
                    cells += xs
            return cells, notes
        except UnicodeDecodeError:
            continue
    return [], []


def _marked(notes, kw):
    """정해진 가상 · 교육 표시가 있는가 (문서 속성 키워드 · 위 3행 안내문). 「가상」 글자만으로는 아님."""
    text = ' '.join(notes)
    return VIRTUAL_TAG in (kw or '') or any(m in text for m in MARKS)


def _blank_dict(path):
    """데이터 사전이 빈 양식인가 (표 데이터 행마다 채운 칸이 번호 하나 이하) · virtual_data.py 와 같은 판정."""
    try:
        lines = open(path, encoding='utf-8').read().splitlines()
    except Exception:
        return False
    for i, ln in enumerate(lines):
        if not ln.startswith('|') or re.match(r'^\|\s*:?-{2,}', ln):
            continue
        if i + 1 < len(lines) and re.match(r'^\|\s*:?-{2,}', lines[i + 1]):
            continue
        cells = [c.strip() for c in ln.strip().strip('|').split('|')]
        if any(c and not re.fullmatch(r'\d+', c) for c in cells):
            return False
    return True


def _is_id_cell(s):
    return len(s) <= 20 and bool(BAN_COL.search(s))   # 열 이름처럼 짧은 칸만 (안내문 문장 「개인 식별정보 없음」은 셈하지 않음)


def check(root):
    rules, src = ignore_rules(root)
    files = included(root, rules)
    block, warn, folders = [], [], []
    for rf in files:   # 개인 식별 열 (올라갈 표 파일 모두 · 플러그인 사본 제외)
        if rf.startswith(PLUGIN_DIR + '/') or not rf.lower().endswith(DATA_EXT):
            continue
        p = os.path.join(root, rf)
        cells = _xlsx_head(p)[0] if rf.lower().endswith(('.xlsx', '.xlsm')) else _csv_head(p)[0] if rf.lower().endswith('.csv') else []
        bad = sorted({c for c in cells if _is_id_cell(c) and '없음' not in c})
        if bad:
            block.append(f'{rf}: 개인 식별 열 {" · ".join(bad[:5])}')
    data = os.path.join(root, 'data')
    for co in sorted(os.listdir(data)) if os.path.isdir(data) else []:
        d = os.path.join(data, co)
        if not os.path.isdir(d) or ignored(f'data/{co}', rules, is_dir=True):
            continue
        mine = [rf for rf in files if rf.startswith(f'data/{co}/') and rf.lower().endswith(DATA_EXT)]   # 표 파일로 판정 (문서는 표시를 읽지 않음)
        marker = os.path.exists(os.path.join(d, VIRTUAL_MARK))
        if not mine:
            docs = [rf for rf in files if rf.startswith(f'data/{co}/') and rf.lower().endswith(DOC_EXT) and os.path.basename(rf) not in SKIP_DOCS
                    and not (os.path.basename(rf).startswith('데이터사전') and _blank_dict(os.path.join(root, rf)))]
            if docs and not marker:   # 문서만 있는 폴더는 안전하다고 보지 않음 (사업계획 · 보고서 원문일 수 있음) → 사람에게 묻기
                folders.append(dict(co=co, kind='문서만 있음 (확인 필요)', files=len(docs)))
                warn.append(f'data/{co}/: 표 파일 없이 문서 {len(docs)}개 ({" · ".join(os.path.basename(x) for x in docs[:4])}) · 가상 · 교육 자료인지 사람에게 확인 '
                            f'(실자료면 폴더 이름을 「{co}_실데이터」로 바꾸면 .gitignore 가 뺌)')
            else:
                folders.append(dict(co=co, kind='표 파일 없음', files=0))
            continue
        std, extra_bad = [], []
        for rf in mine:
            fn = os.path.basename(rf)
            p = os.path.join(root, rf)
            notes, kw = ([], '')
            if fn.lower().endswith(('.xlsx', '.xlsm')):
                _, notes, kw = _xlsx_head(p)
            elif fn.lower().endswith('.csv'):
                notes = _csv_head(p)[1]
            synth = _marked(notes, kw)   # 정해진 표시만 (파일 이름 · 안내문의 「가상」 글자만으로는 인정하지 않음)
            if fn.startswith(EXTRA_PREFIX):
                if not synth:
                    extra_bad.append(rf)
            elif not synth:
                std.append(rf)
        kind = '가상 데이터 (표시 파일)' if marker else ('K-Co. 교육 데이터 · 가상 표시' if not std else '가상 표시 없음')
        folders.append(dict(co=co, kind=kind, files=len(mine)))
        if std and not marker:
            block.append(f'data/{co}/: 가상 · 교육 데이터 표시가 없는 파일 {len(std)}개 (실데이터일 수 있음): ' + ' · '.join(std[:4])
                         + f' → 실데이터면 폴더 이름을 「{co}_실데이터」로 바꾸면 .gitignore 가 뺌')
        for rf in extra_bad:
            warn.append(f'{rf}: 추가_ 자료에 정해진 가상 표시 (2행 「{VIRTUAL_STAMP}」 또는 「{KIT_STAMP}」) 가 없음 → 사람에게 가상으로 만든 자료인지 확인 (실데이터면 _실데이터 자리로)')
    plugin = os.path.exists(os.path.join(root, PLUGIN_DIR, '.claude-plugin', 'plugin.json'))
    info = []
    if not plugin:   # 2026-11-15: 플러그인은 로컬 폴더에서 설치 · Day 3 가져가기 때 사본을 넣음 (경고 아님)
        info.append(f'플러그인 사본 {PLUGIN_DIR}/ 없음 (Day 1 · 2 는 괜찮음) · Day 3 가져가기 때 `take_home.py copy-plugin` 으로 넣음 (동료가 편집자가 될 때 이 사본으로 설치)')
    if not os.path.exists(os.path.join(root, '.gitignore')):
        warn.append('.gitignore 가 없음 → 견본 규칙으로 셈함 · 올리기 전에 `양식/gitignore_견본` 을 .gitignore 로 복사')
    return dict(rules=src, files=files, block=block, warn=warn, info=info, folders=folders, plugin=plugin)


def changelog(root):
    p = os.path.join(root, 'CLAUDE.md')
    if not os.path.exists(p):
        return []
    t = open(p, encoding='utf-8').read()
    i = t.rfind('## 변경 이력')
    rows = []
    for ln in (t[i:] if i >= 0 else '').splitlines():
        if ln.startswith('|') and not re.match(r'^\|\s*:?-{2,}', ln):
            c = [x.strip() for x in ln.strip().strip('|').split('|')]
            if c and c[0] not in ('날짜',) and len(c) >= 2:
                rows.append(c)
    return rows


def plugin_version(root):
    for p in (os.path.join(root, PLUGIN_DIR, '.claude-plugin', 'plugin.json'),
              os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), '.claude-plugin', 'plugin.json')):
        try:
            return json.load(open(p, encoding='utf-8')).get('version')
        except Exception:
            continue
    try:   # 설치 기록 (/hr-init 기본 사본 README · 「pwc-hr-harness 0.10.0 · 날짜」)
        m = re.search(r'pwc-hr-harness ([\d.]+)', open(os.path.join(root, '.claude', 'hr-lib', '_standard', 'README.md'), encoding='utf-8').read())
        return m.group(1) if m else None
    except Exception:
        return None


def has_commit(root):
    """이 폴더가 이미 커밋이 있는 git 저장소인가 (git rev-parse HEAD)."""
    if not os.path.isdir(os.path.join(root, '.git')):
        return False
    rc, out, err = _run(_cmd('TAKE_HOME_GIT', 'git') + ['rev-parse', '--verify', '-q', 'HEAD'], root)
    return rc == 0 and bool(out)


def message(root):
    rows = changelog(root)
    ver = plugin_version(root)
    if has_commit(root):   # 이미 커밋이 있으면 저장 메시지 (Day 1 · 2 · 3 마감)
        last = ' · '.join(x for x in rows[-1][1:2] if x) if rows else ''
        head = f'hr-agent 저장 {datetime.date.today():%Y-%m-%d} (변경 이력 {len(rows)}행' + (f' · 최근: {last[:60]}' if last else '') + ')'
    else:
        head = f'hr-agent 첫 커밋: 기본 팀 설치 (pwc-hr-harness {ver or "?"} · 변경 이력 {len(rows)}행)'
    body = [f'- {" · ".join(x for x in r[:3] if x)}' for r in rows[-12:]]
    return head + ('\n\n변경 이력 (CLAUDE.md):\n' + '\n'.join(body) if body else '')


_CHO = ['g', 'kk', 'n', 'd', 'tt', 'r', 'm', 'b', 'pp', 's', 'ss', '', 'j', 'jj', 'ch', 'k', 't', 'p', 'h']
_JUNG = ['a', 'ae', 'ya', 'yae', 'eo', 'e', 'yeo', 'ye', 'o', 'wa', 'wae', 'oe', 'yo', 'u', 'wo', 'we', 'wi', 'yu', 'eu', 'ui', 'i']
_JONG = ['', 'k', 'k', 'k', 'n', 'n', 'n', 't', 'l', 'k', 'm', 'l', 'l', 'l', 'p', 'l', 'm', 'p', 'p', 't', 't', 'ng', 't', 't', 'k', 't', 'p', 't']


def romanize(s):
    """한글 음절 → 소리 나는 대로 로마자 (국어의 로마자 표기법 기본 대응 · 음운 변화 규칙은 넣지 않음 · 저장소 이름 예 만들기용)."""
    out = []
    for ch in s:
        o = ord(ch) - 0xAC00
        if 0 <= o < 11172:
            out.append(_CHO[o // 588] + _JUNG[(o % 588) // 28] + _JONG[o % 28])
        else:
            out.append(ch)
    return ''.join(out)


def repo_name(root):
    """저장소 이름 예: 언제나 hr-agent (2026-11-15 · 작업 폴더 = 내 HR-Agent)."""
    return 'hr-agent'


def repo_alt(root):
    """같은 이름이 이미 있을 때 대안 (GitHub 이름은 영문 · 숫자 · . _ - 만): hr-agent-<내 회사 폴더 이름 (한글은 로마자)> · 없으면 hr-agent-2."""
    data = os.path.join(root, 'data')
    cos = [d for d in sorted(os.listdir(data)) if os.path.isdir(os.path.join(data, d)) and '실데이터' not in d] if os.path.isdir(data) else []
    mine = [c for c in cos if not c.lower().startswith(('k-co', 'e-co'))] + [c for c in cos if c.lower().startswith('e-co')]
    for c in mine:
        s = re.sub(r'[^A-Za-z0-9._-]+', '-', romanize(c)).strip('-_.').lower()
        s = re.sub(r'-{2,}', '-', s)
        if re.search(r'[a-z]', s):
            return f'hr-agent-{s}'
    return 'hr-agent-2'


def tools():
    return dict(git=shutil.which('git') is not None, gh=shutil.which('gh') is not None)


def _cmd(env, default):
    """명령 (시험에서는 TAKE_HOME_GIT · TAKE_HOME_GH 로 바꿈 · 예: 「python fake_gh.py」)."""
    import shlex
    v = os.environ.get(env)
    return [x[1:-1] if len(x) > 1 and x[0] == x[-1] == '"' else x for x in shlex.split(v, posix=os.name != 'nt')] if v else [default]


def _run(cmd, root):
    import subprocess
    try:
        p = subprocess.run(cmd, cwd=root, capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=60)
        return p.returncode, (p.stdout or '').strip(), (p.stderr or '').strip()
    except (OSError, subprocess.TimeoutExpired) as e:
        return 127, '', str(e)


def remotes(root):
    """git remote -v → [dict(name, url, repo)] · repo = GitHub 소유자/이름 (GitHub 주소가 아니면 '')."""
    rc, out, err = _run(_cmd('TAKE_HOME_GIT', 'git') + ['remote', '-v'], root)
    seen, res = set(), []
    for ln in out.splitlines() if rc == 0 else []:
        x = ln.split()
        if len(x) >= 2 and (x[0], x[1]) not in seen:
            seen.add((x[0], x[1]))
            m = re.search(r'github\.com[:/]+([^/\s]+/[^/\s]+?)(?:\.git)?/?$', x[1])
            res.append(dict(name=x[0], url=x[1], repo=m.group(1) if m else ''))
    return rc, res, err


def visibility(root, repo=''):
    """gh repo view [repo] --json visibility,nameWithOwner,url → (visibility 대문자 | '', 저장소, 주소, 오류)."""
    rc, out, err = _run(_cmd('TAKE_HOME_GH', 'gh') + ['repo', 'view'] + ([repo] if repo else []) + ['--json', 'visibility,nameWithOwner,url'], root)
    if rc != 0:
        return '', repo, '', err or out or f'종료 코드 {rc}'
    try:
        j = json.loads(out)
    except Exception:
        return '', repo, '', '응답을 읽지 못함: ' + out[:120]
    return str(j.get('visibility') or '').upper(), j.get('nameWithOwner') or repo, j.get('url') or '', ''


def not_private(repo, vis, err, when):
    what = (f'저장소 {repo or "(이 폴더 원격)"} 가 비공개 (PRIVATE) 가 아님 ({vis})' if vis else f'저장소 {repo or "(이 폴더 원격)"} 가 비공개 (PRIVATE) 인지 확인하지 못함') + f' · {when}'
    return dict(status='멈춤', kind='가져가기', what=what,
                evidence=[f'gh repo view --json visibility → {vis or "(확인 못 함)"}'] + ([err[:200]] if err else []) + ['공개 저장소 금지 (교육 참가자 사내 사용 조건 · 하네스에 회사 조건)'],
                options=[f'gh repo edit {repo or "<소유자/이름>"} --visibility private --accept-visibility-change-consequences 뒤 `take_home.py visibility` 로 다시 확인',
                         '사람이 브라우저에서 저장소 Settings 의 공개 범위를 확인 · 고친 뒤 다시 확인 (gh 로그인이 풀렸으면 terminal 패널에서 gh auth login)',
                         '보류 (push 하지 않음 · 막히면 take_home.py zip)'],
                answer='확정 · 추가 확인 · 보류 · 이관')


def main():
    ap = argparse.ArgumentParser(description='내 하네스 가져가기 (점검 · 커밋 메시지 · 플러그인 사본 · zip)')
    ap.add_argument('cmd', choices=['check', 'message', 'remote', 'visibility', 'copy-plugin', 'zip'])
    ap.add_argument('--repo', default='', help='visibility: 소유자/이름 (없으면 이 폴더 원격)')
    ap.add_argument('--root', default=None)
    ap.add_argument('--plugin', default=None, help='copy-plugin (선택 · 오프라인 · zip 용): 플러그인 폴더 (보통 ${CLAUDE_PLUGIN_ROOT})')
    ap.add_argument('--out', default=None, help='zip: 만들 파일 (기본 작업 폴더 옆 <폴더>_가져가기_<날짜>.zip)')
    a = ap.parse_args()
    root = find_root(a.root or os.getcwd())
    out = lambda d, rc=0: (print(json.dumps(d, ensure_ascii=False, indent=1)), rc)[1]
    if a.cmd == 'message':
        print(message(root))
        return 0
    if a.cmd == 'remote':   # push 전: 원격 확인 · 원격이 있으면 그 저장소가 PRIVATE 인지
        rc, rs, err = remotes(root)
        if rc != 0:
            return out(dict(status='통과', note='git 저장소가 아니거나 원격이 없음 → gh repo create hr-agent --private --source . --push 로 새로 만듦', remotes=[], detail=err[:200]))
        if not rs:
            return out(dict(status='통과', note='원격 없음 → gh repo create hr-agent --private --source . --push 로 새로 만듦', remotes=[]))
        bad = []
        for r_ in rs:
            if not r_['repo']:
                continue
            vis, repo, url, e = visibility(root, r_['repo'])
            r_['visibility'] = vis or '(확인 못 함)'
            if vis != 'PRIVATE':
                bad.append(not_private(repo, vis, e, 'push 전 (git remote -v)'))
        if bad:
            return out(dict(bad[0], remotes=rs), EXIT_STOP)
        return out(dict(status='통과', remotes=rs, note='원격이 있으면 그 저장소에 push · GitHub 원격은 모두 PRIVATE' if any(r_['repo'] for r_ in rs) else '원격이 GitHub 이 아님 (사내 Git) · 사내 정책으로 비공개인지 사람에게 확인'))
    if a.cmd == 'visibility':   # push 뒤: 반드시 PRIVATE
        vis, repo, url, e = visibility(root, a.repo)
        if vis != 'PRIVATE':
            return out(not_private(repo, vis, e, 'push 뒤'), EXIT_STOP)
        return out(dict(status='통과', visibility=vis, repo=repo, url=url, note='비공개 (PRIVATE) 확인 · 주소를 활용 메모에 적게 안내'))
    if a.cmd == 'copy-plugin':
        dst = os.path.join(root, PLUGIN_DIR)
        if os.path.exists(os.path.join(dst, '.claude-plugin', 'plugin.json')):
            return out(dict(status='완료', note='플러그인 사본이 이미 있음 (그대로 둠)', path=PLUGIN_DIR, version=plugin_version(root)))
        src = a.plugin or os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
        if not os.path.exists(os.path.join(src, '.claude-plugin', 'plugin.json')):
            return out(dict(status='멈춤', kind='가져가기', what='플러그인 폴더를 찾지 못함', evidence=[src],
                            options=['--plugin 에 플러그인 폴더 (.claude-plugin/plugin.json 이 있는 곳)를 주기', '보류'], answer='확정 · 추가 확인 · 보류 · 이관'), EXIT_STOP)
        shutil.copytree(src, dst, ignore=shutil.ignore_patterns('__pycache__', '*.pyc', '.git'))
        return out(dict(status='완료', path=PLUGIN_DIR, version=plugin_version(root),
                        note='비공개 저장소에만 올린다 (교육 참가자 사내 사용 조건 · 공개 저장소 금지)'))
    r = check(root)
    tl = tools()
    if a.cmd == 'check':
        saved = has_commit(root)
        nxt = ('막음을 먼저 해결 (자료를 빼거나 _실데이터 자리로)' if r['block'] else
               ('git add -A → 커밋 → take_home.py remote (PRIVATE 확인) → git push' if saved and tl['git'] else
                'git init → 커밋 → gh repo create hr-agent --private --source . --push' if tl['git'] and tl['gh'] else
                'git 은 있고 gh 가 없음 → 사람에게 묻고 winget install --id GitHub.cli -e (관리자 창 「예」는 사람) → gh auth login 뒤 다시 · 급하면 `take_home.py zip`' if tl['git'] else
                'git 이 없음 → 사람에게 묻고 winget install --id Git.Git -e (관리자 창 「예」는 사람) 뒤 다시 · 급하면 `take_home.py zip`'))
        return out(dict(status='막음' if r['block'] else '통과', root=root.replace('\\', '/'), tools=tl, ignore_rules=r['rules'],
                        포함_파일_수=len(r['files']), 회사_폴더=r['folders'], 막음=r['block'], 확인_필요=r['warn'],
                        참고=r['info'], 플러그인_사본=r['plugin'], 저장소_이름_예=repo_name(root), 저장소_이름_대안=repo_alt(root), 공개=False, 이미_저장소=saved,
                        커밋_메시지=message(root).splitlines()[0], next=nxt), EXIT_STOP if r['block'] else 0)
    # zip
    if r['block']:
        return out(dict(status='멈춤', kind='가져가기', what='올리면 안 되는 자료가 있어 zip 을 만들지 않음', evidence=r['block'],
                        options=['그 자료를 빼거나 _실데이터 자리로 옮긴 뒤 다시', '보류'], answer='확정 · 추가 확인 · 보류 · 이관'), EXIT_STOP)
    dst = os.path.abspath(a.out or os.path.join(os.path.dirname(root), f'{os.path.basename(root)}_가져가기_{datetime.date.today():%Y%m%d}.zip'))
    skip = os.path.relpath(dst, root).replace('\\', '/')
    files = [f for f in r['files'] if f != skip]
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(os.path.join(root, f), os.path.basename(root) + '/' + f)
    return out(dict(status='완료', zip=dst.replace('\\', '/'), 파일_수=len(files), 확인_필요=r['warn'], 참고=r['info'],
                    note='.gitignore 로 뺀 파일 (실데이터 자리 · 잠금 파일 · 캐시 · 개인 설정)은 들어가지 않음 · 받은 곳에서 풀고 Code 탭에서 열어 「<회사> 인력계획 처음부터 돌려 줘」 (플러그인 없이 실행 · 고치려면 폴더 안 pwc-hr-harness/ 사본으로 설치)'))


if __name__ == '__main__':
    sys.exit(main())
