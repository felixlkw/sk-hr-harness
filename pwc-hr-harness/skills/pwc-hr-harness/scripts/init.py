# -*- coding: utf-8 -*-
"""pwc-hr-harness · 기본 HR 인력계획 팀 설치 (/hr-init).

사용: python -X utf8 <플러그인>/skills/pwc-hr-harness/scripts/init.py [--target 작업폴더] [--force] [--check]

1. 환경 점검: Python 3.9+ · openpyxl
2. 현황 감사: 작업 폴더에 이미 하네스가 있는지 (있으면 멈추고 알림 · --force 로 덮어쓰기)
3. 기본 팀 복사: CLAUDE.md · .claude/agents (5) · .claude/skills (7) · .claude/hr-lib (계산 · 뷰어 · 양식)
4. 기본 사본 보관: .claude/hr-lib/_standard/ (자사 버전 「변경점」 비교 기준 · 고치지 않음)
5. 폴더: data/ · outputs/ · 양식/ (PRD · 기획안 · 데이터 사전 · 로그 · Day 3 양식 · K-Co. 조건 예시) · 회사별 결정로그 · 검증로그
6. .gitignore: 없으면 견본 (templates/forms/gitignore_견본) 을 깖 (실데이터 자리 · 잠금 파일 · 캐시 · 개인 설정 제외 · GitHub 가져가기)
7. 뷰어 생성: hr-viewer.html
"""
import os, sys, json, shutil, datetime, argparse, subprocess, glob

PLUGIN = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
STD = os.path.join(PLUGIN, 'templates', 'standard-team')
FORMS = os.path.join(PLUGIN, 'templates', 'forms')
GITIGNORE = 'gitignore_견본'


def version():
    try:
        with open(os.path.join(PLUGIN, '.claude-plugin', 'plugin.json'), encoding='utf-8') as f:
            return json.load(f).get('version', '?')
    except Exception:
        return '?'


def env_check():
    out = dict(python=sys.version.split()[0], ok=sys.version_info >= (3, 9))
    try:
        import openpyxl  # noqa
        out['openpyxl'] = openpyxl.__version__
    except ImportError:
        out['openpyxl'] = None
        out['ok'] = False
        out['fix'] = f'"{sys.executable}" -m pip install openpyxl python-docx'
    try:   # 워드 리포트 · 내보내기용 (없어도 계산 · 엑셀은 됨 · 2026-11-16)
        import docx  # noqa
        out['python-docx'] = 'ok'
    except ImportError:
        out['python-docx'] = None
        out.setdefault('fix', f'"{sys.executable}" -m pip install python-docx')
        out['note'] = 'python-docx 가 없으면 워드 리포트 · 워드 내보내기만 건너뜀 (계산 · 엑셀은 그대로)'
    return out


def audit(t):
    agents = sorted(os.path.basename(p) for p in glob.glob(os.path.join(t, '.claude', 'agents', '*.md')))
    skills = sorted(os.path.basename(os.path.dirname(p)) for p in glob.glob(os.path.join(t, '.claude', 'skills', '*', 'SKILL.md')))
    has_std = os.path.isdir(os.path.join(t, '.claude', 'hr-lib', '_standard'))
    mode = '신규 구축' if not any(a.startswith('wfp-') for a in agents) else ('자사 버전 진화' if has_std else '기존 하네스 (기본 사본 없음)')
    datas = sorted(d for d in os.listdir(os.path.join(t, 'data')) if os.path.isdir(os.path.join(t, 'data', d))) if os.path.isdir(os.path.join(t, 'data')) else []
    return dict(mode=mode, agents=agents, skills=skills, data=datas, claude_md=os.path.exists(os.path.join(t, 'CLAUDE.md')))


def _blank(path):
    """양식 그대로인가 (hrlib.is_blank_form 과 같은 기준 · 작업 폴더 hrlib 을 씀)."""
    sys.path.insert(0, os.path.join(STD, '.claude', 'hr-lib'))
    try:
        import hrlib as H
        with open(path, encoding='utf-8') as f:
            return H.is_blank_form(f.read())
    except Exception:
        return os.path.getsize(path) < 400
    finally:
        sys.path.pop(0)


def copy_tree(src, dst):
    for root, dirs, files in os.walk(src):
        rel = os.path.relpath(root, src)
        if '__pycache__' in rel:
            continue
        os.makedirs(os.path.join(dst, rel), exist_ok=True)
        for f in files:
            if f.endswith('.pyc'):
                continue
            shutil.copy2(os.path.join(root, f), os.path.join(dst, rel, f))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--target', default=os.getcwd())
    ap.add_argument('--force', action='store_true')
    ap.add_argument('--check', action='store_true', help='설치하지 않고 환경 · 현황만 봄')
    a = ap.parse_args()
    t = os.path.abspath(a.target)
    env = env_check()
    au = audit(t)
    if a.check or not env['ok']:
        print(json.dumps(dict(status='점검', target=t, env=env, audit=au), ensure_ascii=False, indent=1))
        sys.exit(0 if env['ok'] else 2)
    if au['mode'] != '신규 구축' and not a.force:
        print(json.dumps(dict(status='멈춤', reason=f'이미 하네스가 있음 ({au["mode"]}) · 덮어쓰려면 --force · 자사 버전 작업은 /hr-customize',
                              audit=au), ensure_ascii=False, indent=1))
        sys.exit(3)
    today = datetime.date.today().isoformat()
    ver = version()
    with open(os.path.join(STD, 'CLAUDE.md'), encoding='utf-8') as f:
        cm = f.read().replace('{today}', today).replace('{version}', ver).replace('{target}', '계획연도')
    cpath = os.path.join(t, 'CLAUDE.md')
    if os.path.exists(cpath) and not a.force:
        with open(cpath, encoding='utf-8') as f:
            old = f.read()
        if '## 하네스: 전략적 인력계획' not in old:
            cm = old.rstrip() + '\n\n' + cm
    with open(cpath, 'w', encoding='utf-8') as f:
        f.write(cm)
    copy_tree(os.path.join(STD, '.claude'), os.path.join(t, '.claude'))
    # 기준 스냅샷 · 회귀를 작업 폴더에서 말로 부를 수 있게 도구 사본을 둔다 (원본은 플러그인)
    # 가상 데이터 만들기 (/hr-customize B2 · 자사 파일이 없을 때)도 같은 방식으로 사본을 둔다
    for tool in ('customize.py', 'virtual_data.py', 'take_home.py', 'adapt.py'):
        shutil.copy2(os.path.join(os.path.dirname(os.path.abspath(__file__)), tool), os.path.join(t, '.claude', 'hr-lib', tool))
    # 기본 사본 (비교 기준)
    snap = os.path.join(t, '.claude', 'hr-lib', '_standard')
    if os.path.isdir(snap):
        shutil.rmtree(snap)
    os.makedirs(snap)
    shutil.copy2(cpath, os.path.join(snap, 'CLAUDE.md'))
    for sub in ('agents', 'skills'):
        copy_tree(os.path.join(t, '.claude', sub), os.path.join(snap, '.claude', sub))
    # config.json 도 보관 (설치 직후 lock · diff · 뷰어가 「추가」로 세지 않게) · hrlib.py 는 lock 의 엔진 변경 확인 기준
    os.makedirs(os.path.join(snap, '.claude', 'hr-lib'), exist_ok=True)
    for f in ('config.json', 'hrlib.py'):
        shutil.copy2(os.path.join(t, '.claude', 'hr-lib', f), os.path.join(snap, '.claude', 'hr-lib', f))
    # 법전 사본 (references/codex-*.md)은 작업 폴더에 깔지 않는다 (방법론 원문 배포 범위 · 판단 규칙은 각 Skill 본문에 있음)
    with open(os.path.join(snap, 'README.md'), 'w', encoding='utf-8') as f:
        f.write(f'# 기본 사본 (pwc-hr-harness {ver} · {today})\n\n- /hr-init 이 설치한 그대로의 사본. 고치지 않는다. 뷰어 「기본 대비 변경점」과 /hr-customize 잠금 검사의 기준이다\n')
    # 폴더 · 양식
    for d in ('data', 'outputs', '양식'):
        os.makedirs(os.path.join(t, d), exist_ok=True)
    for f in glob.glob(os.path.join(FORMS, '*.md')) + glob.glob(os.path.join(FORMS, '*.json')) + [os.path.join(FORMS, GITIGNORE)]:
        dst = os.path.join(t, '양식', os.path.basename(f))
        if os.path.exists(f) and not os.path.exists(dst):
            shutil.copy2(f, dst)
    gi = os.path.join(t, '.gitignore')
    gitignore = '있던 것 그대로' if os.path.exists(gi) else '견본 설치'
    if not os.path.exists(gi) and os.path.exists(os.path.join(FORMS, GITIGNORE)):
        shutil.copy2(os.path.join(FORMS, GITIGNORE), gi)
    datas = sorted(d for d in os.listdir(os.path.join(t, 'data')) if os.path.isdir(os.path.join(t, 'data', d)))
    for co in datas:
        o = os.path.join(t, 'outputs', co)
        os.makedirs(o, exist_ok=True)
        for f in ('결정로그.md', '검증로그.md'):
            dst = os.path.join(o, f)
            if not os.path.exists(dst) or _blank(dst):   # 없거나 양식 그대로 (행이 비어 있음) → 플러그인 양식으로
                shutil.copy2(os.path.join(t, '.claude', 'hr-lib', 'forms', f), dst)
    r = subprocess.run([sys.executable, '-X', 'utf8', os.path.join(t, '.claude', 'hr-lib', 'render.py'), '--root', t],
                       capture_output=True, text=True, encoding='utf-8')
    print(json.dumps(dict(status='설치 완료', target=t, version=ver, env=env, mode=au['mode'], data=datas,
                          installed=dict(agents=sorted(os.listdir(os.path.join(t, '.claude', 'agents'))),
                                         skills=sorted(os.listdir(os.path.join(t, '.claude', 'skills')))),
                          viewer=os.path.join(t, 'hr-viewer.html').replace('\\', '/'), viewer_log=r.stdout.strip()[-300:], gitignore=gitignore,
                          next='Claude Code 를 이 폴더에서 다시 열고 /agents 로 담당 5명 확인 → 「K-Co 사업계획 분석부터 시작해 줘」'),
                     ensure_ascii=False, indent=1))


if __name__ == '__main__':
    main()
