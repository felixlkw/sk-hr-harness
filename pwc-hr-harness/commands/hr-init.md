---
description: 지금 폴더에 기본 HR 인력계획 팀(리드 · 담당 5 · Skill 7 · 뷰어)을 설치한다
allowed-tools: Bash, Read, Glob
---

`pwc-hr-harness` 스킬의 **모드 A (기본 팀 설치)** 를 따른다.

1. 점검: `python -X utf8 "${CLAUDE_PLUGIN_ROOT}/skills/pwc-hr-harness/scripts/init.py" --check` (Mac 은 `python3`)
   - openpyxl 이 없으면 안내된 설치 명령을 사람에게 보여 주고 허락을 받은 뒤 실행한다 (python-docx 도 없으면 같은 명령에 함께 · 워드 리포트용 · 없어도 계산은 됨)
2. 설치: `python -X utf8 "${CLAUDE_PLUGIN_ROOT}/skills/pwc-hr-harness/scripts/init.py"` (이미 있으면 멈춤 → 덮어쓸지 묻고 `--force`)
3. 결과를 짧게 알린다: 설치된 담당 5 · Skill 7 · `data/` 아래 회사 폴더 · `hr-viewer.html` 경로 (「하네스」 탭에서 팀 구성도를 보라고 안내)
4. 마지막에: 「Claude Code 를 이 폴더에서 다시 열면 `/agents` 에 담당 5명이 보입니다. 이 폴더가 나의 HR-Agent 입니다. 「이 폴더를 내 GitHub 비공개 저장소 hr-agent 로 올려 줘」라고 말하면 첫 버전을 저장합니다 (모드 E). 그다음 「K-Co 사업계획 분석부터 시작해 줘」」

$ARGUMENTS
