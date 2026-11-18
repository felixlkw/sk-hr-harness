---
description: "인력계획 4단계를 리드 순서대로 진행한다 (예: /hr-run K-co · /hr-run 자사 ②)"
---

`wfp-lead` 스킬 (작업 폴더 `.claude/skills/wfp-lead/SKILL.md`)을 읽고 리드로서 진행한다.

- 인자: 회사 이름과 (선택) 시작 단계. 예: `K-co` · `E-co` · `자사 ②`
- 체크포인트 · 예외 · 관문에서는 반드시 멈추고 4요소를 보여 준 뒤 사람 답을 받는다
- 단계가 끝날 때와 CP · 예외 · 관문에서 멈출 때마다 `python -X utf8 .claude/hr-lib/render.py` 로 뷰어를 다시 만들고 (브라우저로 열림) 볼 탭을 알려 준다
- 작업 폴더에 `.claude/skills/wfp-lead` 가 없으면 먼저 `/hr-init` 을 안내한다

인자: $ARGUMENTS
