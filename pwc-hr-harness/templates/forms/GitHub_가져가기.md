# 내 hr-agent 저장소 (Day 1 첫 푸시 · 매일 저장 · Day 3 가져가기)

> 내 작업 폴더가 곧 **나의 HR-Agent** 입니다. 이 폴더를 내 계정의 **비공개 (Private)** 저장소 `hr-agent` 에 올리고, 3일 동안 고친 내용을 계속 커밋 · 푸시해 회사로 가져갑니다. 공개 저장소는 금지입니다 (플러그인은 교육 참가자 사내 사용 조건 · 하네스에 회사 조건이 들어 있음). GitHub 이 사내 정책으로 막혔으면 6절 zip 으로 가져갑니다.
>
> - **첫 버전** = `/hr-init` 이 깐 기본 팀 그대로 (모두 같음). Day 2 · 3 에 고친 만큼 회사마다 달라집니다
> - **플러그인** (`pwc-hr-harness`) 은 강사가 준 로컬 폴더에서 설치하고 GitHub 에 따로 올리지 않습니다. Day 3 가져가기 때 hr-agent 안에 사본을 넣습니다
>
> 작업은 Claude Desktop 의 Code 탭에서 합니다. 지시는 채팅창에, `claude plugin …` · `git` · `gh` 명령은 **terminal 패널**에 입력합니다. Desktop 채팅창의 `/plugin …` 명령은 동작하지 않습니다. 요즘은 명령을 직접 치기보다 **채팅창에 말로 시키는 것**이 기본입니다. 아래 표마다 명령과 「💬 Claude 에게」 문장을 같이 적었습니다.

## 0. 준비 (Day 1 전 · 처음 한 번)

**Day 1 부록 교안 「Windows 에서 Git · GitHub 처음 준비하기」** 를 따라 합니다. 직접 하는 것은 GitHub 가입과 로그인 한 줄뿐이고, 설치 · 서명 · 확인은 채팅창에 말로 시킵니다.

| 할 일 | 💬 Claude 에게 | 내가 할 것 |
|---|---|---|
| GitHub 가입 | | github.com 에서 가입 · 사용자명 메모 |
| Git · gh 설치 | 「git 이랑 gh 설치해 줘」 (Claude 가 `winget install` 을 돌림) | 허용 창 「허용」 · Windows 관리자 창 「예」 · 끝나면 Claude Desktop 재시작 |
| 로그인 | | terminal 패널에 `gh auth login -h github.com -p https -w -c` → 브라우저에서 Ctrl+V → Authorize |
| 서명 · 마무리 | 「내 GitHub 계정으로 git 이름 · 이메일 설정하고 마무리해 줘」 | 허용 창 「허용」 |

끝나면 한 번에 확인합니다.

| 확인 | 명령 (terminal 패널) | 💬 Claude 에게 |
|---|---|---|
| 설치 | `git --version` · `gh --version` | 「git 이랑 gh 가 설치돼 있는지 확인해 줘」 |
| 로그인 | `gh auth status` | 「GitHub 에 로그인돼 있는지 확인해 줘」 |
| 서명 | `git config --global user.name` · `user.email` | 「git 에 이름과 이메일이 등록돼 있는지 보여 줘」 |

- 로그인은 내 계정을 승인하는 단계라 **본인이 terminal 패널에서 직접** 합니다 (Claude 가 대신하지 않음) · 한 줄에 옵션을 다 넣어 묻는 질문이 거의 없고, 일회용 코드는 클립보드에 복사돼 있음

## 1. 저장소는 하나: hr-agent

| 때 | 하는 일 | 💬 Claude 에게 |
|---|---|---|
| Day 1 실습 (`/hr-init` 직후) | 저장소 만들고 첫 푸시 | 「이 폴더를 내 GitHub **비공개** 저장소 hr-agent 로 올려 줘」 |
| Day 1 · 2 · 3 마감 (고칠 때마다) | 커밋 · 푸시 | 「오늘 작업 커밋하고 푸시해 줘」 (메시지를 정하려면 「'Day 2 CP3 확정' 으로 커밋하고 푸시해 줘」) |
| Day 3 마무리 | 플러그인 사본을 넣고 푸시 · 동료 공유 안내 | 「회사로 가져갈 수 있게 마무리해 줘」 |

## 2. 무엇을 올리고 무엇을 안 올리나

| 올림 | 안 올림 (`.gitignore` · `/hr-init` 이 깖 · 견본 `양식/gitignore_견본`) |
|---|---|
| `CLAUDE.md` (변경 이력 포함) · `.claude/` 개정본 | 실데이터 자리: 이름 끝 `_실데이터` 폴더 (`data/자사_실데이터/` · `outputs/자사_실데이터/`) · `data/<회사>/_실데이터/` |
| `양식/` · 가상 데이터 `data/<자사>/` (가상 표시 파일 포함) · K-Co. 교육 데이터 | 엑셀 잠금 파일 `~$*` |
| `outputs/` (가상 결과 · 리포트) · `hr-viewer.html` | `__pycache__/` · `*.pyc` · `.DS_Store` · `Thumbs.db` |
| (Day 3) 플러그인 사본 `pwc-hr-harness/` | `.claude/settings.local.json` (PC 마다 다른 권한 기록) |

## 3. Claude 가 하는 순서 (`pwc-hr-harness` 스킬 모드 E)

| 순서 | Claude 가 하는 일 | 같은 일을 직접 할 때 (terminal 패널) |
|---|---|---|
| 1 | `git` · `gh` 가 있는지, 이미 저장소인지 봄 | `python -X utf8 .claude/hr-lib/take_home.py check` |
| 2 | (새 저장소일 때) **비공개**인지 묻고, 공개라고 하면 멈춤 | |
| 3 | 올라갈 파일에 실데이터 · 개인 식별 열이 없는지 봄 (있으면 멈춤 · `추가_` 자료는 가상인지 물음) | 같은 `check` |
| 4 | (Day 3) 플러그인 사본 넣기 | `take_home.py copy-plugin --plugin <플러그인 폴더>` |
| 5 | 커밋 (첫 커밋 = 「hr-agent 첫 커밋: 기본 팀 설치」 · 이후 = 「hr-agent 저장 <날짜>」) | `git add -A` → `git status --short` → `git commit -m "…"` |
| 6 | 원격이 없으면 새 비공개 저장소를 만들며 푸시 · 있으면 푸시 | `gh repo create hr-agent --private --source . --push` · 이후 `git push` |
| 7 | 올린 뒤 **PRIVATE** 인지 확인 (아니면 멈추고 물음) | `gh repo view --json visibility` |
| 8 | 저장소 주소와 이번 커밋 한 줄을 알려 줌 → 브라우저에서 「Private」 표시와 파일 확인 | |

- 같은 이름 저장소가 이미 있으면 `hr-agent-<회사>` 처럼 이름을 바꿉니다 (Claude 가 대안을 제안)
- 푸시가 거절되면 (다른 PC 에서 먼저 올렸을 때) Claude 가 `git pull --rebase` 뒤 다시 올립니다

## 4. gh 로그인이 풀렸을 때

- terminal 패널에서 `gh auth login -h github.com -p https -w -c` 를 다시 하고 (Day 1 부록 교안과 같은 한 줄), 채팅창에 같은 말을 한 번 더 합니다
- 그래도 안 되면 6절 zip 으로 가져갑니다

## 5. 회사에서 다시 쓰기 (운영자 · 사용자 · 편집자)

| 역할 | 누구 | 준비 | 하는 일 |
|---|---|---|---|
| **운영자** | 교육 이수자 (나) | 내 PC 에 플러그인 설치 (이미 됨) | 자사 PRD를 고치고 `/hr-customize` 질문 9 로 내 하네스에 반영 → 뷰어 「변경점」 확인 → 「커밋하고 푸시해 줘」 · 동료를 저장소에 초대 (github.com → 저장소 Settings → Collaborators → Add people) |
| **사용자** | 회사 동료 | Claude Desktop · Git · GitHub 계정 · 초대 수락 | 아래 1 · 2 · 3 · 플러그인 없이 인력계획 실행 (받은 하네스에 회사 규칙이 들어 있어 회사 기준으로 돎) · 고칠 점은 운영자에게 요청 |
| **편집자** | 전파교육을 받은 동료 | 사용자 준비 + 폴더 안 사본으로 플러그인 설치 (아래 4) | 운영자와 같이 `/hr-customize` 로 고침 |

| 순서 | 명령 (terminal 패널) | 💬 Claude 에게 |
|---|---|---|
| 1 받기 | `git clone https://github.com/<운영자계정>/hr-agent.git` | (받기 전이라 명령으로) |
| 2 열기 | Code 탭 → 「새로 만들기」 → 「이 컴퓨터」 → 「폴더 열기…」 → 받은 `hr-agent` 폴더 (영문 화면: New session → Local) | |
| 3 실행 | | 「K-Co 인력계획 처음부터 돌려 줘」 · 「<회사> 인력계획 이어서 해 줘」 (`/hr-init` 은 하지 않음) |
| 최신 받기 | `git pull` | 「운영자가 올린 최신 버전 받아 줘」 |
| 4 편집자 | `claude plugin marketplace add ./pwc-hr-harness` → `claude plugin install pwc-hr-harness@pwc-hr-marketplace` | (설치는 terminal 패널에서 · 새 세션부터 `/hr-customize` 사용) |

- 고칠 때마다 커밋 한 번 (CLAUDE.md 변경 이력 행은 질문 9 가 씀 · Git 기록과 하네스 기록이 같이 남음)
- 실데이터로 바꿀 때는 `양식/실데이터_전환_점검표.md`
- 사내 Git (GitHub Enterprise 등) 이 있으면 주소만 바뀌고 절차는 같습니다

## 6. 막히면: zip 으로 가져가기

「GitHub 이 막혀 있어. zip 으로 만들어 줘」 → `take_home.py zip` (같은 점검 뒤 `.gitignore` 로 뺀 파일 없이 작업 폴더 옆에 `<폴더>_가져가기_<날짜>.zip`). 사내 Git 이 있으면 그 주소로 같은 순서를 밟습니다 (3절 6번만 `git remote add origin <사내 주소>` → `git push -u origin HEAD`).
