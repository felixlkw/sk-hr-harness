# sk-hr-harness

mySUNI 「AI Powered Workforce Planning」 과정 교육용 배포 저장소입니다.

| 폴더 | 내용 |
|---|---|
| `pwc-hr-harness/` | 플러그인 새 판 (0.13.1) · 인력계획 기본 팀 + 새 하네스 만들기 (`harness` · `evolve`) |
| `교안/` | New3강 교안 PDF (준비되면 올라옵니다) |

## 아침에 할 일 · 플러그인 바꿔 끼우기 (약 5분)

Day 1 에 교육장 PC `C:\PwC\pwc-hr-harness` 폴더로 플러그인을 설치했습니다. 이 폴더를 새 판으로 바꾸고 업데이트합니다.

1. **내려받기:** 이 페이지 위쪽 초록색 **Code** 단추 → **Download ZIP** → `sk-hr-harness-main.zip`
2. **압축 풀기:** 받은 zip 을 마우스 오른쪽 단추 → **압축 풀기** (`sk-hr-harness-main` 폴더가 생김)
3. **바꿔 끼우기:** `C:\PwC\pwc-hr-harness` 폴더를 지우고, 압축을 푼 폴더 안의 `pwc-hr-harness` 폴더를 `C:\PwC\` 에 그대로 복사
   - 결과: `C:\PwC\pwc-hr-harness\NOTICE.md` 가 보이면 맞음
4. **업데이트:** Claude Desktop **Code** 탭 → **terminal 패널**에서 두 줄을 차례로 붙여 넣음

   ```
   claude plugin marketplace update pwc-hr-marketplace
   claude plugin update pwc-hr-harness@pwc-hr-marketplace
   ```

   `updated from 0.12.2 to 0.13.1 … Restart to apply changes.` 가 나오면 맞음
   (terminal 이 어려우면 채팅에 「위 두 명령을 terminal 에서 실행해 줘」라고 해도 됩니다)
5. **다시 시작:** Claude Desktop 을 완전히 끄고 다시 켬 → 새 세션
6. **확인:** terminal 패널에서 `claude plugin list` → `pwc-hr-harness … 0.13.1`

### 막히면

| 화면 | 할 일 |
|---|---|
| `Marketplace 'pwc-hr-marketplace' not found` | 폴더 위치가 다름 · `C:\PwC` 에서 `claude plugin marketplace add ./pwc-hr-harness` → `claude plugin install pwc-hr-harness@pwc-hr-marketplace` |
| `already at the latest version` | 3번 폴더가 옛 판 그대로임 · `C:\PwC\pwc-hr-harness\NOTICE.md` 가 있는지 확인 |
| 폴더를 지울 수 없음 | Claude Desktop 을 끄고 다시 지움 |

Day 1~2 에 만든 내 작업 폴더 (hr-agent)는 건드리지 않습니다. 그대로 남습니다.

## 이용 조건

- `pwc-hr-harness/skills/harness/` · `skills/evolve/` 는 [revfactory/harness](https://github.com/revfactory/harness) v2.1.0 (Apache-2.0)에서 가져왔습니다 · 자세한 내용 `pwc-hr-harness/NOTICE.md`
- 그 밖의 파일은 과정 참가자의 교육 실습 · 소속 회사 안 업무 활용 · 전파교육용입니다. 무단 재배포 · 상업적 이용을 금지합니다
- 들어 있는 회사 자료 (K-Co. · E-Co.)는 모두 가상입니다
