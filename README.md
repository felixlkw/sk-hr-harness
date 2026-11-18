# sk-hr-harness

mySUNI 「AI Powered Workforce Planning」 과정 Day 3 (New3강 「내가 필요한 HR-Harness 만들기」) 교육용 배포 저장소입니다.

| 폴더 | 내용 |
|---|---|
| `pwc-hr-harness/` | 플러그인 새 판 (0.13.3) · 인력계획 기본 팀 + 새 하네스 만들기 (`harness`) · 하네스 개선 (`evolve`) · 자사 파일 정리 Agent 틀 · 한 장 정리 |
| `교안/` | New3강 교안 PDF (36장) · 실습 예시 자료 (`HR_harness_mock_files/` · 모두 가상) · 자료 쓰는 법은 [`교안/README.md`](교안/README.md) |

## 오늘 하루

| 시간 | 내용 | 손에 남는 것 |
|---|---|---|
| 09:30~10:30 | 1부 · 하네스 이해와 주제 선정 (하네스 구조 · 메타하네스 · 플러그인 교체 · 하네스가 필요한 업무) | 새 버전 플러그인 · 내 업무 후보 |
| 10:40~12:00 · 13:20~15:00 | 2부 · 내 업무 하네스 구축 (과제 정의 · 데이터 정비 · 요청 · 구축 · 사람 확인 · 개선 · 실습) | 내 하네스 1개 · 변경 이력 |
| 15:00~16:00 | 3부 · 표준 인력계획 팀 자사화 (자사 파일 정리 Agent 추가 · 원본 정리) | 정리 Agent · 기본 틀로 정리된 6종 |
| 16:00~17:30 | 4부 · 공유와 피드백 · 한 장 정리 · 가져가기 | 피드백 반영한 내 하네스 · 한 장 정리 (수료 확인) |

## 아침에 할 일 · 플러그인 바꿔 끼우기 (약 5분 · 교안 8쪽)

Day 1 에 교육장 PC `C:\PwC\pwc-hr-harness` 폴더로 플러그인을 설치했습니다. 이 폴더를 새 판으로 바꾸고 업데이트합니다.

1. **내려받기:** 이 페이지 위쪽 초록색 **Code** 단추 → **Download ZIP** → `sk-hr-harness-main.zip`
2. **압축 풀기:** 받은 zip 을 마우스 오른쪽 단추 → **모두 압축 풀기** (`sk-hr-harness-main` 폴더가 생김 · 같은 이름이 한 번 더 겹치면 그 안으로 들어감)
3. **바꿔 끼우기:** `C:\PwC\pwc-hr-harness` 폴더를 지우고, 압축을 푼 폴더 안의 `pwc-hr-harness` 폴더를 `C:\PwC\` 에 그대로 복사
   - 결과: `C:\PwC\pwc-hr-harness\NOTICE.md` 가 보이면 맞음
4. **업데이트:** Claude 채팅에 「플러그인을 새 버전으로 업데이트해 줘」 라고 요청함 · 직접 하려면 Claude Desktop **Code** 탭 → **terminal 패널**에서 두 줄을 차례로 붙여 넣음

   ```
   claude plugin marketplace update pwc-hr-marketplace
   claude plugin update pwc-hr-harness@pwc-hr-marketplace
   ```

   `updated … 0.13.3 … Restart to apply changes.` 가 나오면 맞음
5. **다시 시작:** Claude Desktop 을 완전히 끄고 다시 켬 → 새 세션
6. **확인:** 플러그인 목록 (또는 terminal 패널 `claude plugin list`) 에서 `pwc-hr-harness … 0.13.3`

### 막히면

| 화면 | 할 일 |
|---|---|
| `Marketplace 'pwc-hr-marketplace' not found` | 폴더 위치가 다름 · `C:\PwC` 에서 `claude plugin marketplace add ./pwc-hr-harness` → `claude plugin install pwc-hr-harness@pwc-hr-marketplace` |
| `already at the latest version` | 3번 폴더가 옛 판 그대로임 · `C:\PwC\pwc-hr-harness\NOTICE.md` 가 있는지 확인 |
| 폴더를 지울 수 없음 | Claude Desktop 을 끄고 다시 지움 |

Day 1~2 에 만든 내 작업 폴더 (hr-agent)는 건드리지 않습니다. 그대로 남습니다.

## 실습 자료 (2부 · 3부)

압축을 푼 `sk-hr-harness-main → 교안 → HR_harness_mock_files` 안에 있습니다.

- **2부 · 내 자료가 없을 때:** `예시1_인력계획` · `예시2_월간인력보고` · `예시3_교육결과보고` 중 하나를 hr-agent **밖** 새 폴더로 복사하고, 그 폴더를 Claude Desktop 에서 연 뒤 교안 17쪽의 요청 문장을 그대로 입력합니다
- **3부:** `E-co_원본` 폴더를 Day 1~2 의 `hr-agent\data\` 안에 복사합니다 (`hr-agent\data\E-co_원본\`)

## 수업이 끝나면 · 내 저장소로 가져가기 (교안 34~35쪽)

플러그인과 하네스는 이 공개 저장소가 아니라 **각자의 GitHub 비공개 저장소**에 두고 회사에서 이어서 씁니다.

1. hr-agent 폴더 세션에서 「새 버전 플러그인 사본을 hr-agent 안에 넣어 줘」 → 교육장 PC 의 플러그인이 `hr-agent\pwc-hr-harness\` 로 복사됨
2. 같은 세션에서 「오늘 작업 커밋하고 푸시해 줘」 → 정리 Agent 와 플러그인 사본이 내 hr-agent 저장소에 올라감
3. 2부에서 만든 하네스 폴더 세션에서 「이 폴더를 내 GitHub 비공개 저장소 ○○-harness 로 새로 만들어 올려 줘」
4. 회사 PC 에서는 두 저장소를 내려받고 「hr-agent 안 플러그인 사본으로 설치해 줘」 라고 요청하면 그대로 이어서 씀

## 이용 조건

- `pwc-hr-harness/skills/harness/` · `skills/evolve/` 는 [revfactory/harness](https://github.com/revfactory/harness) v2.1.0 (Apache-2.0)에서 가져왔습니다 · 자세한 내용 `pwc-hr-harness/NOTICE.md`
- 그 밖의 파일은 과정 참가자의 교육 실습 · 소속 회사 안 업무 활용 · 전파교육용입니다. 무단 재배포 · 상업적 이용을 금지합니다
- 들어 있는 회사 자료 (K-Co. · E-Co. · M-Co.)는 모두 가상입니다
