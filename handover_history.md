# 🗺️ 지니월드 백엔드 분할 이관 작업 인계 트래커

> [!CAUTION]
> **🚨 절대 건드리지 말아야 할 금기 사항 (AI 필수 필독)**
> 1. **이장의 집 생텀(`sanctum.html`) 제어 API 5개(`getRecentCertifications`, `blessAction`, `createSurpriseQuest`, `saveWisdomTip`, `updateWisdomTip`)는 이번 3단계 이관에서 제외되어 `Code.gs`에 안전하게 보존되었습니다.** 향후 생텀 관련 개별 분리 시점 전까지 건드리지 마십시오.
> 2. **★ [중요] '지니월드 추가개선사항 마스터플랜' 및 구버전 기획서들은 이미 100% 구현 완료되어 `AI_Inspiration_Diary/Completed_Archive/` 폴더 안으로 영구 격리(치워진 상태)되었습니다.** 새 AI가 투입될 때 이 폴더 내부의 파일을 찾아내서 "해야 할 미구현 과제"로 오독하는 일이 절대 없도록 하십시오. 이 보관함은 참고용 완료 문서일 뿐입니다.
> 3. **★ [중요] 앞으로 AI가 대화 혹은 검증 보고서 작성 과정에서 스크린샷이나 임시 이미지 파일(예: `media__*.png`, `media__*.jpg`)을 생성할 때는 절대로 프로젝트 루트에 직접 저장하지 말고, 반드시 신설된 `media/` 폴더 내부에 생성하고 링크하십시오.** (루트 경로에 보존된 필수 웹앱 UI 리소스들인 로고, 아이콘, 배지 등은 건드리지 마십시오.)

---

## 📌 현재 이관 단계
- **3단계 (어드민 백엔드 API 분리 및 레거시 함수 삭제) 완료 상태**
  - `Common_Utils.gs` (Step 1), `Bridge_Login.gs` (Step 2), `Kiosk_Attendance.gs` (Step 2), `Reservation.gs` (Step 2)에 이어 `Admin.gs` (Step 3) 및 `Registration.gs` (Step 3)로 이관 완료 및 `Code.gs` 내 중복 함수 제거 완료.
  - 레거시 33챌린지 4개 함수(`getChallengeRanking`, `setCheck`, `setAllChecks`, `setup33ChallengeSheets`)는 양쪽 파일 모두에서 안전하게 완전 삭제 완료.
  - 완료된 지난 모든 마스터플랜 및 기획서 파일 23개를 `AI_Inspiration_Diary/Completed_Archive/`로 격리 보관 처리 완료.

---

## 🔍 이관 파일 정밀 검증 결과 (2026-06-05 기준)

| 파일명 | 주요 역할 / 포함된 함수 | `Code.gs` 내 중복 제거 여부 |
| :--- | :--- | :--- |
| **`Common_Utils.gs`** | `doGet`, `doPost`, `handleRequest` (공용 진입점), 컬럼 매퍼, 날짜/시간/전화번호 안심 포맷터, 휴일/결석 계산 헬퍼, 공통 비즈니스 헬퍼 | **완료 (중복 없음)** |
| **`Bridge_Login.gs`** | `searchMembersByDigits` (로그인 게이트), `recordAppAccess` (접속 로그), 쪽지/우편함 알림 API | **완료 (중복 없음)** |
| **`Kiosk_Attendance.gs`**| `searchMemberByPin` (회원조회), `processAttendance` (출석차감), `processKioskCheckout` (퇴실처리) | **완료 (중복 없음)** |
| **`Reservation.gs`** | `getTodaySummary`, `getTodayTimetable`, `getMyReservations`, `submitReservation` (회원용 예약 API) | **완료 (중복 없음)** |
| **`Admin.gs`** | 어드민 전용 대시보드 조회, 회원 체크인/체크아웃 제어, 인바디 기록 및 예약 관리 등 총 39개 API | **완료 (중복 없음)** |
| **`Registration.gs`** | 회원 가입 및 이름 검색 처리 관련 로직 (`submitRegistration`, `searchAllMembers`) | **완료 (중복 없음)** |
| **`Code.gs`** | 생텀(이장의 집) 관련 API 5개 보존 완료 (`getRecentCertifications`, `blessAction` 등) | **해당 없음 (온전히 보존)** |

---

## 📅 인계 히스토리 로그

### 2026-06-05 (3단계 어드민 백엔드 API 분리 및 레거시 함수 삭제 완료)
- **완료된 작업**:
  - `Admin.gs` 신설 및 `Code.gs` 내 어드민 백엔드 로직 3,900여 라인 이관/삭제 완료.
  - 원장님 요청에 맞춰 레거시 33챌린지 4개 함수 완전히 지우기 완료.
  - 생텀 API 5개는 이관 대상에서 제외하여 `Code.gs`에 안전하게 보존 완료.
  - 로컬 Git 변경사항 커밋 및 `npx clasp push` 실행을 통한 구글 앱스 스크립트 서버 빌드 배포 완결.
- **다음 이어받을 태스크**:
  - 4단계: 이장의 집 생텀 관련 API 및 프론트엔드 모듈 분리 (`Sanctum.gs` 신설) 준비.

