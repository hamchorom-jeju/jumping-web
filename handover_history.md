# 🗺️ 지니월드 백엔드 분할 이관 작업 인계 트래커

> [!CAUTION]
> **🚨 절대 건드리지 말아야 할 금기 사항 (AI 필수 필독)**
> 1. **★ [중요] '지니월드 추가개선사항 마스터플랜' 및 구버전 기획서들은 이미 100% 구현 완료되어 `AI_Inspiration_Diary/Completed_Archive/` 폴더 안으로 영구 격리(치워진 상태)되었습니다.** 새 AI가 투입될 때 이 폴더 내부의 파일을 찾아내서 "해야 할 미구현 과제"로 오독하는 일이 절대 없도록 하십시오. 이 보관함은 참고용 완료 문서일 뿐입니다.
> 2. **★ [중요] 앞으로 AI가 대화 혹은 검증 보고서 작성 과정에서 스크린샷이나 임시 이미지 파일(예: `media__*.png`, `media__*.jpg`)을 생성할 때는 절대로 프로젝트 루트에 직접 저장하지 말고, 반드시 신설된 `media/` 폴더 내부에 생성하고 링크하십시오.** (루트 경로에 보존된 필수 웹앱 UI 리소스들인 로고, 아이콘, 배지 등은 건드리지 마십시오.)
> 3. **★ [초핵심] 세션이 교대되어 부팅된 새로운 AI는 다른 어떤 수정 도구를 실행하기 전에 본 `handover_history.md` 파일과 `AI_협업_프로토콜_지침.md` 파일을 무조건 `view_file`로 먼저 읽어들여 직전 세션과의 정렬을 완료해야 합니다.** (이를 무시하고 멋대로 reset을 실행하는 등의 독단적 파괴 행동을 절대 금합니다.)

---

## 📌 현재 이관 단계
- **9부 & 10부 (공용 유틸 통합 및 테라피 예약 기능 재분배 완결) 완료 상태**
  - **`Common_Utils.gs` 완전 해체 및 `Code.gs`로의 재통합**: 역할이 겹치던 공용 유틸 1,300줄을 `Code.gs` 하단에 병합 완료하고 `Common_Utils.gs`는 완전 삭제 처리되었습니다.
  - **`Reservation.gs` 완전 해체 및 재분배**: `getMemberIDList` 조회 API는 `Code.gs` 하단에 공용 헬퍼로 남기고, 나머지 5개 예약 관리 핵심 API는 `Bridge_Login.gs` 하단으로 병합 완료 후 `Reservation.gs` 원본 파일은 삭제하였습니다.
  - **날짜/전화번호 정규화 포맷터 정리 완료**: 낡은 래퍼 및 중복 정의들(`formatPhoneSafely`, `formatDateSafely`, `formatPhoneForSms`)을 `Code.gs` 본문에서 완전히 지우고, 고성능 단일 헬퍼(`normalizePhoneDigits`, `normalizeDateStr`, `formatPhoneNumber`)로 완벽 단일화했습니다.

---

## 🔍 이관 파일 정밀 검증 결과 (2026-06-05 기준)

| 파일명 | 주요 역할 / 포함된 함수 | `Code.gs` 내 중복 제거 여부 |
| :--- | :--- | :--- |
| **`Common_Utils.gs`** | **[해체 완료]** `Code.gs` 하단 유틸 영역으로 100% 병합 및 통합 완료. 파일 삭제됨. | **통합 완료** |
| **`Reservation.gs`** | **[해체 완료]** 공용 회원조회(`getMemberIDList`)는 `Code.gs` 공용부로, 나머지 예약 5개 API는 `Bridge_Login.gs` 하단으로 병합. 파일 삭제됨. | **분배 통합 완료** |
| **`Bridge_Login.gs`** | `searchMembersByDigits` (로그인 게이트), `recordAppAccess` (접속 로그), 쪽지/우편함 알림 API + **[추가] 테라피 예약 관리 전용 API 5개** | **완료 (중복 없음)** |
| **`Kiosk_Attendance.gs`**| `searchMemberByPin` (회원조회), `processAttendance` (출석차감) | **완료 (중복 없음)** |
| **`Admin.gs`** | 어드민 전용 대시보드 조회, 회원 체크인/체크아웃 제어, 인바디 기록, 예약 관리 및 키오스크 퇴실 처리(`processKioskCheckout`) 등 총 40개 API | **완료 (중복 없음)** |
| **`Registration.gs`** | 회원 가입 및 이름 검색 처리 관련 로직 (`submitRegistration`, `searchAllMembers`) | **완료 (중복 없음)** |
| **`Sanctum.gs`** | 생텀(이장의 집) 관련 공지, 날씨/테마(기후), 지식 꿀팁 및 돌발 퀘스트/인증 승인 전체 통제 API (총 33개) | **완료 (중복 없음)** |
| **`Code.gs`** | **[통합 코어]** 사용자 대시보드 데이터 조회, 커뮤니티, 자동 정산 배치 스케줄러 본체 + **도 doGet/doPost 진입점 라우터, 컬럼 매퍼, 정규화 유틸리티 일체 내장** | **해당 없음 (최종 통합본)** |

---

## 📅 인계 히스토리 로그

### 2026-06-05 (10부 최종 완결: 롤백 복구 및 공용화 리팩토링)
- **완료된 작업**:
  - `Common_Utils.gs`를 `Code.gs` 하단으로 병합 완료 및 `Common_Utils.gs` 파일 삭제 완료.
  - `Reservation.gs` 해체 완료: `getMemberIDList` 조회 API -> `Code.gs` 병합, 예약 관리 API 5개 -> `Bridge_Login.gs` 병합 완료 및 `Reservation.gs` 파일 삭제.
  - `Code.gs` 내 중복되던 낡은 포맷터 3개 및 중복 검색 함수 1개(`searchMemberRegistryByDigits` 낡은 버전) 삭제 완료.
  - `Bridge_Login.gs`와 `Kiosk_Attendance.gs` 내 4자리 검색 루프를 공용 `searchMemberRegistryByDigits` 호출 구조로 리팩토링 완료.
  - `Code.gs` 내 `getMemberIDList` 함수 닫는 중괄호 `}` 누락 패치로 구문 오류 해결.
  - `npx clasp push` 실행 완료: 22개 파일 중복 충돌 없이 구글 앱스 스크립트 서버 배포 최종 빌드 검증 통과.
- **다음 이어받을 태스크 (완결)**:
  - 현재 아키텍처 정리 및 공용화 리팩토링이 최종 완료된 상태이므로, 추가적인 파일 분리 없이 정상 운영 모니터링 가능.
