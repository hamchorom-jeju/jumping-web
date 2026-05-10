# 🏃‍♂️ 노형 점핑 & 테라피 통합 관리 시스템 (Nohyung Jumping ERP)

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Platform](https://img.shields.io/badge/Platform-Google_Apps_Script-green)
![Tech](https://img.shields.io/badge/Tech-HTML5%20%7C%20JS%20%7C%20CSS3-orange)

노형 점핑 & 테라피 센터의 회원 예약, 출석 기록, 및 회원권 갱신을 효율적으로 관리하기 위한 통합 웹 애플리케이션입니다. Google Sheets를 데이터베이스로 활용하며, 강력한 보안과 실시간 데이터 동기화를 제공합니다.

---

## ✨ 주요 기능 (Key Features)

### 📅 스마트 예약 시스템
- **점핑 & 테라피 선택**: 각 프로그램별 실시간 예약 신청 기능.
- **잔여 횟수 체크**: 회원별 잔여 세션 수에 따른 지능형 예약 제한.
- **실시간 대시보드**: 관리자가 한눈에 확인 가능한 예약 현황판.

### 📊 효율적인 출석 및 로그 관리
- **실시간 입장 체크**: 회원의 입장 시간을 기록하고 현황을 관리자에게 노출.
- **자동 퇴실 시스템**: 매일 밤 운영 종료 후 미퇴실 인원을 자동으로 정리하는 Force-Exit 로직 탑재.
- **상세 통계**: 일별, 회원별 출석 데이터를 시각화하여 제공.

### 💳 멤버십 & 갱신 관리
- **회원권 갱신 로직**: 남은 횟수 이월(Carry-over) 처리 및 자동 만료일 계산.
- **데이터 무결성**: 구글 스프레드시트와의 완벽한 연동으로 데이터 유실 방지.

### 📱 최신 웹 디자인 (Premium UI/UX)
- **반응형 레이아웃**: PC와 모바일 어디서든 최적화된 화면 제공.
- **글래스모피즘(Glassmorphism)**: 현대적이고 세련된 투명 레이어 디자인 적용.
- **부드러운 인터랙션**: 사용자 경험을 극대화하는 마이크로 애니메이션 및 호버 효과.

---

## 🛠 기술 스택 (Tech Stack)

- **Backend**: Google Apps Script (GAS)
- **Frontend**: HTML5, Vanilla JavaScript, CSS3
- **Database**: Google Sheets (Spreadsheet API)
- **Deployment**: Google Web App Executable

---

## 📂 프로젝트 구조 (Structure)

- `Code.gs`: 서버 사이드 로직 및 API 핸들러
- `admin.html`: 관리자 전용 대시보드 및 통계 페이지
- `reservation.html`: 회원 예약 신청 인터페이스
- `attendance.html`: 출석 현황 모니터링 페이지
- `registration.html`: 신규 회원 등록 및 정보 수정
- `renewal.html`: 회원권 갱신 및 결제 관리

---

## 👤 개발 및 유지보수
- **개발자**: [hamchorom-jeju](https://github.com/hamchorom-jeju)
- **이메일**: moonjin05@gmail.com

---
*본 프로젝트는 노형 점핑 & 테라피 센터의 운영 효율화를 위해 커스텀 제작되었습니다.*
