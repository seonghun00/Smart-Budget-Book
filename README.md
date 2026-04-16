<p align="center">
  <img src="https://capsule-render.vercel.app/api?type=waving&color=auto&height=80&section=header" width="100%" />
</p>

# 📱 Automated Expense Logging System
> **iPhone Short cuts + Google Apps Script (GAS)**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=flat&logo=google-apps-script&logoColor=white)](https://developers.google.com/apps-script)
[![iOS Shortcuts](https://img.shields.io/badge/iOS%20Shortcuts-FF4500?style=flat&logo=shortcuts&logoColor=white)](https://support.apple.com/guide/shortcuts/welcome/ios)

아이폰 단축어를 활용하여 단 몇 초 만에 소비 내역을 구글 스프레드시트에 업데이트하고, 실시간 대시보드로 시각화하는 자동화 시스템입니다.

## 🚀 Key Features
- **One-Tap Logging**: 아이폰 단축어 통해 앱 실행 없이 즉시 지출 입력.
- **Smart Parsing**: 자연어 입력(`카페 5000`)을 분석하여 카테고리와 금액을 자동 분리.
- **Dynamic Dashboard**: 실시간 구글 시트 차트를 통해 지출 비중 시각화.
- **Monthly Auto-Backup**: 매달 말일 밤, 데이터를 월별 요약 테이블로 자동 백업하고 일별 기록을 초기화하는 서버리스 트리거 탑재.
- **Error Handling**: 마이너스 입력(`-5000`)을 통한 오기입 수정 기능 및 데이터 형식 예외 처리.

## 🛠 Tech Stack
- **Client**: iOS Shortcuts (iOS 단축어)
- **Backend**: Google Apps Script (JavaScript 기반 서버리스 API)
- **Database/UI**: Google Sheets

## 🏗 System Architecture
1. **Input**: 사용자가 아이폰 단축어 버튼을 눌러 "항목 금액" 입력.
2. **Transfer**: 단축어가 구글 웹 앱 URL로 `GET` 요청 전송.
3. **Processing**: GAS가 데이터를 받아 파싱 후, 오늘 날짜와 일치하는 행(Row)에 데이터 기록.
4. **Maintenance**: 매달 말일, 트리거가 작동하여 월간 합계를 하단 DB로 이동 및 시트 초기화.

## 📸 Screen Shots
<img width="1710" height="681" alt="image" src="https://github.com/user-attachments/assets/50d2b0a1-e077-4c6e-bf44-7a932a3eb1df" />
<img width="1713" height="681" alt="image" src="https://github.com/user-attachments/assets/c053db7a-3ae0-49f1-81d6-fe513b103f60" />


## 💻 How to Setup
1. **Google Sheets**: 제공된 템플릿 형식으로 시트를 구성합니다.
2. **Apps Script**: `src/code.gs`의 코드를 복사하여 구글 시트 내 앱스 스크립트에 배포합니다.
3. **iOS Shortcut**: 아이폰 단축어를 생성하고 웹 앱 URL을 연결합니다.
4. **Action Button**: 아이폰 설정에서 해당 단축어를 액션 버튼으로 지정합니다. (아이폰 15 이상 가능)

<br>

**자세한 추가 내용들은 assets/README.md 에서 확인 가능.**

<br>

---
© 2026 Seong-hun Bae.
