# ZoomCheck

Zoom 미팅 참석자 확인 및 관리를 위한 GUI 프로그램입니다.

## 주요 기능
- Zoom 참가자 창 자동 감지
- 실시간 참가자 목록 추출
- 참가자 입퇴장 모니터링
- 참가자 목록 클립보드 복사
- 참가자 수 자동 계산

## 설치 요구사항
- Python 3.x
- Windows 운영체제
- Zoom 클라이언트

## 설치 방법
1. 저장소 클론 또는 다운로드
2. 가상환경 생성 및 활성화
   ```powershell
   python -m venv venv
   .\venv\Scripts\activate
   ```
3. 필수 패키지 설치
   ```powershell
   pip install -r requirements.txt
   ```
4. 종료방법 
   ```powershell
   deactivate
   ```

## 사용 방법
1. Zoom 미팅에 참가합니다
2. 참가자 목록 창을 엽니다 ("참가자" 버튼 클릭)
3. 프로그램 실행:
   ```powershell
   python src\gui.py
   ```

### 버튼 기능
- **현재 참가자 목록**: 현재 Zoom 참가자 목록을 가져와서 표시
- **참가자 목록 복사**: 가장 최근에 가져온 참가자 목록을 클립보드에 복사
- **알림 켜기/끄기**: 참가자 입퇴장 모니터링 시작/중지

## 프로젝트 구조
```
ZoomCheck/
├── src/
│   ├── gui.py           # GUI 및 메인 프로그램
│   ├── window_finder.py # Zoom 창 검색 기능
│   ├── text_extractor.py # 참가자 정보 추출
│   └── utils.py         # 유틸리티 함수
└── requirements.txt     # 필수 패키지 목록
```

## 필수 패키지
- pywin32: Windows API 접근
- pywinauto: GUI 자동화
- PyQt6: GUI 프레임워크
- pyperclip: 클립보드 조작

## 주의사항
- Windows 운영체제에서만 실행 가능합니다
- Zoom 클라이언트가 설치되어 있어야 합니다
- 참가자 목록 창이 열려있어야 합니다
- 참가자 창이 다른 창에 가려지지 않도록 해주세요

## 문제 해결
- 참가자 목록이 보이지 않는 경우: Zoom에서 참가자 창을 다시 열어보세요
- 프로그램이 응답하지 않는 경우: 프로그램을 재시작하세요
- 클립보드 복사가 안 되는 경우: 다시 "현재 참가자 목록" 버튼을 누른 후 시도해보세요