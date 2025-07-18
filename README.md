# ZoomCheck

Zoom 미팅 참가자 목록을 자동으로 추출하는 Windows 전용 프로그램입니다.

---

## 실행파일(.exe) 사용법 (비개발자용)

1. **`ZoomCheck.exe` 파일을 다운로드/복사합니다.**
2. **Zoom을 실행하고, 참가자 목록 창을 엽니다.**
3. **`ZoomCheck.exe`를 더블클릭하여 실행합니다.**
   - 참가자 목록이 자동으로 추출되어 파일로 저장됩니다.
4. 문제가 생기면 오류 메시지를 복사해서 문의해 주세요.

> ※ Python, 추가 설치, 복잡한 설정이 필요 없습니다.

---

## 주요 기능
- Zoom 참가자 창 자동 감지 및 참가자 목록 추출
- 참가자 수 자동 계산 및 파일 저장
- (옵션) 참가자 입퇴장 모니터링, 클립보드 복사 등

---

## 자주 묻는 질문(FAQ) 및 문제 해결
- **참가자 목록이 일부만 추출됨:**
  - Zoom 참가자 창이 완전히 열려 있는지, 다른 창에 가려지지 않았는지 확인하세요.
  - 참가자 수가 많을 경우, 추출에 시간이 더 걸릴 수 있습니다.
- **실행이 안 되거나 오류가 발생:**
  - 오류 메시지를 복사해서 문의해 주세요.
- **보안 프로그램에서 차단:**
  - 직접 만든 프로그램이므로, 신뢰할 수 있는 경우 예외처리 후 사용하세요.

---

## (개발자용) 소스코드 실행 및 실행파일 빌드 방법

### 1. 설치 요구사항
- Python 3.11.x (반드시 3.11 버전 사용)
- Windows 운영체제
- Zoom 클라이언트

### 2. 소스코드 실행
```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python -m src.gui
```

### 3. 실행파일(.exe) 빌드
1. **기존 가상환경 비활성화**
   ```bash
deactivate
```
2. **Python 3.11로 새 가상환경 생성 및 활성화**
   ```bash
py -3.11 -m venv venv311
venv311\Scripts\activate
```
3. **의존성 및 PyInstaller 설치**
   ```bash
pip install -r requirements.txt
pip install pyinstaller
```
4. **실행파일 빌드**
   ```bash
pyinstaller --noconsole --onefile -n ZoomCheck src/main.py
```
   - 빌드가 완료되면 `dist/ZoomCheck.exe` 파일이 생성됩니다.

### 4. 참고 및 주의사항
- Python 3.12 이상에서는 빌드 오류가 발생하므로 반드시 Python 3.11을 사용하세요.
- 문제가 발생하면 오류 메시지를 복사해서 문의해 주세요.

---

## 프로젝트 구조(참고)
```
ZoomCheck/
├── src/
│   ├── common/          # 공통 모듈
│   ├── extractors/      # 추출기 모듈
│   ├── gui.py           # GUI 및 메인 프로그램
│   ├── window_finder.py # Zoom 창 검색 기능
│   ├── text_extractor.py # 참가자 정보 추출
│   ├── utils.py         # 유틸리티 함수
│   └── main.py          # CLI 진입점
├── tests/               # 테스트 파일
└── requirements.txt     # 필수 패키지 목록
```