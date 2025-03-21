from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QPushButton, QTextEdit, QLabel, QHBoxLayout)
from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QClipboard
import sys
import os
import logging
import pythoncom  # COM 초기화를 위해 추가
import pyperclip  # 새로운 라이브러리 사용

# 로깅 설정
logging.basicConfig(level=logging.INFO)

# 현재 스크립트의 디렉토리를 파이썬 경로에 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

from src.text_extractor import TextExtractor
from src.window_finder import WindowFinder

class ZoomCheckGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Zoom 참가자 체크")
        self.setGeometry(100, 100, 800, 600)
        
        # COM 초기화
        pythoncom.CoInitialize()
        
        # 중앙 위젯 설정
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # 상태 레이블
        self.status_label = QLabel("대기 중...")
        layout.addWidget(self.status_label)
        
        # 버튼 컨테이너
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        
        # 1. 현재 참가자 목록 버튼
        self.refresh_button = QPushButton("현재 참가자 목록")
        self.refresh_button.clicked.connect(self.refresh_participants)
        button_layout.addWidget(self.refresh_button)
        
        # 2. 복사 버튼
        self.copy_button = QPushButton("참가자 목록 복사")
        self.copy_button.clicked.connect(self.copy_to_clipboard)
        button_layout.addWidget(self.copy_button)
        
        # 3. 알림 토글 버튼
        self.notification_button = QPushButton("알림 켜기")
        self.notification_button.clicked.connect(self.toggle_notifications)
        button_layout.addWidget(self.notification_button)
        
        layout.addWidget(button_container)
        
        # 로그 표시 영역
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area)
        
        # 클래스 초기화
        self.window_finder = WindowFinder()
        self.text_extractor = TextExtractor()
        
        # 타이머 설정
        self.notification_timer = QTimer()
        self.notification_timer.timeout.connect(self.check_participants_changes)
        self.is_monitoring = False
        
        # 이전 참가자 목록 저장
        self.previous_participants = set()
        
        # 최근 참가자 목록을 저장할 변수 추가
        self.latest_participants_log = ""
        
    def refresh_participants(self):
        """현재 참가자 목록을 가져와서 표시"""
        try:
            self.status_label.setText("참가자 창 검색 중...")
            
            zoom_window = self.window_finder.find_zoom_window()
            if zoom_window:
                self.status_label.setText("참가자 목록 추출 중...")
                participants = self.text_extractor.extract_participants(zoom_window)
                if participants:
                    current_time = self.text_extractor.get_current_time()
                    log_entry = f"[{current_time}] 현재 참가자 목록 ({len(participants)}명):\n"
                    log_entry += "\n".join(participants) + "\n\n"
                    
                    # 최근 참가자 목록 업데이트
                    self.latest_participants_log = log_entry
                    
                    self.log_area.append(log_entry)
                    self.previous_participants = set(participants)
                    self.status_label.setText(f"참가자 목록을 불러왔습니다. (총 {len(participants)}명)")
                else:
                    self.status_label.setText("참가자를 찾을 수 없습니다.")
            else:
                self.status_label.setText("Zoom 참가자 창을 찾을 수 없습니다.")
                
        except Exception as e:
            self.status_label.setText(f"오류 발생: {str(e)}")
        finally:
            # COM 해제
            pythoncom.CoUninitialize()
    
    def copy_to_clipboard(self):
        """최근 참가자 목록을 클립보드에 복사"""
        try:
            if self.latest_participants_log:
                pyperclip.copy(self.latest_participants_log)
                self.status_label.setText("최근 참가자 목록이 클립보드에 복사되었습니다.")
            else:
                self.status_label.setText("복사할 참가자 목록이 없습니다. '현재 참가자 목록' 버튼을 먼저 클릭하세요.")
        except Exception as e:
            self.status_label.setText(f"복사 중 오류 발생: {str(e)}")
    
    def toggle_notifications(self):
        """참가자 입퇴장 알림 토글"""
        if not self.is_monitoring:
            self.is_monitoring = True
            self.notification_button.setText("알림 끄기")
            self.notification_timer.start(5000)  # 5초마다 확인
            self.status_label.setText("참가자 변동 모니터링 시작")
        else:
            self.is_monitoring = False
            self.notification_button.setText("알림 켜기")
            self.notification_timer.stop()
            self.status_label.setText("참가자 변동 모니터링 중지")
    
    def check_participants_changes(self):
        """참가자 변동 사항 확인"""
        try:
            # COM 초기화
            pythoncom.CoInitialize()
            
            zoom_window = self.window_finder.find_zoom_window()
            if zoom_window:
                current_participants = set(self.text_extractor.extract_participants(zoom_window))
                
                # 새로 입장한 참가자 확인
                new_participants = current_participants - self.previous_participants
                if new_participants:
                    log_entry = f"[{self.text_extractor.get_current_time()}] 새로운 참가자:\n"
                    log_entry += "\n".join(new_participants) + "\n\n"
                    self.log_area.append(log_entry)
                
                # 퇴장한 참가자 확인
                left_participants = self.previous_participants - current_participants
                if left_participants:
                    log_entry = f"[{self.text_extractor.get_current_time()}] 퇴장한 참가자:\n"
                    log_entry += "\n".join(left_participants) + "\n\n"
                    self.log_area.append(log_entry)
                
                self.previous_participants = current_participants
                
        except Exception as e:
            self.status_label.setText(f"모니터링 중 오류 발생: {str(e)}")
        finally:
            # COM 해제
            pythoncom.CoUninitialize()

    def closeEvent(self, event):
        """프로그램 종료 시 정리"""
        try:
            if self.is_monitoring:
                self.notification_timer.stop()
            pythoncom.CoUninitialize()
        except:
            pass
        super().closeEvent(event)

def main():
    # DPI 인식 설정
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    window = ZoomCheckGUI()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 