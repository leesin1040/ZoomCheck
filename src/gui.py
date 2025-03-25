import sys
import logging
import os
import threading
from datetime import datetime

# PyQt6 import
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QPushButton, QTextEdit, QHBoxLayout, QLabel,
    QSplitter, QTabWidget
)
from PyQt6.QtCore import QObject, pyqtSignal, Qt
from PyQt6.QtGui import QFont

from .window_finder import WindowFinder
from .text_extractor import TextExtractor

logger = logging.getLogger(__name__)

# 로깅 핸들러를 생성하여 GUI로 로그 전달
class QTextEditLogger(logging.Handler):
    def __init__(self, widget):
        super().__init__()
        self.widget = widget
        self.setLevel(logging.INFO)
        self.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
        self.widget.setReadOnly(True)

    def emit(self, record):
        msg = self.format(record)
        self.widget.append(msg)
        # 자동 스크롤
        self.widget.verticalScrollBar().setValue(self.widget.verticalScrollBar().maximum())

class SignalManager(QObject):
    """UI 업데이트를 위한 시그널 관리 클래스"""
    update_log = pyqtSignal(str)
    update_participant_list = pyqtSignal(str)
    enable_buttons = pyqtSignal(bool)

class ZoomCheckGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # 기본 변수 초기화
        self.window_finder = WindowFinder()
        self.text_extractor = TextExtractor()
        self.signal_manager = SignalManager()
        self.extraction_in_progress = False
        
        # 결과 저장용 변수
        self.current_participants = []
        self.current_duplicate_info = {}
        
        # UI 설정
        self.init_ui()
        
        # 시그널 연결
        self.signal_manager.update_log.connect(self.append_to_log)
        self.signal_manager.update_participant_list.connect(self.append_to_participant_list)
        self.signal_manager.enable_buttons.connect(self.set_buttons_enabled)
        
        # 로깅 설정
        self.setup_logging()
    
    def setup_logging(self):
        """로깅 핸들러 설정"""
        # 기존 핸들러 제거
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        
        # GUI 로깅 핸들러 추가
        log_handler = QTextEditLogger(self.log_tab)
        logger.addHandler(log_handler)
        
        # 로그 레벨 설정
        logger.setLevel(logging.INFO)
    
    def init_ui(self):
        """UI 요소 초기화"""
        self.setWindowTitle('Zoom 참가자 체크')
        self.setGeometry(100, 100, 1000, 700)
        
        # 중앙 위젯 및 레이아웃
        central_widget = QWidget()
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        
        # 버튼 영역
        button_layout = QHBoxLayout()
        
        # 참가자 새로고침 버튼
        self.refresh_button = QPushButton('참가자 목록 가져오기')
        self.refresh_button.clicked.connect(self.refresh_participants)
        button_layout.addWidget(self.refresh_button)
        
        # 중지 버튼
        self.stop_button = QPushButton('중지')
        self.stop_button.clicked.connect(self.stop_extraction)
        self.stop_button.setEnabled(False)
        button_layout.addWidget(self.stop_button)
        
        # 클립보드 복사 버튼
        self.copy_button = QPushButton('참가자 목록 복사')
        self.copy_button.clicked.connect(self.copy_to_clipboard)
        button_layout.addWidget(self.copy_button)
        
        main_layout.addLayout(button_layout)
        
        # 탭 위젯 생성
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        
        # 참가자 목록 탭
        self.participant_tab = QTextEdit()
        self.participant_tab.setReadOnly(True)
        font = QFont("Consolas", 10)
        self.participant_tab.setFont(font)
        self.tab_widget.addTab(self.participant_tab, "참가자 목록")
        
        # 로그 탭
        self.log_tab = QTextEdit()
        self.log_tab.setReadOnly(True)
        self.log_tab.setFont(font)
        self.tab_widget.addTab(self.log_tab, "로그")
        
        # 초기 안내 메시지
        instructions = (
            "사용 방법:\n"
            "1. Zoom 미팅에 참가 후 '참가자' 버튼을 클릭하여 참가자 목록 창을 여세요.\n"
            "2. '참가자 목록 가져오기' 버튼을 클릭하여 현재 참가자 목록을 불러옵니다.\n"
            "3. '참가자 목록 복사' 버튼을 클릭하여 클립보드에 참가자 목록을 복사할 수 있습니다.\n"
            "4. '로그' 탭에서 자세한 실행 정보를 확인할 수 있습니다.\n"
        )
        self.participant_tab.append(instructions)
        
    def refresh_participants(self):
        """참가자 목록 새로고침"""
        if not self.extraction_in_progress:
            self.extraction_in_progress = True
            self.refresh_button.setEnabled(False)
            self.copy_button.setEnabled(False)
            self.stop_button.setEnabled(True)
            
            # 참가자 탭으로 전환
            self.tab_widget.setCurrentIndex(0)
            
            # 백그라운드 스레드에서 실행
            extraction_thread = threading.Thread(target=self._extract_participants)
            extraction_thread.daemon = True
            extraction_thread.start()
    
    def stop_extraction(self):
        """참가자 추출 중단"""
        if self.extraction_in_progress:
            self.text_extractor.stop_extraction()
            self.signal_manager.update_log.emit("참가자 목록 추출 중단 요청됨...")
            self.stop_button.setEnabled(False)
    
    def _extract_participants(self):
        """별도 스레드에서 참가자 추출 실행"""
        try:
            self.signal_manager.update_log.emit("Zoom 참가자 창 검색 중...")
            self.signal_manager.update_participant_list.emit("Zoom 참가자 창 검색 중...")
            
            # 올바른 메서드 호출 방식으로 수정
            # window_handle = self.window_finder.find_zoom_window() -> 잘못된 방식
            window_handle = WindowFinder.find_zoom_window()  # 올바른 방식 (정적 메서드이므로 클래스에서 직접 호출)
            
            if window_handle:
                self.signal_manager.update_log.emit(f"참가자 창 발견! 목록 추출 중...")
                self.signal_manager.update_participant_list.emit("참가자 창 발견! 목록 추출 중...")
                
                try:
                    # 값 추출 시도
                    result = self.text_extractor.extract_participants(window_handle)
                    
                    # 반환 값이 튜플(리스트, 딕셔너리)인지 확인
                    if isinstance(result, tuple) and len(result) >= 2:
                        participants = result[0]  # 첫 번째 항목은 참가자 목록
                        duplicate_info = result[1] if len(result) > 1 else {}  # 두 번째 항목은 중복 정보
                        
                        # 결과 저장
                        self.current_participants = participants
                        self.current_duplicate_info = duplicate_info
                        
                    elif isinstance(result, list):
                        # 이전 버전 호환성 - 단일 리스트만 반환하는 경우
                        participants = result
                        duplicate_info = {}
                        
                        # 결과 저장
                        self.current_participants = participants
                        self.current_duplicate_info = {}
                        
                    else:
                        # 기타 경우
                        participants = []
                        duplicate_info = {}
                        
                        # 결과 초기화
                        self.current_participants = []
                        self.current_duplicate_info = {}
                
                except Exception as e:
                    self.signal_manager.update_log.emit(f"값 추출 중 오류: {str(e)}")
                    self.signal_manager.update_participant_list.emit(f"값 추출 중 오류: {str(e)}")
                    participants = []
                    duplicate_info = {}
                    
                    # 결과 초기화
                    self.current_participants = []
                    self.current_duplicate_info = {}
                
                if participants:
                    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    # 중복 개수 계산
                    true_duplicates = {}
                    possible_duplicates = {}
                    
                    for name, info in duplicate_info.items():
                        if info.get('type') == '정확한 중복':
                            true_duplicates[name] = info
                        elif info.get('type') == '동명이인 가능성':
                            possible_duplicates[name] = info
                    
                    # 헤더 표시
                    header = f"\n[{now}] 현재 참가자 목록 ({len(participants)}명)"
                    if true_duplicates:
                        total_true_dups = sum(info['count'] - 1 for info in true_duplicates.values())
                        header += f", 확실한 중복 {total_true_dups}개"
                    if possible_duplicates:
                        possible_dup_count = len(possible_duplicates)
                        header += f", 동명이인 가능성 {possible_dup_count}개"
                    
                    # 참가자 탭에만 결과 표시
                    self.signal_manager.update_participant_list.emit(header + ":")
                    
                    # 참가자 목록 표시
                    for participant in participants:
                        self.signal_manager.update_participant_list.emit(participant)
                    
                else:
                    self.signal_manager.update_log.emit("참가자를 찾을 수 없습니다.")
                    self.signal_manager.update_participant_list.emit("참가자를 찾을 수 없습니다.")
            else:
                self.signal_manager.update_log.emit("Zoom 참가자 창을 찾을 수 없습니다. Zoom 미팅이 실행 중이고 참가자 목록이 열려있는지 확인해주세요.")
                self.signal_manager.update_participant_list.emit("Zoom 참가자 창을 찾을 수 없습니다. Zoom 미팅이 실행 중이고 참가자 목록이 열려있는지 확인해주세요.")
                
        except Exception as e:
            self.signal_manager.update_log.emit(f"오류 발생: {str(e)}")
            self.signal_manager.update_participant_list.emit(f"오류 발생: {str(e)}")
            import traceback
            self.signal_manager.update_log.emit(traceback.format_exc())
        
        finally:
            # 완료 후 상태 초기화
            self.extraction_in_progress = False
            self.signal_manager.enable_buttons.emit(True)
    
    def copy_to_clipboard(self):
        """최신 참가자 목록을 클립보드에 복사합니다."""
        try:
            if self.current_participants:
                # 중복 분류
                true_duplicates = {}
                possible_duplicates = {}
                
                for name, info in self.current_duplicate_info.items():
                    if info.get('type') == '정확한 중복':
                        true_duplicates[name] = info
                    elif info.get('type') == '동명이인 가능성':
                        possible_duplicates[name] = info
                
                # 현재 시간 기준 헤더 생성
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                header = f"[{now}] 현재 참가자 목록 ({len(self.current_participants)}명)"
                
                if true_duplicates:
                    total_true_dups = sum(info['count'] - 1 for info in true_duplicates.values())
                    header += f", 확실한 중복 {total_true_dups}개"
                if possible_duplicates:
                    possible_dup_count = len(possible_duplicates)
                    header += f", 동명이인 가능성 {possible_dup_count}개"
                
                # 클립보드용 텍스트 생성 - 참가자 목록 먼저
                clipboard_text = f"{header}:\n" + "\n".join(self.current_participants)
                
                # 중복 정보도 클립보드에 추가
                if self.current_duplicate_info:
                    clipboard_text += "\n\n=== 중복 이름 분석 ===\n"
                    
                    if true_duplicates:
                        clipboard_text += "\n[확실한 중복]\n"
                        for name, info in true_duplicates.items():
                            clipboard_text += f"• '{name}'이(가) {info['count']}번 발견됨 (동일한 상태)\n"
                            if isinstance(info['details'], str):
                                clipboard_text += f"  → 상태: {info['details']}\n"
                            elif isinstance(info['details'], list):
                                clipboard_text += f"  → 상태: {info['details'][0]}\n"
                    
                    if possible_duplicates:
                        clipboard_text += "\n[동명이인 가능성]\n"
                        for name, info in possible_duplicates.items():
                            clipboard_text += f"• '{name}'이(가) {info['count']}번 발견됨 (서로 다른 상태)\n"
                            if isinstance(info['details'], list):
                                for i, status in enumerate(info['details']):
                                    clipboard_text += f"  → 상태 {i+1}: {status}\n"
                
                # 클립보드에 복사
                clipboard = QApplication.clipboard()
                clipboard.setText(clipboard_text)
                self.signal_manager.update_participant_list.emit("\n클립보드에 참가자 목록이 복사되었습니다.")
            else:
                self.signal_manager.update_participant_list.emit("\n복사할 참가자 목록이 없습니다. 먼저 목록을 가져오세요.")
        except Exception as e:
            self.signal_manager.update_log.emit(f"\n클립보드 복사 중 오류 발생: {str(e)}")
            self.signal_manager.update_participant_list.emit(f"\n클립보드 복사 중 오류 발생: {str(e)}")
            import traceback
            self.signal_manager.update_log.emit(traceback.format_exc())
    
    def append_to_log(self, message):
        """로그 영역에 메시지 추가"""
        self.log_tab.append(message)
        # 자동 스크롤 (PyQt6 호환)
        scrollbar = self.log_tab.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def append_to_participant_list(self, message):
        """참가자 목록 영역에 메시지 추가"""
        self.participant_tab.append(message)
        # 자동 스크롤 (PyQt6 호환)
        scrollbar = self.participant_tab.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def set_buttons_enabled(self, enabled):
        """버튼 활성화/비활성화"""
        self.refresh_button.setEnabled(enabled)
        self.copy_button.setEnabled(enabled)
        self.stop_button.setEnabled(not enabled)

def main():
    app = QApplication(sys.argv)
    gui = ZoomCheckGUI()
    gui.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 