import sys
import logging
import os
import threading
from datetime import datetime
import time

# PyQt6 import
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QPushButton, QTextEdit, QHBoxLayout, QLabel,
    QSplitter, QMenu
)
from PyQt6.QtGui import QAction
from PyQt6.QtCore import QObject, pyqtSignal, Qt, QTimer
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
        try:
            msg = self.format(record)
            if self.widget and hasattr(self.widget, 'append'):
                self.widget.append(msg)
                # 자동 스크롤
                try:
                    scrollbar = self.widget.verticalScrollBar()
                    if scrollbar:
                        scrollbar.setValue(scrollbar.maximum())
                except Exception as scroll_error:
                    # 스크롤 오류는 무시
                    pass
        except Exception as e:
            # GUI 업데이트 실패 시 콘솔에 출력
            print(f"로그 핸들러 오류: {e}")
            print(f"원본 메시지: {record.getMessage()}")

class SignalManager(QObject):
    """UI 업데이트를 위한 시그널 관리 클래스"""
    update_log = pyqtSignal(str)
    update_participant_list = pyqtSignal(str)
    enable_buttons = pyqtSignal(bool)

class ZoomCheckGUI(QMainWindow):
    def __init__(self):
        try:
            super().__init__()
            print("GUI 초기화 시작...")
            
            # 기본 변수 초기화
            self.extraction_in_progress = False
            self.current_participants = []
            self.current_duplicate_info = {}
            
            # TextExtractor 초기화
            try:
                self.text_extractor = TextExtractor()
                print("TextExtractor 초기화 완료")
            except Exception as e:
                print(f"TextExtractor 초기화 실패: {e}")
                self.text_extractor = None
            
            # 추출 진행 상태
            self.extraction_in_progress = False
            
            # 시그널 매니저 초기화
            try:
                self.signal_manager = SignalManager()
                print("SignalManager 초기화 완료")
            except Exception as e:
                print(f"SignalManager 초기화 실패: {e}")
                return
            
            # UI 설정
            try:
                self.init_ui()
                print("UI 초기화 완료")
            except Exception as e:
                print(f"UI 초기화 실패: {e}")
                return
            
            # 시그널 연결
            try:
                self.signal_manager.update_log.connect(self.append_to_log)
                self.signal_manager.update_participant_list.connect(self.append_to_participant_list)
                self.signal_manager.enable_buttons.connect(self.set_buttons_enabled)
                print("시그널 연결 완료")
            except Exception as e:
                print(f"시그널 연결 실패: {e}")
            
            # 로깅 설정
            try:
                self.setup_logging()
                print("로깅 설정 완료")
            except Exception as e:
                print(f"로깅 설정 실패: {e}")
                
            print("GUI 초기화 완료")
        except Exception as e:
            print(f"GUI 초기화 중 치명적 오류: {e}")
            import traceback
            print(f"초기화 오류 상세: {traceback.format_exc()}")
            raise
        
    def setup_logging(self):
        """로깅 핸들러 설정"""
        # 기존 핸들러 제거
        logger = logging.getLogger()
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        
        # GUI 로깅 핸들러 추가
        log_handler = QTextEditLogger(self.log_area)
        log_handler.setLevel(logging.INFO)
        logger.addHandler(log_handler)
        
        # 콘솔 핸들러도 유지
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        logger.addHandler(console_handler)
        
        # 루트 로거 레벨 설정
        logger.setLevel(logging.INFO)
        
        # pywinauto 로깅도 캡처
        pywinauto_logger = logging.getLogger('pywinauto')
        pywinauto_logger.setLevel(logging.INFO)
        
        # 로그 출력 테스트
        logger.info("로깅 시스템 초기화 완료")
    
    def init_ui(self):
        """UI 요소 초기화"""
        self.setWindowTitle('Zoom 참가자 체크')
        self.setGeometry(100, 100, 1200, 800)
        
        # 중앙 위젯 및 레이아웃
        central_widget = QWidget()
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        
        # 참가자 목록 관련 버튼들
        button_layout = QHBoxLayout()
        
        # 참가자 목록 가져오기 버튼
        self.refresh_button = QPushButton('참가자 목록 가져오기')
        self.refresh_button.clicked.connect(self.refresh_participants)
        button_layout.addWidget(self.refresh_button)
        
        # 클립보드에 복사 버튼
        self.copy_button = QPushButton('클립보드에 복사')
        self.copy_button.clicked.connect(self.copy_to_clipboard)
        button_layout.addWidget(self.copy_button)
        
        main_layout.addLayout(button_layout)
        
        # 분할기 생성 (수평 분할)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setSizes([600, 600])  # 초기 크기 설정
        main_layout.addWidget(splitter)
        
        # 왼쪽: 참가자 목록 영역
        participant_widget = QWidget()
        participant_layout = QVBoxLayout()
        participant_widget.setLayout(participant_layout)
        
        participant_label = QLabel("참가자 목록")
        participant_layout.addWidget(participant_label)
        
        self.participant_area = QTextEdit()
        self.participant_area.setReadOnly(True)
        font = QFont("Consolas", 10)
        self.participant_area.setFont(font)
        participant_layout.addWidget(self.participant_area)
        
        # 오른쪽: 로그 영역
        log_widget = QWidget()
        log_layout = QVBoxLayout()
        log_widget.setLayout(log_layout)
        
        log_label = QLabel("로그")
        log_layout.addWidget(log_label)
        
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setFont(font)
        # 로그창을 선택 가능하게 만들고 기본 복사 기능 활성화
        self.log_area.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextSelectableByKeyboard)
        # 로그창에 컨텍스트 메뉴 추가 (복사 기능)
        self.log_area.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.log_area.customContextMenuRequested.connect(self.show_log_context_menu)
        
        # 키보드 단축키 추가
        from PyQt6.QtGui import QKeySequence, QShortcut
        copy_shortcut = QShortcut(QKeySequence.StandardKey.Copy, self.log_area)
        copy_shortcut.activated.connect(self.copy_selected_log)
        select_all_shortcut = QShortcut(QKeySequence.StandardKey.SelectAll, self.log_area)
        select_all_shortcut.activated.connect(self.log_area.selectAll)
        
        log_layout.addWidget(self.log_area)
        
        # 분할기에 위젯 추가
        splitter.addWidget(participant_widget)
        splitter.addWidget(log_widget)
        
        # 초기 안내 메시지
        instructions = (
            "사용 방법:\n"
            "1. Zoom 미팅에 참가 후 '참가자' 버튼을 클릭하여 참가자 목록 창을 여세요.\n"
            "2. '참가자 목록 가져오기' 버튼을 클릭하여 현재 참가자 목록을 불러옵니다.\n"
            "3. '참가자 목록 복사' 버튼을 클릭하여 클립보드에 참가자 목록을 복사할 수 있습니다.\n"
            "4. 오른쪽 '로그' 영역에서 자세한 실행 정보를 확인할 수 있습니다.\n"
        )
        self.participant_area.append(instructions)
    
    def refresh_participants(self):
        """참가자 목록을 새로 가져옵니다."""
        if not self.extraction_in_progress:
            self.extraction_in_progress = True
            self.set_buttons_enabled(False)
            
            # 백그라운드 스레드에서 실행
            extraction_thread = threading.Thread(target=self._extract_participants)
            extraction_thread.daemon = True
            extraction_thread.start()
        else:
            self.signal_manager.update_log.emit("이미 참가자 목록을 가져오는 중입니다.")

    def _extract_participants(self):
        """백그라운드에서 참가자 목록을 추출합니다."""
        try:
            if not hasattr(self, 'text_extractor') or self.text_extractor is None:
                self.signal_manager.update_log.emit("TextExtractor가 초기화되지 않았습니다.")
                return
                
            # Zoom 창 찾기
            try:
                window_handle = WindowFinder.find_zoom_window()
            except Exception as e:
                self.signal_manager.update_log.emit(f"Zoom 창 찾기 중 오류: {str(e)}")
                return
            
            if window_handle:
                self.signal_manager.update_log.emit("Zoom 창을 찾았습니다. 참가자 목록을 가져오는 중...")
                
                # 참가자 추출
                try:
                    result = self.text_extractor.extract_participants_with_retry(window_handle)
                except Exception as e:
                    self.signal_manager.update_log.emit(f"참가자 추출 중 오류: {str(e)}")
                    import traceback
                    self.signal_manager.update_log.emit(f"추출 오류 상세: {traceback.format_exc()}")
                    return
                
                if isinstance(result, tuple) and len(result) >= 2:
                    participants = result[0]
                    duplicate_info = result[1]
                    
                    # 결과 저장 (복사 기능용)
                    try:
                        self.current_participants = participants
                        self.current_duplicate_info = duplicate_info
                    except Exception as e:
                        self.signal_manager.update_log.emit(f"결과 저장 중 오류: {str(e)}")
                    
                    if participants:
                        # 참가자 목록 표시
                        try:
                            participant_text = f"\n=== 참가자 목록 ({len(participants)}명) ===\n"
                            for i, participant in enumerate(participants, 1):
                                participant_text += f"{i:3d}. {participant}\n"
                            
                            # 중복 정보 표시
                            if duplicate_info:
                                participant_text += "\n[중복 정보]\n"
                                for name, info in duplicate_info.items():
                                    participant_text += f"• '{name}': {info['count']}번 발견\n"
                                    if isinstance(info['details'], str):
                                        participant_text += f"  → 상태: {info['details']}\n"
                                    elif isinstance(info['details'], list):
                                        participant_text += f"  → 상태: {info['details'][0]}\n"
                            
                            self.signal_manager.update_participant_list.emit(participant_text)
                            self.signal_manager.update_log.emit(f"✅ 참가자 목록을 성공적으로 가져왔습니다. (총 {len(participants)}명)")
                        except Exception as e:
                            self.signal_manager.update_log.emit(f"참가자 목록 표시 중 오류: {str(e)}")
                        
                        # 안전한 트래킹 실행
                        try:
                            if hasattr(self.text_extractor, 'tracking_enabled') and self.text_extractor.tracking_enabled:
                                new_participants, left_participants = self.text_extractor.track_participant_changes(participants)
                                # 트래킹 결과는 이미 로그에 출력되므로 추가 처리 불필요
                        except Exception as tracking_error:
                            self.signal_manager.update_log.emit(f"⚠️ 트래킹 처리 중 오류: {str(tracking_error)}")
                            # 트래킹 오류가 전체 프로그램을 중단시키지 않도록 함
                            import traceback
                            self.signal_manager.update_log.emit(f"트래킹 오류 상세: {traceback.format_exc()}")
                    else:
                        self.signal_manager.update_participant_list.emit("\n참가자 목록을 가져올 수 없습니다.")
                        self.signal_manager.update_log.emit("❌ 참가자 목록이 비어있습니다.")
                else:
                    self.signal_manager.update_participant_list.emit("\n참가자 목록 추출에 실패했습니다.")
                    self.signal_manager.update_log.emit("❌ 참가자 목록 추출 결과가 올바르지 않습니다.")
            else:
                self.signal_manager.update_participant_list.emit("\nZoom 창을 찾을 수 없습니다.")
                self.signal_manager.update_log.emit("❌ Zoom 창을 찾을 수 없습니다.")
                
        except Exception as e:
            self.signal_manager.update_log.emit(f"❌ 참가자 목록 가져오기 중 치명적 오류 발생: {str(e)}")
            import traceback
            self.signal_manager.update_log.emit(f"치명적 오류 상세: {traceback.format_exc()}")
        finally:
            try:
                self.extraction_in_progress = False
                self.set_buttons_enabled(True)
            except Exception as e:
                self.signal_manager.update_log.emit(f"상태 초기화 중 오류: {str(e)}")
    
    def copy_to_clipboard(self):
        """최신 참가자 목록을 클립보드에 복사합니다."""
        try:
            if hasattr(self, 'current_participants') and self.current_participants:
                # 현재 시간 기준 헤더 생성
                from datetime import datetime
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                # 클립보드용 텍스트 생성
                clipboard_text = f"[{now}] 참가자 목록 ({len(self.current_participants)}명):\n"
                
                # 참가자 목록 추가 (번호 없이)
                for participant in self.current_participants:
                    clipboard_text += f"{participant}\n"
                
                # 중복 정보 추가
                if hasattr(self, 'current_duplicate_info') and self.current_duplicate_info:
                    clipboard_text += "\n=== 중복 정보 ===\n"
                    for name, info in self.current_duplicate_info.items():
                        clipboard_text += f"• '{name}': {info['count']}번 발견\n"
                        if isinstance(info['details'], str):
                            clipboard_text += f"  → 상태: {info['details']}\n"
                        elif isinstance(info['details'], list):
                            clipboard_text += f"  → 상태: {info['details'][0]}\n"
                
                # 클립보드에 복사
                try:
                    import win32clipboard
                    import win32con
                    
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(clipboard_text, win32con.CF_UNICODETEXT)
                    win32clipboard.CloseClipboard()
                    
                    self.signal_manager.update_log.emit("✅ 참가자 목록이 복사되었습니다.")
                except Exception as clipboard_error:
                    # 실패 시 PyQt 방식으로 다시 시도
                    try:
                        clipboard = QApplication.clipboard()
                        clipboard.setText(clipboard_text)
                        self.signal_manager.update_log.emit("✅ 참가자 목록이 복사되었습니다.")
                    except Exception as qt_error:
                        self.signal_manager.update_log.emit(f"❌ 참가자 목록 복사 중 오류: {str(qt_error)}")
                    
            else:
                self.signal_manager.update_log.emit("⚠️ 복사할 참가자 목록이 없습니다. 먼저 목록을 가져오세요.")
        except Exception as e:
            self.signal_manager.update_log.emit(f"❌ 클립보드 복사 중 오류: {str(e)}")
    
    def append_to_log(self, message):
        """로그 영역에 메시지 추가"""
        try:
            self.log_area.append(message)
            # 자동 스크롤
            try:
                scrollbar = self.log_area.verticalScrollBar()
                if scrollbar:
                    scrollbar.setValue(scrollbar.maximum())
            except Exception as e:
                # 스크롤 오류는 무시
                pass
        except Exception as e:
            # GUI 업데이트 실패 시 콘솔에 출력
            print(f"로그 업데이트 실패: {e}")
    
    def append_to_participant_list(self, message):
        """참가자 목록 영역에 메시지 추가"""
        try:
            self.participant_area.append(message)
            # 자동 스크롤
            try:
                scrollbar = self.participant_area.verticalScrollBar()
                if scrollbar:
                    scrollbar.setValue(scrollbar.maximum())
            except Exception as e:
                # 스크롤 오류는 무시
                pass
        except Exception as e:
            # GUI 업데이트 실패 시 콘솔에 출력
            print(f"참가자 목록 업데이트 실패: {e}")
    
    def set_buttons_enabled(self, enabled):
        """버튼 활성화/비활성화"""
        try:
            self.refresh_button.setEnabled(enabled)
            self.copy_button.setEnabled(enabled)
        except Exception as e:
            # 버튼 상태 변경 실패 시 콘솔에 출력
            print(f"버튼 상태 변경 실패: {e}")

    def show_log_context_menu(self, pos):
        """로그 영역에 컨텍스트 메뉴를 표시합니다."""
        menu = QMenu(self.log_area)
        
        # 기본 복사 액션 추가
        copy_action = QAction("복사", self.log_area)
        copy_action.triggered.connect(self.copy_selected_log)
        menu.addAction(copy_action)
        
        # 전체 복사 액션 추가
        copy_all_action = QAction("전체 복사", self.log_area)
        copy_all_action.triggered.connect(self.copy_all_log)
        menu.addAction(copy_all_action)
        
        # 구분선 추가
        menu.addSeparator()
        
        # 기본 컨텍스트 메뉴 액션들 추가
        select_all_action = QAction("전체 선택", self.log_area)
        select_all_action.triggered.connect(self.log_area.selectAll)
        menu.addAction(select_all_action)
        
        menu.exec(self.log_area.mapToGlobal(pos))

    def copy_selected_log(self):
        """선택된 로그 텍스트를 클립보드에 복사합니다."""
        try:
            cursor = self.log_area.textCursor()
            selected_text = cursor.selectedText()
            if selected_text:
                # 줄바꿈 처리: QTextEdit의 selectedText()는 줄바꿈을 \u2029로 반환하므로 \n으로 변환
                selected_text = selected_text.replace('\u2029', '\n')
                
                # 윈도우 API를 사용한 클립보드 접근
                try:
                    import win32clipboard
                    import win32con
                    
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(selected_text, win32con.CF_UNICODETEXT)
                    win32clipboard.CloseClipboard()
                    
                    self.signal_manager.update_log.emit("✅ 선택된 로그가 복사되었습니다.")
                except Exception as clipboard_error:
                    # 실패 시 PyQt 방식으로 다시 시도
                    clipboard = QApplication.clipboard()
                    clipboard.setText(selected_text)
                    self.signal_manager.update_log.emit("✅ 선택된 로그가 복사되었습니다.")
            else:
                self.signal_manager.update_log.emit("⚠️ 복사할 텍스트가 선택되지 않았습니다.")
        except Exception as e:
            self.signal_manager.update_log.emit(f"❌ 로그 복사 중 오류: {str(e)}")

    def copy_all_log(self):
        """로그 영역의 모든 텍스트를 클립보드에 복사합니다."""
        try:
            all_text = self.log_area.toPlainText()
            if all_text:
                # 윈도우 API를 사용한 클립보드 접근
                try:
                    import win32clipboard
                    import win32con
                    
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(all_text, win32con.CF_UNICODETEXT)
                    win32clipboard.CloseClipboard()
                    
                    self.signal_manager.update_log.emit("✅ 모든 로그가 복사되었습니다.")
                except Exception as clipboard_error:
                    # 실패 시 PyQt 방식으로 다시 시도
                    clipboard = QApplication.clipboard()
                    clipboard.setText(all_text)
                    self.signal_manager.update_log.emit("✅ 모든 로그가 복사되었습니다.")
            else:
                self.signal_manager.update_log.emit("⚠️ 복사할 로그가 없습니다.")
        except Exception as e:
            self.signal_manager.update_log.emit(f"❌ 전체 로그 복사 중 오류: {str(e)}")

def main():
    try:
        print("프로그램 시작...")
        app = QApplication(sys.argv)
        print("QApplication 생성 완료")
        
        gui = ZoomCheckGUI()
        print("GUI 생성 완료")
        
        gui.show()
        print("GUI 표시 완료")
        
        print("이벤트 루프 시작...")
        sys.exit(app.exec())  # PyQt6에서는 exec_() 대신 exec() 사용
    except Exception as e:
        print(f"프로그램 실행 중 치명적 오류: {e}")
        import traceback
        print(f"오류 상세: {traceback.format_exc()}")
        input("엔터를 눌러 종료하세요...")
        sys.exit(1)

if __name__ == "__main__":
    main() 