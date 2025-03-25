# src/window_finder.py
import win32gui
import win32con
import win32process
import logging
import re
from pywinauto import Application
from pywinauto.findwindows import find_windows

# Logger 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class WindowFinder:
    @staticmethod
    def find_zoom_window():
        """Zoom 참가자 창을 찾습니다."""
        try:
            def callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if "참가자" in title and "Zoom" not in title:  # "Zoom 참가자 체크"가 아닌 창을 찾음
                        windows.append(hwnd)

            windows = []
            win32gui.EnumWindows(callback, windows)
            
            if windows:
                hwnd = windows[0]
                title = win32gui.GetWindowText(hwnd)
                logger.info(f"*** Zoom 참가자 창 발견!: '{title}' ***")
                rect = win32gui.GetWindowRect(hwnd)
                logger.info(f"창 위치: {rect}")
                return hwnd
            else:
                logger.error("Zoom 참가자 창을 찾을 수 없습니다.")
                return None

        except Exception as e:
            logger.error(f"Zoom 창 찾기 중 오류 발생: {e}")
            return None

    def print_all_windows(self):
        """디버깅용: 모든 창 정보를 출력합니다."""
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if window_text:
                    self.logger.info(f"창: '{window_text}' (핸들: {hwnd})")
            return True
        
        self.logger.info("=== 현재 열린 모든 창 목록 ===")
        win32gui.EnumWindows(callback, None)
        self.logger.info("===========================")

    @staticmethod
    def get_process_id_from_window(hwnd):
        """주어진 핸들에서 프로세스 ID를 가져옵니다."""
        try:
            app = Application().connect(handle=hwnd)
            process_id = app.process
            return process_id
        except Exception as e:
            logger.error(f"프로세스 ID 가져오기 중 오류 발생: {e}")
            return None

def focus_window(hwnd):
    """창을 최상위로 가져오는 함수
    
    Args:
        hwnd (int): 윈도우 핸들
        
    Returns:
        bool: 성공 여부
    """
    if not hwnd:
        logger.error("유효하지 않은 핸들입니다.")
        return False
    
    try:
        # 최소화된 창이면 복원
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        
        # 창을 최상위로 가져오기
        win32gui.SetForegroundWindow(hwnd)
        logger.debug(f"핸들 {hwnd}의 창을 최상위로 가져왔습니다.")
        return True
    except Exception as e:
        logger.error(f"창 포커스 중 오류 발생: {e}")
        return False

def get_window_rect(hwnd):
    """창의 위치와 크기를 가져오는 함수
    
    Args:
        hwnd (int): 윈도우 핸들
        
    Returns:
        tuple: (left, top, right, bottom) 창의 위치와 크기
    """
    try:
        rect = win32gui.GetWindowRect(hwnd)
        logger.debug(f"창 위치: {rect}")
        return rect
    except Exception as e:
        logger.error(f"창 위치 가져오기 중 오류 발생: {e}")
        return None

def get_window_text(hwnd):
    """창의 제목을 가져오는 함수
    
    Args:
        hwnd (int): 윈도우 핸들
        
    Returns:
        str: 창 제목
    """
    try:
        title = win32gui.GetWindowText(hwnd)
        logger.debug(f"창 제목: {title}")
        return title
    except Exception as e:
        logger.error(f"창 제목 가져오기 중 오류 발생: {e}")
        return ""

def copy_to_clipboard(self):
    """최신 참가자 목록을 클립보드에 복사합니다."""
    try:
        # 디버깅 정보 출력
        self.signal_manager.update_log.emit(f"클립보드 복사 시도: current_participants 길이 = {len(self.current_participants)}")
        
        if self.current_participants:
            # ... 기존 코드 ...
            
            # 클립보드에 복사 시도
            try:
                clipboard = QApplication.clipboard()
                self.signal_manager.update_log.emit("QApplication.clipboard() 호출 성공")
                clipboard.setText(clipboard_text)
                self.signal_manager.update_log.emit("clipboard.setText() 호출 성공")
                
                # 복사 확인
                copied_text = clipboard.text()
                if copied_text:
                    self.signal_manager.update_log.emit(f"클립보드 텍스트 확인 - 길이: {len(copied_text)} 문자")
                else:
                    self.signal_manager.update_log.emit("클립보드에 텍스트가 없습니다!")
                
                self.signal_manager.update_participant_list.emit("\n클립보드에 참가자 목록이 복사되었습니다.")
            except Exception as e:
                self.signal_manager.update_log.emit(f"클립보드 객체 사용 중 오류: {str(e)}")
                
        else:
            self.signal_manager.update_log.emit("복사할 참가자 목록이 없습니다.")
            self.signal_manager.update_participant_list.emit("\n복사할 참가자 목록이 없습니다. 먼저 목록을 가져오세요.")
    except Exception as e:
        self.signal_manager.update_log.emit(f"\n클립보드 복사 중 오류 발생: {str(e)}")
        self.signal_manager.update_participant_list.emit(f"\n클립보드 복사 중 오류 발생: {str(e)}")
        import traceback
        self.signal_manager.update_log.emit(traceback.format_exc())