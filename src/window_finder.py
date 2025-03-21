# src/window_finder.py
import win32gui
import win32con
import win32process
import logging
import re

class WindowFinder:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def find_zoom_window(self):
        """Zoom 참가자 창을 찾아서 핸들을 반환합니다."""
        try:
            def callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    window_text = win32gui.GetWindowText(hwnd)
                    # 디버깅: 모든 창 정보 출력
                    if window_text:
                        self.logger.info(f"발견된 창: '{window_text}'")
                    
                    # "참가자" 또는 "Participants"로 시작하는 창 찾기
                    if (window_text.startswith("참가자") or 
                        window_text.startswith("Participants")):
                        self.logger.info(f"*** Zoom 참가자 창 발견!: '{window_text}' ***")
                        windows.append(hwnd)
                return True

            windows = []
            win32gui.EnumWindows(callback, windows)
            
            if windows:
                selected_window = windows[0]
                window_text = win32gui.GetWindowText(selected_window)
                self.logger.info(f"선택된 Zoom 참가자 창: '{window_text}' (핸들: {selected_window})")
                
                # 창 위치와 크기 정보 출력
                rect = win32gui.GetWindowRect(selected_window)
                self.logger.info(f"창 위치: {rect}")
                
                return selected_window
            else:
                self.logger.warning("Zoom 참가자 창을 찾을 수 없습니다.")
                return None
                
        except Exception as e:
            self.logger.error(f"Zoom 창 검색 중 오류 발생: {str(e)}")
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

def get_process_id_from_window(hwnd):
    """창 핸들로부터 프로세스 ID를 얻는 함수
    
    Args:
        hwnd (int): 윈도우 핸들
        
    Returns:
        int: 프로세스 ID
    """
    _, process_id = win32process.GetWindowThreadProcessId(hwnd)
    logger.debug(f"핸들 {hwnd}의 프로세스 ID: {process_id}")
    return process_id

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