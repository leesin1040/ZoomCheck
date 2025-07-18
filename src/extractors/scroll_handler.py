# src/extractors/scroll_handler.py
import time
import win32gui
import win32con
import win32api
import logging
from pywinauto.keyboard import send_keys

logger = logging.getLogger(__name__)

class ScrollHandler:
    """스크롤 관련 기능을 처리하는 클래스"""
    
    def __init__(self, scroll_delay=0.8):
        self.scroll_delay = scroll_delay
    
    def scroll_to_top(self, hwnd):
        """스크롤을 제일 위로 올립니다."""
        if not hwnd:
            logger.error("유효하지 않은 핸들입니다.")
            return False
        
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.1)
            win32api.keybd_event(win32con.VK_HOME, 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(win32con.VK_HOME, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.1)
            return True
        except Exception as e:
            logger.error(f"스크롤 상단 이동 중 오류: {e}")
            return False
    
    def scroll_down_and_return_to_top(self, hwnd, scroll_amount=10):
        """스크롤을 아래로 내린 후 다시 맨 위로 올립니다."""
        if not hwnd:
            logger.error("유효하지 않은 핸들입니다.")
            return False
        
        try:
            # 먼저 스크롤을 맨 위로 올림
            self.scroll_to_top(hwnd)
            time.sleep(0.5)
            
            # 스크롤 다운
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.1)
            
            for _ in range(scroll_amount):
                win32api.keybd_event(win32con.VK_PGDN, 0, 0, 0)
                time.sleep(0.05)
                win32api.keybd_event(win32con.VK_PGDN, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.1)
            
            # 다시 맨 위로 스크롤 올림
            self.scroll_to_top(hwnd)
            return True
        except Exception as e:
            logger.error(f"스크롤 다운 후 상단 이동 중 오류: {e}")
            return False
    
    def do_scroll(self, window, method='keyboard'):
        """창에서 스크롤을 수행합니다."""
        try:
            # 포커스를 여러 번 줌
            for _ in range(2):
                window.set_focus()
                time.sleep(0.1)
            
            # 키보드 스크롤만 사용
            send_keys('{PGDN 2}')
            time.sleep(0.5)
            return True
        except Exception as e:
            logger.error(f"스크롤 수행 중 오류: {e}")
            return False
    
    def background_scroll(self, window, max_scrolls=30, stop_flag=None):
        """백그라운드에서 자동 스크롤을 수행하는 함수"""
        try:
            for i in range(max_scrolls):
                if stop_flag and stop_flag():
                    logger.info("스크롤 중단 요청됨")
                    break
                
                try:
                    window.set_focus()
                    send_keys('{PGDN}')
                except Exception as e:
                    logger.warning(f"스크롤 중 오류 (무시됨): {e}")
                
                time.sleep(self.scroll_delay)
            
            logger.info(f"백그라운드 스크롤 완료: {i+1}회 수행")
        except Exception as e:
            logger.error(f"백그라운드 스크롤 중 오류: {e}") 