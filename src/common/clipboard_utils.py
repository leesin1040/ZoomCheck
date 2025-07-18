# src/common/clipboard_utils.py
import win32clipboard
import win32con
import logging

logger = logging.getLogger(__name__)

def get_clipboard_text():
    """클립보드의 텍스트를 가져옵니다."""
    try:
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return data
    except Exception as e:
        logger.error(f"클립보드 읽기 오류: {str(e)}")
        return ""

def set_clipboard_text(text):
    """클립보드에 텍스트를 설정합니다."""
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32con.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return True
    except Exception as e:
        logger.error(f"클립보드 쓰기 오류: {str(e)}")
        return False 