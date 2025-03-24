# src/text_extractor.py
import logging
from datetime import datetime
import time
import win32gui
import win32con
import win32ui
from ctypes import create_unicode_buffer, sizeof
import re
import pythoncom
from pywinauto import Application, timings
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.keyboard import send_keys
from src.window_finder import WindowFinder
import win32api
import win32clipboard

# Logger 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class TextExtractor:
    def __init__(self):
        self.logger = logger
        try:
            pythoncom.CoInitialize()
        except:
            pass

    def __del__(self):
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    def get_current_time(self):
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def clean_participant_name(self, raw_text):
        """참가자 이름만 깔끔하게 추출"""
        try:
            # 콤마로 구분된 첫 번째 부분만 가져오기
            name = raw_text.split(',')[0].strip()
            
            # (호스트), (나) 등의 태그 제거
            name = name.split('(')[0].strip()
            
            # 앞에 붙은 # 제거
            if name.startswith('#'):
                name = name[1:].strip()
                
            return name
        except:
            return raw_text

    def get_clipboard_text(self):
        """클립보드의 텍스트를 가져옵니다."""
        try:
            win32clipboard.OpenClipboard()
            data = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
            return data
        except Exception as e:
            self.logger.error(f"클립보드 읽기 오류: {str(e)}")
            return ""

    def extract_participants(self, window_handle):
        """Zoom 참가자 창에서 참가자 목록을 추출합니다."""
        try:
            self.logger.info(f"창 핸들 {window_handle}에 연결 시도")
            
            # 창을 활성화하고 포커스를 줍니다
            if win32gui.IsWindow(window_handle):
                win32gui.SetForegroundWindow(window_handle)
                time.sleep(0.5)

            # pywinauto로 창에 연결
            app = Application(backend='uia').connect(handle=window_handle)
            window = app.window(handle=window_handle)
            
            self.logger.info(f"참가자 창 이름: {window.window_text()}")
            self.logger.info("참가자 목록 검색 중...")

            participants = []  # 순서 유지를 위해 list 사용
            seen_participants = set()  # 중복 체크용
            last_count = 0
            scroll_attempts = 0
            max_scroll_attempts = 10

            # 먼저 맨 위로 스크롤
            for _ in range(max_scroll_attempts):
                send_keys('{PGUP}')
                time.sleep(0.1)

            while scroll_attempts < max_scroll_attempts:
                # 현재 보이는 참가자들을 가져옵니다
                for child in window.descendants():
                    try:
                        text = child.window_text()
                        if text and "컴퓨터 오디오" in text and "비디오" in text:
                            name = text.split(',')[0].strip()
                            if name and not any(x in name.lower() for x in ['검색', '초대', '총 참가자']):
                                if name not in seen_participants:  # 중복 체크
                                    participants.append(name)  # 순서대로 추가
                                    seen_participants.add(name)  # 중복 체크용 set에 추가
                                    self.logger.info(f"참가자 발견: {name}")
                    except Exception as e:
                        continue

                current_count = len(participants)
                
                # 새로운 참가자가 발견되지 않으면 스크롤을 시도합니다
                if current_count == last_count:
                    scroll_attempts += 1
                else:
                    scroll_attempts = 0  # 새 참가자가 발견되면 카운터 리셋
                
                last_count = current_count

                # Page Down 키를 눌러 스크롤합니다
                try:
                    window.set_focus()
                    send_keys('{PGDN}')
                    time.sleep(0.5)  # 스크롤 후 잠시 대기
                except Exception as e:
                    self.logger.error(f"스크롤 중 오류: {str(e)}")

            # 마지막으로 맨 위로 스크롤
            for _ in range(max_scroll_attempts):
                send_keys('{PGUP}')
                time.sleep(0.1)
            
            self.logger.info(f"총 {len(participants)}명의 참가자를 찾았습니다")
            self.logger.info(f"참가자 목록: {participants}")
            
            return participants

        except Exception as e:
            self.logger.error(f"참가자 목록 추출 중 오류 발생: {str(e)}")
            return []

def connect_to_window(hwnd):
    """창 핸들을 통해 pywinauto Application에 연결
    
    Args:
        hwnd (int): 윈도우 핸들
        
    Returns:
        pywinauto.application.Application: 연결된 애플리케이션 객체
    """
    try:
        # 프로세스 ID 가져오기
        process_id = WindowFinder.get_process_id_from_window(hwnd)
        
        # pywinauto로 애플리케이션 연결
        app = Application(backend="uia").connect(process=process_id)
        logger.info(f"프로세스 ID {process_id}에 성공적으로 연결했습니다.")
        return app
    except Exception as e:
        logger.error(f"애플리케이션 연결 중 오류 발생: {e}")
        return None

def find_participant_list_control(app, hwnd):
    """참가자 목록 컨트롤 찾기
    
    Args:
        app (pywinauto.application.Application): 연결된 애플리케이션
        hwnd (int): 윈도우 핸들
        
    Returns:
        pywinauto.controls.uia_controls.ListControl: 참가자 목록 컨트롤
    """
    try:
        # 창에 연결
        window = app.window(handle=hwnd)
        
        # 참가자 목록 컨트롤 찾기 시도
        for attempt in range(3):
            try:
                # 방법 1: ListView 컨트롤 찾기
                participant_list = window.child_window(control_type="List")
                logger.info("List 컨트롤을 찾았습니다.")
                return participant_list
            except ElementNotFoundError:
                try:
                    # 방법 2: ListItem 컨테이너 찾기
                    participant_list = window.child_window(control_type="Custom", class_name="ListItemContainer")
                    logger.info("ListItemContainer를 찾았습니다.")
                    return participant_list
                except ElementNotFoundError:
                    # 방법 3: UI 구조 탐색
                    logger.debug(f"시도 {attempt+1}: 컨트롤 패턴 매칭 시도 중...")
                    # 모든 컨트롤 확인
                    for control in window.descendants():
                        control_name = control.friendly_class_name().lower()
                        if "list" in control_name:
                            logger.info(f"패턴 매칭으로 컨트롤 찾음: {control_name}")
                            return control
            
            # 잠시 대기 후 다시 시도
            time.sleep(0.5)
        
        # 시각적 디버깅을 위한 정보 출력
        logger.warning("참가자 목록 컨트롤을 찾을 수 없습니다. UI 구조 덤프:")
        window.print_control_identifiers(depth=3)
        return None
        
    except Exception as e:
        logger.error(f"컨트롤 찾기 중 오류 발생: {e}")
        return None

def extract_participants_from_control(list_control):
    """컨트롤에서 참가자 목록 추출
    
    Args:
        list_control: 참가자 목록 컨트롤
        
    Returns:
        list: 참가자 이름 목록
    """
    if not list_control:
        logger.error("유효한 컨트롤이 아닙니다.")
        return []
    
    participants = []
    
    try:
        # 방법 1: items() 메소드 사용
        try:
            items = list_control.items()
            for item in items:
                item_text = item.text()
                if item_text:
                    participants.append(item_text)
            logger.info(f"items() 메소드로 {len(participants)}명의 참가자를 찾았습니다.")
        except (AttributeError, NotImplementedError) as e:
            logger.debug(f"items() 메소드 실패: {e}")
            
            # 방법 2: children() 메소드 사용
            if not participants:
                children = list_control.children()
                for child in children:
                    if hasattr(child, 'texts') and child.texts():
                        text = child.texts()[0]
                        if text and text.strip():
                            participants.append(text.strip())
                logger.info(f"children() 메소드로 {len(participants)}명의 참가자를 찾았습니다.")
            
            # 방법 3: window_text() 메소드 사용
            if not participants:
                list_text = list_control.window_text()
                if list_text:
                    # 줄바꿈으로 분리하여 각 참가자 추출
                    participants = [line.strip() for line in list_text.split('\n') if line.strip()]
                    logger.info(f"window_text() 메소드로 {len(participants)}명의 참가자를 찾았습니다.")
        
        # 참가자 정보 정리 (이모티콘, 호스트 표시 등 제거)
        cleaned_participants = []
        for participant in participants:
            # 호스트, 공동호스트 표시 제거
            clean_name = participant.replace('(호스트)', '').replace('(Host)', '')
            clean_name = clean_name.replace('(공동호스트)', '').replace('(Co-Host)', '')
            # 앞뒤 공백 제거
            clean_name = clean_name.strip()
            if clean_name:
                cleaned_participants.append(clean_name)
        
        return cleaned_participants
        
    except Exception as e:
        logger.error(f"참가자 추출 중 오류 발생: {e}")
        return []

def extract_participants(hwnd):
    """창 핸들로부터 참가자 목록 추출
    
    Args:
        hwnd (int): 윈도우 핸들
        
    Returns:
        list: 참가자 이름 목록
    """
    # 창 포커스
    WindowFinder.focus_window(hwnd)
    time.sleep(0.5)  # UI가 반응할 시간 주기
    
    # 애플리케이션 연결
    app = connect_to_window(hwnd)
    if not app:
        return []
    
    # 참가자 목록 컨트롤 찾기
    list_control = find_participant_list_control(app, hwnd)
    if not list_control:
        return []
    
    # 참가자 추출
    return extract_participants_from_control(list_control)