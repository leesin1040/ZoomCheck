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
import threading

# Logger 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class TextExtractor:
    def __init__(self):
        self.logger = logger
        self.stop_scrolling = False
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

    def background_scroll(self, window):
        """백그라운드에서 자동 스크롤을 수행하는 함수"""
        try:
            for _ in range(30):  # 최대 30번 스크롤
                if self.stop_scrolling:
                    break
                try:
                    window.set_focus()
                    send_keys('{PGDN}')
                except:
                    pass
                time.sleep(0.1)  # 0.1초마다 스크롤
        except:
            pass

    def extract_participants(self, window_handle):
        """Zoom 참가자 창에서 참가자 목록을 추출합니다."""
        try:
            # 시작 시간 기록
            start_time = time.time()
            self.logger.info(f"[시간 측정] 참가자 추출 시작: {datetime.now().strftime('%H:%M:%S.%f')[:-3]}")
            
            self.logger.info(f"창 핸들 {window_handle}에 연결 시도")
            
            # 창을 활성화
            if win32gui.IsWindow(window_handle):
                win32gui.SetForegroundWindow(window_handle)
                time.sleep(0.3)

            # pywinauto로 창에 연결
            app = Application(backend='uia').connect(handle=window_handle)
            window = app.window(handle=window_handle)
            
            # 창 제목에서 총 참가자 수 추출
            window_title = window.window_text()
            total_expected = 0
            try:
                if '(' in window_title and ')' in window_title:
                    total_str = window_title.split('(')[1].split(')')[0]
                    total_expected = int(total_str)
                    self.logger.info(f"창 제목에서 찾은 총 참가자 수: {total_expected}명")
            except:
                self.logger.info("창 제목에서 참가자 수를 추출할 수 없습니다.")
            
            self.logger.info(f"참가자 창 이름: {window_title}")
            self.logger.info("참가자 목록 검색 중...")

            participants = []
            seen_participants = set()
            last_count = 0
            no_new_count = 0  # 새로운 참가자를 찾지 못한 연속 횟수
            
            # 처음에 Home 키를 눌러 맨 위로 스크롤
            window.set_focus()
            send_keys('{HOME}')
            time.sleep(0.2)

            # 추출 시간 측정 시작
            extraction_start = time.time()
            scroll_count = 0
            total_scroll_time = 0
            
            # 마지막 요소 텍스트를 저장하여 스크롤 끝을 감지
            last_element_text = None
            end_detected = False

            while True:
                scroll_count += 1
                scan_start = time.time()
                prev_count = len(participants)
                
                # 모든 요소를 한번에 가져와서 확인
                try:
                    all_elements = window.descendants()
                    
                    # 마지막 요소의 텍스트 저장 (스크롤 끝 감지용)
                    last_visible_element = None
                    
                    for child in all_elements:
                        try:
                            text = child.window_text()
                            if text and "컴퓨터 오디오" in text and "비디오" in text:
                                last_visible_element = text  # 현재 보이는 마지막 요소 업데이트
                                name = text.split(',')[0].strip()
                                if name and not any(x in name.lower() for x in ['검색', '초대', '총 참가자']):
                                    if name not in seen_participants:
                                        participants.append(name)
                                        seen_participants.add(name)
                        except Exception:
                            continue
                
                    # 마지막 요소를 이용해 스크롤 끝 감지
                    if last_visible_element == last_element_text:
                        no_new_count += 1
                        if no_new_count >= 2:  # 두 번 연속 같은 마지막 요소면 끝으로 간주
                            self.logger.info("스크롤 끝 감지: 마지막 요소가 변하지 않음")
                            end_detected = True
                    else:
                        no_new_count = 0
                        last_element_text = last_visible_element
                
                except Exception as e:
                    self.logger.error(f"요소 처리 중 오류: {str(e)}")
                
                current_count = len(participants)
                if current_count > prev_count:
                    self.logger.info(f"[시간 측정] 스크롤 {scroll_count}: {current_count - prev_count}명 발견 (총 {current_count}명)")
                
                # 모든 참가자를 찾았거나 스크롤 끝에 도달했는지 확인
                if (total_expected > 0 and current_count >= total_expected) or end_detected:
                    self.logger.info(f"참가자 목록 추출 완료 조건 충족: 총 {current_count}명 발견")
                    break
                
                # 새 참가자를 찾지 못한 경우
                if current_count == last_count:
                    no_new_count += 1
                    if no_new_count >= 3:  # 세 번 연속 새 참가자 없으면 종료
                        self.logger.info("새 참가자 없음: 스크롤 종료")
                        break
                else:
                    no_new_count = 0  # 새 참가자 발견 시 카운터 리셋
                
                last_count = current_count

                # 스크롤 시간 측정
                scroll_start = time.time()
                
                # 페이지 다운으로 스크롤
                try:
                    window.set_focus()
                    send_keys('{PGDN}')
                    scroll_time = 0.2  # 스크롤 후 대기 시간 단축
                    time.sleep(scroll_time)
                    total_scroll_time += scroll_time
                except Exception as e:
                    self.logger.error(f"스크롤 중 오류: {str(e)}")

            # 추출 시간 계산
            extraction_time = time.time() - extraction_start
            
            # 마지막으로 맨 위로 스크롤
            window.set_focus()
            send_keys('{HOME}')
            
            # 총 소요 시간 계산
            total_time = time.time() - start_time
            
            self.logger.info(f"[시간 측정] 참가자 추출 완료: 총 {len(participants)}명")
            self.logger.info(f"[시간 측정] 총 소요 시간: {total_time:.3f}초")
            self.logger.info(f"[시간 측정] 스크롤 대기 시간: {total_scroll_time:.3f}초 ({(total_scroll_time/total_time*100):.1f}%)")
            self.logger.info(f"[시간 측정] 추출 시간: {extraction_time:.3f}초 ({(extraction_time/total_time*100):.1f}%)")
            self.logger.info(f"[시간 측정] 평균 참가자 추출 속도: {len(participants)/extraction_time:.1f}명/초")
            
            return participants

        except Exception as e:
            self.logger.error(f"참가자 목록 추출 중 오류 발생: {str(e)}")
            end_time = time.time()
            self.logger.info(f"[시간 측정] 오류로 인한 중단: {end_time - start_time:.3f}초")
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