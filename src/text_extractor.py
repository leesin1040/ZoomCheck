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
import queue
from pywinauto.mouse import scroll as mouse_scroll

# Logger 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class TextExtractor:
    def __init__(self):
        self.logger = logger
        self.extraction_queue = queue.Queue()
        self.seen_participants = set()
        self.participants_list = []
        self.extraction_active = False
        self.stop_scrolling = False
        self._should_stop = False  # 중지 플래그 추가
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

    def stop_extraction(self):
        """참가자 추출을 중단합니다."""
        self.logger.info("참가자 추출 중단 요청 받음")
        self._should_stop = True  # 중지 플래그 설정

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
                time.sleep(0.5)  # 0.5초마다 스크롤
        except:
            pass

    def extract_participants(self, window_handle):
        """Zoom 참가자 창에서 참가자 목록을 추출합니다."""
        try:
            # COM 초기화 (필요할 경우)
            try:
                pythoncom.CoInitialize()
            except:
                pass

            # 중지 플래그 초기화
            self._should_stop = False

            # 시작 시간 기록
            start_time = time.time()
            self.logger.info(f"[시간 측정] 참가자 추출 시작: {datetime.now().strftime('%H:%M:%S.%f')[:-3]}")
            
            self.logger.info(f"창 핸들 {window_handle}에 연결 시도")
            
            # 창 유효성 검사
            if not win32gui.IsWindow(window_handle):
                self.logger.error("유효하지 않은 창 핸들입니다.")
                return []
            
            # 창이 보이는지 확인
            if not win32gui.IsWindowVisible(window_handle):
                self.logger.error("창이 보이지 않습니다.")
                return []
            
            # 창을 활성화
            try:
                win32gui.SetForegroundWindow(window_handle)
                time.sleep(0.1)  # 활성화 대기 시간
            except Exception as e:
                self.logger.warning(f"창 활성화 실패: {e}")

            # pywinauto로 창에 연결
            try:
                app = Application(backend='uia').connect(handle=window_handle)
                window = app.window(handle=window_handle)
            except Exception as e:
                self.logger.error(f"pywinauto 연결 실패: {e}")
                return []
            
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

            # 처음에 Home 키를 눌러 맨 위로 스크롤
            try:
                window.set_focus()
                send_keys('{HOME}')
                time.sleep(1.0)
            except Exception as e:
                self.logger.warning(f"초기 스크롤 실패: {e}")
            
            # 참가자 수에 따른 대기 시간 조정
            from src.common.constants import (
                LARGE_PARTICIPANT_THRESHOLD, MEDIUM_PARTICIPANT_THRESHOLD,
                LARGE_PARTICIPANT_INITIAL_DELAY, NORMAL_PARTICIPANT_INITIAL_DELAY,
                LARGE_PARTICIPANT_SCROLL_ATTEMPTS, LARGE_PARTICIPANT_SCROLL_DELAY, LARGE_PARTICIPANT_CONSECUTIVE_LIMIT,
                MEDIUM_PARTICIPANT_SCROLL_ATTEMPTS, MEDIUM_PARTICIPANT_SCROLL_DELAY, MEDIUM_PARTICIPANT_CONSECUTIVE_LIMIT,
                NORMAL_PARTICIPANT_SCROLL_ATTEMPTS, NORMAL_PARTICIPANT_SCROLL_DELAY, NORMAL_PARTICIPANT_CONSECUTIVE_LIMIT,
                PROGRESS_LOG_INTERVAL
            )
            
            if total_expected > MEDIUM_PARTICIPANT_THRESHOLD:
                time.sleep(LARGE_PARTICIPANT_INITIAL_DELAY)  # 대규모 참가자일 경우 더 오래 대기
            else:
                time.sleep(NORMAL_PARTICIPANT_INITIAL_DELAY)  # 기존 대기 시간

            # 추출 시간 측정 시작
            extraction_start = time.time()
            
            participants = []
            seen_participants = set()
            
            # 참가자 수에 따른 스크롤 설정 조정
            if total_expected > LARGE_PARTICIPANT_THRESHOLD:
                max_scroll_attempts = LARGE_PARTICIPANT_SCROLL_ATTEMPTS
                scroll_delay = LARGE_PARTICIPANT_SCROLL_DELAY
                consecutive_limit = LARGE_PARTICIPANT_CONSECUTIVE_LIMIT
            elif total_expected > MEDIUM_PARTICIPANT_THRESHOLD:
                max_scroll_attempts = MEDIUM_PARTICIPANT_SCROLL_ATTEMPTS
                scroll_delay = MEDIUM_PARTICIPANT_SCROLL_DELAY
                consecutive_limit = MEDIUM_PARTICIPANT_CONSECUTIVE_LIMIT
            else:
                # 소규모 참가자(50명 이하)는 더 빠르게 처리
                if total_expected <= 50:
                    max_scroll_attempts = 20  # 매우 적은 스크롤
                    scroll_delay = 0.5        # 빠른 대기
                    consecutive_limit = 2     # 빠른 종료
                else:
                    max_scroll_attempts = NORMAL_PARTICIPANT_SCROLL_ATTEMPTS
                    scroll_delay = NORMAL_PARTICIPANT_SCROLL_DELAY
                    consecutive_limit = NORMAL_PARTICIPANT_CONSECUTIVE_LIMIT
            
            # 스크롤 관련 변수
            scroll_count = 0
            consecutive_same_view = 0
            prev_count = 0
            last_progress_time = time.time()
            error_count = 0  # 오류 카운터 추가

            def do_scroll(window):
                try:
                    for _ in range(2):
                        window.set_focus()
                        time.sleep(0.1)
                    send_keys('{PGDN 2}')
                    time.sleep(0.5)
                    return True
                except Exception as e:
                    self.logger.warning(f"스크롤 실패: {e}")
                    return False

            self.logger.info(f"스크롤 설정: 최대 {max_scroll_attempts}회, 대기시간 {scroll_delay}초, 연속제한 {consecutive_limit}회")

            while scroll_count < max_scroll_attempts and consecutive_same_view < consecutive_limit:
                if self._should_stop:
                    self.logger.info("중지 요청으로 추출 중단")
                    break
                
                scroll_count += 1
                current_batch = []
                
                # 진행 상황 로깅 (설정된 간격마다)
                current_time = time.time()
                if current_time - last_progress_time > PROGRESS_LOG_INTERVAL:
                    self.logger.info(f"진행 상황: {scroll_count}/{max_scroll_attempts} 스크롤, 현재 {len(participants)}명 발견")
                    last_progress_time = current_time
                
                try:
                    all_elements = window.descendants()
                    for child in all_elements:
                        try:
                            text = child.window_text()
                            if not text or len(text.strip()) < 2:
                                continue
                            
                            # 개선된 참가자 패턴 매칭
                            is_participant = False
                            
                            # 먼저 확실히 제외할 UI 요소들 체크
                            exclude_patterns = [
                                # 창 컨트롤
                                '시스템', '최소화', '최대화', '닫기', '복원',
                                # 메뉴 항목
                                '전화 참가자 목록', '초대', '모두 음소거', '모든 참가자 관리',
                                '추가 옵션', '음소거 해제', '음소거', '비디오 시작', '비디오 중지',
                                # 기타 UI 요소
                                '검색', '총 참가자', '모두에게 메시지 보내기', 
                                '참가자', '참석자', 'host', 'co-host', '참가자 목록',
                                'button', 'edit', 'combo', 'list', 'scroll', 'menu', 'toolbar',
                                '상태', '정보', '설정', '옵션', '관리', '보기', '도움말'
                            ]
                            
                            # 제외 패턴에 해당하면 건너뛰기
                            if any(x in text for x in exclude_patterns):
                                continue
                            
                            # 참가자 패턴 매칭
                            # 1. 쉼표가 있는 텍스트 (참가자 이름 + 상태) - 가장 확실한 패턴
                            if ',' in text and len(text.split(',')[0].strip()) > 1:
                                name_part = text.split(',')[0].strip()
                                # 이름 부분이 제외 패턴에 없고, 적절한 길이인 경우
                                if (len(name_part) > 1 and len(name_part) < 50 and 
                                    not any(x in name_part.lower() for x in exclude_patterns)):
                                    is_participant = True
                            
                            # 2. 괄호가 있는 텍스트 (참가자 이름 + 역할)
                            elif '(' in text and ')' in text and len(text.split('(')[0].strip()) > 1:
                                name_part = text.split('(')[0].strip()
                                if (len(name_part) > 1 and len(name_part) < 50 and 
                                    not any(x in name_part.lower() for x in exclude_patterns)):
                                    is_participant = True
                            
                            # 3. 특정 키워드가 포함된 텍스트 (참가자 상태 정보)
                            elif any(pattern in text for pattern in [
                                "컴퓨터 오디오", "비디오", "마이크", "스피커", "카메라", "오디오"
                            ]):
                                # 키워드가 포함되어 있지만, 실제 이름 부분이 있는지 확인
                                if ',' in text:
                                    name_part = text.split(',')[0].strip()
                                    if (len(name_part) > 1 and len(name_part) < 50 and 
                                        not any(x in name_part.lower() for x in exclude_patterns)):
                                        is_participant = True
                            
                            # 4. 단순 텍스트 (참가자 이름일 가능성) - 가장 엄격하게 체크
                            elif (len(text.strip()) > 2 and len(text.strip()) < 30 and 
                                  not any(x in text.lower() for x in exclude_patterns)):
                                # 한글이나 영문이 포함되어 있는지 확인
                                import re
                                if re.search(r'[가-힣a-zA-Z]', text):
                                    is_participant = True
                            
                            if is_participant:
                                # 참가자 이름 추출
                                if ',' in text:
                                    name = text.split(',')[0].strip()
                                elif '(' in text:
                                    name = text.split('(')[0].strip()
                                else:
                                    name = text.strip()
                                
                                # 이름 정리
                                clean_name = name.split('(')[0].strip()
                                if clean_name and len(clean_name) > 1:
                                    current_batch.append(clean_name)
                                    if clean_name not in seen_participants:
                                        participants.append(clean_name)
                                        seen_participants.add(clean_name)
                                        
                        except Exception as e:
                            error_count += 1
                            if error_count <= 5:  # 처음 5개 오류만 로깅
                                self.logger.debug(f"요소 처리 중 오류 (무시됨): {e}")
                            continue
                            
                except Exception as e:
                    error_count += 1
                    self.logger.error(f"요소 처리 중 오류: {str(e)}")
                    if error_count > 10:  # 오류가 너무 많으면 중단
                        self.logger.error("오류가 너무 많아 추출을 중단합니다.")
                        break
                
                # 스크롤 후 충분히 대기
                time.sleep(scroll_delay)
                
                # 동일 참가자 수 반복 체크
                if len(participants) == prev_count:
                    consecutive_same_view += 1
                else:
                    consecutive_same_view = 0
                prev_count = len(participants)
                
                # 조기 종료 조건: 모든 참가자를 찾았거나 충분히 스크롤했을 때
                if total_expected > 0 and len(participants) >= total_expected:
                    self.logger.info(f"모든 참가자를 찾았습니다! ({len(participants)}명)")
                    break
                elif scroll_count >= 5 and len(participants) > 0 and consecutive_same_view >= 2:
                    self.logger.info(f"충분히 스크롤했고 새로운 참가자가 없어 조기 종료합니다. (현재 {len(participants)}명)")
                    break
                
                # 스크롤 실행
                if not do_scroll(window):
                    error_count += 1
                    if error_count > 5:
                        self.logger.error("스크롤 오류가 너무 많아 중단합니다.")
                        break
            
            # 추출 후 안내
            if total_expected > 0:
                if len(participants) < total_expected:
                    missing_count = total_expected - len(participants)
                    missing_percentage = (missing_count / total_expected) * 100
                    self.logger.warning(f"Zoom에 표시된 참가자 수({total_expected})와 추출된 참가자 수({len(participants)})가 다릅니다.")
                    self.logger.warning(f"누락된 참가자: {missing_count}명 ({missing_percentage:.1f}%)")
                    if missing_percentage > 10:
                        self.logger.warning("높은 누락률입니다. 스크롤 횟수나 대기 시간을 늘려보세요.")
                else:
                    self.logger.info(f"모든 참가자를 성공적으로 추출했습니다! ({len(participants)}명)")
            
            total_time = time.time() - start_time
            self.logger.info(f"[시간 측정] 참가자 추출 완료: 총 {len(participants)}명")
            self.logger.info(f"[시간 측정] 총 소요 시간: {total_time:.3f}초")
            if total_time > 0:
                self.logger.info(f"[시간 측정] 평균 처리 속도: {len(participants)/total_time:.1f}명/초")
            
            # 오류 통계
            if error_count > 0:
                self.logger.info(f"처리 중 발생한 오류: {error_count}개")
            
            return participants
        except Exception as e:
            self.logger.error(f"참가자 목록 추출 중 오류 발생: {str(e)}")
            import traceback
            self.logger.error(traceback.format_exc())
            return []
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def _extract_with_ui_elements(self, window, total_expected):
        """UI 요소 탐색 방식으로 참가자 추출 (폴백용)"""
        # 기존의 UI 요소 탐색 코드를 여기에 구현
        # (기존 extract_participants 메서드의 UI 요소 탐색 부분)
        return [], {}

    def extract_text_with_ocr(self, hwnd=None, capture_method='window'):
        """
        OCR을 사용하여 텍스트를 추출합니다.
        """
        if hwnd is None:
            hwnd = self.window_finder.current_hwnd
        
        if not hwnd:
            raise ValueError("윈도우 핸들이 없습니다.")
        
        # 처음에 스크롤을 맨 위로 올립니다
        self.window_finder.scroll_to_top(hwnd)
        time.sleep(0.5)
        
        # ... existing code ... (기존 OCR 처리 코드)
        
        # 텍스트 추출 작업이 끝난 후 다시 스크롤을 맨 위로 올립니다
        self.window_finder.scroll_to_top(hwnd)
        
        return extracted_text

    def extract_text_with_accessibility(self, hwnd=None):
        """
        접근성 API를 사용하여 텍스트를 추출합니다.
        """
        if hwnd is None:
            hwnd = self.window_finder.current_hwnd
        
        if not hwnd:
            raise ValueError("윈도우 핸들이 없습니다.")
        
        # 처음에 스크롤을 맨 위로 올립니다
        self.window_finder.scroll_to_top(hwnd)
        time.sleep(0.5)
        
        # ... existing code ... (기존 텍스트 추출 코드)
        
        # 텍스트 추출 작업이 끝난 후 다시 스크롤을 맨 위로 올립니다
        self.window_finder.scroll_to_top(hwnd)
        
        return text

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