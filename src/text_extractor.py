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
            # COM 초기화 (필요할 경우)
            try:
                pythoncom.CoInitialize()
            except:
                pass

            # 시작 시간 기록
            start_time = time.time()
            self.logger.info(f"[시간 측정] 참가자 추출 시작: {datetime.now().strftime('%H:%M:%S.%f')[:-3]}")
            
            self.logger.info(f"창 핸들 {window_handle}에 연결 시도")
            
            # 창을 활성화
            if win32gui.IsWindow(window_handle):
                win32gui.SetForegroundWindow(window_handle)
                time.sleep(0.1)  # 활성화 대기 시간

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

            # 처음에 Home 키를 눌러 맨 위로 스크롤
            window.set_focus()
            send_keys('{HOME}')
            time.sleep(0.2)  # 홈 키 적용 대기

            # 추출 시간 측정 시작
            extraction_start = time.time()
            
            participants = []
            seen_participants = set()
            scroll_count = 0
            total_scroll_time = 0
            
            # 스크롤 관련 변수
            max_scroll_attempts = 60  # 스크롤 제한을 60회로 설정
            no_new_count = 0
            consecutive_same_view = 0
            last_batch = []
            
            # 스크롤 딜레이 조정 - 속도 개선
            scroll_delay = 0.18  # 초기값 (이전보다 약간 빠르게)
            
            while scroll_count < max_scroll_attempts and consecutive_same_view < 3:
                scroll_count += 1
                current_batch = []
                scroll_start = time.time()
                
                # 현재 화면에서 참가자 추출
                try:
                    all_elements = window.descendants()
                    
                    for child in all_elements:
                        try:
                            text = child.window_text()
                            if text and ("컴퓨터 오디오" in text or "비디오" in text):
                                name = text.split(',')[0].strip()
                                if name and not any(x in name.lower() for x in ['검색', '초대', '총 참가자', '모두에게 메시지 보내기']):
                                    current_batch.append(name)
                                    if name not in seen_participants:
                                        participants.append(name)
                                        seen_participants.add(name)
                        except Exception:
                            continue
                
                except Exception as e:
                    self.logger.error(f"요소 처리 중 오류: {str(e)}")
                
                # 발견된 참가자 확인
                current_count = len(participants)
                found_now = len(current_batch)
                
                # 현재 배치가 이전과 동일한지 확인
                current_batch_set = set(current_batch)
                if current_batch and current_batch_set == set(last_batch):
                    consecutive_same_view += 1
                    # 2번 연속 동일하면 스크롤 딜레이 증가
                    if consecutive_same_view >= 2:
                        scroll_delay = min(0.2, scroll_delay + 0.02)  # 딜레이 증가
                    
                    if consecutive_same_view >= 3:
                        self.logger.info("스크롤 끝 감지: 마지막 요소가 변하지 않음")
                else:
                    consecutive_same_view = 0
                    
                # 마지막 배치 업데이트
                last_batch = current_batch
                
                # 새 참가자 발견 여부 로깅
                prev_count = current_count - found_now
                
                if current_count > prev_count and found_now > 0:
                    self.logger.info(f"[시간 측정] 스크롤 {scroll_count}: {found_now}명 발견 (총 {current_count}명)")
                    no_new_count = 0
                    
                    # 발견 속도가 좋으면 스크롤 딜레이 유지/감소
                    if found_now >= 20:
                        scroll_delay = max(0.15, scroll_delay - 0.01)  # 딜레이 약간 감소
                else:
                    no_new_count += 1
                    # 새 참가자를 찾지 못하면 스크롤 딜레이 증가
                    if no_new_count >= 2:
                        scroll_delay = min(0.2, scroll_delay + 0.01)  # 딜레이 증가
                
                # 모든 참가자를 찾았는지 확인
                if total_expected > 0 and current_count >= total_expected:
                    self.logger.info(f"참가자 목록 추출 완료: 모든 참가자 {total_expected}명 찾음")
                    break
                
                # 더 이상 찾지 못하는 경우 전략 변경
                if no_new_count >= 3:
                    if no_new_count == 3:
                        self.logger.info("연속 3번 새 참가자 없음: 대형 스크롤 시도")
                    
                    # 진행률 확인
                    progress = current_count / total_expected if total_expected > 0 else 0
                    
                    # 진행률에 따라 다른 전략 시도
                    if progress < 0.6 and scroll_count <= 40:
                        # 절반 조금 더 넘은 수준이라면 대형 스크롤 시도
                        window.set_focus()
                        send_keys('{PGDN 2}')  # 2페이지 점프
                        time.sleep(0.25)  # 대형 스크롤은 더 길게 대기
                        
                        # 추가 전략: 다시 위로 올라가서 재시도
                        if no_new_count >= 5 and scroll_count >= 20:
                            self.logger.info("새로운 전략: 맨 위로 이동 후 천천히 스크롤")
                            window.set_focus()
                            send_keys('{HOME}')
                            time.sleep(0.3)
                            
                            # 스크롤 딜레이 초기화 및 증가
                            scroll_delay = 0.2
                            no_new_count = 0
                    else:
                        # 절반 이상 찾았거나, 충분히 스크롤했다면
                        window.set_focus()
                        send_keys('{PGDN}')
                        time.sleep(0.2)
                    
                    # 매우 오랫동안 새 참가자를 찾지 못하면 종료 고려
                    if no_new_count >= 7:
                        if progress >= 0.3 or scroll_count >= 40:
                            self.logger.info(f"참가자 목록 추출 완료 조건 충족: 총 {current_count}명 발견 (진행률: {progress*100:.1f}%)")
                            break
                else:
                    # 일반 스크롤
                    window.set_focus()
                    
                    # 10번마다 대형 스크롤하되, 적절한 대기 시간 적용
                    if scroll_count % 10 == 0:
                        self.logger.info(f"대형 스크롤 시도 (10회 단위)")
                        send_keys('{PGDN 2}')  # 2페이지 점프
                        time.sleep(0.25)  # 충분한 로딩 시간
                    else:
                        # 20회에서 30회 사이는 DOWN 키로 더 작게 스크롤 (오버랩 많이)
                        if 20 <= scroll_count <= 30:
                            send_keys('{DOWN 20}')  # 약간만 스크롤 (더 많은 오버랩)
                        else:
                            send_keys('{PGDN}')  # 일반 스크롤
                        
                        # 스크롤 딜레이는 현재 설정된 값 사용
                        time.sleep(scroll_delay)
                
                total_scroll_time += scroll_delay
            
            # 추출 시간 계산
            extraction_time = time.time() - extraction_start
            
            # 추출 완료 후 추가 검사 (누락이 있으면)
            if total_expected > 0 and len(participants) < total_expected * 0.95:
                self.logger.info(f"누락된 참가자가 많음: 추가 검사 실행 중... ({len(participants)}/{total_expected}명)")
                
                # 맨 아래로 스크롤한 후 위로 올라오면서 확인
                window.set_focus()
                send_keys('{END}')
                time.sleep(0.3)
                
                # 위로 스크롤하며 추가 참가자 확인
                extra_scrolls = min(20, max(10, (total_expected - len(participants)) // 15))
                for i in range(extra_scrolls):
                    window.set_focus()
                    send_keys('{PGUP}')
                    time.sleep(0.2)
                    
                    try:
                        extra_elements = window.descendants()
                        initial_count = len(participants)
                        
                        for child in extra_elements:
                            try:
                                text = child.window_text()
                                if text and ("컴퓨터 오디오" in text or "비디오" in text):
                                    name = text.split(',')[0].strip()
                                    if name and not any(x in name.lower() for x in ['검색', '초대', '총 참가자']):
                                        if name not in seen_participants:
                                            participants.append(name)
                                            seen_participants.add(name)
                            except Exception:
                                continue
                        
                        # 새로 찾은 참가자 수 로깅
                        new_found = len(participants) - initial_count
                        if new_found > 0:
                            self.logger.info(f"추가 검사 {i+1}: {new_found}명 추가 발견 (총 {len(participants)}명)")
                    except Exception as e:
                        self.logger.error(f"추가 검사 중 오류: {str(e)}")
                
                # 맨 위로 다시 스크롤 (원래 상태로 복원)
                window.set_focus()
                send_keys('{HOME}')
            
            # 총 소요 시간 계산
            total_time = time.time() - start_time
            
            # 참가자 수 불일치 로깅
            if total_expected > 0 and total_expected != len(participants):
                self.logger.info(f"참고: 예상 참가자 수({total_expected}명)와 찾은 참가자 수({len(participants)}명)가 다릅니다.")
                self.logger.info(f"      {total_expected - len(participants)}명 누락됨 ({(total_expected - len(participants))/total_expected*100:.1f}%)")
            
            self.logger.info(f"[시간 측정] 참가자 추출 완료: 총 {len(participants)}명")
            self.logger.info(f"[시간 측정] 총 소요 시간: {total_time:.3f}초")
            self.logger.info(f"[시간 측정] 스크롤 대기 시간: {total_scroll_time:.3f}초 ({(total_scroll_time/total_time*100):.1f}%)")
            self.logger.info(f"[시간 측정] 추출 시간: {extraction_time:.3f}초 ({(extraction_time/total_time*100):.1f}%)")
            self.logger.info(f"[시간 측정] 평균 참가자 추출 속도: {len(participants)/extraction_time:.1f}명/초")
            
            return participants

        except Exception as e:
            self.logger.error(f"참가자 목록 추출 중 오류 발생: {str(e)}")
            return []
            
        finally:
            # COM 해제
            try:
                pythoncom.CoUninitialize()
            except:
                pass

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