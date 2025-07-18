# src/extractors/ui_element_finder.py
import logging
import re
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError
from src.common.constants import PARTICIPANT_CONTROL_PATTERNS, SUPPORTED_CONTROL_TYPES

logger = logging.getLogger(__name__)

class UIElementFinder:
    """UI 요소 검색을 담당하는 클래스"""
    
    @staticmethod
    def find_participant_list_control(app, hwnd):
        """참가자 목록 컨트롤을 찾습니다."""
        try:
            window = app.window(handle=hwnd)
            
            # 모든 하위 요소 탐색
            all_elements = window.descendants()
            
            # 참가자 목록 관련 컨트롤 찾기
            list_elements = []
            for element in all_elements:
                try:
                    control_type = element.control_type()
                    class_name = element.friendly_class_name().lower()
                    
                    # 참가자 목록 관련 컨트롤 타입들
                    if any(pattern in class_name for pattern in PARTICIPANT_CONTROL_PATTERNS):
                        list_elements.append(element)
                    elif control_type in SUPPORTED_CONTROL_TYPES:
                        list_elements.append(element)
                except Exception as e:
                    logger.debug(f"요소 분석 중 오류 (무시됨): {e}")
                    continue
            
            if list_elements:
                logger.info(f"참가자 목록 관련 컨트롤 {len(list_elements)}개 발견")
                return list_elements[0]  # 첫 번째 요소 반환
            
            logger.warning("참가자 목록 컨트롤을 찾을 수 없습니다.")
            return None
            
        except Exception as e:
            logger.error(f"참가자 목록 컨트롤 검색 중 오류: {e}")
            return None
    
    @staticmethod
    def extract_participants_from_control(list_control):
        """컨트롤에서 참가자 목록을 추출합니다."""
        participants = []
        seen_participants = set()
        
        try:
            # 컨트롤의 모든 하위 요소 탐색
            all_elements = list_control.descendants()
            
            for element in all_elements:
                try:
                    # 텍스트가 있는 요소만 처리
                    element_text = element.window_text()
                    if not element_text or element_text.strip() == "":
                        continue
                    
                    # 참가자 이름 정리
                    clean_name = UIElementFinder.clean_participant_name(element_text)
                    
                    if clean_name and clean_name not in seen_participants:
                        participants.append(clean_name)
                        seen_participants.add(clean_name)
                        
                except Exception as e:
                    logger.debug(f"요소 텍스트 추출 중 오류 (무시됨): {e}")
                    continue
            
            logger.info(f"컨트롤에서 {len(participants)}명의 참가자 추출")
            return participants
            
        except Exception as e:
            logger.error(f"컨트롤에서 참가자 추출 중 오류: {e}")
            return []
    
    @staticmethod
    def clean_participant_name(raw_text):
        """참가자 이름만 깔끔하게 추출"""
        try:
            # 콤마로 구분된 첫 번째 부분만 가져오기
            name = raw_text.split(',')[0].strip()
            
            # (호스트), (나) 등의 태그 제거
            name = name.split('(')[0].strip()
            
            # 앞에 붙은 # 제거
            if name.startswith('#'):
                name = name[1:].strip()
            
            # 빈 문자열이나 너무 짧은 이름 제외
            if len(name) < 2:
                return None
                
            return name
        except Exception as e:
            logger.debug(f"이름 정리 중 오류 (무시됨): {e}")
            return None
    
    @staticmethod
    def connect_to_window(hwnd):
        """창에 연결합니다."""
        try:
            app = Application(backend='uia').connect(handle=hwnd)
            return app
        except Exception as e:
            logger.error(f"창 연결 중 오류: {e}")
            return None 