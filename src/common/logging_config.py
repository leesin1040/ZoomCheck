# src/common/logging_config.py
import os
import logging
import datetime

def setup_logging(log_level=logging.INFO, log_to_file=True):
    """로깅 설정
    
    Args:
        log_level: 로깅 레벨
        log_to_file: 파일 로깅 여부
    """
    # 로그 디렉토리 생성
    if log_to_file:
        log_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'logs')
        os.makedirs(log_dir, exist_ok=True)
        
        # 로그 파일 이름 설정
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        log_file = os.path.join(log_dir, f'zoom_check_{timestamp}.log')
    
    # 로깅 포맷 설정
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    
    # 로깅 핸들러 설정
    handlers = [logging.StreamHandler()]
    
    if log_to_file:
        handlers.append(logging.FileHandler(log_file, encoding='utf-8'))
    
    # 로깅 설정
    logging.basicConfig(
        level=log_level,
        format=log_format,
        handlers=handlers
    )

def setup_gui_logging(log_widget, log_level=logging.INFO):
    """GUI용 로깅 설정
    
    Args:
        log_widget: 로그를 표시할 위젯
        log_level: 로깅 레벨
    """
    # GUI 로깅 핸들러 생성
    class QTextEditLogger(logging.Handler):
        def __init__(self, widget):
            super().__init__()
            self.widget = widget
            self.setLevel(log_level)
            self.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
            self.widget.setReadOnly(True)

        def emit(self, record):
            msg = self.format(record)
            self.widget.append(msg)
            # 자동 스크롤
            scrollbar = self.widget.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())
    
    # 기존 핸들러 제거
    logger = logging.getLogger()
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # GUI 로깅 핸들러 추가
    log_handler = QTextEditLogger(log_widget)
    log_handler.setLevel(log_level)
    logger.addHandler(log_handler)
    
    # 콘솔 핸들러도 유지
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    logger.addHandler(console_handler)
    
    # 루트 로거 레벨 설정
    logger.setLevel(log_level)
    
    # pywinauto 로깅도 캡처
    pywinauto_logger = logging.getLogger('pywinauto')
    pywinauto_logger.setLevel(log_level) 