# src/common/logging_config.py
import os
import logging
import datetime

def setup_logging(log_level=logging.INFO, log_to_file=True, log_format=None):
    """로깅 설정
    
    Args:
        log_level: 로깅 레벨
        log_to_file: 파일 로깅 여부
        log_format: 로그 포맷 (None이면 기본값 사용)
    """
    # 로그 디렉토리 생성
    if log_to_file:
        log_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'logs')
        os.makedirs(log_dir, exist_ok=True)
        
        # 로그 파일 이름 설정
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        log_file = os.path.join(log_dir, f'zoom_check_{timestamp}.log')
    
    # 로깅 포맷 설정
    if log_format is None:
        log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    
    # 로깅 핸들러 설정
    handlers = [logging.StreamHandler()]
    
    if log_to_file:
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)  # 파일에는 모든 로그 기록
        handlers.append(file_handler)
    
    # 로깅 설정
    logging.basicConfig(
        level=log_level,
        format=log_format,
        handlers=handlers
    )
    
    # 특정 모듈의 로그 레벨 조정
    logging.getLogger('pywinauto').setLevel(logging.WARNING)
    logging.getLogger('win32gui').setLevel(logging.WARNING)
    
    # 성능 모니터링을 위한 로거
    perf_logger = logging.getLogger('performance')
    perf_logger.setLevel(logging.INFO)
    
    return log_file if log_to_file else None

def log_performance(operation, start_time, end_time, details=None):
    """성능 로깅 헬퍼 함수"""
    duration = end_time - start_time
    perf_logger = logging.getLogger('performance')
    message = f"PERF: {operation} - {duration:.3f}초"
    if details:
        message += f" - {details}"
    perf_logger.info(message)

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