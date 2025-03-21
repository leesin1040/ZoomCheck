# src/utils.py
import os
import logging
import datetime

def setup_logging(log_level=logging.INFO):
    """로깅 설정
    
    Args:
        log_level: 로깅 레벨
    """
    # 로그 디렉토리 생성
    log_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'logs')
    os.makedirs(log_dir, exist_ok=True)
    
    # 로그 파일 이름 설정
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f'zoom_check_{timestamp}.log')
    
    # 로깅 포맷 설정
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    
    # 로깅 핸들러 설정
    handlers = [
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
    
    # 로깅 설정
    logging.basicConfig(
        level=log_level,
        format=log_format,
        handlers=handlers
    )

def save_participants_to_file(participants, filename=None):
    """참가자 목록을 파일로 저장
    
    Args:
        participants (list): 참가자 목록
        filename (str, optional): 저장할 파일 이름
    
    Returns:
        str: 저장된 파일 경로
    """
    # 기본 파일 이름 설정
    if not filename:
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'zoom_participants_{timestamp}.txt'
    
    # 출력 디렉토리 생성
    output_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'output')
    os.makedirs(output_dir, exist_ok=True)
    
    # 파일 경로 설정
    file_path = os.path.join(output_dir, filename)
    
    # 파일 저장
    with open(file_path, 'w', encoding='utf-8') as f:
        for participant in participants:
            f.write(f"{participant}\n")
    
    return file_path