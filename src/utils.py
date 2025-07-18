# src/utils.py
import os
import logging
import datetime
from src.common.constants import DEFAULT_OUTPUT_DIR, PARTICIPANTS_FILE_PREFIX

def setup_logging(log_level=logging.INFO):
    """로깅 설정 (기존 호환성을 위해 유지)
    
    Args:
        log_level: 로깅 레벨
    """
    from src.common.logging_config import setup_logging as setup_logging_common
    setup_logging_common(log_level)

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
        filename = f'{PARTICIPANTS_FILE_PREFIX}_{timestamp}.txt'
    
    # 출력 디렉토리 생성
    output_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), DEFAULT_OUTPUT_DIR)
    os.makedirs(output_dir, exist_ok=True)
    
    # 파일 경로 설정
    file_path = os.path.join(output_dir, filename)
    
    # 파일 저장
    with open(file_path, 'w', encoding='utf-8') as f:
        for participant in participants:
            f.write(f"{participant}\n")
    
    return file_path