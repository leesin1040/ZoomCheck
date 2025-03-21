# src/main.py
import sys
import time
import logging
import argparse
from . import window_finder
from . import text_extractor
from . import utils

def parse_arguments():
    """명령줄 인자 파싱
    
    Returns:
        argparse.Namespace: 파싱된 인자
    """
    parser = argparse.ArgumentParser(description='줌 참가자 목록 추출 프로그램')
    parser.add_argument('--debug', action='store_true', help='디버그 모드 활성화')
    parser.add_argument('--interval', type=int, default=0, help='주기적 추출 간격(초), 0이면 1회만 실행')
    parser.add_argument('--output', type=str, help='출력 파일 이름')
    return parser.parse_args()

def main():
    """메인 함수"""
    # 인자 파싱
    args = parse_arguments()
    
    # 로깅 설정
    log_level = logging.DEBUG if args.debug else logging.INFO
    utils.setup_logging(log_level)
    
    logger = logging.getLogger(__name__)
    logger.info("줌 참가자 목록 추출 프로그램 시작")
    
    try:
        # 주기적 실행 여부 확인
        if args.interval > 0:
            logger.info(f"{args.interval}초 간격으로 주기적 실행 모드")
            
            while True:
                run_extraction(args.output)
                time.sleep(args.interval)
        else:
            # 1회 실행
            run_extraction(args.output)
    
    except KeyboardInterrupt:
        logger.info("사용자에 의해 프로그램 종료됨")
    except Exception as e:
        logger.error(f"프로그램 실행 중 오류 발생: {e}", exc_info=True)
        return 1
    
    return 0

def run_extraction(output_filename=None):
    """참가자 목록 추출 실행
    
    Args:
        output_filename (str, optional): 출력 파일 이름
    """
    logger = logging.getLogger(__name__)
    
    # 줌 참가자 창 찾기
    hwnd = window_finder.find_zoom_participants_window()
    
    if not hwnd:
        logger.error("줌 참가자 창을 찾을 수 없습니다. 줌이 실행 중이고 참가자 목록이 열려 있는지 확인하세요.")
        return
    
    # 창 정보 출력
    window_title = window_finder.get_window_text(hwnd)
    logger.info(f"창 제목: {window_title}")
    
    # 참가자 목록 추출
    participants = text_extractor.extract_participants(hwnd)
    
    if participants:
        logger.info(f"{len(participants)}명의 참가자를 찾았습니다:")
        for i, participant in enumerate(participants, 1):
            logger.info(f"{i}. {participant}")
        
        # 파일로 저장
        file_path = utils.save_participants_to_file(participants, output_filename)
        logger.info(f"참가자 목록이 {file_path}에 저장되었습니다.")
    else:
        logger.warning("참가자를 찾을 수 없습니다.")

if __name__ == "__main__":
    sys.exit(main())