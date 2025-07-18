# tests/test_utils.py
import unittest
import sys
import os
import tempfile
import shutil
import logging
from unittest.mock import patch, MagicMock

# 상위 디렉토리의 모듈을 import할 수 있도록 경로 추가
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src import utils

class TestUtils(unittest.TestCase):
    """유틸리티 함수 테스트"""
    
    def setUp(self):
        """테스트 설정"""
        self.temp_dir = tempfile.mkdtemp()
    
    def tearDown(self):
        """테스트 정리"""
        shutil.rmtree(self.temp_dir)
    
    @patch('src.common.logging_config.setup_logging')
    def test_setup_logging(self, mock_setup):
        """로깅 설정 테스트"""
        # 테스트 실행
        utils.setup_logging(logging.INFO)
        
        # 결과 확인
        mock_setup.assert_called_once_with(logging.INFO)
    
    @patch('os.makedirs')
    @patch('builtins.open', create=True)
    def test_save_participants_to_file(self, mock_open, mock_makedirs):
        """참가자 목록 파일 저장 테스트"""
        # 테스트 데이터
        participants = ["참가자1", "참가자2", "참가자3"]
        
        # 모킹 설정
        mock_file = MagicMock()
        mock_open.return_value.__enter__.return_value = mock_file
        
        # 테스트 실행
        result = utils.save_participants_to_file(participants)
        
        # 결과 확인
        self.assertIn("zoom_participants", result)
        self.assertIn(".txt", result)
        mock_makedirs.assert_called_once()
        mock_file.write.assert_called()
    
    @patch('os.makedirs')
    @patch('builtins.open', create=True)
    def test_save_participants_to_file_with_custom_filename(self, mock_open, mock_makedirs):
        """사용자 정의 파일명으로 참가자 목록 저장 테스트"""
        # 테스트 데이터
        participants = ["참가자1", "참가자2"]
        custom_filename = "custom_participants.txt"
        
        # 모킹 설정
        mock_file = MagicMock()
        mock_open.return_value.__enter__.return_value = mock_file
        
        # 테스트 실행
        result = utils.save_participants_to_file(participants, custom_filename)
        
        # 결과 확인
        self.assertIn(custom_filename, result)
        mock_file.write.assert_called()

if __name__ == '__main__':
    unittest.main() 