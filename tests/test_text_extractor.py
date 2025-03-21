# tests/test_text_extractor.py
import unittest
import sys
import os
from unittest.mock import patch, MagicMock

# 상위 디렉토리의 모듈을 import할 수 있도록 경로 추가
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src import text_extractor

class TestTextExtractor(unittest.TestCase):
    """텍스트 추출 관련 기능 테스트"""
    
    @patch('src.window_finder.get_process_id_from_window')
    @patch('pywinauto.Application')
    def test_connect_to_window(self, mock_application, mock_get_process_id):
        """애플리케이션 연결 테스트"""
        # 모킹 설정
        mock_get_process_id.return_value = 1234
        mock_app = MagicMock()
        mock_application.return_value.connect.return_value = mock_app
        
        # 테스트 실행
        result = text_extractor.connect_to_window(12345)
        
        # 결과 확인
        self.assertEqual(result, mock_app)
        mock_get_process_id.assert_called_once_with(12345)
        mock_application.return_value.connect.assert_called_once_with(process=1234)
    
    @patch('src.text_extractor.connect_to_window')
    @patch('src.text_extractor.find_participant_list_control')
    @patch('src.text_extractor.extract_participants_from_control')
    @patch('src.window_finder.focus_window')
    def test_extract_participants(self, mock_focus_window, mock_extract_from_control, 
                                 mock_find_control, mock_connect):
        """참가자 추출 테스트"""
        # 모킹 설정
        mock_focus_window.return_value = True
        mock_app = MagicMock()
        mock_connect.return_value = mock_app
        mock_control = MagicMock()
        mock_find_control.return_value = mock_control
        mock_extract_from_control.return_value = ["참가자1", "참가자2", "참가자3"]
        
        # 테스트 실행
        result = text_extractor.extract_participants(12345)
        
        # 결과 확인
        self.assertEqual(result, ["참가자1", "참가자2", "참가자3"])
        mock_focus_window.assert_called_once_with(12345)
        mock_connect.assert_called_once_with(12345)
        mock_find_control.assert_called_once_with(mock_app, 12345)
        mock_extract_from_control.assert_called_once_with(mock_control)

if __name__ == '__main__':
    unittest.main()