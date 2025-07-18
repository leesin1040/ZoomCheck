# src/common/constants.py
# Zoom 관련 상수들

# 창 검색 관련
ZOOM_PARTICIPANTS_WINDOW_TITLE = "참가자"
ZOOM_WINDOW_TITLE_CONTAINS = ["참가자", "Participants"]

# 스크롤 관련
DEFAULT_SCROLL_DELAY = 0.8
MAX_SCROLL_ATTEMPTS = 60
SCROLL_PAGE_DOWN_COUNT = 2

# UI 요소 관련
PARTICIPANT_CONTROL_PATTERNS = ['list', 'item', 'participant', 'user']
SUPPORTED_CONTROL_TYPES = ['ListItem', 'List', 'Custom']

# 파일 관련
DEFAULT_OUTPUT_DIR = "output"
DEFAULT_LOG_DIR = "logs"
PARTICIPANTS_FILE_PREFIX = "zoom_participants"

# 로깅 관련
DEFAULT_LOG_LEVEL = "INFO"
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
GUI_LOG_FORMAT = '%(levelname)s: %(message)s' 