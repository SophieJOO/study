"""
스터디 요약 자동화 - 설정 모듈
"""
import os
import json
from pathlib import Path
from dotenv import load_dotenv

# .env 파일 로드
load_dotenv()

# 프로젝트 경로
PROJECT_DIR = Path(__file__).parent
OUTPUT_DIR = PROJECT_DIR / os.getenv("OUTPUT_DIR", "output")
LOG_DIR = PROJECT_DIR / os.getenv("LOG_DIR", "logs")

# 폴더 생성
OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

# Apps Script 웹앱 URL (Drive 스캔 대체)
APPS_SCRIPT_URL = os.getenv("APPS_SCRIPT_URL", "")

# Gemini API 설정
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# Slack 설정
SLACK_BOT_TOKEN = os.getenv("SLACK_BOT_TOKEN", "")
SLACK_USER_ID = os.getenv("SLACK_USER_ID", "")

# =========================================
# 회원 목록 및 폴더 ID 매핑 (members.json에서 로드)
# =========================================
MEMBERS_FILE = PROJECT_DIR / "members.json"
MEMBERS = {}
if MEMBERS_FILE.exists():
    with open(MEMBERS_FILE, "r", encoding="utf-8") as _f:
        _members_data = json.load(_f)
    MEMBERS = {
        m["name"]: m["folder_id"]
        for m in _members_data.get("members", [])
        if m.get("active") and m.get("name") and m.get("folder_id")
    }

# =========================================
# 이미지 생성 설정
# =========================================
IMAGE_STYLE_PROMPT = """
인포그래픽 스타일의 이미지를 생성해주세요.
- 깔끔하고 현대적인 디자인
- 부드러운 파스텔 톤 배경
- 아이콘과 시각적 요소 사용
- 각 회원별로 구분된 섹션
- 한국어 텍스트
- 고해상도 (1920x1080)
"""

# 마감 시간 (새벽 5시 이전 실행 시 전날 날짜 사용)
DEADLINE_HOUR = 5
