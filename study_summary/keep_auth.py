"""
NotebookLM 인증 세션 유지 스크립트

Playwright 영구 브라우저 프로필을 사용하여 Google 세션 쿠키를
주기적으로 갱신합니다. 작업 스케줄러로 8~12시간마다 실행하면
storage_state.json의 쿠키가 만료되기 전에 자동 갱신됩니다.

동작 원리:
1. Playwright Chromium을 영구 프로필로 열기 (headless)
2. NotebookLM 페이지 방문 → 브라우저 프로필의 쿠키로 자동 인증
3. 인증 성공 시 storage_state.json에 갱신된 쿠키 저장
4. 인증 실패(로그인 페이지 리다이렉트) 시 오류 반환
"""

import sys
import json
import logging
from pathlib import Path
from datetime import datetime

# notebooklm-py 패키지 경로
NOTEBOOKLM_HOME = Path.home() / ".notebooklm"
STORAGE_PATH = NOTEBOOKLM_HOME / "storage_state.json"
BROWSER_PROFILE = NOTEBOOKLM_HOME / "browser_profile"
LOG_DIR = Path(__file__).parent / "logs"

logger = logging.getLogger(__name__)


def setup_logging():
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"keep_auth_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def refresh_session() -> bool:
    """Playwright 영구 프로필로 NotebookLM 방문 후 storage_state 갱신.

    Returns:
        True: 세션 갱신 성공
        False: 인증 만료로 수동 로그인 필요
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        logger.error("Playwright가 설치되어 있지 않습니다: pip install playwright && playwright install chromium")
        return False

    if not BROWSER_PROFILE.exists():
        logger.error(f"브라우저 프로필이 없습니다: {BROWSER_PROFILE}")
        logger.error("먼저 'notebooklm login'을 실행하세요.")
        return False

    logger.info("Playwright 세션 갱신 시작...")
    logger.info(f"  브라우저 프로필: {BROWSER_PROFILE}")
    logger.info(f"  저장 경로: {STORAGE_PATH}")

    with sync_playwright() as p:
        # headless shell은 persistent context를 지원하지 않으므로
        # 전체 Chromium을 --headless=new 모드로 실행
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(BROWSER_PROFILE),
            headless=False,
            args=[
                "--headless=new",
                "--disable-gpu",
                "--disable-blink-features=AutomationControlled",
                "--password-store=basic",
            ],
            ignore_default_args=["--enable-automation"],
        )

        try:
            page = context.pages[0] if context.pages else context.new_page()

            # NotebookLM 방문
            logger.info("  NotebookLM 페이지 로드 중...")
            response = page.goto(
                "https://notebooklm.google.com/",
                wait_until="networkidle",
                timeout=60000,
            )

            final_url = page.url
            logger.info(f"  최종 URL: {final_url}")

            # 로그인 페이지로 리다이렉트되었는지 확인
            if "accounts.google.com" in final_url:
                logger.error("인증 만료: Google 로그인 페이지로 리다이렉트됨")
                logger.error("수동으로 'notebooklm login'을 실행해주세요.")
                return False

            if "notebooklm.google.com" not in final_url:
                logger.warning(f"예상치 못한 URL: {final_url}")
                return False

            # storage_state 저장
            context.storage_state(path=str(STORAGE_PATH))
            logger.info("storage_state.json 갱신 완료!")

            # 검증: 저장된 파일 확인
            data = json.loads(STORAGE_PATH.read_text(encoding="utf-8"))
            cookie_count = len(data.get("cookies", []))
            logger.info(f"  저장된 쿠키 수: {cookie_count}")

            return True

        finally:
            context.close()


def main():
    setup_logging()
    logger.info("=" * 50)
    logger.info("NotebookLM 세션 유지 실행")
    logger.info("=" * 50)

    success = refresh_session()

    if success:
        logger.info("세션 갱신 성공")
    else:
        logger.error("세션 갱신 실패 - 수동 로그인 필요")

    return success


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
