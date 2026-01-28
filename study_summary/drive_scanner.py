"""
Drive 스캔 모듈 - Apps Script 웹앱 연동
- Apps Script가 매일 생성하는 digest HTML을 가져와 회원별 학습 내용 추출
"""
import sys
import os
import json
import re
import logging
import time
from datetime import datetime, timedelta
from typing import Dict, List, Optional
from pathlib import Path

# Windows cp949 인코딩 문제 해결
if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import requests
from bs4 import BeautifulSoup

from config import APPS_SCRIPT_URL, DEADLINE_HOUR

logger = logging.getLogger(__name__)

# members.json 경로
MEMBERS_FILE = Path(__file__).parent / "members.json"


def get_target_date() -> str:
    """
    대상 날짜 계산 (새벽 3시 기준)
    - 현재 시간이 새벽 3시 이전이면 전날 날짜
    - 새벽 3시 이후면 오늘 날짜
    """
    now = datetime.now()
    if now.hour < DEADLINE_HOUR:
        target = now - timedelta(days=1)
    else:
        target = now
    return target.strftime("%Y-%m-%d")


def _load_member_names() -> List[str]:
    """members.json에서 활성 회원 이름 목록 로드"""
    if not MEMBERS_FILE.exists():
        return []
    with open(MEMBERS_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    return [
        m["name"]
        for m in data.get("members", [])
        if m.get("active") and m.get("name")
    ]


def _extract_inner_html(raw_html: str) -> str:
    """
    Apps Script HtmlService의 iframe wrapper에서 실제 콘텐츠 HTML을 추출.
    Google은 doGet 응답을 sandboxFrame iframe 안에 이스케이프된 JS 문자열로 감싼다.
    """
    logger.debug(f"_extract_inner_html: raw_html length={len(raw_html)}")

    # 스크립트 블록에서 가장 긴 문자열 리터럴을 찾음 (= 실제 HTML)
    script_match = re.search(r'<script[^>]*>(.*?)</script>\s*</body>', raw_html, re.DOTALL)
    if not script_match:
        logger.warning(f"_extract_inner_html: no <script> block found in wrapper. raw_html[:200]={raw_html[:200]}")
        return raw_html  # wrapper가 아니면 그대로 반환

    script = script_match.group(1)
    strings = re.findall(r'"([^"]{500,})"', script)
    if not strings:
        logger.warning(f"_extract_inner_html: no long strings found in script block. script[:300]={script[:300]}")
        return raw_html

    longest = max(strings, key=len)
    logger.debug(f"_extract_inner_html: found {len(strings)} strings, longest={len(longest)} chars")

    # JS 이중 이스케이프 디코딩
    # Google의 iframe wrapper는 HTML을 JS 문자열로 2중 이스케이프함
    decoded = longest
    # 1차: \\x → \x, \\" → \", \\\\ → \\, \\n → \n, \\/ → \/
    decoded = re.sub(r'\\\\x([0-9a-fA-F]{2})', lambda m: chr(int(m.group(1), 16)), decoded)
    decoded = decoded.replace('\\\\"', '"')
    decoded = decoded.replace('\\\\/', '/')
    decoded = decoded.replace('\\\\n', '\n')
    decoded = decoded.replace('\\\\t', '\t')
    decoded = decoded.replace('\\\\\\\\', '\\')
    # 2차: 남은 단일 이스케이프
    decoded = decoded.replace('\\n', '\n')
    decoded = decoded.replace('\\t', '\t')
    decoded = decoded.replace('\\/', '/')
    decoded = decoded.replace('\\"', '"')
    decoded = re.sub(r'\\x([0-9a-fA-F]{2})', lambda m: chr(int(m.group(1), 16)), decoded)
    decoded = decoded.replace('\\\\', '\\')

    # JSON 설정 부분을 건너뛰고 실제 HTML 시작점 찾기
    html_start = decoded.find('<!DOCTYPE')
    if html_start < 0:
        html_start = decoded.find('<html')
    if html_start < 0:
        html_start = decoded.find('<')
    if html_start < 0:
        logger.warning("_extract_inner_html: no HTML start tag found in decoded string")
        return raw_html

    result = decoded[html_start:]
    has_member = '.member-section' in result or 'class="member-section"' in result
    logger.debug(f"_extract_inner_html: decoded length={len(result)}, has_member_section={has_member}")
    return result


def fetch_digest_html(date: str, max_retries: int = 3) -> str:
    """
    Apps Script 웹앱에서 digest HTML 가져오기 (재시도 포함)

    Args:
        date: YYYY-MM-DD 형식 날짜
        max_retries: 최대 재시도 횟수 (파싱 실패 시)

    Returns:
        실제 콘텐츠 HTML 문자열 (iframe wrapper 제거됨)

    Raises:
        RuntimeError: 요청 실패 시
    """
    if not APPS_SCRIPT_URL:
        raise RuntimeError("APPS_SCRIPT_URL이 설정되지 않았습니다. .env 파일을 확인하세요.")

    url = f"{APPS_SCRIPT_URL}?date={date}"

    for attempt in range(1, max_retries + 1):
        try:
            logger.info(f"  Apps Script 요청 (시도 {attempt}/{max_retries}): {date}")
            resp = requests.get(url, timeout=60)
            resp.raise_for_status()

            raw_text = resp.text
            raw_len = len(raw_text)
            html = _extract_inner_html(raw_text)
            extracted_len = len(html)
            logger.info(f"  응답: raw={raw_len} chars → extracted={extracted_len} chars")

            # 추출 성공 여부 판단: 추출 후 크기가 변했거나, raw 자체가 콘텐츠인 경우
            extraction_ok = (html is not raw_text) or raw_len < 100000
            if extraction_ok:
                return html

            # 추출 실패: raw HTML이 그대로 반환됨 (iframe wrapper 추출 실패)
            logger.warning(f"  iframe wrapper 추출 실패 (시도 {attempt}/{max_retries}), raw={raw_len} chars")
            if attempt < max_retries:
                wait = attempt * 5
                logger.info(f"  {wait}초 후 재시도...")
                time.sleep(wait)
        except requests.RequestException as e:
            logger.warning(f"  요청 실패 (시도 {attempt}/{max_retries}): {e}")
            if attempt < max_retries:
                wait = attempt * 5
                logger.info(f"  {wait}초 후 재시도...")
                time.sleep(wait)
            else:
                raise RuntimeError(f"Apps Script 요청 {max_retries}회 실패: {e}")

    # 모든 재시도 후에도 추출 실패하면 마지막 결과 반환
    logger.warning("  모든 재시도 완료. 마지막 추출 결과 반환")
    return html


def parse_digest_html(html: str) -> List[Dict]:
    """
    digest HTML을 파싱하여 회원별 데이터 추출

    Args:
        html: Apps Script가 생성한 digest HTML

    Returns:
        [{"name": str, "text_content": str, "files": [str, ...]}, ...]
    """
    soup = BeautifulSoup(html, "html.parser")
    members = []

    for section in soup.select(".member-section"):
        # 회원 이름
        h2 = section.find("h2")
        if not h2:
            continue
        name = h2.get_text(strip=True)

        # 파일 목록
        files = []
        for li in section.select(".file-list li"):
            files.append(li.get_text(strip=True))

        # 학습 내용 텍스트
        content_body = section.select_one(".content-body")
        text_content = content_body.get_text(separator="\n", strip=True) if content_body else ""

        members.append({
            "name": name,
            "text_content": text_content,
            "files": files,
        })

    return members


def scan_all_members(target_date: Optional[str] = None) -> List[Dict]:
    """
    모든 회원의 학습 데이터 수집 (Apps Script 웹앱 경유)

    Returns:
        기존과 동일한 형식:
        [{"name", "date", "has_submission", "text_content", "files"}, ...]
    """
    if target_date is None:
        target_date = get_target_date()

    print(f"📅 대상 날짜: {target_date}")

    # 1) 웹앱에서 HTML 가져오기
    print("  🌐 Apps Script 웹앱에서 데이터 가져오는 중...")
    html = fetch_digest_html(target_date)
    logger.info(f"  fetch_digest_html 결과: {len(html)} chars, has_member_section={'member-section' in html}")

    # 2) HTML 파싱 → 제출한 회원 데이터
    parsed = parse_digest_html(html)
    submitted_names = {m["name"] for m in parsed}
    print(f"  📊 HTML에서 {len(parsed)}명 데이터 파싱 완료")
    if len(parsed) == 0:
        logger.warning(f"  파싱 결과 0명! HTML 앞부분: {html[:500]}")

    # 3) 결과 조립
    results = []
    for m in parsed:
        results.append({
            "name": m["name"],
            "date": target_date,
            "has_submission": True,
            "text_content": m["text_content"],
            "files": [{"name": f, "type": "unknown"} for f in m["files"]],
        })

    # 4) members.json에서 미제출 회원 추가
    all_names = _load_member_names()
    for name in all_names:
        if name not in submitted_names:
            results.append({
                "name": name,
                "date": target_date,
                "has_submission": False,
                "text_content": "",
                "files": [],
            })

    # 5) 요약 출력
    submitted_count = sum(1 for r in results if r["has_submission"])
    print(f"  👥 전체: {len(results)}명 (제출 {submitted_count} / 미제출 {len(results) - submitted_count})")
    for r in results:
        status = "✅" if r["has_submission"] else "❌"
        print(f"    {status} {r['name']}")

    return results


def test_connection() -> bool:
    """Apps Script 웹앱 연결 테스트"""
    try:
        if not APPS_SCRIPT_URL:
            print("❌ APPS_SCRIPT_URL이 설정되지 않았습니다.")
            return False

        print(f"  URL: {APPS_SCRIPT_URL}")
        target_date = get_target_date()
        html = fetch_digest_html(target_date)

        parsed = parse_digest_html(html)
        print(f"✅ Apps Script 웹앱 연결 성공! ({len(parsed)}명 데이터 수신)")
        return True
    except Exception as e:
        print(f"❌ Apps Script 웹앱 연결 실패: {e}")
        return False


if __name__ == "__main__":
    print("=== Apps Script 웹앱 연결 테스트 ===")
    if test_connection():
        print("\n=== 회원 데이터 스캔 테스트 ===")
        results = scan_all_members()

        print("\n=== 스캔 결과 ===")
        for r in results:
            status = "✅" if r["has_submission"] else "❌"
            files_count = len(r["files"])
            content_len = len(r["text_content"])
            print(f"{status} {r['name']}: {files_count}개 파일, {content_len}자 텍스트")
