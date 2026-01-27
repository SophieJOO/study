"""
Drive 스캔 모듈 - Apps Script 웹앱 연동
- Apps Script가 매일 생성하는 digest HTML을 가져와 회원별 학습 내용 추출
"""
import json
import re
from datetime import datetime, timedelta
from typing import Dict, List, Optional
from pathlib import Path

import requests
from bs4 import BeautifulSoup

from config import APPS_SCRIPT_URL, DEADLINE_HOUR

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
    # 스크립트 블록에서 가장 긴 문자열 리터럴을 찾음 (= 실제 HTML)
    script_match = re.search(r'<script[^>]*>(.*?)</script>\s*</body>', raw_html, re.DOTALL)
    if not script_match:
        return raw_html  # wrapper가 아니면 그대로 반환

    script = script_match.group(1)
    strings = re.findall(r'"([^"]{500,})"', script)
    if not strings:
        return raw_html

    longest = max(strings, key=len)

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
        return raw_html

    return decoded[html_start:]


def fetch_digest_html(date: str) -> str:
    """
    Apps Script 웹앱에서 digest HTML 가져오기

    Args:
        date: YYYY-MM-DD 형식 날짜

    Returns:
        실제 콘텐츠 HTML 문자열 (iframe wrapper 제거됨)

    Raises:
        RuntimeError: 요청 실패 시
    """
    if not APPS_SCRIPT_URL:
        raise RuntimeError("APPS_SCRIPT_URL이 설정되지 않았습니다. .env 파일을 확인하세요.")

    url = f"{APPS_SCRIPT_URL}?date={date}"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    return _extract_inner_html(resp.text)


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

    # 2) HTML 파싱 → 제출한 회원 데이터
    parsed = parse_digest_html(html)
    submitted_names = {m["name"] for m in parsed}
    print(f"  📊 HTML에서 {len(parsed)}명 데이터 파싱 완료")

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
