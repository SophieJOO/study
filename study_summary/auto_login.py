"""
NotebookLM 자동 로그인 스크립트

브라우저 프로필에 Google 로그인이 유지되어 있으면
자동으로 NotebookLM 페이지 로드 후 Enter를 눌러 인증 저장.
"""
import subprocess
import time
import sys
from pathlib import Path

NOTEBOOKLM_CLI = Path.home() / "AppData/Roaming/Python/Python314/Scripts/notebooklm.exe"
WAIT_SECONDS = 30  # 브라우저 로딩 대기 시간 (NotebookLM 홈페이지 완전 로드까지)


def auto_login():
    """NotebookLM 로그인 자동화"""
    print(f"NotebookLM 자동 로그인 시작...")
    print(f"CLI 경로: {NOTEBOOKLM_CLI}")

    if not NOTEBOOKLM_CLI.exists():
        print(f"오류: {NOTEBOOKLM_CLI} 파일을 찾을 수 없습니다.")
        return False

    # 로그인 프로세스 시작
    proc = subprocess.Popen(
        [str(NOTEBOOKLM_CLI), "login"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )

    print(f"브라우저 열림, {WAIT_SECONDS}초 대기 중...")
    time.sleep(WAIT_SECONDS)

    # Enter 전송
    print("Enter 전송...")
    proc.stdin.write("\n")
    proc.stdin.flush()

    # 결과 대기
    stdout, _ = proc.communicate(timeout=30)
    print(stdout)

    if proc.returncode == 0:
        print("로그인 성공!")
        return True
    else:
        print(f"로그인 실패 (exit code: {proc.returncode})")
        return False


if __name__ == "__main__":
    success = auto_login()
    sys.exit(0 if success else 1)
