"""
스터디 요약 자동화 - 인포그래픽 생성 모듈 (notebooklm-py)

NotebookLM API를 사용하여 회원별 학습 인포그래픽을 생성합니다.
"""
import asyncio
import logging
import time
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from config import OUTPUT_DIR

logger = logging.getLogger(__name__)

MAX_RETRIES = 3
RETRY_DELAY = 5  # seconds
GENERATION_TIMEOUT = 180.0  # seconds


def _overlay_label(image_path: Path, member_name: str, date: str):
    """인포그래픽 이미지 상단에 이름과 날짜 라벨을 오버레이한다."""
    img = Image.open(image_path)
    draw = ImageDraw.Draw(img)

    label = f"{member_name} | {date}"

    # 폰트: 이미지 너비의 ~2.5% 크기
    font_size = max(24, img.width // 40)
    try:
        font = ImageFont.truetype("malgun.ttf", font_size)
    except OSError:
        try:
            font = ImageFont.truetype("C:/Windows/Fonts/malgun.ttf", font_size)
        except OSError:
            font = ImageFont.load_default()

    # 텍스트 크기 계산
    bbox = draw.textbbox((0, 0), label, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]

    # 배경 박스 (상단 중앙)
    padding = 12
    box_x = (img.width - text_w) // 2 - padding
    box_y = padding
    draw.rounded_rectangle(
        [box_x, box_y, box_x + text_w + padding * 2, box_y + text_h + padding * 2],
        radius=8,
        fill=(0, 0, 0, 180),
    )

    # 텍스트
    draw.text(
        (box_x + padding, box_y + padding),
        label,
        font=font,
        fill=(255, 255, 255),
    )

    img.save(image_path)
    logger.info(f"  라벨 오버레이 완료: {label}")


async def _generate_infographic_async(
    member_name: str,
    study_content: str,
    date: str,
    output_dir: Path = OUTPUT_DIR,
) -> Path | None:
    """
    NotebookLM을 사용하여 인포그래픽을 비동기 생성합니다.

    Args:
        member_name: 회원 이름
        study_content: 학습 내용 텍스트
        date: 대상 날짜 (YYYY-MM-DD)
        output_dir: 출력 디렉토리

    Returns:
        생성된 이미지 파일 경로, 실패 시 None
    """
    from notebooklm import NotebookLMClient, InfographicOrientation, InfographicDetail

    output_dir.mkdir(parents=True, exist_ok=True)
    notebook = None

    try:
        async with await NotebookLMClient.from_storage() as client:
            # 1. 임시 노트북 생성
            notebook_title = f"{member_name}_{date}"
            notebook = await client.notebooks.create(notebook_title)
            logger.info(f"  노트북 생성: {notebook_title}")

            # 2. 소스 추가 (학습 내용 텍스트 + 이름/날짜 헤더)
            header = f"[{member_name}] {date} 학습 인증\n\n"
            await client.sources.add_text(
                notebook.id,
                title=f"{date} 학습 인증 - {member_name}",
                content=header + study_content,
                wait=True,
            )
            logger.info(f"  소스 추가 완료")

            # 3. 인포그래픽 생성 요청
            status = await client.artifacts.generate_infographic(
                notebook.id,
                language="ko",
                orientation=InfographicOrientation.PORTRAIT,
                detail_level=InfographicDetail.DETAILED,
            )
            logger.info(f"  인포그래픽 생성 요청 완료, 대기 중...")

            # 4. 완료 대기
            await client.artifacts.wait_for_completion(
                notebook.id,
                status.task_id,
                timeout=GENERATION_TIMEOUT,
            )
            logger.info(f"  인포그래픽 생성 완료")

            # 5. 다운로드
            output_path = output_dir / f"infographic_{member_name}_{date}.png"
            await client.artifacts.download_infographic(
                notebook.id, str(output_path)
            )
            logger.info(f"  다운로드 완료: {output_path}")

            # 5-1. 이름/날짜 오버레이
            _overlay_label(output_path, member_name, date)

            # 6. 노트북 삭제 (정리)
            await client.notebooks.delete(notebook.id)
            notebook = None
            logger.info(f"  노트북 삭제 완료")

            return output_path

    except Exception as e:
        logger.error(f"  인포그래픽 생성 오류: {e}")
        # 노트북이 생성된 경우 정리 시도
        if notebook is not None:
            try:
                async with await NotebookLMClient.from_storage() as cleanup_client:
                    await cleanup_client.notebooks.delete(notebook.id)
                    logger.info(f"  오류 후 노트북 정리 완료")
            except Exception:
                logger.warning(f"  노트북 정리 실패 (수동 삭제 필요)")
        return None


def generate_infographic(
    member_name: str,
    study_content: str,
    date: str,
    output_dir: Path = OUTPUT_DIR,
) -> Path | None:
    """
    인포그래픽 생성 (동기 래퍼, 최대 3회 재시도).

    Args:
        member_name: 회원 이름
        study_content: 학습 내용 텍스트
        date: 대상 날짜 (YYYY-MM-DD)
        output_dir: 출력 디렉토리

    Returns:
        생성된 이미지 파일 경로, 실패 시 None
    """
    for attempt in range(1, MAX_RETRIES + 1):
        logger.info(f"  [{member_name}] 시도 {attempt}/{MAX_RETRIES}")
        result = asyncio.run(
            _generate_infographic_async(member_name, study_content, date, output_dir)
        )
        if result is not None:
            return result

        if attempt < MAX_RETRIES:
            logger.warning(f"  [{member_name}] 재시도 대기 {RETRY_DELAY}초...")
            time.sleep(RETRY_DELAY)

    logger.error(f"  [{member_name}] {MAX_RETRIES}회 시도 모두 실패")
    return None


# main.py 호환 별칭
generate_educational_infographic = generate_infographic


def test_single_infographic() -> Path | None:
    """
    테스트용 인포그래픽 생성.
    간단한 샘플 데이터로 NotebookLM 연결 및 생성을 검증합니다.

    Returns:
        생성된 테스트 이미지 경로, 실패 시 None
    """
    test_content = """
# Python 리스트 컴프리헨션
## 기본 문법
- [표현식 for 항목 in 반복가능객체]
- 예: squares = [x**2 for x in range(10)]

## 조건부 필터링
- [x for x in range(20) if x % 2 == 0]
- 짝수만 필터링

## 장점
1. 간결한 코드
2. 일반 for문보다 빠름
3. 가독성 향상
"""
    from datetime import datetime

    test_date = datetime.now().strftime("%Y-%m-%d")
    print("NotebookLM 인포그래픽 생성 테스트...")

    result = generate_infographic(
        member_name="테스트",
        study_content=test_content,
        date=test_date,
    )

    if result:
        print(f"  테스트 성공: {result}")
    else:
        print("  테스트 실패: 인포그래픽 생성에 실패했습니다.")

    return result


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    test_single_infographic()
