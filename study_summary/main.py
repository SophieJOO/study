"""
스터디 요약 자동화 - 메인 실행 스크립트

각 회원의 공부 내용을:
1. Google Drive에서 수집
2. 교육적인 인포그래픽 이미지로 변환 (회원별 개별 생성)
3. Slack DM으로 전송
"""
import sys
import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict

# 설정 모듈
from config import LOG_DIR, OUTPUT_DIR, DEADLINE_HOUR


def setup_logging():
    """로깅 설정"""
    log_file = LOG_DIR / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    return logging.getLogger(__name__)


def get_target_date() -> str:
    """대상 날짜 계산 (새벽 3시 기준)"""
    now = datetime.now()
    if now.hour < DEADLINE_HOUR:
        target = now - timedelta(days=1)
    else:
        target = now
    return target.strftime("%Y-%m-%d")


def run_pipeline(test_mode: bool = False):
    """
    전체 파이프라인 실행 - 각 회원별 개별 인포그래픽 생성
    
    Args:
        test_mode: True면 테스트 데이터 사용
    """
    logger = setup_logging()
    
    logger.info("=" * 50)
    logger.info("🚀 스터디 인포그래픽 자동 생성 시작")
    logger.info("=" * 50)
    
    target_date = get_target_date()
    logger.info(f"📅 대상 날짜: {target_date}")
    
    try:
        # 1단계: Google Drive 스캔 또는 테스트 데이터
        logger.info("\n📂 1단계: 공부 내용 수집")
        
        if test_mode:
            scan_results = [
                {
                    "name": "홍길동", 
                    "date": target_date, 
                    "has_submission": True,
                    "text_content": """
# JavaScript 화살표 함수
## 기본 문법
- 기존: function(x) { return x * 2; }
- 화살표: (x) => x * 2

## 특징
1. 간결한 문법 - 코드가 짧아짐
2. this 바인딩이 렉시컬 방식
3. 콜백 함수에 특히 유용

## 예시
const doubled = [1,2,3].map(n => n * 2);
"""
                },
                {
                    "name": "김철수",
                    "date": target_date,
                    "has_submission": True,
                    "text_content": """
# React useState 훅
## 상태 관리의 기본
- 함수형 컴포넌트에서 상태 사용
- const [state, setState] = useState(초기값)

## 특징
1. 불변성 유지 필요
2. 비동기로 업데이트됨
3. 이전 상태 기반 업데이트: setState(prev => prev + 1)

## 예시
const [count, setCount] = useState(0);
"""
                },
                {
                    "name": "박민수",
                    "date": target_date,
                    "has_submission": False,
                    "text_content": ""
                },
            ]
            logger.info(f"  테스트 모드: {len(scan_results)}명 데이터")
        else:
            from drive_scanner import scan_all_members
            scan_results = scan_all_members(target_date)
        
        # 제출한 회원 필터링
        submitted = [r for r in scan_results if r.get("has_submission") and r.get("text_content")]
        logger.info(f"  제출 완료: {len(submitted)}/{len(scan_results)}명")
        
        if not submitted:
            logger.warning("❌ 인포그래픽을 생성할 회원이 없습니다.")
            return True  # 에러는 아님
        
        # 2단계: 각 회원별 인포그래픽 생성
        logger.info("\n🎨 2단계: 개별 인포그래픽 생성")
        
        from infographic_generator import generate_educational_infographic
        
        generated_images = []
        for member in submitted:
            logger.info(f"\n  📝 {member['name']} 인포그래픽 생성 중...")
            
            try:
                image_path = generate_educational_infographic(
                    member_name=member["name"],
                    study_content=member["text_content"],
                    date=member["date"]
                )
                
                if image_path:
                    generated_images.append({
                        "name": member["name"],
                        "path": image_path
                    })
                    logger.info(f"  ✅ {member['name']}: {image_path}")
                else:
                    logger.warning(f"  ⚠️ {member['name']}: 이미지 생성 실패")
            except Exception as e:
                logger.error(f"  ❌ {member['name']}: 오류 - {e}")
        
        logger.info(f"\n📊 생성 결과: {len(generated_images)}/{len(submitted)}개 성공")
        
        # 3단계: Slack DM 전송
        if generated_images and not test_mode:
            logger.info("\n📤 3단계: Slack DM 전송")
            
            from slack_sender import send_dm_with_image
            
            for img in generated_images:
                message = f"📚 {target_date} {img['name']}님의 학습 인포그래픽\nSlack에 공유해주세요! 💪"
                success = send_dm_with_image(img["path"], message)
                if success:
                    logger.info(f"  ✅ {img['name']} 이미지 전송 완료")
                else:
                    logger.warning(f"  ⚠️ {img['name']} 전송 실패")
        elif test_mode:
            logger.info("\n📤 3단계: Slack 전송 (테스트 모드 - 스킵)")
            for img in generated_images:
                logger.info(f"  📁 {img['name']}: {img['path']}")
        
        logger.info("\n" + "=" * 50)
        logger.info("✅ 파이프라인 완료!")
        logger.info("=" * 50)
        
        return True
        
    except Exception as e:
        logger.error(f"❌ 파이프라인 오류: {e}", exc_info=True)
        return False


def run_tests():
    """연결 테스트 실행"""
    print("=" * 50)
    print("🧪 연결 테스트")
    print("=" * 50)
    
    from drive_scanner import test_connection as test_drive
    from slack_sender import test_connection as test_slack
    
    print("\n1️⃣ Apps Script 웹앱 연결 테스트")
    drive_ok = test_drive()
    
    print("\n2️⃣ NotebookLM 인포그래픽 생성 테스트")
    try:
        from infographic_generator import test_single_infographic
        image_path = test_single_infographic()
        gemini_ok = image_path is not None
    except Exception as e:
        print(f"❌ NotebookLM 인포그래픽 생성 실패: {e}")
        gemini_ok = False
    
    print("\n3️⃣ Slack 연결 테스트")
    slack_ok = test_slack()
    
    print("\n" + "=" * 50)
    print("테스트 결과:")
    print(f"  Apps Script:  {'✅' if drive_ok else '❌'}")
    print(f"  NotebookLM:   {'✅' if gemini_ok else '❌'}")
    print(f"  Slack:        {'✅' if slack_ok else '❌'}")
    print("=" * 50)
    
    return all([drive_ok, gemini_ok, slack_ok])


def main():
    """메인 함수"""
    import argparse
    
    parser = argparse.ArgumentParser(description="스터디 인포그래픽 자동 생성")
    parser.add_argument("--test", action="store_true", help="테스트 데이터로 실행")
    parser.add_argument("--check", action="store_true", help="연결 테스트만 실행")
    
    args = parser.parse_args()
    
    if args.check:
        success = run_tests()
        sys.exit(0 if success else 1)
    else:
        success = run_pipeline(test_mode=args.test)
        sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
