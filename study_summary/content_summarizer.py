"""
콘텐츠 요약 모듈
- Gemini API를 사용하여 공부 내용 요약
"""
from typing import Dict, List
import google.generativeai as genai

from config import GEMINI_API_KEY


def init_gemini():
    """Gemini API 초기화"""
    if not GEMINI_API_KEY:
        raise ValueError("GEMINI_API_KEY가 설정되지 않았습니다. .env 파일을 확인하세요.")
    
    genai.configure(api_key=GEMINI_API_KEY)
    return genai.GenerativeModel('gemini-2.0-flash')


def summarize_content(text_content: str, max_length: int = 50) -> str:
    """
    공부 내용을 짧게 요약
    
    Args:
        text_content: 원본 텍스트
        max_length: 요약 최대 길이 (글자 수)
    
    Returns:
        요약된 내용 (짧은 문장 또는 키워드)
    """
    if not text_content.strip():
        return "내용 없음"
    
    model = init_gemini()
    
    prompt = f"""다음 공부 인증 내용을 {max_length}자 이내로 아주 간결하게 요약해주세요.
핵심 키워드나 주제만 추출하세요. 존댓말 없이 명사형으로 끝내세요.

예시 출력:
- "JavaScript 화살표 함수 학습"
- "알고리즘 문제 5개 풀이"
- "React useState 훅 정리"

공부 내용:
{text_content[:2000]}  # 너무 길면 자름

요약:"""

    try:
        response = model.generate_content(prompt)
        summary = response.text.strip()
        
        # 줄바꿈이 있으면 첫 줄만
        if '\n' in summary:
            summary = summary.split('\n')[0]
        
        # 너무 길면 자르기
        if len(summary) > max_length + 10:
            summary = summary[:max_length] + "..."
        
        return summary
    except Exception as e:
        print(f"요약 실패: {e}")
        return "요약 생성 실패"


def summarize_from_files(files: List[Dict]) -> str:
    """
    파일 목록에서 요약 생성 (텍스트 내용이 없을 때)
    """
    if not files:
        return "제출 없음"
    
    file_types = {}
    for f in files:
        ftype = f.get("type", "other")
        file_types[ftype] = file_types.get(ftype, 0) + 1
    
    parts = []
    if file_types.get("image", 0):
        parts.append(f"이미지 {file_types['image']}개")
    if file_types.get("md", 0) or file_types.get("txt", 0):
        text_count = file_types.get("md", 0) + file_types.get("txt", 0)
        parts.append(f"문서 {text_count}개")
    if file_types.get("code", 0):
        parts.append(f"코드 {file_types['code']}개")
    if file_types.get("pdf", 0):
        parts.append(f"PDF {file_types['pdf']}개")
    
    return " + ".join(parts) if parts else f"파일 {len(files)}개"


def summarize_all_members(scan_results: List[Dict]) -> List[Dict]:
    """
    모든 회원의 공부 내용 요약
    
    Args:
        scan_results: drive_scanner.scan_all_members()의 결과
    
    Returns:
        요약 정보가 추가된 결과 리스트
    """
    summarized = []
    
    for result in scan_results:
        member_summary = {
            "name": result["name"],
            "date": result["date"],
            "has_submission": result["has_submission"],
            "summary": ""
        }
        
        if not result["has_submission"]:
            member_summary["summary"] = "미제출"
        elif result.get("text_content"):
            # 텍스트 내용이 있으면 AI로 요약
            print(f"  📝 {result['name']} 내용 요약 중...")
            member_summary["summary"] = summarize_content(result["text_content"])
        else:
            # 텍스트 없으면 파일 목록으로 요약
            member_summary["summary"] = summarize_from_files(result.get("files", []))
        
        summarized.append(member_summary)
        print(f"    → {member_summary['summary']}")
    
    return summarized


def test_api():
    """Gemini API 연결 테스트"""
    try:
        model = init_gemini()
        response = model.generate_content("안녕하세요. 테스트입니다. 짧게 인사해주세요.")
        print(f"✅ Gemini API 연결 성공!")
        print(f"   응답: {response.text.strip()}")
        return True
    except Exception as e:
        print(f"❌ Gemini API 연결 실패: {e}")
        return False


if __name__ == "__main__":
    print("=== Gemini API 테스트 ===")
    if test_api():
        print("\n=== 요약 테스트 ===")
        test_content = """
# 오늘 공부 내용

## JavaScript 화살표 함수
- 기존 function 키워드 대신 => 사용
- this 바인딩이 다름
- 간결한 문법

## React useState
- 상태 관리 훅
- const [state, setState] = useState(초기값)
"""
        summary = summarize_content(test_content)
        print(f"요약 결과: {summary}")
