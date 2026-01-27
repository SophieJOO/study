"""
Slack 메시지 전송 모듈
- 생성된 이미지를 Slack DM으로 전송
"""
from pathlib import Path
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

from config import SLACK_BOT_TOKEN, SLACK_USER_ID


def get_slack_client() -> WebClient:
    """Slack 클라이언트 생성"""
    if not SLACK_BOT_TOKEN:
        raise ValueError("SLACK_BOT_TOKEN이 설정되지 않았습니다. .env 파일을 확인하세요.")
    
    return WebClient(token=SLACK_BOT_TOKEN)


def send_dm_with_image(image_path: Path, message: str = None) -> bool:
    """
    Slack DM으로 이미지 전송
    
    Args:
        image_path: 전송할 이미지 파일 경로
        message: 함께 보낼 메시지 (optional)
    
    Returns:
        성공 여부
    """
    if not SLACK_USER_ID:
        raise ValueError("SLACK_USER_ID가 설정되지 않았습니다. .env 파일을 확인하세요.")
    
    client = get_slack_client()
    
    try:
        # DM 채널 열기
        response = client.conversations_open(users=[SLACK_USER_ID])
        channel_id = response["channel"]["id"]
        
        # 이미지 업로드 및 전송
        with open(image_path, "rb") as file:
            response = client.files_upload_v2(
                channel=channel_id,
                file=file,
                filename=image_path.name,
                initial_comment=message or "📊 오늘의 스터디 인증 현황입니다!"
            )
        
        print(f"✅ Slack DM 전송 완료!")
        return True
        
    except SlackApiError as e:
        print(f"❌ Slack 전송 실패: {e.response['error']}")
        return False
    except Exception as e:
        print(f"❌ Slack 전송 오류: {e}")
        return False


def send_text_message(message: str) -> bool:
    """
    Slack DM으로 텍스트 메시지만 전송
    """
    if not SLACK_USER_ID:
        raise ValueError("SLACK_USER_ID가 설정되지 않았습니다.")
    
    client = get_slack_client()
    
    try:
        response = client.conversations_open(users=[SLACK_USER_ID])
        channel_id = response["channel"]["id"]
        
        client.chat_postMessage(
            channel=channel_id,
            text=message
        )
        
        print(f"✅ Slack 메시지 전송 완료!")
        return True
        
    except SlackApiError as e:
        print(f"❌ Slack 전송 실패: {e.response['error']}")
        return False


def test_connection() -> bool:
    """Slack 연결 테스트"""
    try:
        client = get_slack_client()
        response = client.auth_test()
        
        print(f"✅ Slack 연결 성공!")
        print(f"   Bot: {response['user']}")
        print(f"   Workspace: {response['team']}")
        return True
        
    except SlackApiError as e:
        print(f"❌ Slack 연결 실패: {e.response['error']}")
        return False
    except Exception as e:
        print(f"❌ Slack 연결 오류: {e}")
        return False


if __name__ == "__main__":
    print("=== Slack 연결 테스트 ===")
    test_connection()
