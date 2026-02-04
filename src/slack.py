# slack.py
import os
from slack_sdk import WebClient
from datetime import date, timedelta, datetime
from slack_sdk.errors import SlackApiError  # <--- 제거해도 됩니다.


def send_files_to_slack(token, channel, file_name, output_bytes, err_message):
    """
    slack 전송 함수
    (내부 try...except 제거됨. 오류 발생 시 예외를 상위로 전달)
    """
    client = WebClient(token=token)

    files_to_upload = []  # 업로드 파일 목록
    initial_comment = ""

    # 에러 발생 여부 확인
    is_error = err_message is not None

    last_week_date = date.today() - timedelta(days=7)
    thisweek = f"{last_week_date.isocalendar()[1]}주차"

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S KST")

    if is_error:
        # 실패 시 로직
        initial_comment = (
            f"🚨 *Coupang Play Weekly Report Failed*\n"
            f"• Week: {thisweek}\n"
            f"• Timestamp: {timestamp}\n"
            f"• Status: Failed  ❌"
        )
        # 에러 메시지를 텍스트 파일로 만들어 첨부
        files_to_upload.append(
            {"content": err_message, "filename": "error.txt", "filetype": "text"}
        )

    else:
        # 성공 시 로직
        if output_bytes and file_name:
            initial_comment = (
                f"📊 *Coupang Play Weekly Report Generated Successfully*\n"
                f"• Week: {thisweek}\n"
                f"• Timestamp: {timestamp}\n"
                f"• Status: Completed ✅"
            )
            # 결과 파일(바이트)만 첨부
            files_to_upload.append({"content": output_bytes, "filename": file_name})
        else:
            # 이 경우는 리포트 생성은 성공했으나, 알 수 없는 이유로
            # output_bytes나 file_name이 없는 경우입니다. (이것도 에러로 처리)
            initial_comment = (
                f"⚠️ *Coupang Play Weekly Report Anomaly*\n"
                f"• Week: {thisweek}\n"
                f"• Timestamp: {timestamp}\n"
                f"• Status: ⚠️ Warning (Report generated but no file content)"
            )

    # 파일 업로드 API 호출
    client.files_upload_v2(
        channel=channel,
        file_uploads=files_to_upload,
        initial_comment=initial_comment,
    )
