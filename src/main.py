import traceback
from datetime import datetime
import threading
from config import config
from slack import send_files_to_slack
from make_report import create_weekly_report
import argparse
import sys

parser = argparse.ArgumentParser()
parser.add_argument("--env", type=str, default="prod", help="실행 환경 (prod 또는 dev)")
args = parser.parse_args()


def run_report_and_notify():
    file_name = None
    output_bytes = None
    err_message = None

    try:
        # --- 1. 리포트 생성 시도 ---
        print("리포트 생성 시도...")
        file_name, output_bytes = create_weekly_report(config.NDAP_INFO)
        print("리포트 생성 성공.")

    except Exception:
        # --- 2. 리포트 생성 실패 ---
        # 리포트 생성 중 발생한 모든 에러를 캡처
        err_message = traceback.format_exc()
        print(f"스크립트 실행 중 에러가 발생했습니다:\n{err_message}", file=sys.stderr)

    finally:
        # --- 3. 알림 전송 시도 (성공/실패 모두) ---
        print("슬랙 알림 시도...")

    if args.env == "prod":
        slack_channel = config.SLACK_CHANNEL
    else:
        slack_channel = config.SLACK_CHANNEL_TEST

    # 슬랙으로 파일 또는 에러 메시지 전송
    send_files_to_slack(
        token=config.SLACK_TOKEN,
        channel=slack_channel,
        file_name=file_name,
        output_bytes=output_bytes,
        err_message=err_message,  # 리포트 생성 에러(있다면)를 전달
    )
    print("슬랙 알림 전송 성공")


def main():

    report_thread = threading.Thread(target=run_report_and_notify)
    report_thread.start()


if __name__ == "__main__":
    main()
