import json
import os
import sys
import urllib.error
import urllib.request


RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "")
TEST_RECIPIENT = os.environ.get("TEST_EMAIL_RECIPIENT", "")
API_URL = "https://api.resend.com/emails"


def send_test_email() -> str:
    if not RESEND_API_KEY:
        raise RuntimeError("请先设置 RESEND_API_KEY")
    if not TEST_RECIPIENT:
        raise RuntimeError("请先设置 TEST_EMAIL_RECIPIENT")
    payload = {
        "from": "Zhanyi Metal <contact@zhanyimetal.com>",
        "to": [TEST_RECIPIENT],
        "subject": "Zhanyi Metal 测试邮件",
        "html": (
            "<p>您好，这是一封通过 Resend 从 "
            "contact@zhanyimetal.com 发出的测试邮件。</p>"
            "<p>如果您收到此邮件，说明邮件发送配置正常。</p>"
        ),
    }

    request = urllib.request.Request(
        API_URL,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "zhanyi-metal-email/1.0",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        details = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Resend API 请求失败 ({error.code}): {details}") from error
    except urllib.error.URLError as error:
        raise RuntimeError(f"网络请求失败: {error.reason}") from error

    email_id = result.get("id")
    if not email_id:
        raise RuntimeError(f"Resend 返回了意外结果: {result}")
    return email_id


if __name__ == "__main__":
    try:
        sent_email_id = send_test_email()
    except RuntimeError as error:
        print(error, file=sys.stderr)
        raise SystemExit(1)

    print(f"测试邮件已提交，邮件 ID: {sent_email_id}")
