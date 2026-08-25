"""이메일 발송 공통 설정 로더.

notebooks/.env 파일 한 곳에서 발송 계정 / 앱비밀번호 / 수신자 목록을 읽어옵니다.
4개 자동화 프로젝트의 py_files/common_functions.py 가 모두 이 모듈을 사용하므로,
비밀번호를 변경할 일이 생기면 notebooks/.env 의 EMAIL_PASSWORD 한 줄만 고치면 됩니다.

크롤링/점검 단계는 메일을 보내지 않으므로 비밀번호 없이도 실행됩니다.
비밀번호는 실제 발송 직전 require_password() 에서만 검사합니다.

사용 예)
    from mail_env import load_mail_config, require_password
    cfg = load_mail_config("RECV_ADDRS_FSS_SANCTION")
    cfg.send_addr, cfg.password, cfg.recv_addrs ...
"""
import os
from pathlib import Path
from types import SimpleNamespace

from dotenv import load_dotenv

## 이 파일(notebooks/mail_env.py)과 같은 폴더의 .env 를 읽음.
## __file__ 기준이므로 노트북을 어느 폴더에서 실행하든 항상 같은 파일을 가리킴.
ENV_PATH = Path(__file__).resolve().parent / ".env"

## .env 를 새로 만들었을 때 들어있는 안내 문구. 이 값이 그대로면 아직 미설정 상태임.
PASSWORD_PLACEHOLDER = "여기에_새로_발급받은_앱비밀번호_입력"


def _fail(message):
    raise RuntimeError(f"[이메일 설정 오류] {message}\n  설정 파일 위치: {ENV_PATH}")


def get_env(key, required=True, default=None):
    """.env 에서 값 하나를 읽음. 값이 비어 있으면 원인을 알려주고 중단."""
    value = os.getenv(key)
    if value is None or not str(value).strip():
        if required:
            _fail(f"'{key}' 항목이 비어 있습니다. 해당 파일을 열어 값을 채워주세요.")
        return default
    return str(value).strip()


def get_recv_addrs(key):
    """콤마로 구분된 수신자 목록을 리스트로 변환."""
    addrs = [addr.strip() for addr in get_env(key).split(",") if addr.strip()]
    if not addrs:
        _fail(f"'{key}' 에 유효한 수신자 이메일이 없습니다. 콤마(,)로 구분해 입력해주세요.")
    return addrs


def require_password(password):
    """메일 발송 직전에 호출. 비밀번호가 준비되지 않았으면 발송을 막고 안내."""
    if not password or password == PASSWORD_PLACEHOLDER:
        _fail(
            "EMAIL_PASSWORD 가 아직 설정되지 않아 메일을 보낼 수 없습니다."
            "\n  네이버 > 내정보 > 보안설정 > 2단계 인증 > 애플리케이션 비밀번호 관리"
            "\n  에서 발급받은 값을 EMAIL_PASSWORD= 뒤에 입력한 뒤 다시 실행해주세요."
            "\n  (크롤링/점검 단계는 비밀번호 없이도 실행할 수 있습니다)"
        )
    return password


def load_mail_config(recv_addrs_key):
    """공통 발송 설정 + 프로젝트별 수신자 목록을 함께 읽어옴.

    recv_addrs_key : .env 에 정의된 수신자 항목 이름
                     (예: "RECV_ADDRS_FSS_SANCTION")
    """
    if not ENV_PATH.exists():
        _fail(".env 파일을 찾을 수 없습니다. .env.example 을 복사해 .env 로 만들어주세요.")

    ## override=True : 이미 셸 환경변수에 같은 이름이 있어도 .env 값을 우선 적용
    ##                 (노트북에서 .env 를 고치고 재실행했을 때 바로 반영되도록)
    load_dotenv(ENV_PATH, override=True)

    ## 비밀번호는 여기서 검사하지 않음 (점검 단계는 비밀번호 불필요).
    ## 실제 검사는 발송 직전 require_password() 에서 수행함.
    return SimpleNamespace(
        send_addr=get_env("SEND_ADDR"),
        password=get_env("EMAIL_PASSWORD", required=False, default=""),
        smtp_server=get_env("SMTP_SERVER"),
        smtp_port=int(get_env("SMTP_PORT", required=False, default="587")),
        download_folder=get_env("DOWNLOAD_FOLDER"),
        recv_addrs=get_recv_addrs(recv_addrs_key),
    )
