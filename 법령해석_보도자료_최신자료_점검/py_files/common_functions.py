import warnings, os, re, smtplib, time
from webdriver_manager.chrome import ChromeDriverManager

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.encoders import encode_base64

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
os.environ['WDM_SSL_VERIFY'] = '0' ## ChromeDriverManager().install() 실행 시 verify = False 옵션
warnings.filterwarnings('ignore')

options = Options()
prefs = {
    "safebrowsing.enabled": True, ## 자동다운로드 차단 강제로 끄기
    "plugins.always_open_pdf_externally": True, ## 파일 자동열기 끄기
    } 
options.add_experimental_option("prefs", prefs)

############################################################################
## 이메일 발송 설정: notebooks/.env 에서 읽어옴 (프로젝트별 .env 아님)
## 비밀번호를 이 파일이나 노트북에 직접 적지 말 것 (.env 만 고치면 됨)
############################################################################
import sys
from pathlib import Path
_NOTEBOOKS_ROOT = Path(__file__).resolve().parents[2]   ## notebooks 폴더
if str(_NOTEBOOKS_ROOT) not in sys.path:
    sys.path.insert(0, str(_NOTEBOOKS_ROOT))
from mail_env import load_mail_config, require_password

## display() 는 노트북에서만 기본 제공됨. 아래 import 로 일반 python 실행도 가능해짐
from IPython.display import display

_cfg = load_mail_config("RECV_ADDRS_PRESS")
SEND_ADDR       = _cfg.send_addr
EMAIL_PASSWORD  = _cfg.password
SMTP_SERVER     = _cfg.smtp_server
SMTP_PORT       = _cfg.smtp_port
RECV_ADDRS      = _cfg.recv_addrs
DOWNLOAD_FOLDER = _cfg.download_folder


"""
    FSS: 금융감독원
    FSC: 금융위원회
    FIU: 금융정보분석원
    PRESS: 보도자료
    ITP: 법령해석(Interpreter)
"""

## 다운로드 받으려는 PDF 파일이 있을 경우 먼저 삭제
## remove_file -> 파일의 이름(확장자 제외)
def remove_dup_files(remove_file):   
    remove_file = re.sub(" ", "", remove_file) 
    folder_path = DOWNLOAD_FOLDER
    for file in os.listdir(folder_path):
        fn, _ = os.path.splitext(file)    
        fn = re.sub(" ", "", fn)
        if fn.find(remove_file)>-1:        
            os.remove(os.path.join(folder_path, file))

## 첨부파일 다운로드 완료 대기.
##   download_files : 확장자를 뺀 파일명 리스트(공백이 제거된 형태)
##   반환값         : 다운로드 폴더에서 찾은 실제 파일 경로 리스트
##
## 사이트가 알려주는 파일명과 실제 저장되는 파일명이 다른 경우가 있어
## (공백 처리 차이) 파일이 영영 나타나지 않을 수 있다.
## 제한시간이 지나면 그때까지 받아진 것만 반환하고 넘어간다.
def wait_for_downloads(download_files, timeout=30.0, interval=0.1):
    def _found():
        return [
            os.path.join(DOWNLOAD_FOLDER, f)
            for f in os.listdir(DOWNLOAD_FOLDER)
            if re.sub(" ", "", os.path.splitext(f)[0]) in download_files
        ]

    if not download_files:  ## 첨부파일이 없는 게시글도 있음
        return []

    deadline = time.monotonic() + timeout
    while True:
        files = _found()
        if len(files) == len(download_files):
            time.sleep(interval)  ## 파일 기록이 끝나도록 짧게 여유를 둠
            return _found()
        if time.monotonic() >= deadline:
            done = set(re.sub(" ", "", os.path.splitext(os.path.basename(f))[0]) for f in files)
            missing = sorted(set(download_files) - done)
            print(f"[경고] 첨부파일 다운로드가 {timeout:.0f}초 안에 끝나지 않았습니다. "
                  f"누락 {len(missing)}건: {missing}")
            return files
        time.sleep(interval)  ## 반드시 쉬어줄 것. 없으면 CPU 를 100% 점유함


## 발송 대상 제목인지 확인한다.
##   keywords 가 비어 있으면(None/빈 목록) 전부 대상으로 본다.
## 사용자가 사이트에서 제목을 복사해 넘기는 경우가 많아 띄어쓰기가 어긋나기 쉬우므로,
## 공백을 모두 없앤 뒤 부분일치로 비교한다.
def subject_matches(subject, keywords=None):
    if not keywords:
        return True
    text = re.sub(r'\s+', '', str(subject))
    return any(re.sub(r'\s+', '', str(k)) in text for k in keywords if str(k).strip())


## 게시판의 날짜 문자열에서 날짜만 뽑아 'yymmdd' 로 변환.
##
## 사이트가 날짜 칸에 다른 정보를 덧붙이는 경우가 있다.
## 예) 금융정보분석원은 2026년경부터 li_date 안에 조회수를 함께 넣어
##     '2026-08-14 / 조회수: 928' 처럼 나오고, pd.to_datetime() 은 여기서 실패한다.
## 그래서 문자열 전체를 파싱하지 않고 날짜 패턴만 찾아서 쓴다.
_DATE_PATTERN = re.compile(r'(\d{4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})')

def parse_press_date(text):
    matched = _DATE_PATTERN.search(str(text))
    if not matched:
        raise ValueError(f"날짜를 찾지 못했습니다: {text!r}")
    year, month, day = (int(g) for g in matched.groups())
    return f"{year % 100:02d}{month:02d}{day:02d}"


def get_browser():    
    while True:
        try:            
            browser = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)            
            break
        except:
            continue
    return browser

def move_to_home(browser, work_type):
    if work_type=='fss_itp': ## 금융위 법령해석
        browser.get("https://better.fsc.go.kr/fsc_new/replyCase/PastReplyList.do?stNo=11&muNo=171&muGpNo=75")
    elif work_type=='fsc_press': ## 금융위 보도자료
        browser.get("https://www.fsc.go.kr/no010101")
    elif work_type=='fss_press': ## 금감원 보도자료
        browser.get("https://www.fss.or.kr/fss/bbs/B0000188/list.do?menuNo=200218")
    elif work_type=='fiu_press': ## 금융정보분석원 보도자료
        browser.get("https://www.kofiu.go.kr/kor/notification/report.do")
    else:
        pass

## 이메일 발송
def send_mail(mail_title, mail_body, files):
    require_password(EMAIL_PASSWORD)  ## 비밀번호 미설정 시 발송 전에 중단
    # 메일 객체 생성 및 로그인
    mail_server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    mail_server.ehlo()
    mail_server.starttls()
    mail_server.ehlo()
    mail_server.login(SEND_ADDR, EMAIL_PASSWORD)
    
    for recv_addr in RECV_ADDRS:
        # 제목, 본문 작성
        msg = MIMEMultipart()
        msg['From'] = SEND_ADDR
        msg['To'] = recv_addr
        msg['Subject'] = mail_title
        msg.attach(MIMEText(mail_body, _charset='utf-8'))           
        
        # 파일첨부
        if files: ## 첨부파일이 없을수도 있을까봐            
            for file in files:
                if file:                
                    part = MIMEBase('application', "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    part.set_payload(open(file, "rb").read())
                    encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
                    msg.attach(part)

        ## 최종 발송
        mail_server.sendmail(SEND_ADDR, msg['To'].split(','), msg.as_string())

    mail_server.quit()  ## SMTP 연결 정리