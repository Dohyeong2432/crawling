import pandas as pd
import time, warnings, re, os, logging, smtplib, pdfplumber, datetime
warnings.filterwarnings('ignore')
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.encoders import encode_base64
from bs4 import BeautifulSoup as bs
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup as bs
import xlwings as xw
from webdriver_manager.chrome import ChromeDriverManager
os.environ['WDM_SSL_VERIFY'] = '0' ## ChromeDriverManager().install() 실행 시 verify = False 옵션
# pdfplumber 로그 끄기
logging.getLogger("pdfplumber").setLevel(logging.ERROR)
logging.getLogger("pdfminer").setLevel(logging.ERROR)

############################################################################
## 이메일 발송 설정: notebooks/.env 에서 읽어옴
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

_cfg = load_mail_config("RECV_ADDRS_FSS_SANCTION")
SEND_ADDR       = _cfg.send_addr
EMAIL_PASSWORD  = _cfg.password
SMTP_SERVER     = _cfg.smtp_server
SMTP_PORT       = _cfg.smtp_port
RECV_ADDRS      = _cfg.recv_addrs
DOWNLOAD_FOLDER = _cfg.download_folder



############################################################################
## 제재 검색 조건. 노트북과 run_check.py 가 이 값을 함께 사용하므로
## 조건을 바꿀 일이 있으면 반드시 여기만 고칠 것.
############################################################################
## 관련부서에 이 문자열이 들어가면 대상
FIND_CATEGORIES = [
    '자금세탁',
]

## PDF 본문에 이 키워드가 있으면 대상
FIND_PDF_KEYWORDS = [
    'CDO', '고객확인', 'STR', '의심거래', 'CTR', '고액현금거래',
    'AML', '자금세탁', '특정금융거래정보', '전기통신', '고객알기', '지주',
]

## 금감원 제재사이트 주소
def get_fss_url(): 
    return "https://www.fss.or.kr/fss/job/openInfo/list.do?menuNo=200476"

def get_browser():    
    return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

def move_to_fss_home(browser):
    url = get_fss_url()
    browser.get(url)
    time.sleep(0.2)

## 페이지 정보(url)와 마지막 페이지 번호를 가져옴
## 페이지 정보(url) 주소의 pageIndex 값을 변경하면서 페이지 이동함
def get_last_page_info(browser):
    """목록 페이지의 URL(파라미터 포함)과 마지막 페이지 번호를 반환한다.

    [2026-09 사이트 개편]
    예전에는 'pagination-set' 안의 '마지막목록' 버튼을 클릭해 마지막 페이지로 갔다.
    개편 후 그 클래스와 버튼이 모두 사라졌고, 지금은 이런 구조다.
        <ul class="pagination-krsd pagination-centered">
            <a href="javascript:fnSearch(2)">2</a> ... <a href="javascript:fnSearch(26)">26</a>
    - 보이는 링크의 fnSearch(N) 중 가장 큰 N 이 마지막 페이지다.
      (가운데 '...' 링크가 마지막 페이지로 바로 가는 역할을 한다)
    - 첫 화면 URL 에는 pageIndex 가 없다. fnSearch 를 한 번 실행해야
      pageIndex/sdate/edate 등 파라미터가 붙은 URL 이 만들어지고,
      그래야 move_to_page() 가 pageIndex 만 바꿔가며 이동할 수 있다.
    """
    page_nums = []
    for ul_tag in browser.find_elements(By.TAG_NAME, 'ul'):
        ul_class = ul_tag.get_attribute('class') or ''
        if 'pagination' not in ul_class:
            continue
        for a_tag in ul_tag.find_elements(By.TAG_NAME, 'a'):
            href = a_tag.get_attribute('href') or ''
            match = re.search(r'fnSearch\((\d+)\)', href)
            if match:
                page_nums.append(int(match.group(1)))

    if not page_nums:
        print('마지막 페이지를 찾는데 실패하였음 (페이지 링크를 찾지 못함)')
        return None

    last_page_num = max(page_nums)

    ## 파라미터가 붙은 URL 을 얻기 위해 1페이지로 한 번 이동한다.
    browser.execute_script('fnSearch(1);')
    time.sleep(1.0)
    page_url = browser.current_url

    if not re.search(r'pageIndex=\d+', page_url):
        print(f'마지막 페이지를 찾는데 실패하였음 (URL 에 pageIndex 가 없음: {page_url})')
        return None

    return page_url, last_page_num

## 페이지 정보와 페이지 넘버로 페이지를 이동
def move_to_page(browser, url, page_num):
    new_url = re.sub(r'pageIndex=\d+', f'pageIndex={str(page_num)}', url)
    browser.get(new_url)
    time.sleep(0.2)

## 해당페이지의 테이블 정보와 태그를 dataframe에 저장하고 반환
def get_table_info(browser):
    """현재 페이지의 표 내용과 각 칸의 태그를 dataframe 두 개로 반환한다.

    [2026-09 사이트 개편]
    예전에는 'bd-list' div 안에서 표를 찾았는데 그 클래스가 사라졌다.
    지금은 <div class="krds-table-wrap"><table class="tbl col list-data"> 구조라
    표를 직접 찾는다. 컬럼 구성(번호/제재대상기관/제재조치요구일/제재조치요구내용/
    관련부서/조회수)은 개편 전후가 같다.
    """
    table_tag = None
    for candidate in browser.find_elements(By.TAG_NAME, 'table'):
        if 'list-data' in (candidate.get_attribute('class') or ''):
            table_tag = candidate
            break
    if table_tag is None:  ## 클래스가 또 바뀌었을 때를 대비한 대비책
        table_tag = browser.find_element(By.TAG_NAME, 'table')

    head_tag = table_tag.find_element(By.TAG_NAME, 'thead')
    body_tag = table_tag.find_element(By.TAG_NAME, 'tbody')

    table_df_cols = []
    for th_tag in head_tag.find_elements(By.TAG_NAME, 'th'):
        table_df_cols.append(th_tag.text.strip())
    table_df = pd.DataFrame(columns=table_df_cols)
    tag_df = pd.DataFrame(columns=table_df_cols)
    for tr_tag in body_tag.find_elements(By.TAG_NAME, 'tr'):
        row_idx = len(table_df)
        for col_idx, td_tag in enumerate(tr_tag.find_elements(By.TAG_NAME, 'td')):
            table_df.loc[row_idx, table_df_cols[col_idx]] = td_tag.text
            tag_df.loc[row_idx, table_df_cols[col_idx]] = td_tag
    return table_df, tag_df

## 다운로드 받으려는 PDF 파일이 있을 경우 먼저 삭제
def remove_dup_files(download_folder, remove_file):    
    folder_path = download_folder
    for file in os.listdir(folder_path):
        fn, _ = os.path.splitext(file)    
        if fn.find(remove_file)>-1:        
            os.remove(os.path.join(folder_path, file))

## 방금 내려받은 파일을 지운다.
## Windows 에서는 브라우저가 파일을 잠시 붙들고 있어 곧바로 지우면 실패할 수 있으므로
## 짧게 기다리며 몇 번 다시 시도한다.
def remove_downloaded_file(path, attempts=5, interval=0.4):
    for attempt in range(attempts):
        try:
            os.remove(path)
            return True
        except OSError:
            if attempt == attempts - 1:
                return False
            time.sleep(interval)
    return False

## 안내할 제재리스트 찾기
def get_notify_list(min_date, find_categories=None, find_pdf_keywords=None, download_folder=None):
    ## 조건을 따로 넘기지 않으면 이 파일 상단의 공용 상수를 사용한다.
    if find_categories is None:
        find_categories = FIND_CATEGORIES
    if find_pdf_keywords is None:
        find_pdf_keywords = FIND_PDF_KEYWORDS
    if download_folder is None:  ## 지정하지 않으면 .env 의 DOWNLOAD_FOLDER 사용
        download_folder = DOWNLOAD_FOLDER
    browser = get_browser()
    move_to_fss_home(browser)
    page_url, last_page_num = get_last_page_info(browser)

    send_df = []
    send_files = []
    find_keywords_list = []
    removed_files = []  ## 안내 대상이 아니어서 지운 PDF
    unremoved_files = []  ## 지우려 했으나 실패한 PDF
    started_at = time.time()  ## 이 시각 이후 저장된 파일이 이번 실행의 다운로드분
    find_continue = True
    for page_num in range(1, int(last_page_num)+1):
        if find_continue==False:
            break
        
        move_to_page(browser, page_url, page_num)    
        table_df, tag_df = get_table_info(browser)    
    
        ## 해당페이지의 페이지 정보를 한줄씩 확인하며 작업을 처리함
        for idx, rows in table_df.iterrows():
            if rows['제재조치요구일']<min_date:
                find_continue = False
                break
            send_mail_yn = False
            find_keywords = ''
    
            ## 관련부서 해당 여부 확인
            for find_category in find_categories:
                if rows['관련부서'].find(find_category)>-1:
                    send_mail_yn = True
        
            tag_df.loc[idx, '제재조치요구내용'].click() ## 일단 제재조치요구내용을 클릭함
            
            ## 기존에 받은 파일이 있다면 삭제        
            file_name = browser.find_element(By.CLASS_NAME, 'name').text    
            remove_dup_files(download_folder, os.path.splitext(file_name)[0])
            browser.find_element(By.CLASS_NAME, 'name').click()  ## 파일 다운로드   
            
    
            ## PDF를 읽고 관련 키워드가 있는지 확인
            pdf_file = os.path.join(download_folder, file_name)
            pdf_text = ""
            ## 파일이 내려올 때까지 대기. 다운로드가 실패하면 영영 끝나지 않으므로 제한시간을 둔다.
            deadline = time.time() + 60
            while not os.path.exists(pdf_file):
                if time.time() >= deadline:
                    print(f'[경고] 60초 안에 다운로드되지 않아 건너뜁니다: {file_name}')
                    break
                time.sleep(0.2)
            if not os.path.exists(pdf_file):
                move_to_page(browser, page_url, page_num)
                time.sleep(0.3)
                _, tag_df = get_table_info(browser)
                continue
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    pdf_text = pdf_text + page.extract_text()        
            pdf_text = re.sub(r'[^가-힣]', '', pdf_text).strip()
            
            
            ## pdf에서 키워드를 찾고. 해당키워드가 존재하면 find_keywords에 저장
            for find_pdf_keyword in find_pdf_keywords:
                if pdf_text.find(find_pdf_keyword)>-1:
                    send_mail_yn = True
                    find_keywords = find_keywords + "," + find_pdf_keyword
            
            if send_mail_yn:
                send_df.append(table_df.loc[idx].to_frame().T)
                send_files.append(pdf_file)
                find_keywords_list.append(find_keywords[1:]) ## 맨처음 콤마(,)는 지우고...
            else:
                ## 안내 대상이 아니면 받아둔 PDF 를 지운다.
                ## 키워드 확인용으로만 내려받은 파일이라 다운로드 폴더에 남길 이유가 없다.
                if remove_downloaded_file(pdf_file):
                    removed_files.append(os.path.basename(pdf_file))
                else:
                    unremoved_files.append(os.path.basename(pdf_file))
            
            move_to_page(browser, page_url, page_num)
            time.sleep(0.3)
            _, tag_df = get_table_info(browser)
        
    
    ## 마무리 정리.
    ## 행마다 지워도, 삭제 직후 브라우저가 같은 이름으로 파일을 다시 쓰는 경우가 있어
    ## 파일이 되살아난다. 그래서 끝에서 한 번 더 훑는다.
    ## 기준: 이번 실행 중에 저장되었고(started_at 이후), 안내 대상이 아닌 PDF.
    keep = set(os.path.abspath(f) for f in send_files)
    for filename in os.listdir(download_folder):
        file_path = os.path.join(download_folder, filename)
        if not os.path.isfile(file_path) or os.path.splitext(filename)[1].lower() != '.pdf':
            continue
        if os.path.abspath(file_path) in keep:          ## 안내 대상은 남긴다
            continue
        if os.path.getmtime(file_path) < started_at:    ## 이번 실행과 무관한 기존 파일
            continue
        if remove_downloaded_file(file_path):
            if filename not in removed_files:
                removed_files.append(filename)
        elif filename not in unremoved_files:
            unremoved_files.append(filename)

    if removed_files:
        print(f'안내 대상이 아닌 PDF {len(removed_files)}건을 다운로드 폴더에서 삭제했습니다.')
    if unremoved_files:
        print(f'[알림] 아래 {len(unremoved_files)}건은 삭제하지 못했습니다. 직접 정리해주세요.')
        for name in unremoved_files:
            print(f'         - {name}')

    if send_df:
        send_df = pd.concat(send_df)
        send_df['find_keywords'] = find_keywords_list
        send_df['pdf_files'] = send_files
        return send_df
    else:
        print('안내할 제재 내용이 없습니다.')
        return None
    
## 최종적으로 이메일 발송
def send_mail(update_df):
    require_password(EMAIL_PASSWORD)  ## 비밀번호 미설정 시 발송 전에 중단
    update_df = update_df.copy()
    today = datetime.date.today().strftime('%y%m%d')
    mail_title = f"{(today)}_금융감독원_제재사항_안내"
    mail_body = mail_title

    file_list = update_df.pdf_files.values.tolist()
    update_df['pdf_files'] = update_df['pdf_files'].map(lambda x: os.path.basename(x))
    file_name = f"./update_history/{today}_금융감독원_제재사항_리스트.xlsx"
    update_df.to_excel(file_name, index=False)

    app = xw.App(visible=True)
    wb = app.books.open(file_name)
    sht = wb.sheets[0]    
    sht.autofit()  # 열/행 자동 맞춤
    wb.save()
    wb.close()
    app.quit()

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
        for file in [file_name] + file_list:
            part = MIMEBase('application', "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(open(file, "rb").read())
            encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
            msg.attach(part)
        
        mail_server.sendmail(SEND_ADDR, msg['To'].split(','), msg.as_string())

    mail_server.quit()  ## SMTP 연결 정리
    
