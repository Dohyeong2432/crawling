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
    find_url_yn = False
    last_list_yn = True 

    """
    ### 260227 업데이트 ###
        move_to_fss_home()으로 이동하면 주소에 페이지 번호가 없음
        페이지가 많아지만 연초에는 없던, 마지막 목록이 나타나고, 마지막 목록을 클릭하면 주소가 변화면서 마지막 페이지를 잡는 구조
            그러므로 마지막목록이 있으면 마지막목록 이동후 pageIndex를 찾고, 아니면 현재 active페이지의 페이지 번호 중 맨 마지막 페이지를 마지막 페이지로 특정
        특정 목록을 클릭해야 주소가 변홤(세부적인 파라미터들이 생기기 시작)
    """
    ## (260227) 새해가 되면 마지막목록이 disabled되므로 분개를 넣어줌 이부분 확인
    for div_tags in browser.find_elements(By.CLASS_NAME, 'pagination-set'):
        for li_tag in div_tags.find_elements(By.TAG_NAME, 'li'):
            class_name = li_tag.get_attribute('class')
            if 'end' in class_name:
                if 'disabled' in class_name: ## 마지막 페이지가 비활성화 되어 있으면 last_list_yn을 False로
                    last_list_yn = False

    if last_list_yn: ## 마지막목록이 있으면 원래 방법대로 처리하고
        for div_tags in browser.find_elements(By.CLASS_NAME, 'pagination-set'): ## (260227) pagination-set은 1개임
            for a_tag in div_tags.find_elements(By.TAG_NAME, 'a'):
                find_text = '마지막목록'
                a_tag_title = re.sub(r'[^가-힣]', '', a_tag.get_attribute('title')).strip()
                if a_tag_title==find_text:
                    find_url_yn = True
                    a_tag.click()
                    break
                    
    else: ## 마지막 목록이 없으면

        ## ul 태그 안에 숨겨진 마지막 페이지를 찾기. 먼저 centered 라는 클래스 검색(By.CLASS_NAME으로 검색안됨)
        for ul_tag in browser.find_elements(By.TAG_NAME, 'ul'):
            ul_class_name = ul_tag.get_attribute('class')
            if "pagination-centered" in ul_class_name:
                break

        ## 마지막 페이지를 일단 찾고
        page_nums = []
        for span_tag in ul_tag.find_elements(By.TAG_NAME, 'span'):
            if span_tag.text:
                page_nums.append(span_tag.text.strip())        
        
        ## 마지막 페이지로 이동
        for span_tag in ul_tag.find_elements(By.TAG_NAME, 'span'):
            if span_tag.text:
                if span_tag.text.strip()==page_nums[-1]:
                    find_url_yn = True
                    span_tag.click()
                    break
    
    if find_url_yn:
        ## 파라미터가 포함된 페이지 주소를 받음, move_to_page에서 본 주소의 페이지 번호만 정규표현식으로 수정하며 이동
        page_url = browser.current_url
        
        match = re.search(r'pageIndex=(\d+)', page_url)
        if match:
            page_idx = match.group(1)        
        
        time.sleep(0.1)        
        return page_url, page_idx
        
    else:
        print('마지막 페이지를 찾는데 실패하였음')    

## 페이지 정보와 페이지 넘버로 페이지를 이동
def move_to_page(browser, url, page_num):
    new_url = re.sub(r'pageIndex=\d+', f'pageIndex={str(page_num)}', url)
    browser.get(new_url)
    time.sleep(0.2)

## 해당페이지의 테이블 정보와 태그를 dataframe에 저장하고 반환
def get_table_info(browser):
    list_tag = browser.find_element(By.CLASS_NAME, 'bd-list')
    table_tag = list_tag.find_element(By.TAG_NAME, 'table')
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
                move_to_page(browser, page_num)
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
            
            move_to_page(browser, page_url, page_num)
            time.sleep(0.3)
            _, tag_df = get_table_info(browser)
        
    
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
    
