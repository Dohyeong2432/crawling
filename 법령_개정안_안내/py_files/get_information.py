import pandas as pd
import time, re, os, smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.encoders import encode_base64

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By

os.environ['WDM_SSL_VERIFY'] = '0' ## ChromeDriverManager().install() 실행 시 verify = False 옵션
from .common_functions import *

#######################################################################
## 작업을 수행해야할 법령/행동규칙을 찾고, 법규리스트의 최근개정일을 업데이트함
## update_df의 law_name은 한글이외 문자 모두 삭제한 버전
def get_update_df():    
    law_file_name = "./database/법규리스트.xlsx"
    law_list = pd.read_excel(law_file_name)

    df_dict = dict()
    site_categories = ['law', 'reg']
    for site_category in site_categories:
        name_col, date_col = get_column_name(site_category)    
        file = sorted(glob.glob(f"./database/*{site_category}*.parquet"))[-1]
        df = pd.read_parquet(file)
        df[name_col] = df[name_col].map(lambda x: re.sub(r'[^가-힣]', '', x).strip())
        df[date_col] = pd.to_datetime(df[date_col])
        df_dict[site_category] = df

    law_list['최근발송일'] = pd.to_datetime(law_list['최근발송일'].astype(str))

    ### 업데이트할 법령을 찾음
    update_df = pd.DataFrame(columns=['site_category', 'law_name', '최근발송일', '최근개정일'])
    for idx, rows in law_list.iterrows():
        law_name, update_date = rows['법령명'], rows['최근발송일']
        law_update_date = rows['최근개정일']
        site_categories = ['law', 'reg']
        for site_category in site_categories:
            name_col, date_col = get_column_name(site_category)            
            url = get_url(site_category)
            law_df = df_dict[site_category]
            if law_name in law_df[name_col].values.tolist():
                ## 업데이트 되어 있지 않으면 업데이트 목록에 추가하고 종료
                ## 기준은 최근발송일이 최근개정일보다 작을때임!
                if update_date < law_df[law_df[name_col]==law_name][date_col].max():
                    update_df.loc[len(update_df)] = [
                        site_category, law_name, update_date.strftime("%Y%m%d"), law_update_date
                        ]
                    break

    ## 엑셀파일을 업데이트함(최근개정일 업데이트)
    ## 아래 루프에서 문자열 날짜를 대입하므로 미리 문자열 컬럼으로 맞춰둠
    ## (int64 컬럼에 문자열을 넣으면 pandas FutureWarning, 향후 버전에서는 오류)
    law_list['최근개정일'] = law_list['최근개정일'].astype(str)

    ## law/reg 어디에서도 찾지 못한 법령명을 모아 마지막에 경고함
    not_found_names = []
    for idx, rows in law_list.iterrows():
        law_name = rows['법령명']
        ## 행마다 반드시 초기화할 것.
        ## (초기화가 없으면 매칭 실패 시 '직전 행'의 날짜가 그대로 기록되어
        ##  잘못된 값이 조용히 저장되고, 명칭 오류를 알아챌 수 없음)
        lastest_update_date = None
        site_categories = ['law', 'reg']
        for site_category in site_categories:
            name_col, date_col = get_column_name(site_category)
            law_df = df_dict[site_category]
            if law_name in law_df[name_col].values.tolist():
                lastest_update_date = law_df[law_df[name_col]==law_name][date_col].max().strftime("%Y%m%d")
                break  ## 위쪽 탐지 루프와 동일하게 law 를 우선함

        if lastest_update_date is None:  ## 못 찾았으면 기존 값을 건드리지 않고 기록만 남김
            not_found_names.append(law_name)
            continue
        law_list.loc[idx, '최근개정일'] = lastest_update_date

    ## 매칭 실패는 곧 '개정 안내가 나가지 않는 법령'을 뜻하므로 반드시 표면화함
    if not_found_names:
        print("=" * 78)
        print(f"[경고] 아래 {len(not_found_names)}건은 법령/행정규칙 목록에서 찾지 못했습니다.")
        print("       법규리스트.xlsx 의 법령명이 국가법령정보센터 명칭과 다르거나,")
        print("       해당 사이트에 없는 자료(약관 등)일 수 있습니다.")
        print("       이 항목들은 개정되어도 안내 메일이 발송되지 않으니 확인이 필요합니다.")
        for not_found_name in not_found_names:
            print(f"         - {not_found_name}")
        print("=" * 78)
    law_list['최근발송일'] = law_list['최근발송일'].dt.strftime("%Y%m%d")

    for col in ['최근발송일', '최근개정일']: ## 혹시몰라 한번더 데이터타입 정리
        law_list[col] = law_list[col].astype(int).astype(str)

    law_list.to_excel(law_file_name, index=False)       
    
    app = open_excel_app()
    wb = app.books.open(law_file_name)
    sht = wb.sheets[0]    
    sht.autofit()  # 열/행 자동 맞춤
    wb.save()
    wb.close()
    app.quit()

    return update_df

## 법률/규칙에서 제정개정이유를 클릭하고 이동한 후, 우측 상단의 파일 저장 수행
def download_file(browser):
    ## 우측 상단의 저장 버튼을 클릭
    while True:
        try:
            browser.find_element(By.ID, "bdySaveBtn").click()
            break
        except:
            time.sleep(1)
    time.sleep(0.3)

    ## 저장 버튼 클릭시 파일 저장이 가능한 inner_box가 새로 호출됨. 그걸 찾음(제목이 법령은 목록저장, 규칙은 내용저장임)
    inner_box_tag = None
    for bx_tag in browser.find_elements(By.CLASS_NAME, 'bx_inner'):
        menu_text = bx_tag.find_element(By.TAG_NAME, 'p').text.strip()
        if menu_text == '목록저장' or menu_text == '내용저장':
            inner_box_tag = bx_tag
            break

    ## 저장 파일 양식을 hwp로 설정하고 저장함
    for tag in inner_box_tag.find_elements(By.TAG_NAME, 'label'):
        if tag.text.strip() == "HWP(한글)":
            tag.click()
            time.sleep(0.3)
            break

    for tag in inner_box_tag.find_elements(By.TAG_NAME, 'a'):
        if tag.text.strip() == '저장':
            tag.click()
            time.sleep(0.3)
            break

## 해당 법령 페이지에서 법령번호를 찾아 클릭함.
## find_num. 클릭할 행의 이름, law_name. 해당페이지 로딩 완료 여부 확인용.
def click_law_row(browser, find_num, law_name):
    find_continue = True
    for tr_tag in browser.find_elements(By.TAG_NAME, 'tr'):
        if find_continue==False:
            break
        for td_tag in tr_tag.find_elements(By.TAG_NAME, 'td'):
            if td_tag.text.strip() == str(find_num):  ## 페이지 번호 확인하고 맞으면 다음 tag를 가져옴
                click_tag = tr_tag.find_elements(By.TAG_NAME, 'td')[1]
                find_continue = False
                break

    click_tag.find_element(By.TAG_NAME, 'a').click()

    # 세부페이지 로드 될때까지 대기
    break_yn = False
    while True:
        page_source = bs(browser.page_source)    
        for tag in page_source.findAll(name='h2'): ## 제목이 해당법령명과 동일함(h2태그)
            if re.sub(r'[^가-힣]', '', tag.text).strip()==law_name:
                break_yn = True
        if break_yn:
            break
        else:
            time.sleep(1)
    
    time.sleep(0.5)

################################################################################################################
## 법령번호를 찾아 클릭한 후 제정개정이유와 신구법비교를 클릭해 파일을 저장하고, 메일을 발송할 본문(제정개정이유)를 텍스트로 반환
def get_law_information(browser, law_name):
    ## 각각의 메뉴에서 작업할 메뉴들
    get_menus = ['제정개정이유', '신구법비교']

    # 기존 창의 핸들 저장
    original_window = browser.current_window_handle

    ## 반드시 미리 초기화할 것.
    ## '제정개정이유' 메뉴가 없는 법령을 만나면 아래 if 문에서 UnboundLocalError 가 나서
    ## 발송 작업 전체가 중단됨(그 법령뿐 아니라 뒤에 남은 법령까지 전부 못 보냄).
    mail_text = None

    for get_menu in get_menus:
        menu_tag = browser.find_element(By.CLASS_NAME, "body_top_area").find_element(By.CLASS_NAME, "l_bx")
        new_window = None  ## 이번 메뉴에서 실제로 새 창이 열렸는지 여부
        for tag in menu_tag.find_elements(By.TAG_NAME, 'a'):
            text = re.sub(r'[^가-힣]', '', tag.text).strip()
            if text==get_menu: ## 작업을 수행하는 메뉴인 경우에만 작업하고 넘어감
                tag.click()
                # 새 창이 뜰 때까지 기다리기
                WebDriverWait(browser, 10).until(lambda x: len(x.window_handles) > 1)
                # 새 창으로 전환
                for handle in browser.window_handles:
                    if handle != original_window:
                        browser.switch_to.window(handle)
                        new_window = handle
                        break

                # 세부페이지 로드 될때까지 대기 (열리지 않는 경우가 있어 제한시간을 둠)
                break_yn = False
                deadline = time.time() + 30
                while True:
                    page_source = bs(browser.page_source, 'html.parser')
                    for h2_tag in page_source.findAll(name='h2'): ## h2 태그의 제목이 해당 법령과 동일하거나(제정개정이유)
                        if re.sub(r'[^가-힣]', '', h2_tag.text).strip()==law_name:
                            break_yn = True
                    for img_tag in page_source.findAll(name='img'): ## 신구법비교의 이전법령 버튼이 나타나거나
                        if img_tag.get('alt')=='이전법령':
                            break_yn = True
                    if break_yn:
                        break
                    if time.time() >= deadline:
                        print(f"[경고] '{law_name}' 의 '{get_menu}' 페이지가 30초 안에 열리지 않았습니다.")
                        break
                    time.sleep(1)
                time.sleep(1)
                download_file(browser) ## 새로 열린 창에서 hwp 파일을 저장

                if text=='제정개정이유': ## 제정개정이유 메뉴에서는 본문의 제정개정이유 텍스트를 별도 저장
                    page_source = bs(browser.page_source, 'html.parser')
                    for p_tag in page_source.findAll(name='p'):
                        if re.sub(r'[^가-힣]', '', p_tag.text).strip()=='제정개정이유':
                            found_title = p_tag.text.strip()
                            found_body = None
                            sibling = p_tag
                            while True: ## 제정개정이유 본문은 div 태그의 pgroup 클래스
                                sibling = sibling.find_next_sibling()
                                if sibling is None:  ## 형제 태그가 끝났는데 못 찾은 경우(무한루프/AttributeError 방지)
                                    break
                                if sibling.get('class') and sibling.get('class')[0]=='pgroup':
                                    found_body = sibling.get_text(separator="\n", strip=True)
                                    break
                            if found_body:
                                mail_text = found_title + "\n\n" + found_body
                            break
                break

        ## 새 창이 실제로 열렸을 때만 닫을 것.
        ## 메뉴를 못 찾았는데 close() 하면 '원래 창'이 닫혀서 이후 모든 작업이 실패함.
        if new_window:
            browser.close()
            browser.switch_to.window(original_window)
            time.sleep(0.5)
        else:
            print(f"[알림] '{law_name}' 에는 '{get_menu}' 메뉴가 없어 건너뜁니다.")

    ## 제정/개정이유가 없는 법령도 있음.
    ## 그런 경우에도 첨부파일과 함께 메일은 나가야 하므로 None 대신 안내 문구를 돌려줌
    ## (None 을 그대로 넘기면 MIMEText(None) 에서 발송 자체가 실패함)
    if mail_text:
        return mail_text
    return "이 법령은 제정·개정이유 본문이 제공되지 않습니다. 첨부한 파일을 참고해주세요."


## 날짜 표기를 'YYYYMMDD' 문자열로 통일
## 사이트의 공포일자는 '2026. 3. 5.' 형태라 그대로 쓰면 비교가 안 됨
def to_yyyymmdd(value):
    nums = re.findall(r'\d+', str(value))
    if len(nums) == 1 and len(nums[0]) == 8:   ## 이미 20260305 형태
        return nums[0]
    if len(nums) >= 3:                          ## '2026. 3. 5.' -> 20260305
        return f"{int(nums[0]):04d}{int(nums[1]):02d}{int(nums[2]):02d}"
    return pd.to_datetime(value).strftime("%Y%m%d")


## 개정 안내를 발송한 법령의 '최근발송일'을 법규리스트에 기록.
##
## 발송 대상 판정 기준이 '최근발송일 < 최근개정일' 이므로,
## 이 기록을 남기지 않으면 다음 실행 때 같은 법령이 다시 대상으로 잡혀
## 수신자에게 중복 메일이 발송된다. 반드시 발송 직후에 호출할 것.
def mark_as_sent(law_name, sent_date):
    law_file_name = "./database/법규리스트.xlsx"
    close_law_list_excel()  ## 엑셀에서 열려 있으면 저장이 실패하므로 먼저 닫음
    law_list = pd.read_excel(law_file_name)

    ## 비교용으로만 정규화함. 원본 법령명 컬럼은 그대로 둠
    normalized = law_list['법령명'].map(lambda x: re.sub(r'[^가-힣]', '', str(x)).strip())
    mask = normalized == re.sub(r'[^가-힣]', '', str(law_name)).strip()

    if not mask.any():
        print(f"[경고] '{law_name}' 을 법규리스트에서 찾지 못해 최근발송일을 기록하지 못했습니다.")
        print("       다음 실행 때 같은 법령이 다시 발송될 수 있으니 확인이 필요합니다.")
        return False

    law_list['최근발송일'] = law_list['최근발송일'].astype(str)
    law_list.loc[mask, '최근발송일'] = to_yyyymmdd(sent_date)
    for col in ['최근발송일', '최근개정일']:  ## 파일의 원래 형식(정수)으로 되돌림
        law_list[col] = law_list[col].astype(int)
    law_list.to_excel(law_file_name, index=False)
    return True


## 최종적으로 이메일 발송
def send_mail(mail_title, mail_body, file_list):
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
        for file in file_list:
            part = MIMEBase('application', "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(open(file, "rb").read())
            encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
            msg.attach(part)
        
        mail_server.sendmail(SEND_ADDR, msg['To'].split(','), msg.as_string())

    mail_server.quit()  ## SMTP 연결 정리