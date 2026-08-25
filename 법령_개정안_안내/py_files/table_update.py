import pandas as pd
import xlwings as xw

from bs4 import BeautifulSoup as bs
from .common_functions import *

## 목록 재수집 구간(일). 공포/발령 후 사이트 등록까지의 지연을 흡수하기 위한 폭.
## 실제 관측된 지연은 4일 수준이었으나, 비용이 낮으므로(law 9p + reg 21p) 넉넉히 잡는다.
LOOKBACK_DAYS = 60

## 해당 사이트에서, from_date보다 같거나 큰 날짜의 모든 리스트를 수집
def get_updated_table(browser, site_category, from_date):
    last_page_number = 200
    page_ranges = get_page_range(last_page_number)
    name_col, date_col = get_column_name(site_category)

    table_df_list = []

    find_continue = True
    for page_range in page_ranges:  # 각 페이지 범위들을 돌면서 데이터를 크롤링함

        for page in range(page_range[0], page_range[1] + 1):
            while True:
                try:
                    move_to_page(browser, page)  ## 해당 페이지로 이동해서
                    break
                except:
                    time.sleep(1)
            time.sleep(0.5)

            while True:  ## 이전 테이블과 동일하면. 크롤링에 랙이 발생한것이므로 다시
                page_source = bs(browser.page_source, 'html.parser')
                table_df = get_page_table_info(page_source)  ## 테이블 데이터 긁고
                if page == 1:
                    break
                if table_df_prev[name_col].equals(table_df[name_col]):
                    continue
                else:
                    break

            ## 더이상 업데이트할게 없으면 종료
            if pd.to_datetime(table_df[date_col]).max() < from_date:
                find_continue = False
                break
            table_df_list.append(table_df)
            table_df_prev = table_df.copy()
        if find_continue == False:
            break
        if page_range[1] == last_page_number:
            break
        else:
            click_next_page(browser)
            time.sleep(0.5)

    return pd.concat(table_df_list).reset_index(drop=True).drop('번호', axis=1)

############################################################################################
### 아래 코드는 모든 페이지를 돌면서 법령명과/공포일자/법령종류들 표 데이터를 크롤링하는 코드임
### 기본적으로 아래 테이블 정보를 기초로 실제 크롤링할 사이트를 결정함
### 평소에는 굳이 할일이 없고. 최초 개발시에만 DB원본을 만들기 위해 사용하였음
############################################################################################
def get_total_table_info(site_category):    
    name_col, date_col = get_column_name(site_category)
    browser = get_browser()
    move_to_home(browser, site_category=site_category)
    time.sleep(0.5)

    ## 마지막 페이지 위치 잡고. 페이지 범위 가져옴
    last_page_number = get_last_page_number(browser, site_category=site_category)
    page_ranges = get_page_range(last_page_number)

    table_df_list = []

    find_continue = True
    table_df_prev = None
    for page_range in page_ranges:  # 각 페이지 범위들을 돌면서 데이터를 크롤링함
        for page in range(page_range[0], page_range[1] + 1):
            while True:
                try:
                    move_to_page(browser, page)  ## 해당 페이지로 이동해서
                    break
                except:
                    time.sleep(1)
            time.sleep(0.5)

            while True:  ## 이전 테이블과 동일하면. 크롤링에 랙이 발생한것이므로 다시
                page_source = bs(browser.page_source, 'html.parser')
                table_df = get_page_table_info(page_source)  ## 테이블 데이터 긁고
                if page == 1: ## 맨 처음엔 table_df_prev가 존재하지 않으므로 넘어감
                    break
                if table_df_prev[name_col].equals(table_df[name_col]):
                    continue
                else:
                    break

            if table_df[date_col].max()[:4] < '2000':  ## 너무 옛날 데이터는 취합하지 않음
                find_continue = False
                break
            table_df_list.append(table_df)
            table_df_prev = table_df.copy()

        if find_continue == False:
            break
        if page_range[1] == last_page_number:
            break
        else:
            click_next_page(browser)
            time.sleep(0.5)

    table_df = pd.concat(table_df_list).reset_index(drop=True).drop('번호', axis=1).drop_duplicates()
    ## 행정규칙중 날짜가 잘못된게 딱 2줄 있음. 그걸 삭제
    table_df = table_df[~table_df[date_col].str.contains('20063')] 
    table_df = table_df[~table_df[date_col].str.contains('200611')]
    today = datetime.date.today().strftime('%y%m%d')
    table_df.to_excel(f"./database/{today}_lastest_{site_category}.xlsx", index=False)
    table_df.to_parquet(f"./database/{today}_lastest_{site_category}.parquet", index=False)

## 오래된 parquet 파일 삭제
def remove_old_parquet_files(max_file_num=20):
    parquet_files = glob.glob("./database/*.parquet")
    if len(parquet_files) > max_file_num:
        remove_files = list(f for f in parquet_files if f not in sorted(parquet_files, reverse=True)[:max_file_num])
        for file in remove_files:
            os.remove(file)

### 법령을 업데이트하고, 법령리스트 엑셀파일을 최신화함
def update_raw_excel():
    law_file = sorted(glob.glob("./database/*law.parquet"))[-1]
    reg_file = sorted(glob.glob("./database/*reg.parquet"))[-1]
    law_df = pd.read_parquet(law_file)
    reg_df = pd.read_parquet(reg_file).rename({
        '행정규칙명':'법령명', '발령일자':'공포일자'
    }, axis=1)

    df = pd.concat([law_df, reg_df])
    df['공포일자'] = pd.to_datetime(df['공포일자'])
    df['공포일자'] = df['공포일자'].dt.strftime("%Y%m%d")
    
    close_law_list_excel()
    law_file = './database/법규리스트.xlsx'    
    law_df = pd.read_excel(law_file)
    law_df['법령명'] = law_df['법령명'].map(lambda x: re.sub(r'[^가-힣]', '', x).strip())
    law_df['최근개정일'] = law_df['최근개정일'].astype(str)
    law_df['최근발송일'] = law_df['최근발송일'].astype(str)
    for idx, rows in law_df.iterrows():
        law_name, law_date = rows['법령명'], rows['최근개정일']
        if law_name in df.법령명.values.tolist():
            ## 법령일자가 최신화 안되어 있으면 최신화함
            latest_date = df[df.법령명==law_name].공포일자.max()
            if rows['최근개정일'] < latest_date:
                law_df.loc[idx, '최근개정일'] = latest_date
    law_df.to_excel(law_file, index=False)
    app = open_excel_app()
    wb = app.books.open(law_file)
    sht = wb.sheets[0]    
    sht.autofit()  # 열/행 자동 맞춤
    wb.save()
    wb.close()
    app.quit()

#####################################################
### 법령과 규칙 테이블 업데이트 필요하면 업데이트함
#####################################################
def table_update(lookback_days=LOOKBACK_DAYS):
    """법령/행정규칙 목록을 최근 lookback_days 일 구간만큼 다시 수집해 DB 에 병합한다.

    [왜 '오늘 - N일' 인가]
    예전에는 from_date 를 'DB 에 있는 가장 최신 공포일자'로 잡았다.
    그런데 국가법령정보센터는 공포/발령 후 며칠 지나서야 목록에 등록되는 경우가 있다.
    기준점이 DB 최신일이면 실행할 때마다 앞으로 밀리기 때문에,
    등록이 늦은 항목은 이미 지나간 구간에 나타나 영영 수집되지 않았다.
      실제 사례) 신용정보업감독규정 2026-08-13 발령 -> 8/17 경 등록 -> 끝내 누락
                 (같은 규정의 2026-07-24 개정도 같은 이유로 누락)
    기준점을 '오늘 - N일' 로 고정하면 창이 뒤로 밀리지 않으므로
    N일 이내에 등록되는 항목은 반드시 걸린다. 중복은 drop_duplicates() 가 정리한다.
    """
    browser = get_browser()
    time.sleep(0.2)

    remove_old_parquet_files(max_file_num=20) ## 오래된 parquet 파일은 삭제함
    update_yn = False  ## 해당일에 새로 들어온 항목이 있었는지 여부 체크
    today = datetime.date.today().strftime('%y%m%d')
    from_date = pd.Timestamp(datetime.date.today()) - pd.Timedelta(days=lookback_days)

    site_categories = ['law', 'reg']
    new_table_list = []
    for site_category in site_categories:
        file = sorted(glob.glob(f"./database/*{site_category}*.parquet"))[-1]
        table_raw = pd.read_parquet(file)

        move_to_home(browser, site_category=site_category)
        time.sleep(0.5)
        name_col, date_col = get_column_name(site_category)

        print(f"{site_category}_{from_date.strftime('%Y-%m-%d')} 이후 구간을 수집합니다. (최근 {lookback_days}일)")
        table_df = get_updated_table(browser, site_category, from_date)

        ## 수집분 중 기존 DB 에 없던 행만 골라냄. 창을 넓게 잡으므로 대부분은 이미 있는 행이다.
        new_rows = table_df.merge(table_raw, how='left', indicator=True)
        new_rows = new_rows[new_rows['_merge'] == 'left_only'].drop('_merge', axis=1)

        print(f"    수집 {len(table_df)}건 중 신규 {len(new_rows)}건")
        if len(new_rows) > 0:
            update_yn = True
            new_table_list.append(new_rows)

        ## 신규가 없어도 저장해 둔다(그날의 스냅샷을 남겨 추적이 가능하도록)
        table_raw = pd.concat([table_df, table_raw]).drop_duplicates().reset_index(drop=True)
        table_raw.to_parquet(f"./database/{today}_lastest_{site_category}.parquet")

    ## 최종적으로 법령 및 규칙 확인을 위해 다시 처음으로 이동
    move_to_home(browser, site_categories[0])
    url = get_url(site_categories[1])
    browser.execute_script(f"window.open('{url}', '_blank');")

    if update_yn: ## 변경사항이 있다면 엑셀을 띄움
        for new_table in new_table_list:
            if excel_is_visible():  ## 자동 실행 중에는 창을 띄우지 않음
                xw.view(new_table, table=False)

        print('데이터 업데이트 완료. 엑셀과 실제 웹페이지를 비교해 오류가 없는지 확인해주세요.')

        ## 기존 엑셀 법규리스트 파일을 업데이트함
        update_raw_excel()
    else:
        print(f'({today})_새로 등록된 법령/행정규칙이 없습니다.')
