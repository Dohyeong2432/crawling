# -*- coding: utf-8 -*-
"""법령 제·개정 점검 실행 스크립트 (메일 발송 없음).

실행 방법
    cd C:\\Windows\\Python\\notebooks\\법령_개정안_안내
    python run_check.py                 ## 기본: 최근 60일 구간 재수집
    python run_check.py 480             ## 소급 수집이 필요할 때 일수를 직접 지정
    python run_check.py --no-crawl      ## 크롤링 없이 현재 DB 로 판정만

이 스크립트가 하는 일
    1) table_update()   국가법령정보센터 목록을 재수집해 DB(parquet) 갱신
    2) get_update_df()  법규리스트와 대조해 안내가 필요한 법령을 추림
    3) 결과를 표로 출력

하지 않는 일
    - 메일을 보내지 않는다. send_mail() 을 호출하는 코드가 이 파일에 없다.
      발송은 법령제개정수집자동화.ipynb 마지막 셀에서 사람이 내용을 확인한 뒤 실행한다.
    - Excel 창을 띄우지 않는다(set_excel_visible(False)).
      서식 저장은 그대로 수행되고 화면에만 안 뜬다.

참고
    Chrome 창은 뜬다. 대상 사이트가 클릭·다운로드 기반이라 headless 로 바꿀 수 없다.
    따라서 로그인된 데스크톱 세션에서만 동작한다.
"""
import os
import sys
from pathlib import Path

## 내부 코드가 "./database/..." 같은 상대경로를 쓰므로 프로젝트 폴더를 기준으로 맞춘다.
ROOT = Path(__file__).resolve().parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))

import pandas as pd

from py_files.common_functions import set_excel_visible, close_law_list_excel
from py_files.table_update import table_update, LOOKBACK_DAYS
from py_files.get_information import get_update_df


def parse_args(argv):
    lookback, crawl = LOOKBACK_DAYS, True
    for arg in argv[1:]:
        if arg == "--no-crawl":
            crawl = False
        elif arg.isdigit():
            lookback = int(arg)
        else:
            sys.exit(f"알 수 없는 인자입니다: {arg}\n{__doc__}")
    return lookback, crawl


def main(argv):
    lookback, crawl = parse_args(argv)
    set_excel_visible(False)  ## 자동 실행 중 Excel 창이 뜨지 않도록

    print("=" * 72)
    print("법령 제·개정 점검  (메일 발송 없음)")
    print("=" * 72)

    if crawl:
        print(f"\n[1/2] 목록 수집 — 최근 {lookback}일 구간")
        new_tables = table_update(lookback_days=lookback)
        ## 새로 등록된 항목을 화면에 남긴다.
        ## 대부분은 감시 대상 163건과 무관하지만, 무엇이 들어왔는지는 보여야 한다.
        ## (Excel 창을 끄고 실행하므로 노트북의 xw.view 를 이 출력이 대신한다)
        for new_table in new_tables or []:
            print(f"\n  [신규 등록 {len(new_table)}건]")
            with pd.option_context("display.max_colwidth", 45, "display.width", 200):
                print(new_table.head(30).to_string(index=False))
            if len(new_table) > 30:
                print(f"  ... 외 {len(new_table) - 30}건")
    else:
        print("\n[1/2] 목록 수집 건너뜀 (--no-crawl)")

    print("\n[2/2] 법규리스트 대조")
    ## 법규리스트가 Excel 에서 열려 있으면 저장이 실패하므로 먼저 닫는다.
    ## (저장하지 않고 닫으므로, 직접 편집 중이던 내용이 있으면 사라진다)
    close_law_list_excel()
    update_df = get_update_df()

    print("\n" + "=" * 72)
    if len(update_df) == 0:
        print("안내가 필요한 법령/행정규칙이 없습니다.")
        print("=" * 72)
        return 0

    update_df = update_df.sort_values("최근개정일", ascending=False)
    print(f"안내가 필요한 법령/행정규칙 {len(update_df)}건")
    print("=" * 72)
    with pd.option_context("display.max_colwidth", 50, "display.width", 200):
        print(update_df.to_string(index=False))
    print("=" * 72)
    print("발송하려면 법령제개정수집자동화.ipynb 를 열어 내용을 확인한 뒤")
    print("마지막 셀(이메일 발송)을 실행하세요.")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
