# -*- coding: utf-8 -*-
"""금융감독원 제재사항 점검 스크립트 (메일 발송 없음).

실행 방법
    cd C:\\Windows\\Python\\notebooks\\금감원_제재내역_안내
    python run_check.py 20260720        ## 이 날짜 이후의 제재만 검토
    python run_check.py 260720          ## 6자리로 줘도 된다

하는 일
    제재 목록을 최신 페이지부터 거슬러 올라가며, 제재조치요구일이 기준일 이상인 건에 대해
      - 관련부서가 FIND_CATEGORIES 에 걸리는지
      - 첨부 PDF 본문에 FIND_PDF_KEYWORDS 가 있는지
    를 확인해 안내 대상을 추린다. PDF 는 다운로드 폴더에 저장된다.

하지 않는 일
    - 메일을 보내지 않는다. send_mail() 을 호출하는 코드가 이 파일에 없다.
      발송은 감독원제재수집자동화.ipynb 마지막 셀에서 내용을 확인한 뒤 실행한다.
    - Excel 창을 띄우지 않는다.

검색 조건
    py_files/common_functions.py 의 FIND_CATEGORIES / FIND_PDF_KEYWORDS 에서 관리한다.
    노트북도 같은 값을 쓰므로 조건 변경은 그 파일 한 곳만 고치면 된다.

참고
    Chrome 창은 뜬다. 대상 사이트가 클릭·다운로드 기반이라 headless 로 바꿀 수 없다.
"""
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))

import pandas as pd

from py_files.common_functions import (
    FIND_CATEGORIES,
    FIND_PDF_KEYWORDS,
    DOWNLOAD_FOLDER,
    get_notify_list,
)


def parse_min_date(argv):
    """20260720 / 260720 을 모두 받아 8자리 'YYYYMMDD' 로 맞춘다.

    사이트의 제재조치요구일이 8자리 문자열이라 그 형식에 맞춰야 비교가 된다.
    """
    if len(argv) < 2:
        sys.exit("기준일을 넣어주세요.  예)  python run_check.py 20260720")
    raw = re.sub(r"[^0-9]", "", argv[1])
    if len(raw) == 6:       ## 260720 -> 20260720
        raw = "20" + raw
    if len(raw) != 8:
        sys.exit(f"날짜 형식을 알 수 없습니다: {argv[1]}  (예: 20260720)")
    return raw


def main(argv):
    min_date = parse_min_date(argv)

    print("=" * 72)
    print("금융감독원 제재사항 점검  (메일 발송 없음)")
    print("=" * 72)
    print(f"  기준일     : {min_date} 이후")
    print(f"  관련부서   : {', '.join(FIND_CATEGORIES)}")
    print(f"  PDF 키워드 : {', '.join(FIND_PDF_KEYWORDS)}")
    print(f"  다운로드   : {DOWNLOAD_FOLDER}")
    print("\n제재 목록을 확인하고 첨부 PDF 를 읽는 중입니다. 시간이 걸릴 수 있습니다...\n")

    update_df = get_notify_list(min_date)

    print("\n" + "=" * 72)
    if update_df is None or len(update_df) == 0:
        print("안내가 필요한 제재 내용이 없습니다.")
        print("=" * 72)
        return 0

    show = update_df.copy()
    if "pdf_files" in show.columns:
        show["pdf_files"] = show["pdf_files"].map(os.path.basename)
    print(f"안내가 필요한 제재 {len(show)}건")
    print("=" * 72)
    with pd.option_context("display.max_colwidth", 40, "display.width", 220):
        print(show.to_string(index=False))
    print("=" * 72)
    print("발송하려면 감독원제재수집자동화.ipynb 를 열어 내용을 확인한 뒤")
    print("마지막 셀(이메일 발송)을 실행하세요.")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
