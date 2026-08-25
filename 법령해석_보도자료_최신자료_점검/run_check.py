# -*- coding: utf-8 -*-
"""법령해석 · 보도자료 신규 자료 점검 스크립트 (메일 발송 없음).

실행 방법
    cd C:\\Windows\\Python\\notebooks\\법령해석_보도자료_최신자료_점검
    python run_check.py 20260820        ## 이 날짜 이후 등록된 자료만 확인
    python run_check.py 260820          ## 6자리로 줘도 된다

하는 일
    아래 4개 사이트에 기준일 이후 새로 올라온 자료가 있는지만 확인해 표로 보여준다.
      - 금융위 법령해석   (fss_itp)
      - 금융위 보도자료   (fsc_press)
      - 금감원 보도자료   (fss_press)
      - 금융정보분석원 보도자료 (fiu_press)

하지 않는 일
    - 메일을 보내지 않는다. send_mail() 을 호출하는 코드가 이 파일에 없다.
      발송은 법령해석_보도자료_최신자료_점검.ipynb 에서 사이트별로 확인한 뒤 실행한다.
    - 첨부파일을 내려받지 않는다. 점검 단계는 목록만 읽는다.

한 사이트가 실패해도 나머지는 계속 진행한다.
사이트 구조가 바뀌면 그 사이트만 오류로 표시되므로, 조용히 지나가지 않는다.

참고
    Chrome 창은 뜬다. 대상 사이트가 클릭·다운로드 기반이라 headless 로 바꿀 수 없다.
"""
import os
import re
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))

from py_files.common_functions import get_browser
import py_files.fss_itp as fss_itp
import py_files.fsc_press as fsc_press
import py_files.fss_press as fss_press
import py_files.fiu_press as fiu_press

## (표시 이름, 점검 함수). 노트북의 점검 셀과 같은 순서.
CHECKS = [
    ("금융위 법령해석", fss_itp.update_check),
    ("금융위 보도자료", fsc_press.update_check),
    ("금감원 보도자료", fss_press.update_check),
    ("금융정보분석원 보도자료", fiu_press.update_check),
]


def parse_today(argv):
    """20260820 / 260820 을 모두 받아 6자리 'yymmdd' 로 맞춘다.

    각 사이트의 등록일을 '%y%m%d' 로 변환해 문자열 비교하므로 6자리여야 한다.
    """
    if len(argv) < 2:
        sys.exit("기준일을 넣어주세요.  예)  python run_check.py 20260820")
    raw = re.sub(r"[^0-9]", "", argv[1])
    if len(raw) == 8:       ## 20260820 -> 260820
        raw = raw[2:]
    if len(raw) != 6:
        sys.exit(f"날짜 형식을 알 수 없습니다: {argv[1]}  (예: 20260820)")
    return raw


def main(argv):
    today = parse_today(argv)

    print("=" * 72)
    print("법령해석 · 보도자료 신규 자료 점검  (메일 발송 없음)")
    print("=" * 72)
    print(f"  기준일 : 20{today[:2]}-{today[2:4]}-{today[4:]} 이후 등록분\n")

    started = time.time()
    browser = get_browser()
    browser.set_page_load_timeout(60)

    failed = []
    try:
        for label, check in CHECKS:
            print(f"--- {label} ---")
            begin = time.time()
            try:
                check(browser, today)
                print(f"    ({time.time() - begin:.1f}초)\n")
            except Exception:
                failed.append(label)
                print(f"    [오류] 점검에 실패했습니다 ({time.time() - begin:.1f}초)")
                traceback.print_exc(limit=3)
                print()
    finally:
        browser.quit()

    print("=" * 72)
    print(f"점검 완료 — 성공 {len(CHECKS) - len(failed)}/{len(CHECKS)} ({time.time() - started:.0f}초)")
    if failed:
        print(f"실패 : {', '.join(failed)}")
        print("사이트 구조가 바뀌었을 수 있습니다. 위 오류 내용을 확인하세요.")
    print("발송하려면 법령해석_보도자료_최신자료_점검.ipynb 에서")
    print("해당 사이트의 발송 셀을 실행하세요.")
    print("=" * 72)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
