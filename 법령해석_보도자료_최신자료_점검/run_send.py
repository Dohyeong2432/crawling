# -*- coding: utf-8 -*-
"""법령해석 · 보도자료 이메일 발송 스크립트.

**이 스크립트는 실제로 메일을 보냅니다.** 되돌릴 수 없습니다.
그래서 기본 동작은 "보낼 목록만 보여주기"이고, 실제 발송은 --yes 를 붙여야 합니다.

실행 방법
    cd C:\\Windows\\Python\\notebooks\\법령해석_보도자료_최신자료_점검

    ## 1) 무엇이 나갈지 먼저 확인 (발송 안 함)
    python run_send.py 260825 --site fss_press
    python run_send.py 260825 --site fss_press --title "공모펀드"

    ## 2) 확인했으면 --yes 를 붙여 실제 발송
    python run_send.py 260825 --site fss_press --title "공모펀드" --yes

인자
    <날짜>            20260825 / 260825 둘 다 가능. 이 날짜 이후 등록분이 대상.
    --site <이름>     fss_itp | fsc_press | fss_press | fiu_press | all
    --title "..."     제목에 이 문자열이 들어간 건만 발송.
                      콤마로 여러 개 지정 가능(하나라도 맞으면 대상).
                      공백은 무시하고 비교하므로 사이트에서 제목을 복사해 붙여도 된다.
                      생략하면 해당 사이트의 그 날짜 이후 자료 전부.
    --yes             실제 발송. 없으면 목록만 출력한다.

미리보기는 각 사이트 1페이지만 훑는다(보통 그 안에 다 들어온다).
실제 발송은 페이지를 넘겨가며 조건에 맞는 건을 모두 처리한다.
"""
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))

from py_files.common_functions import get_browser, move_to_home, subject_matches, RECV_ADDRS
import py_files.fss_itp as fss_itp
import py_files.fsc_press as fsc_press
import py_files.fss_press as fss_press
import py_files.fiu_press as fiu_press

## 사이트 이름 -> (표시명, 목록 조회 함수, 발송 함수, move_to_home 용 키)
SITES = {
    "fss_itp":   ("금융위 법령해석",        fss_itp.get_table_data,   fss_itp.notice_fss_ipt,     "fss_itp"),
    "fsc_press": ("금융위 보도자료",        fsc_press.get_table_data, fsc_press.notice_fsc_press, "fsc_press"),
    "fss_press": ("금감원 보도자료",        fss_press.get_table_data, fss_press.notice_fsc_press, "fss_press"),
    "fiu_press": ("금융정보분석원 보도자료", fiu_press.get_table_data, fiu_press.notice_fiu_press, "fiu_press"),
}


def parse_args(argv):
    date, sites, keywords, confirmed = None, None, None, False
    i = 1
    while i < len(argv):
        arg = argv[i]
        if arg == "--yes":
            confirmed = True
        elif arg == "--site":
            i += 1
            name = argv[i] if i < len(argv) else ""
            if name == "all":
                sites = list(SITES)
            elif name in SITES:
                sites = [name]
            else:
                sys.exit(f"알 수 없는 사이트: {name}\n선택 가능: {', '.join(SITES)}, all")
        elif arg == "--title":
            i += 1
            raw = argv[i] if i < len(argv) else ""
            keywords = [k.strip() for k in raw.split(",") if k.strip()]
        elif re.fullmatch(r"[0-9]{6,8}", re.sub(r"[^0-9]", "", arg)):
            digits = re.sub(r"[^0-9]", "", arg)
            date = digits[2:] if len(digits) == 8 else digits
        else:
            sys.exit(f"알 수 없는 인자: {arg}\n{__doc__}")
        i += 1

    if date is None:
        sys.exit("기준일을 넣어주세요.  예)  python run_send.py 260825 --site fss_press")
    if sites is None:
        sys.exit("--site 를 지정해주세요.  선택 가능: " + ", ".join(SITES) + ", all")
    return date, sites, keywords, confirmed


def preview(browser, name, today, keywords):
    """1페이지를 훑어 발송 대상이 될 항목을 보여준다."""
    label, get_table, _, home_key = SITES[name]
    move_to_home(browser, home_key)
    result = get_table(browser)
    table_df = result[0] if isinstance(result, tuple) else result
    table_df = table_df[table_df.등록일 >= today]
    if keywords:
        table_df = table_df[table_df["제목"].map(lambda s: subject_matches(s, keywords))]
    print(f"--- {label} : {len(table_df)}건 ---")
    for _, row in table_df.iterrows():
        print(f"    [{row['등록일']}] {row['제목']}")
    if len(table_df) == 0:
        print("    (조건에 맞는 자료 없음)")
    return len(table_df)


def main(argv):
    today, sites, keywords, confirmed = parse_args(argv)

    print("=" * 72)
    print("법령해석 · 보도자료 발송" + ("" if confirmed else "  [미리보기 — 발송하지 않음]"))
    print("=" * 72)
    print(f"  기준일   : 20{today[:2]}-{today[2:4]}-{today[4:]} 이후 등록분")
    print(f"  대상     : {', '.join(SITES[s][0] for s in sites)}")
    print(f"  제목조건 : {' 또는 '.join(keywords) if keywords else '(전체)'}")
    print(f"  수신자   : {', '.join(RECV_ADDRS)}")
    print()

    browser = get_browser()
    browser.set_page_load_timeout(60)
    try:
        total = 0
        for name in sites:
            total += preview(browser, name, today, keywords)

        print()
        if total == 0:
            print("=" * 72)
            print("조건에 맞는 자료가 없어 발송할 것이 없습니다.")
            print("=" * 72)
            return 0

        if not confirmed:
            print("=" * 72)
            print(f"위 {total}건이 발송 대상입니다. 아직 보내지 않았습니다.")
            print("실제로 보내려면 같은 명령에 --yes 를 붙여 다시 실행하세요.")
            print("=" * 72)
            return 0

        print("=" * 72)
        print("발송을 시작합니다.")
        print("=" * 72)
        for name in sites:
            label, _, notice, _ = SITES[name]
            print(f"\n--- {label} ---")
            notice(browser, today, keywords)
    finally:
        browser.quit()

    print("\n" + "=" * 72)
    print("발송 완료.")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
