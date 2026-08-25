# [사용하지 않음] 금융위 법령해석 최신사례

**2026-08-25부로 이 프로젝트는 사용하지 않습니다.**

## 이유

`법령해석_보도자료_최신자료_점검` 프로젝트의 `py_files/fss_itp.py` 모듈이
**완전히 같은 사이트**를 수집합니다.

```
이 프로젝트          → better.fsc.go.kr/fsc_new/replyCase/PastReplyList.do?stNo=11&muNo=171&muGpNo=75
fss_itp.py          → better.fsc.go.kr/fsc_new/replyCase/PastReplyList.do?stNo=11&muNo=171&muGpNo=75
```

URL이 동일하고, `fss_itp.py` 쪽이 더 발전된 버전입니다.

- 첨부 엑셀에 서식 적용(자동 줄바꿈, 열 너비, 상단 정렬)
- 제목을 첫 번째 표에서 가져옴 (두 번째 표에 제목이 없는 게시글 대응)
- 보도자료 3종과 함께 한 번에 점검/발송

## 앞으로 법령해석은 여기서 확인하세요

```
notebooks/법령해석_보도자료_최신자료_점검/법령해석_보도자료_최신자료_점검.ipynb
```

## 파일은 그대로 두었습니다

폴더와 `update_list/`의 과거 수집 기록(11건)을 포함해 **아무것도 삭제하지 않았습니다.**
다시 쓰실 일이 생기면 노트북을 열어 그대로 실행하시면 됩니다.

`notebooks/.env` 의 `RECV_ADDRS_FSC_INTERPRETATION` 항목도 지우지 않고 남겨두었으므로
설정 변경 없이 바로 동작합니다.

## 참고: 수신자가 달랐습니다

| | 수신자 |
|---|---|
| 이 프로젝트 (`RECV_ADDRS_FSC_INTERPRETATION`) | h09144, h09087, h09153 |
| `fss_itp` (`RECV_ADDRS_PRESS`) | h09144 |

이 프로젝트를 사용하지 않으면 **h09087, h09153 두 분은 법령해석 안내를 받지 못합니다.**
필요하면 `notebooks/.env` 의 `RECV_ADDRS_PRESS` 에 추가하시면 되는데,
그 경우 보도자료 3종(금융위/금감원/FIU)도 함께 발송된다는 점을 감안하세요.
