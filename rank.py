import datetime
import urllib.request
import json
import gspread  # pip install gspread
import re
import warnings
import ssl
import sys
import turtle
import pyautogui

ssl._create_default_https_context = ssl._create_unverified_context

# parameter 0:시트기록 x  / 1:시트기록 o
# is_write = 1
# turtle.setup(width=600, height=200, startx=0, starty=0)
# turtle.title("Tracing Writing Check")
# is_write = turtle.numinput(
#     title="시트에 기록하시겠습니까?", prompt="(1:기록/0:무기록)", default=1, minval=0, maxval=1
# )
# turtle.bye()
is_write = pyautogui.prompt("(1:기록/0:무기록)", "시트에 기록하시겠습니까?", default=1)

if is_write == None:
    sys.exit(1)

if is_write != "0" or is_write != "1":
    pyautogui.alert("1 또는 0을 입력하세요\n프로그램을 종료합니다")
    sys.exit(1)

is_write = int(is_write)

if is_write == 0:
    msg = "결과를 시트에 기록하지 않습니다"
else:
    msg = "결과를 시트에 기록합니다"
msg += "\n[OK]를 누르시고 잠시 기다려주세요"
pyautogui.alert(msg)

# Naver Search API id, secret key (cg-lab)
CLIENT_ID = "4L9CRB_dZ8HKe4R1WmNL"
CLIENT_SECRET = "HLoLebsBkd"

# google sheet API (inspi2k@gmail)
gs_json = "cglab-python-9750d891fb6e.json"
gs_url = "https://docs.google.com/spreadsheets/d/18O363rcAZY7bz6l7BAjCcekyx08Ytv1Ana1LLtO8Uhs"
gs_sheet_keyword = "list"
gs_sheet_rank = "rank"

# 1. 구글 시트에서 순위 조회할 리스트 가져오기
try:
    print("{} \tstart for google sheet reading".format(datetime.datetime.now()))

    gc = gspread.service_account(filename="./" + gs_json)
    doc = gc.open_by_url(gs_url)
    worksheet = doc.worksheet(gs_sheet_keyword)

    items = worksheet.get_all_records()
    # print(list_of_keyword)

    print("{} \tfinish for google sheet reading".format(datetime.datetime.now()))

except Exception as e:
    print("google sheet reading error:", e)

# sys.exit(0)
list_of_search = []


# Naver Search API
def get_nv_api(sstore, ccatalog_t, kkeyword):
    encText = urllib.parse.quote(kkeyword)

    pparam_start = 1
    list_return = []

    while pparam_start <= 1000:
        if pparam_start == 1:
            pparam_display = 99
        else:
            pparam_display = 100  # 찾아오는 아이템 단위 get_nv_api (max:100)
        url = "https://openapi.naver.com/v1/search/shop"
        url += (
            "?start={}".format(pparam_start)
            + "&display={}".format(pparam_display)
            + "&query="
            + encText
        )

        request = urllib.request.Request(url)
        request.add_header("X-Naver-Client-Id", CLIENT_ID)
        request.add_header("X-Naver-Client-Secret", CLIENT_SECRET)
        response = urllib.request.urlopen(request)
        rescode = response.getcode()
        if rescode == 200:
            response_body = response.read()
        else:
            print("Error Code:" + rescode)
            return []

        data = json.loads(response_body.decode("utf-8"))  # JSON 형태의 문자열 읽기

        for prd in data["items"]:
            if (sstore in prd["mallName"]) or (
                ccatalog_t != ""
                and ccatalog_t in re.sub("(<([^>]+)>)", "", prd["title"])
            ):
                # ffind_rank = 1
                list_return.append(
                    {
                        "storename": sstore,
                        "keyword": kkeyword,
                        "mid": prd["productId"],
                        "rank": pparam_start,
                    }
                )

                num_page = (pparam_start - 1) // 40 + 1
                num_ppos = (pparam_start - 1) % 40 + 1
                print(
                    "{} \tRank:{:>3} ({:>2}p {:>2}) {} / {} ({}) {}".format(
                        datetime.datetime.now(),
                        pparam_start,
                        num_page,
                        num_ppos,
                        kkeyword,
                        sstore,
                        prd["productId"],
                        re.sub("(<([^>]+)>)", "", prd["title"]),
                    ),
                    flush=True,
                )
                # pyautogui.confirm(kkeyword, pparam_start)
            pparam_start += 1

    return list_return


# google sheet - search, storename, keyword, catalog_title, startdate
# 1.1. 키워드만큼 반복하며 순위 조회 루틴 시작
try:
    for item in items:
        if item["search"] == "":
            continue

        # 2. 순위 찾기 루틴
        list_r = get_nv_api(item["storename"], item["catalog_title"], item["keyword"])

        if len(list_r) < 1:
            print(
                "Can't search '{}' in '{}'".format(item["keyword"], item["storename"])
            )

        for l in list_r:
            list_of_search.append(l)

except IndexError as e:
    print("Index Error:", e)
    print("item=", item)

except Exception as e:
    print("Error:", e)


# 3. 구글 시트에 기록
try:
    if is_write != 0:
        print("{} \tstart for google sheet writing".format(datetime.datetime.now()))

        wsheet = doc.worksheet(gs_sheet_rank)

        r = len(wsheet.col_values(1)) + 1

        list_values = []
        counter = 0

        for search in list_of_search:
            counter += 1
            list_row = []
            list_row.append(datetime.datetime.now().strftime("%Y-%m-%d"))
            list_row.append(datetime.datetime.now().strftime("%H:%M:%S"))
            list_row.append(1)
            # list_row.append(list_storename[key_i])
            # list_row.append(list_keyword[key_i])
            # list_row.append(list_rank[key_i])
            # list_row.append((int(list_rank[key_i]) - 1) // 40 + 1)
            list_row.append(search["storename"])
            list_row.append(search["keyword"])
            list_row.append(search["mid"])
            list_row.append(search["rank"])
            list_values.append(list_row)

        # print(list_values)
        warnings.filterwarnings(action="ignore")
        wsheet.update("A" + str(r) + ":G" + str(r + counter), list_values)
        warnings.filterwarnings(action="default")

        print("{} \tfinish for google sheet writing".format(datetime.datetime.now()))

        # pyautogui.confirm(list_values)
except Exception as e:
    print("writing sheet exception:", e)
