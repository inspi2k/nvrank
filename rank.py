import datetime
import urllib.request
import json
import gspread  # pip install gspread
import re
import warnings
import ssl
import sys
import time
import pyautogui  # pip install pyautogui

ssl._create_default_https_context = ssl._create_unverified_context

# parameter 0:시트기록 x  / 1:시트기록 o
is_write = pyautogui.prompt("(1:기록/0:무기록)", "시트에 기록하시겠습니까?", default=1)

if is_write == None:
    sys.exit(1)

# if is_write != "0" or is_write != "1":
#     pyautogui.alert("1 또는 0을 입력하세요\n프로그램을 종료합니다")
#     sys.exit(1)

is_write = int(is_write)

if is_write == 0:
    msg = "결과를 시트에 기록하지 않습니다"
elif is_write == 1:
    msg = "결과를 시트에 기록합니다"
else:
    pyautogui.alert("1 또는 0을 입력하세요")
    sys.exit(1)
# msg += "\n[OK]를 누르시고 잠시 기다려주세요"
# pyautogui.alert(msg)

is_mid = pyautogui.prompt(
    "(1:정해진 상품(MID)만 순위 검색 / 0:스토어 내 모든 상품 순위 검색)", "어떤 상품의 순위를 찾습니까?", default=1
)

if is_mid == None:
    sys.exit(1)

is_mid = int(is_mid)

if is_mid == 1:
    msg += "\n\n정해진 상품만 순위 검색합니다 (MID)"
elif is_mid == 0:
    msg += "\n\n스토어의 모든 상품의 순위 검색합니다"
else:
    pyautogui.alert("1 또는 0을 입력하세요")
    sys.exit(1)
msg += "\n\n[OK]를 누르시면 순위 검색을 시작합니다 ([Cancel]은 중지)"
if pyautogui.confirm(msg) == "Cancel":
    sys.exit(0)

# Naver Search API id, secret key (cg-lab)
CLIENT_ID = "4L9CRB_dZ8HKe4R1WmNL"
CLIENT_SECRET = "HLoLebsBkd"

# google sheet API (inspi2k@gmail)
gs_json = "cglab-python-9750d891fb6e.json"
gs_url = "https://docs.google.com/spreadsheets/d/18O363rcAZY7bz6l7BAjCcekyx08Ytv1Ana1LLtO8Uhs"
gs_sheet_keyword = "list"
gs_sheet_rank = "rank"


# Naver Search API - https://developers.naver.com
# API Call Limit - 1초에 10회
def get_nv_api(sstore, kkeyword, ccatalog_t, mmid):
    # print("{} - {} - {} - {}".format(sstore, kkeyword, ccatalog_t, mmid))

    encText = urllib.parse.quote(kkeyword)

    param_start = 1
    list_return = []

    while param_start <= 1000:
        if param_start == 1:
            param_display = 99
        else:
            param_display = 100  # 찾아오는 아이템 단위 get_nv_api (max:100)
        
        if (param_start % 5) == 0:
            time.sleep(1)
            # print("{}\t sleep / {:>4},{} / {}".format(datetime.datetime.now(), param_start, param_display, kkeyword))

        url = "https://openapi.naver.com/v1/search/shop"
        url += (
            "?start={}".format(param_start)
            + "&display={}".format(param_display)
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

        if mmid != "":  # MID로 상품 순위 찾기
            # print("keyword={} \tmid={}".format(kkeyword, mmid))
            for prd in data["items"]:
                # print(
                #     "productId={} \tmallName={}".format(
                #         prd["productId"], prd["mallName"]
                #     )
                # )
                if mmid == prd["productId"]:
                    # print(
                    #     "I got it! {},{},{},{}".format(
                    #         sstore, kkeyword, mmid, param_start
                    #     )
                    # )
                    list_return.append(
                        {
                            "storename": sstore,
                            "keyword": kkeyword,
                            "mid": mid,
                            "rank": param_start,
                            "title": re.sub("(<([^>]+)>)", "", prd["title"]),
                        }
                    )
                    return list_return

                # pyautogui.confirm(kkeyword, pparam_start)
                param_start += 1
        elif mmid == "":  # 모든 상품 순위 찾기
            # print("keyword={}".format(kkeyword))
            for prd in data["items"]:
                # 상품 검색된 상황
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
                            "rank": param_start,
                            "title": re.sub("(<([^>]+)>)", "", prd["title"]),
                        }
                    )

                # pyautogui.confirm(kkeyword, pparam_start)
                param_start += 1

        else:
            print("No searching - no mid")
            return []

        # # console log
        # num_page = (param_start - 1) // 40 + 1
        # num_ppos = (param_start - 1) % 40 + 1
        # print(
        #     "{} \tRank:{:>3} ({:>2}p {:>2}) {} / {} ({}) {}".format(
        #         datetime.datetime.now(),
        #         param_start,
        #         num_page,
        #         num_ppos,
        #         kkeyword,
        #         sstore,
        #         mmid,
        #         re.sub("(<([^>]+)>)", "", prd["title"]),
        #     ),
        #     flush=True,
        # )

    return list_return


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

# google sheet - search, storename, keyword, catalog_title, startdate
# 1.1. 키워드만큼 반복하며 순위 조회 루틴 시작 - get_nv_api()
try:
    for item in items:
        if item["search"] == "":
            continue

        # 2. 순위 찾기 루틴
        if is_mid == 1:  # MID로 검색
            mid = (
                item["ctMid"]
                if item["nvMid"] != "" and item["ctMid"] != ""
                else item["nvMid"]
            )
            mid = str(mid)
        else:  # 모든 상품 검색
            mid = ""

        # print(
        #     "{}-{}-{}-{}".format(
        #         item["storename"], item["keyword"], item["catalog_title"], mid
        #     )
        # )
        list_r = get_nv_api(
            item["storename"], item["keyword"], item["catalog_title"], mid
        )

        if len(list_r) < 1:
            print(
                "Can't search '{}' in '{}'(~1,099 ranks / 28p 19th)".format(
                    item["keyword"], item["storename"]
                )
            )

        for l in list_r:
            list_of_search.append(l)

            # console log
            num_page = (l["rank"] - 1) // 40 + 1
            num_ppos = (l["rank"] - 1) % 40 + 1

            print(
                "{}\tRank:{:>4} ({:>2}p {:>2}) {} / {} ({}) {}".format(
                    datetime.datetime.now(),
                    format(l["rank"], ","),
                    num_page,
                    num_ppos,
                    l["keyword"],
                    l["storename"],
                    l["mid"],
                    l["title"][:30] + "..." if len(l["title"]) > 30 else l["title"],
                )
            )

except IndexError as e:
    print("Index Error:", e)
    print("item=", item)

except Exception as e:
    print("Error:", e)

# 3. 구글 시트에 기록
if is_write != 0:
    try:
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
            list_row.append(search["storename"])
            list_row.append(search["keyword"])
            list_row.append(search["title"])
            list_row.append(search["mid"])
            list_row.append(search["rank"])
            list_values.append(list_row)

        # print(list_values)
        warnings.filterwarnings(action="ignore")
        wsheet.update("A" + str(r) + ":H" + str(r + counter), list_values)
        warnings.filterwarnings(action="default")

        print("{} \tfinish for google sheet writing".format(datetime.datetime.now()))

        # pyautogui.confirm(list_values)
    except Exception as e:
        print("writing sheet exception:", e)

# 4. traffic sheet 읽어와서 비교 - 업데이트
# API Call Limit - 1유저 당 60초간 60회
if is_write != 0:
    try:
        print("{} \t-- update start".format(datetime.datetime.now()))

        tr_doc = gc.open_by_key('1D3xNE6orWM4gPSXunAYaf1ohteqJMSEcTkzP5wejfoM')
        # tr_doc = gc.open_by_url('https://docs.google.com/spreadsheets/d/1D3xNE6orWM4gPSXunAYaf1ohteqJMSEcTkzP5wejfoM')
        tr_wsheet = tr_doc.worksheet('keywords')

        # 직전의 page 값 복사 - 비교하기 위해
        prev_page = tr_wsheet.col_values(10)
        pprev = []
        nn = [datetime.datetime.now().strftime("%Y-%m-%d %H:%M")]
        pprev.append(nn)
        for p in range(1,len(prev_page)):
            nn = [int(prev_page[p])]
            pprev.append(nn)
        # print(pprev)
        warnings.filterwarnings(action="ignore")
        tr_wsheet.update('P1:P'+str(len(prev_page)),pprev)
        warnings.filterwarnings(action="default")

        tr_items = tr_wsheet.get_all_records()

        for search in list_of_search:
            for tr_i in tr_items:
                if (search["mid"] == str(tr_i["nvMid"])) and (search["keyword"] == tr_i["keyword"]):
                    cell = tr_wsheet.find(search["mid"])

                    warnings.filterwarnings(action="ignore")
                    tr_wsheet.update_cell(cell.row, cell.col+2, (search["rank"]-1)//40+1)
                    warnings.filterwarnings(action="default")

                    print('\tupdate values - r{}c{}={}({}p)\t{}({})'.format(cell.row, cell.col+2, search["rank"], (search["rank"]-1)//40+1,search["keyword"], search["mid"]))

                    continue
                elif (search["mid"] == str(tr_i["ctMid"])) and (search["keyword"] == tr_i["keyword"]):
                    cell = tr_wsheet.find(search["mid"])

                    warnings.filterwarnings(action="ignore")
                    tr_wsheet.update_cell(cell.row, cell.col+1, (search["rank"]-1)//40+1)
                    warnings.filterwarnings(action="default")

                    print('\tupdate values - r{}c{}={}({}p)\t{}({})'.format(cell.row, cell.col+1, search["rank"], (search["rank"]-1)//40+1,search["keyword"], search["mid"]))

                    continue

        print("{} \t-- update finish".format(datetime.datetime.now()))
        
    except Exception as e:
        print("update error:", e)