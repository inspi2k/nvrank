# 230919 화 오후5
# git 연동

# 230905 화 오후2
# 검색페이지 api의 json은 권한 오류로 사용불가
# 대체 - catalog 상품은 제목으로 찾기
# 구글 시트에 기록

# 230904 월
# 네이버 검색 API 이용하여 순위 찾기
# - 28페이지 29위까지 검색가능 (1000위부터 100개까지)
# 구글 시트에서 읽기

import datetime
import urllib.request
import json
import gspread  # pip install gspread
import re
import warnings
import ssl

ssl._create_default_https_context = ssl._create_unverified_context

# Variable list
param_display = 100     # 찾아오는 아이템 단위 get_nv_api (max:100)

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

    gc = gspread.service_account(filename="./"+gs_json)
    doc = gc.open_by_url(gs_url)
    worksheet = doc.worksheet(gs_sheet_keyword)

    list_check = worksheet.col_values(1)
    list_storename = worksheet.col_values(2)
    list_keyword = worksheet.col_values(3)
    list_catalog_title = worksheet.col_values(4)
    list_rank = ['rank']

    print("{} \tfinish for google sheet reading".format(datetime.datetime.now()))
except Exception as e:
    print("google sheet reading error:", e)

# print(list_check)
# print(list_storename)
# print(list_keyword)
# print(list_catalog_title)
# print(list_rank)

# Naver Search API
def get_nv_api(sstore, ccatalog_t, kkeyword, pparam_start=1, pparam_display=100, rrank=0, ffind_rank=0):
    encText = urllib.parse.quote(kkeyword)
    url = "https://openapi.naver.com/v1/search/shop"
    if(pparam_start==1): pparam_display -= 1
    url += "?start={}".format(pparam_start) + "&display={}".format(pparam_display) + "&query=" + encText
    # print("s:{}/d:{}/k:{}".format(pparam_start,pparam_display,kkeyword))
    if(pparam_start==1): pparam_display += 1
    
    request = urllib.request.Request(url)
    request.add_header("X-Naver-Client-Id", CLIENT_ID)
    request.add_header("X-Naver-Client-Secret", CLIENT_SECRET)
    response = urllib.request.urlopen(request)
    rescode = response.getcode()
    if(rescode==200):
        response_body = response.read()
    else:
        print("Error Code:" + rescode)
        return -1, -1

    data = json.loads(response_body.decode('utf-8')) # JSON 형태의 문자열 읽기

    for item in data['items']:
        rrank += 1
        # print("productType={}, cat_title={}, title={}".format(item['productType'],ccatalog_t,re.sub('(<([^>]+)>)','',item['title'])))
        if( (sstore in item['mallName']) or (ccatalog_t !="" and ccatalog_t in re.sub('(<([^>]+)>)','',item['title'])) ):
            ffind_rank = 1
            break

    return rrank, ffind_rank

# 1.1. 키워드만큼 반복하며 순위 조회 루틴 시작
try:
    for key_i in range(1,len(list_check)):
#        if(key_i > len(list_check)):
#            list_check.insert(key_i,"")
#            list_rank.insert(key_i, "")
#            continue

        if(list_check[key_i]==""):
            list_keyword[key_i] = ""
        
        keyword = list_keyword[key_i]
        if(keyword==""):
            list_rank.insert(key_i, "")
            continue

        storename = list_storename[key_i]
        if key_i < len(list_catalog_title): catalog_title = list_catalog_title[key_i]
        else: catalog_title = ""

        rank = 0
        find_rank = 0

        param_start = 1
        
        # print("{} \tindex={} \tstarting for keyword:{}(catalog:{})".format(datetime.datetime.now(), key_i, keyword, catalog))

        # 2. 순위 찾기 루틴
        # catalog 가격비교상품 / 일반상품
        while find_rank==0 :
            if (param_start > 1000): break
            
            # print("find_rank:{}\tget_nv_api(start:{},display:{},keyword:{})".format(find_rank,param_start,param_display,keyword))
            (rank, find_rank) = get_nv_api(storename, catalog_title, keyword, param_start, param_display, rank, find_rank)

            if(find_rank==1): 
                list_rank.insert(key_i, rank)
                # print("found it in keyword({})".format(keyword))

            # if (param_start % 5==0): print(".", end="")

            if(param_start==1): param_start = param_display
            else: param_start += param_display

        # print("key_i={} find_rank={}, rank={}, keyword={}".format(key_i, find_rank, rank, keyword))
        if(find_rank==0 and rank>1): 
            list_rank.insert(key_i, 0)
            rank = ""
            # print("not found it in keyword({})".format(keyword))

        if( rank == "" ):
            num_page = ""
            num_ppos = ""
        else:
            num_page = (rank-1) // 40 + 1
            num_ppos = (rank-1) % 40 + 1
        print("{} \tRank:{:>3} ({:>2}p {:>2}) {} / {}".format(datetime.datetime.now(),rank,num_page,num_ppos,keyword,storename))
except IndexError as e:
    print("Index Error:",e)
    print("key_i=",key_i)
    
    print(list_check)
    print(list_storename)
    print(list_keyword)
    print(list_catalog_title)
    print(list_rank)

except Exception as e:
    print("Error:",e)



# 3. 구글 시트에 기록
try:
    if(list_check[0] == "1"):
        print("{} \tstart for google sheet writing".format(datetime.datetime.now()))

        wsheet = doc.worksheet(gs_sheet_rank)

        r = len(wsheet.col_values(1)) + 1

        list_values = []
        counter = 0

        for key_i in range(1,len(list_check)):
            if(list_check[key_i]==""): continue

            counter += 1
            list_row = []    
            list_row.append(datetime.datetime.now().strftime("%Y-%m-%d"))
            list_row.append(datetime.datetime.now().strftime("%H:%M:%S"))
            list_row.append(1)
            list_row.append(list_storename[key_i])
            list_row.append(list_keyword[key_i])
            list_row.append(list_rank[key_i])
            list_row.append( (int(list_rank[key_i])-1) // 40 + 1)
            list_values.append(list_row)

        # print(list_values)
        warnings.filterwarnings(action='ignore')
        wsheet.update('A'+str(r)+':G'+str(r+counter), list_values)
        warnings.filterwarnings(action='default')

        print("{} \tfinish for google sheet writing".format(datetime.datetime.now()))
except Exception as e:
    print("writing sheet exception:", e)