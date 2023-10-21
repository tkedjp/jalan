from time import sleep

import requests
from bs4 import BeautifulSoup
import pandas as pd

import math

hotel_list = []

year = input('チェックインする年を入力してください(入力できるのは1年後までです)：')
month = input('チェックインする月を入力してくださいい(入力できるのは1年後までです)：')
month = month.zfill(2)
day = input('チェックインする日を入力してください：')
day = day.zfill(2)
stay_count = int(input('泊数を入力してください：'))
if stay_count <= 9:
    stay_count = str(stay_count)
else:
    stay_count = input('泊数は1から9の間で入力してください：')
room_count = input('室数を入力してください：')
adult_num = input('人数を入力してください：')

base_url = 'https://www.jalan.net/040000/LRG_040200/SML_040202/?screenId=UWW1402&distCd=01&listId=0&activeSort=0&mvTabFlg=1&stayYear=2023&stayMonth=' + month + '&stayDay=' + day + '&stayCount=' + stay_count + '&roomCount=' + room_count +'&adultNum=' + adult_num +'&yadHb=1&roomCrack=200000&kenCd=040000&lrgCd=040200&smlCd=040202&vosFlg=6&idx={}'

user_agent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36'
header = {
    'User-Agent': user_agent
}

sleep(3)

r = requests.get(base_url, timeout=7.5, headers=header)
if r.status_code >= 400:
    print(F'{base_url}は無効です')
soup = BeautifulSoup(r.content, 'lxml')

#詳細ページ数と一覧ページ数取得
number = soup.select_one('span.jlnpc-listInformation--count').text
max_page_index = int(number) // 30 + 1
max_page_index = math.floor(max_page_index)
# print(max_page_index)

for i in range(max_page_index):
    url = base_url.format(30*i)
    # print(f'{i+1}ページ目URL：{url}')

    sleep(3)

    page_r = requests.get(url, timeout=7.5, headers=header)
    if page_r.status_code >= 400:
        print(F'{url}は無効です')
        continue

    page_soup = BeautifulSoup(page_r.content, 'lxml')

    #最安料金と1人あたりの料金
    table_soup = page_soup.select('div.p-yadoCassette__body.p-searchResultItem__body')
    for table in table_soup:
        hotel = table.select_one('a > div > div > div.p-searchResultItem__summaryInner > div.p-searchResultItem__summaryLeft > h2').text
        room_price = table.select_one('a > div > div > div.p-searchResultItem__summaryInner > div.p-searchResultItem__summaryRight > dl > dd > span.p-searchResultItem__lowestPriceValue').text
        per_price_tag = table.select_one('a > div > div > div.p-searchResultItem__summaryInner > div.p-searchResultItem__summaryRight > dl > dd > span.p-searchResultItem__lowestUnitPrice')
        if per_price_tag is None:
            per_price = None

        else:
            per_price = per_price_tag.text

        page_urls = table.select('a.jlnpc-yadoCassette__link')

        for i, page_url in enumerate(page_urls):
            page_url = 'https://www.jalan.net' + page_url.get('href')
            if 'javascript' in page_url:
                page_urls = None
            
            else:
                sleep(3)

                hotel_page_r = requests.get(page_url, timeout=7.5, headers=header)
                if hotel_page_r.status_code >= 400:
                    print(F'{page_url}は無効です')
                    continue
            
                #ホテル名
                hotel_page_soup = BeautifulSoup(hotel_page_r.content, 'lxml')
                # hotel_name = hotel_page_soup.select_one('#pankuzu > h1').text

                #住所
                address_tags = hotel_page_soup.select_one('#jlnpc-main-contets-area > div.shisetsu-accesspartking_body_wrap > table tr:nth-child(1) > td')
                if address_tags is None:
                    address = None

                else:
                    address = address_tags.text
                    address = address.replace('大きな地図をみる', '')
                    address = address.strip()
            
                #駐車場
                parking_tags = hotel_page_soup.select_one('#jlnpc-main-contets-area > div.shisetsu-accesspartking_body_wrap > table tr:nth-child(3) > td')  
                if parking_tags is None:
                    parking = None
                
                else:
                    parking = parking_tags.text
                    parking = parking.replace('\n','')
                    parking = parking.strip()

                #タイプ別の室数
                room_tag = hotel_page_soup.select_one('.shisetsu-roomsetsubi_body')
                if room_tag is None:
                    single = None
                    double = None
                    twin = None
                    sweet = None
                    total = None

                else:
                    tags = room_tag.text

                    if '総部屋数' not in tags:
                        single = room_tag.select_one('tr:nth-child(2) > td > div > table tr:nth-child(2) > td:first-child').text
                        double = room_tag.select_one('tr:nth-child(2) > td > div > table tr:nth-child(2) > td:nth-child(2)').text
                        twin = room_tag.select_one('tr:nth-child(2) > td > div > table tr:nth-child(2) > td:nth-child(3)').text
                        sweet = room_tag.select_one('tr:nth-child(2) > td > div > table tr:nth-child(2) > td:last-child').text
                        total = None

                    elif 'シングル' not in tags:
                        single = None
                        double = None
                        twin = None
                        sweet = None
                        total = room_tag.select_one('tr:nth-child(1) > td > div > table tr:nth-child(2) > td:nth-child(5)').text
                        total = total.strip()

                    else:
                        single = room_tag.select_one('tr:nth-child(3) > td > div > table tr:nth-child(2) > td:first-child').text
                        double = room_tag.select_one('tr:nth-child(3) > td > div > table tr:nth-child(2) > td:nth-child(2)').text
                        twin = room_tag.select_one('tr:nth-child(3) > td > div > table tr:nth-child(2) > td:nth-child(3)').text
                        sweet = room_tag.select_one('tr:nth-child(3) > td > div > table tr:nth-child(2) > td:last-child').text
                        total = room_tag.select_one('tr:nth-child(1) > td > div > table tr:nth-child(2) > td:nth-child(5)').text
                        total = total.strip()

                hotel_list.append({
                    'ホテル名': hotel,
                    '詳細ページ': page_url,
                    '住所': address,
                    'シングル': single,
                    'ダブル': double,
                    'ツイン': twin,
                    'スイート': sweet,
                    '総部屋数': total,
                    '室料': room_price,
                    '1人あたり': per_price,
                    '駐車場': parking
                })
                print(hotel_list[-1])

#スプレッドシート出力
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# データフレームをGoogleスプレッドシートに書き出す
from gspread_dataframe import set_with_dataframe

#認証
SCOPES = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'xxxxx.json'
credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, SCOPES)
gs = gspread.authorize(credentials)
SPREADSHEET_KEY = 'xxxxx'

#書込
df = pd.DataFrame(hotel_list)
workbook = gs.open_by_key(SPREADSHEET_KEY)
set_with_dataframe(workbook.worksheet("シート1"), df, include_index=False)