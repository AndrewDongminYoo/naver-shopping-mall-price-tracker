# -*- coding: utf-8 -*-
import os
import re
from urllib.parse import quote
import requests
from dotenv import load_dotenv
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet

load_dotenv()
filename = "data.xlsx"
sheet_name = "Sheet1"

NAVER_CLIENT_ID = os.getenv('NAVER_CLIENT_ID')
NAVER_CLIENT_SECRET = os.getenv('NAVER_CLIENT_SECRET')

product_type = {
    "1": "일반상품	가격비교 상품",
    "2": "가격비교 비매칭 일반상품",
    "3": "가격비교 매칭 일반상품",
    "4": "중고상품	가격비교 상품",
    "5": "가격비교 비매칭 일반상품",
    "6": "가격비교 매칭 일반상품",
    "7": "단종상품	가격비교 상품",
    "8": "가격비교 비매칭 일반상품",
    "9": "가격비교 매칭 일반상품",
    "10": "판매예정상품	가격비교 상품	",
    "11": "가격비교 비매칭 일반상품",
    "12": "가격비교 매칭 일반상품",
}


def extract_title(string):
    string = re.sub(r"\[?유닛]?", "", string)
    string = re.sub(r"\[?롯데백화점]?", "", string)
    string = re.sub(r'\[?\(?AVE\)?]?', '', string)
    string = re.sub(r'\[?\(?<b>.+</b>\)?]?', '', string)
    string = re.sub(r'\s+', ' ', string)
    return string.strip()


def naver_shopping_search(word, low_price):
    display = 100
    start = 1

    keyword = quote(word)
    url = f"https://openapi.naver.com/v1/search/shop?" \
          f"query={keyword}" \
          f"&display={display}" \
          f"&start={start}"
    headers = {
        "X-Naver-Client-Id": NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": NAVER_CLIENT_SECRET
    }
    response = requests.get(url, headers=headers)
    start += display
    body = response.json()
    if "items" in body:
        for data in body["items"]:
            print(data["mallName"], product_type[data["productType"]])
            if not bigger_than(data['lprice'], low_price):
                title = extract_title(data['title'])
                print(title, data["lprice"], low_price, data["link"])


def bigger_than(is_bigger, is_smaller):
    return int(is_bigger) >= int(is_smaller)


def main():
    wb: Workbook = load_workbook(
        filename=filename,
        data_only=True,
    )
    sheet_ranges: Worksheet = wb[sheet_name]
    values = sheet_ranges.iter_rows(
        max_col=sheet_ranges.max_column,
        max_row=sheet_ranges.max_row,
        min_row=3,
        min_col=1,
    )
    for value in values:
        store = ["본점", "잠실", "청량리", "부산본점", "동탄", "중동", "대전", "광주", "강남", "수원", "동래", "노원", "대구", "안산", "평촌"]
        index, code, name, season, tag_price, dsc_price, ten, fifteen, *stores = value
        # index, code, name, season, tag_price, dsc_price, ten, fifteen, *stores = map(lambda x: x.value, value)
        naver_shopping_search(code.value, fifteen.value)


if __name__ == '__main__':
    main()
