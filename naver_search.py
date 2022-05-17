# -*- coding: utf-8 -*-
import os
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


def naver_shopping_search(word, low_price):
    display = 100
    start = 1
    brand_list = []

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
    for data in body["items"]:
        print(data['title'], compare(data['lprice'], low_price))


def compare(price1, price2):
    print(price1, price2)
    return int(price1) >= price2

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
        values_only=True,
    )
    for value in values:
        print(value)
        # naver_shopping_search(value[1].value, value[7].value)


if __name__ == '__main__':
    main()
