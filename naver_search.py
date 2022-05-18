# -*- coding: utf-8 -*-
import os
import re
import time
from urllib.parse import quote, urlparse
import requests
from dotenv import load_dotenv
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

load_dotenv()
filename = "data.xlsx"
sheet_name = "아이템 정보"

NAVER_CLIENT_ID = os.getenv('NAVER_CLIENT_ID')
NAVER_CLIENT_SECRET = os.getenv('NAVER_CLIENT_SECRET')

host_dict = dict(
    eleven_street="https://www.11st.co.kr/",
    lotte_i_mall="https://www.lotteimall.com/",
    lotte_on="https://www.lotteon.com/",
    g_market="http://item.gmarket.co.kr/",
    gee9="https://www.g9.co.kr/",
    gs_shop="https://with.gsshop.com/",
    on_style="https://display.cjonstyle.com/",
    auction="http://itempage3.auction.co.kr/",
    interpark="https://shopping.interpark.com/",
    smart_store="https://smartstore.naver.com/",
    we_make_price="https://front.wemakeprice.com/",
    tmon="https://www.tmon.co.kr/",
    unit="http://mahaknit.com/",
    shopping="https://shopping.naver.com/",
)
hosts = list(host_dict.values())

product_type = {
    "1": "일반상품 가격비교 상품",
    "2": "가격비교 비매칭 일반상품",
    "3": "가격비교 매칭 일반상품",
    "4": "중고상품 가격비교 상품",
    "5": "가격비교 비매칭 중고상품",
    "6": "가격비교 매칭 중고상품",
    "7": "단종상품 가격비교 상품",
    "8": "가격비교 비매칭 단종상품",
    "9": "가격비교 매칭 단종상품",
    "10": "판매예정상품 가격비교 상품",
    "11": "가격비교 비매칭 판매예정상품",
    "12": "가격비교 매칭 판매예정상품",
}


def extract_title(string):
    string = re.sub(r"\[?유닛]?", "", string)
    string = re.sub(r"\[?롯데백화점]?", "", string)
    string = re.sub(r"\[?\(?AVE\)?]?", "", string)
    string = re.sub(r"\[?\(?<b>.+</b>\)?]?", "", string)
    string = re.sub(r"\s+", " ", string)
    return string.strip()


def redirect_url(code, url, product_types):
    with Chrome() as driver:
        driver.get(url)
        driver.implicitly_wait(10)
        wait = WebDriverWait(driver, 30)
        try:
            btn1 = driver.find_element(By.XPATH, "//a[text()='사러가기']") or None
            btn1.click() if btn1 else None
            if not product_types.startswith("가격비교"):
                btn2 = driver.find_element(By.XPATH, "//a[text()='사러가기']") or None
                btn2.click() if btn2 else None
                tabs = driver.window_handles
                driver.switch_to.window(tabs[1])
            wait.until_not(EC.url_contains("cr.shopping.naver.com"))
        except NoSuchElementException as e:
            # print("실패:", product_types)
            pass
        except TimeoutException as e:
            print("응답 시간 초과")
        else:
            # print("성공:", product_types)
            pass
        parsed = urlparse(driver.current_url)
        host_url = f"{parsed.scheme}://{parsed.netloc}/"
        if host_url not in hosts:
            hosts.append(host_url)
            print(hosts)
        return driver.current_url


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
            # print(data["mallName"], product_type[data["productType"]])
            if not bigger_than(data['lprice'], low_price):
                url = redirect_url(word, data["link"], product_type[data["productType"]])
                title = extract_title(data['title'])
                print(title, data["lprice"], low_price, url)


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
        min_row=2,
        min_col=1,
    )
    for value in values:
        store = ["본점", "잠실", "청량리", "부산본점", "동탄", "중동", "대전", "광주", "강남", "수원", "동래", "노원", "대구", "안산", "평촌"]
        index, code, korean_name, on_off, year, season, tag_price, dsc_price, percent = value
        # index, code, name, season, tag_price, dsc_price, ten, fifteen, *stores = map(lambda x: x.value, value)
        naver_shopping_search(code.value, dsc_price.value)


if __name__ == '__main__':
    main()
