# -*- coding: utf-8 -*-
import os
import re
import csv
import time
from _csv import _writer

import requests
from datetime import datetime
from urllib.parse import quote, urlparse
from dotenv import load_dotenv
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
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

store_types = {'02-772-3343': '본점', '02-2143-7251': '잠실점', '02-2164-5353': '영등포점', '02-3707-1321': '청량리점',
               '02-3289-8007': '관악점', '02-531-2224': '강남점', '02-950-2268': '노원점', '02-944-2274': '미아점',
               '02-2218-3404': '건대점', '02-6116-3051': '김포공항점', '02-6965-2661': '서울역점', '031-738-2217': '분당점',
               '031-909-3486': '일산점', '032-320-7298': '중동점', '031-412-7772': '안산점', '031-8086-9250': '평촌점',
               '031-8066-0286': '수원점', '032-242-2234': '인천터미널점', '031-8036-3746': '동탄점', '051-810-4280': '부산본점',
               '062-221-1308': '광주점', '042-601-2337': '대전점', '054-230-1345': '포항점', '052-960-4954': '울산점',
               '051-668-4254': '동래점', '055-279-3377': '창원점', '053-660-3322': '대구점', '053-258-3213': '상인점',
               '063-289-3252': '전주점', '061-801-2156': '남악점', '051-678-3488': '광복점'}

store_types2 = {
    "EB": "본점", "IM": "본점", "BJ": "본점", "IJS": "잠실점", "EYD": "영등포점", "YDP": "영등포점", "ECL": "청량리점", "EGN": "강남점",
    "ENW": "노원점", "INW": "노원점", "TNW": "노원점", "EGP": "김포공항점", "EIS": "일산점", "ISS": "일산점", "EJD": "중동점", "IJD": "중동점",
    "EAS": "안산점", "IAS": "안산점", "AS": "안산점", "EPC": "평촌점", "IM": "평촌점", "ESW": "수원점", "ISW": "수원점", "ICC": "인천터미널점",
    "IC": "인천터미널점", "EDT": "동탄점", "IMT": "동탄점", "MDT": "동탄점", "EBS": "부산본점", "EGJ": "광주점", "IJ": "광주점", "EDJ": "대전점",
    "IDJ": "대전점", "EDR": "동래점", "IDR": "동래점", "EDG": "대구점", "IDG": "대구점", "DGG": "대구점", "EGB": "광복점",
}


def extract_title(string):
    string = re.sub(r"\[?유닛]?", "", string)
    string = re.sub(r"\[?롯데백화점]?", "", string)
    string = re.sub(r"\[?\(?AVE\)?]?", "", string)
    string = re.sub(r"\[?\(?<b>.+</b>\)?]?", "", string)
    string = re.sub(r"\s+", " ", string)
    return string.strip()


def extract_phone(string):
    phone_regex = re.compile(r"(0\d{1,2}-\d{3,4}-\d{4})|(1\d{3}-\d{4})")
    match = phone_regex.search(string)
    if match:
        return match.group()
    return "none"


def find_cs_number(page_source: str):
    phone = extract_phone(page_source)
    if phone in store_types:
        return store_types[phone]
    return phone


def get_host_from_url(url_string):
    hosts = list(host_dict.values())
    parsed = urlparse(url_string)
    host_url = f"{parsed.scheme}://{parsed.netloc}/"
    if host_url not in hosts:
        hosts.append(host_url)
    return host_url


def redirect_url(url, product_types):
    with Chrome() as driver:
        driver.implicitly_wait(10)
        wait = WebDriverWait(driver, 30)
        driver.get(url)
        try:
            btn1 = driver.find_element(By.XPATH, "//a[text()='사러가기']") or None
            btn1.click() if btn1 else None
            if not product_types.startswith("가격비교"):
                btn2 = driver.find_element(By.XPATH, "//a[text()='사러가기']") or None
                btn2.click() if btn2 else None
            tabs = driver.window_handles
            if len(tabs) > 1:
                driver.switch_to.window(tabs[1])
            if driver.current_url.startswith("https://display.cjonstyle.com/"):
                driver.execute_script('document.querySelector("#promotion_layer > div > div > div > a").click()')
            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "body > div")))
            wait.until_not(EC.url_matches(r"https://cr.shopping.naver.com/.*"))
            time.sleep(0.5)
            page_source = driver.page_source
            return page_source, driver.current_url
        except Exception as e:
            print("예외 발생", driver.current_url)


def find_model_name(source, word):
    model_name = re.search(f"{word}-([A-Za-z]+)", source)
    if model_name:
        if (model_name.group(1)).upper() in store_types2:
            return store_types2[model_name.group(1)]
    return "none"


def naver_shopping_search(csv_writer: _writer, index: int, season: str, word: str, low_price: int | str):
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
            if not bigger_than(data['lprice'], low_price):
                title = extract_title(data['title'])
                source, url = redirect_url(data["link"], product_type[data["productType"]])
                cs_number = find_cs_number(source)
                model_name = find_model_name(source, word)
                csv_writer.writerow([index, word, title, cs_number, model_name, season, int(data["lprice"]), int(low_price)])


def bigger_than(is_bigger, is_smaller):
    return int(is_bigger) >= int(is_smaller)


def main():
    dateformat = datetime.now().strftime("%Y%m%d-%H%M%S")
    new_filename = f".\\data\\result_{dateformat}.csv"
    with open(new_filename, "w", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=",", lineterminator="\n")
        writer.writerow(["NO", "스타일코드", "한글명", "지점_연락처", "지점_코드", "시즌", "판매가", "공식할인가"])
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
            index, code, korean_name, on_off, year, season, tag_price, dsc_price, percent = value
            # index, code, name, season, tag_price, dsc_price, ten, fifteen, *stores = map(lambda x: x.value, value)
            naver_shopping_search(writer, index.value, season.value, code.value, dsc_price.value)
        f.close()


if __name__ == '__main__':
    main()
