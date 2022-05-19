# -*- coding: utf-8 -*-
import os
import re
from urllib.parse import quote, urlparse
import requests
import selenium.webdriver
from bs4 import BeautifulSoup
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


def extract_title(string):
    string = re.sub(r"\[?유닛]?", "", string)
    string = re.sub(r"\[?롯데백화점]?", "", string)
    string = re.sub(r"\[?\(?AVE\)?]?", "", string)
    string = re.sub(r"\[?\(?<b>.+</b>\)?]?", "", string)
    string = re.sub(r"\s+", " ", string)
    return string.strip()


def extract_phone(string):
    phone_regex = re.compile(r"(\d{2,3}-\d{3,4}-\d{4})|(\d{4}-\d{4})")
    match = phone_regex.search(string)
    if match:
        return match.group()
    return "[no-phone-number]"


def get_soup(source):
    return BeautifulSoup(source, "html.parser")


def get_number(soup, css_selector):
    string = soup.select_one(css_selector)
    return string.text if string else soup.find(css_selector).text


def find_cs_number(page_source: str, soup: BeautifulSoup, url_string):
    switch = {
        'https://www.lotteon.com/': "table > tr:-soup-contains('업체명') > td > p",
        'https://www.11st.co.kr/': "div.prodc_return_wrap > table > tbody > tr:nth-child(9) > td",
        'http://itempage3.auction.co.kr/':
            "#ucOfficialNotice_ulMain > li:nth-child(12) > span.cont",
        'https://www.lotteimall.com/':
            "#contents > div.detail_sec > div.division_product_tab > div.content_detail > div.wrap_detail.content2 > "
            "div > div:nth-child(3) > table > tbody > tr:last-child > td",
        'http://item.gmarket.co.kr/':
            "#vip-tab_detail > div.vip-detailarea_productinfo > div.box__product-notice-list > "
            "table:nth-child(2) > tbody > tr:-soup-contains('A/S') > td",
        'https://with.gsshop.com/':
            "#ProTab04 > div.normalN_table_wrap.more > table > tbody > tr:-soup-contains('A/S') > td",
        'https://display.cjonstyle.com/':
            "#_INFO > div:nth-child(3) > table > tbody > tr:last-child > td",
        'https://shopping.interpark.com/':
            "#productInfoProvideNotification > div:nth-child(3) > dl:-soup-contains('A/S') > dd",
        'https://smartstore.naver.com/':
            "#INTRODUCE > div > div.product_info_notice > div > table > tbody > tr:last-child",
        'https://front.wemakeprice.com/':
            "#productdetails > div > div.deal_detailinfo > ul > li > div > table > tbody > tr:-soup-contains('A/S') > "
            "td:nth-child(2)",
        'https://www.tmon.co.kr/':
            "#_wrapProductInfoNotes > div > div > div > table > tbody > tr:-soup-contains('A/S') > td",
        'https://www.g9.co.kr/': "#info_tab1_sub1 > div > div > table > tbody > tr:-soup-contains('A/S') > td > div",
        'https://shopping.naver.com/': "",
        'https://cr.shopping.naver.com/': "",
        'http://mahaknit.com/': "",
    }
    host_url = get_host_from_url(url_string)
    if switch.get(host_url):
        find_text = soup.select_one(switch[host_url])
        if find_text:
            return find_text.text.strip()
    return extract_phone(page_source)


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
            wait.until_not(EC.url_matches(r"https\:\/\/cr\.shopping\.naver\.com\/.*"))
            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "body > div")))
            page_source = driver.page_source or driver.execute_script("return document.body.innerHTML;")
            return page_source, driver.current_url
        except NoSuchElementException as e:
            print("대상 버튼 없음", driver.current_url)
        except TimeoutException as e:
            print("응답 시간 초과:", driver.current_url)


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
                title = extract_title(data['title'])
                source, url = redirect_url(data["link"], product_type[data["productType"]])
                soup = get_soup(source)
                cs_number = find_cs_number(source, soup, url)
                if cs_number == "[cannot-find]" or cs_number == "[no-phone-number]":
                    print(url)
                print(title, cs_number, data["lprice"], low_price)


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
        index, code, korean_name, on_off, year, season, tag_price, dsc_price, percent = value
        # index, code, name, season, tag_price, dsc_price, ten, fifteen, *stores = map(lambda x: x.value, value)
        naver_shopping_search(code.value, dsc_price.value)


if __name__ == '__main__':
    main()
