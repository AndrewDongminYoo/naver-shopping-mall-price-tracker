import os
import re
import csv
import time
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
import chromedriver_autoinstaller

chrome_path = chromedriver_autoinstaller.install()
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

host_ko = {
    "https://www.11st.co.kr/": "11번가",
    "https://www.lotteimall.com/": "롯데아이몰",
    "https://www.lotteon.com/": "롯데온",
    "http://item.gmarket.co.kr/": "지마켓",
    "https://www.g9.co.kr/": "지구",
    "https://with.gsshop.com/": "지에스샵",
    "https://display.cjonstyle.com/": "온스타일",
    "http://itempage3.auction.co.kr/": "옥션",
    "https://shopping.interpark.com/": "인터파크",
    "https://smartstore.naver.com/": "스마트스토어",
    "https://front.wemakeprice.com/": "위메프",
    "https://www.tmon.co.kr/": "티몬",
    "http://mahaknit.com/": "마하유닛",
    "https://shopping.naver.com/": "네이버쇼핑",
}

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

store_types = {
    '02-772-3343': '본점', '02-2143-7251': '잠실점', '02-2164-5353': '영등포점', '02-3707-1321': '청량리점',
    '02-3289-8007': '관악점', '02-531-2224': '강남점', '02-950-2268': '노원점', '02-944-2274': '미아점',
    '02-2218-3404': '건대점', '02-6116-3051': '김포공항점', '02-6965-2661': '서울역점', '031-738-2217': '분당점',
    '031-909-3486': '일산점', '032-320-7298': '중동점', '031-412-7772': '안산점', '031-8086-9250': '평촌점',
    '031-8066-0286': '수원점', '032-242-2234': '인천터미널점', '031-8036-3746': '동탄점', '051-810-4280': '부산본점',
    '062-221-1308': '광주점', '042-601-2337': '대전점', '054-230-1345': '포항점', '052-960-4954': '울산점',
    '051-668-4254': '동래점', '055-279-3377': '창원점', '053-660-3322': '대구점', '053-258-3213': '상인점',
    '063-289-3252': '전주점', '061-801-2156': '남악점', '051-678-3488': '광복점'
}

store_types2 = {
    "EBJ": "본점", "IBJ": "본점", "EB": "본점", "IB": "본점", "BJ": "본점", "IJS": "잠실점", "EYD": "영등포점", "YDP": "영등포점",
    "ECL": "청량리점", "EGN": "강남점", "ENW": "노원점", "INW": "노원점", "TNW": "노원점", "EGP": "김포공항점", "EIS": "일산점", "ISS": "일산점",
    "EJD": "중동점", "IJD": "중동점", "EAS": "안산점", "IAS": "안산점", "AS": "안산점", "EPC": "평촌점", "IM": "평촌점", "ESW": "수원점",
    "ISW": "수원점", "ICC": "인천터미널점", "IC": "인천터미널점", "EDT": "동탄점", "IMT": "동탄점", "MDT": "동탄점", "EBS": "부산본점",
    "EGJ": "광주점", "IJ": "광주점", "EDJ": "대전점", "IDJ": "대전점", "EDR": "동래점", "IDR": "동래점", "EDG": "대구점", "IDG": "대구점",
    "DGG": "대구점", "EGB": "광복점",
}


def extract_title(string):
    string = re.sub(r"\[?유닛]?", "", string)
    string = re.sub(r"\[?롯데백화점( 2관)?]?", "", string)
    string = re.sub(r"\[?\(?AVE\)?]?", "", string)
    string = re.sub(r"</?b>", "", string)
    string = re.sub(r"\s+", " ", string)
    return string.strip()


def extract_phone(string):
    phone_regex = re.compile(r"(0\d{1,2}-\d{3,4}-\d{4})|(1\d{3}-\d{4})")
    match = phone_regex.search(string)
    if match:
        return match.group()
    return ""


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


def redirect_url(driver, url, product_types):
    find_all_href = [url]
    req = None
    while find_all_href:
        req = requests.get(find_all_href[0], allow_redirects=True)
        find_all_href = re.compile(r"https://cr.shopping.naver.com/[a-zA-Z\d%&?.=/]+").findall(req.text)
    find_all_redirects = re.compile(r'targetUrl = "([a-zA-Z\d:%&?.=/_]+)"').findall(req.text)
    driver.get(find_all_redirects.pop())
    tabs = driver.window_handles
    while len(tabs) > 1:
        driver.switch_to.window(tabs[1])
        driver.close()
        driver.switch_to.window(tabs[0])
    wait = WebDriverWait(driver, 30)
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "body > div")))
    return driver.page_source, driver.current_url


def find_model_name(source, word):
    model_name = re.search(f"{word}-([A-Za-z]+)", source)
    if model_name:
        if (model_name.group(1)).upper() in store_types2:
            return store_types2[model_name.group(1)]
    return ""


def naver_shopping_search(csv_writer, index: int, season: str, word: str, low_price: int):
    display = 100
    start = 1

    url = f"https://openapi.naver.com/v1/search/shop?" \
          f"query={quote(word)}" \
          f"&display={str(display)}" \
          f"&start={str(start)}"
    headers = {
        "X-Naver-Client-Id": NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": NAVER_CLIENT_SECRET
    }
    response = requests.get(url, headers=headers)
    start += display
    body = response.json()
    if "items" in body:
        with Chrome(executable_path=chrome_path) as driver:
            driver.implicitly_wait(10)
            for data in body["items"]:
                if not bigger_than(data['lprice'], low_price):
                    title = extract_title(data['title'])
                    try:
                        source, url = redirect_url(driver, data["link"], product_type[data["productType"]])
                        cs_number = find_cs_number(source)
                        model_name = find_model_name(source, word)
                        host_key = get_host_from_url(url)
                        host_name = host_ko.get(host_key, host_key)
                        row = [
                            index, word, title, cs_number, model_name, season,
                            int(data["lprice"]), low_price, int(low_price * 0.9),
                            host_name, url
                        ]
                        csv_writer.writerow(row)
                        print(row)
                    except TypeError:
                        pass


def bigger_than(is_bigger: str, is_smaller: int):
    discounted_price = is_smaller * 0.9
    return int(is_bigger) >= discounted_price


def main():
    dateformat = datetime.now().strftime("%Y%m%d-%H%M%S")
    new_filename = f"result_{dateformat}.csv"
    with open(new_filename, mode="w", newline="") as f:
        writer = csv.writer(f, delimiter=",", lineterminator="\n")
        header = [
            "NO", "스타일코드", "한글명",
            "지점_연락처", "지점_코드", "시즌",
            "판매가", "공식할인가", "추가할인가",
            "판매처", "링크"
        ]
        writer.writerow(header)
        wb: Workbook = load_workbook(
            filename=filename,
            data_only=True,
        )
        worksheet: Worksheet = wb[sheet_name]
        sheet_rows = worksheet.iter_rows(
            max_col=worksheet.max_column,
            max_row=worksheet.max_row,
            min_row=2,
            min_col=1,
            values_only=True,
        )
        for sheet_row in sheet_rows:
            index, code, korean_name, on_off, year, season, tag_price, dsc_price, percent = sheet_row
            naver_shopping_search(writer, index, season, code, dsc_price)
        f.close()


if __name__ == '__main__':
    main()
