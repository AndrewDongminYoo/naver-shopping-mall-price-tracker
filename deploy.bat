pyinstaller --onefile --add-data "prod.env;." --add-binary "chromedriver.exe;." naver_search.py --distpath .
