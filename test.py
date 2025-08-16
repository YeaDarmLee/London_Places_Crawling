import time, subprocess, json, re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

subprocess.Popen('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\kmongCookie\\London_Places_Crawling"'.format("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"))

# Selenium 옵션 설정
options = Options()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--ignore-certificate-errors')
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

# ChromeDriver 실행
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# 사이트 진입
driver.get("https://www.google.com/maps/place/London+Bridge/@51.5058503,-0.0886118,17.5z/data=!4m14!1m7!3m6!1s0x4876035159bb13c5:0xa61e28267c3563ac!2z65-w642YIOq1kA!8m2!3d51.5078788!4d-0.0877321!16zL20vMHA3N2c!3m5!1s0x4876035747ecc86f:0x949a2d8ba1bca2df!8m2!3d51.5058624!4d-0.0869692!16s%2Fg%2F1ptxs65xf?entry=ttu&g_ep=EgoyMDI1MDgxMC4wIKXMDSoASAFQAw%3D%3D")

# search_input = driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div[9]/div[3]/div[1]/div[1]/div/div[2]/form/input")
# search_input.send_keys('서울')