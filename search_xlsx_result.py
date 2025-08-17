import time, subprocess, json, re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# 전화번호 저장용 리스트
phone_numbers = []
idx = 0

# 엑셀 파일 명
xlsx_name = "lawyer_result.xlsx"

# ... 네가 가진 import/옵션/driver/엑셀 로드/normalize_e164 그대로 ...

def xpath_literal(s: str) -> str:
  if "'" not in s:
    return f"'{s}'"
  if '"' not in s:
    return f'"{s}"'
  parts = s.split("'")
  return "concat(" + ", \"'\", ".join([f"'{p}'" for p in parts]) + ")"

def title_matches(driver, name: str) -> bool:
  try:
    t = WebDriverWait(driver, 6).until(
      EC.presence_of_element_located((By.CLASS_NAME, "DUwDvf"))
    ).text.strip()
    return t.lower() == (name or "").strip().lower()
  except:
    return False

def click_exact_in_list(driver, target_name: str, hint_addr: str = None) -> bool:
  """
  좌측 결과 리스트에서 target_name과 제목이 정확히 같은 카드를 클릭.
  hint_addr가 주어지면, 동일 제목이 여러 개인 경우 카드의 보이는 텍스트에 hint_addr가 포함된 것을 우선 클릭.
  """
  # 리스트 컨테이너 존재 확인
  try:
    WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='feed']")))
  except:
    return False

  # 제목 정확 일치 경로들 (Maps DOM 변동 대비 3가지 전략)
  xpaths = [
    f"//div[@role='feed']//a[contains(@href,'/place')]" \
    f"[.//div[contains(@class,'fontHeadline')][normalize-space()={xpath_literal(target_name)}]]",

    f"//div[@role='feed']//a[contains(@href,'/place')]" \
    f"[.//div[contains(@class,'qBF1Pd')][normalize-space()={xpath_literal(target_name)}]]",

    # aria-label에 가게명이 들어오는 케이스
    f"//div[@role='feed']//*[(@aria-label={xpath_literal(target_name)}) and (self::a or self::div or self::span)]" \
    "/ancestor::a[contains(@href,'/place')][1]"
  ]

  # 후보 수집
  candidates = []
  for xp in xpaths:
    els = driver.find_elements(By.XPATH, xp)
    for el in els:
      if el not in candidates:
        candidates.append(el)

  if not candidates:
    return False

  # hint_addr가 있으면 hint를 포함한 카드 우선
  def card_text(e):
    try:
      return e.text or ""
    except:
      return ""

  if hint_addr:
    hinted = [c for c in candidates if hint_addr.strip() and hint_addr.strip().lower() in card_text(c).lower()]
    order = hinted + [c for c in candidates if c not in hinted]
  else:
    order = candidates

  # 순서대로 클릭 시도
  for card in order:
    try:
      driver.execute_script("arguments[0].scrollIntoView({block:'center'});", card)
      time.sleep(0.15)
      card.click()
      WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "DUwDvf")))
      return True
    except:
      continue

  return False

# ====== 전화번호 추출 ======
def normalize_e164(raw: str) -> str:
  if not raw:
    return ""
  s = raw.strip()
  if s.startswith("tel:"):
    s = s[4:]
  s = re.sub(r"[^\d+]", "", s)  # +와 숫자만 남김
  if s.count("+") > 1:
    s = "+" + s.replace("+", "")
  return s

subprocess.Popen('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\chromeCookie\\kmong_Rohmin_leisure"'.format("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"))

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
driver.get("https://www.google.com/maps")

time.sleep(0.5)

# 엑셀 파일 불러오기
df = pd.read_excel(xlsx_name)

# 1) 특정 컬럼만 리스트로 변환
name_list = df["회사명"].tolist()
addr_list = df["주소"].tolist()

# 2) 행 단위 순회
for row in name_list:
  # 검색창 찾기
  search_box = driver.find_element(By.ID, "searchboxinput")
  search_box.clear()
  search_box.send_keys(f"\"{row}\" {addr_list[idx]}")
  search_box.send_keys(Keys.ENTER)

  # 검색 결과/상세 로딩 대기
  WebDriverWait(driver, 12).until(
    EC.any_of(
      EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='feed']")),
      EC.presence_of_element_located((By.CLASS_NAME, "DUwDvf"))
    )
  )

  # 상세가 아니라면(여러 결과) → 좌측 리스트에서 '정확 일치' 클릭 (주소 힌트 사용)
  if not title_matches(driver, row):
    click_exact_in_list(driver, row, hint_addr=str(addr_list[idx]) if idx < len(addr_list) else None)
    # 그래도 상세가 아니거나 제목 불일치라면 이번 건은 빈값 처리하고 다음으로
    if not title_matches(driver, row):
      phone_numbers.append("")
      print(f"{idx} :: 리스트 다중결과 - 정확 일치 미탐: {row}")
      idx += 1
      continue

  time.sleep(1)

  # ===== 전화번호 추출 (네 기존 1/2/3단계 그대로) =====
  phone_display, phone_e164, phone_source = "", "", ""

  try:
    btn = WebDriverWait(driver, 8).until(
      EC.presence_of_element_located((By.CSS_SELECTOR, "button[data-item-id^='phone']"))
    )
    txt = btn.text.strip()
    if txt:
      phone_display = txt
      phone_e164 = normalize_e164(txt)
      phone_source = "data-item-id=phone button"
  except:
    pass

  if not phone_display:
    try:
      tel_a = driver.find_element(By.CSS_SELECTOR, "a[href^='tel:']")
      href = (tel_a.get_attribute("href") or "").strip()
      if href:
        phone_display = href.replace("tel:", "").strip()
        phone_e164 = normalize_e164(href)
        phone_source = "tel: href"
    except:
      pass

  if not phone_display:
    try:
      nodes = driver.find_elements(By.XPATH, "//*[@aria-label[contains(., '전화:')]]")
      for el in nodes:
        al = el.get_attribute("aria-label") or ""
        m = re.search(r"전화:\s*([^\n\r]+)", al)
        if m:
          phone_display = m.group(1).strip()
          phone_e164 = normalize_e164(phone_display)
          phone_source = "aria-label"
          break
    except:
      pass

  phone_numbers.append(phone_display)
  print(f"{idx} :: {phone_display} ({phone_source})")
  idx += 1

# 루프 끝난 후 저장
df["전화번호"] = phone_numbers
df.to_excel(xlsx_name, index=False)
print(f"저장 완료: {xlsx_name}")