import time, re, html
import pandas as pd
from urllib.parse import urlparse, urljoin
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

# ==============================
# 1) 설정값
# ==============================
HARD_WAIT = 7                 # 요소 대기(초) - 느리면 9~10
URL_HARD_LIMIT = 60           # URL 단위 최대 처리 시간(초)
CHECKPOINT_EVERY = 1000       # N행마다 저장
OUT_XLSX_NAME = "beauty_result_filled.xlsx"
SKIP_ON_CRASH = True          # 탭 크래시/브라우저 크래시 시 즉시 스킵

# ==============================
# 2) 정규식/필터/스코어링
# ==============================
import re
EMAIL_REGEX = re.compile(r'\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,24}\b', re.IGNORECASE)
BAD_TLDS = {
  "css","js","map","json","png","jpg","jpeg","gif","webp","svg","ico",
  "woff","woff2","ttf","otf","mp4","webm","mov","avi","pdf","zip",
  "rar","7z","gz","tar","xml","html","htm"
}

def link_weight(txt_lower, href_lower):
  # contact > about > support
  if "contact" in txt_lower or "contact" in href_lower: return 3
  if "about"   in txt_lower or "about"   in href_lower: return 2
  if "support" in txt_lower or "support" in href_lower: return 1
  return 0

def is_valid_email(e):
  if not e or "@" not in e: return False
  e = e.strip()
  if not EMAIL_REGEX.fullmatch(e): return False
  if any(ch in e for ch in [' ', ',', ';', '<', '>', '"', "'"]): return False
  if e.count("@") != 1: return False
  local, _, dom = e.rpartition("@")
  if not local or not dom: return False
  if local.startswith(".") or local.endswith(".") or ".." in local: return False
  if dom.startswith(".")   or dom.endswith(".")   or ".." in dom:   return False
  tld = dom.split(".")[-1].lower()
  if tld in BAD_TLDS: return False
  if len(e) > 254 or len(local) > 64: return False
  return True

# ==============================
# 3) 드라이버 생성 (URL마다 새 창)
# ==============================
CHROMEDRIVER_BIN = ChromeDriverManager().install()

def start_driver():
  options = Options()
  # 브라우저 보이게(헤드리스 X)
  # options.add_argument('--headless=new')  # 필요시 헤드리스, 지금은 표시 목적이라 주석
  options.add_argument('--no-sandbox')
  options.add_argument('--disable-dev-shm-usage')
  options.add_argument('--ignore-certificate-errors')
  options.add_argument('--disable-extensions')
  options.add_argument('--disable-gpu')
  options.add_argument('--disable-features=site-per-process')
  options.add_argument('--js-flags=--max-old-space-size=128')
  options.add_argument("--disable-blink-features=AutomationControlled")
  options.add_argument('--start-maximized')
  options.add_argument('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36')

  # 완전 로드 기다리지 않고 DOM 등장까지만 대기
  options.set_capability('pageLoadStrategy', 'none')

  # 무거운 리소스 차단(이미지/폰트/영상). CSS는 기본 허용
  prefs = {
    "profile.managed_default_content_settings.images": 2,
    "profile.managed_default_content_settings.javascript": 1,
    # "profile.managed_default_content_settings.stylesheets": 2,
  }
  options.add_experimental_option("prefs", prefs)

  service = Service(CHROMEDRIVER_BIN)
  drv = webdriver.Chrome(service=service, options=options)
  try:
    drv.set_page_load_timeout(HARD_WAIT)
  except Exception:
    pass

  # 네트워크 차단(CDP)
  try:
    drv.execute_cdp_cmd("Network.enable", {})
    drv.execute_cdp_cmd("Network.setBlockedURLs", {
      "urls": [
        "*.png","*.jpg","*.jpeg","*.gif","*.webp","*.svg",
        "*.ico","*.mp4","*.avi","*.webm","*.mov",
        "*.woff","*.woff2","*.ttf","*.otf",
        # "*.css",
      ]
    })
  except Exception:
    pass
  return drv

def fast_get(driver, url, wait_selector=("tag name", "body")):
  """pageLoadStrategy='none' 기준: body 등장까지만 대기. Timeout 시 stopLoading 후 진행."""
  try:
    driver.set_page_load_timeout(HARD_WAIT)
  except Exception:
    pass
  try:
    driver.get(url)
    by, sel = wait_selector
    if by.lower() == "css":
      WebDriverWait(driver, HARD_WAIT).until(EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
    elif by.lower() == "xpath":
      WebDriverWait(driver, HARD_WAIT).until(EC.presence_of_element_located((By.XPATH, sel)))
    else:
      WebDriverWait(driver, HARD_WAIT).until(EC.presence_of_element_located((By.TAG_NAME, sel)))
  except TimeoutException:
    try:
      driver.execute_cdp_cmd("Page.stopLoading", {})
    except Exception:
      pass
  except Exception:
    raise

# ==============================
# 4) 데이터 적재
# ==============================
xlsx_name = "beauty_result.xlsx"
df = pd.read_excel(xlsx_name)

# 결과 컬럼 보장
if "이메일 주소" not in df.columns:
  df["이메일 주소"] = "-"

# ==============================
# 5) 메인 루프 (URL마다 새 드라이버 생성/종료)
# ==============================
row_no = 0
for i in df.index:
  row_no += 1
  url_val = df.at[i, "웹사이트 주소"]

  result_text = "조회할 사이트정보 없음"

  if pd.isna(url_val):
    df.at[i, "이메일 주소"] = result_text
    print(f"{row_no} :: 조회할 사이트정보 없음")
    if row_no % CHECKPOINT_EVERY == 0:
      df.to_excel(OUT_XLSX_NAME, index=False)
      print(f"[checkpoint] saved {row_no} rows → {OUT_XLSX_NAME}")
    continue

  url = str(url_val).strip()
  print(f"{row_no} :: {url}")

  attempt = 0
  written = False
  t0 = time.time()

  while attempt < 2 and not written:
    # URL 하드 타임박스
    if time.time() - t0 > URL_HARD_LIMIT:
      print("  -> url hard timeout, skip")
      df.at[i, "이메일 주소"] = "조회 중 시간초과"
      written = True
      break

    driver = None
    try:
      # === URL마다 새 브라우저 띄우기 ===
      driver = start_driver()

      # 메인 페이지
      fast_get(driver, url)

      all_candidates = set()

      # 1) mailto:
      try:
        for a in driver.find_elements(By.CSS_SELECTOR, 'a[href^="mailto:"]'):
          href = (a.get_attribute("href") or "").strip()
          if not href: continue
          base = href.split("?", 1)[0].replace("mailto:", "").strip()
          if is_valid_email(base): all_candidates.add(base)
          if "?" in href:
            qs = href.split("?", 1)[1]
            for kv in qs.split("&"):
              if "=" in kv:
                k, v = kv.split("=", 1)
                if k.lower() in ("to","cc","bcc"):
                  for e in v.split(","):
                    e = e.strip()
                    if is_valid_email(e): all_candidates.add(e)
      except Exception:
        pass

      # 2) 보이는 텍스트만 분석
      try:
        raw_html = driver.page_source or ""
        soup = BeautifulSoup(raw_html, "lxml")
        for tag in soup(["script","style","noscript","template"]):
          tag.decompose()
        text = soup.get_text(" ", strip=True)

        deob = html.unescape(text)
        for pat, rep in [
          (r'\s*\[?\s*at\s*\]?\s*', '@'),
          (r'\s*\(?\s*at\s*\)?\s*', '@'),
          (r'\s+at\s+', '@'),
          (r'\s*\[?\s*dot\s*\]?\s*', '.'),
          (r'\s*\(?\s*dot\s*\)?\s*', '.'),
          (r'\s+dot\s+', '.'),
          (r'골뱅이', '@'),
          (r'\s*점\s*', '.'),
          (r'닷', '.'),
        ]:
          deob = re.sub(pat, rep, deob, flags=re.IGNORECASE)
        deob = re.sub(r'\s*@\s*', '@', deob)
        deob = re.sub(r'\s*\.\s*', '.', deob)

        for e in EMAIL_REGEX.findall(deob):
          if is_valid_email(e): all_candidates.add(e)
      except Exception:
        pass

      # 3) 후보 링크(가중치, 최대 3)
      try:
        anchors = driver.execute_script("""
          return Array.from(document.querySelectorAll('a[href]'))
            .map(a => [a.href, (a.textContent||'').trim().toLowerCase()]);
        """) or []

        host = urlparse(url).netloc
        scored = []
        for href, txt in anchors:
          if not href or href.startswith("mailto:"): continue
          href2 = urljoin(url, href)
          if urlparse(href2).netloc != host: continue
          low = href2.lower()
          w = link_weight(txt, low)
          if w > 0:
            scored.append((w, len(href2), href2))

        scored.sort(key=lambda x: (-x[0], x[1], x[2]))
        cand_links = [h for _,_,h in scored[:3]]

        for link in cand_links:
          if time.time() - t0 > URL_HARD_LIMIT:
            print("  -> url hard timeout during subpage, cut off")
            break

          try:
            fast_get(driver, link)

            # mailto
            for a in driver.find_elements(By.CSS_SELECTOR, 'a[href^="mailto:"]'):
              href = (a.get_attribute("href") or "").strip()
              if not href: continue
              base = href.split("?", 1)[0].replace("mailto:", "").strip()
              if is_valid_email(base): all_candidates.add(base)
              if "?" in href:
                qs = href.split("?", 1)[1]
                for kv in qs.split("&"):
                  if "=" in kv:
                    k, v = kv.split("=", 1)
                    if k.lower() in ("to","cc","bcc"):
                      for e in v.split(","):
                        e = e.strip()
                        if is_valid_email(e): all_candidates.add(e)

            # 보이는 텍스트
            raw_html2 = driver.page_source or ""
            soup2 = BeautifulSoup(raw_html2, "lxml")
            for tag in soup2(["script","style","noscript","template"]):
              tag.decompose()
            text2 = soup2.get_text(" ", strip=True)

            deob2 = html.unescape(text2)
            for pat2, rep2 in [
              (r'\s*\[?\s*at\s*\]?\s*', '@'),
              (r'\s*\(?\s*at\s*\)?\s*', '@'),
              (r'\s+at\s+', '@'),
              (r'\s*\[?\s*dot\s*\]?\s*', '.'),
              (r'\s*\(?\s*dot\s*\)?\s*', '.'),
              (r'\s+dot\s+', '.'),
              (r'골뱅이', '@'),
              (r'\s*점\s*', '.'),
              (r'닷', '.'),
            ]:
              deob2 = re.sub(pat2, rep2, deob2, flags=re.IGNORECASE)
            deob2 = re.sub(r'\s*@\s*', '@', deob2)
            deob2 = re.sub(r'\s*\.\s*', '.', deob2)

            for e in EMAIL_REGEX.findall(deob2):
              if is_valid_email(e): all_candidates.add(e)

          except Exception:
            pass
      except Exception:
        pass

      # 4) 점수화 & 상위 1~3 콤마 저장
      parts = urlparse(url).netloc.lower().split(".")
      if len(parts) >= 3 and parts[-2:] == ["co", "uk"]:
        base_dom = ".".join(parts[-3:])
      else:
        base_dom = ".".join(parts[-2:]) if len(parts) >= 2 else urlparse(url).netloc.lower()

      score_map = {
        "info": 4, "hello": 4, "contact": 4, "support": 3, "help": 3,
        "sales": 3, "admin": 2, "team": 2, "office": 2, "enquiries": 2
      }

      ranked = []
      for e in sorted(all_candidates):
        local, _, dom = e.partition("@")
        s = 0
        if base_dom and base_dom in dom: s -= 5
        s -= score_map.get(local, 0)
        ranked.append((s, e))
      ranked.sort()

      if ranked:
        top_emails = [e for _, e in ranked[:3]]
        result_text = ", ".join(top_emails)
      else:
        result_text = "조회결과 없음"

      df.at[i, "이메일 주소"] = result_text
      print(f"  -> TOP3: {result_text}")
      written = True

    except WebDriverException as e:
      msg = str(e).lower()
      if any(x in msg for x in ["tab crashed", "chrome not reachable", "no such window"]):
        print("  -> 크래시 감지", ("(스킵)" if SKIP_ON_CRASH else "(재시도)"))
        if SKIP_ON_CRASH:
          df.at[i, "이메일 주소"] = "크래시로 스킵"
          written = True
        else:
          attempt += 1  # 재시도 모드일 때만 증가
      else:
        print("  -> WebDriver 예외:", e)
        df.at[i, "이메일 주소"] = "조회 중 오류"
        written = True

    except Exception as e:
      print("  -> 일반 예외:", e)
      df.at[i, "이메일 주소"] = "조회 중 오류"
      written = True

    finally:
      # === 이 URL용 브라우저 닫기 ===
      try:
        if driver: driver.quit()
      except Exception:
        pass

    # 재시도 플래그 처리
    if not SKIP_ON_CRASH and not written:
      attempt += 1

  # 체크포인트 저장
  if row_no % CHECKPOINT_EVERY == 0:
    df.to_excel(OUT_XLSX_NAME, index=False)
    print(f"[checkpoint] saved {row_no} rows → {OUT_XLSX_NAME}")

# ==============================
# 6) 최종 저장
# ==============================
try:
  df.to_excel(OUT_XLSX_NAME, index=False)
  print(f"최종 저장: {OUT_XLSX_NAME}")
except Exception as e:
  print("엑셀 저장 중 오류:", e)
