import time, subprocess, json, re, html
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

# ------------------------------------------------------------
# 0) 원격 디버깅 크롬 실행(프로필 재사용)
# ------------------------------------------------------------
subprocess.Popen(
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe '
  '--remote-debugging-port=9222 '
  '--user-data-dir="C:\\chromeCookie\\kmong_Rohmin_leisure"'
)

# ------------------------------------------------------------
# 1) 옵션/성능 튜닝
# ------------------------------------------------------------
options = Options()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--ignore-certificate-errors')
options.add_argument('--disable-extensions')
options.add_argument('--disable-gpu')
options.add_argument('--disable-features=site-per-process')
options.add_argument('--js-flags=--max-old-space-size=128')
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

# 완전 로드를 기다리지 않고, 필요한 요소만 명시적으로 대기
options.set_capability('pageLoadStrategy', 'none')

# 무거운 리소스 차단(필요 시 CSS도 차단 가능)
prefs = {
  "profile.managed_default_content_settings.images": 2,
  "profile.managed_default_content_settings.javascript": 1,
  # "profile.managed_default_content_settings.stylesheets": 2,  # 문제 있으면 주석 해제
}
options.add_experimental_option("prefs", prefs)

# ------------------------------------------------------------
# 2) 타임아웃/네비게이션 헬퍼
# ------------------------------------------------------------
HARD_WAIT = 7  # 느리면 9~10으로만 올려봐

def fast_get(driver, url, wait_selector=("tag name", "body")):
  """
  pageLoadStrategy='none' 기준: 'body' 등장까지만 기다림.
  타임아웃시 로딩 중단(Page.stopLoading) 후 현재까지 로드된 DOM으로 진행.
  """
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

# ------------------------------------------------------------
# 3) 이메일 정규식/필터/스코어링
# ------------------------------------------------------------
EMAIL_REGEX = re.compile(r'\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,24}\b', re.IGNORECASE)

# 파일 확장자처럼 보이는 TLD는 이메일로 취급하지 않음(오탐 컷)
BAD_TLDS = {
  "css","js","map","json","png","jpg","jpeg","gif","webp","svg","ico",
  "woff","woff2","ttf","otf","mp4","webm","mov","avi","pdf","zip",
  "rar","7z","gz","tar","xml","html","htm"
}

def link_weight(txt_lower, href_lower):
  # 후보 링크 가중치: contact > about > support
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

def start_driver():
  service = Service(ChromeDriverManager().install())
  drv = webdriver.Chrome(service=service, options=options)
  try:
    drv.set_page_load_timeout(HARD_WAIT)
  except Exception:
    pass
  # 리소스 차단(CDP)
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

# ------------------------------------------------------------
# 4) 데이터 적재
# ------------------------------------------------------------
out_rows = []  # 콤마로 합친 Top1~3 저장
idx = 0
xlsx_name = "clinic_result.xlsx"
df = pd.read_excel(xlsx_name)
url_list = df["웹사이트 주소"].tolist()

# 최초 드라이버 기동
driver = start_driver()

# ------------------------------------------------------------
# 5) 메인 루프(크래시 즉시 재기동 + 1회 재시도)
# ------------------------------------------------------------
for url in url_list:
  idx += 1
  top_joined = "조회할 사이트정보 없음"

  # NaN 처리
  if pd.isna(url):
    out_rows.append({'이메일 주소': top_joined})
    print(f"{idx} :: 조회할 사이트정보 없음")
    continue

  url = str(url).strip()
  print(f"{idx} :: {url}")

  attempt = 0
  success = False
  appended = False

  while attempt < 2 and not success:
    try:
      # 메인 페이지
      fast_get(driver, url)

      visited = [url]
      all_candidates = set()

      # ---- 1) mailto: 회수 ----
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

      # ---- 2) 보이는 텍스트에서만 추출 ----
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

      # ---- 3) 후보 링크: 가중치(contact>about>support), 최대 3 ----
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
          try:
            fast_get(driver, link)
            visited.append(link)

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

            # 보이는 텍스트 재추출
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

      # ---- 4) 점수화 후 상위 1~3 → 콤마 저장 ----
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
        if base_dom and base_dom in dom: s -= 5      # 회사 도메인 가점(낮을수록 상위)
        s -= score_map.get(local, 0)                 # 로컬파트 가점
        ranked.append((s, e))
      ranked.sort()

      if ranked:
        top_emails = [e for _, e in ranked[:3]]
        top_joined = ", ".join(top_emails)
      else:
        top_joined = "조회결과 없음"

      out_rows.append({'이메일 주소': top_joined})
      print(f"  -> TOP3: {top_joined}")
      appended = True
      success = True

    except WebDriverException as e:
      msg = str(e).lower()
      if any(x in msg for x in ["tab crashed", "chrome not reachable", "no such window"]):
        # 드라이버 재기동 + 재시도
        print("  -> 드라이버 크래시 감지: 재기동 후 재시도")
        try:
          driver.quit()
        except Exception:
          pass
        driver = start_driver()
        attempt += 1
        if attempt >= 2 and not appended:
          out_rows.append({'이메일 주소': '조회 중 오류'})
          appended = True
        continue
      else:
        if not appended:
          out_rows.append({'이메일 주소': '조회 중 오류'})
          appended = True
        break
    except Exception as e:
      if not appended:
        out_rows.append({'이메일 주소': '조회 중 오류'})
        appended = True
      break

  if not appended:
    # 방어적: 어떤 이유로도 미기록 상태면 오류로 채움
    out_rows.append({'이메일 주소': '조회 중 오류'})

# ------------------------------------------------------------
# 6) 드라이버 종료 & 저장
# ------------------------------------------------------------
try:
  driver.quit()
except Exception:
  pass

try:
  if "이메일 주소" not in df.columns:
    df["이메일 주소"] = "-"
  df["이메일 주소"] = [r["이메일 주소"] for r in out_rows]
  df.to_excel("clinic_result_filled.xlsx", index=False)
  print("엑셀 저장: clinic_result_filled.xlsx")
except Exception as e:
  print("엑셀 저장 중 오류:", e)
