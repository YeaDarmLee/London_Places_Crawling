# retry_error_only.py
# -*- coding: utf-8 -*-
import time, subprocess, re, html, random
import pandas as pd
from urllib.parse import urlparse, urljoin, urlsplit, parse_qs, unquote
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

INPUT_XLSX = "wealth_result_filled_retry.xlsx"         # 기존 결과 파일
OUTPUT_XLSX = "wealth_result_filled_retry_retry.xlsx"  # 리트라이 결과 저장
TARGET_COL_URL = "웹사이트 주소"
TARGET_COL_EMAIL = "이메일 주소"
RETRY_LABEL = "조회 중 오류"

EMAIL_REGEX = re.compile(r'\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,24}\b', re.IGNORECASE)
BAD_TLDS = {
  "css","js","map","json","png","jpg","jpeg","gif","webp","svg","ico",
  "woff","woff2","ttf","otf","mp4","webm","mov","avi","pdf","zip","rar","7z","gz","tar","xml","html","htm"
}

def link_weight(txt_lower, href_lower):
  if "contact" in txt_lower or "contact" in href_lower: return 3
  if "about"   in txt_lower or "about"   in href_lower: return 2
  if "support" in txt_lower or "support" in href_lower: return 1
  return 0

def is_valid_email(e: str) -> bool:
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

def base_domain(u: str) -> str:
  host = urlparse(u).netloc.lower()
  parts = host.split(".")
  if len(parts) >= 3 and parts[-2:] == ["co", "uk"]:
    return ".".join(parts[-3:])
  return ".".join(parts[-2:]) if len(parts) >= 2 else host

def score_and_pick(emails: set, u: str, k: int = 3) -> str:
  bd = base_domain(u)
  score_map = {
    "info": 4, "hello": 4, "contact": 4, "support": 3, "help": 3,
    "sales": 3, "admin": 2, "team": 2, "office": 2, "enquiries": 2
  }
  ranked = []
  for e in sorted(emails):
    local, _, dom = e.partition("@")
    s = 0
    if bd and bd in dom.lower():
      s -= 5
    s -= score_map.get(local.lower(), 0)
    ranked.append((s, e))
  ranked.sort()
  return ", ".join([e for _, e in ranked[:k]]) if ranked else RETRY_LABEL

def deobfuscate_and_extract(text: str) -> set:
  out = set()
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
    if is_valid_email(e):
      out.add(e)
  return out

def get_visible_text(html_str: str) -> str:
  soup = BeautifulSoup(html_str or "", "lxml")
  for tag in soup(["script","style","noscript","template"]):
    tag.decompose()
  # aria-label 힌트도 텍스트로 포함
  for el in soup.select('[aria-label]'):
    try:
      el.append(soup.new_string(' ' + el['aria-label']))
    except Exception:
      pass
  return soup.get_text(" ", strip=True)

def collect_from_mailto(driver, emails: set):
  anchors = driver.find_elements(By.CSS_SELECTOR, 'a[href^="mailto:"]')
  for a in anchors:
    href = (a.get_attribute("href") or "").strip()
    if not href: continue
    base = href.split("?", 1)[0].replace("mailto:", "").strip()
    base = unquote(base)
    if is_valid_email(base):
      emails.add(base)
    qs = urlsplit(href).query
    if not qs: continue
    qd = parse_qs(qs)
    for key in ("to","cc","bcc"):
      for val in qd.get(key, []):
        for e in unquote(val).split(","):
          e = e.strip()
          if is_valid_email(e):
            emails.add(e)

def find_top_emails(url: str, driver) -> str:
  try:
    driver.get(url)
    WebDriverWait(driver, 12).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    WebDriverWait(driver, 8).until(lambda d: d.execute_script('return document.readyState') in ('interactive','complete'))
    time.sleep(0.4)
  except Exception:
    return RETRY_LABEL

  all_candidates = set()

  # 1) mailto:
  try:
    collect_from_mailto(driver, all_candidates)
  except Exception:
    pass

  # 2) 본문 가시 텍스트
  try:
    text = get_visible_text(driver.page_source or "")
    all_candidates |= deobfuscate_and_extract(text)
  except Exception:
    pass

  # 3) 내부 contact/about/support 상위 3개 링크 진입
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
      w = link_weight(txt, href2.lower())
      if w > 0:
        scored.append((w, len(href2), href2))
    scored.sort(key=lambda x: (-x[0], x[1], x[2]))
    cand_links = [h for _,_,h in scored[:3]]

    for link in cand_links:
      try:
        driver.get(link)
        WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(0.3)
        collect_from_mailto(driver, all_candidates)
        text2 = get_visible_text(driver.page_source or "")
        all_candidates |= deobfuscate_and_extract(text2)
        time.sleep(random.uniform(0.15, 0.35))
      except Exception:
        continue
  except Exception:
    pass

  return score_and_pick(all_candidates, url, k=3)

def make_driver():
  # 기존 크롬 프로필 + 디버깅 포트에 붙고 싶다면 주석 해제
  # subprocess.Popen('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\chromeCookie\\kmong_Rohmin_leisure"')
  options = Options()
  options.add_argument("--no-sandbox")
  options.add_argument("--disable-dev-shm-usage")
  options.add_argument("--ignore-certificate-errors")
  options.add_argument("--disable-blink-features=AutomationControlled")
  options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/126.0.0.0 Safari/537.36")
  # 디버깅 세션에 붙을 때는 아래 줄 활성화
  # options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
  options.page_load_strategy = 'eager'

  service = Service(ChromeDriverManager().install())
  driver = webdriver.Chrome(service=service, options=options)
  driver.set_page_load_timeout(20)

  try:
    driver.execute_cdp_cmd("Network.enable", {})
    driver.execute_cdp_cmd("Network.setBlockedURLs", {
      "urls": [
        "*.png","*.jpg","*.jpeg","*.gif","*.webp","*.svg","*.ico",
        "*.mp4","*.avi","*.webm","*.mov",
        "*.woff","*.woff2","*.ttf","*.otf",
      ]
    })
  except Exception:
    pass
  return driver

def sanitize(v):
  if isinstance(v, str):
    return re.sub(r'[\x00-\x1F\x7F]', '', v)
  return v

def main():
  df = pd.read_excel(INPUT_XLSX)
  if TARGET_COL_URL not in df.columns or TARGET_COL_EMAIL not in df.columns:
    raise RuntimeError(f"엑셀에 '{TARGET_COL_URL}' 또는 '{TARGET_COL_EMAIL}' 컬럼이 없습니다.")

  mask = (df[TARGET_COL_EMAIL].astype(str).str.strip() == RETRY_LABEL)
  error_idxs = list(df.index[mask])

  if not error_idxs:
    print("리트라이 대상(조회 중 오류) 행이 없습니다.")
    return

  print(f"리트라이 대상 행 수: {len(error_idxs)}")

  driver = make_driver()
  processed = 0
  RESTART_EVERY = 80
  CHECKPOINT_EVERY = 50

  try:
    for i in error_idxs:
      url = str(df.at[i, TARGET_COL_URL]).strip()
      if not url or url.lower() in ("nan", "none", "-"):
        df.at[i, TARGET_COL_EMAIL] = RETRY_LABEL
        continue
      print(f"[{processed+1}/{len(error_idxs)}] 재조회: {url}")

      try:
        result = find_top_emails(url, driver)
      except Exception:
        result = RETRY_LABEL

      df.at[i, TARGET_COL_EMAIL] = sanitize(result)
      processed += 1

      # 체크포인트 저장
      if processed % CHECKPOINT_EVERY == 0:
        df.to_excel(OUTPUT_XLSX, index=False)
        print(f"체크포인트 저장: {OUTPUT_XLSX}")

      # 장시간 안정성 위해 주기적 재기동
      if processed % RESTART_EVERY == 0:
        try:
          driver.quit()
        except Exception:
          pass
        time.sleep(1.0)
        driver = make_driver()

      time.sleep(random.uniform(0.2, 0.5))  # 매너 슬립
  finally:
    try:
      driver.quit()
    except Exception:
      pass

  df.to_excel(OUTPUT_XLSX, index=False)
  print(f"리트라이 완료: {OUTPUT_XLSX}")

if __name__ == "__main__":
  main()
