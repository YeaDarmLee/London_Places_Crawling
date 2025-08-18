# text_api_search_pandas.py
import math
import requests
import pandas as pd
from typing import Dict, List, Tuple, Optional

API_KEY = "YOUR_API_KEY"   # ← 교체
URL_TEXT = "https://places.googleapis.com/v1/places:searchText"

# 검색 설정 ==============================
LANG = "en"
REGION = "GB"
CENTER_LAT = 51.5055
CENTER_LNG = -0.0865
SEARCH_RADIUS_M = 3000.0
MAX_PAGES_PER_QUERY = 10   # 페이지 한도(안전장치). 모두 끝까지면 None

CATEGORY_LABEL = "디자인/마케팅 에이전시"

USE_EAST_OF_LONGITUDE = True
CUTOFF_LNG = -0.09038947216087369

CREATIVE_QUERIES = [
  "design agency",
  "digital marketing agency",
  "branding agency",
  "creative agency",
  "advertising agency",
  "graphic design studio",
  "web design agency",
  "SEO agency",
  "content marketing agency",
  "social media marketing agency",
  "UX UI agency",
  "performance marketing agency",
]

# === FieldMask: 필요한 것만 지정해야 실제 응답에 포함됨 ===
FIELD_MASK = ",".join([
  "places.id",
  "places.displayName.text",
  "places.formattedAddress",
  "places.location",
  "places.types",
  "places.primaryType",
  "places.postalAddress",     # postalAddress.addressLines/locality/administrativeArea/regionCode/postalCode
  "places.addressComponents",   # component fallback (street_number, route, postal_code 등)
  "places.websiteUri",
  "places.nationalPhoneNumber",
  "places.internationalPhoneNumber",
  "nextPageToken",
])

HEADERS = {
  "Content-Type": "application/json",
  "X-Goog-Api-Key": API_KEY,
  "X-Goog-FieldMask": FIELD_MASK,
}

# ===== 유틸 =====
def offset_latlng(lat: float, lng: float, north_m: float = 0.0, east_m: float = 0.0) -> Tuple[float, float]:
  dlat = north_m / 111_320.0
  dlng = east_m / (111_320.0 * math.cos(math.radians(lat)))
  return lat + dlat, lng + dlng

def make_restriction_rectangle(lat: float, lng: float, radius_m: float) -> Dict:
  north_lat, _ = offset_latlng(lat, lng, north_m=+radius_m, east_m=0)
  south_lat, _ = offset_latlng(lat, lng, north_m=-radius_m, east_m=0)
  _, east_lng = offset_latlng(lat, lng, north_m=0, east_m=+radius_m)
  _, west_lng = offset_latlng(lat, lng, north_m=0, east_m=-radius_m)
  low = {"latitude": min(south_lat, north_lat), "longitude": min(west_lng, east_lng)}
  high = {"latitude": max(south_lat, north_lat), "longitude": max(west_lng, east_lng)}
  return {"rectangle": {"low": low, "high": high}}

def is_right_of_meridian(lng: float, cutoff_lng: float) -> bool:
  return lng > cutoff_lng

# ===== 주소/우편주소 생성 =====
def extract_postal_code_from_components(components: List[Dict]) -> Optional[str]:
  for c in components or []:
    types = c.get("types", [])
    if "postal_code" in types:
      return c.get("longText") or c.get("shortText")
  return None

def extract_street_from_components(components: List[Dict]) -> Optional[str]:
  street_number = None
  route = None
  for c in components or []:
    types = set(c.get("types", []))
    val = c.get("longText") or c.get("shortText")
    if "street_number" in types and val:
      street_number = val
    if ("route" in types or "street_address" in types) and val:
      route = val
  if street_number or route:
    return " ".join([t for t in [street_number, route] if t])
  return None

def build_address_fields(place: Dict) -> Tuple[str, str, Optional[str]]:
  """
  반환: (주소, 우편주소, 우편번호)
  - 주소: 거리/번지(line) 중심. postalAddress.addressLines 우선, 없으면 addressComponents에서 구성,
      최후에는 formattedAddress.
  - 우편주소: 거리 + 도시/주 + 우편번호 + 국가를 조합한 완전 주소(가능한 경우),
        없으면 formattedAddress.
  """
  formatted = place.get("formattedAddress") or ""
  pa = place.get("postalAddress") or {}
  comps = place.get("addressComponents") or []

  # 1) street line (주소)
  addr_lines = pa.get("addressLines") or []
  street_line = ", ".join(addr_lines) if addr_lines else extract_street_from_components(comps)
  if not street_line:
    # 마지막 수단: formattedAddress에서 첫 콤마 전까지 추출(영국 주소가 일관적이지 않을 수 있어 보수적으로)
    street_line = formatted.split(",")[0].strip() if formatted else ""

  # 2) postal code (우편번호)
  postal_code = pa.get("postalCode") or extract_postal_code_from_components(comps)

  # 3) city/state/country
  locality = pa.get("locality")
  admin = pa.get("administrativeArea")
  country = pa.get("regionCode")

  # 우편주소 문자열 구성(가능한 한 구조화)
  parts = []
  if street_line: parts.append(street_line)
  city_state = ", ".join([p for p in [locality, admin] if p])
  if city_state: parts.append(city_state)
  if postal_code: parts.append(postal_code)
  if country: parts.append(country)
  postal_addr = ", ".join(parts) if parts else formatted

  return street_line or formatted, postal_addr or formatted, postal_code

# ===== API =====
def text_search_once(query: str, lat: float, lng: float, radius_m: float, page_token: Optional[str] = None):
  body = {
    "textQuery": query,
    "languageCode": LANG,
    "regionCode": REGION,
    "locationRestriction": make_restriction_rectangle(lat, lng, radius_m),
    "pageSize": 20,
  }
  if page_token:
    body["pageToken"] = page_token

  r = requests.post(URL_TEXT, headers=HEADERS, json=body, timeout=30)
  if r.status_code >= 400:
    try:
      print(f"[HTTP {r.status_code}] {r.text[:500]}")
    finally:
      r.raise_for_status()
  data = r.json()
  return data.get("places", []), data.get("nextPageToken")

# ===== 실행 & 저장 =====
def run_text_search_to_excel(output_path: str):
  seen = set()
  rows = []

  for q in CREATIVE_QUERIES:
    token = None
    pages = 0
    while True:
      places, token = text_search_once(q, CENTER_LAT, CENTER_LNG, SEARCH_RADIUS_M, token)

      for p in places:
        pid = p.get("id")
        if not pid or pid in seen:
          continue

        loc = p.get("location") or {}
        plat = loc.get("latitude")
        plng = loc.get("longitude")
        if USE_EAST_OF_LONGITUDE and isinstance(plng, (int, float)) and not is_right_of_meridian(plng, CUTOFF_LNG):
          continue

        seen.add(pid)

        display_name = (p.get("displayName") or {}).get("text") or ""
        website = p.get("websiteUri") or ""
        phone = p.get("nationalPhoneNumber") or p.get("internationalPhoneNumber") or ""
        primary_type = p.get("primaryType") or ""

        addr_line, postal_addr, postal_code = build_address_fields(p)

        row = {
          "회사명": display_name,
          "업종": CATEGORY_LABEL,
          "기본 유형": primary_type,
          "주소": addr_line,     # 도로/번지 중심
          "우편주소": postal_addr,  # 완전 우편주소(도시/주/우편번호/국가 포함)
          "우편번호": postal_code or "",
          "이메일 주소": "",      # Places는 이메일 미제공(웹 크롤링 필요)
          "전화번호": phone,
          "웹사이트 주소": website,
          "위도": plat,
          "경도": plng,
        }
        rows.append(row)

      pages += 1
      if not token or (MAX_PAGES_PER_QUERY is not None and pages >= MAX_PAGES_PER_QUERY):
        break

  df = pd.DataFrame(rows, columns=[
    "회사명", "업종", "기본 유형", "주소", "우편주소", "우편번호",
    "이메일 주소", "전화번호", "웹사이트 주소", "위도", "경도"
  ]).fillna("")
  df.to_excel(output_path, index=False)
  print(f"저장 완료: {output_path} (총 {len(df)}건)")

if __name__ == "__main__":
  run_text_search_to_excel("text_results_creative.xlsx")
