# text_api_search.py
# ------------------
# Google Places API (New) - Text Search 전용 수집기
# - locationRestriction.rectangle 사용 (원형은 Bias에서만 지원)
# - nextPageToken 끝까지 추적 (MAX_TEXT_PAGES_PER_QUERY=None)
# - 반경 밖 결과 이중 필터
# - 경도 기준선 오른쪽만 결과로 남기는 옵션
# - 엑셀(.xlsx) 저장

import math
import time
from collections import deque
from typing import List, Dict, Tuple, Optional

import requests
from openpyxl import Workbook

# =======================
# 기본 설정
# =======================
API_KEY = "AIzaSyArCS_QsBpm0TMybY-yriu6SLFb_wudibc"
TEXT_URL = "https://places.googleapis.com/v1/places:searchText"

HEADERS = {
  "Content-Type": "application/json",
  "X-Goog-Api-Key": API_KEY,
  # nextPageToken 포함(페이지네이션용)
  "X-Goog-FieldMask": (
    "places.id,places.displayName,places.location,places.formattedAddress,"
    "places.websiteUri,places.types,nextPageToken"
  )
}

# 시작 위치(중심) 및 파일럿/전체 반경
START_LAT = 51.5055
START_LNG = -0.0865
PILOT_RADIUS_M = 1000.0   # 중심부 1회 수집 반경(파일럿)
BIG_RADIUS_M   = 3500.0   # 전체 목표 커버 반경

# 타일 계획(annulus)
MARGIN_M = 10.0       # inner cutoff = pilot_maxdist + margin
MAX_CELL_RADIUS = 400.0   # 타일 반경(작을수록 누락↓ / 호출수↑)
OVERLAP_RATIO   = 0.2     # 인접 타일 겹침 비율

# 세분화(분할) 파라미터
MIN_RADIUS_M = 120.0
MAX_DEPTH = 4
SPLIT_COUNT_THRESHOLD = 200   # 한 타일 결과가 이 이상이면 강제 분할
MAX_TOTAL_CALLS = 10000     # 전체 API 호출 상한(세이프가드)

# 텍스트 쿼리 버킷(디자인/마케팅 예시 — 필요 시 확장)
CREATIVE_QUERIES = [
  "marketing agency", "digital marketing agency", "advertising agency",
  "creative agency", "branding agency", "web design agency", "seo agency"
]

# 페이지네이션: None이면 끝까지, 숫자면 해당 페이지 수까지만
MAX_TEXT_PAGES_PER_QUERY: Optional[int] = None

# 위치 제한/필터
USE_LOCATION_RESTRICTION = True   # True → rectangle로 반경 강제 제한
FILTER_BY_RADIUS = True       # 반경 밖 결과를 거리로 한 번 더 버림(안전)
# 경도 기준선 오른쪽만 결과로 남길지(커버리지를 위해 타일은 전체 돌리고 결과에서만 필터)
FILTER_RESULTS_TO_RIGHT_ONLY = True
CUTOFF_LNG = -0.09038947216087369

# 저장(엑셀)
XLSX_PATH = "text_results_creative.xlsx"


# =======================
# 유틸
# =======================
def haversine_meters(lat1: float, lng1: float, lat2: float, lng2: float) -> float:
  R = 6371000.0
  dlat = math.radians(lat2 - lat1)
  dlng = math.radians(lng2 - lng1)
  a = (math.sin(dlat / 2) ** 2
     + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlng / 2) ** 2)
  c = 2 * math.asin(math.sqrt(a))
  return R * c

def offset_latlng(lat: float, lng: float, north_m: float = 0.0, east_m: float = 0.0) -> Tuple[float, float]:
  dlat = north_m / 111_320.0
  dlng = east_m / (111_320.0 * math.cos(math.radians(lat)))
  return lat + dlat, lng + dlng

def is_right_of_meridian(lng: float, cutoff_lng: float) -> bool:
  # 서경(음수)에서 값이 "더 크면" 동쪽
  return lng > cutoff_lng

def make_viewport_rectangle(lat: float, lng: float, radius_m: float) -> Dict:
  # 반경 r 을 포함하는 사각형(뷰포트) 구성 (SW=low, NE=high)
  south, west = offset_latlng(lat, lng, north_m=-radius_m, east_m=-radius_m)
  north, east = offset_latlng(lat, lng, north_m=+radius_m, east_m=+radius_m)
  return {
    "rectangle": {
      "low":  {"latitude": south, "longitude": west},   # SW
      "high": {"latitude": north, "longitude": east}  # NE
    }
  }

def build_ring_tiles_plan(
  lat0: float,
  lng0: float,
  R: float,
  max_distance: float,
  margin: float,
  max_cell_radius: float,
  overlap_ratio: float
) -> List[Dict]:
  """
  inner_cutoff = max_distance + margin 부터 R까지의 띠(annulus)를
  반지름 r_tile짜리 원 타일들로 커버하는 계획 생성.
  """
  r_inner = max_distance + margin
  w = R - r_inner
  if w <= 0:
    return []

  n_rings = max(1, math.ceil(w / max_cell_radius))
  ring_thickness = w / n_rings

  tiles: List[Dict] = []
  for i in range(n_rings):
    r_tile = min(ring_thickness, max_cell_radius)
    r_center = r_inner + (i + 0.5) * ring_thickness

    L = 2 * math.pi * r_center
    spacing = max(1.0, 2 * r_tile * (1 - overlap_ratio))
    n = max(1, math.ceil(L / spacing))

    for k in range(n):
      theta = (360.0 * k) / n
      rad = math.radians(theta)
      north = r_center * math.cos(rad)
      east = r_center * math.sin(rad)
      lat, lng = offset_latlng(lat0, lng0, north_m=north, east_m=east)
      tiles.append({
        "center": (lat, lng),
        "radius": r_tile,
        "ring_index": i,
        "bearing_deg": theta,
        "depth": 0
      })
  return tiles

def split_circle_7(center_lat: float, center_lng: float, radius_m: float, parent_depth: int) -> List[Dict]:
  """r/2로 7분할(중심 + 육각)"""
  r2 = radius_m / 2.0
  out = [{"center": (center_lat, center_lng), "radius": r2, "depth": parent_depth + 1}]
  for deg in [0, 60, 120, 180, 240, 300]:
    rad = math.radians(deg)
    north = (radius_m / 2.0) * math.cos(rad)
    east = (radius_m / 2.0) * math.sin(rad)
    lat, lng = offset_latlng(center_lat, center_lng, north, east)
    out.append({"center": (lat, lng), "radius": r2, "depth": parent_depth + 1})
  return out


# =======================
# Text Search
# =======================
def text_search_once(lat: float, lng: float, radius_m: float, query: str, page_token: Optional[str] = None):
  # rectangle로 강제 제한 (Text Search의 locationRestriction은 rectangle만 허용)
  payload = {
    "textQuery": query,
    "locationRestriction": make_viewport_rectangle(lat, lng, radius_m),
    "pageSize": 20
  }
  if page_token:
    payload["pageToken"] = page_token

  res = requests.post(TEXT_URL, headers=HEADERS, json=payload, timeout=30)
  res.raise_for_status()
  data = res.json()
  return data.get("places", []), data.get("nextPageToken")

def search_text_tile(lat: float, lng: float, radius_m: float, queries: List[str], max_pages_per_query: Optional[int]):
  """
  한 타일 수집: 여러 쿼리 × 페이지네이션(끝까지 혹은 한도까지)
  반환: (places_unique, count, max_dist, saturated, calls_used)
    - saturated: max_pages_per_query 제한으로 '더 남은' 상태에서 중단됐는지
  """
  by_id: Dict[str, Dict] = {}
  saturated = False
  calls_used = 0

  for q in queries:
    token = None
    pages = 0
    while True:
      places, token = text_search_once(lat, lng, radius_m, q, token)
      calls_used += 1

      for p in places:
        loc = p.get("location") or {}
        lon = loc.get("longitude")
        la  = loc.get("latitude")
        if lon is None or la is None:
          continue

        # 경도 기준선: 결과에서만 필터
        if FILTER_RESULTS_TO_RIGHT_ONLY and not is_right_of_meridian(lon, CUTOFF_LNG):
          continue

        # 반경 이중 필터(원 밖 노이즈 제거)
        d = haversine_meters(lat, lng, la, lon)
        if FILTER_BY_RADIUS and d > radius_m:
          continue

        pid = p.get("id")
        if not pid or pid in by_id:
          continue
        p["distanceMeters"] = d
        by_id[pid] = p

      pages += 1
      if not token:
        break
      if (max_pages_per_query is not None) and (pages >= max_pages_per_query):
        saturated = True
        break

  dists = [p.get("distanceMeters") for p in by_id.values() if isinstance(p.get("distanceMeters"), (int, float))]
  max_dist = max(dists) if dists else 0.0
  return list(by_id.values()), len(by_id), max_dist, saturated, calls_used


# =======================
# 저장(엑셀)
# =======================
def save_to_excel(items: List[Dict], path: str):
  cols = ["id", "name", "address", "lat", "lng", "distance_m", "website", "types", "gmaps_url"]

  wb = Workbook()
  ws = wb.active
  ws.title = "results"
  ws.append(cols)

  if not items:
    wb.save(path)
    return

  for p in items:
    loc = p.get("location") or {}
    lat = loc.get("latitude")
    lng = loc.get("longitude")
    gmaps = f"https://www.google.com/maps?q={lat},{lng}" if (lat is not None and lng is not None) else None
    ws.append([
      p.get("id"),
      (p.get("displayName") or {}).get("text"),
      p.get("formattedAddress"),
      lat,
      lng,
      p.get("distanceMeters"),
      p.get("websiteUri"),
      ",".join(p.get("types") or []),
      gmaps
    ])
  wb.save(path)


# =======================
# 메인
# =======================
def main():
  # 0) 파일럿: 중심 원 한 번 수집(내부 컷오프 계산용)
  pilot_places, pilot_count, pilot_maxdist, pilot_sat, pilot_calls = search_text_tile(
    START_LAT, START_LNG, PILOT_RADIUS_M, CREATIVE_QUERIES, MAX_TEXT_PAGES_PER_QUERY
  )
  print(f"[pilot] right_count={pilot_count}, maxDist={pilot_maxdist:.1f}m, calls={pilot_calls}, saturated={pilot_sat}")

  # 1) 바깥 띠(annulus) 타일 계획(전체 반경 기준)
  seed_plan = build_ring_tiles_plan(
    lat0=START_LAT, lng0=START_LNG, R=BIG_RADIUS_M,
    max_distance=pilot_maxdist, margin=MARGIN_M,
    max_cell_radius=MAX_CELL_RADIUS, overlap_ratio=OVERLAP_RATIO
  )
  print(f"[seed] tiles={len(seed_plan)}")

  # 2) BFS 큐
  queue = deque(seed_plan)
  visited = set()
  total_calls = pilot_calls
  processed = 0

  # 결과(파일럿 포함) 디듀프
  results_by_id: Dict[str, Dict] = {p.get("id"): p for p in pilot_places if p.get("id")}

  while queue and total_calls < MAX_TOTAL_CALLS:
    print(f"[queue] remaining={len(queue)} processed={processed} calls={total_calls}")
    tile = queue.popleft()
    lat, lng = tile["center"]
    r = tile["radius"]
    depth = tile.get("depth", 0)

    key = (round(lat, 6), round(lng, 6), round(r, 1))
    if key in visited:
      continue
    visited.add(key)

    places_right, count, maxdist, saturated, calls_used = search_text_tile(
      lat, lng, r, CREATIVE_QUERIES, MAX_TEXT_PAGES_PER_QUERY
    )
    total_calls += calls_used
    processed += 1

    for p in places_right:
      pid = p.get("id")
      if pid and pid not in results_by_id:
        results_by_id[pid] = p

    print(f"[tile] depth={depth} r={r:.1f} center=({lat:.6f},{lng:.6f}) "
        f"-> count={count} maxDist={maxdist:.1f}m saturated={saturated} uniq_total={len(results_by_id)}")

    # 분할 조건: 페이지 한도로 끊겼거나 OR 결과가 매우 많을 때
    if (saturated or count >= SPLIT_COUNT_THRESHOLD) and r > MIN_RADIUS_M and depth < MAX_DEPTH:
      for child in split_circle_7(lat, lng, r, parent_depth=depth):
        queue.append(child)

    time.sleep(0.03)  # QPS 여유

  print("\n=== FINAL SUMMARY ===")
  print(f"Unique places (right side only={FILTER_RESULTS_TO_RIGHT_ONLY}): {len(results_by_id)}")
  print(f"HTTP calls (approx): {total_calls}")
  print(f"Visited tiles: {len(visited)}")

  # 3) 저장
  items = list(results_by_id.values())
  save_to_excel(items, XLSX_PATH)
  print(f"Saved Excel -> {XLSX_PATH}")


if __name__ == "__main__":
  try:
    main()
  except requests.HTTPError as e:
    resp = e.response
    if resp is not None:
      print(f"[HTTP {resp.status_code}] {resp.text[:500]}")
    raise
