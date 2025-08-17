import math
import requests
from typing import List, Dict, Tuple
from collections import deque
import pandas as pd

# =======================
# 설정값
# =======================
API_KEY = "AIzaSyArCS_QsBpm0TMybY-yriu6SLFb_wudibc"
URL = "https://places.googleapis.com/v1/places:searchNearby"

# "law": ["lawyer"] : 로펌/법률 사무소
# "estate": ["real_estate_agency"] : 부동산 중개
# "clinic": ["doctor", "dentist", "physiotherapist", "spa", "beauty_salon"] : 의료·미용 클리닉
# "creative": ["advertising_agency", "graphic_designer"] : 디자인/마케팅 에이전시
# "beauty": ["hair_care", "beauty_salon"] : 헤어살롱/뷰티숍
# "gallery": ["art_gallery", "art_studio"] : 갤러리/스튜디오
# "wealth": ["accounting", "bank", "insurance_agency"] : 금융/자산관리
includedTypes = ["marketing agency","digital marketing agency","performance marketing agency","advertising agency","media buying agency","seo agency","search marketing agency","social media agency","social media marketing agency","content marketing agency","influencer marketing agency","growth marketing agency","demand generation agency","brand consultancy","marketing consultancy","design agency","creative agency","creative studio","branding agency","brand design studio","brand consultancy","graphic design studio","packaging design agency","web design agency","website designer","UX UI design studio"]

headers = {
  "Content-Type": "application/json",
  "X-Goog-Api-Key": API_KEY,
  "X-Goog-FieldMask": "places.id,places.displayName,places.formattedAddress,places.postalAddress,places.location,places.primaryType,places.types,places.websiteUri"
}

# 최종 결과 리스트
RESULT_LIST = []

# 시작위치 및 반경 설정
START_LAT = 51.5055
START_LNG = -0.0865
START_RADIUS = 2000.0  # 전체 목표 반경

# r_inner = maxDistance + MARGIN_M
MARGIN_M = 10.0
# annulus 1차 타일링 파라미터
MAX_CELL_RADIUS = 800.0
OVERLAP_RATIO = 0.2

# 분할(세로 탐색) 파라미터
MIN_RADIUS_M = 120.0
MAX_DEPTH = 4
MAX_CALLS = 2000

# 기준 경도선(오른쪽만 수집/호출)
CUTOFF_LNG = -0.09038947216087369
SKIP_TILES_LEFT_OF_LINE = True  # True면 왼쪽 타일은 큐에 안 넣음

# =======================
# 유틸 함수
# =======================
def haversine_meters(lat1: float, lng1: float, lat2: float, lng2: float) -> float:
  R = 6371000.0
  to_rad = math.radians
  dlat = to_rad(lat2 - lat1)
  dlng = to_rad(lng2 - lng1)
  a = (math.sin(dlat / 2) ** 2
       + math.cos(to_rad(lat1)) * math.cos(to_rad(lat2)) * math.sin(dlng / 2) ** 2)
  c = 2 * math.asin(math.sqrt(a))
  return R * c

def offset_latlng(lat: float, lng: float, north_m: float = 0.0, east_m: float = 0.0) -> Tuple[float, float]:
  dlat = north_m / 111_320.0
  dlng = east_m / (111_320.0 * math.cos(math.radians(lat)))
  return lat + dlat, lng + dlng

def is_right_of_meridian(lng: float, cutoff_lng: float) -> bool:
  return lng > cutoff_lng

def build_ring_tiles_plan(
  lat0: float,
  lng0: float,
  R: float,
  max_distance: float,
  margin: float,
  max_cell_radius: float,
  overlap_ratio: float
) -> List[Dict]:
  """최상위 한 번만 사용: inner_cutoff~R 띠를 링으로 커버하는 원들을 생성"""
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
  """count==20일 때만 사용: 반경 r/2로 7개(중심+육각 방향) 분할"""
  r2 = radius_m / 2.0
  subtiles = []
  # 중심
  subtiles.append({"center": (center_lat, center_lng), "radius": r2, "depth": parent_depth + 1})
  # 육각 방향
  for deg in [0, 60, 120, 180, 240, 300]:
    rad = math.radians(deg)
    north = (radius_m / 2.0) * math.cos(rad)
    east = (radius_m / 2.0) * math.sin(rad)
    lat, lng = offset_latlng(center_lat, center_lng, north_m=north, east_m=east)
    subtiles.append({"center": (lat, lng), "radius": r2, "depth": parent_depth + 1})
  return subtiles

def nearby_once(center_lat: float, center_lng: float, radius: float) -> List[Dict]:
  """한 원을 조회하고 원시 places를 반환"""
  payload = {
    "languageCode": "en",
    "regionCode": "GB",
    "includedTypes": includedTypes,
    "maxResultCount": 20,
    "rankPreference": "DISTANCE",
    "locationRestriction": {
      "circle": {"center": {"latitude": center_lat, "longitude": center_lng}, "radius": radius}
    }
  }
  res = requests.post(URL, headers=headers, json=payload, timeout=30)
  res.raise_for_status()
  data = res.json()
  return data.get("places", [])

def search_nearby(center_lat: float, center_lng: float, radius: float):
  """결과 요약: 오른쪽만, 개수/최대거리"""
  places = nearby_once(center_lat, center_lng, radius)
  places_right = []
  dists = []
  for p in places:
    loc = p.get("location") or {}
    lng = loc.get("longitude")
    if lng is None or not is_right_of_meridian(lng, CUTOFF_LNG):
      continue
    lat = loc.get("latitude")
    if lat is not None:
      d = haversine_meters(center_lat, center_lng, lat, lng)
      p["distanceMeters"] = d
      dists.append(d)
    places_right.append(p)
  
  for p in places_right:
    name = p.get("displayName", {}).get("text")
    types = p.get("types")
    primaryType = p.get('primaryType')
    addr = p.get("formattedAddress")
    postalAddress = p.get("postalAddress")
    websiteUri = p.get("websiteUri")
    loc = p.get("location") or {}
    lng = loc.get("longitude")
    lat = loc.get("latitude")

    RESULT_LIST.append({
      "회사명": name,
      "업종": types,
      "기본 유형" : primaryType,
      "주소": addr,
      "우편주소" : postalAddress,
      "이메일 주소": "-",
      "전화번호": "-",
      "웹사이트 주소": websiteUri,
      "위도" : lat,
      "경도" : lng
    })

  count = len(places_right)
  max_dist = max(dists) if dists else 0.0
  return places_right, count, max_dist

# =======================
# 메인: 종료 조건이 명확한 BFS
# =======================
# 0) 파일럿: 시작 원 한 번 조회해서 inner_cutoff 계산 재료
pilot_places, pilot_count, pilot_maxdist = search_nearby(START_LAT, START_LNG, START_RADIUS)
print(f"[pilot] right_count={pilot_count}, maxDist={pilot_maxdist:.1f}m")

# 1) annulus 타일링은 '한 번만' 해서 큐 시드 생성
seed_plan = build_ring_tiles_plan(
  lat0=START_LAT, lng0=START_LNG, R=START_RADIUS,
  max_distance=pilot_maxdist, margin=MARGIN_M,
  max_cell_radius=MAX_CELL_RADIUS, overlap_ratio=OVERLAP_RATIO
)
# 오른쪽 타일만 큐에 추가(필요 시 False로 두고 결과만 필터)
queue = deque([t for t in seed_plan if is_right_of_meridian(t["center"][1], CUTOFF_LNG)])

visited = set()
calls = 1  # pilot에서 1회
results_by_id = {p.get("id"): p for p in pilot_places if p.get("id")}
processed = 0

while queue:
  print(f"[queue]\tremaining={len(queue)} processed={processed} calls={calls}")
  tile = queue.popleft()
  lat, lng = tile["center"]
  r = tile["radius"]
  depth = tile.get("depth", 0)

  # 중복 타일 방지
  key = (round(lat, 6), round(lng, 6), round(r, 1))
  if key in visited:
    continue
  visited.add(key)

  # 왼쪽 타일 건너뛰기(비용 절감)
  if SKIP_TILES_LEFT_OF_LINE and not is_right_of_meridian(lng, CUTOFF_LNG):
    continue

  # 조회
  places_right, count, maxdist = search_nearby(lat, lng, r)
  calls += 1
  processed += 1

  # 결과 합치기(디듀프)
  for p in places_right:
    pid = p.get("id")
    if pid and pid not in results_by_id:
      results_by_id[pid] = p

  print(f"[tile]\tdepth={depth} r={r:.1f}m center=({lat:.6f},{lng:.6f}) "
        f"→ count={count} maxDist={maxdist:.1f}m uniq_total={len(results_by_id)}")

  # **종료/분할 로직**
  if count == 20 and r > MIN_RADIUS_M and depth < MAX_DEPTH and calls < MAX_CALLS:
    # 꽉 찼으니 더 쪼갠다(세로 탐색)
    for child in split_circle_7(lat, lng, r, parent_depth=depth):
      # 오른쪽 타일만 큐에 추가(경계 누락이 걱정되면 조건 제거)
      if not SKIP_TILES_LEFT_OF_LINE or is_right_of_meridian(child["center"][1], CUTOFF_LNG):
        queue.append(child)
  # else: 20 미만 → 이 타일은 완료(자식 없음)

print("\n=== FINAL SUMMARY ===")
print(f"Unique places RIGHT of {CUTOFF_LNG}: {len(results_by_id)}")
print(f"HTTP calls (approx): {calls}")

# 엑셀 파일로 저장
# 1) 결과를 DF로 만들기
df = pd.DataFrame(RESULT_LIST)

# 2) 엑셀로 저장 (openpyxl)
df.to_excel("creative_result.xlsx", index=False, engine='openpyxl')