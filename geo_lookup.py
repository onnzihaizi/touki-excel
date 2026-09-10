import json
import math
import os
import re
import time
import urllib.parse
import urllib.request

CACHE_PATH = os.path.join(os.path.dirname(__file__), ".geo_cache.json")
USER_AGENT = "ToukiExcelFree/2.4 (local personal-use app)"


def _load_cache():
    try:
        with open(CACHE_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_cache(cache):
    try:
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _get_json(url, timeout=12):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _normalize_address_for_search(address: str) -> str:
    s = (address or "").strip()
    s = s.replace(" ", "").replace("　", "")
    # GSI tends to accept either registry-style 番地 or hyphen notation.
    return s


def geocode_japan(address: str, fallback_name: str = ""):
    """GSI address search. Returns {lat, lon, matched_address, query} or None."""
    cache = _load_cache()
    queries = [q for q in (_normalize_address_for_search(address), (fallback_name or "").strip()) if q]
    for q in queries:
        key = "geo:" + q
        if key in cache:
            return cache[key]
        url = "https://msearch.gsi.go.jp/address-search/AddressSearch?" + urllib.parse.urlencode({"q": q})
        try:
            data = _get_json(url)
            if data:
                item = data[0]
                coords = item.get("geometry", {}).get("coordinates", [])
                if len(coords) >= 2:
                    result = {
                        "lon": float(coords[0]),
                        "lat": float(coords[1]),
                        "matched_address": item.get("properties", {}).get("title", ""),
                        "query": q,
                    }
                    cache[key] = result
                    _save_cache(cache)
                    return result
        except Exception:
            continue
    return None


def _distance_to_meters(distance_text: str):
    """HeartRails distance strings are typically '640m' / '1.2km'."""
    s = (distance_text or "").strip().lower().replace(" ", "")
    m = re.search(r"([0-9.]+)km", s)
    if m:
        return float(m.group(1)) * 1000
    m = re.search(r"([0-9.]+)m", s)
    if m:
        return float(m.group(1))
    try:
        return float(s)
    except Exception:
        return None


def nearest_station(lon: float, lat: float):
    """HeartRails Express nearest-station API. Free, no key. Returns nearest result."""
    cache = _load_cache()
    key = f"station:{lon:.6f},{lat:.6f}"
    if key in cache:
        return cache[key]
    params = {"method": "getStations", "x": f"{lon:.8f}", "y": f"{lat:.8f}"}
    url = "https://express.heartrails.com/api/json?" + urllib.parse.urlencode(params)
    try:
        data = _get_json(url)
        stations = data.get("response", {}).get("station", [])
        if not stations:
            return None
        # API response is nearest-first, but explicitly sort by parsed distance when possible.
        parsed = []
        for st in stations:
            meters = _distance_to_meters(st.get("distance", ""))
            parsed.append((meters if meters is not None else 10**12, st, meters))
        parsed.sort(key=lambda x: x[0])
        _, st, meters = parsed[0]
        # Real-estate convention: walking minutes are rounded up at 80m/min.
        # HeartRails distance is not guaranteed to be actual pedestrian-route distance,
        # so this is intentionally presented as an estimate in the UI.
        walk_minutes = math.ceil(meters / 80) if meters is not None else None
        result = {
            "name": st.get("name", ""),
            "line": st.get("line", ""),
            "distance_text": st.get("distance", ""),
            "distance_m": meters,
            "walk_minutes": walk_minutes,
        }
        cache[key] = result
        _save_cache(cache)
        return result
    except Exception:
        return None


def enrich_location(property_address: str, property_name: str = "") -> dict:
    geo = geocode_japan(property_address, property_name)
    if not geo:
        return {
            "geocoded_address": "", "station_name": "", "station_line": "",
            "station_distance_m": None, "walk_minutes": None,
            "warning": "住所の位置検索に失敗しました。住所を確認して手入力してください。",
        }
    st = nearest_station(geo["lon"], geo["lat"])
    if not st:
        return {
            "geocoded_address": geo.get("matched_address", ""), "station_name": "", "station_line": "",
            "station_distance_m": None, "walk_minutes": None,
            "warning": "最寄駅検索に失敗しました。インターネット接続を確認するか手入力してください。",
        }
    return {
        "geocoded_address": geo.get("matched_address", ""),
        "station_name": st.get("name", ""),
        "station_line": st.get("line", ""),
        "station_distance_m": st.get("distance_m"),
        "walk_minutes": st.get("walk_minutes"),
        "warning": "",
    }
