import html
import json
import os
import re
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from collections import defaultdict

CACHE_PATH = os.path.join(os.path.dirname(__file__), ".building_cache_v370.json")
UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152 Safari/537.36"

# Ver.3.7: 特定サイトの本文スクレイピングに依存しない。
# 検索結果スニペットを複数サイトから集め、同じ情報が複数ソースで一致した候補を優先する。
SOURCE_WEIGHTS = {
    "homes.co.jp": 6,
    "suumo.jp": 5,
    "athome.co.jp": 5,
    "mansion-review.jp": 4,
    "nomu.com": 4,
    "livable.co.jp": 4,
    "smtrc.jp": 4,
    "rehouse.co.jp": 4,
    "mansion-market.com": 3,
    "ieshil.com": 3,
    "home4u.jp": 3,
}


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


def _get_text(url, timeout=12):
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": UA,
            "Accept-Language": "ja-JP,ja;q=0.9,en;q=0.6",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read()
        charset = resp.headers.get_content_charset() or "utf-8"
        return raw.decode(charset, errors="ignore")


def _strip_tags(s):
    s = re.sub(r"<script\b.*?</script>|<style\b.*?</style>", " ", s or "", flags=re.I | re.S)
    s = re.sub(r"</(?:tr|td|th|li|div|p|dt|dd|h[1-6])>", " | ", s, flags=re.I)
    s = re.sub(r"<br\s*/?>", " | ", s, flags=re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = html.unescape(s)
    return re.sub(r"\s+", " ", s).strip()


def _normalize_digits(s):
    return (s or "").translate(str.maketrans("０１２３４５６７８９，", "0123456789,"))


def _normalize_name(s):
    s = html.unescape(s or "")
    s = s.replace("　", "").replace(" ", "")
    s = s.replace("‐", "-").replace("－", "-").replace("―", "-").replace("−", "-")
    s = re.sub(r"[・･\.．,，'\"「」『』()（）\[\]【】]", "", s)
    return s.lower()


def _item_mentions_property(text, property_name):
    n = _normalize_name(property_name)
    t = _normalize_name(text)
    if not n:
        return False
    if n in t:
        return True
    core = re.sub(r"(?:壱番館|一番館|弐番館|二番館|参番館|三番館|[0-9]+号棟)$", "", n)
    return len(core) >= 6 and core in t


def _municipality_hint(address):
    s = (address or "").replace("　", "").replace(" ", "")
    m = re.search(r"(?:東京都)?([^0-9０-９]{1,18}区)", s)
    if m:
        return m.group(1)
    m = re.search(r"(?:.{2,4}[都道府県])?([^0-9０-９]{1,18}(?:市|町|村))", s)
    return m.group(1) if m else ""


def _source_domain(url):
    try:
        host = urllib.parse.urlparse(url or "").netloc.lower()
    except Exception:
        host = ""
    host = host.split(":")[0]
    if host.startswith("www."):
        host = host[4:]
    for d in SOURCE_WEIGHTS:
        if host == d or host.endswith("." + d):
            return d
    return host or "search-snippet"


def _source_weight(url):
    return SOURCE_WEIGHTS.get(_source_domain(url), 2)


def _search_bing_rss(query):
    url = "https://www.bing.com/search?" + urllib.parse.urlencode({"q": query, "format": "rss", "setlang": "ja-JP"})
    xml_text = _get_text(url)
    root = ET.fromstring(xml_text)
    items = []
    for item in root.findall(".//item"):
        title = item.findtext("title") or ""
        desc = item.findtext("description") or ""
        link = item.findtext("link") or ""
        if "t23m-navi.jp" in link.lower() or "t23m-navi.jp" in (title + desc).lower():
            continue
        items.append({"title": _strip_tags(title), "description": _strip_tags(desc), "url": link})
    return items


def _search_html(query):
    # RSSが使えない環境用の保険。検索結果本文全体を1スニペットとして返す。
    urls = [
        "https://www.google.com/search?" + urllib.parse.urlencode({"q": query, "hl": "ja"}),
        "https://www.bing.com/search?" + urllib.parse.urlencode({"q": query, "setlang": "ja-JP"}),
        "https://html.duckduckgo.com/html/?" + urllib.parse.urlencode({"q": query}),
    ]
    for url in urls:
        try:
            text = _strip_tags(_get_text(url))
            if text:
                return [{"title": "", "description": text, "url": ""}]
        except Exception:
            continue
    return []


def _search(query):
    try:
        items = _search_bing_rss(query)
        if items:
            return items
    except Exception:
        pass
    return _search_html(query)


def _clean_address(raw):
    s = html.unescape(raw or "").strip(" ：:|｜、,。・")
    s = re.split(
        r"\s*(?:交通|アクセス|総戸数|戸数|管理会社|管理形態|築年|築年月|竣工|構造|階数|間取り|専有面積|販売価格|賃料|地図|周辺)",
        s,
        maxsplit=1,
    )[0]
    s = re.sub(r"\s+", "", s)
    s = s.replace("−", "-").replace("ー", "-").replace("―", "-").replace("‐", "-").replace("－", "-")
    # 「13丁目12番」のような住所も残すが、末尾のサイト説明は削る。
    s = re.sub(r"[、,。].*$", "", s)
    if len(s) < 6 or len(s) > 70:
        return ""
    if not re.search(r"(?:都|道|府|県|区|市|町|村)", s) or not re.search(r"[0-9]", s):
        return ""
    return s


def _canonical_address(s):
    s = _clean_address(s)
    if not s:
        return ""
    s = s.replace("丁目", "-").replace("番地", "-").replace("番", "-").replace("号", "")
    s = re.sub(r"-+", "-", s).strip("-")
    return s


def _extract_addresses(text, registry_address=""):
    text = _normalize_digits(_strip_tags(text))
    hint = _municipality_hint(registry_address)
    patterns = [
        r"(?:所在地|住所)\s*(?:[:：|｜]\s*)?((?:東京都|北海道|(?:京都|大阪)府|.{2,3}県)[^|｜。]{5,75})",
        r"((?:東京都|北海道|(?:京都|大阪)府|.{2,3}県)[^|｜。\s]{3,50}[0-9]+(?:[-丁目番地号0-9]+){0,12})",
    ]
    out = []
    for pat in patterns:
        for m in re.finditer(pat, text, flags=re.I):
            c = _clean_address(m.group(1))
            if not c:
                continue
            if hint and hint not in c:
                continue
            if c not in out:
                out.append(c)
    return out


def _extract_station_pairs(text):
    text = _normalize_digits(_strip_tags(text))
    # 「京王井の頭線 / 神泉駅 徒歩3分」「神泉駅徒歩3分」どちらも対応。
    pats = [
        r"(?:^|[|｜。])\s*[^|｜。/]{0,60}/\s*([^/|｜。]{1,30})駅\s*徒歩\s*([0-9]{1,3})分",
        r"([^/|｜。\s]{1,30})駅\s*徒歩\s*([0-9]{1,3})分",
    ]
    out = []
    for pat in pats:
        for m in re.finditer(pat, text):
            station = m.group(1).strip(" /|｜")
            station = re.sub(r"^(?:JR|京王|東急|東京メトロ|都営|小田急|西武|東武|京成|相鉄|横浜市営|大阪メトロ)[^/|｜。\s]*[線本]", "", station)
            station = station.strip(" /|｜")
            if station and len(station) <= 25:
                pair = (station, int(m.group(2)))
                if pair not in out:
                    out.append(pair)
    return out


def _extract_total_units_all(text):
    text = _normalize_digits(_strip_tags(text))
    out = []
    pats = [
        r"総戸数(?:\s*\(\s*棟総戸数\s*\))?\s*[:：|｜]?\s*([0-9,]+)\s*戸",
        r"(?:全|戸数)\s*[:：|｜]?\s*([0-9,]+)\s*戸",
    ]
    for pat in pats:
        for m in re.finditer(pat, text):
            try:
                n = int(m.group(1).replace(",", ""))
            except Exception:
                continue
            if 1 <= n <= 10000 and n not in out:
                out.append(n)
    return out


def _clean_company(raw):
    s = html.unescape(raw or "").strip(" ：:|｜、,。・")
    s = re.split(
        r"\s*(?:総戸数|戸数|管理形態|管理人|施工会社|施工|旧分譲主|分譲会社|用途地域|土地権利|建物構造|修繕積立金|管理費|築年|階数|住所|所在地|最寄り駅|最寄駅)",
        s,
        maxsplit=1,
    )[0]
    s = re.sub(r"\s+", " ", s).strip()
    return "" if len(s) > 80 else s


def _extract_management_all(text):
    text = _strip_tags(text)
    out = []
    pats = [
        r"(?:建物)?管理会社\s*[:：|｜]?\s*([^|｜。\n]{2,80})",
        r"管理会社名\s*[:：|｜]?\s*([^|｜。\n]{2,80})",
    ]
    for pat in pats:
        for m in re.finditer(pat, text, flags=re.I):
            c = _clean_company(m.group(1))
            if c and c not in out:
                out.append(c)
    return out


def _collect_search_items(property_name, registry_address=""):
    name = (property_name or "").strip()
    hint = _municipality_hint(registry_address)
    base = '"{}"{}'.format(name, (' "{}"'.format(hint)) if hint else "")

    # 1サイトが検索エンジンに出なくても別サイトが拾えるよう、一般検索＋主要サイトを混ぜる。
    queries = [
        base + ' 所在地 交通 総戸数',
        base + ' 住所 最寄駅 徒歩 総戸数',
        base + ' "総戸数"',
    ]
    for domain in ["homes.co.jp", "suumo.jp", "athome.co.jp", "mansion-review.jp", "nomu.com", "livable.co.jp", "smtrc.jp", "rehouse.co.jp", "mansion-market.com"]:
        queries.append(base + ' site:' + domain)

    out = []
    seen = set()
    for q in queries:
        try:
            items = _search(q)
        except Exception:
            items = []
        for item in items[:15]:
            text = " ".join([item.get("title") or "", item.get("description") or ""])
            if not _item_mentions_property(text, name):
                continue
            url = item.get("url") or ""
            key = (url, text[:500])
            if key in seen:
                continue
            seen.add(key)
            out.append(item)
        # 無料検索に負荷をかけすぎない。
        if len(out) >= 18:
            break
        time.sleep(0.04)
    return out


def _candidate_score_base(item, property_name, registry_address=""):
    text = " ".join([item.get("title") or "", item.get("description") or ""])
    score = _source_weight(item.get("url") or "")
    if _item_mentions_property(text, property_name):
        score += 8
    hint = _municipality_hint(registry_address)
    if hint and hint in text:
        score += 3
    return score


def _pick_consensus(candidates, canon_fn=lambda x: x):
    """candidates: [(value, score, domain, url), ...]. ソースの多様性＋重みで選ぶ。"""
    if not candidates:
        return None, [], 0
    groups = defaultdict(list)
    for value, score, domain, url in candidates:
        key = canon_fn(value)
        if key:
            groups[key].append((value, score, domain, url))
    if not groups:
        return None, [], 0

    ranked = []
    for key, rows in groups.items():
        domains = {r[2] for r in rows}
        # 同一サイトからの重複は票を水増ししない。
        best_by_domain = {}
        for row in rows:
            d = row[2]
            if d not in best_by_domain or row[1] > best_by_domain[d][1]:
                best_by_domain[d] = row
        uniq = list(best_by_domain.values())
        total = sum(r[1] for r in uniq) + max(0, len(domains) - 1) * 7
        ranked.append((total, len(domains), max(r[1] for r in uniq), key, uniq))
    ranked.sort(reverse=True)
    total, domain_count, _, _, rows = ranked[0]
    # 表示値はそのグループ内で最も信頼度の高い表記。
    rows.sort(key=lambda r: r[1], reverse=True)
    return rows[0][0], rows, domain_count


def _lookup_multisource(property_name, registry_address=""):
    items = _collect_search_items(property_name, registry_address)

    address_candidates = []
    station_candidates = []
    units_candidates = []
    management_candidates = []

    for item in items:
        text = " ".join([item.get("title") or "", item.get("description") or ""])
        url = item.get("url") or ""
        domain = _source_domain(url)
        base = _candidate_score_base(item, property_name, registry_address)

        for addr in _extract_addresses(text, registry_address):
            score = base + (5 if re.search(r"所在地|住所", text) else 0)
            address_candidates.append((addr, score, domain, url))

        for station, walk in _extract_station_pairs(text):
            # 駅徒歩はペア単位で一致を見る。
            score = base + (4 if "交通" in text else 0)
            station_candidates.append(((station, walk), score, domain, url))

        for units in _extract_total_units_all(text):
            score = base + (6 if "総戸数" in text else 0)
            units_candidates.append((units, score, domain, url))

        for company in _extract_management_all(text):
            score = base + 5
            management_candidates.append((company, score, domain, url))

    address, addr_rows, addr_count = _pick_consensus(address_candidates, _canonical_address)
    station_pair, station_rows, station_count = _pick_consensus(
        station_candidates, lambda x: _normalize_name(x[0]) + "|" + str(x[1])
    )
    units, units_rows, units_count = _pick_consensus(units_candidates, lambda x: str(x))
    management, mgmt_rows, mgmt_count = _pick_consensus(management_candidates, _normalize_name)

    # 1サイトしかヒットしなくても、物件名・市区町村が一致している高得点候補なら採用。
    # ただし弱い一般スニペットだけなら誤爆防止で採用しない。
    def accepted(rows, count, minimum_score):
        if not rows:
            return False
        return count >= 2 or rows[0][1] >= minimum_score

    result = {}
    source_urls = []
    source_domains = []

    if address and accepted(addr_rows, addr_count, 17):
        result["searched_address"] = address
        result["address_source"] = "複数不動産サイト検索"
        source_urls += [r[3] for r in addr_rows if r[3]]
        source_domains += [r[2] for r in addr_rows]
        result["address_source_count"] = addr_count

    if station_pair and accepted(station_rows, station_count, 16):
        result["homes_station_name"] = station_pair[0]
        result["homes_walk_minutes"] = station_pair[1]
        result["station_source"] = "複数不動産サイト検索"
        source_urls += [r[3] for r in station_rows if r[3]]
        source_domains += [r[2] for r in station_rows]
        result["station_source_count"] = station_count

    if units is not None and accepted(units_rows, units_count, 18):
        result["total_units"] = int(units)
        result["units_source"] = "複数不動産サイト検索"
        source_urls += [r[3] for r in units_rows if r[3]]
        source_domains += [r[2] for r in units_rows]
        result["units_source_count"] = units_count

    if management and accepted(mgmt_rows, mgmt_count, 17):
        result["management_company"] = management
        result["management_source"] = "複数不動産サイト検索"
        source_urls += [r[3] for r in mgmt_rows if r[3]]
        source_domains += [r[2] for r in mgmt_rows]
        result["management_source_count"] = mgmt_count

    missing = []
    if not result.get("searched_address"):
        missing.append("所在地")
    if not result.get("homes_station_name") or result.get("homes_walk_minutes") is None:
        missing.append("駅徒歩")
    if result.get("total_units") is None:
        missing.append("総戸数")

    result["missing_fields"] = missing
    result["source_urls"] = list(dict.fromkeys(source_urls))[:8]
    result["source_domains"] = list(dict.fromkeys([x for x in source_domains if x]))[:8]
    return result


def lookup_building_info(property_name, address=""):
    """Ver.3.7: 複数不動産サイトの検索結果を照合して建物情報を取得。"""
    name = (property_name or "").strip()
    if not name:
        return {
            "management_company": "", "total_units": None, "searched_address": "",
            "homes_station_name": "", "homes_walk_minutes": None,
            "building_info_source_urls": [], "building_info_warning": "",
        }

    cache = _load_cache()
    cache_key = "building_v370_multisource:" + name + "|" + (address or "").strip()
    if cache_key in cache:
        return cache[cache_key]

    multi = _lookup_multisource(name, address)
    missing = list(multi.get("missing_fields") or [])
    warning = ""
    if missing:
        warning = "複数の不動産サイト検索で{}を十分に確認できなかったため、その項目のみ従来方法/登記情報を使用します。".format("・".join(missing))

    result = {
        "searched_address": multi.get("searched_address") or "",
        "homes_station_name": multi.get("homes_station_name") or "",
        "homes_walk_minutes": multi.get("homes_walk_minutes"),
        "management_company": multi.get("management_company") or "",
        "total_units": multi.get("total_units"),
        "building_info_source_urls": multi.get("source_urls") or [],
        "building_info_source_domains": multi.get("source_domains") or [],
        "building_info_warning": warning,
        "address_source": multi.get("address_source") or "",
        "station_lookup_source": multi.get("station_source") or "",
        "units_source": multi.get("units_source") or "",
        "management_source": multi.get("management_source") or "",
        "address_source_count": multi.get("address_source_count", 0),
        "station_source_count": multi.get("station_source_count", 0),
        "units_source_count": multi.get("units_source_count", 0),
    }
    cache[cache_key] = result
    _save_cache(cache)
    return result
