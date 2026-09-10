import html
import json
import os
import re
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

CACHE_PATH = os.path.join(os.path.dirname(__file__), ".building_cache.json")
UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152 Safari/537.36"


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
            "Accept-Language": "ja,en;q=0.8",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read()
        charset = resp.headers.get_content_charset() or "utf-8"
        return raw.decode(charset, errors="ignore")


def _strip_tags(s):
    s = re.sub(r"<script\b.*?</script>|<style\b.*?</style>", " ", s or "", flags=re.I | re.S)
    s = re.sub(r"<[^>]+>", " ", s)
    s = html.unescape(s)
    return re.sub(r"\s+", " ", s).strip()


def _normalize_digits(s):
    return (s or "").translate(str.maketrans("０１２３４５６７８９，", "0123456789,"))


def _search_bing_rss(query):
    """Bingの通常Web検索RSS。APIキー不要。タイトル・説明・URLを返す。"""
    url = "https://www.bing.com/search?" + urllib.parse.urlencode({"q": query, "format": "rss", "setlang": "ja-JP"})
    xml_text = _get_text(url)
    root = ET.fromstring(xml_text)
    items = []
    for item in root.findall(".//item"):
        title = item.findtext("title") or ""
        desc = item.findtext("description") or ""
        link = item.findtext("link") or ""
        # マンションナビはユーザー指定で検索対象から除外。
        if "t23m-navi.jp" in link.lower() or "t23m-navi.jp" in (title + desc).lower():
            continue
        items.append({"title": _strip_tags(title), "description": _strip_tags(desc), "url": link})
    return items


def _search_html(query):
    """RSSが使えない環境向けの通常検索フォールバック。"""
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


def _clean_company(raw):
    s = html.unescape(raw or "").strip(" ：:|｜、,。・")
    # 次の項目ラベルや検索結果UIが巻き込まれた場合にそこで止める。
    s = re.split(
        r"\s*(?:総戸数|戸数|管理形態|管理人|施工会社|施工|旧分譲主|分譲会社|用途地域|土地権利|建物構造|修繕積立金|管理費|築年|階数|住所|所在地|最寄り駅|最寄駅)\s*[:：]?",
        s,
        maxsplit=1,
    )[0]
    s = re.sub(r"\s+", " ", s).strip()
    # 明らかに検索説明文が長く入ったものは採用しない。
    if len(s) > 80:
        return ""
    return s


def _extract_management(text):
    text = _strip_tags(text)
    patterns = [
        r"(?:建物)?管理会社\s*[:：|｜]?\s*([^|｜。\n]{2,80})",
        r"管理会社名\s*[:：|｜]?\s*([^|｜。\n]{2,80})",
    ]
    for pat in patterns:
        m = re.search(pat, text, flags=re.I)
        if m:
            candidate = _clean_company(m.group(1))
            if candidate:
                return candidate
    return ""


def _extract_total_units(text):
    text = _normalize_digits(_strip_tags(text))
    patterns = [
        r"総戸数(?:\s*\(\s*棟総戸数\s*\))?\s*[:：|｜]?\s*([0-9,]+)\s*戸",
        r"戸数\s*[:：|｜]?\s*([0-9,]+)\s*戸",
        r"全\s*([0-9,]+)\s*戸",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            try:
                n = int(m.group(1).replace(",", ""))
                if 1 <= n <= 10000:
                    return n
            except Exception:
                pass
    return None


def _location_hint(address):
    s = (address or "").replace("　", " ")
    m = re.search(r"((?:東京都|北海道|(?:京都|大阪)府|.{2,3}県)?[^0-9０-９]{1,14}(?:区|市))", s)
    return m.group(1).strip() if m else ""


def lookup_building_info(property_name, address=""):
    """
    物件名を一般Web検索し、公開検索結果から建物管理会社・総戸数を拾う。
    特定サイトには依存しない。見つからない項目は空欄。
    """
    name = (property_name or "").strip()
    if not name:
        return {
            "management_company": "", "total_units": None,
            "building_info_source_urls": [], "building_info_warning": "",
        }

    cache = _load_cache()
    cache_key = "building:" + name + "|" + (address or "").strip()
    if cache_key in cache:
        return cache[cache_key]

    hint = _location_hint(address)
    base = f'"{name}"' + (f" {hint}" if hint else "")
    # ユーザー指定によりマンションナビは使わない。
    queries = [
        f'{base} "管理会社" -site:t23m-navi.jp',
        f'{base} "総戸数" -site:t23m-navi.jp',
        f'{base} 管理会社 総戸数 -site:t23m-navi.jp',
    ]

    management = ""
    total_units = None
    source_urls = []
    errors = []

    for q in queries:
        try:
            items = _search(q)
        except Exception as e:
            errors.append(str(e))
            continue
        for item in items[:10]:
            url = item.get("url") or ""
            if "t23m-navi.jp" in url.lower():
                continue
            text = " ".join([item.get("title") or "", item.get("description") or ""])
            if not management:
                management = _extract_management(text)
            if total_units is None:
                total_units = _extract_total_units(text)
            if url and url not in source_urls:
                source_urls.append(url)
            if management and total_units is not None:
                break
        if management and total_units is not None:
            break
        time.sleep(0.15)

    warning = ""
    if not management and total_units is None:
        warning = "一般Web検索で建物管理会社・総戸数を確認できませんでした。空欄で出力します。"
    elif not management:
        warning = "一般Web検索で建物管理会社を確認できませんでした。管理会社は空欄で出力します。"
    elif total_units is None:
        warning = "一般Web検索で総戸数を確認できませんでした。総戸数は空欄で出力します。"

    result = {
        "management_company": management,
        "total_units": total_units,
        "building_info_source_urls": source_urls[:5],
        "building_info_warning": warning,
    }
    cache[cache_key] = result
    _save_cache(cache)
    return result
