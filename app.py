import io
import copy
import os
import re
import zipfile
import urllib.parse
from datetime import date
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET

import streamlit as st

from geo_lookup import enrich_location
from building_lookup import lookup_building_info
from touki_parser import analyze_registry, calculate_monthly_payment, extract_pdf_text

st.set_page_config(page_title="登記簿 → CS Excel（無料版）", page_icon="🏢", layout="wide")

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "sample", "CS_登記簿自動入力.xlsx")

# 1物件目から左→右へ、3列ずつずらして入力する。
SLOT_MAPS = [
    {
        "property_name": "C9", "floors": "C10", "built_year": "C11", "building_age": "D11", "room": "E10", "area_sqm": "E11",
        "property_address": "C12", "station_walk": "C13", "management_company": "C15", "total_units": "D10", "loan_company": "C16", "loan_amount_yen": "C17", "two_months_later_ym": "C18",
        "estimated_balance_yen": "C19", "interest_rate": "C22", "purchase_date": "C23", "payoff_ym": "C25", "monthly_payment_yen": "C27",
    },
    {
        "property_name": "F9", "floors": "F10", "built_year": "F11", "building_age": "G11", "room": "H10", "area_sqm": "H11",
        "property_address": "F12", "station_walk": "F13", "management_company": "F15", "total_units": "G10", "loan_company": "F16", "loan_amount_yen": "F17", "two_months_later_ym": "F18",
        "estimated_balance_yen": "F19", "interest_rate": "F22", "purchase_date": "F23", "payoff_ym": "F25", "monthly_payment_yen": "F27",
    },
    {
        "property_name": "I9", "floors": "I10", "built_year": "I11", "building_age": "J11", "room": "K10", "area_sqm": "K11",
        "property_address": "I12", "station_walk": "I13", "management_company": "I15", "total_units": "J10", "loan_company": "I16", "loan_amount_yen": "I17", "two_months_later_ym": "I18",
        "estimated_balance_yen": "I19", "interest_rate": "I22", "purchase_date": "I23", "payoff_ym": "I25", "monthly_payment_yen": "I27",
    },
    {
        "property_name": "L9", "floors": "L10", "built_year": "L11", "building_age": "M11", "room": "N10", "area_sqm": "N11",
        "property_address": "L12", "station_walk": "L13", "management_company": "L15", "total_units": "M10", "loan_company": "L16", "loan_amount_yen": "L17", "two_months_later_ym": "L18",
        "estimated_balance_yen": "L19", "interest_rate": "L22", "purchase_date": "L23", "payoff_ym": "L25", "monthly_payment_yen": "L27",
    },
    {
        "property_name": "O9", "floors": "O10", "built_year": "O11", "building_age": "P11", "room": "Q10", "area_sqm": "Q11",
        "property_address": "O12", "station_walk": "O13", "management_company": "O15", "total_units": "P10", "loan_company": "O16", "loan_amount_yen": "O17", "two_months_later_ym": "O18",
        "estimated_balance_yen": "O19", "interest_rate": "O22", "purchase_date": "O23", "payoff_ym": "O25", "monthly_payment_yen": "O27",
    },
    {
        "property_name": "R9", "floors": "R10", "built_year": "R11", "building_age": "S11", "room": "T10", "area_sqm": "T11",
        "property_address": "R12", "station_walk": "R13", "management_company": "R15", "total_units": "S10", "loan_company": "R16", "loan_amount_yen": "R17", "two_months_later_ym": "R18",
        "estimated_balance_yen": "R19", "interest_rate": "R22", "purchase_date": "R23", "payoff_ym": "R25", "monthly_payment_yen": "R27",
    },
    {
        "property_name": "U9", "floors": "U10", "built_year": "U11", "building_age": "V11", "room": "W10", "area_sqm": "W11",
        "property_address": "U12", "station_walk": "U13", "management_company": "U15", "total_units": "V10", "loan_company": "U16", "loan_amount_yen": "U17", "two_months_later_ym": "U18",
        "estimated_balance_yen": "U19", "interest_rate": "U22", "purchase_date": "U23", "payoff_ym": "U25", "monthly_payment_yen": "U27",
    },
]

# 固定資産税はVer.3.8.1で廃止。テンプレート側の旧値だけ明示的に空欄化する。
REMOVED_TAX_CELLS = []

# テンプレート側に元々ある計算式は、ダウンロード時の一時非表示を防ぐため毎回明示的に再設定する。
FORCED_FORMULAS = {
    "E23": 'IFERROR(IF(C25=0,"",DATEDIF($C$7,C25,"ｍ")),"")',
    "E25": 'IF(C23=0,"",DATEDIF(C23,C25,"y"))',
    "H23": 'IFERROR(IF(F25=0,"",DATEDIF($C$7,F25,"ｍ")),"")',
    "H25": 'IF(F23=0,"",DATEDIF(F23,F25,"y"))',
    "K23": 'IFERROR(IF(I25=0,"",DATEDIF($C$7,I25,"ｍ")),"")',
    "K25": 'IF(I23=0,"",DATEDIF(I23,I25,"y"))',
    "N23": 'IFERROR(IF(L25=0,"",DATEDIF($C$7,L25,"ｍ")),"")',
    "N25": 'IF(L23=0,"",DATEDIF(L23,L25,"y"))',
    "Q23": 'IFERROR(IF(O25=0,"",DATEDIF($C$7,O25,"ｍ")),"")',
    "Q25": 'IF(O23=0,"",DATEDIF(O23,O25,"y"))',
    "T23": 'IFERROR(IF(R25=0,"",DATEDIF($C$7,R25,"ｍ")),"")',
    "T25": 'IF(R23=0,"",DATEDIF(R23,R25,"y"))',
    "W23": 'IFERROR(IF(U25=0,"",DATEDIF($C$7,U25,"ｍ")),"")',
    "W25": 'IF(U23=0,"",DATEDIF(U23,U25,"y"))',
}


def _cell_xml(ref, value, style=None, prefix=""):
    p = prefix
    s = f' s="{style}"' if style is not None else ""
    if value is None or value == "":
        return f'<{p}c r="{ref}"{s}/>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<{p}c r="{ref}"{s}><{p}v>{value}</{p}v></{p}c>'
    return f'<{p}c r="{ref}"{s} t="inlineStr"><{p}is><{p}t>{escape(str(value))}</{p}t></{p}is></{p}c>'




def _col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _num_to_col(n: int) -> str:
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def _ensure_cell(sheet_xml: str, ref: str, prefix="") -> str:
    """空欄セルがXML上に存在しない場合でも、左側の同型セルのスタイルを引き継いで作成する。"""
    p = re.escape(prefix)
    cell_pat = rf'<{p}c\b[^>]*\br="{re.escape(ref)}"[^>]*(?:/>|>.*?</{p}c>)'
    if re.search(cell_pat, sheet_xml, flags=re.S):
        return sheet_xml

    mref = re.fullmatch(r'([A-Z]+)(\d+)', ref)
    if not mref:
        raise ValueError(f"セル参照 {ref} が不正です。")
    col, row = mref.group(1), mref.group(2)

    # 物件枠は3列刻みなので、まず3列左の同じ行からスタイルを引き継ぐ。
    style = None
    col_num = _col_to_num(col)
    candidates = []
    if col_num > 3:
        candidates.append(f"{_num_to_col(col_num - 3)}{row}")
    # 念のため、さらに左の既存セルも探索。
    candidates += [f"{_num_to_col(n)}{row}" for n in range(col_num - 1, 0, -1)]
    seen = set()
    for src in candidates:
        if src in seen:
            continue
        seen.add(src)
        sm = re.search(rf'<{p}c\b([^>]*\br="{re.escape(src)}"[^>]*)', sheet_xml, flags=re.S)
        if sm:
            ssm = re.search(r'\bs="([^"]+)"', sm.group(1) or "")
            if ssm:
                style = ssm.group(1)
            break

    row_pat = rf'(<{p}row\b[^>]*\br="{row}"[^>]*>)(.*?)(</{p}row>)'
    rm = re.search(row_pat, sheet_xml, flags=re.S)
    if not rm:
        raise ValueError(f"テンプレート内に行 {row} が見つかりませんでした。")
    new_cell = _cell_xml(ref, "", style, prefix)
    return sheet_xml[:rm.start()] + rm.group(1) + rm.group(2) + new_cell + rm.group(3) + sheet_xml[rm.end():]


def _replace_formula_cell(sheet_xml: str, ref: str, formula: str, prefix="") -> str:
    """指定セルへ数式を必ず再設定する（既存スタイルは維持）。"""
    sheet_xml = _ensure_cell(sheet_xml, ref, prefix)
    p = re.escape(prefix)
    pat = rf'<{p}c\b([^>]*\br="{re.escape(ref)}"[^>]*)/>|<{p}c\b([^>]*\br="{re.escape(ref)}"[^>/]*)>(.*?)</{p}c>'
    m = re.search(pat, sheet_xml, flags=re.S)
    attrs = (m.group(1) or m.group(2) or "")
    sm = re.search(r'\bs="([^"]+)"', attrs)
    style_attr = f' s="{sm.group(1)}"' if sm else ""
    # t="str" + 空のキャッシュ値。Excel/Numbers側で開いた際に再計算させる。
    new_xml = f'<{prefix}c r="{ref}"{style_attr} t="str"><{prefix}f>{escape(formula)}</{prefix}f><{prefix}v/></{prefix}c>'
    return sheet_xml[:m.start()] + new_xml + sheet_xml[m.end():]

def _replace_cell(sheet_xml: str, ref: str, value, prefix="") -> str:
    sheet_xml = _ensure_cell(sheet_xml, ref, prefix)
    p = re.escape(prefix)
    pat = rf'<{p}c\b([^>]*\br="{re.escape(ref)}"[^>]*)/>|<{p}c\b([^>]*\br="{re.escape(ref)}"[^>/]*)>(.*?)</{p}c>'
    m = re.search(pat, sheet_xml, flags=re.S)
    attrs = (m.group(1) or m.group(2) or "")
    sm = re.search(r'\bs="([^"]+)"', attrs)
    style = sm.group(1) if sm else None
    new_xml = _cell_xml(ref, value, style, prefix)
    return sheet_xml[:m.start()] + new_xml + sheet_xml[m.end():]


def _ym_after_months(months: int, base=None) -> str:
    base = base or date.today()
    total = base.year * 12 + (base.month - 1) + months
    y, m0 = divmod(total, 12)
    return f"{y}/{m0 + 1}時点"


def _building_age(built_year):
    try:
        y = int(str(built_year).strip())
        return max(0, date.today().year - y)
    except Exception:
        return ""


def _payoff_ym(purchase_date: str, years: int = 35) -> str:
    """購入日からちょうど35年後。日まで取れている場合は日付も維持する。"""
    s = (purchase_date or "").strip().replace("-", "/")
    m = re.match(r"^(\d{4})/(\d{1,2})(?:/(\d{1,2}))?$", s)
    if not m:
        return ""
    y = int(m.group(1)) + years
    mo = int(m.group(2))
    d = int(m.group(3)) if m.group(3) else None
    if d is None:
        return f"{y}/{mo}"
    # 2/29取得など、35年後に同日が存在しない場合はその月末に丸める。
    import calendar
    d = min(d, calendar.monthrange(y, mo)[1])
    return f"{y}/{mo}/{d}"


def _format_room(room) -> str:
    """部屋番号は「113号室」の形で表示・出力する。"""
    s = str(room or "").strip()
    if not s:
        return ""
    s = re.sub(r"\s*号室$", "", s).strip()
    return f"{s}号室"


def _result_values(result: dict) -> dict:
    mortgage_status = result.get("mortgage_status") or ("active" if result.get("active_mortgage_count", 0) else "none")
    no_active_loan = mortgage_status != "active"
    past_cancelled = mortgage_status == "cancelled_current_owner"
    station = (result.get("station_name") or "").strip()
    walk = result.get("walk_minutes")
    if station and walk:
        station_walk = f"{station}駅 徒歩{int(walk)}分"
    elif walk:
        station_walk = f"徒歩{int(walk)}分"
    else:
        station_walk = station

    rate = result.get("interest_rate_pct")
    interest_text = "" if (mortgage_status == "none" or rate is None) else f"{float(rate):g}%"
    principal = 0 if mortgage_status == "none" else int(result.get("loan_amount_yen") or 0)
    monthly_payment = 0 if no_active_loan else int(result.get("monthly_payment_yen") or calculate_monthly_payment(principal, rate or 2.3, 35))
    if past_cancelled:
        company = (result.get("loan_company") or "").strip()
        loan_company_text = f"{company}（抹消済）" if company else "過去抵当あり（抹消済）"
    elif mortgage_status == "none":
        loan_company_text = "抵当なし"
    else:
        loan_company_text = result.get("loan_company") or ""

    return {
        "property_name": result.get("property_name") or "",
        "floors": int(result["floors"]) if str(result.get("floors") or "").isdigit() else (result.get("floors") or ""),
        "built_year": int(result["built_year"]) if str(result.get("built_year") or "").isdigit() else (result.get("built_year") or ""),
        "building_age": _building_age(result.get("built_year")),
        "room": _format_room(result.get("room")),
        "area_sqm": result.get("area_sqm") if result.get("area_sqm") not in (None, "") else "",
        "property_address": result.get("property_address") or result.get("geocoded_address") or "",
        "station_walk": station_walk,
        "management_company": result.get("management_company") or "",
        "total_units": int(result["total_units"]) if result.get("total_units") not in (None, "") else "",
        "loan_company": loan_company_text,
        "loan_amount_yen": principal,
        "two_months_later_ym": _ym_after_months(2),
        "estimated_balance_yen": 0 if no_active_loan else int(result.get("estimated_balance_yen") or 0),
        "interest_rate": interest_text,
        "purchase_date": result.get("purchase_date") or result.get("current_owner_acquired_date_slash") or "",
        "payoff_ym": _payoff_ym(result.get("purchase_date") or result.get("current_owner_acquired_date_slash") or ""),
        "monthly_payment_yen": monthly_payment,
    }



def _set_cell_style(sheet_xml: str, ref: str, style_idx: int, prefix="") -> str:
    p = re.escape(prefix)
    pat = rf'(<{p}c\b[^>]*\br="{re.escape(ref)}"[^>]*)(/?>)'
    m = re.search(pat, sheet_xml)
    if not m:
        return sheet_xml
    start = m.group(1)
    if re.search(r'\bs="[^"]+"', start):
        start = re.sub(r'\bs="[^"]+"', f' s="{style_idx}"', start)
    else:
        start += f' s="{style_idx}"'
    return sheet_xml[:m.start()] + start + m.group(2) + sheet_xml[m.end():]


def _gray_style_map(styles_xml: bytes, sheet_xml: str, refs: list[str], prefix=""):
    """元のセル書式を維持したまま、背景だけグレーにしたスタイルを追加する。"""
    root = ET.fromstring(styles_xml)
    ns_uri = root.tag.split('}')[0].strip('{') if '}' in root.tag else ''
    q = (lambda name: f'{{{ns_uri}}}{name}' if ns_uri else name)
    fills = root.find(q('fills'))
    cell_xfs = root.find(q('cellXfs'))
    if fills is None or cell_xfs is None:
        return styles_xml, sheet_xml

    fill = ET.Element(q('fill'))
    pf = ET.SubElement(fill, q('patternFill'), {'patternType': 'solid'})
    ET.SubElement(pf, q('fgColor'), {'rgb': 'FFD9D9D9'})
    ET.SubElement(pf, q('bgColor'), {'indexed': '64'})
    gray_fill_id = len(list(fills))
    fills.append(fill)
    fills.set('count', str(len(list(fills))))

    style_cache = {}
    for ref in refs:
        m = re.search(rf'<{re.escape(prefix)}c\b([^>]*\br="{re.escape(ref)}"[^>]*)', sheet_xml)
        if not m:
            continue
        sm = re.search(r'\bs="(\d+)"', m.group(1))
        old_idx = int(sm.group(1)) if sm else 0
        if old_idx not in style_cache:
            xfs = list(cell_xfs)
            if old_idx >= len(xfs):
                continue
            new_xf = copy.deepcopy(xfs[old_idx])
            new_xf.set('fillId', str(gray_fill_id))
            new_xf.set('applyFill', '1')
            cell_xfs.append(new_xf)
            new_idx = len(list(cell_xfs)) - 1
            style_cache[old_idx] = new_idx
        sheet_xml = _set_cell_style(sheet_xml, ref, style_cache[old_idx], prefix)
    cell_xfs.set('count', str(len(list(cell_xfs))))
    return ET.tostring(root, encoding='utf-8', xml_declaration=True), sheet_xml


def fill_cs_template(results: list[dict]) -> bytes:
    """アップロード順に、同じExcelの左→右の枠へ最大7物件を書き込む。"""
    if len(results) > len(SLOT_MAPS):
        raise ValueError(f"このテンプレートは最大{len(SLOT_MAPS)}物件までです。")

    with open(TEMPLATE_PATH, "rb") as f:
        src = f.read()

    zin = zipfile.ZipFile(io.BytesIO(src), "r")
    sheet_name = "xl/worksheets/sheet1.xml"
    styles_name = "xl/styles.xml"
    sheet_xml = zin.read(sheet_name).decode("utf-8")
    styles_xml = zin.read(styles_name)
    prefix = "x:" if "<x:worksheet" in sheet_xml else ""

    # テンプレートに残っている外部参照数式等を消し、7枠すべてを空の状態から作る。
    for slot in SLOT_MAPS:
        for ref in slot.values():
            sheet_xml = _replace_cell(sheet_xml, ref, "", prefix)
    for ref in REMOVED_TAX_CELLS:
        sheet_xml = _replace_cell(sheet_xml, ref, "", prefix)

    gray_refs = []
    for idx, result in enumerate(results):
        values = _result_values(result)
        for key, ref in SLOT_MAPS[idx].items():
            sheet_xml = _replace_cell(sheet_xml, ref, values[key], prefix)
        if result.get("mortgage_status") == "cancelled_current_owner":
            col = ["C", "F", "I", "L", "O", "R", "U"][idx]
            gray_refs.extend([f"{col}{row}" for row in range(16, 26)])

    # E23/E25（および2〜7件目の対応セル）の数式を毎回必ず入れ直す。
    for ref, formula in FORCED_FORMULAS.items():
        sheet_xml = _replace_formula_cell(sheet_xml, ref, formula, prefix)

    if gray_refs:
        styles_xml, sheet_xml = _gray_style_map(styles_xml, sheet_xml, gray_refs, prefix)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == sheet_name:
                data = sheet_xml.encode("utf-8")
            elif item.filename == styles_name:
                data = styles_xml
            else:
                data = zin.read(item.filename)
            zout.writestr(item, data)
    zin.close()
    return out.getvalue()


def parse_uploaded(uploaded_file):
    pdf_bytes = uploaded_file.getvalue()
    pdf_text = extract_pdf_text(pdf_bytes)
    if len(pdf_text) < 80:
        raise ValueError("PDFから十分な文字を取得できませんでした。画像だけのスキャンPDFの可能性があります。")
    result = analyze_registry(pdf_text)
    result["source_file"] = uploaded_file.name
    result["pdf_text"] = pdf_text
    return result


# ===== Ver.3.3 Dark UI =====
st.markdown(
    """
    <style>
    .stApp {
        background: radial-gradient(circle at 20% 0%, #1f2937 0%, #111827 34%, #0b1220 100%);
        color: #f8fafc;
    }
    [data-testid="stHeader"] { background: rgba(0,0,0,0); }
    [data-testid="stToolbar"] { right: 1rem; }
    [data-testid="stSidebar"] {
        background: #0a0f1a;
        border-right: 1px solid #263244;
    }
    .block-container {
        max-width: 1460px;
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }
    h1, h2, h3, p, label, .stMarkdown { color: #f8fafc; }
    .cs-header {
        display:flex; align-items:center; justify-content:space-between;
        padding: 4px 2px 18px 2px;
    }
    .cs-title-wrap { display:flex; align-items:center; gap:14px; }
    .cs-logo {
        width:46px; height:46px; border-radius:12px;
        border:1px solid #41506a; background:#182131;
        display:flex; align-items:center; justify-content:center;
        font-size:23px; box-shadow:0 8px 30px rgba(0,0,0,.22);
    }
    .cs-title { font-size:28px; font-weight:800; letter-spacing:.02em; margin:0; }
    .cs-subtitle { margin-top:4px; color:#97a6ba; font-size:13px; }
    .cs-brand { color:#cbd5e1; font-weight:700; letter-spacing:.18em; font-size:12px; }
    .cs-upload-note {
        background:#151e2d; border:1px solid #334155; border-radius:14px;
        padding:12px 16px; color:#cbd5e1; margin-bottom:10px;
    }
    [data-testid="stFileUploader"] {
        background:#151e2d; border:1px solid #334155; border-radius:15px;
        padding:10px 14px 2px 14px; box-shadow:0 14px 30px rgba(0,0,0,.13);
    }
    [data-testid="stFileUploaderDropzone"] {
        background:#101827; border:1px dashed #53637d; border-radius:12px;
    }
    [data-testid="stExpander"] {
        background: linear-gradient(180deg,#172131,#111a28);
        border:1px solid #334155; border-radius:14px;
        overflow:hidden; box-shadow:0 12px 28px rgba(0,0,0,.15);
    }
    [data-testid="stExpander"] summary { background:#172131; }
    [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input {
        background:#0d1522 !important; color:#f8fafc !important;
        border:1px solid #39475d !important; border-radius:9px !important;
    }
    [data-testid="stMetric"] {
        background:#0d1522; border:1px solid #334155; border-radius:10px;
        padding:8px 12px;
    }
    .stButton > button, .stDownloadButton > button {
        min-height:48px; border-radius:10px; font-weight:800;
        background:linear-gradient(180deg,#d5ad62,#b98736) !important;
        color:#111827 !important; border:1px solid #e2bf79 !important;
        box-shadow:0 8px 20px rgba(185,135,54,.18);
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        border-color:#f1d39a !important; transform:translateY(-1px);
    }
    .cs-section-title {
        color:#e5e7eb; font-size:15px; font-weight:800; margin:20px 0 10px;
        letter-spacing:.03em;
    }
    .cs-status-ok { color:#34d399; font-weight:700; }
    hr { border-color:#2d3748 !important; }
    .small-note { color:#94a3b8; font-size:12px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="cs-header">
      <div class="cs-title-wrap">
        <div class="cs-logo">▣</div>
        <div>
          <div class="cs-title">登記簿 自動入力ツール</div>
          <div class="cs-subtitle">登記簿PDFから物件情報を自動抽出し、CS Excelへ入力します。</div>
        </div>
      </div>
      <div class="cs-brand">CS REAL ESTATE</div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("### CS REAL ESTATE")
    st.caption("業務用 登記簿入力ツール")
    st.divider()
    st.markdown("**▣ 登記簿入力**")
    st.caption("Ver.3.8.12 / 最大7物件対応・空欄セル自動生成修正")
    st.divider()
    st.success("OpenAI APIキー不要 / 利用料0円")
    st.info("住所・総戸数・管理会社は手入力。HOME'S・マンションナビ確認リンク付き。現所有者の過去抵当が抹消済みの場合は、過去ローン情報を残しExcelのローン欄をグレー表示します。")
    st.warning("残債・毎月返済額は35年・元利均等。金利記載なしは2.3%。固定資産税項目は廃止しました。")

st.markdown(
    '<div class="cs-upload-note">☁️ <b>登記簿PDFをアップロード</b>　最初のPDF→左端、2件目→その右…の順で最大7件まで入力できます。</div>',
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "PDFをここにドラッグ＆ドロップ、またはファイルを選択",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="visible",
)

if uploaded_files:
    st.caption(f"選択中：{len(uploaded_files)}件")
    if len(uploaded_files) > 7:
        st.error("このExcelテンプレートは7物件分までです。PDFを7件以内にしてください。")
    elif st.button("解析＋物件情報検索を開始", type="primary", use_container_width=True):
        results = []
        progress = st.progress(0)
        status = st.empty()
        for i, uploaded in enumerate(uploaded_files, start=1):
            status.write(f"{i}/{len(uploaded_files)}件目：{uploaded.name} を解析中...")
            try:
                r = parse_uploaded(uploaded)
                registry_address = r.get("property_address", "")
                building = lookup_building_info(r.get("property_name", ""), registry_address)
                r.update(building)

                # Ver.3.8.1: 住所・総戸数・管理会社は自動入力しない。
                # 駅徒歩だけは従来どおりWeb検索を試す。
                r["registry_address"] = registry_address
                r["property_address"] = ""
                r["management_company"] = ""
                r["total_units"] = None

                homes_station = (building.get("homes_station_name") or "").strip()
                homes_walk = building.get("homes_walk_minutes")
                if homes_station and homes_walk not in (None, ""):
                    r["station_name"] = homes_station
                    r["walk_minutes"] = int(homes_walk)
                    r["station_source"] = "Web検索"
                    loc = {}
                else:
                    # 駅徒歩がWeb検索で取れない場合のみ、登記上の所在から最寄駅を推定。
                    loc = enrich_location(registry_address, r.get("property_name", ""))
                    r.update(loc)
                    r["station_source"] = "登記所在から推定"

                # 建物情報検索の「住所・総戸数が取れない」警告はVer.3.8では不要。
                if loc.get("warning"):
                    r.setdefault("warnings", []).append(loc["warning"])
                results.append(r)
            except Exception as e:
                results.append({
                    "source_file": uploaded.name,
                    "warnings": [f"解析に失敗しました: {e}"],
                    "property_name": "", "room": "", "floors": "", "area_sqm": None,
                    "built_year": "", "loan_company": "", "loan_amount_yen": 0,
                    "property_address": "", "registry_address": "", "searched_address": "", "station_name": "", "walk_minutes": None,
                    "management_company": "", "total_units": None, "building_info_source_urls": [],
                    "loan_date": "", "interest_rate_pct": None, "interest_rate_source": "",
                    "estimated_balance_yen": 0, "monthly_payment_yen": 0, "active_mortgage_count": 0,
                })
            progress.progress(i / len(uploaded_files))
        status.success("解析完了。下のカードで内容を確認・修正できます。")
        st.session_state["results"] = results

if "results" in st.session_state:
    st.markdown('<div class="cs-section-title">解析結果</div>', unsafe_allow_html=True)
    edited_results = []

    for idx, r in enumerate(st.session_state["results"]):
        slot_col = ["C", "F", "I", "L", "O", "R", "U"][idx]
        title = r.get("property_name") or r.get("source_file") or f"物件{idx+1}"
        with st.expander(f"物件{idx+1}　{title}　｜　Excel {slot_col}列側　● 解析済み", expanded=True):
            for w in r.get("warnings", []):
                st.warning(w)

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.markdown("**物件概要**")
                property_name = st.text_input("物件名", value=r.get("property_name") or "", key=f"name_{idx}")
                if property_name.strip():
                    homes_query = urllib.parse.quote_plus(f'"{property_name.strip()}" site:homes.co.jp')
                    homes_url = f"https://www.google.com/search?q={homes_query}"
                    mansion_navi_query = urllib.parse.quote_plus(f'"{property_name.strip()}" マンションナビ')
                    mansion_navi_url = f"https://www.google.com/search?q={mansion_navi_query}"
                    link_col1, link_col2 = st.columns(2)
                    with link_col1:
                        st.link_button("🔎 HOME'S", homes_url, use_container_width=True)
                    with link_col2:
                        st.link_button("🔎 マンションナビ", mansion_navi_url, use_container_width=True)
                    st.caption("住所・総戸数・管理会社を確認して、右の建物情報へ手入力")
                room = st.text_input("部屋番号", value=_format_room(r.get("room")), key=f"room_{idx}")
                floors = st.text_input("階数", value=r.get("floors") or "", key=f"floors_{idx}")
                area = st.number_input("閉米数", min_value=0.0, value=float(r.get("area_sqm") or 0), step=0.01, key=f"area_{idx}")

            with c2:
                st.markdown("**建物情報**")
                built_year = st.text_input("築年", value=r.get("built_year") or "", key=f"year_{idx}")

                # 最寄駅と徒歩分数は、確認・修正しやすいよう1つの入力欄にまとめる。
                default_station = (r.get("station_name") or "").strip()
                default_walk = r.get("walk_minutes")
                if default_station and not default_station.endswith("駅"):
                    default_station += "駅"
                if default_station and default_walk:
                    default_station_walk = f"{default_station} 徒歩{int(default_walk)}分"
                else:
                    default_station_walk = default_station
                station_walk_text = st.text_input(
                    "最寄駅 駅徒歩",
                    value=default_station_walk,
                    placeholder="例：神泉駅 徒歩3分",
                    key=f"station_walk_{idx}",
                )

                # Excel出力用に駅名と徒歩分数へ戻す。
                station_name = station_walk_text.strip()
                walk_minutes = None
                m_walk = re.search(r"徒歩\s*(\d+)\s*分", station_name)
                if m_walk:
                    walk_minutes = int(m_walk.group(1))
                    station_name = station_name[:m_walk.start()].strip()
                if station_name.endswith("駅"):
                    station_name = station_name[:-1].strip()

                property_address = st.text_input("物件住所（手入力）", value="", placeholder="例：東京都渋谷区神泉町13-12", key=f"address_{idx}")
                management_company = st.text_input("建物管理会社（手入力）", value="", key=f"mgmt_{idx}")
                total_units_text = st.text_input("総戸数（手入力）", value="", placeholder="例：33", key=f"units_{idx}")
                total_units = int(total_units_text) if total_units_text.strip().isdigit() else None
                st.caption("住所・総戸数・管理会社：空欄から手入力 / 最寄駅・駅徒歩：自動検索後に1つの欄で修正可能")

            with c3:
                st.markdown("**ローン概要**")
                mortgage_status = r.get("mortgage_status") or ("active" if r.get("active_mortgage_count", 0) else "none")
                base_company = (r.get("loan_company") or "").strip()
                if mortgage_status == "cancelled_current_owner":
                    default_company = f"{base_company}（抹消済）" if base_company else "過去抵当あり（抹消済）"
                elif mortgage_status == "none":
                    default_company = "抵当なし"
                else:
                    default_company = base_company
                loan_company = st.text_input("ローン会社", value=default_company, key=f"loan_{idx}")
                loan_amount = st.number_input("借入額", min_value=0, value=int(r.get("loan_amount_yen") or 0), step=10000, key=f"amount_{idx}")
                estimated_balance = st.number_input("推定残債", min_value=0, value=0 if mortgage_status != "active" else int(r.get("estimated_balance_yen") or 0), step=10000, key=f"balance_{idx}")

            with c4:
                st.markdown("**返済情報**")
                interest_rate = st.number_input("金利（%）", min_value=0.0, value=float(r.get("interest_rate_pct") if r.get("interest_rate_pct") is not None else 0.0), step=0.01, key=f"rate_{idx}")
                purchase_date = st.text_input("購入時期", value=r.get("purchase_date") or r.get("current_owner_acquired_date_slash") or "", key=f"purchase_date_{idx}")
                no_loan = mortgage_status != "active"
                monthly_payment = 0 if no_loan else calculate_monthly_payment(loan_amount, interest_rate or 2.3, 35)
                st.metric("毎月の返済額", f"{monthly_payment:,}円")
                building_age = _building_age(built_year)
                payoff_ym = _payoff_ym(purchase_date)
                two_months_later_ym = _ym_after_months(2)
                st.caption(f"築年数 {building_age if building_age != '' else '-'}年 ｜ 完済 {payoff_ym or '-'} ｜ 2か月後 {two_months_later_ym}")
                if mortgage_status == "cancelled_current_owner":
                    st.caption("現所有者の過去抵当は抹消済み → 過去の借入内容を表示 / 現在残債・返済額は0円 / Excel C16:C25相当をグレー表示")
                elif mortgage_status == "none":
                    st.caption("現所有者の抵当履歴なし → 抵当なし / 借入額・残債・返済額は0円")
                else:
                    st.caption(f"35年・元利均等 / 金利根拠: {r.get('interest_rate_source') or ''}")

            edited = dict(r)
            edited.update({
                "property_name": property_name, "room": room, "floors": floors,
                "area_sqm": area if area else None, "built_year": built_year,
                "property_address": property_address, "station_name": station_name,
                "walk_minutes": walk_minutes,
                "management_company": management_company, "total_units": total_units if total_units else None,
                "loan_company": loan_company, "loan_amount_yen": loan_amount,
                "estimated_balance_yen": estimated_balance, "interest_rate_pct": interest_rate,
                "purchase_date": purchase_date, "monthly_payment_yen": monthly_payment,
                "building_age": building_age, "payoff_ym": payoff_ym,
                "two_months_later_ym": two_months_later_ym,
            })
            edited_results.append(edited)

    st.divider()
    xlsx = fill_cs_template(edited_results)
    st.download_button(
        f"⬇ {len(edited_results)}物件を入力したCS Excelを作成",
        data=xlsx,
        file_name="CS_登記簿自動入力.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
    st.caption("PDF順：1件目=C列側 / 2件目=F列側 / 3件目=I列側 / 4件目=L列側 / 5件目=O列側 / 6件目=R列側 / 7件目=U列側")
