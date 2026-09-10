import io
import os
import re
import zipfile
from datetime import date
from xml.sax.saxutils import escape

import streamlit as st

from geo_lookup import enrich_location
from building_lookup import lookup_building_info
from touki_parser import analyze_registry, calculate_monthly_payment, extract_pdf_text

st.set_page_config(page_title="登記簿 → CS Excel（無料版）", page_icon="🏢", layout="wide")

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "sample", "CS_登記簿自動入力.xlsx")

# 1物件目から左→右へ、3列ずつずらして入力する。
SLOT_MAPS = [
    {
        "property_name": "C1", "floors": "C2", "built_year": "C3", "building_age": "D3", "room": "E2", "area_sqm": "E3",
        "property_address": "C4", "station_walk": "C5", "management_company": "C7", "total_units": "D2", "loan_company": "C8", "loan_amount_yen": "C9", "two_months_later_ym": "C10",
        "estimated_balance_yen": "C11", "interest_rate": "C14", "purchase_date": "C15", "payoff_ym": "C17", "monthly_payment_yen": "C19", "estimated_property_tax_yen": "C26",
    },
    {
        "property_name": "F1", "floors": "F2", "built_year": "F3", "building_age": "G3", "room": "H2", "area_sqm": "H3",
        "property_address": "F4", "station_walk": "F5", "management_company": "F7", "total_units": "G2", "loan_company": "F8", "loan_amount_yen": "F9", "two_months_later_ym": "F10",
        "estimated_balance_yen": "F11", "interest_rate": "F14", "purchase_date": "F15", "payoff_ym": "F17", "monthly_payment_yen": "F19", "estimated_property_tax_yen": "F26",
    },
    {
        "property_name": "I1", "floors": "I2", "built_year": "I3", "building_age": "J3", "room": "K2", "area_sqm": "K3",
        "property_address": "I4", "station_walk": "I5", "management_company": "I7", "total_units": "J2", "loan_company": "I8", "loan_amount_yen": "I9", "two_months_later_ym": "I10",
        "estimated_balance_yen": "I11", "interest_rate": "I14", "purchase_date": "I15", "payoff_ym": "I17", "monthly_payment_yen": "I19", "estimated_property_tax_yen": "I26",
    },
    {
        "property_name": "L1", "floors": "L2", "built_year": "L3", "building_age": "M3", "room": "N2", "area_sqm": "N3",
        "property_address": "L4", "station_walk": "L5", "management_company": "L7", "total_units": "M2", "loan_company": "L8", "loan_amount_yen": "L9", "two_months_later_ym": "L10",
        "estimated_balance_yen": "L11", "interest_rate": "L14", "purchase_date": "L15", "payoff_ym": "L17", "monthly_payment_yen": "L19", "estimated_property_tax_yen": "L26",
    },
]


def _cell_xml(ref, value, style=None, prefix=""):
    p = prefix
    s = f' s="{style}"' if style is not None else ""
    if value is None or value == "":
        return f'<{p}c r="{ref}"{s}/>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<{p}c r="{ref}"{s}><{p}v>{value}</{p}v></{p}c>'
    return f'<{p}c r="{ref}"{s} t="inlineStr"><{p}is><{p}t>{escape(str(value))}</{p}t></{p}is></{p}c>'


def _replace_cell(sheet_xml: str, ref: str, value, prefix="") -> str:
    p = re.escape(prefix)
    pat = rf'<{p}c\b([^>]*\br="{re.escape(ref)}"[^>]*)/>|<{p}c\b([^>]*\br="{re.escape(ref)}"[^>/]*)>(.*?)</{p}c>'
    m = re.search(pat, sheet_xml, flags=re.S)
    if not m:
        raise ValueError(f"テンプレート内にセル {ref} が見つかりませんでした。")
    attrs = (m.group(1) or m.group(2) or "")
    sm = re.search(r'\bs="([^"]+)"', attrs)
    style = sm.group(1) if sm else None
    new_xml = _cell_xml(ref, value, style, prefix)
    return sheet_xml[:m.start()] + new_xml + sheet_xml[m.end():]


def _ym_after_months(months: int, base=None) -> str:
    base = base or date.today()
    total = base.year * 12 + (base.month - 1) + months
    y, m0 = divmod(total, 12)
    return f"{y}/{m0 + 1}"


def _building_age(built_year):
    try:
        y = int(str(built_year).strip())
        return max(0, date.today().year - y)
    except Exception:
        return ""


def _payoff_ym(purchase_date: str, years: int = 35) -> str:
    s = (purchase_date or "").strip().replace("-", "/")
    m = re.match(r"^(\d{4})/(\d{1,2})(?:/(\d{1,2}))?$", s)
    if not m:
        return ""
    return f"{int(m.group(1)) + years}/{int(m.group(2))}"


def _estimate_property_tax_monthly(result: dict):
    """固定資産税の簡易月額推定。

    登記簿だけでは固定資産税評価額そのものは分からないため、
    現在または過去の抵当権の債権額を価格の代用値として使い、
    評価額70% × 固定資産税標準税率1.4% ÷ 12 で概算する。
    根拠額が取れない場合は空欄。
    """
    basis = int(result.get("tax_basis_amount_yen") or result.get("loan_amount_yen") or 0)
    if basis <= 0:
        return ""
    return int(round(basis * 0.70 * 0.014 / 12))


def _result_values(result: dict) -> dict:
    no_loan = result.get("active_mortgage_count", 0) == 0
    station = (result.get("station_name") or "").strip()
    walk = result.get("walk_minutes")
    if station and walk:
        station_walk = f"{station}駅 徒歩{int(walk)}分"
    elif walk:
        station_walk = f"徒歩{int(walk)}分"
    else:
        station_walk = station

    rate = result.get("interest_rate_pct")
    interest_text = "" if no_loan or rate is None else f"{float(rate):g}%"
    principal = 0 if no_loan else int(result.get("loan_amount_yen") or 0)
    monthly_payment = 0 if no_loan else int(result.get("monthly_payment_yen") or calculate_monthly_payment(principal, rate or 2.3, 35))

    return {
        "property_name": result.get("property_name") or "",
        "floors": int(result["floors"]) if str(result.get("floors") or "").isdigit() else (result.get("floors") or ""),
        "built_year": int(result["built_year"]) if str(result.get("built_year") or "").isdigit() else (result.get("built_year") or ""),
        "building_age": _building_age(result.get("built_year")),
        "room": result.get("room") or "",
        "area_sqm": result.get("area_sqm") if result.get("area_sqm") not in (None, "") else "",
        "property_address": result.get("property_address") or result.get("geocoded_address") or "",
        "station_walk": station_walk,
        "management_company": result.get("management_company") or "",
        "total_units": int(result["total_units"]) if result.get("total_units") not in (None, "") else "",
        "loan_company": "" if no_loan else (result.get("loan_company") or ""),
        "loan_amount_yen": principal,
        "two_months_later_ym": _ym_after_months(2),
        "estimated_balance_yen": 0 if no_loan else int(result.get("estimated_balance_yen") or 0),
        "interest_rate": interest_text,
        "purchase_date": result.get("purchase_date") or result.get("current_owner_acquired_date_slash") or "",
        "payoff_ym": _payoff_ym(result.get("purchase_date") or result.get("current_owner_acquired_date_slash") or ""),
        "monthly_payment_yen": monthly_payment,
        "estimated_property_tax_yen": result.get("estimated_property_tax_yen") if result.get("estimated_property_tax_yen") not in (None, "") else _estimate_property_tax_monthly(result),
    }


def fill_cs_template(results: list[dict]) -> bytes:
    """アップロード順に、同じExcelの左→右の枠へ最大4物件を書き込む。"""
    if len(results) > len(SLOT_MAPS):
        raise ValueError(f"このテンプレートは最大{len(SLOT_MAPS)}物件までです。")

    with open(TEMPLATE_PATH, "rb") as f:
        src = f.read()

    zin = zipfile.ZipFile(io.BytesIO(src), "r")
    sheet_name = "xl/worksheets/sheet1.xml"
    sheet_xml = zin.read(sheet_name).decode("utf-8")
    prefix = "x:" if "<x:worksheet" in sheet_xml else ""

    # テンプレートに残っている外部参照数式等を消し、4枠すべてを空の状態から作る。
    for slot in SLOT_MAPS:
        for ref in slot.values():
            sheet_xml = _replace_cell(sheet_xml, ref, "", prefix)

    for idx, result in enumerate(results):
        values = _result_values(result)
        for key, ref in SLOT_MAPS[idx].items():
            sheet_xml = _replace_cell(sheet_xml, ref, values[key], prefix)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = sheet_xml.encode("utf-8") if item.filename == sheet_name else zin.read(item.filename)
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
    st.caption("Ver.3.3 / ダークUI")
    st.divider()
    st.success("OpenAI APIキー不要 / 利用料0円")
    st.info("住所・最寄駅・管理会社・総戸数の検索時のみインターネット接続を使います。")
    st.warning("残債・毎月返済額は35年・元利均等。金利記載なしは2.3%。固定資産税は簡易推定です。")

st.markdown(
    '<div class="cs-upload-note">☁️ <b>登記簿PDFをアップロード</b>　最初のPDF→左端、2件目→その右…の順で最大4件まで入力できます。</div>',
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
    if len(uploaded_files) > 4:
        st.error("このExcelテンプレートは4物件分までです。PDFを4件以内にしてください。")
    elif st.button("解析＋物件情報検索を開始", type="primary", use_container_width=True):
        results = []
        progress = st.progress(0)
        status = st.empty()
        for i, uploaded in enumerate(uploaded_files, start=1):
            status.write(f"{i}/{len(uploaded_files)}件目：{uploaded.name} を解析中...")
            try:
                r = parse_uploaded(uploaded)
                loc = enrich_location(r.get("property_address", ""), r.get("property_name", ""))
                r.update(loc)
                building = lookup_building_info(r.get("property_name", ""), r.get("property_address", ""))
                r.update(building)
                if building.get("building_info_warning"):
                    r.setdefault("warnings", []).append(building["building_info_warning"])
                if loc.get("warning"):
                    r.setdefault("warnings", []).append(loc["warning"])
                results.append(r)
            except Exception as e:
                results.append({
                    "source_file": uploaded.name,
                    "warnings": [f"解析に失敗しました: {e}"],
                    "property_name": "", "room": "", "floors": "", "area_sqm": None,
                    "built_year": "", "loan_company": "", "loan_amount_yen": 0,
                    "property_address": "", "station_name": "", "walk_minutes": None,
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
        slot_col = ["C", "F", "I", "L"][idx]
        title = r.get("property_name") or r.get("source_file") or f"物件{idx+1}"
        with st.expander(f"物件{idx+1}　{title}　｜　Excel {slot_col}列側　● 解析済み", expanded=True):
            for w in r.get("warnings", []):
                st.warning(w)

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.markdown("**物件概要**")
                property_name = st.text_input("物件名", value=r.get("property_name") or "", key=f"name_{idx}")
                room = st.text_input("部屋番号", value=r.get("room") or "", key=f"room_{idx}")
                floors = st.text_input("階数", value=r.get("floors") or "", key=f"floors_{idx}")
                area = st.number_input("閉米数", min_value=0.0, value=float(r.get("area_sqm") or 0), step=0.01, key=f"area_{idx}")

            with c2:
                st.markdown("**建物情報**")
                built_year = st.text_input("築年", value=r.get("built_year") or "", key=f"year_{idx}")
                property_address = st.text_input("物件住所", value=r.get("property_address") or r.get("geocoded_address") or "", key=f"address_{idx}")
                station_name = st.text_input("最寄駅", value=r.get("station_name") or "", key=f"station_{idx}")
                walk_minutes = st.number_input("駅徒歩（分）", min_value=0, value=int(r.get("walk_minutes") or 0), step=1, key=f"walk_{idx}")
                management_company = st.text_input("建物管理会社", value=r.get("management_company") or "", key=f"mgmt_{idx}")
                total_units = st.number_input("総戸数", min_value=0, value=int(r.get("total_units") or 0), step=1, key=f"units_{idx}")
                if r.get("building_info_source_urls"):
                    st.caption("一般Web検索から取得")

            with c3:
                st.markdown("**ローン概要**")
                loan_company = st.text_input("ローン会社", value=r.get("loan_company") or "", key=f"loan_{idx}")
                loan_amount = st.number_input("借入額", min_value=0, value=int(r.get("loan_amount_yen") or 0), step=10000, key=f"amount_{idx}")
                estimated_balance = st.number_input("推定残債", min_value=0, value=int(r.get("estimated_balance_yen") or 0), step=10000, key=f"balance_{idx}")

            with c4:
                st.markdown("**返済・税金**")
                interest_rate = st.number_input("金利（%）", min_value=0.0, value=float(r.get("interest_rate_pct") if r.get("interest_rate_pct") is not None else 0.0), step=0.01, key=f"rate_{idx}")
                purchase_date = st.text_input("購入時期", value=r.get("purchase_date") or r.get("current_owner_acquired_date_slash") or "", key=f"purchase_date_{idx}")
                no_loan = r.get("active_mortgage_count", 0) == 0
                monthly_payment = 0 if no_loan else calculate_monthly_payment(loan_amount, interest_rate or 2.3, 35)
                st.metric("毎月の返済額", f"{monthly_payment:,}円")
                building_age = _building_age(built_year)
                payoff_ym = _payoff_ym(purchase_date)
                two_months_later_ym = _ym_after_months(2)
                default_tax = _estimate_property_tax_monthly({**r, "loan_amount_yen": loan_amount})
                estimated_property_tax = st.number_input(
                    "想定固定資産税（月額）", min_value=0,
                    value=int(default_tax or 0), step=100, key=f"tax_{idx}"
                )
                st.caption(f"築年数 {building_age if building_age != '' else '-'}年 ｜ 完済 {payoff_ym or '-'} ｜ 2か月後 {two_months_later_ym}")
                if no_loan:
                    st.caption("現在有効な抵当権なし → 借入額・残債・返済額は0円")
                else:
                    st.caption(f"35年・元利均等 / 金利根拠: {r.get('interest_rate_source') or ''}")

            edited = dict(r)
            edited.update({
                "property_name": property_name, "room": room, "floors": floors,
                "area_sqm": area if area else None, "built_year": built_year,
                "property_address": property_address, "station_name": station_name,
                "walk_minutes": walk_minutes if walk_minutes else None,
                "management_company": management_company, "total_units": total_units if total_units else None,
                "loan_company": loan_company, "loan_amount_yen": loan_amount,
                "estimated_balance_yen": estimated_balance, "interest_rate_pct": interest_rate,
                "purchase_date": purchase_date, "monthly_payment_yen": monthly_payment,
                "building_age": building_age, "payoff_ym": payoff_ym,
                "two_months_later_ym": two_months_later_ym,
                "estimated_property_tax_yen": estimated_property_tax if estimated_property_tax else "",
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
    st.caption("PDF順：1件目=C列側 / 2件目=F列側 / 3件目=I列側 / 4件目=L列側")
