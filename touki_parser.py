import io
import re
import unicodedata
import calendar
from datetime import date
from pypdf import PdfReader

ERA_BASE = {"明治": 1867, "大正": 1911, "昭和": 1925, "平成": 1988, "令和": 2018}


def extract_pdf_text(pdf_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pages = []
    for i, page in enumerate(reader.pages, start=1):
        txt = page.extract_text() or ""
        pages.append(f"\n--- PAGE {i} ---\n{txt}")
    return "\n".join(pages).strip()


def normalize(text: str) -> str:
    text = unicodedata.normalize("NFKC", text)
    # Remove private-use glyphs often emitted by registry PDFs.
    text = "".join(ch for ch in text if not (0xE000 <= ord(ch) <= 0xF8FF))
    out = []
    for line in text.splitlines():
        line = line.replace("　", " ")
        line = re.sub(r"[ \t]+", "", line)
        if line:
            out.append(line)
    return "\n".join(out)


def jp_date_to_iso(s: str):
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", s).replace("元年", "1年")
    m = re.search(r"(明治|大正|昭和|平成|令和)(\d+)年(\d+)月(\d+)日", s)
    if not m:
        return ""
    era, y, mo, d = m.groups()
    year = ERA_BASE[era] + int(y)
    try:
        return date(year, int(mo), int(d)).isoformat()
    except ValueError:
        return ""


def calculate_age_years(iso_date: str):
    if not iso_date:
        return None
    try:
        y, m, d = map(int, iso_date.split("-"))
        built = date(y, m, d)
    except Exception:
        return None
    today = date.today()
    return max(0, today.year - built.year - ((today.month, today.day) < (built.month, built.day)))


def _slice(text: str, start_key: str, end_keys):
    start = text.find(start_key)
    if start < 0:
        return ""
    end = len(text)
    for key in end_keys:
        p = text.find(key, start + len(start_key))
        if p >= 0:
            end = min(end, p)
    return text[start:end]


def _first_field(section: str, label: str):
    # Registry table rows usually use │ as cell separators.
    m = re.search(re.escape(label) + r"[│|]([^│|┃┠┗┏\n]+)", section)
    if m:
        return m.group(1).strip()
    # Conservative fallback for PDFs where separators are lost.
    m = re.search(re.escape(label) + r"([^\n]{1,80})", section)
    return m.group(1).strip("┃│| ") if m else ""


def _number(s: str):
    return int(re.sub(r"\D", "", s)) if re.search(r"\d", s or "") else None


def _entries(section: str):
    entries = []
    current = None
    for line in section.splitlines():
        m = re.match(r"^[┃|]?(\d+)[│|]", line)
        if m:
            if current:
                entries.append(current)
            current = {"rank": int(m.group(1)), "lines": [line]}
        elif current:
            current["lines"].append(line)
    if current:
        entries.append(current)
    for e in entries:
        e["text"] = "\n".join(e["lines"])
        e["flat"] = re.sub(r"[\n┃│|┠┗┏━─┬┴┼]+", "", e["text"])
    return entries


def _extract_amount_yen(flat: str):
    # Standard mortgage uses 債権額; revolving mortgage uses 極度額.
    for label in ("債権額", "極度額"):
        m = re.search(label + r"金?([0-9,，]+)(万円|円)", flat)
        if m:
            n = int(m.group(1).replace(",", "").replace("，", ""))
            return n * 10000 if m.group(2) == "万円" else n
    return None


def _extract_lender(flat: str):
    pos = -1
    for label in ("抵当権者", "根抵当権者"):
        p = flat.find(label)
        if p >= 0:
            pos = p + len(label)
            break
    if pos < 0:
        return ""

    tail = flat[pos:]
    tail = re.split(r"順位\d+番|共同担保|債務者|債権の範囲|確定期日", tail)[0]

    # 登記簿では「抵当権者 住所 会社名」の順で出ることが多い。
    # 例: 港区浜松町二丁目4番1号オリックス株式会社
    # → 住所を捨てて「オリックス株式会社」だけ返す。
    corporate_forms = ["株式会社", "有限会社", "合同会社", "合資会社", "合名会社"]

    # 前株: 株式会社○○
    for form in corporate_forms:
        m = re.search(form + r"[一-龥々ぁ-んァ-ヶーA-Za-z0-9・ー]+", tail)
        if m:
            return m.group(0).strip()

    # 後株: ○○株式会社 / ○○有限会社 など。
    for form in corporate_forms:
        idx = tail.find(form)
        if idx >= 0:
            before = tail[:idx]
            # 住所が「...号」で終わる場合は、その後だけを法人名とする。
            if "号" in before:
                before = before.rsplit("号", 1)[-1]
            else:
                # 「...番地」終わりにも対応。
                for marker in ("番地",):
                    if marker in before:
                        before = before.rsplit(marker, 1)[-1]
            # なお住所が残る場合に備え、末尾側の社名らしい連続文字だけを取得。
            mname = re.search(r"([一-龥々ぁ-んァ-ヶーA-Za-z0-9・ー]{1,40})$", before)
            if mname:
                return (mname.group(1) + form).strip()

    # 銀行・信金など法人格が株式会社表記でないケース。
    for suffix in ("銀行", "信用金庫", "信用組合", "労働金庫", "農業協同組合"):
        idx = tail.find(suffix)
        if idx >= 0:
            before = tail[:idx]
            if "号" in before:
                before = before.rsplit("号", 1)[-1]
            mname = re.search(r"([一-龥々ぁ-んァ-ヶーA-Za-z0-9・ー]{1,40})$", before)
            if mname:
                return (mname.group(1) + suffix).strip()

    if "住宅金融支援機構" in tail:
        return "住宅金融支援機構"
    return ""



def _extract_interest_rate(flat: str):
    """登記簿の「利息 年3・10％」等から年利(%)を取得。なければNone。"""
    m = re.search(r"利息年([0-9]+(?:[\.・][0-9]+)?)%", flat)
    if not m:
        # NFKC後も全角％や中黒が残るケースに備える
        m = re.search(r"利息年([0-9]+(?:[\.・][0-9]+)?)[％%]", flat)
    if not m:
        return None
    try:
        return float(m.group(1).replace("・", "."))
    except Exception:
        return None


def _extract_loan_date(entry: dict):
    """金銭消費貸借の原因日を優先し、取れなければ抵当権設定の受付日を返す。"""
    flat = entry.get("flat", "")
    text = entry.get("text", "")
    # 例: 原因平成11年6月25日金銭消費貸借同日設定
    m = re.search(r"原因((?:明治|大正|昭和|平成|令和)\d+年\d+月\d+日)[^\n]{0,30}?(?:金銭消費貸借|保証委託|設定)", flat)
    if m:
        iso = jp_date_to_iso(m.group(1))
        if iso:
            return iso
    # 原因日一般
    m = re.search(r"原因((?:明治|大正|昭和|平成|令和)\d+年\d+月\d+日)", flat)
    if m:
        iso = jp_date_to_iso(m.group(1))
        if iso:
            return iso
    # 受付年月日をフォールバック
    m = re.search(r"((?:明治|大正|昭和|平成|令和)\d+年\d+月\d+日)", text)
    return jp_date_to_iso(m.group(1)) if m else ""


def _months_elapsed(start_iso: str, as_of=None):
    if not start_iso:
        return 0
    try:
        y, m, d = map(int, start_iso.split("-"))
        start = date(y, m, d)
    except Exception:
        return 0
    today = as_of or date.today()
    months = (today.year - start.year) * 12 + (today.month - start.month)
    if today.day < start.day:
        months -= 1
    return max(0, months)


def _date_after_months(base: date, months: int) -> date:
    total = base.year * 12 + (base.month - 1) + months
    y, m0 = divmod(total, 12)
    m = m0 + 1
    d = min(base.day, calendar.monthrange(y, m)[1])
    return date(y, m, d)


def estimate_remaining_balance(principal_yen, loan_date_iso: str, annual_rate_pct: float, years: int = 35, as_of=None):
    """元利均等・35年ローン想定の推定残債。円単位で返す。"""
    if not principal_yen or principal_yen <= 0 or not loan_date_iso:
        return 0
    total_months = years * 12
    paid = _months_elapsed(loan_date_iso, as_of=as_of)
    if paid >= total_months:
        return 0
    principal = float(principal_yen)
    monthly_rate = float(annual_rate_pct) / 100.0 / 12.0
    if monthly_rate <= 0:
        payment = principal / total_months
        balance = principal - payment * paid
    else:
        payment = principal * monthly_rate / (1 - (1 + monthly_rate) ** (-total_months))
        balance = principal * (1 + monthly_rate) ** paid - payment * (((1 + monthly_rate) ** paid - 1) / monthly_rate)
    return max(0, int(round(balance)))



def calculate_monthly_payment(principal_yen, annual_rate_pct: float, years: int = 35):
    """元利均等・35年ローン想定の毎月返済額。円単位で返す。"""
    if not principal_yen or principal_yen <= 0:
        return 0
    total_months = years * 12
    principal = float(principal_yen)
    monthly_rate = float(annual_rate_pct or 0) / 100.0 / 12.0
    if monthly_rate <= 0:
        payment = principal / total_months
    else:
        payment = principal * monthly_rate / (1 - (1 + monthly_rate) ** (-total_months))
    return max(0, int(round(payment)))


def _extract_party_name(flat: str, label: str):
    """権利部の「所有者」「債務者」から氏名/法人名らしい末尾部分を取得。"""
    pos = flat.find(label)
    if pos < 0:
        return ""
    tail = flat[pos + len(label):]
    # 次の権利関係ラベルまでを対象にする。
    tail = re.split(r"抵当権者|根抵当権者|債権額|極度額|利息|損害金|共同担保|順位\d+番|住所変更", tail)[0]
    # 法人は法人格を含む名称を優先。
    for form in ("株式会社", "有限会社", "合同会社", "合資会社", "合名会社"):
        m = re.search(form + r"[一-龥々ぁ-んァ-ヶーA-Za-z0-9・ー]+", tail)
        if m:
            return m.group(0)
        idx = tail.find(form)
        if idx >= 0:
            before = tail[:idx]
            # 住所末尾の号/番地より後ろを会社名候補にする。
            for marker in ("号", "番地"):
                if marker in before:
                    before = before.rsplit(marker, 1)[-1]
            m2 = re.search(r"([一-龥々ぁ-んァ-ヶーA-Za-z0-9・ー]{1,50})$", before)
            if m2:
                return m2.group(1) + form
    # 個人名は住所の末尾（号/番地等）以降に残る日本語文字列を採用。
    cleaned = tail
    for marker in ("号", "番地"):
        if marker in cleaned:
            cleaned = cleaned.rsplit(marker, 1)[-1]
    # 数字・記号を除いた末尾の連続文字列。通常は「藤岡綱昭」「米谷俊輔」等。
    candidates = re.findall(r"[一-龥々ぁ-んァ-ヶー]{2,30}", cleaned)
    return candidates[-1] if candidates else ""


def _same_party(a: str, b: str) -> bool:
    def norm(x):
        return re.sub(r"[\s　・･,，.．\-ー]", "", unicodedata.normalize("NFKC", x or ""))
    aa, bb = norm(a), norm(b)
    return bool(aa and bb and aa == bb)

def analyze_registry(pdf_text: str) -> dict:
    t = normalize(pdf_text)
    warnings = []

    one = _slice(t, "表題部(一棟の建物の表示)", ["表題部(敷地権の目的である土地の表示)", "表題部(専有部分の建物の表示)"])
    unit = _slice(t, "表題部(専有部分の建物の表示)", ["表題部(敷地権の表示)", "権利部(甲区)"])
    kou = _slice(t, "権利部(甲区)", ["権利部(乙区)"])
    otsu = _slice(t, "権利部(乙区)", [])

    property_name = _first_field(one, "建物の名称")
    # 部屋番号は「建物の名称」ではなく、専有部分の「家屋番号」の末尾から取得する。
    # 例: 神泉町8番7の502 → 502
    house_number = _first_field(unit, "家屋番号")
    room = ""
    m_room = re.search(r"(?:の|ノ)(\d+)$", house_number or "")
    if m_room:
        room = m_room.group(1)
    elif house_number:
        # 区切りが崩れたPDF向けの保守的フォールバック。
        m_room = re.search(r"(\d+)$", house_number)
        room = m_room.group(1) if m_room else ""

    # 一棟の建物の「所在」を物件住所として取得。東京23区は登記簿で都道府県が
    # 省略されるため「東京都」を補う。
    property_address = _first_field(one, "所在")
    if property_address and re.match(r"^[^都道府県]{1,12}区", property_address):
        property_address = "東京都" + property_address

    floors = ""
    # 「何階建て」は構造欄の「◯階建」ではなく、床面積欄に実際に列挙された
    # 階数の最大値から算出する。PDFによって構造欄が崩れても拾いやすい。
    floor_area_matches = []
    for line in one.splitlines():
        # 例: 「1階592:54」「8階531:14」「8階 531.14」
        m_floor = re.search(r"(\d+)階(?=[^\n]{0,20}\d+[：:.]\d{1,2})", line)
        if m_floor:
            floor_area_matches.append(int(m_floor.group(1)))
    if floor_area_matches:
        floors = str(max(floor_area_matches))
    else:
        # 念のため従来の構造欄もフォールバックとして残す。
        floor_matches = re.findall(r"(\d+)階建", one)
        if floor_matches:
            floors = str(max(map(int, floor_matches)))

    # 「閉米数」列は専有部分の床面積ではなく、③敷地権の割合の分子を
    # 小数点2桁として出力する。
    # 例: 68248分の2083 → 20.83
    area = None
    land_right = _slice(t, "表題部(敷地権の表示)", ["所有者", "権利部(甲区)"])
    m = re.search(r"(\d+)分の(\d+)", land_right)
    if m:
        numerator = int(m.group(2))
        area = numerator / 100.0

    # 「築年数」列は経過年数ではなく、表題部の「原因及びその日付」から西暦年だけを出す。
    # 例: 平成11年5月17日 → 1999
    built_date = ""
    built_year = ""
    m = re.search(r"((?:明治|大正|昭和|平成|令和)\d+年\d+月\d+日)(?:新築|変更|増築)?", unit)
    if m:
        built_date = jp_date_to_iso(m.group(1))
        if built_date:
            built_year = built_date[:4]

    acquired_date = ""
    current_owner_name = ""
    owner_entries = []
    for e in _entries(kou):
        if re.search(r"所有権(?:移転|保存)|持分(?:全部|一部)?移転", e["flat"]):
            dm = re.search(r"原因((?:明治|大正|昭和|平成|令和)\d+年\d+月\d+日)", e["flat"])
            owner_entries.append({
                "rank": e["rank"],
                "date": jp_date_to_iso(dm.group(1)) if dm else "",
                "name": _extract_party_name(e["flat"], "所有者"),
            })
    if owner_entries:
        latest_owner = sorted(owner_entries, key=lambda x: x["rank"])[-1]
        acquired_date = latest_owner["date"]
        current_owner_name = latest_owner["name"]

    mortgage_entries = _entries(otsu)
    settings = {}
    cancelled = set()
    for e in mortgage_entries:
        flat = e["flat"]
        if "抵当権設定" in flat or "根抵当権設定" in flat:
            settings[e["rank"]] = e
        for cm in re.finditer(r"(?:^|[│|┃])(\d+)番(?:根)?抵当権(?:登記)?抹消", e["text"]):
            cancelled.add(int(cm.group(1)))

    active = [e for rank, e in sorted(settings.items()) if rank not in cancelled]
    # 固定資産税の概算用。完済・抹消済みでも過去の抵当権額を参考値として残す。
    tax_basis_amount = 0
    if settings:
        latest_setting = sorted(settings.items())[-1][1]
        tax_basis_amount = _extract_amount_yen(latest_setting["flat"]) or 0

    loan_company = ""
    loan_amount = None
    loan_date = ""
    interest_rate = None
    interest_rate_source = ""
    estimated_balance = 0
    monthly_payment = 0
    mortgage_details = []

    if active:
        # 画面の単一項目には最新順位の抵当権を表示。残債は有効な抵当権すべてを合算。
        chosen = active[-1]
        loan_company = _extract_lender(chosen["flat"])
        loan_amount = _extract_amount_yen(chosen["flat"])
        loan_date = _extract_loan_date(chosen)
        explicit_rate = _extract_interest_rate(chosen["flat"])
        interest_rate = explicit_rate if explicit_rate is not None else 2.3
        interest_rate_source = "登記記載" if explicit_rate is not None else "既定2.3%"

        total_balance = 0
        total_monthly_payment = 0
        for e in active:
            amount = _extract_amount_yen(e["flat"]) or 0
            date_iso = _extract_loan_date(e)
            rate_found = _extract_interest_rate(e["flat"])
            rate = rate_found if rate_found is not None else 2.3
            bal = estimate_remaining_balance(amount, date_iso, rate, years=35, as_of=_date_after_months(date.today(), 3))
            payment = calculate_monthly_payment(amount, rate, years=35)
            total_balance += bal
            total_monthly_payment += payment
            mortgage_details.append({
                "rank": e["rank"],
                "company": _extract_lender(e["flat"]),
                "amount_yen": amount,
                "loan_date": date_iso,
                "interest_rate": rate,
                "interest_rate_source": "登記記載" if rate_found is not None else "既定2.3%",
                "estimated_balance_yen": bal,
                "monthly_payment_yen": payment,
            })
        estimated_balance = total_balance
        monthly_payment = total_monthly_payment
        if len(active) > 1:
            warnings.append(f"有効な抵当権が{len(active)}件あります。推定残債はExcelの「2か月後」欄に入る月時点で全件を合算しています。ローン会社等の表示は最新順位です。")
    else:
        # 現在有効な抵当権がない場合、現所有者本人が過去に設定した抵当権だけを
        # 「過去抵当あり（抹消済み）」として表示する。前所有者の抵当は引き継がない。
        estimated_balance = 0
        monthly_payment = 0
        historical_current_owner = []
        for rank, e in sorted(settings.items()):
            if rank not in cancelled:
                continue
            debtor = _extract_party_name(e["flat"], "債務者")
            setting_date = _extract_loan_date(e)
            # 名義一致を主条件にし、取得日前の設定は除外する。
            if _same_party(debtor, current_owner_name) and (not acquired_date or not setting_date or setting_date >= acquired_date):
                historical_current_owner.append(e)

        if historical_current_owner:
            chosen = historical_current_owner[-1]
            loan_company = _extract_lender(chosen["flat"])
            loan_amount = _extract_amount_yen(chosen["flat"]) or 0
            loan_date = _extract_loan_date(chosen)
            explicit_rate = _extract_interest_rate(chosen["flat"])
            interest_rate = explicit_rate if explicit_rate is not None else 2.3
            interest_rate_source = "登記記載（抹消済み）" if explicit_rate is not None else "既定2.3%（抹消済み）"
            mortgage_details.append({
                "rank": chosen["rank"], "company": loan_company, "amount_yen": loan_amount,
                "loan_date": loan_date, "interest_rate": interest_rate,
                "interest_rate_source": interest_rate_source,
                "estimated_balance_yen": 0, "monthly_payment_yen": 0, "status": "cancelled_current_owner",
            })
            warnings.append("現所有者が過去に設定した抵当権は抹消済みです。過去の借入内容を表示し、現在残債は0円としています。")
        elif settings:
            warnings.append("現在有効な抵当権はありません。過去の抵当権は現所有者のものではないため、ローン情報には使用しません。")

    checks = {
        "物件名": bool(property_name), "号室": bool(room), "何階建て": bool(floors),
        "平米数": area is not None, "新築年月日": bool(built_date),
        "購入年月日": bool(acquired_date)
    }
    missing = [k for k, ok in checks.items() if not ok]
    if missing:
        warnings.append("自動取得できなかった項目: " + "、".join(missing) + "。PDF抽出テキストと照合してください。")

    return {
        "property_name": property_name,
        "property_address": property_address,
        "room": room,
        "floors": floors,
        "area_sqm": area,
        "built_date": built_date,
        "built_year": built_year,
        "age_years": calculate_age_years(built_date),
        "current_owner_acquired_date": acquired_date,
        "current_owner_acquired_date_slash": (lambda parts: f"{int(parts[0])}/{int(parts[1])}/{int(parts[2])}")(acquired_date.split("-")) if acquired_date else "",
        "purchase_date": (lambda parts: f"{int(parts[0])}/{int(parts[1])}/{int(parts[2])}")(acquired_date.split("-")) if acquired_date else "",
        "current_owner_name": current_owner_name,
        "loan_company": loan_company,
        "loan_amount_yen": loan_amount,
        "tax_basis_amount_yen": tax_basis_amount,
        "loan_date": (lambda parts: f"{int(parts[0])}/{int(parts[1])}/{int(parts[2])}")(loan_date.split("-")) if loan_date else "",
        "interest_rate_pct": interest_rate,
        "interest_rate_source": interest_rate_source,
        "estimated_balance_yen": estimated_balance,
        "monthly_payment_yen": monthly_payment,
        "mortgage_details": mortgage_details,
        "active_mortgage_count": len(active),
        "mortgage_status": "active" if active else ("cancelled_current_owner" if mortgage_details else "none"),
        "cancelled_mortgage_ranks": sorted(cancelled),
        "warnings": warnings,
    }
