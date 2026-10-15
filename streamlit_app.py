import hashlib
import hmac
import io
import re
import secrets
import sqlite3
from contextlib import contextmanager
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

import streamlit.components.v1 as components


st.set_page_config(page_title="Ruang Anggaran — Dashboard RAB & TOR",page_icon="balmon.png", layout="wide")
ROOT = Path(__file__).resolve().parent
IDENTITY_COMPONENT = components.declare_component(
    "balmon_identity_check",
    path=str(ROOT / "identity_component"),
)
COLORS = ["#092759", "#087cb5", "#3d91b8", "#7a9bb5", "#a8bbc9", "#c3ccd4"]
MAX_UPLOAD_SIZE = 20 * 1024 * 1024
ACCOUNT_DB_PATH = ROOT / ".data" / "accounts.sqlite3"


def clean_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def to_number(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value) if pd.notna(value) else None
    text = re.sub(r"[^\d,.-]", "", str(value)).strip()
    if not text or re.fullmatch(r"-+", text):
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        head, tail = text.rsplit(",", 1)
        text = head + tail if len(tail) == 3 and head.lstrip("-").isdigit() else head + "." + tail
    elif "." in text:
        if re.fullmatch(r"-?\d{1,3}(?:\.\d{3})+", text):
            text = text.replace(".", "")
        else:
            head, tail = text.rsplit(".", 1)
            text = head + tail if len(tail) == 3 and head.lstrip("-").isdigit() else text
    try:
        value = float(text)
        return value if pd.notna(value) else None
    except ValueError:
        return None


def trim_rows(rows):
    result = [list(row) for row in rows]
    while result and all(value is None or clean_text(value) == "" for value in result[-1]):
        result.pop()
    return result


def split_text_rows(text):
    rows = []
    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if not line.strip():
            continue
        delimiter = "\t" if "\t" in line else re.compile(r" {2,}")
        rows.append([part.strip() or None for part in delimiter.split(line.strip())])
    return trim_rows(rows)


def parse_sp2d_rows(rows):
    joined = " ".join(clean_text(value) for row in rows for value in row)
    is_sp2d = (
        re.search(r"LAPORAN REALISASI SP2D", joined, re.I)
        and any(re.fullmatch(r"Lock Pagu", clean_text(value), re.I) for row in rows for value in row)
        and any(re.fullmatch(r"Uraian", clean_text(value), re.I) for row in rows for value in row)
    )
    if not is_sp2d:
        return None
    normalized = [["Kode Anggaran", "Uraian", "Pagu", "Sisa Anggaran"]]
    for row in rows:
        padded = list(row) + [None] * max(0, 9 - len(row))
        code = clean_text(padded[0])
        description = clean_text(padded[2])
        amount = to_number(padded[8])
        remaining = to_number(padded[6])
        if re.fullmatch(r"(?:\d{3}|\d{6})", code) and description and amount and amount > 0:
            if remaining is not None and remaining >= 0:
                normalized.append([code, description, amount, remaining])
    if len(normalized) == 1:
        raise ValueError("Kolom pagu dan sisa pada laporan SP2D tidak dapat dipastikan.")
    return normalized


def parse_upload(name, content):
    extension = Path(name).suffix.lower()
    if len(content) > MAX_UPLOAD_SIZE:
        raise ValueError("Ukuran file melebihi batas 20 MB.")
    if extension in (".xlsx", ".xlsm", ".xls"):
        engine = "xlrd" if extension == ".xls" else "openpyxl"
        workbook = pd.ExcelFile(io.BytesIO(content), engine=engine)
        sheets = {}
        for sheet_name in workbook.sheet_names:
            frame = pd.read_excel(
                io.BytesIO(content),
                sheet_name=sheet_name,
                header=None,
                engine=engine,
                dtype=object,
            )
            sheets[sheet_name] = trim_rows(frame.where(pd.notna(frame), None).values.tolist())
        return sheets
    if extension == ".csv":
        text = None
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                text = content.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError("CSV tidak menggunakan encoding teks yang didukung.")
        try:
            frame = pd.read_csv(io.StringIO(text), header=None, sep=None, engine="python", dtype=object)
        except pd.errors.ParserError as error:
            raise ValueError("Struktur kolom CSV tidak dapat dibaca.") from error
        return {"Data RAB": trim_rows(frame.where(pd.notna(frame), None).values.tolist())}
    if extension == ".docx":
        doc = Document(io.BytesIO(content))
        rows = []
        for table in doc.tables:
            rows.extend([[cell.text.strip() or None for cell in row.cells] for row in table.rows])
            rows.append([])
        if not rows:
            rows = split_text_rows("\n".join(paragraph.text for paragraph in doc.paragraphs))
        return {"Data RAB": trim_rows(rows)}
    if extension == ".pdf":
        try:
            import pdfplumber
        except ImportError as error:
            raise ValueError("Dukungan PDF belum tersedia. Pasang dependensi yang tercantum di requirements.txt.") from error
        pages_text = []
        table_rows = []
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables() or []:
                    table_rows.extend(table)
                pages_text.append(page.extract_text(layout=False) or "")
        rows = trim_rows(table_rows) if table_rows else split_text_rows("\n".join(pages_text))
        sp2d = parse_sp2d_rows(rows)
        if sp2d:
            return {"Realisasi SP2D": sp2d}
        if not rows:
            raise ValueError("PDF tidak memiliki teks yang dapat dibaca. PDF hasil pindai perlu OCR terlebih dahulu.")
        header = find_header(rows)
        if not header:
            raise ValueError("Struktur tabel RAB pada PDF tidak dapat dipastikan. Simpan PDF dengan teks/kolom terpisah atau unggah Excel/CSV.")
        return {"Data RAB": rows}
    if extension == ".doc":
        raise ValueError("Format Word .doc lama belum didukung di Streamlit. Simpan ulang sebagai .docx lalu unggah.")
    raise ValueError("Format belum didukung. Gunakan Excel, CSV, Word .docx, atau PDF berbasis teks.")


def find_header(rows):
    best = None
    for row_index, row in enumerate(rows[:40]):
        labels = [clean_text(value) for value in row]
        normalized = [label.lower() for label in labels]
        desc_index = next(
            (
                index for index, label in enumerate(normalized)
                if re.fullmatch(
                    r"uraian(?: kegiatan)?|deskripsi(?: kegiatan)?|nama kegiatan|rincian(?: kegiatan)?|kegiatan|belanja",
                    label,
                )
            ),
            -1,
        )
        if desc_index < 0:
            continue
        amount_matches = sum(bool(re.search(r"jumlah|pagu|anggaran|nilai|sisa", label)) for label in normalized)
        score = amount_matches * 2 + sum(bool(label) for label in normalized)
        if best is None or score >= best["score"]:
            best = {"row_index": row_index, "labels": labels, "normalized": normalized, "description": desc_index, "score": score}
    return best


def find_revision(rows):
    layout = None
    for marker_index, row in enumerate(rows[:40]):
        markers = [clean_text(value).lower() for value in row]
        original_start = markers.index("semula") if "semula" in markers else -1
        revised_start = markers.index("menjadi") if "menjadi" in markers else -1
        if original_start < 0 or revised_start <= original_start:
            continue
        for header_index in range(marker_index + 1, min(len(rows), marker_index + 5)):
            labels = [clean_text(value).lower() for value in rows[header_index]]
            original_desc = next((i for i, label in enumerate(labels) if original_start <= i < revised_start and label == "uraian"), -1)
            revised_desc = next((i for i, label in enumerate(labels) if i >= revised_start and label == "uraian"), -1)
            delta_col = next((i for i, label in enumerate(labels) if i >= revised_start and label == "selisih"), -1)
            remaining_col = next((i for i, label in enumerate(labels) if label == "sisa anggaran"), -1)
            orig_amounts = [
                i for i, label in enumerate(labels)
                if original_start <= i < revised_start and re.sub(r"\s+", " ", label).strip() == "jumlah"
            ]
            original_code = next(
                (i for i, label in enumerate(labels) if original_start <= i < revised_start and label == "kode anggaran"),
                -1,
            )
            revised_code = next(
                (i for i, label in enumerate(labels) if i >= revised_start and label == "kode anggaran"),
                -1,
            )
            revised_amount = next(
                (i for i, label in enumerate(labels) if revised_start <= i < delta_col and label == "jumlah"),
                -1,
            ) if delta_col >= 0 else -1
            if original_desc >= 0 and revised_desc >= 0 and delta_col >= 0 and remaining_col >= 0 and orig_amounts and revised_amount >= 0:
                original_amount = max(
                    orig_amounts,
                    key=lambda column: sum(to_number(row[column] if column < len(row) else None) is not None for row in rows[header_index + 1:]),
                )
                layout = {
                    "first": header_index + 1,
                    "original_code": original_code,
                    "original_desc": original_desc,
                    "original_amount": original_amount,
                    "revised_code": revised_code,
                    "revised_desc": revised_desc,
                    "revised_amount": revised_amount,
                    "delta": delta_col,
                    "remaining": remaining_col,
                }
                break
        if layout:
            break
    if not layout:
        return None

    changes = []
    for row_index in range(layout["first"], len(rows)):
        row = rows[row_index]
        def get(column):
            return row[column] if column >= 0 and column < len(row) else None

        original_desc = clean_text(get(layout["original_desc"]))
        revised_desc = clean_text(get(layout["revised_desc"]))
        if get(layout["original_code"]) or get(layout["revised_code"]) or (not original_desc and not revised_desc):
            continue
        original_amount = to_number(get(layout["original_amount"]))
        revised_amount = to_number(get(layout["revised_amount"]))
        delta = to_number(get(layout["delta"]))
        if delta is None and original_amount is not None and revised_amount is not None:
            delta = revised_amount - original_amount
        if delta is None or delta == 0:
            continue
        remaining = to_number(get(layout["remaining"]))
        status = "decrease" if delta < 0 else "unknown" if remaining is None else "enough" if remaining >= delta else "short"
        changes.append({
            "Uraian semula": original_desc or "—",
            "Pagu semula": original_amount,
            "Uraian menjadi": revised_desc or "—",
            "Pagu menjadi": revised_amount,
            "Selisih": delta,
            "Sisa anggaran": remaining,
            "Kecukupan": {"decrease": "Pengurangan", "unknown": "Cek sisa", "enough": "Sisa cukup", "short": "Sisa kurang"}[status],
            "_status": status,
        })
    increases = [row for row in changes if row["Selisih"] > 0]
    decreases = [row for row in changes if row["Selisih"] < 0]
    insufficient = [row for row in increases if row["_status"] == "short"]
    unknown = [row for row in increases if row["_status"] == "unknown"]
    return {
        "changes": changes,
        "increases": increases,
        "decreases": decreases,
        "increase_total": sum(row["Selisih"] for row in increases),
        "decrease_total": sum(abs(row["Selisih"]) for row in decreases),
        "sufficient_count": sum(row["_status"] == "enough" for row in increases),
        "insufficient_count": len(insufficient),
        "unknown_count": len(unknown),
        "shortfall": sum(row["Selisih"] - row["Sisa anggaran"] for row in insufficient),
    }


def analyze_rows(rows, sheet_name):
    header = find_header(rows)
    revision = find_revision(rows)
    if not header:
        return {
            "sheet": sheet_name,
            "header_found": False,
            "revision": revision,
            "rows": [],
            "categories": [],
            "columns": [],
            "pagu_total": None,
            "remaining_total": None,
            "source_count": len(rows),
            "excluded_count": len(rows),
            "grouping": "",
            "insights": ["Struktur tabel belum dikenali. Pastikan ada kolom uraian/kegiatan dan angka anggaran."],
        }

    header_row = header["row_index"]
    labels = header["labels"]
    normalized = header["normalized"]
    start = 0
    for previous_row in rows[max(0, header_row - 3):header_row]:
        marker = next((i for i, value in enumerate(previous_row) if clean_text(value).lower() == "menjadi"), -1)
        if marker >= 0:
            start = marker

    columns = [(i, label) for i, label in enumerate(labels) if i >= start and label]
    description_col = next((i for i, label in columns if re.search(r"uraian|deskripsi|kegiatan|rincian|belanja", label.lower())), header["description"])
    code_col = next((i for i, label in columns if re.search(r"kode anggaran|^kode$|^akun$", label.lower())), -1)
    amount_candidates = [(i, label) for i, label in columns if re.search(r"jumlah|pagu|total anggaran|nilai anggaran", label.lower())]
    money_col = amount_candidates[-1][0] if amount_candidates else -1
    remaining_candidates = [(i, label) for i, label in columns if re.search(r"sisa anggaran|^sisa$", label.lower())]
    remaining_col = remaining_candidates[-1][0] if remaining_candidates else -1
    prepared = []
    source_count = max(0, len(rows) - header_row - 1)
    for row_index, raw in enumerate(rows[header_row + 1:], header_row + 2):
        row = list(raw)
        description = clean_text(row[description_col] if description_col < len(row) else None)
        code = clean_text(row[code_col] if code_col >= 0 and code_col < len(row) else None)
        amount = to_number(row[money_col] if money_col >= 0 and money_col < len(row) else None)
        remaining = to_number(row[remaining_col] if remaining_col >= 0 and remaining_col < len(row) else None)
        if not description or re.fullmatch(r"\d+(?:[.,]\d+)?", description):
            continue
        if description.lower() in ("uraian", "kode anggaran", "vol.", "satuan", "harga satuan", "jumlah", "pagu", "sisa anggaran"):
            continue
        if amount is None and remaining is None:
            continue
        prepared.append({"Kode": code, "Uraian": description, "Pagu": amount, "Sisa anggaran": remaining, "Baris": row_index, "_raw": row})

    positive = [row["Pagu"] for row in prepared if row["Pagu"] is not None and row["Pagu"] > 0]
    pagu_total = max(positive) if positive else None
    detailed = [
        row for row in prepared
        if row["Pagu"] is not None and row["Pagu"] > 0
        and (pagu_total is None or row["Pagu"] < pagu_total * 0.9)
        and (row["Uraian"].startswith(("-", "•")) or row["Kode"])
    ]
    ranked = sorted(detailed or [row for row in prepared if row["Pagu"] and row["Pagu"] > 0], key=lambda row: row["Pagu"], reverse=True)[:5]
    activity = [row for row in prepared if re.fullmatch(r"\d{3}", row["Kode"]) and row["Pagu"] and row["Pagu"] > 0]
    account = [row for row in prepared if re.fullmatch(r"\d{6}", row["Kode"]) and row["Pagu"] and row["Pagu"] > 0]
    categories = []
    grouping = "kelompok kegiatan"
    if len(activity) >= 2:
        categories = [{"Kode": row["Kode"], "Kategori": row["Kode"] + " · " + row["Uraian"], "Pagu": row["Pagu"], "Sisa": row["Sisa anggaran"]} for row in activity]
    elif account:
        grouped = {}
        for row in account:
            category = grouped.setdefault(row["Kode"], {"Kode": row["Kode"], "Kategori": row["Kode"] + " · " + row["Uraian"], "Pagu": 0, "Sisa": 0, "_sisa_count": 0, "_count": 0})
            category["Pagu"] += row["Pagu"]
            category["_count"] += 1
            if row["Sisa anggaran"] is not None:
                category["Sisa"] += row["Sisa anggaran"]
                category["_sisa_count"] += 1
        categories = [
            {key: value for key, value in category.items() if not key.startswith("_")}
            for category in grouped.values()
        ]
        for category, source in zip(categories, grouped.values()):
            if source["_sisa_count"] != source["_count"]:
                category["Sisa"] = None
        grouping = "akun belanja"
    else:
        categories = [{"Kode": row["Kode"], "Kategori": row["Uraian"], "Pagu": row["Pagu"], "Sisa": row["Sisa anggaran"]} for row in detailed if row["Pagu"] and row["Pagu"] > 0]
        grouping = "rincian belanja"
    categories.sort(key=lambda row: row["Pagu"], reverse=True)
    category_total = sum(row["Pagu"] for row in categories)
    remaining_categories = [row for row in categories if row["Sisa"] is not None]
    category_remaining_total = sum(row["Sisa"] for row in remaining_categories) if remaining_categories else None

    insights = []
    if revision and revision["increases"]:
        if revision["insufficient_count"]:
            insights.append("{} rincian kenaikan memiliki sisa lebih kecil dari selisih; kekurangan terhitung {}.".format(
                revision["insufficient_count"], format_rupiah(revision["shortfall"])
            ))
        elif revision["unknown_count"]:
            insights.append("{} rincian naik, namun {} rincian belum memiliki nilai sisa yang bisa dibandingkan.".format(
                len(revision["increases"]), revision["unknown_count"]
            ))
        else:
            insights.append("Sisa pada semua {} rincian kenaikan yang terbaca mencukupi.".format(len(revision["increases"])))
    if categories and category_total:
        top = categories[0]
        insights.append("{} merupakan porsi pagu terbesar ({:.0f}%).".format(top["Kategori"], top["Pagu"] / category_total * 100))
        if category_remaining_total is not None:
            insights.append("Sisa yang terbaca pada kategori terpilih berjumlah {}.".format(format_rupiah(category_remaining_total)))
        if any(row["Sisa"] is not None and (row["Sisa"] < 0 or row["Sisa"] > row["Pagu"]) for row in categories):
            insights.append("Ada nilai sisa di luar rentang pagu; periksa perubahan anggaran atau formula sumber.")
    elif not categories:
        insights.append("Belum ada kategori bernilai yang bisa dipetakan untuk grafik.")

    return {
        "sheet": sheet_name,
        "header_found": True,
        "revision": revision,
        "rows": prepared,
        "ranked": ranked,
        "categories": categories,
        "columns": columns,
        "money_col": money_col,
        "remaining_col": remaining_col,
        "pagu_total": pagu_total,
        "remaining_total": max([row["Sisa anggaran"] for row in prepared if row["Sisa anggaran"] is not None], default=None),
        "category_total": category_total,
        "category_remaining_total": category_remaining_total,
        "grouping": grouping,
        "source_count": source_count,
        "excluded_count": source_count - len(prepared),
        "insights": insights[:4],
    }


def format_rupiah(value):
    if value is None:
        return "—"
    return "Rp {:,.0f}".format(value).replace(",", ".")


def format_compact_rupiah(value):
    if value is None:
        return "—"
    amount = abs(value)
    if amount >= 1_000_000_000:
        scaled, suffix = value / 1_000_000_000, " m"
    elif amount >= 1_000_000:
        scaled, suffix = value / 1_000_000, " jt"
    elif amount >= 1_000:
        scaled, suffix = value / 1_000, " rb"
    else:
        return "Rp {:,.0f}".format(value).replace(",", ".")
    number = "{:,.1f}".format(scaled).replace(",", "X").replace(".", ",").replace("X", ".")
    if number.endswith(",0"):
        number = number[:-2]
    return number + suffix


def hash_password(password, salt=None):
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 180000)
    return salt.hex(), digest.hex()


def password_matches(password, salt_hex, digest_hex):
    _, candidate = hash_password(password, bytes.fromhex(salt_hex))
    return hmac.compare_digest(candidate, digest_hex)


@contextmanager
def account_connection():
    ACCOUNT_DB_PATH.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    ACCOUNT_DB_PATH.parent.chmod(0o700)
    connection = sqlite3.connect(ACCOUNT_DB_PATH, timeout=10)
    try:
        with connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS accounts (
                    email TEXT PRIMARY KEY,
                    salt TEXT NOT NULL,
                    password_hash TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            yield connection
    finally:
        connection.close()
        ACCOUNT_DB_PATH.chmod(0o600)


def get_account(email):
    with account_connection() as connection:
        row = connection.execute(
            "SELECT salt, password_hash FROM accounts WHERE email = ?",
            (email,),
        ).fetchone()
    if row is None:
        return None
    return {"salt": row[0], "digest": row[1]}


def create_account(email, salt, digest):
    with account_connection() as connection:
        cursor = connection.execute(
            "INSERT OR IGNORE INTO accounts (email, salt, password_hash) VALUES (?, ?, ?)",
            (email, salt, digest),
        )
        return cursor.rowcount == 1


def logout():
    for key in ("authenticated_email", "upload_name", "sheets", "selected_sheet", "tor_result"):
        st.session_state.pop(key, None)


def render_auth():
    st.markdown(
        """
        <style>
        .stApp { background: #f5f8fb; }
        .st-key-auth_card { width: min(100%, 1100px); margin: clamp(1.5rem, 6vh, 4rem) auto 0;
            padding: clamp(1.5rem, 4vw, 3rem);
            background: #1e1e24; border: 1px solid #303038; border-radius: 12px; }
        .st-key-auth_card [data-testid="stMarkdownContainer"] h1 { color: #fff; font-size: clamp(2rem, 3vw, 2.7rem); }
        .st-key-auth_card [data-testid="stMarkdownContainer"] p,
        .st-key-auth_card [data-testid="stWidgetLabel"] p,
        .st-key-auth_card [data-testid="stCaptionContainer"] p { color: #c5c7ce; }
        .st-key-auth_card [data-testid="stTextInput"] input { background: #38383f; color: #f4f5f7; border-color: #4a4a53; }
        .st-key-auth_card [data-testid="stRadio"] label { color: #f2f3f5; }
        .st-key-auth_card [data-testid="stBaseButton-primary"] { background: #087cb5; border-color: #087cb5; }
        .st-key-auth_card [data-testid="stForm"] { border: 0; padding: 0; }
        @media (max-width: 700px) {
            .st-key-auth_card { margin-top: 1rem; padding: 1.25rem; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    left, center, right = st.columns([0.1, 4, 0.1])
    with center:
        with st.container(key="auth_card"):
            brand, form_panel = st.columns([0.9, 1.1], gap="large", vertical_alignment="center")
            with brand:
                logos = st.columns([1, 1, 1], gap="small")
                for column, filename in zip(logos, ("balmon.png", "djid.png", "komdigi.png")):
                    path = ROOT / "image" / filename
                    if path.exists():
                        column.image(str(path), use_container_width=True)
                st.markdown("<h1>RAB &amp; TOR</h1>", unsafe_allow_html=True)
                st.markdown("Masuk untuk membaca anggaran, melihat insight, dan menyusun draft TOR.")
                st.caption("Foto kartu dan selfie diproses di browser. Pemeriksaan demo bukan bukti identitas atau akses pegawai; akun hanya aktif selama sesi.")

            with form_panel:
                mode = st.radio("Akses demo", ("Masuk", "Daftar"), horizontal=True, label_visibility="collapsed")
                if mode == "Masuk":
                    with st.form("login_form"):
                        email = st.text_input("Email", key="login_email", autocomplete="email")
                        password = st.text_input("Kata sandi", type="password", key="login_password")
                        submitted = st.form_submit_button("Masuk", type="primary", use_container_width=True)
                    if submitted:
                        normalized_email = email.strip().lower()
                        account = get_account(normalized_email)
                        if account and password_matches(password, account["salt"], account["digest"]):
                            st.session_state.authenticated_email = normalized_email
                            st.rerun()
                        else:
                            st.error("Email atau kata sandi tidak cocok.")
                else:
                    with st.form("signup_form"):
                        email = st.text_input("Email", key="signup_email", autocomplete="email")
                        password = st.text_input("Kata sandi", type="password", key="signup_password")
                        employee_number = st.text_input("Nomor pegawai pada kartu", max_chars=24)
                        submitted = st.form_submit_button("Lanjut verifikasi kartu", type="primary", use_container_width=True)
                    if submitted:
                        if not re.fullmatch(r"\d{5,24}", re.sub(r"\D", "", employee_number)):
                            st.error("Nomor pegawai harus berisi sedikitnya 5 angka.")
                        elif "@" not in email or len(password) < 8:
                            st.error("Masukkan email yang valid dan kata sandi minimal 8 karakter.")
                        elif get_account(email.strip().lower()):
                            st.error("Email sudah terdaftar.")
                        else:
                            salt, digest = hash_password(password)
                            st.session_state.pending_signup = {
                                "email": email.strip().lower(),
                                "employee_number": re.sub(r"\D", "", employee_number),
                                "salt": salt,
                                "digest": digest,
                            }
                    pending = st.session_state.get("pending_signup")
                    if pending:
                        st.info("Pilih foto kartu, lalu ambil selfie dengan wajah dan seluruh nomor kartu terlihat jelas.")
                        result = IDENTITY_COMPONENT(
                            employee_number=pending["employee_number"],
                            key="identity-verification",
                            default=None,
                            height=430,
                        )
                        if isinstance(result, dict) and result.get("verified"):
                            if create_account(pending["email"], pending["salt"], pending["digest"]):
                                st.session_state.authenticated_email = pending["email"]
                                st.session_state.pop("pending_signup", None)
                                st.success("Pemeriksaan awal cocok. Akun demo dibuat.")
                                st.rerun()
                            else:
                                st.session_state.pop("pending_signup", None)
                                st.error("Email sudah terdaftar. Silakan masuk dengan akun tersebut.")
                        elif isinstance(result, dict) and result.get("message"):
                            st.error(result["message"])


def load_analysis():
    uploaded = st.file_uploader(
        "Unggah file RAB",
        type=["xlsx", "xls", "xlsm", "csv", "docx", "pdf"],
        help="Maksimum 20 MB. PDF hasil pindai perlu OCR sebelum diunggah.",
    )
    if uploaded is not None:
        try:
            content = uploaded.getvalue()
            upload_key = hashlib.sha256(uploaded.name.encode("utf-8") + content).hexdigest()
            if upload_key != st.session_state.get("upload_key"):
                sheets = parse_upload(uploaded.name, content)
                if not sheets or not any(any(clean_text(value) for row in rows for value in row) for rows in sheets.values()):
                    raise ValueError("Tidak ada isi tabel yang dapat dibaca dari file.")
                st.session_state.upload_name = uploaded.name
                st.session_state.upload_key = upload_key
                st.session_state.sheets = sheets
                preferred = next((name for name in sheets if re.search(r"rab|anggaran", name, re.I)), next(iter(sheets)))
                st.session_state.selected_sheet = preferred
                st.session_state.pop("tor_result", None)
                for key in ("tor_activity", "tor_location", "tor_year", "tor_notes"):
                    st.session_state.pop(key, None)
        except Exception as error:
            st.error("File tidak dapat diproses: {}".format(error))
            return None
    sheets = st.session_state.get("sheets")
    if not sheets:
        st.info("Unggah RAB untuk melihat ringkasan, grafik, dan membuat draft TOR.")
        return None
    names = list(sheets)
    chosen = st.selectbox("Pilih sheet / tabel", names, index=names.index(st.session_state.get("selected_sheet", names[0])) if st.session_state.get("selected_sheet") in names else 0)
    st.session_state.selected_sheet = chosen
    analysis = analyze_rows(sheets[chosen], chosen)
    return chosen, analysis


def render_dashboard(data):
    if not data:
        return
    sheet, analysis = data
    st.subheader("Ringkasan anggaran")
    st.caption("{} · {} baris anggaran disiapkan; {} baris disaring.".format(
        analysis["sheet"], len(analysis["rows"]), analysis["excluded_count"]
    ))
    categories = analysis.get("categories", [])
    pagu_for_card = analysis.get("category_total") if categories else analysis.get("pagu_total")
    remaining_for_card = analysis.get("category_remaining_total") if categories else analysis.get("remaining_total")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total pagu kategori", format_rupiah(pagu_for_card))
    col2.metric("Sisa anggaran", format_rupiah(remaining_for_card))
    col3.metric("Kategori dianalisis", len(categories))
    col4.metric("Rincian bernilai", len(analysis["rows"]))

    revision = analysis.get("revision")
    if revision:
        st.markdown("### Perubahan RAB & kecukupan dana")
        if revision["insufficient_count"]:
            st.warning("{} rincian kenaikan perlu ditinjau: sisa lebih kecil dari selisih.".format(revision["insufficient_count"]))
        elif revision["unknown_count"]:
            st.info("{} rincian kenaikan belum dapat dibandingkan karena nilai sisa tidak terbaca.".format(revision["unknown_count"]))
        elif revision["increases"]:
            st.success("Sisa anggaran pada semua rincian kenaikan yang terbaca mencukupi.")
        else:
            st.info("Tidak ada kenaikan pagu yang terdeteksi.")
        rev1, rev2, rev3, rev4 = st.columns(4)
        rev1.metric("Total kenaikan", format_rupiah(revision["increase_total"]))
        rev2.metric("Sisa mencukupi", revision["sufficient_count"])
        rev3.metric("Sisa lebih kecil", revision["insufficient_count"])
        rev4.metric("Kekurangan terhitung", format_rupiah(revision["shortfall"]))
        if revision["changes"]:
            revision_frame = pd.DataFrame(revision["changes"]).drop(columns=["_status"])
            st.dataframe(revision_frame, use_container_width=True, hide_index=True)

    if categories:
        chart_frame = pd.DataFrame(categories)
        st.markdown("### Pagu terbesar per kategori")
        pagu_data = chart_frame.nlargest(10, "Pagu").sort_values("Pagu")
        axis_divisor = 1_000_000_000 if pagu_data["Pagu"].max() >= 1_000_000_000 else 1_000_000 if pagu_data["Pagu"].max() >= 1_000_000 else 1
        axis_unit = "miliar" if axis_divisor == 1_000_000_000 else "juta" if axis_divisor == 1_000_000 else "rupiah"
        axis_suffix = " m" if axis_divisor == 1_000_000_000 else " jt" if axis_divisor == 1_000_000 else ""
        pagu_data = pagu_data.assign(
            **{
                "Pagu (skala)": pagu_data["Pagu"] / axis_divisor,
                "Label ringkas": pagu_data["Pagu"].map(format_compact_rupiah),
            }
        )
        fig = px.bar(
            pagu_data,
            x="Pagu (skala)",
            y="Kategori",
            orientation="h",
            text="Label ringkas",
            color_discrete_sequence=[COLORS[0]],
            custom_data=["Pagu"],
        )
        fig.update_traces(
            textposition="outside",
            cliponaxis=False,
            hovertemplate="%{y}<br>Pagu: Rp %{customdata[0]:,.0f}<extra></extra>",
        )
        fig.update_layout(
            height=max(330, 48 * len(pagu_data)),
            xaxis_title="Pagu (Rp {})".format(axis_unit),
            separators=",.",
            xaxis=dict(ticksuffix=axis_suffix),
            yaxis_title=None,
            showlegend=False,
            margin=dict(l=10, r=55, t=10, b=40),
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("### Porsi pagu")
        pie_data = chart_frame.nlargest(5, "Pagu").copy()
        pie_data["Label"] = pie_data.apply(
            lambda row: row["Kode"] or row["Kategori"][:24],
            axis=1,
        )
        if len(chart_frame) > 5:
            other = chart_frame.nsmallest(len(chart_frame) - 5, "Pagu")["Pagu"].sum()
            pie_data = pd.concat(
                [
                    pie_data,
                    pd.DataFrame([{"Kategori": "Kategori lainnya", "Label": "Lainnya", "Pagu": other}]),
                ],
                ignore_index=True,
            )
        pie = px.pie(
            pie_data,
            values="Pagu",
            names="Label",
            hole=0.58,
            color_discrete_sequence=COLORS,
            custom_data=["Kategori"],
        )
        pie.update_traces(
            textposition="inside",
            textinfo="percent",
            hovertemplate="%{customdata[0]}<br>Rp %{value:,.0f} (%{percent})<extra></extra>",
            sort=True,
        )
        pie.update_layout(
            height=410,
            margin=dict(l=10, r=240, t=10, b=10),
            legend_title=None,
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=0.68, font=dict(size=12)),
        )
        st.plotly_chart(pie, use_container_width=True)
        remaining = chart_frame.dropna(subset=["Sisa"]).nlargest(8, "Sisa")
        if not remaining.empty:
            st.markdown("### Sisa anggaran terbesar")
            remaining = remaining.sort_values("Sisa")
            remaining_divisor = 1_000_000_000 if remaining["Sisa"].max() >= 1_000_000_000 else 1_000_000 if remaining["Sisa"].max() >= 1_000_000 else 1
            remaining_unit = "miliar" if remaining_divisor == 1_000_000_000 else "juta" if remaining_divisor == 1_000_000 else "rupiah"
            remaining_suffix = " m" if remaining_divisor == 1_000_000_000 else " jt" if remaining_divisor == 1_000_000 else ""
            remaining = remaining.assign(
                **{
                    "Sisa (skala)": remaining["Sisa"] / remaining_divisor,
                    "Label ringkas": remaining["Sisa"].map(format_compact_rupiah),
                }
            )
            bar = px.bar(
                remaining,
                x="Sisa (skala)",
                y="Kategori",
                orientation="h",
                text="Label ringkas",
                color_discrete_sequence=[COLORS[1]],
                custom_data=["Sisa"],
            )
            bar.update_traces(
                textposition="outside",
                cliponaxis=False,
                hovertemplate="%{y}<br>Sisa: Rp %{customdata[0]:,.0f}<extra></extra>",
            )
            bar.update_layout(
                height=max(280, min(440, 48 * len(remaining))),
                xaxis_title="Sisa (Rp {})".format(remaining_unit),
                separators=",.",
                xaxis=dict(ticksuffix=remaining_suffix),
                yaxis_title=None,
                showlegend=False,
                margin=dict(l=10, r=45, t=10, b=40),
            )
            st.plotly_chart(bar, use_container_width=True)
    st.markdown("### Temuan dari data")
    for insight in analysis.get("insights", []):
        st.markdown("- " + insight)

    st.markdown("### Data anggaran yang sudah diproses")
    search = st.text_input("Cari uraian atau kode")
    prepared = pd.DataFrame(analysis["rows"])
    if not prepared.empty:
        if search:
            mask = prepared["Uraian"].str.contains(search, case=False, na=False) | prepared["Kode"].str.contains(search, case=False, na=False)
            prepared = prepared.loc[mask]
        st.dataframe(prepared.drop(columns=["_raw"], errors="ignore"), use_container_width=True, hide_index=True)
    else:
        st.warning("Struktur tabel belum dikenali. Pastikan terdapat header uraian dan pagu/jumlah.")


def infer_activity(analysis):
    coded = next((row for row in analysis.get("rows", []) if re.fullmatch(r"\d+(?:\.[A-Z0-9]+){2,}", row["Kode"], re.I) and row["Pagu"] is not None), None)
    if coded:
        return coded["Uraian"]
    if analysis.get("ranked"):
        return re.sub(r"^[\-•]\s*", "", analysis["ranked"][0]["Uraian"])
    return analysis.get("rows", [{}])[0].get("Uraian", "")


def build_tor_sections(activity, location, year, notes, analysis):
    rows = analysis.get("rows", [])
    source = next((row for row in rows if row["Uraian"].lower() == activity.lower() and row["Pagu"] is not None), None)
    if source:
        amount_text = "Pagu: {}{}".format(
            format_rupiah(source["Pagu"]),
            " · kode {}".format(source["Kode"]) if source["Kode"] else "",
        )
        remaining_text = "Sisa anggaran: {}".format(format_rupiah(source["Sisa anggaran"]))
    else:
        amount_text = "Baris kegiatan belum ditemukan. Cocokkan nama kegiatan dengan RAB sebelum menetapkan pagu."
        remaining_text = "Sisa anggaran belum dapat dipastikan."
    purpose = notes or "Lengkapi tujuan khusus dan indikator keberhasilan kegiatan."
    return [
        ("I. Latar Belakang", "TOR ini menjadi acuan pelaksanaan {} pada tahun {}, berdasarkan data pada sheet {}.".format(activity, year, analysis.get("sheet", "RAB"))),
        ("II. Maksud dan Tujuan", purpose),
        ("III. Sasaran", "Sasaran dan penerima manfaat: lengkapi sesuai target kegiatan."),
        ("IV. Ruang Lingkup", "Persiapan, koordinasi, pelaksanaan, pemantauan, dan pelaporan; sesuaikan dengan kebutuhan kegiatan."),
        ("V. Keluaran", "Target keluaran terukur: lengkapi sebelum TOR disahkan."),
        ("VI. Waktu dan Tempat", "Tahun: {} · Lokasi: {}. Lengkapi jadwal pelaksanaan secara rinci.".format(year, location)),
        ("VII. Pembiayaan", "{} · {}".format(amount_text, remaining_text)),
        ("VIII. Penutup", "Lengkapi bagian sasaran, keluaran, jadwal, dan persetujuan. Verifikasi pembiayaan dengan RAB yang disahkan."),
    ]


def make_tor_docx(activity, location, year, sheet_name, sections):
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    heading = document.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = heading.add_run("KERANGKA ACUAN KERJA")
    run.bold = True
    run.font.size = Pt(16)
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run(activity)
    title_run.bold = True
    title_run.font.size = Pt(15)
    for label, value in (
        ("Lokasi", location),
        ("Tahun anggaran", year),
        ("Sumber data", sheet_name),
    ):
        paragraph = document.add_paragraph()
        paragraph.add_run(label + ": ").bold = True
        paragraph.add_run(value)
    for title_text, body in sections:
        heading = document.add_heading(title_text, level=1)
        for run in heading.runs:
            run.font.color.rgb = RGBColor(18, 61, 112)
            run.font.name = "Arial"
        paragraph = document.add_paragraph(body)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in paragraph.runs:
            run.font.name = "Arial"
            run.font.size = Pt(11)
    end = document.add_paragraph()
    end.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    end.add_run("Draft untuk ditinjau dan dilengkapi\n\n\nPenanggung Jawab Kegiatan").italic = True
    output = io.BytesIO()
    document.save(output)
    return output.getvalue()


def render_tor(data):
    st.subheader("Generator TOR")
    if not data:
        st.info("Unggah RAB pada tab Dashboard untuk menghubungkan data dan membuat draft TOR.")
        return
    sheet, analysis = data
    default_activity = infer_activity(analysis)
    activity = st.text_input("Nama kegiatan", value=default_activity, key="tor_activity")
    location = st.text_input("Lokasi pelaksanaan", value="", key="tor_location")
    year = st.text_input("Tahun pelaksanaan", value="2026", key="tor_year")
    notes = st.text_area("Tujuan / catatan khusus", key="tor_notes")
    st.caption("Sumber data: {} · {} baris bernilai.".format(sheet, len(analysis["rows"])))
    if st.button("Buat draft TOR", type="primary"):
        if not activity.strip():
            st.error("Nama kegiatan perlu diisi.")
        else:
            sections = build_tor_sections(activity.strip(), location.strip() or "Lokasi pelaksanaan ditetapkan kemudian", year.strip() or "2026", notes.strip(), analysis)
            st.session_state.tor_result = {
                "activity": activity.strip(),
                "location": location.strip() or "Lokasi pelaksanaan ditetapkan kemudian",
                "year": year.strip() or "2026",
                "sections": sections,
            }
    result = st.session_state.get("tor_result")
    if result:
        st.markdown("### Draft TOR siap")
        summary_columns = st.columns(3)
        summary_columns[0].markdown("**Kegiatan**  \n{}".format(result["activity"]))
        summary_columns[1].markdown("**Lokasi**  \n{}".format(result["location"]))
        summary_columns[2].markdown("**Tahun**  \n{}".format(result["year"]))
        with st.expander("Lihat isi lengkap TOR"):
            st.markdown("#### KERANGKA ACUAN KERJA")
            st.markdown("**Sumber data:** {}".format(sheet))
            for section_title, body in result["sections"]:
                st.markdown("**{}**".format(section_title))
                st.write(body)
        document = make_tor_docx(result["activity"], result["location"], result["year"], sheet, result["sections"])
        safe_name = re.sub(r"[^a-z0-9]+", "-", result["activity"].lower()).strip("-") or "draft"
        st.download_button(
            "Unduh TOR dalam Word",
            data=document,
            file_name="TOR-{}.docx".format(safe_name),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


def main():
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.4rem; padding-bottom: 2rem; }
        h1, h2, h3 { color: #092759; }
        [data-testid="stMetric"] { background: #fff; border: 1px solid #e4ebf2;
            border-radius: 10px; padding: 14px 16px; }
        .stTabs [data-baseweb="tab-list"] { gap: 8px; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    if not st.session_state.get("authenticated_email"):
        render_auth()
        return

    with st.sidebar:
        st.markdown("## Ruang Anggaran")
        logo_columns = st.columns(3)
        for column, filename in zip(logo_columns, ("balmon.png", "djid.png", "komdigi.png")):
            path = ROOT / "image" / filename
            if path.exists():
                column.image(str(path), use_container_width=True)
        st.caption("Balai Monitor SFR Kelas I Surabaya")
        st.write(st.session_state.authenticated_email)
        if st.button("Keluar", use_container_width=True):
            logout()
            st.rerun()
        st.divider()
        st.caption("Draft dan insight perlu ditinjau sebelum digunakan resmi.")

    st.title("Dashboard RAB & Generator TOR")
    data = load_analysis()
    dashboard_tab, tor_tab = st.tabs(["Dashboard", "Generator TOR"])
    with dashboard_tab:
        render_dashboard(data)
    with tor_tab:
        render_tor(data)


if __name__ == "__main__":
    main()
