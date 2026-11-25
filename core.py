"""core.py - Logika inti SEARAH Sistem Cataloging & SPF (dapat diuji tanpa Streamlit/Gemini).

Alur:
1. read_any / detect_columns      -> baca file SAP (xlsx/csv) dan deteksi kolom otomatis
2. clean_text                     -> pembersihan teks untuk template (karakter rusak, pecahan, spasi)
3. match_text                     -> teks internal untuk pencocokan (kamus sinonim)
4. classify_rows                  -> label contoh -> aturan -> AI (opsional) -> REVIEW
5. validate_template1 / build_template1_bytes -> validasi & file Template 1 (Source_Data_Upload)
6. extract_template2_rows / validate_template2 / build_template2_bytes -> Template 2 (Extracted_Data)
7. build_spf_bytes                -> file Stock Proposal Form (SPF)
"""
from __future__ import annotations

import io
import random
import re
import time

import openpyxl
import pandas as pd


# ----------------------------------------------------------------------------
# 1. Membaca file
# ----------------------------------------------------------------------------
def read_any(name: str, raw: bytes) -> pd.DataFrame:
    """Baca xlsx atau csv. CSV dicoba UTF-8 dulu, lalu cp1252."""
    if name.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "cp1252"):
            try:
                return pd.read_csv(
                    io.BytesIO(raw), dtype=str, encoding=enc,
                    keep_default_na=False
                )
            except UnicodeDecodeError:
                continue
        return pd.read_csv(
            io.BytesIO(raw), dtype=str, encoding="latin1",
            keep_default_na=False
        )
    return pd.read_excel(io.BytesIO(raw), dtype=str).fillna("")


def detect_columns(cols) -> dict:
    """Tebak kolom kode / short / long / uom dari nama header ekspor SAP."""
    low = {c: str(c).lower() for c in cols}

    def pick(pred):
        for c, lc in low.items():
            if pred(lc):
                return c
        return None

    short = pick(lambda x: "description" in x or "short" in x)
    long_ = pick(lambda x: "po text" in x or "long" in x)
    uom = pick(lambda x: "unit" in x or "uom" in x)
    code = pick(
        lambda x: ("material" in x or "code" in x)
        and "description" not in x
    )

    guess = [code, short, long_, uom]
    fallback = list(cols)[:4]
    return {
        "code": code or (fallback[0] if fallback else None),
        "short": short or (fallback[1] if len(fallback) > 1 else (fallback[0] if fallback else None)),
        "long": long_ or (fallback[2] if len(fallback) > 2 else (fallback[0] if fallback else None)),
        "uom": uom or (fallback[3] if len(fallback) > 3 else (fallback[0] if fallback else None)),
        "_guessed": all(guess),
    }


def load_master(df_raw: pd.DataFrame):
    """Material Standard -> (df_nm: pasangan noun|modifier, df_ms: semua atribut)."""
    df = df_raw.fillna("").copy()
    df.columns = [str(c).strip().upper() for c in df.columns]

    for c in ("NOUN", "MODIFIER", "MANDATORY", "ATTRIBUTE"):
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip().str.upper()

    if "SEQ" in df.columns:
        df["SEQ"] = df["SEQ"].astype(str).str.strip().astype(int)

    required = {"NOUN", "MODIFIER"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            "Kolom wajib tidak ditemukan di Material Standard: "
            + ", ".join(sorted(missing))
        )

    df_nm = (
        df[["NOUN", "MODIFIER"]]
        .drop_duplicates()
        .sort_values(["NOUN", "MODIFIER"])
        .reset_index(drop=True)
    )
    return df_nm, df


# ----------------------------------------------------------------------------
# 2. Pembersihan teks (masuk ke template)
# ----------------------------------------------------------------------------
CHAR_FIXES = [
    ("Â°", "°"),
    ("GÇ¥", '"'),
    ("GÇô", "-"),
    ("GÇ£", '"'),
    ("Â", ""),
    ("”", '"'), ("“", '"'),
    ("’", "'"), ("‘", "'"),
    ("–", "-"), ("—", "-"),
    ("\xa0", " "),
]
ASCII_SYMBOLS = [
    ("®", ""), ("™", ""), ("•", "-"),
    ("±", "+/-"), ("³", "3")
]


def clean_text(s, ascii_symbols: bool = False) -> str:
    s = "" if s is None else str(s)

    for a, b in CHAR_FIXES:
        s = s.replace(a, b)

    for ch, frac in (("½", "1/2"), ("¼", "1/4"), ("¾", "3/4")):
        s = re.sub(rf"(\d){ch}", rf"\1-{frac}", s)
        s = s.replace(ch, frac)

    if ascii_symbols:
        for a, b in ASCII_SYMBOLS:
            s = s.replace(a, b)

    s = re.sub(r"[\t\r\n]+", " ", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s


# ----------------------------------------------------------------------------
# 3. Teks internal untuk pencocokan (kamus sinonim)
# ----------------------------------------------------------------------------
SYNONYMS = [
    (r"\bPUP ?JOINTS?\b", "PUP JOINT"),
    (r"\bTBG\b", "TUBING"),
    (r"\bCSG\b", "CASING"),
    (r"\bX-?OVER\b|\bCROSS ?OVER\b", "CROSSOVER"),
    (r"\bHAND ?GLOVES?\b", "HAND GLOVE"),
    (r"\b(?:V-?BELTS?|BELT-?V)\b", "BELT V"),
    (r"\bELEC(?:TRIC)?\b", "ELECTRICAL"),
    (r"\bPNEU\b", "PNEUMATIC"),
    (r"\bTRUST\b", "THRUST"),
]


def match_text(s) -> str:
    t = str(s or "").upper()
    for pat, rep in SYNONYMS:
        t = re.sub(pat, rep, t)
    return t


# ----------------------------------------------------------------------------
# 4. Konteks & Klasifikasi (Template 1)
# ----------------------------------------------------------------------------
def build_context(
    df_nm: pd.DataFrame,
    labeled: pd.DataFrame | None = None
) -> dict:
    """Siapkan struktur pencarian dari master dan file contoh berlabel (Template 2 terisi)."""
    ctx = {
        "valid": set(zip(df_nm["NOUN"], df_nm["MODIFIER"])),
        "nouns": set(df_nm["NOUN"]),
        "mods_by_noun": df_nm.groupby("NOUN")["MODIFIER"].apply(list).to_dict(),
        "pairs_text": "\n".join(
            f"{n} | {m}" for n, m in zip(df_nm["NOUN"], df_nm["MODIFIER"])
        ),
        "labels": {},
        "index": [],
        "attr_examples": {},
    }

    if labeled is not None and not labeled.empty:
        lab = labeled.fillna("")
        required = {
            "PSC Material Code", "Noun", "Modifier",
            "Material - Long Description"
        }
        missing = required - set(lab.columns)
        if missing:
            raise ValueError(
                "Kolom wajib Template 2 tidak ditemukan: "
                + ", ".join(sorted(missing))
            )

        first = lab[
            lab["Material - Long Description"].astype(str).str.strip() != ""
        ]

        for _, r in first.iterrows():
            code = str(r["PSC Material Code"]).strip()
            noun = str(r["Noun"]).strip().upper()
            mod = str(r["Modifier"]).strip().upper()
            text = str(r["Material - Long Description"])
            ctx["labels"][code] = (noun, mod)
            ctx["index"].append((_tokens(text), text, noun, mod))

        if {"Attribute", "Value"}.issubset(set(lab.columns)):
            grouped = {}
            cur_code = None
            cur_nm = None
            cur_desc = ""
            cur_attrs = {}
            for _, r in lab.iterrows():
                mc = str(r["PSC Material Code"]).strip()
                desc = str(r["Material - Long Description"]).strip()
                if desc:
                    if cur_code and cur_nm and cur_attrs:
                        grouped.setdefault(cur_nm, []).append((cur_desc, cur_attrs))
                    cur_code = mc
                    cur_nm = (str(r["Noun"]).strip().upper(), str(r["Modifier"]).strip().upper())
                    cur_desc = desc
                    cur_attrs = {}
                attr = str(r["Attribute"]).strip().upper()
                val = str(r["Value"]).strip().upper()
                if attr and val:
                    cur_attrs[attr] = val
            if cur_code and cur_nm and cur_attrs:
                grouped.setdefault(cur_nm, []).append((cur_desc, cur_attrs))
            ctx["attr_examples"] = grouped

    return ctx


def _tokens(text: str) -> frozenset:
    return frozenset(re.findall(r"[A-Z]{3,}", match_text(text)))


def similar_examples(
    text: str,
    ctx: dict,
    k: int = 3,
    exclude_text: str | None = None
) -> list:
    """k contoh berlabel paling mirip (Jaccard pada kata) untuk few-shot AI."""
    q = _tokens(text)
    if not q or not ctx["index"]:
        return []

    scored = []
    for toks, t, noun, mod in ctx["index"]:
        if exclude_text is not None and t == exclude_text:
            continue
        inter = len(q & toks)
        if inter:
            scored.append((inter / len(q | toks), t, noun, mod))

    scored.sort(key=lambda x: -x[0])
    return [(t, n, m) for _, t, n, m in scored[:k]]


def _first_token(t: str) -> str:
    return re.split(r"[,;]", t)[0].strip()


def rule_classify(short: str, long: str, ctx: dict):
    """Aturan berpresisi tinggi. Return tuple atau None."""
    s, l = match_text(short), match_text(long)
    comb = f"{s} , {l}"
    valid = ctx["valid"]
    mods_by_noun = ctx["mods_by_noun"]

    if (
        re.match(r"^CROSSOVER\b", _first_token(l))
        or re.match(r"^CROSSOVER\b", _first_token(s))
    ):
        if ("CROSS-OVER", "--") in valid:
            return (
                "CROSS-OVER", "--", 90, "ATURAN",
                "Kata pertama CROSSOVER/X-OVER"
            )

    if re.search(r"\bPUP\b", comb) and re.search(
        r"\bPUP JOINT\b|\bPUP\b.*\b(CASING|TUBING)\b|"
        r"\b(CASING|TUBING)\b.*\bPUP\b",
        comb
    ):
        mod = (
            "TUBING"
            if re.search(r"\bTUBING\b", comb)
            and not re.search(r"\bCASING\b", comb)
            else "CASING"
        )
        if ("PUP JOINT", mod) in valid:
            return (
                "PUP JOINT", mod, 78, "ATURAN",
                f"PUP JOINT + kata {mod}"
            )

    for tok in (_first_token(l), _first_token(s)):
        if tok in mods_by_noun:
            for mod in sorted(mods_by_noun[tok], key=len, reverse=True):
                if mod != "--" and re.search(
                    rf"\b{re.escape(mod)}\b", comb
                ):
                    return (
                        tok, mod, 80, "ATURAN",
                        f"Noun + modifier tertulis ({tok} | {mod})"
                    )

    return None


# ---- AI Helper & Error Handling ---------------------------------------------
class AIFatalError(Exception):
    """Error AI yang tidak akan sembuh dengan mengulang (kunci, model, kuota)."""


def is_fatal_error(error: Exception) -> bool:
    """Kunci salah, model tidak ada/sudah dimatikan, izin ditolak, kuota harian habis."""
    status = getattr(error, "status_code", None) or getattr(error, "code", None)
    try:
        status = int(status)
    except (TypeError, ValueError):
        status = None
    if status in (400, 401, 403, 404):
        return True
    msg = str(error).upper()
    fatal_terms = (
        "ACCESS_TOKEN_TYPE_UNSUPPORTED", "UNAUTHENTICATED", "API_KEY_INVALID",
        "API KEY NOT VALID", "PERMISSION_DENIED", "NO LONGER AVAILABLE",
        "NOT_FOUND", "PERDAY", "PER_DAY", "LIMIT: 0",
    )
    return any(term in msg for term in fatal_terms)


def _is_retryable(error: Exception) -> bool:
    status = (
        getattr(error, "status_code", None)
        or getattr(error, "code", None)
    )
    try:
        status = int(status)
    except (TypeError, ValueError):
        status = None

    if status in (408, 425, 429, 500, 502, 503, 504):
        return True

    msg = str(error).upper()
    retry_terms = (
        "503", "UNAVAILABLE",
        "429", "RESOURCE_EXHAUSTED",
        "500", "INTERNAL SERVER ERROR",
        "502", "BAD GATEWAY",
        "504", "GATEWAY TIMEOUT",
        "408", "TIMEOUT", "TIMED OUT",
        "CONNECTION RESET", "CONNECTION ERROR",
        "TEMPORARILY", "TEMPORARY",
    )
    return any(term in msg for term in retry_terms)


def make_gemini_batch_fn(
    api_key: str,
    model: str,
    ctx: dict,
    retries: int = 5,
    base_delay: float = 2,
    max_delay: float = 30,
):
    """Kembalikan fungsi f(payload) -> {row_index: (noun, modifier, conf, alasan)}."""
    from google import genai
    from google.genai import types
    from pydantic import BaseModel

    class Item(BaseModel):
        row_index: int
        noun: str
        modifier: str
        confidence: int
        reason: str

    class Result(BaseModel):
        items: list[Item]

    client = genai.Client(api_key=api_key.strip())

    def run(payload: list[dict]) -> dict:
        blocks = []

        for p in payload:
            examples = "\n".join(
                f"   - {t[:200]} => {n} | {m}"
                for t, n, m in p.get("examples", [])
            )

            block = (
                f"[row_index {p['row_index']}]\n"
                f"SHORT: {p['short'][:150]}\n"
                f"PO TEXT: {p['long'][:600]}"
            )
            if examples:
                block += (
                    "\nCONTOH SERUPA (sudah dikatalogkan):\n"
                    + examples
                )
            blocks.append(block)

        prompt = (
            "Anda cataloguer senior SKK Migas. Untuk SETIAP item, pilih "
            "tepat satu pasangan NOUN | MODIFIER dari DAFTAR RESMI.\n"
            "Catatan penting:\n"
            "- SHORT dibatasi 40 karakter dan sering disingkat/terpotong "
            "(PUPJOINT, X-OVER, TBG=tubing, CSG=casing). PO TEXT lebih "
            "lengkap; utamakan PO TEXT.\n"
            "- Tentukan benda utamanya, bukan peralatan tempat ia dipasang.\n"
            "- 'SPARE PARTS | --' sah dipakai untuk suku cadang yang tidak "
            "punya kelas sendiri.\n"
            "- Ikuti pola CONTOH SERUPA bila relevan. Jangan mengarang "
            "pasangan di luar daftar.\n"
            "- confidence 0-100 (jujur; turunkan bila ragu).\n\n"
            "ITEM:\n"
            + "\n\n".join(blocks)
            + "\n\nDAFTAR RESMI:\n"
            + ctx["pairs_text"]
        )

        last_error = None
        total_attempts = max(1, int(retries))

        for attempt in range(total_attempts):
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=Result,
                    ),
                )

                if not getattr(response, "text", None):
                    raise ValueError("Gemini mengembalikan respons kosong.")

                parsed = Result.model_validate_json(response.text)

                return {
                    item.row_index: (
                        item.noun.strip().upper(),
                        item.modifier.strip().upper(),
                        item.confidence,
                        item.reason,
                    )
                    for item in parsed.items
                }

            except Exception as error:
                last_error = error
                if is_fatal_error(error):
                    raise AIFatalError(str(error)[:500]) from error
                if not _is_retryable(error):
                    raise
                if attempt < total_attempts - 1:
                    delay = min(
                        float(base_delay) * (2 ** attempt),
                        float(max_delay),
                    )
                    delay += random.uniform(0, min(1.0, delay * 0.2))
                    time.sleep(delay)

        raise last_error

    return run


def classify_rows(
    rows: list[dict],
    ctx: dict,
    ai_batch=None,
    batch_size: int = 5,
    progress=None
) -> list[dict]:
    """Klasifikasi: label contoh -> aturan -> AI -> REVIEW."""
    results: dict[int, tuple] = {}
    pending: list[int] = []
    errors: dict[int, str] = {}
    fatal_message = None

    for i, row in enumerate(rows):
        label = ctx["labels"].get(row["code"])

        if label and label in ctx["valid"]:
            results[i] = (
                label[0], label[1], 100,
                "CONTOH", "Label dari file contoh berlabel"
            )
            continue

        rule = rule_classify(row["short"], row["long"], ctx)
        if rule:
            results[i] = rule
        else:
            pending.append(i)

    if ai_batch is not None and pending:
        safe_batch_size = max(1, int(batch_size))

        for start in range(0, len(pending), safe_batch_size):
            chunk = pending[start:start + safe_batch_size]

            payload = [
                {
                    "row_index": i,
                    "short": rows[i]["short"],
                    "long": rows[i]["long"],
                    "examples": similar_examples(
                        rows[i]["long"] or rows[i]["short"],
                        ctx
                    ),
                }
                for i in chunk
            ]

            try:
                decisions = ai_batch(payload)
                if not isinstance(decisions, dict):
                    raise TypeError("Hasil AI harus berupa dictionary berdasarkan row_index.")
            except AIFatalError as error:
                decisions = {}
                fatal_message = str(error)[:500]
                for i in chunk:
                    errors[i] = fatal_message
            except Exception as error:
                decisions = {}
                message = str(error)[:500]
                for i in chunk:
                    errors[i] = message

            for i in chunk:
                decision = decisions.get(i)

                if not decision:
                    note = (
                        f"Gagal AI: {errors[i]}"
                        if i in errors
                        else "AI tidak mengembalikan hasil untuk baris ini."
                    )
                    results[i] = ("", "", 0, "REVIEW", note)
                    continue

                try:
                    noun, modifier, confidence, reason = decision
                    noun = str(noun).strip().upper()
                    modifier = str(modifier).strip().upper()
                    confidence = max(0, min(100, int(confidence)))
                    reason = str(reason)

                    if (noun, modifier) in ctx["valid"]:
                        results[i] = (noun, modifier, confidence, "AI", reason)
                    else:
                        results[i] = (
                            "", "", 0, "REVIEW",
                            "AI mengusulkan pasangan tidak ada di master: "
                            f"{noun} | {modifier}"
                        )
                except (TypeError, ValueError) as error:
                    results[i] = (
                        "", "", 0, "REVIEW",
                        f"Format jawaban AI tidak valid: {error}"
                    )

            if fatal_message:
                for j in pending[start + len(chunk):]:
                    results[j] = (
                        "", "", 0, "REVIEW",
                        f"AI dihentikan: {fatal_message}"
                    )
                if progress:
                    progress(len(pending), len(pending))
                break

            if progress:
                progress(min(start + len(chunk), len(pending)), len(pending))

    output = []
    for i, _row in enumerate(rows):
        noun, modifier, confidence, source, note = results.get(
            i, ("", "", 0, "REVIEW", "Belum diklasifikasi")
        )
        output.append({
            "noun": noun,
            "modifier": modifier,
            "confidence": confidence,
            "source": source,
            "note": note,
        })

    return output


# ----------------------------------------------------------------------------
# 5. Template 1: validasi dan pembuatan file
# ----------------------------------------------------------------------------
LIMITS = {
    "code": 20,
    "noun": 30,
    "modifier": 30,
    "short": 100,
    "long": 3000,
    "uom": 10,
}


def make_short(long: str, limit: int = 100) -> str:
    """Potong PO Text rapi (di koma/spasi) agar muat batas Short Description."""
    if len(long) <= limit:
        return long

    cut = long[:limit]
    idx = max(cut.rfind(","), cut.rfind(" "))
    return (cut[:idx] if idx > 40 else cut).rstrip(" ,;")


def parse_uom_map(text: str) -> dict:
    mapping = {}
    for line in (text or "").splitlines():
        if "=" in line:
            source, target = line.split("=", 1)
            if source.strip():
                mapping[source.strip().upper()] = target.strip().upper()
    return mapping


def validate_template1(
    df: pd.DataFrame,
    valid_pairs: set
) -> pd.DataFrame:
    """df kolom: code, noun, modifier, short, long, uom. Return tabel masalah."""
    issues = []

    required = ["code", "noun", "modifier", "short", "long", "uom"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(
            "Kolom validasi Template 1 tidak ditemukan: "
            + ", ".join(missing)
        )

    def add(index, field, message):
        issues.append({
            "baris": index + 1,
            "kode": df.iloc[index]["code"],
            "kolom": field,
            "masalah": message,
        })

    duplicates = df["code"].duplicated(keep=False)

    for i in range(len(df)):
        row = df.iloc[i]

        if not str(row["code"]).strip():
            add(i, "code", "Kode material kosong")

        if duplicates.iloc[i]:
            add(i, "code", "Kode material duplikat")

        noun = str(row["noun"]).strip()
        modifier = str(row["modifier"]).strip()

        if not noun or not modifier:
            add(i, "noun/modifier", "Belum diklasifikasi")
        elif (noun, modifier) not in valid_pairs:
            add(
                i, "noun/modifier",
                f"Pasangan tidak ada di master: {noun} | {modifier}"
            )

        if not str(row["long"]).strip():
            add(i, "long", "Long Description kosong")

        if not str(row["uom"]).strip():
            add(i, "uom", "Uom kosong")

        for field, limit in LIMITS.items():
            length = len(str(row[field]))
            if length > limit:
                add(
                    i, field,
                    f"Melebihi {limit} karakter ({length})"
                )

    return pd.DataFrame(
        issues,
        columns=["baris", "kode", "kolom", "masalah"]
    )


FORMAT_ROWS = [
    [
        "INTEGER", "ALPHANUMERIC", "ALPHANUMERIC",
        "ALPHANUMERIC", "ALPHANUMERIC", "ALPHANUMERIC",
        "ALPHANUMERIC", "HEADER DATA", "DATA TYPE"
    ],
    [
        "No. of the material data ",
        "The Code/KEYMAP# of the material",
        "NOUN as per the CLASS master",
        "MODIFIER as per the CLASS master",
        "Short description of the material",
        "Long description of the material",
        "Uom of the material",
        None,
        "DESCRIPTION"
    ],
    ["2", "20", "30", "30", "100", "3000", "10", None, "SIZE"],
    [
        "No.", "PSC Material code", "Noun", "Modifier",
        "Material - Short Description",
        "Material - Long Description", "Uom", None, "HEADER"
    ],
]


def build_template1_bytes(df: pd.DataFrame) -> bytes:
    """Buat file Template 1."""
    required = ["code", "noun", "modifier", "short", "long", "uom"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(
            "Kolom untuk membuat Template 1 tidak ditemukan: "
            + ", ".join(missing)
        )

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Source_Data_Upload"

    worksheet.append([
        "No.", "PSC Material Code", "Noun", "Modifier",
        "Material - Short Description",
        "Material - Long Description", "Uom"
    ])

    for number, row in enumerate(
        df[required].itertuples(index=False, name="MaterialRow"),
        start=1
    ):
        worksheet.append([
            number,
            row.code,
            row.noun,
            row.modifier,
            row.short,
            row.long,
            row.uom,
        ])

    format_sheet = workbook.create_sheet("Source_Data_Upload_Format")
    for row in FORMAT_ROWS:
        format_sheet.append(row)

    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


# ----------------------------------------------------------------------------
# 6. Template 2 (Extracted_Data): Ekstraksi Atribut & Pembuatan File
# ----------------------------------------------------------------------------
TEMPLATE2_COLUMNS = [
    "No.", "PSC Code", "PSC Material Code", "Ref.No",
    "Noun", "Modifier", "Sequence Number", "Attribute",
    "Value", "Remarks", "Mandatory", "Uom", "Material - Long Description"
]


def make_gemini_extract_fn(
    api_key: str,
    model: str,
    retries: int = 5,
    base_delay: float = 2,
    max_delay: float = 30,
):
    """Kembalikan fungsi f(noun, modifier, long_desc, schema_rows, examples) -> {seq: value}."""
    from google import genai
    from google.genai import types
    from pydantic import BaseModel

    class AttrVal(BaseModel):
        seq: int
        value: str

    class ExtractOut(BaseModel):
        values: list[AttrVal]

    client = genai.Client(api_key=api_key.strip())

    def run(
        noun: str,
        modifier: str,
        long_desc: str,
        schema_rows: list[dict],
        examples: list[tuple[str, dict]] | None = None,
    ) -> dict[int, str]:
        schema_lines = "\n".join(
            f"- SEQ {r['seq']} | Mandatory={r['mandatory']} | Attribute={r['attribute']}"
            for r in schema_rows
        )

        ex_block = ""
        if examples:
            ex_lines = []
            for ex_desc, ex_map in examples[:2]:
                pairs = ", ".join(f"{k}: {v}" for k, v in ex_map.items())
                ex_lines.append(f"  Deskripsi: {ex_desc[:250]}\n  Hasil: {pairs}")
            ex_block = "\nCONTOH FORMAT PENGISIAN E-CATALOGING SEBELUMNYA:\n" + "\n".join(ex_lines) + "\n"

        prompt = (
            "Anda adalah cataloguer senior SKK Migas. Ekstrak nilai spesifikasi dari DESKRIPSI MATERIAL "
            "ke dalam daftar ATRIBUT BAKU di bawah ini.\n"
            "Aturan wajib:\n"
            "- Semua nilai (value) wajib HURUF KAPITAL (UPPERCASE).\n"
            "- Jika informasi tidak tercantum di deskripsi, isi string kosong (''). Jangan mengarang nilai.\n"
            "- Gunakan satuan dan penulisan standar SKK Migas seperti pada contoh (misal: '5 IN', '13 PPF', "
            "'GRADE L80', 'API SPEC 5CT', 'STAINLESS STEEL-316', 'CLASS 150', '150 PSI').\n"
            "- Jika ada spesifikasi teknis sisa yang penting namun tidak punya slot atribut khusus, masukkan ke "
            "'ADDITIONAL DETAILS' (bila atribut tersebut tersedia di skema).\n\n"
            f"NOUN: {noun}\n"
            f"MODIFIER: {modifier}\n"
            f"DESKRIPSI MATERIAL: \"{long_desc}\"\n"
            f"{ex_block}\n"
            f"SKEMA ATRIBUT:\n{schema_lines}"
        )

        last_error = None
        total_attempts = max(1, int(retries))

        for attempt in range(total_attempts):
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=ExtractOut,
                    ),
                )
                if not getattr(response, "text", None):
                    raise ValueError("Gemini mengembalikan respons kosong.")
                parsed = ExtractOut.model_validate_json(response.text)
                return {
                    int(item.seq): str(item.value or "").strip().upper()
                    for item in parsed.values
                }
            except Exception as error:
                last_error = error
                if is_fatal_error(error):
                    raise AIFatalError(str(error)[:500]) from error
                if not _is_retryable(error):
                    raise
                if attempt < total_attempts - 1:
                    delay = min(float(base_delay) * (2 ** attempt), float(max_delay))
                    delay += random.uniform(0, min(1.0, delay * 0.2))
                    time.sleep(delay)

        raise last_error

    return run


def extract_template2_rows(
    items: list[dict],
    df_ms: pd.DataFrame,
    ctx: dict | None = None,
    ai_extract=None,
    psc_code: str = "",
    ref_prefix: str = "",
    progress=None,
) -> tuple[pd.DataFrame, str | None]:
    """Bangun tabel panjang Template 2 (1 baris per atribut per material)."""
    if not {"SEQ", "MANDATORY", "ATTRIBUTE"}.issubset(set(df_ms.columns)):
        raise ValueError("File Material Standard tidak memiliki kolom SEQ, MANDATORY, dan ATTRIBUTE.")

    schema_cache: dict[tuple[str, str], list[dict]] = {}
    out_rows = []
    seq_no = 1
    fatal_error = None
    total = len(items)

    for idx, item in enumerate(items):
        code = str(item.get("code", "")).strip()
        noun = str(item.get("noun", "")).strip().upper()
        mod = str(item.get("modifier", "")).strip().upper()
        long_desc = str(item.get("long", "")).strip()
        uom = str(item.get("uom", "")).strip().upper()

        item_ref = str(item.get("ref_no", "")).strip()
        if not item_ref and ref_prefix.strip():
            item_ref = f"{ref_prefix.strip()}_{idx + 1}"

        key = (noun, mod)
        if key not in schema_cache:
            sub = df_ms[(df_ms["NOUN"] == noun) & (df_ms["MODIFIER"] == mod)].sort_values("SEQ")
            schema_cache[key] = [
                {
                    "seq": int(r["SEQ"]),
                    "mandatory": str(r["MANDATORY"]).strip().upper(),
                    "attribute": str(r["ATTRIBUTE"]).strip().upper(),
                }
                for _, r in sub.iterrows()
            ]

        schema_rows = schema_cache[key]
        val_map: dict[int, str] = {}
        remark_seq1 = ""

        if not schema_rows:
            out_rows.append({
                "No.": str(seq_no),
                "PSC Code": psc_code.strip(),
                "PSC Material Code": code,
                "Ref.No": item_ref,
                "Noun": noun,
                "Modifier": mod,
                "Sequence Number": "1",
                "Attribute": "",
                "Value": "",
                "Remarks": f"Skema atribut tidak ditemukan di Material Standard untuk {noun} | {mod}",
                "Mandatory": "",
                "Uom": uom,
                "Material - Long Description": long_desc,
            })
            seq_no += 1
            if progress:
                progress(idx + 1, total)
            continue

        if ai_extract is not None and fatal_error is None:
            ex_list = (ctx or {}).get("attr_examples", {}).get(key, [])
            try:
                val_map = ai_extract(noun, mod, long_desc, schema_rows, ex_list)
            except AIFatalError as err:
                fatal_error = str(err)[:500]
                remark_seq1 = f"AI dihentikan: {fatal_error}"
            except Exception as err:
                remark_seq1 = f"Gagal ekstrak AI: {str(err)[:300]}"
        elif fatal_error is not None:
            remark_seq1 = f"AI dihentikan: {fatal_error}"

        for s_row in schema_rows:
            s_num = s_row["seq"]
            is_first = (s_num == 1)
            out_rows.append({
                "No.": str(seq_no),
                "PSC Code": psc_code.strip(),
                "PSC Material Code": code,
                "Ref.No": item_ref,
                "Noun": noun,
                "Modifier": mod,
                "Sequence Number": str(s_num),
                "Attribute": s_row["attribute"],
                "Value": val_map.get(s_num, ""),
                "Remarks": remark_seq1 if is_first else "",
                "Mandatory": s_row["mandatory"],
                "Uom": uom if is_first else "",
                "Material - Long Description": long_desc if is_first else "",
            })
            seq_no += 1

        if progress:
            progress(idx + 1, total)

    return pd.DataFrame(out_rows, columns=TEMPLATE2_COLUMNS), fatal_error


def validate_template2(df: pd.DataFrame) -> pd.DataFrame:
    """Periksa atribut Mandatory (M) yang masih kosong pada Template 2."""
    issues = []
    if df.empty:
        return pd.DataFrame(columns=["baris", "kode", "noun_modifier", "atribut", "masalah"])

    for i, r in df.iterrows():
        mand = str(r.get("Mandatory", "")).strip().upper()
        val = str(r.get("Value", "")).strip()
        if mand == "M" and not val:
            issues.append({
                "baris": i + 1,
                "kode": r.get("PSC Material Code", ""),
                "noun_modifier": f"{r.get('Noun', '')} | {r.get('Modifier', '')}",
                "atribut": f"SEQ {r.get('Sequence Number', '')} - {r.get('Attribute', '')}",
                "masalah": "Atribut Mandatory (M) belum terisi",
            })
    return pd.DataFrame(issues, columns=["baris", "kode", "noun_modifier", "atribut", "masalah"])


def build_template2_bytes(df: pd.DataFrame, sheet_name: str = "Extracted_Data") -> bytes:
    """Buat file Excel Template 2 (Extracted_Data) sesuai format baku SKK Migas."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]

    header_font = openpyxl.styles.Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = openpyxl.styles.PatternFill(fill_type="solid", fgColor="1F4E78")

    ws.append(TEMPLATE2_COLUMNS)
    for c_idx in range(1, len(TEMPLATE2_COLUMNS) + 1):
        cell = ws.cell(1, c_idx)
        cell.font = header_font
        cell.fill = header_fill

    for row_num, (_, r) in enumerate(df.iterrows(), start=1):
        ws.append([
            str(row_num),
            str(r.get("PSC Code", "")).strip(),
            str(r.get("PSC Material Code", "")).strip(),
            str(r.get("Ref.No", "")).strip(),
            str(r.get("Noun", "")).strip().upper(),
            str(r.get("Modifier", "")).strip().upper(),
            str(r.get("Sequence Number", "")).strip(),
            str(r.get("Attribute", "")).strip().upper(),
            str(r.get("Value", "")).strip().upper(),
            str(r.get("Remarks", "")).strip(),
            str(r.get("Mandatory", "")).strip().upper(),
            str(r.get("Uom", "")).strip().upper(),
            str(r.get("Material - Long Description", "")).strip(),
        ])

    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# 7. Stock Proposal Form (SPF) - Format Resmi SEARAH (SPF- NEW FORMAT)
# ----------------------------------------------------------------------------
SPF_ROW_MERGES = [
    (4, 6),    # D:F Short Description
    (7, 16),   # G:P Long Description
    (17, 18),  # Q:R Criticality (V/E/D)
    (19, 24),  # S:X Part/Item No. or Model/Type No.
    (25, 26),  # Y:Z UOM
    (27, 28),  # AA:AB Qty Installed Per Eqpt.
    (29, 30),  # AC:AD Qty Installed Total
    (31, 32),  # AE:AF Est. Annual Usage
    (33, 35),  # AG:AI Reorder Level
    (36, 38),  # AJ:AL Max Level
    (39, 41),  # AM:AO Initial Order Qty.
    (42, 44),  # AP:AR Est. Unit Price (IDR)
    (45, 47),  # AS:AU Material Type
    (48, 49),  # AV:AW Val. Class
    (50, 51),  # AX:AY Certificate
    (52, 53),  # AZ:BA Lead Time
    (54, 57),  # BB:BE MESC No.
    (58, 61),  # BF:BI Material No.
    (62, 64),  # BJ:BL Date Catalogue Created
]


def make_short_40(short: str, long: str = "", limit: int = 40) -> str:
    """Pastikan Short Description SPF maksimal 40 karakter SAP."""
    s = clean_text(short).upper()
    if not s and long:
        s = clean_text(long).upper()
    if len(s) <= limit:
        return s
    cut = s[:limit]
    idx_sep = max(cut.rfind(","), cut.rfind(" "))
    return (cut[:idx_sep] if idx_sep > 20 else cut).rstrip(" ,;-")


def _to_num_or_str(val):
    s = str(val or "").strip()
    if not s:
        return None
    try:
        if "." in s:
            return float(s)
        return int(s)
    except ValueError:
        return s


def _copy_spf_row_style(ws, src_row: int, dst_row: int, max_col: int = 64):
    from copy import copy
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    for c in range(1, max_col + 1):
        src_cell = ws.cell(src_row, c)
        dst_cell = ws.cell(dst_row, c)
        if src_cell.has_style:
            dst_cell.font = copy(src_cell.font)
            dst_cell.fill = copy(src_cell.fill)
            dst_cell.border = copy(src_cell.border)
            dst_cell.alignment = copy(src_cell.alignment)
            dst_cell.number_format = src_cell.number_format
    for c1, c2 in SPF_ROW_MERGES:
        ws.merge_cells(start_row=dst_row, start_column=c1, end_row=dst_row, end_column=c2)


def build_spf_bytes(
    df_spf: pd.DataFrame,
    meta: dict,
    template_bytes: bytes | None = None,
) -> bytes:
    """Bangun file Excel Stock Proposal Form (SPF) mengikuti format resmi SEARAH (SPF- NEW FORMAT)."""
    if template_bytes:
        wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
        ws = wb["SPF- NEW FORMAT"] if "SPF- NEW FORMAT" in wb.sheetnames else wb.active

        # Bersihkan nilai contoh pada header template
        for coord in [
            "BF1", "AC5", "BC5", "H8", "X8", "AS8", "BH8",
            "H9", "X9", "AS9", "BH9", "H10", "X10", "AS10", "BH10",
            "H11", "X11", "AS11", "BH11", "K12", "AS12", "BH12",
        ]:
            ws[coord].value = None

        # Isi metadata header form SPF
        header_map = {
            "BF1": meta.get("spf_ref", ""),
            "H5": meta.get("originating_dept", ""),
            "AC5": meta.get("originator_ref", ""),
            "BC5": meta.get("required_date", ""),
            "H8": meta.get("equipment", ""),
            "X8": meta.get("size_capacity", ""),
            "AS8": meta.get("equipment_tag", ""),
            "BH8": meta.get("bom_no", ""),
            "H9": meta.get("model_type", ""),
            "X9": meta.get("manufacturer", ""),
            "AS9": meta.get("serial_no", ""),
            "BH9": meta.get("asset_no", ""),
            "H10": meta.get("sub_model", ""),
            "X10": meta.get("country_origin", ""),
            "AS10": meta.get("epc_contractor", ""),
            "BH10": meta.get("spir_no", ""),
            "H11": meta.get("applicable_std", ""),
            "X11": meta.get("mfr_drg", ""),
            "AS11": meta.get("project_drg", ""),
            "K12": meta.get("reason", ""),
            "AS12": meta.get("units_installed", ""),
            "BH12": meta.get("units_in_use", ""),
        }
        for coord, val in header_map.items():
            if val is not None and str(val).strip() != "":
                ws[coord].value = str(val).strip()

        eq_crit = str(meta.get("eqpmt_criticality", "")).strip()
        if eq_crit:
            ws["BH11"].value = _to_num_or_str(eq_crit)

        # Bersihkan baris item existing (18..61)
        clear_cols = [1, 2, 3, 4, 7, 17, 19, 25, 27, 29, 31, 33, 36, 39, 42, 45, 48, 50, 52, 54, 58, 62]
        for r in range(18, 62):
            for c in clear_cols:
                ws.cell(r, c).value = None

        n_items = len(df_spf)
        base_capacity = 44  # Baris 18..61

        # Jika item > 44 baris, sisipkan baris ekstra sebelum baris 62 & geser blok tanda tangan
        if n_items > base_capacity:
            extra = n_items - base_capacity
            footer_merges = [
                rng for rng in list(ws.merged_cells.ranges)
                if rng.min_row >= 62
            ]
            for rng in footer_merges:
                ws.unmerge_cells(str(rng))

            ws.insert_rows(62, amount=extra)

            for rng in footer_merges:
                ws.merge_cells(
                    start_row=rng.min_row + extra,
                    start_column=rng.min_col,
                    end_row=rng.max_row + extra,
                    end_column=rng.max_col,
                )

            for r in range(62, 62 + extra):
                _copy_spf_row_style(ws, src_row=18, dst_row=r)

        # Isi data per baris mulai row 18
        for idx, (_, row) in enumerate(df_spf.iterrows()):
            r = 18 + idx
            short_val = make_short_40(
                row.get("Short Description", ""),
                row.get("Long Description", ""),
                limit=40,
            )
            long_val = str(row.get("Long Description", "")).strip()

            ws.cell(r, 1).value = idx + 1
            ws.cell(r, 2).value = str(row.get("Plant", meta.get("default_plant", "ID05"))).strip().upper()
            ws.cell(r, 3).value = f"=LEN(D{r})"
            ws.cell(r, 4).value = short_val
            ws.cell(r, 7).value = long_val
            ws.cell(r, 17).value = str(row.get("Criticality", "D")).strip().upper()
            ws.cell(r, 19).value = str(row.get("Part / Category", "")).strip()
            ws.cell(r, 25).value = str(row.get("UoM", "EA")).strip().upper()
            ws.cell(r, 27).value = _to_num_or_str(row.get("Qty Per Eqpt", ""))
            ws.cell(r, 29).value = _to_num_or_str(row.get("Qty Total", ""))
            ws.cell(r, 31).value = _to_num_or_str(row.get("Est. Annual Usage", "0"))
            ws.cell(r, 33).value = _to_num_or_str(row.get("Reorder Level", ""))
            ws.cell(r, 36).value = _to_num_or_str(row.get("Max Level", ""))
            ws.cell(r, 39).value = _to_num_or_str(row.get("Initial Order Qty", ""))
            ws.cell(r, 42).value = _to_num_or_str(row.get("Est. Unit Price (IDR)", ""))
            ws.cell(r, 45).value = str(row.get("Material Type", "")).strip().upper()
            ws.cell(r, 48).value = str(row.get("Val. Class", "")).strip().upper()
            ws.cell(r, 50).value = str(row.get("Certificate", "")).strip()
            ws.cell(r, 52).value = str(row.get("Lead Time", "")).strip().upper()

            if str(row.get("MESC No.", "")).strip():
                ws.cell(r, 54).value = str(row.get("MESC No.", "")).strip()
            if str(row.get("Material Code", "")).strip() and meta.get("fill_mat_code", False):
                ws.cell(r, 58).value = str(row.get("Material Code", "")).strip()

            est_lines = max(2, (len(long_val) // 75) + 1)
            ws.row_dimensions[r].height = max(28.0, min(160.0, est_lines * 16.5))

        # Isi blok tanda tangan di footer (6 kolom approval)
        offset = max(0, n_items - base_capacity)
        header_row = 62 + offset
        name_row = 64 + offset
        date_row = 65 + offset

        # Cari nomor kolom otomatis berdasarkan label header baris 62
        col_by_label = {}
        for c in range(1, 65):
          lbl = str(ws.cell(header_row, c).value or "").upper()
          if "INITIATED BY" in lbl:
            col_by_label["init"] = c
          elif "ENDORSED BY" in lbl and "SENIOR" not in lbl:
            col_by_label["mgr"] = c
          elif "SENIOR MANAGER" in lbl:
            col_by_label["srmgr"] = c
          elif "CHECKED BY" in lbl or "INVENTORY" in lbl:
            col_by_label["inv"] = c
          elif "REVIEWED BY" in lbl:
            col_by_label["rev_scm"] = c
          elif "APPROVED BY" in lbl:
            col_by_label["app_scm"] = c

        c_init = col_by_label.get("init", 1)
        c_mgr = col_by_label.get("mgr", 7)
        c_srmgr = col_by_label.get("srmgr", 14)
        c_inv = col_by_label.get("inv", 26)
        c_rev = col_by_label.get("rev_scm", 45)
        c_app = col_by_label.get("app_scm", 55)

        # Tulis nama (kosongkan 'NAME/CODE' jika user tidak mengisi)
        ws.cell(name_row, c_init).value = (
            str(meta.get("initiated_by", "")).strip().upper() or None
        )
        ws.cell(date_row, c_init).value = (
            f"DATE:  {str(meta['initiated_date']).strip().upper()}"
            if meta.get("initiated_date")
            else "DATE:"
        )
        ws.cell(name_row, c_mgr).value = (
            str(meta.get("endorsed_mgr", "")).strip().upper() or None
        )
        ws.cell(name_row, c_srmgr).value = (
            str(meta.get("endorsed_srmgr", "")).strip().upper() or None
        )
        ws.cell(name_row, c_inv).value = (
            str(meta.get("checked_inv", "")).strip().upper() or None
        )
        ws.cell(name_row, c_rev).value = (
            str(meta.get("reviewed_scm", "")).strip().upper() or None
        )
        ws.cell(name_row, c_app).value = (
            str(meta.get("approved_scm", "")).strip().upper() or None
        )

    # Fallback jika template resmi tidak tersedia
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stock_Proposal_Form"

    header_font = openpyxl.styles.Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = openpyxl.styles.PatternFill(fill_type="solid", fgColor="1F4E78")
    title_font = openpyxl.styles.Font(name="Calibri", size=14, bold=True)
    meta_font = openpyxl.styles.Font(name="Calibri", size=10, bold=True)

    ws["A1"] = "STOCK PROPOSAL FORM (SPF) - SEARAH"
    ws["A1"].font = title_font
    ws["A2"] = "Originating Dept:"
    ws["A2"].font = meta_font
    ws["B2"] = meta.get("originating_dept", "")
    ws["E2"] = "Applicable Standard:"
    ws["E2"].font = meta_font
    ws["F2"] = meta.get("applicable_std", "")
    ws["A3"] = "Equipment / Model:"
    ws["A3"].font = meta_font
    ws["B3"] = f"{meta.get('equipment', '')} - {meta.get('model_type', '')}".strip(" -")
    ws["E3"] = "Required Date:"
    ws["E3"].font = meta_font
    ws["F3"] = meta.get("required_date", "")

    start_row = 5
    cols = list(df_spf.columns)
    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(start_row, c_idx, col_name)
        cell.font = header_font
        cell.fill = header_fill

    for r_offset, (_, row) in enumerate(df_spf.iterrows(), start=1):
        for c_idx, col_name in enumerate(cols, start=1):
            val = row[col_name]
            ws.cell(start_row + r_offset, c_idx, "" if pd.isna(val) else str(val))

    ws.freeze_panes = f"A{start_row + 1}"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

from google import genai
from google.genai import types
from pydantic import BaseModel, Field


class SPFItemRecommendation(BaseModel):
  recommended_short_text: str = Field(
      description=(
          "Short text standar e-Cat/SAP, HURUF KAPITAL, maks 40 karakter,"
          " format NOUN,MODIFIER,SPEC."
      )
  )
  recommended_long_text: str = Field(
      description=(
          "Long text terstruktur baris demi baris (NOUN, MODIFIER, ATTRIBUTES,"
          " PART NO, MANUFACTURER), HURUF KAPITAL."
      )
  )
  keterangan_ai: str = Field(
      description=(
          "Penjelasan singkat bagian yang diperbaiki/disingkat dan catatan"
          " jika ada spesifikasi teknis yang kurang."
      )
  )


def recommend_spf_item(
    api_key: str,
    model: str,
    short_in: str,
    long_in: str,
    part_no: str = "",
    eq_info: str = "",
) -> SPFItemRecommendation:
  client = genai.Client(api_key=api_key)
  prompt = f"""
    Anda adalah Senior Material Cataloger Migas (standar SAP Material Master & SKK Migas e-Cataloging).
    Tinjau dan rekomendasikan perbaikan Short Description dan Long Description untuk item Stock Proposal Form (SPF) berikut:

    - Konteks Equipment/Manufacturer: {eq_info}
    - Part/Category/Model No: {part_no}
    - Input Short Description: "{short_in}"
    - Input Long Description: "{long_in}"

    Aturan Wajib:
    1. Short Description MAKSIMAL 40 karakter, HURUF KAPITAL, format NOUN,MODIFIER,SPEC, tanpa spasi setelah koma, gunakan singkatan standar (IN, SS316, CL150, RF, NPT, ASSY, dll).
    2. Long Description HURUF KAPITAL, rapi per baris atribut (NOUN, MODIFIER, SIZE/RATING, MATERIAL, PART NO, MANUFACTURER).
    3. Berikan keterangan jelas apa yang diperbaiki dan peringatan jika ada spek penting yang belum tercantum.
    """
  response = client.models.generate_content(
      model=model,
      contents=prompt,
      config=types.GenerateContentConfig(
          response_mime_type="application/json",
          response_schema=SPFItemRecommendation,
          temperature=0.2,
      ),
  )
  return response.parsed
