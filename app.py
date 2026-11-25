"""app.py - SEARAH Sistem Cataloging & Stock Proposal Form (SPF)

Menu Utama:
  0. Menu Awal (Beranda & Status Sistem)
  1. Output Template 1 (Source_Data_Upload)
  2. Output Template 2 (Extracted_Data / Atribut)
  3. Generate SPF (Stock Proposal Form)
"""
from __future__ import annotations

import io
import os

import pandas as pd
import streamlit as st

import core


st.set_page_config(
    page_title="SEARAH - Sistem Cataloging & SPF",
    layout="wide",
)

BASE = os.path.dirname(os.path.abspath(__file__))
EDITABLE_T1 = ["Kelas", "Short (final)", "Uom"]
EDITABLE_T2 = ["PSC Code", "Ref.No", "Value", "Remarks", "Uom", "Material - Long Description"]


# ------------------------------------------------------------------
# Helper UI
# ------------------------------------------------------------------
def find_local(pred):
    for filename in sorted(os.listdir(BASE)):
        if pred(filename.lower()):
            return os.path.join(BASE, filename)
    return None


def secret(name, default=""):
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = ""
    return value or os.getenv(name, default)


@st.cache_data(show_spinner=False)
def cached_master(raw: bytes):
    return core.load_master(pd.read_excel(io.BytesIO(raw), dtype=str))


@st.cache_data(show_spinner=False)
def cached_labeled(raw: bytes):
    return pd.read_excel(io.BytesIO(raw), dtype=str).fillna("")


def apply_edits(base: pd.DataFrame, edits: dict) -> pd.DataFrame:
    current = base.copy()
    for index, changes in edits.items():
        for column, value in changes.items():
            current.at[index, column] = value
        if "Kelas" in changes:
            current.at[index, "Sumber"] = "MANUAL"
            current.at[index, "Keyakinan"] = 100
    return current


def to_template1(current: pd.DataFrame) -> pd.DataFrame:
    parts = current["Kelas"].fillna("").astype(str).str.split(
        " | ", n=1, regex=False, expand=True
    )
    if parts.shape[1] < 2:
        parts[1] = ""
    return pd.DataFrame({
        "code": current["Kode"],
        "noun": parts[0].fillna(""),
        "modifier": parts[1].fillna(""),
        "short": current["Short (final)"],
        "long": current["Long"],
        "uom": current["Uom"],
    })


# ------------------------------------------------------------------
# State Navigasi
# ------------------------------------------------------------------
MENU_LIST = [
    "🏠 Menu Awal (Beranda)",
    "1️⃣ Output Template 1 (Source_Data_Upload)",
    "2️⃣ Output Template 2 (Extracted_Data)",
    "3️⃣ Generate SPF (Stock Proposal Form)",
]

if "selected_menu" not in st.session_state:
    st.session_state["selected_menu"] = MENU_LIST[0]


def go_to(menu_name: str):
    st.session_state["selected_menu"] = menu_name


# ------------------------------------------------------------------
# Sidebar: Navigasi Utama & Konfigurasi Global
# ------------------------------------------------------------------
with st.sidebar:
    st.title("SEARAH Cataloging")
    selected_menu = st.radio(
        "Pilih Menu Kerja:",
        MENU_LIST,
        key="selected_menu",
    )

    st.divider()
    st.subheader("File Referensi")

    master_path = find_local(
        lambda filename:
            filename.startswith("material_standard")
            and filename.endswith(".xlsx")
    )
    master_raw = open(master_path, "rb").read() if master_path else None
    if master_raw is None:
        uploaded_master = st.file_uploader("Material Standard (.xlsx)", type=["xlsx"], key="ms")
        master_raw = uploaded_master.getvalue() if uploaded_master else None

    example_path = find_local(
        lambda filename:
            "extracted_data" in filename
            and "kosongan" not in filename
            and filename.endswith(".xlsx")
    )
    example_raw = open(example_path, "rb").read() if example_path else None
    if example_raw is None:
        uploaded_example = st.file_uploader("Contoh Terisi Template 2 (opsional)", type=["xlsx"], key="ex")
        example_raw = uploaded_example.getvalue() if uploaded_example else None

    st.subheader("Konfigurasi AI (Gemini)")
    api_key = secret("GEMINI_API_KEY") or st.text_input(
        "Gemini API Key",
        type="password",
        placeholder="Tempel API Key jika belum di Secrets",
    ).strip()

    use_ai = st.checkbox("Aktifkan AI", value=bool(api_key), disabled=not api_key)
    model = st.text_input("Model Gemini", value=secret("GEMINI_MODEL", "gemini-3.5-flash")).strip()
    batch_size = st.slider("Batch AI (Template 1)", min_value=1, max_value=10, value=5)
    limit_rows = st.number_input("Batas uji baris (0 = semua)", min_value=0, value=0, step=50)


# ------------------------------------------------------------------
# Load Master Data Global
# ------------------------------------------------------------------
if master_raw is None:
    st.warning("Silakan letakkan file `Material_Standard...xlsx` di folder aplikasi atau unggah di Sidebar.")
    st.stop()

try:
    df_nm, df_ms = cached_master(master_raw)
    labeled = cached_labeled(example_raw) if example_raw else None
    ctx = core.build_context(df_nm, labeled)
except Exception as error:
    st.error(f"Gagal membaca file referensi: {error}")
    st.stop()

with st.sidebar:
    st.success(
        f"Master Aktif: {len(df_nm):,} Noun|Modifier\n\n"
        f"Atribut Master: {len(df_ms):,} baris\n\n"
        f"Contoh Berlabel: {len(ctx['labels']):,} material"
    )


# ==================================================================
# 0. MENU AWAL (BERANDA)
# ==================================================================
if selected_menu == MENU_LIST[0]:
    st.title("SEARAH - Sistem Cataloging & SPF")
    st.markdown(
        "Pilih modul kerja di bawah ini atau melalui menu navigasi di sebelah kiri untuk memproses data material."
    )

    st.divider()
    col1, col2, col3 = st.columns(3)

    with col1:
        st.subheader("1️⃣ Output Template 1")
        st.write(
            "Membersihkan deskripsi material SAP, memetakan **Noun | Modifier** "
            "berdasarkan Master SKK Migas & AI, serta menghasilkan file **`Source_Data_Upload.xlsx`**."
        )
        t1_count = len(st.session_state["current_t1"]) if "current_t1" in st.session_state else 0
        st.caption(f"Status sesi ini: **{t1_count:,} material** siap.")
        st.button("Buka Menu Template 1 →", on_click=go_to, args=(MENU_LIST[1],), type="primary", use_container_width=True)

    with col2:
        st.subheader("2️⃣ Output Template 2")
        st.write(
            "Mengekstrak spesifikasi teknis material ke slot atribut baku dari **Material Standard** "
            "(13 kolom) dan menghasilkan file **`Extracted_Data.xlsx`**."
        )
        t2_count = st.session_state["df_t2_result"]["PSC Material Code"].nunique() if "df_t2_result" in st.session_state else 0
        st.caption(f"Status sesi ini: **{t2_count:,} material** terekstrak.")
        st.button("Buka Menu Template 2 →", on_click=go_to, args=(MENU_LIST[2],), type="primary", use_container_width=True)

    with col3:
        st.subheader("3️⃣ Generate SPF")
        st.write(
            "Menyusun dokumen **Stock Proposal Form (SPF)** lengkap dengan parameter stok "
            "(Min/Max, ROP, Criticality, Material Type) siap cetak/ekspor Excel."
        )
        spf_count = len(st.session_state["df_spf_result"]) if "df_spf_result" in st.session_state else 0
        st.caption(f"Status sesi ini: **{spf_count:,} item** SPF tersusun.")
        st.button("Buka Menu Generate SPF →", on_click=go_to, args=(MENU_LIST[3],), type="primary", use_container_width=True)

    st.divider()
    st.info(
        "💡 **Alur Kerja Terintegrasi:** Mulai dari **Menu 1 (Template 1)**, lalu hasilnya otomatis "
        "terbawa ke **Menu 2 (Template 2)** dan **Menu 3 (Generate SPF)** tanpa perlu mengunggah ulang file."
    )


# ==================================================================
# 1. MENU OUTPUT TEMPLATE 1
# ==================================================================
elif selected_menu == MENU_LIST[1]:
    st.title("1️⃣ Output Template 1 (Source_Data_Upload)")

    with st.expander("Opsi Pembersihan Teks & Pemetaan UoM", expanded=False):
        c_opt1, c_opt2, c_opt3 = st.columns([1.5, 1.5, 2])
        short_mode = c_opt1.radio(
            "Mode Short Description",
            ["Short SAP apa adanya", "Dari PO Text (dipotong rapi, maks 100)"],
        )
        ascii_symbols = c_opt2.checkbox("Ganti simbol ® ™ • ± ³ ke ASCII", value=False)
        uom_text = c_opt3.text_area("Pemetaan UoM (SAP=SKK)", placeholder="PCS=PC\nBX=BOX", height=70)

    uploaded_sap = st.file_uploader("Upload Data Material SAP (.xlsx / .csv)", type=["xlsx", "csv"], key="sap_up")

    if uploaded_sap is not None:
        raw_df = core.read_any(uploaded_sap.name, uploaded_sap.getvalue())
        cols = list(raw_df.columns)
        guess = core.detect_columns(cols)

        def col_idx(name):
            return cols.index(name) if name in cols else 0

        c1, c2, c3, c4 = st.columns(4)
        col_code = c1.selectbox("Kolom Kode Material", cols, index=col_idx(guess["code"]))
        col_short = c2.selectbox("Kolom Short (SAP)", cols, index=col_idx(guess["short"]))
        col_long = c3.selectbox("Kolom Long (PO Text)", cols, index=col_idx(guess["long"]))
        col_uom = c4.selectbox("Kolom UoM", cols, index=col_idx(guess["uom"]))

        st.caption(f"{len(raw_df):,} baris terbaca.")
        st.dataframe(raw_df.head(5), use_container_width=True)

        if st.button("Bersihkan & Klasifikasi Noun | Modifier", type="primary"):
            data = raw_df[raw_df[col_code].astype(str).str.strip() != ""].copy()
            if limit_rows > 0:
                data = data.head(int(limit_rows))

            uom_map = core.parse_uom_map(uom_text)
            rows = []
            for _, row in data.iterrows():
                uom = str(row[col_uom]).strip().upper()
                rows.append({
                    "code": str(row[col_code]).strip(),
                    "short": core.clean_text(row[col_short], ascii_symbols),
                    "long": core.clean_text(row[col_long], ascii_symbols),
                    "uom": uom_map.get(uom, uom),
                })

            ai_fn = None
            if use_ai and api_key:
                try:
                    ai_fn = core.make_gemini_batch_fn(api_key=api_key, model=model, ctx=ctx)
                except Exception as err:
                    st.error(f"AI tidak dapat diaktifkan: {err}")

            pbar = st.progress(0.0, text="Menerapkan label contoh & aturan...")

            def on_prog(done, total):
                if total > 0:
                    pbar.progress(min(done / total, 1.0), text=f"AI mengklasifikasi {done:,} / {total:,} item")

            output = core.classify_rows(rows, ctx, ai_batch=ai_fn, batch_size=batch_size, progress=on_prog)
            pbar.empty()

            base = pd.DataFrame({
                "Kode": [r["code"] for r in rows],
                "Kelas": [f"{o['noun']} | {o['modifier']}" if o["noun"] else "" for o in output],
                "Sumber": [o["source"] for o in output],
                "Keyakinan": [o["confidence"] for o in output],
                "Catatan": [o["note"] for o in output],
                "Short SAP": [r["short"] for r in rows],
                "Short (final)": [
                    r["short"] if short_mode.startswith("Short SAP") else core.make_short(r["long"] or r["short"])
                    for r in rows
                ],
                "Long": [r["long"] for r in rows],
                "Uom": [r["uom"] for r in rows],
            })
            st.session_state["base"] = base
            st.session_state["edits"] = {}
            st.session_state["run_id"] = st.session_state.get("run_id", 0) + 1

    if "base" in st.session_state:
        base = st.session_state["base"]
        edits = st.session_state["edits"]
        current = apply_edits(base, edits)

        st.subheader("Review Hasil Klasifikasi")
        only_review = st.checkbox("Hanya tampilkan yang perlu review (kelas kosong / keyakinan < 70)", value=True)
        mask = ((base["Kelas"].fillna("") == "") | (base["Keyakinan"] < 70)) if only_review else pd.Series(True, index=base.index)
        view = current.loc[mask].copy()
        pairs = [""] + [f"{n} | {m}" for n, m in zip(df_nm["NOUN"], df_nm["MODIFIER"])]

        edited = st.data_editor(
            view,
            key=f"ed_t1_{st.session_state['run_id']}_{only_review}",
            use_container_width=True,
            hide_index=True,
            height=450,
            disabled=[c for c in view.columns if c not in EDITABLE_T1],
            column_config={
                "Kelas": st.column_config.SelectboxColumn("Kelas (Noun | Modifier)", options=pairs, width="medium"),
                "Keyakinan": st.column_config.NumberColumn(format="%d"),
            },
        )

        for idx in edited.index:
            for col in EDITABLE_T1:
                val = edited.at[idx, col]
                val = "" if val is None or (isinstance(val, float) and pd.isna(val)) else str(val)
                if val != str(view.at[idx, col]):
                    edits.setdefault(idx, {})[col] = val

        st.session_state["edits"] = edits
        current = apply_edits(base, edits)
        st.session_state["current_t1"] = current

        counts = current["Sumber"].value_counts().to_dict()
        m = st.columns(6)
        m[0].metric("Total", f"{len(current):,}")
        for idx_m, src in enumerate(["CONTOH", "ATURAN", "AI", "MANUAL", "REVIEW"], start=1):
            m[idx_m].metric(src, f"{counts.get(src, 0):,}")

        t1_df = to_template1(current)
        issues = core.validate_template1(t1_df, ctx["valid"])
        if issues.empty:
            st.success("Semua baris lolos validasi Template 1.")
        else:
            st.warning(f"Ditemukan {len(issues):,} catatan validasi pada {issues['baris'].nunique():,} baris.")
            with st.expander("Lihat detail validasi"):
                st.dataframe(issues, use_container_width=True)

        c_dl1, c_dl2, c_next = st.columns([2, 2, 1.5])
        c_dl1.download_button(
            "Download Template 1 (Source_Data_Upload.xlsx)",
            data=core.build_template1_bytes(t1_df),
            file_name="Source_Data_Upload.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        c_dl2.download_button(
            "Download Laporan Audit (.csv)",
            data=current.to_csv(index=False).encode("utf-8-sig"),
            file_name="laporan_klasifikasi.csv",
            mime="text/csv",
            use_container_width=True,
        )
        c_next.button("Lanjut ke Template 2 →", on_click=go_to, args=(MENU_LIST[2],), use_container_width=True)


# ==================================================================
# 2. MENU OUTPUT TEMPLATE 2
# ==================================================================
elif selected_menu == MENU_LIST[2]:
    st.title("2️⃣ Output Template 2 (Extracted_Data / Atribut)")
    st.caption("Skema atribut (SEQ, Attribute, Mandatory) diambil otomatis dari Material Standard. PSC Code & Ref.No boleh dikosongkan.")

    c_p1, c_p2, c_p3 = st.columns([1.5, 1.5, 2])
    psc_code_in = c_p1.text_input("PSC Code (opsional)", value="", placeholder="Kosong atau misal: A095")
    ref_pref_in = c_p2.text_input("Prefix Ref.No (opsional)", value="", placeholder="Kosong atau misal: A095_1313")
    src_mode_t2 = c_p3.radio(
        "Sumber Material:",
        ["Ambil dari hasil Menu 1 (sesi ini)", "Upload file Template 1 / Excel (.xlsx)"],
    )

    items_t2: list[dict] = []
    if src_mode_t2.startswith("Ambil"):
        if "current_t1" not in st.session_state:
            st.info("Belum ada data di Menu 1. Silakan proses di Menu 1 dulu atau pilih opsi Upload Excel.")
        else:
            t1_df = to_template1(st.session_state["current_t1"])
            valid_t1 = t1_df[(t1_df["noun"].str.strip() != "") & (t1_df["modifier"].str.strip() != "")]
            st.write(f"Tersedia **{len(valid_t1):,} material** terklasifikasi dari Menu 1.")
            for _, r in valid_t1.iterrows():
                items_t2.append({
                    "code": str(r["code"]).strip(),
                    "noun": str(r["noun"]).strip().upper(),
                    "modifier": str(r["modifier"]).strip().upper(),
                    "long": str(r["long"]).strip(),
                    "uom": str(r["uom"]).strip().upper(),
                })
    else:
        up_t2 = st.file_uploader("Upload Excel Template 1 (.xlsx)", type=["xlsx"], key="up_t2_src")
        if up_t2 is not None:
            df_u2 = pd.read_excel(io.BytesIO(up_t2.getvalue()), dtype=str).fillna("")
            cols_u2 = list(df_u2.columns)

            def find_c(cands):
                for c in cols_u2:
                    if any(k in c.lower() for k in cands):
                        return cols_u2.index(c)
                return 0

            sc1, sc2, sc3, sc4, sc5 = st.columns(5)
            u_code = sc1.selectbox("PSC Material Code", cols_u2, index=find_c(["material code", "code"]))
            u_noun = sc2.selectbox("Noun", cols_u2, index=find_c(["noun"]))
            u_mod = sc3.selectbox("Modifier", cols_u2, index=find_c(["modifier"]))
            u_long = sc4.selectbox("Long Description", cols_u2, index=find_c(["long"]))
            u_uom = sc5.selectbox("UoM", cols_u2, index=find_c(["uom"]))

            for _, r in df_u2.iterrows():
                n_v = str(r[u_noun]).strip().upper()
                m_v = str(r[u_mod]).strip().upper()
                if n_v and m_v:
                    items_t2.append({
                        "code": str(r[u_code]).strip(),
                        "noun": n_v,
                        "modifier": m_v,
                        "long": str(r[u_long]).strip(),
                        "uom": str(r[u_uom]).strip().upper(),
                    })

    if items_t2 and st.button("Generate & Ekstrak Atribut Template 2", type="primary"):
        run_items = items_t2[:int(limit_rows)] if limit_rows > 0 else items_t2
        ai_ext = None
        if use_ai and api_key:
            try:
                ai_ext = core.make_gemini_extract_fn(api_key=api_key, model=model)
            except Exception as err:
                st.error(f"AI ekstraksi gagal diinisialisasi: {err}")

        pbar2 = st.progress(0.0, text="Mengekstrak atribut material...")

        def on_p2(done, total):
            if total > 0:
                pbar2.progress(min(done / total, 1.0), text=f"Mengekstrak {done:,} / {total:,} material...")

        df_t2, fatal_t2 = core.extract_template2_rows(
            items=run_items,
            df_ms=df_ms,
            ctx=ctx,
            ai_extract=ai_ext,
            psc_code=psc_code_in,
            ref_prefix=ref_pref_in,
            progress=on_p2,
        )
        pbar2.empty()
        st.session_state["df_t2_result"] = df_t2
        st.session_state["t2_fatal"] = fatal_t2

    if "df_t2_result" in st.session_state:
        df_t2_res = st.session_state["df_t2_result"]
        if st.session_state.get("t2_fatal"):
            st.error(f"AI terhenti: {st.session_state['t2_fatal']}")

        only_m = st.checkbox("Hanya tampilkan baris Mandatory (M)", value=False)
        view_t2 = df_t2_res[df_t2_res["Mandatory"] == "M"].copy() if only_m else df_t2_res.copy()

        edited_t2 = st.data_editor(
            view_t2,
            key=f"ed_t2_{only_m}",
            use_container_width=True,
            hide_index=True,
            height=480,
            disabled=[c for c in view_t2.columns if c not in EDITABLE_T2],
        )

        for idx in edited_t2.index:
            for col in EDITABLE_T2:
                val = edited_t2.at[idx, col]
                val = "" if val is None or (isinstance(val, float) and pd.isna(val)) else str(val)
                if col == "Value":
                    val = val.strip().upper()
                df_t2_res.at[idx, col] = val

        st.session_state["df_t2_result"] = df_t2_res

        t2_iss = core.validate_template2(df_t2_res)
        tot_mat = df_t2_res["PSC Material Code"].nunique()
        if t2_iss.empty:
            st.success(f"Seluruh atribut Mandatory (M) pada {tot_mat:,} material sudah lengkap.")
        else:
            st.warning(f"{len(t2_iss):,} atribut Mandatory (M) masih kosong pada {t2_iss['kode'].nunique():,} material.")
            with st.expander("Lihat atribut Mandatory yang kosong"):
                st.dataframe(t2_iss, use_container_width=True)

        c_dl_t2, c_to_spf = st.columns([2.5, 1.5])
        c_dl_t2.download_button(
            "Download Template 2 (Extracted_Data.xlsx)",
            data=core.build_template2_bytes(df_t2_res, sheet_name="Extracted_Data"),
            file_name="Extracted_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        c_to_spf.button("Lanjut ke Generate SPF →", on_click=go_to, args=(MENU_LIST[3],), use_container_width=True)


# ==================================================================
# 3. MENU GENERATE SPF (STOCK PROPOSAL FORM)
# ==================================================================
elif selected_menu == MENU_LIST[3]:
    st.title("3️⃣ Generate SPF (Stock Proposal Form)")
    st.caption(
        "Susun dokumen Stock Proposal Form sesuai format resmi SEARAH (SPF- NEW FORMAT) "
        "dengan perhitungan panjang Short Text 40 karakter SAP otomatis."
    )

    local_spf_tpl = find_local(
        lambda fn: ("spf_template" in fn or ("spf" in fn and fn.endswith(".xlsx")))
    )
    tpl_bytes = open(local_spf_tpl, "rb").read() if local_spf_tpl else None

    with st.expander("Template Excel SPF Resmi (SPF- NEW FORMAT)", expanded=(tpl_bytes is None)):
        if tpl_bytes is not None:
            st.success(f"Template resmi aktif dari folder aplikasi: `{os.path.basename(local_spf_tpl)}`")
        up_tpl = st.file_uploader(
            "Upload / Ganti Template SPF (.xlsx, misal: SPF Insulation Gasket 2026_Rev.1 31 Aug 2026.xlsx)",
            type=["xlsx"],
            key="up_spf_tpl",
        )
        if up_tpl is not None:
            tpl_bytes = up_tpl.getvalue()
            st.info(f"Menggunakan template unggahan: `{up_tpl.name}`")

    with st.expander("1. Informasi Header & Approval SPF (SPF- NEW FORMAT)", expanded=True):
        r1c1, r1c2, r1c3, r1c4 = st.columns([2, 1.5, 1.5, 1.5])
        spf_dept = r1c1.selectbox(
            "Originating Dept / Section",
            [
                "ID05 PCK2L OPERATION",
                "ID 05 PCK2L DRILLING",
                "ID01 PCML OPERATION",
                "ID 01 PCML DRILLING",
                "ID08 PCNM2L DRILLING",
            ],
        )
        default_plant = spf_dept.split()[0] if spf_dept.startswith("ID") and len(spf_dept.split()[0]) == 4 else "ID05"
        spf_req_date = r1c2.text_input("Stock Required Date", value=pd.Timestamp.now().strftime("%d/%m/%Y"))
        spf_orig_ref = r1c3.text_input("Originator Ref. No.", value="")
        spf_ref_no = r1c4.text_input("SPF Ref. No.", value="")

        r2c1, r2c2, r2c3, r2c4 = st.columns(4)
        spf_eqpt = r2c1.text_input("Equipment", value="Consumable Materials")
        spf_model = r2c2.text_input("Model or Type", placeholder="Misal: Insulation Gasket Set")
        spf_submodel = r2c3.text_input("Sub-Model", placeholder="Misal: Very Critical Service")
        spf_std = r2c4.text_input("Applicable Standard", value="Searah Ketapang II Ltd")

        r3c1, r3c2, r3c3, r3c4 = st.columns(4)
        spf_size = r3c1.text_input("Size / Capacity", value="Various")
        spf_mfr = r3c2.text_input("Manufacturer", value="Various")
        spf_tag = r3c3.text_input("Equipment Tag No.", value="Various")
        spf_asset = r3c4.text_input("Asset No.", value="Various P2B")

        r4c1, r4c2 = st.columns([3, 1])
        spf_reason = r4c1.text_input(
            "Reasons for Proposal",
            placeholder="Misal: PROVISION INSULATION GASKET FOR SEARAH KETAPANG II LTD.",
        )
        spf_eq_crit = r4c2.selectbox("Eqpmt Criticality (1/2/3)", ["1", "2", "3"], index=0)

        st.markdown("**Blok Persetujuan (Footer SPF)**")
        s1, s2, s3, s4 = st.columns(4)
        spf_init_by = s1.text_input("Initiated By (Engineer/Supv)", placeholder="Misal: TEGUH CAHYO / SLAMET SUBAGIO")
        spf_end_mgr = s2.text_input("Endorsed By (Manager)", placeholder="Misal: BENTLY NURPATRIAWIRA")
        spf_end_srmgr = s3.text_input("Endorsed By (Senior Manager)", placeholder="Misal: AULIA ARIF HIDAYAT")
        spf_init_date = s4.text_input("Tanggal Tanda Tangan", value=pd.Timestamp.now().strftime("%d %b %Y").upper())

    st.subheader("2. Sumber Data Material")
    c_src1, c_src2, c_src3 = st.columns([2, 1.2, 1.3])
    src_spf = c_src1.radio(
        "Pilih Sumber Item:",
        [
            "Gunakan hasil katalogisasi sesi ini (Menu 1)",
            "Upload file Excel daftar material / Template 1 (.xlsx)",
        ],
        horizontal=True,
    )
    def_mat_type = c_src2.text_input("Default Material Type", value="PM1270602")
    def_cert = c_src3.text_input("Default Certificate", value="COC/COO (2,1)")

    spf_base_rows = []
    if src_spf.startswith("Gunakan"):
        if "current_t1" not in st.session_state:
            st.info("Belum ada data dari Menu 1. Silakan proses di Menu 1 dahulu atau pilih opsi Upload Excel.")
        else:
            t1_df = to_template1(st.session_state["current_t1"])
            for i, r in t1_df.iterrows():
                short_40 = core.make_short_40(str(r["short"]), str(r["long"]), limit=40)
                nm_label = f"{str(r['noun']).strip()} | {str(r['modifier']).strip()}".strip(" |")
                spf_base_rows.append({
                    "No.": i + 1,
                    "Plant": default_plant,
                    "LEN": len(short_40),
                    "Short Description": short_40,
                    "Long Description": str(r["long"]).strip(),
                    "Criticality": "D",
                    "Part / Category": nm_label,
                    "UoM": str(r["uom"]).strip().upper() or "EA",
                    "Qty Per Eqpt": "1",
                    "Qty Total": "1",
                    "Est. Annual Usage": "0",
                    "Reorder Level": "1",
                    "Max Level": "2",
                    "Initial Order Qty": "2",
                    "Est. Unit Price (IDR)": "0",
                    "Material Type": def_mat_type.strip().upper(),
                    "Val. Class": "PD1",
                    "Certificate": def_cert.strip(),
                    "Lead Time": "6 WEEK",
                    "Material Code": str(r["code"]).strip(),
                })
    else:
        up_spf = st.file_uploader("Upload Excel Daftar Material / Template 1 (.xlsx)", type=["xlsx"], key="up_spf_file")
        if up_spf is not None:
            df_s_in = pd.read_excel(io.BytesIO(up_spf.getvalue()), dtype=str).fillna("")
            cols_s = list(df_s_in.columns)

            def idx_s(cands):
                for c in cols_s:
                    if any(k in c.lower() for k in cands):
                        return cols_s.index(c)
                return 0

            sc1, sc2, sc3, sc4, sc5 = st.columns(5)
            s_code = sc1.selectbox("Kode Material (opsional)", cols_s, index=idx_s(["material code", "code", "material"]))
            s_short = sc2.selectbox("Short Description", cols_s, index=idx_s(["short", "description"]))
            s_long = sc3.selectbox("Long Description", cols_s, index=idx_s(["long", "po text"]))
            s_uom = sc4.selectbox("UoM", cols_s, index=idx_s(["uom", "unit"]))
            s_cat = sc5.selectbox("Part / Category / Noun", cols_s, index=idx_s(["noun", "part", "category"]))

            for i, r in df_s_in.iterrows():
                raw_s = str(r[s_short]).strip()
                raw_l = str(r[s_long]).strip()
                short_40 = core.make_short_40(raw_s, raw_l, limit=40)
                spf_base_rows.append({
                    "No.": i + 1,
                    "Plant": default_plant,
                    "LEN": len(short_40),
                    "Short Description": short_40,
                    "Long Description": raw_l or raw_s,
                    "Criticality": "D",
                    "Part / Category": str(r[s_cat]).strip(),
                    "UoM": str(r[s_uom]).strip().upper() or "EA",
                    "Qty Per Eqpt": "1",
                    "Qty Total": "1",
                    "Est. Annual Usage": "0",
                    "Reorder Level": "1",
                    "Max Level": "2",
                    "Initial Order Qty": "2",
                    "Est. Unit Price (IDR)": "0",
                    "Material Type": def_mat_type.strip().upper(),
                    "Val. Class": "PD1",
                    "Certificate": def_cert.strip(),
                    "Lead Time": "6 WEEK",
                    "Material Code": str(r[s_code]).strip(),
                })

    if spf_base_rows:
        df_spf_init = pd.DataFrame(spf_base_rows)
        st.subheader(f"3. Tabel Draft SPF ({len(df_spf_init):,} Material)")
        st.caption(
            "Kolom LEN otomatis menghitung panjang Short Description (maks. 40 karakter SAP). "
            "Edit parameter stok, harga (IDR), criticality (V/V1/E/D), dan deskripsi langsung di tabel:"
        )

        edited_spf = st.data_editor(
            df_spf_init,
            key="editor_spf_table",
            use_container_width=True,
            hide_index=True,
            height=480,
            disabled=["No.", "LEN"],
            column_config={
                "Short Description": st.column_config.TextColumn("Short Description (Max 40)", max_chars=40, width="medium"),
                "Long Description": st.column_config.TextColumn("Long Description", width="large"),
                "Criticality": st.column_config.SelectboxColumn(
                    "Criticality (V/E/D)",
                    options=["V", "V1", "E", "D"],
                ),
                "Val. Class": st.column_config.SelectboxColumn(
                    "Val. Class",
                    options=["PD1", "PD0"],
                ),
            },
        )
        edited_spf["LEN"] = edited_spf["Short Description"].fillna("").astype(str).str.len()
        over_40 = int((edited_spf["LEN"] > 40).sum())
        if over_40 > 0:
            st.warning(f"{over_40} baris memiliki Short Description > 40 karakter (akan dipotong rapi saat ekspor).")

        st.session_state["df_spf_result"] = edited_spf

        fill_mm_code = st.checkbox(
            "Isi juga kolom 'MATERIAL NO.' (kolom BF, bagian MM Dept) dengan Kode Material SAP",
            value=False,
        )

        spf_bytes = core.build_spf_bytes(
            edited_spf,
            meta={
                "spf_ref": spf_ref_no,
                "originating_dept": spf_dept,
                "originator_ref": spf_orig_ref,
                "required_date": spf_req_date,
                "default_plant": default_plant,
                "equipment": spf_eqpt,
                "size_capacity": spf_size,
                "equipment_tag": spf_tag,
                "model_type": spf_model,
                "manufacturer": spf_mfr,
                "asset_no": spf_asset,
                "sub_model": spf_submodel,
                "applicable_std": spf_std,
                "eqpmt_criticality": spf_eq_crit,
                "reason": spf_reason,
                "initiated_by": spf_init_by,
                "initiated_date": spf_init_date,
                "endorsed_mgr": spf_end_mgr,
                "endorsed_srmgr": spf_end_srmgr,
                "fill_mat_code": fill_mm_code,
            },
            template_bytes=tpl_bytes,
        )
        st.download_button(
            "Download Stock Proposal Form (SPF_NEW_FORMAT.xlsx)",
            data=spf_bytes,
            file_name=f"SPF_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
