# SEARAH Sistem Cataloging & Stock Proposal Form (SPF)

Aplikasi berbasis Streamlit untuk otomasi standarisasi katalog material SKK Migas (e-Cataloging) dan pembuatan dokumen Stock Proposal Form (SPF).

## Struktur Modul
1. **Menu Awal (Beranda):** Ringkasan status file referensi (`Material_Standard` & contoh berlabel), status koneksi Gemini AI, dan navigasi cepat.
2. **Menu 1 - Output Template 1 (`Source_Data_Upload.xlsx`):** Pembersihan deskripsi material SAP, pemetaan `Noun | Modifier` berbasis contoh, aturan, dan Gemini AI, serta validasi karakter.
3. **Menu 2 - Output Template 2 (`Extracted_Data.xlsx`):** Pembentukan 13 kolom baku e-Cataloging SKK Migas dengan skema atribut (`SEQ`, `Attribute`, `Mandatory`) langsung dari `Material_Standard` dan ekstraksi nilai spesifikasi otomatis via AI.
4. **Menu 3 - Generate SPF (`Stock_Proposal_Form.xlsx`):** Penyusunan draft Stock Proposal Form terintegrasi dengan hasil katalogisasi maupun upload Excel mandiri.

## Instalasi & Menjalankan Lokal
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Konfigurasi Secrets (Streamlit Cloud / Lokal)
Buat file `.streamlit/secrets.toml` atau atur pada menu **Settings > Secrets** di Streamlit Cloud:
```toml
GEMINI_API_KEY = "ISI_API_KEY_GEMINI_ANDA"
GEMINI_MODEL = "gemini-3.5-flash"
```

## Referensi Data
Letakkan file `Material_Standard...xlsx` di folder utama repository agar otomatis dimuat saat aplikasi dijalankan.
