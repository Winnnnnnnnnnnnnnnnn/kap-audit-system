import io
import json
import os
import shutil
import subprocess
import tempfile
import zipfile
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from google import genai
from google.genai import types
from pypdf import PdfReader

# 0. LOAD ENVIRONMENT VARIABLES & SECRETS
load_dotenv()
api_key = os.environ.get("GEMINI_API_KEY") or st.secrets.get("GEMINI_API_KEY", "")

# KONFIGURASI HALAMAN STREAMLIT
st.set_page_config(
    page_title="Audit Assistant - KAP Working Papers",
    page_icon="📊",
    layout="wide"
)

# --- SISTEM PROTEKSI GERBANG PASSWORD (TETAP AKTIF) ---
def check_password():
    """Mengembalikan True jika pengguna memasukkan kata sandi yang benar."""
    def password_entered():
        target_password = st.secrets.get("APP_PASSWORD")
        if st.session_state["password_input"] == target_password:
            st.session_state["password_correct"] = True
            del st.session_state["password_input"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.subheader("🔒 Akses Terbatas - Portal Audit KAP")
        st.text_input("Masukkan Kata Sandi:", type="password", on_change=password_entered, key="password_input")
        return False
    elif not st.session_state["password_correct"]:
        st.subheader("🔒 Akses Terbatas - Portal Audit KAP")
        st.text_input("Masukkan Kata Sandi:", type="password", on_change=password_entered, key="password_input")
        st.error("Kata sandi salah. Silakan hubungi admin.")
        return False
    else:
        return True

if not check_password():
    st.stop()

# --- KODE APLIKASI UTAMA BERJALAN DI BAWAH SINI ---

# 1. STANDAR AKUN & KODE INDEKS KKP KAP (ALSINDO TEMPLATE)
AUDIT_INDEX_CATALOG = {
    "A-1": "Kas",
    "A-2": "Bank (IDR / USD)",
    "A-3": "Piutang Usaha",
    "A-4": "Piutang Lain-Lain",
    "A-5": "Persediaan",
    "A-6": "Uang Muka",
    "A-7": "Biaya Dibayar Dimuka / Pajak Dimuka",
    "A-8": "Aset Tetap (Nilai Perolehan & Akumulasi)",
    "A-9": "Aset Lain-lain",
    "B-1": "Utang Usaha",
    "B-2": "Utang Pajak",
    "B-3": "Utang Lain-lain",
    "B-4": "Pendapatan Diterima Dimuka",
    "B-5": "Utang Bank",
    "B-6": "Utang Leasing",
    "EQ-1": "Modal Saham",
    "EQ-2": "Saldo Laba",
    "REV": "Pendapatan Usaha",
    "C-1": "Beban Pokok Pendapatan (HPP)",
    "D-1": "Beban Penjualan",
    "D-2": "Beban Umum dan Administrasi",
    "OTHER_REV": "Pendapatan Diluar Usaha",
    "OTHER_EXP": "Beban Diluar Usaha",
    "TAX_EXP": "Beban Pajak Penghasilan"
}

# 2. HELPER DETEKSI BARIS HEADER OTOMATIS (MURNI STRUKTURAL)
def detect_table_header_index(df_sample: pd.DataFrame) -> int:
    best_row_idx = 0
    max_score = -1.0

    for idx, row in df_sample.iterrows():
        non_null_values = [val for val in row if pd.notna(val) and str(val).strip() != ""]
        num_non_null = len(non_null_values)

        if num_non_null < 2:
            continue

        density = num_non_null / len(row)
        text_count = sum(1 for val in non_null_values if isinstance(val, str) and not val.strip().replace(".", "", 1).isdigit())
        text_ratio = text_count / num_non_null
        unique_ratio = len(set(non_null_values)) / num_non_null

        score = (density * 0.5) + (text_ratio * 0.3) + (unique_ratio * 0.2)

        if score > max_score and density >= 0.3:
            max_score = score
            best_row_idx = idx

    return best_row_idx

# 3. HELPER EKSTRAKSI ARSIP (.RAR & .ZIP)
def extract_archive_files(uploaded_file):
    extracted_files = {"excel": {}, "pdf": {}}
    file_ext = uploaded_file.name.split(".")[-1].lower()

    tmpdir = tempfile.mkdtemp()
    temp_archive_path = os.path.join(tmpdir, uploaded_file.name)
    with open(temp_archive_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    success = False
    if file_ext == "zip":
        with zipfile.ZipFile(temp_archive_path, "r") as z:
            z.extractall(tmpdir)
        success = True
    elif file_ext == "rar":
        winrar_paths = [
            r"C:\Program Files\WinRAR\WinRAR.exe",
            r"C:\Program Files\WinRAR\UnRAR.exe",
            r"C:\Program Files (x86)\WinRAR\WinRAR.exe"
        ]
        winrar_cmd = next((p for p in winrar_paths if os.path.exists(p)), None)

        if winrar_cmd:
            try:
                subprocess.run([winrar_cmd, "x", "-y", temp_archive_path, tmpdir], check=True, stdout=subprocess.DEVNULL)
                success = True
            except Exception:
                pass

        if not success:
            try:
                subprocess.run(["tar", "-xf", temp_archive_path, "-C", tmpdir], check=True, stdout=subprocess.DEVNULL)
                success = True
            except Exception:
                pass

        if not success:
            try:
                import rarfile
                if winrar_cmd:
                    rarfile.UNRAR_TOOL = winrar_cmd
                with rarfile.RarFile(temp_archive_path) as rf:
                    rf.extractall(tmpdir)
                success = True
            except Exception:
                pass

        if not success:
            try:
                import patoolib
                patoolib.extract_archive(temp_archive_path, outdir=tmpdir, verbosity=-1)
                success = True
            except Exception as e:
                st.error(f"Gagal mengekstrak .rar: {e}. Kamu bisa mengekstraknya manual dan langsung mengunggah file Excel/PDF.")

    if success:
        for root, _, files in os.walk(tmpdir):
            for f in files:
                full_path = os.path.join(root, f)
                lower_f = f.lower()
                if lower_f.endswith((".xlsx", ".xls")) and not lower_f.startswith("~$"):
                    with open(full_path, "rb") as b:
                        extracted_files["excel"][f] = io.BytesIO(b.read())
                elif lower_f.endswith(".pdf") and not lower_f.startswith("~$"):
                    with open(full_path, "rb") as b:
                        extracted_files["pdf"][f] = io.BytesIO(b.read())

    return extracted_files

# 4. HELPER PARSING TEKS PDF
def extract_text_from_pdf(pdf_stream, max_pages=10):
    reader = PdfReader(pdf_stream)
    text_content = []
    total_pages = min(len(reader.pages), max_pages)
    for i in range(total_pages):
        page_text = reader.pages[i].extract_text() or ""
        text_content.append(f"--- Halaman {i+1} ---\n{page_text}")
    return "\n".join(text_content)

# 5. SANITASI DATAFRAME AGAR KOMPATIBEL DENGAN APACHE ARROW
def sanitize_dataframe(df):
    clean_df = df.copy()
    for col in clean_df.columns:
        if pd.api.types.is_datetime64_any_dtype(clean_df[col]) or clean_df[col].dtype == "object":
            clean_df[col] = clean_df[col].astype(str).replace("nan", "").replace("None", "")
    return clean_df

# 6. PEMBERSIH ANGKA / NOMINAL
def clean_currency_to_float(series: pd.Series) -> pd.Series:
    """Membersihkan simbol mata uang, format ribuan koma/titik, dan tanda kurung negatif."""
    if pd.api.types.is_numeric_dtype(series):
        return series.fillna(0)

    def parse_val(val):
        if pd.isna(val):
            return 0.0
        val = str(val).strip()
        if val in ["-", "", "None", "nan"]:
            return 0.0

        is_negative = False
        if val.startswith("(") and val.endswith(")"):
            is_negative = True
            val = val[1:-1].strip()

        for prefix in ["Rp", "IDR", "rp", "idr"]:
            val = val.replace(prefix, "").strip()

        if "," in val and "." not in val:
            val = val.replace(",", "")
        elif "," in val and "." in val:
            if val.rfind(",") < val.rfind("."):
                val = val.replace(",", "")
            else:
                val = val.replace(".", "").replace(",", ".")
        elif "." in val:
            parts = val.split(".")
            if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
                val = val.replace(".", "")

        try:
            res = float(val)
            return -res if is_negative else res
        except Exception:
            return 0.0

    return series.apply(parse_val)

# 7. GEMINI ENGINES
def map_client_accounts(client_account_names, key):
    client = genai.Client(api_key=key)
    prompt = f"""
    Kamu adalah auditor senior KAP.
    Petakan daftar nama akun internal klien berikut:
    {client_account_names}

    Ke salah satu Kode Indeks KKP standar kami:
    {json.dumps(AUDIT_INDEX_CATALOG, indent=2)}

    Keluarkan HANYA format JSON valid tanpa tanda backtick atau markdown:
    [
      {{"akun_klien": "...", "kode_indeks": "...", "nama_indeks": "..."}}
    ]
    """
    response = client.models.generate_content(
        model="gemini-3.6-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.0
        )
    )
    return json.loads(response.text)

def analyze_pdf_with_gemini(pdf_text, key):
    client = genai.Client(api_key=key)
    prompt = f"""
    Kamu adalah auditor senior KAP.
    Analisis dokumen keuangan / rekening koran PDF berikut:
    {pdf_text[:15000]}

    Tugasmu:
    1. Identifikasi Entitas / Nama Bank / No Rekening (jika ada).
    2. Identifikasi Saldo Awal, Total Mutasi Masuk (Kredit), Total Mutasi Keluar (Debit), dan Saldo Akhir.
    3. Cocokkan ke nomor indeks KKP KAP (misal Kas Bank = A-2, Pendapatan = REV).
    4. Buat 3 temuan audit penting.

    Gunakan Bahasa Indonesia yang formal dan terstruktur rapi.
    """
    response = client.models.generate_content(
        model="gemini-3.6-flash",
        contents=prompt
    )
    return response.text

# 8. BUILD EXCEL LAPORAN HASIL
def create_audit_export_excel(df_mapped, df_rekap):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_rekap.to_excel(writer, sheet_name="Rekap Indeks KKP", index=False)
        df_mapped.to_excel(writer, sheet_name="Detail Mapping Akun", index=False)
    return output.getvalue()

# --- TAMPILAN USER INTERFACE (STREAMLIT) ---
st.title("📊 Sistem Audit & Pemetaan KKP Laporan Keuangan")
st.markdown("Mendukung berkas Excel (`.xlsx`, `.xls`), Dokumen (`.pdf`), dan Paket Arsip (`.rar`, `.zip`).")

# Sidebar: HANYA STATUS KONEKSI (KOTAK INPUT API KEY DIHILANGKAN)
with st.sidebar:
    st.header("⚙️ Status Sistem")
    if api_key:
        st.success("API Key Aktif (Terhubung)")
    else:
        st.error("GEMINI_API_KEY belum disetel di .env / secrets.")

uploaded_file = st.file_uploader(
    "Unggah Berkas Klien (.xlsx, .xls, .pdf, .rar, .zip):",
    type=["xlsx", "xls", "pdf", "rar", "zip"]
)

active_excel = None
active_pdf = None

if uploaded_file is not None:
    fname = uploaded_file.name.lower()

    if fname.endswith((".rar", ".zip")):
        with st.spinner("Mengekstrak isi arsip..."):
            extracted = extract_archive_files(uploaded_file)
            total_excel = len(extracted["excel"])
            total_pdf = len(extracted["pdf"])

            st.info(f"Ditemukan di dalam arsip: **{total_excel} Excel** dan **{total_pdf} PDF**.")

            if total_excel > 0:
                pilihan_excel = st.selectbox("Pilih file Excel yang ingin diaudit:", options=list(extracted["excel"].keys()))
                active_excel = extracted["excel"][pilihan_excel]

            if total_pdf > 0:
                pilihan_pdf = st.selectbox("Pilih file PDF (misal Rekening Koran):", options=list(extracted["pdf"].keys()))
                active_pdf = extracted["pdf"][pilihan_pdf]

    elif fname.endswith((".xlsx", ".xls")):
        active_excel = uploaded_file

    elif fname.endswith(".pdf"):
        active_pdf = uploaded_file

# --- PROSES EXCEL (MAPPING AKUN KKP) ---
if active_excel is not None:
    st.divider()
    st.subheader("📑 Modul Audit Excel (Trial Balance / Akun)")
    try:
        xls = pd.ExcelFile(active_excel)
        sheet_choice = st.selectbox("Pilih Sheet:", options=xls.sheet_names)

        df_sample = pd.read_excel(xls, sheet_name=sheet_choice, header=None, nrows=20)
        detected_header_idx = detect_table_header_index(df_sample)

        df_raw = pd.read_excel(xls, sheet_name=sheet_choice, header=detected_header_idx)
        df_raw = df_raw.dropna(how="all", axis=1)
        df_raw.columns = [str(c).strip() for c in df_raw.columns]

        df_display = sanitize_dataframe(df_raw.dropna(how="all"))
        st.dataframe(df_display.head(5), use_container_width=True)

        cols = list(df_raw.columns)

        default_acc_idx = 0
        for i, c in enumerate(cols):
            c_clean = str(c).strip().lower()
            if any(k in c_clean for k in ["keterangan", "uraian", "nama akun", "deskripsi", "account"]):
                default_acc_idx = i
                break

        default_val_idx = min(1, len(cols) - 1)
        for target in ["saldo", "debet", "debit", "nominal", "kredit"]:
            found = False
            for i, c in enumerate(cols):
                if str(c).strip().lower() == target:
                    default_val_idx = i
                    found = True
                    break
            if found:
                break

        c1, c2 = st.columns(2)
        with c1:
            col_akun = st.selectbox("Kolom Nama Akun / Keterangan:", options=cols, index=default_acc_idx)
        with c2:
            col_saldo = st.selectbox("Kolom Nominal Saldo / Debet / Kredit:", options=cols, index=default_val_idx)

        if st.button("🚀 Petakan Akun Excel via Gemini", type="primary"):
            if not api_key:
                st.error("GEMINI_API_KEY tidak ditemukan di environment.")
            else:
                with st.spinner("Memproses audit dan mapping akun dengan Gemini 3.6 Flash..."):
                    df_clean = df_raw.copy()

                    df_clean[col_saldo] = clean_currency_to_float(df_clean[col_saldo])

                    first_col = df_clean.columns[0]
                    first_str = df_clean[first_col].astype(str).str.strip()

                    is_acc_header = (
                        first_str.str.contains(r"^\d{4}\.", regex=True) & 
                        ~first_str.str.contains(r"^\d{4}-\d{2}", regex=True)
                    )

                    has_nested_accounts = is_acc_header.any()

                    if has_nested_accounts:
                        df_clean["Akun_Terdeteksi"] = df_clean[first_col].where(is_acc_header).ffill()
                        df_clean["Akun_Final"] = df_clean["Akun_Terdeteksi"]
                        df_transaksi = df_clean[~is_acc_header & (df_clean[col_saldo] != 0)].copy()
                    else:
                        df_clean["Akun_Final"] = df_clean[col_akun]
                        df_transaksi = df_clean[df_clean[col_saldo] != 0].copy()

                    unique_accs = df_transaksi["Akun_Final"].dropna().astype(str).unique().tolist()
                    mapping_res = map_client_accounts(unique_accs, api_key)

                    map_dict = {item["akun_klien"]: item["kode_indeks"] for item in mapping_res}
                    df_transaksi["Kode_Indeks"] = df_transaksi["Akun_Final"].astype(str).map(map_dict).fillna("LAIN-LAIN")
                    df_transaksi["Kategori_Standar"] = df_transaksi["Kode_Indeks"].map(AUDIT_INDEX_CATALOG).fillna("Belum Terdefinisi")

                    if has_nested_accounts and "saldo" in col_saldo.lower():
                        saldo_per_akun = (
                            df_transaksi.groupby(["Akun_Final", "Kode_Indeks", "Kategori_Standar"])[col_saldo]
                            .last()
                            .reset_index()
                        )
                        rekap_df = saldo_per_akun.groupby(["Kode_Indeks", "Kategori_Standar"])[col_saldo].sum().reset_index()
                    else:
                        rekap_df = df_transaksi.groupby(["Kode_Indeks", "Kategori_Standar"])[col_saldo].sum().reset_index()

                    rekap_df.columns = ["Kode Indeks", "Pos Laporan Standar", "Total Saldo Teraudit"]

                    st.session_state["rekap_df"] = rekap_df
                    st.session_state["df_mapped"] = sanitize_dataframe(df_transaksi)
                    st.success("Mapping akun dan kalkulasi saldo akhir berhasil diselesaikan!")

    except Exception as e:
        st.error(f"Gagal memproses Excel: {e}")

# Tampilan Hasil Excel
if "rekap_df" in st.session_state:
    st.dataframe(st.session_state["rekap_df"].style.format({"Total Saldo Teraudit": "Rp {:,.0f}"}), use_container_width=True)
    excel_exp = create_audit_export_excel(st.session_state["df_mapped"], st.session_state["rekap_df"])
    st.download_button(
        label="📥 Unduh File Excel KKP Hasil Audit",
        data=excel_exp,
        file_name="Hasil_Mapping_Audit_KKP.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# --- PROSES PDF (REKENING KORAN / LAPORAN AUDIT) ---
if active_pdf is not None:
    st.divider()
    st.subheader("📄 Modul Analisis Dokumen PDF (Rekening Koran / Bukti Audit)")

    if st.button("🔍 Ekstrak & Audit PDF dengan Gemini"):
        if not api_key:
            st.error("GEMINI_API_KEY tidak ditemukan di environment.")
        else:
            with st.spinner("Membaca halaman PDF dan meminta insight auditor dari Gemini..."):
                try:
                    pdf_text = extract_text_from_pdf(active_pdf)
                    audit_analysis = analyze_pdf_with_gemini(pdf_text, api_key)

                    st.success("Analisis PDF Selesai!")
                    st.markdown("### Ringkasan Hasil Pemeriksaan Dokumen:")
                    st.markdown(audit_analysis)

                    with st.expander("Lihat Ekstrak Teks Mentah dari PDF"):
                        st.text(pdf_text[:3000] + "\n\n... (dipotong untuk pratinjau)")
                except Exception as ex:
                    st.error(f"Gagal membaca PDF: {ex}")
