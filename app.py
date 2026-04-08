import streamlit as st
import pandas as pd
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Nagsa Master Data Audit V3", layout="wide")

st.title("🔍 Master Data Audit System - Unifarm Project")
st.write("Sistem Audit: Baris merah menunjukkan data yang WAJIB diperbaiki oleh Admin.")

uploaded_file = st.file_uploader("Unggah File Master Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 1. LOAD DATA
        df = pd.read_excel(uploaded_file)
        # Normalisasi nama kolom (Hapus spasi & Upper)
        df.columns = df.columns.str.strip().str.upper()
        
        # Buat dataframe kerja
        working_df = df.copy()
        working_df['ERROR_CATEGORY'] = ""

        # --- LOGIKA AUDIT (PEMBERIAN LIST MERAH) ---

        # 1. Audit Inisial Nama (Rule 1)
        # Cek apakah nama depan masih mengandung Alfamart, Indomaret, Alfamidi
        invalid_prefixes = ('ALFAMART', 'INDOMARET', 'ALFAMIDI')
        mask_name_issue = working_df['CUST NAME VERSI MASTER'].astype(str).str.upper().str.startswith(invalid_prefixes, na=False)
        working_df.loc[mask_name_issue, 'ERROR_CATEGORY'] += "NAMA_BELUM_INISIAL; "

        # 2. Audit Double (Rule 2)
        # Kecualikan kode tertentu dari pengecekan double
        exclusion_list = ['NOO', 'DISTRIBUTOR', 'KANTOR', 'NEW', 'N00', '#N/A']
        mask_eligible_dup = ~working_df['CUST CODE VERSI MASTER'].astype(str).str.upper().isin(exclusion_list)
        mask_duplicate = working_df.duplicated(subset=['CUST CODE VERSI MASTER', 'CUST NAME VERSI MASTER', 'ADDRESS'], keep='first') & mask_eligible_dup
        working_df.loc[mask_duplicate, 'ERROR_CATEGORY'] += "DATA_DOUBLE; "

        # 3. Audit Alamat (Rule 3)
        # Blank, 0, #N/A, atau kurang dari 3 huruf
        def audit_address(addr):
            a = str(addr).strip().upper()
            if a in ['', '0', '#N/A', 'NAN', 'NONE']: return True
            if len(a) < 3: return True
            return False
        
        mask_bad_addr = working_df['ADDRESS'].apply(audit_address)
        working_df.loc[mask_bad_addr, 'ERROR_CATEGORY'] += "ALAMAT_INVALID; "

        # 4. Audit Account Alignment (Rule 4 & Catatan)
        def audit_account(row):
            name = str(row['CUST NAME VERSI MASTER']).upper()
            acc = str(row['ACCOUNT']).upper()
            
            # SAT -> Harus mengandung ALFAMART
            if name.startswith('SAT'):
                if 'ALFAMART' not in acc: return True
            # IDM -> Harus mengandung INDOMARET
            elif name.startswith('IDM'):
                if 'INDOMARET' not in acc: return True
            # MIDI -> Harus mengandung ALFAMIDI
            elif name.startswith('MIDI'):
                if 'ALFAMIDI' not in acc: return True
            return False

        mask_acc_issue = working_df.apply(audit_account, axis=1)
        working_df.loc[mask_acc_issue, 'ERROR_CATEGORY'] += "ACCOUNT_TIDAK_SESUAI; "

        # --- STATISTIK DASHBOARD ---
        st.subheader("📊 Statistik Master Data")
        total_rows = len(working_df)
        rows_with_error = (working_df['ERROR_CATEGORY'] != "").sum()
        rows_clean = total_rows - rows_with_error

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Semua Data", f"{total_rows} Baris")
        col2.metric("Data Bersih (Putih)", f"{rows_clean} Baris")
        col3.metric("Data Bermasalah (Merah)", f"{rows_with_error} Baris")
        col4.metric("Akurasi Data", f"{(rows_clean/total_rows*100):.1f}%")

        # --- GENERATE EXCEL WITH HIGHLIGHT ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            working_df.to_excel(writer, index=False, sheet_name='AUDIT_OUTLET')
            
            workbook  = writer.book
            worksheet = writer.sheets['AUDIT_OUTLET']
            
            # Format Merah untuk baris bermasalah
            red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            
            # Identifikasi kolom ERROR_CATEGORY untuk trigger warna
            error_col_idx = working_df.columns.get_loc("ERROR_CATEGORY")
            from xlsxwriter.utility import xl_col_to_name
            col_letter = xl_col_to_name(error_col_idx)
            
            # Terapkan highlight merah jika ada teks di kolom ERROR_CATEGORY
            worksheet.conditional_format(1, 0, total_rows, len(working_df.columns) - 1,
                                         {'type':     'formula',
                                          'criteria': f'=${col_letter}2<>""',
                                          'format':   red_format})
            
            worksheet.freeze_panes(1, 0) # Kunci header

        st.divider()
        st.download_button(
            label="📥 Download Hasil Audit (Highlight Merah)",
            data=output.getvalue(),
            file_name="LAPORAN_AUDIT_MASTER_DATA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # PRATINJAU DI WEB
        st.subheader("Preview Data (50 Baris Pertama)")
        def highlight_rows(s):
            return ['background-color: #ffcccc' if s.ERROR_CATEGORY != "" else '' for _ in s]
        
        st.dataframe(working_df.head(50).style.apply(highlight_rows, axis=1))

    except Exception as e:
        st.error(f"Terjadi kesalahan teknis: {e}")
        st.info("Pastikan kolom sesuai acuan: CUST CODE VERSI MASTER, CUST NAME VERSI MASTER, ACCOUNT, ADDRESS")