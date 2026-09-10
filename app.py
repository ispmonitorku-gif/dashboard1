import streamlit as st
import pandas as pd
import numpy as np
import re
import math
import copy
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="OTDR Dashboard NIX-PCM", layout="wide")

st.title("📊 Mesin Generator Dashboard OTDR - NIX PCM")
st.write("Sistem otomatis yang mengonversi file raw `Event_table` menjadi Dashboard identik 100% menggunakan template asli Anda.")

# --- DUA TOMBOL UPLOAD AGAR TIDAK ERROR FILE NOT FOUND ---
st.info("💡 Langkah: Upload file Template terlebih dahulu, lalu upload file Event table.")
col1, col2 = st.columns(2)
with col1:
    uploaded_template = st.file_uploader("1. Upload File TEMPLATE (Dashboard Asli)", type=["xlsx"])
with col2:
    uploaded_file = st.file_uploader("2. Upload File DATA (Event_table)", type=["xlsx", "xls"])

# Mesin baru berjalan jika KEDUA file sudah di-upload
if uploaded_template is not None and uploaded_file is not None:
    try:
        with st.spinner("Memproses data dan menyalin warna template..."):
            # 1. READ & CLEAN DATA
            df_raw = pd.read_excel(uploaded_file, skiprows=3)

            headers = [f"Col_{i}" if pd.isna(val) else str(val).strip() for i, val in enumerate(df_raw.iloc[0])]
            df_raw.columns = headers
            df_data = df_raw.iloc[1:].copy()

            std_cols = ['File', 'Fiber', 'Wavelength', 'Loss_dB', 'Length_km', 'Attenuation_dB_km']
            rename_dict = {df_data.columns[i+1]: std_cols[i] for i in range(6)}
            df_data.rename(columns=rename_dict, inplace=True)

            for col in ['Loss_dB', 'Length_km', 'Attenuation_dB_km']:
                df_data[col] = pd.to_numeric(df_data[col], errors='coerce')

            def extract_core(filename):
                match = re.search(r'(\d+)', str(filename)) if not pd.isna(filename) else None
                return int(match.group(1)) if match else None

            df_data['Core'] = df_data['File'].apply(extract_core)
            df_data = df_data.dropna(subset=['Core', 'Length_km', 'Loss_dB']).copy()
            df_data['Core'] = df_data['Core'].astype(int)
            df_data['Tube'] = df_data['Core'].apply(lambda x: f"Tube {math.ceil(x/12)}")

            df_clean = df_data[['Tube', 'Core', 'File', 'Fiber', 'Wavelength', 'Loss_dB', 'Length_km', 'Attenuation_dB_km']].copy()

            # 2. CALCULATIONS
            REF_LENGTH = 59.67
            df_detail = df_clean.sort_values(by='Core').reset_index(drop=True)
            df_detail['Length_%_of_Route'] = (df_detail['Length_km'] / REF_LENGTH) * 100
            df_detail['Status'] = df_detail['Length_%_of_Route'].apply(lambda pct: "NORMAL" if pct >= 95 else ("WARNING" if pct >= 80 else "CRITICAL"))

            # Tube Analysis
            tube_stats = df_detail.groupby('Tube').agg(
                Core_Count=('Core', 'count'),
                Avg_Length_km=('Length_km', 'mean'),
                Min_Length_km=('Length_km', 'min'),
                Max_Length_km=('Length_km', 'max'),
                Avg_Loss_dB=('Loss_dB', 'mean'),
                Max_Loss_dB=('Loss_dB', 'max'),
                Avg_Attenuation_dB_km=('Attenuation_dB_km', 'mean'),
                Max_Attenuation_dB_km=('Attenuation_dB_km', 'max')
            ).reset_index()

            tube_stats['Delta_vs_Route_km'] = tube_stats['Avg_Length_km'] - REF_LENGTH
            tube_stats['Completeness_%'] = (tube_stats['Core_Count'] / 12) * 100
            tube_stats['Status'] = tube_stats.apply(
                lambda r: "REVIEW - data terbatas" if r['Core_Count'] <= 2 else ("PRIORITAS PERBAIKAN" if r['Avg_Length_km'] < (0.95 * REF_LENGTH) else "NORMAL"), axis=1)
            
            # Length Group
            df_detail_sorted = df_detail.sort_values(by='Length_km').copy()
            groups = []
            current_group = 1
            prev_len = None
            for idx, row in df_detail_sorted.iterrows():
                if prev_len is not None and (row['Length_km'] - prev_len > 0.15):
                    current_group += 1
                groups.append(current_group)
                prev_len = row['Length_km']
            df_detail_sorted['Length_Group'] = groups
            
            df_detail = df_detail.merge(df_detail_sorted[['Core', 'Length_Group']], on='Core', how='left')

            length_summary = df_detail_sorted.groupby('Length_Group').agg(
                Core_Count=('Core', 'count'),
                Avg_Length_km=('Length_km', 'mean'),
                Min_Length_km=('Length_km', 'min'),
                Max_Length_km=('Length_km', 'max'),
                Avg_Loss_dB=('Loss_dB', 'mean'),
                Avg_Attenuation_dB_km=('Attenuation_dB_km', 'mean')
            ).reset_index()
            length_summary.insert(1, 'Panjang Kelompok', length_summary['Avg_Length_km'].apply(lambda x: f"{x:.2f} km (±0.15 km)"))

            # Ranking Masalah
            ranking = tube_stats[['Tube']].copy()
            tube_counts = df_detail.groupby('Tube')['Status'].value_counts().unstack(fill_value=0)
            for col in ['CRITICAL', 'WARNING']:
                if col not in tube_counts.columns: tube_counts[col] = 0
                
            ranking = ranking.merge(tube_counts[['CRITICAL', 'WARNING']].reset_index(), on='Tube', how='left')
            ranking = ranking.merge(tube_stats[['Tube', 'Avg_Length_km', 'Avg_Loss_dB', 'Avg_Attenuation_dB_km', 'Core_Count']], on='Tube')
            ranking['Score'] = (0.60 * ((REF_LENGTH - ranking['Avg_Length_km']) / REF_LENGTH * 100).clip(lower=0) +
                                0.25 * (ranking['Avg_Loss_dB']) + 0.15 * (ranking['Avg_Attenuation_dB_km'] * 100))
            ranking = ranking.sort_values(by='Score', ascending=False)
            ranking.insert(0, 'Rank', range(1, len(ranking) + 1))
            ranking['Status'] = ranking['Score'].apply(lambda x: "CRITICAL" if x > 50 else ("WARNING" if x > 20 else "NORMAL"))

            # 3. MENGISI KE TEMPLATE EXCEL ASLI DARI FILE UPLOADED
            wb = load_workbook(uploaded_template)
            
            def update_table(ws, df, start_row, start_col=1):
                style_row = start_row
                for r_idx, row in enumerate(df.itertuples(index=False), start_row):
                    for c_idx, val in enumerate(row, start_col):
                        cell = ws.cell(row=r_idx, column=c_idx, value=val)
                        ref_cell = ws.cell(row=style_row, column=c_idx)
                        if ref_cell.has_style: 
                            cell.font = copy.copy(ref_cell.font)
                            cell.border = copy.copy(ref_cell.border)
                            cell.fill = copy.copy(ref_cell.fill)
                            cell.number_format = copy.copy(ref_cell.number_format)
                            cell.alignment = copy.copy(ref_cell.alignment)
                
                # Hapus data ghosting di bawahnya jika tabel baru lebih pendek dari template asli
                end_row = start_row + len(df)
                for r in range(end_row, end_row + 50):
                    for c in range(start_col, start_col + len(df.columns)):
                        ws.cell(row=r, column=c).value = None

            # Update Value Visual Dashboard & Dashboard
            ws_vis = wb['Visual_Dashboard']
            if ws_vis['A5'].value: ws_vis['A5'] = f"TOTAL CORE\n{len(df_detail)}"
            if ws_vis['D5'].value: ws_vis['D5'] = f"AVG PANJANG\n{df_detail['Length_km'].mean():.2f}"
            if ws_vis['G5'].value: ws_vis['G5'] = f"AVG LOSS\n{df_detail['Loss_dB'].mean():.2f}"
            
            for row in ws_vis.iter_rows(min_row=3, max_row=7):
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and 'PRIORITAS' in cell.value:
                        cell.value = f"PRIORITAS\n{len(df_detail[df_detail['Status'] != 'NORMAL'])}"

            ws_dash = wb['Dashboard']
            if ws_dash['A5'].value: ws_dash['A5'] = f"TOTAL CORE\n{len(df_detail)}"
            if ws_dash['D5'].value: ws_dash['D5'] = f"JUMLAH TUBE\n{len(tube_stats)}"
            if ws_dash['G5'].value: ws_dash['G5'] = f"PANJANG REFERENSI\n{REF_LENGTH}"
            
            # Injeksi Tabel ke seluruh sheet
            update_table(ws_dash, tube_stats, start_row=10)
            update_table(wb['Core_Detail'], df_detail, start_row=2)
            update_table(wb['Tube_Analysis'], tube_stats, start_row=2)
            update_table(wb['Length_Group'], length_summary, start_row=2)
            update_table(wb['Data_Clean'], df_clean, start_row=2)
            
            cols_order = ['Rank', 'Tube', 'Core_Count', 'Avg_Length_km', 'Avg_Loss_dB', 'Avg_Attenuation_dB_km', 'CRITICAL', 'WARNING', 'Score', 'Status']
            update_table(wb['Ranking_Masalah'], ranking[cols_order], start_row=7)

            # Simpan output
            output = BytesIO()
            wb.save(output)
            output.seek(0)

        st.success("✅ Dashboard berhasil dibuat dengan format identik!")
        st.download_button(
            label="📥 Download Excel Hasil",
            data=output.getvalue(),
            file_name="Dashboard_OTDR_Generator_Result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
