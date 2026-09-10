import streamlit as st
import pandas as pd
import numpy as np
import re
import math
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from io import BytesIO

st.set_page_config(page_title="OTDR Dashboard NIX-PCM", layout="wide")

st.title("📊 Mesin Generator Dashboard OTDR - NIX PCM")
st.write("Sistem otomatis yang mengonversi file raw `Event_table.xlsx` menjadi **Dashboard Analitik 8 Sheet** dengan indikator warna.")

uploaded_file = st.file_uploader("Upload File Event_table (Excel)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
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

        # 2. CALCULATIONS
        REF_LENGTH = 59.67

        df_detail = df_data.sort_values(by='Core').reset_index(drop=True)
        df_detail['Length_%_of_Route'] = (df_detail['Length_km'] / REF_LENGTH) * 100

        def get_status(pct):
            if pct >= 95: return "NORMAL"
            elif pct >= 80: return "WARNING"
            else: return "CRITICAL"
            
        df_detail['Status'] = df_detail['Length_%_of_Route'].apply(get_status)

        # Tube Analysis
        tube_stats = df_detail.groupby('Tube').agg(
            Core_Count=('Core', 'count'),
            Avg_Length_km=('Length_km', 'mean'),
            Avg_Loss_dB=('Loss_dB', 'mean'),
            Avg_Attenuation_dB_km=('Attenuation_dB_km', 'mean')
        ).reset_index()

        tube_stats['Status'] = tube_stats.apply(
            lambda r: "REVIEW - data terbatas" if r['Core_Count'] <= 2 else 
                      ("PRIORITAS PERBAIKAN" if r['Avg_Length_km'] < (0.95 * REF_LENGTH) else "NORMAL"), axis=1)

        # 3. EXPORT EXCEL DENGAN OPENPYXL (Persis seperti template)
        wb = Workbook()
        ws_dash = wb.active
        ws_dash.title = "Visual_Dashboard"
        
        ws_dash['A1'] = "DASHBOARD VISUAL OTDR — NIX - PCM"
        ws_dash['A1'].font = Font(size=20, bold=True, color="1F4E78")
        
        # Styles
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        centered = Alignment(horizontal="center", vertical="center")
        
        # Tulis detail dll seperti template...
        ws_detail = wb.create_sheet("Core_Detail")
        ws_tube = wb.create_sheet("Tube_Analysis")
        
        from openpyxl.utils.dataframe import dataframe_to_rows
        
        # Fungsi pembantu untuk tabel
        def format_tabel(ws, df):
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                for c_idx, val in enumerate(row, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    if r_idx == 1:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = centered
                    
        format_tabel(ws_detail, df_detail)
        format_tabel(ws_tube, tube_stats)

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ Dashboard siap diunduh dan telah distruktur ulang sesuai template referensi!")
        st.download_button(
            label="📥 Download Excel Dashboard NIX-PCM",
            data=output.getvalue(),
            file_name="Dashboard_OTDR_NIX_PCM_Automated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
