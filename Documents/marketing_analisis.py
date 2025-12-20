import pandas as pd
import numpy as np
import os
import glob
from datetime import datetime

class SeniorMarketingAnalyst:
    def __init__(self):
        # Fitur Scan Dokumen tetap dipertahankan
        self.selected_file = self._scan_and_select_file()
        if not self.selected_file:
            return
        self.report_name = f"MARKETING_ANALISIS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    def _scan_and_select_file(self):
        print("=== MARKETING ANALISIS ===")
        files = glob.glob('*.xlsx')
        if not files: return None
        for i, f in enumerate(files, 1):
            print(f"{i}. {f}")
        try:
            return files[int(input("\nPilih nomor file: ")) - 1]
        except: return None

    def run_analysis(self):
        if not self.selected_file: return
        
        # Load Data
        df = pd.read_excel(self.selected_file, sheet_name='Penjualan')
        df['Kategori'] = df['Nama Pelanggan'].str.split().str[0]

        # --- LOGIKA ANALISIS ---
        prof = df.groupby(['Kategori', 'Nama Barang'])['Netto'].sum().reset_index()
        prof = prof.sort_values(['Kategori', 'Netto'], ascending=[True, False]).groupby('Kategori').head(1)

        penetration = pd.pivot_table(df, values='Qty', index='Nama Barang', columns='Kategori', aggfunc='sum', fill_value=0)

        roi = df.groupby('Kategori').agg({'Diskon': 'sum', 'Netto': 'sum', 'No. Faktur': 'count'}).reset_index()
        roi['Rasio_Diskon'] = (roi['Diskon'] / (roi['Netto'] + roi['Diskon'])) * 100

        loyalty = df.groupby('Nama Pelanggan').agg({'No. Faktur': 'count', 'Netto': 'sum', 'Tanggal': lambda x: (datetime.now() - x.max()).days}).rename(columns={'No. Faktur': 'Frekuensi', 'Netto': 'Nilai_Belanja', 'Tanggal': 'Hari_Sejak_Order_Terakhir'})

        # Palette Warna Profesional
        colors = ['#4E79A7', '#F28E2B', '#E15759', '#76B7B2', '#59A14F', '#EDC948', '#B07AA1', '#FF9DA7', '#9C755F', '#BAB0AC'] * 3

        print(f"[*] Menyusun laporan dengan Highlighting Per Produk...")
        with pd.ExcelWriter(self.report_name, engine='xlsxwriter') as writer:
            prof.to_excel(writer, sheet_name='1_Profiling_Pasar', index=False)
            penetration.to_excel(writer, sheet_name='2_Penetrasi_Produk')
            roi.to_excel(writer, sheet_name='3_ROI_Promosi', index=False)
            loyalty.to_excel(writer, sheet_name='4_Analisis_Loyalitas')

            workbook = writer.book
            money_fmt = workbook.add_format({'num_format': '#,##0', 'font_name': 'Arial'})
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
            
            # Format baru untuk Penjualan Tertinggi Per Baris (Produk)
            best_sell_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True})

            for sheet_name in writer.sheets:
                ws = writer.sheets[sheet_name]
                df_temp = prof if 'Profiling' in sheet_name else (penetration.reset_index() if 'Penetrasi' in sheet_name else (roi if 'ROI' in sheet_name else loyalty.reset_index()))

                # 1. AUTO-FIT & HEADER
                for i, col in enumerate(df_temp.columns):
                    max_len = max(df_temp[col].astype(str).map(len).max(), len(str(col))) + 3
                    ws.set_column(i, i, max_len, money_fmt if i > 0 else None)
                    ws.write(0, i, col, header_fmt)

                # A. Penetrasi Produk: Warnai Tertinggi PER PRODUK (Row-based Highlighting)
                if 'Penetrasi' in sheet_name:
                    num_rows = len(penetration)
                    num_cols = len(penetration.columns)
                    
                    # Loop setiap baris (setiap nama produk)
                    for row_idx in range(1, num_rows + 1):
                        # Terapkan conditional formatting hanya pada baris tersebut (kolom B sampai akhir)
                        ws.conditional_format(row_idx, 1, row_idx, num_cols, {
                            'type':     'top',
                            'value':    1,
                            'format':   best_sell_fmt
                        })

                # B. Profiling Pasar: Diagram Batang
                if 'Profiling' in sheet_name:
                    chart_prof = workbook.add_chart({'type': 'column'})
                    chart_prof.add_series({
                        'name': 'Hero Product Revenue',
                        'categories': [sheet_name, 1, 0, len(prof), 0],
                        'values':     [sheet_name, 1, 2, len(prof), 2],
                        'points': [{'fill': {'color': colors[i]}} for i in range(len(prof))]
                    })
                    ws.insert_chart('L2', chart_prof)

                # C. ROI Promosi: Pie Chart
                if 'ROI' in sheet_name:
                    chart_roi = workbook.add_chart({'type': 'pie'})
                    chart_roi.add_series({
                        'categories': [sheet_name, 1, 0, len(roi), 0],
                        'values':     [sheet_name, 1, 1, len(roi), 1],
                        'data_labels': {'percentage': True},
                    })
                    ws.insert_chart('G2', chart_roi)

        print(f"\n[SUKSES] Laporan selesai dibuat.")
        print(f"File: {os.path.abspath(self.report_name)}")

if __name__ == "__main__":
    analyst = SeniorMarketingAnalyst()
    analyst.run_analysis()