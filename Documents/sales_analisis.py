import pandas as pd
import numpy as np
import glob
import os
import xlsxwriter
from datetime import datetime, timedelta

# --- FUNGSI BANTUAN ---

def pilih_file_target_realistic():
    """Mencari file Data V9 terbaru"""
    list_files = glob.glob('Data_V9_Dummy*.xlsx')
    if not list_files:
        print("ERROR: Tidak ditemukan file Data V9. Jalankan generator V9 terlebih dahulu.")
        return None
    latest_file = max(list_files, key=os.path.getctime)
    print(f"--- Menganalisa File: {latest_file} ---")
    return latest_file

def segmentasi_pasar(nama):
    nama = str(nama).lower()
    if 'bakso' in nama: return 'Kuliner - Bakso'
    if 'sate' in nama: return 'Kuliner - Sate'
    if 'soto' in nama: return 'Kuliner - Soto'
    if 'mie' in nama: return 'Kuliner - Mie'
    if 'warung' in nama: return 'Retail - Warung'
    if 'toko' in nama: return 'Retail - Toko'
    if 'ud' in nama or 'cv' in nama or 'agen' in nama: return 'Wholesale/Agen'
    if 'catering' in nama or 'rumah makan' in nama or 'resto' in nama: return 'Horeka'
    return 'Lainnya'

def ekstrak_jalan(alamat):
    try:
        return alamat.split("No.")[0].strip() if "No." in alamat else alamat.split(",")[0].strip()
    except: return "Unknown"

def hitung_umur_faktur(tgl_faktur):
    if pd.isna(tgl_faktur): return 0
    today = datetime.now() 
    delta = (today - tgl_faktur).days
    return delta

def bucket_umur(hari):
    if hari <= 0: return "-30 Hari - 0 Hari (Belum JT)"
    elif hari <= 31: return "1 Hari - 31 Hari"
    elif hari <= 60: return "32 Hari - 60 Hari"
    else: return "> 60 Hari (Macet)"

# --- MAIN LOGIC ---

def run_analyst_v3():
    input_file = pilih_file_target_realistic()
    if not input_file: return

    print("1. Membaca & Membersihkan Data...")
    try:
        df_jual = pd.read_excel(input_file, sheet_name='Penjualan')
        df_bayar = pd.read_excel(input_file, sheet_name='Pembayaran')
        df_saldo = pd.read_excel(input_file, sheet_name='Saldo Awal')
        df_target_raw = pd.read_excel(input_file, sheet_name='Target Sales', header=[0,1], index_col=0)
    except Exception as e:
        print(f"Error load data: {e}")
        return

    # =========================================================================
    # A. DASHBOARD & FORECASTING
    # =========================================================================
    print("2. Menyusun Dashboard & Forecasting...")
    
    df_jual['Tanggal'] = pd.to_datetime(df_jual['Tanggal'])
    daily_sales = df_jual.groupby('Tanggal')['Netto'].sum().reset_index()
    
    total_omzet = df_jual['Netto'].sum()
    total_transaksi = len(df_jual)
    avg_trx = total_omzet / total_transaksi if total_transaksi > 0 else 0
    
    daily_sales['Day_Num'] = (daily_sales['Tanggal'] - daily_sales['Tanggal'].min()).dt.days
    X = daily_sales['Day_Num'].values
    y = daily_sales['Netto'].values
    
    df_forecast = pd.DataFrame()
    trend_desc = "Netral"
    
    if len(X) > 1:
        m, c = np.polyfit(X, y, 1)
        trend_desc = "POSITIF (NAIK)" if m > 0 else "NEGATIF (TURUN)"
        last_real_date = daily_sales['Tanggal'].max()
        future_dates = [last_real_date + timedelta(days=i) for i in range(1, 31)]
        future_vals = [(m * (X[-1] + i)) + c for i in range(1, 31)]
        df_forecast = pd.DataFrame({'Tanggal': future_dates, 'Prediksi': future_vals})

    # =========================================================================
    # B. ANALISIS PASAR & JALAN
    # =========================================================================
    print("3. Analisis Pasar (Segmentasi & Jalan)...")
    
    df_jual['Segmen'] = df_jual['Nama Pelanggan'].apply(segmentasi_pasar)
    segment_value = df_jual.groupby('Segmen')['Netto'].sum().reset_index().sort_values('Netto', ascending=False)
    
    top_segment = segment_value.iloc[0]['Segmen']
    top_contribution = (segment_value.iloc[0]['Netto'] / total_omzet) * 100
    
    rekomendasi_txt = (
        f"REKOMENDASI:\n"
        f"1. Fokus akuisisi pelanggan tipe '{top_segment}' ({top_contribution:.1f}%).\n"
        f"2. Perkuat penetrasi di jalan dengan omzet tertinggi."
    )

    # --- RESTORE: ANALISIS JALAN ---
    df_jual['Jalan'] = df_jual['Alamat'].apply(ekstrak_jalan)
    geo_sales = df_jual.groupby('Jalan')['Netto'].sum().reset_index().sort_values('Netto', ascending=False).head(10)

    # =========================================================================
    # C. EVALUASI TIM (INSENTIF & CHART TARGET)
    # =========================================================================
    print("4. Evaluasi Tim Sales...")
    
    if 'GRAND TOTAL' in df_target_raw.index: df_target_raw = df_target_raw.drop('GRAND TOTAL')
    
    df_jual['Diskon Pct'] = (df_jual['Diskon'] / df_jual['Total']).fillna(0)
    avg_disc_per_sales = df_jual.groupby('Nama Sales')['Diskon Pct'].mean()
    
    sales_kpi_data = []
    sales_names = df_target_raw.columns.get_level_values(0).unique()
    
    for sales in sales_names:
        try:
            tgt_val = df_target_raw[sales]['Target Value'].sum()
            act_val = df_jual[df_jual['Nama Sales'] == sales]['Netto'].sum()
            ach_pct = act_val / tgt_val if tgt_val > 0 else 0
            
            insentif_rp = 0
            rate_info = "0%"
            
            if ach_pct >= 1.00:
                insentif_rp = act_val * 0.025
                rate_info = "2.5%"
            elif ach_pct >= 0.85:
                insentif_rp = act_val * 0.015
                rate_info = "1.5%"
            elif ach_pct >= 0.70:
                insentif_rp = act_val * 0.005
                rate_info = "0.5%"
            
            avg_d = avg_disc_per_sales.get(sales, 0)
            
            sales_kpi_data.append({
                'Nama Sales': sales,
                'Target': tgt_val,
                'Realisasi': act_val,
                'Ach %': ach_pct,
                'Rate Bonus': rate_info,
                'Total Bonus (Rp)': insentif_rp,
                'Avg Diskon': avg_d
            })
        except: continue
        
    df_kpi_team = pd.DataFrame(sales_kpi_data)

    # =========================================================================
    # D. ANALISA PIUTANG (DETAIL)
    # =========================================================================
    print("5. Kalkulasi Detail Piutang...")
    
    df_old = df_saldo[['No. Faktur Lama', 'Tanggal Faktur', 'Nama Pelanggan', 'Nama Sales', 'Sisa Piutang']].copy()
    df_old.columns = ['No. Faktur', 'Tanggal', 'Nama Pelanggan', 'Nama Sales', 'Tagihan Awal']
    
    df_new = df_jual[['No. Faktur', 'Tanggal', 'Nama Pelanggan', 'Nama Sales', 'Netto']].copy()
    df_new.columns = ['No. Faktur', 'Tanggal', 'Nama Pelanggan', 'Nama Sales', 'Tagihan Awal']
    
    df_ar_master = pd.concat([df_old, df_new], ignore_index=True)
    pay_per_inv = df_bayar.groupby('No. Faktur')['Jumlah Bayar'].sum().reset_index()
    
    df_ar_final = pd.merge(df_ar_master, pay_per_inv, on='No. Faktur', how='left').fillna(0)
    df_ar_final['Sisa Piutang'] = df_ar_final['Tagihan Awal'] - df_ar_final['Jumlah Bayar']
    df_ar_final = df_ar_final[df_ar_final['Sisa Piutang'] > 100].copy()
    
    df_ar_final['Tanggal'] = pd.to_datetime(df_ar_final['Tanggal'])
    df_ar_final['Umur Hari'] = df_ar_final['Tanggal'].apply(hitung_umur_faktur)
    df_ar_final['Kategori'] = df_ar_final['Umur Hari'].apply(bucket_umur)
    
    pivot_bucket = df_ar_final.pivot_table(index='Nama Sales', columns='Kategori', values='Sisa Piutang', aggfunc='sum', fill_value=0)
    buckets = ["-30 Hari - 0 Hari (Belum JT)", "1 Hari - 31 Hari", "32 Hari - 60 Hari", "> 60 Hari (Macet)"]
    for b in buckets:
        if b not in pivot_bucket.columns: pivot_bucket[b] = 0
    pivot_bucket = pivot_bucket[buckets]
    pivot_bucket['TOTAL'] = pivot_bucket.sum(axis=1)
    
    top_10_ar = df_ar_final.sort_values('Sisa Piutang', ascending=False).head(10)

    # =========================================================================
    # WRITING TO EXCEL
    # =========================================================================
    output_file = f"Laporan_Analisis_Sales_V3_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    print(f"6. Menyimpan Laporan: {output_file}...")
    
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    wb = writer.book
    
    # Styles
    fmt_head = wb.add_format({'bold':True, 'bg_color':'#2c3e50', 'font_color':'white', 'border':1, 'align':'center'})
    fmt_curr = wb.add_format({'num_format':'Rp #,##0', 'border':1})
    fmt_pct = wb.add_format({'num_format':'0.0%', 'border':1, 'align':'center'})
    fmt_num = wb.add_format({'num_format':'#,##0', 'border':1})
    fmt_date = wb.add_format({'num_format':'dd-mm-yyyy', 'border':1, 'align':'center'})
    fmt_txt_wrap = wb.add_format({'text_wrap': True, 'valign': 'top', 'border': 1, 'bg_color': '#FFFFE0'})

    # --- SHEET 1: DASHBOARD ---
    ws1 = wb.add_worksheet('Dashboard')
    ws1.set_column('A:B', 22)
    ws1.set_column('D:E', 22)
    
    ws1.write('A1', 'KPI DASHBOARD', fmt_head)
    ws1.write('A2', 'Total Omzet', fmt_head)
    ws1.write('B2', total_omzet, fmt_curr)
    ws1.write('A3', 'Total Transaksi', fmt_head)
    ws1.write('B3', total_transaksi, fmt_num)
    
    ws1.write('D1', 'REALISASI & FORECAST', fmt_head)
    ws1.write('D2', 'Tanggal', fmt_head)
    ws1.write('E2', 'Realisasi', fmt_head)
    ws1.write('F2', 'Forecast', fmt_head)
    
    row_idx = 2
    for i, r in daily_sales.iterrows():
        ws1.write(row_idx, 3, r['Tanggal'], fmt_date)
        ws1.write(row_idx, 4, r['Netto'], fmt_curr)
        row_idx += 1
    
    for i, r in df_forecast.iterrows():
        ws1.write(row_idx, 3, r['Tanggal'], fmt_date)
        ws1.write(row_idx, 5, r['Prediksi'], fmt_curr)
        row_idx += 1
        
    chart1 = wb.add_chart({'type': 'line'})
    chart1.add_series({'name': 'Realisasi', 'categories': ['Dashboard', 2, 3, row_idx-1, 3], 'values': ['Dashboard', 2, 4, row_idx-1, 4], 'line': {'color': 'blue'}})
    chart1.add_series({'name': 'Forecast', 'categories': ['Dashboard', 2, 3, row_idx-1, 3], 'values': ['Dashboard', 2, 5, row_idx-1, 5], 'line': {'color': 'red', 'dash_type': 'dash'}})
    ws1.insert_chart('H2', chart1, {'x_scale': 1.8, 'y_scale': 1.2})

    # --- SHEET 2: ANALISIS PASAR ---
    ws2 = wb.add_worksheet('Analisis Pasar')
    ws2.set_column('A:B', 25)
    ws2.set_column('E:F', 22) # Kolom untuk Jalan
    
    ws2.write('A1', 'SEGMENTASI PELANGGAN', fmt_head)
    ws2.write('A2', 'Segmen', fmt_head)
    ws2.write('B2', 'Omzet', fmt_head)
    
    r_seg = 2
    for i, row in segment_value.iterrows():
        ws2.write(r_seg, 0, row['Segmen'], fmt_num)
        ws2.write(r_seg, 1, row['Netto'], fmt_curr)
        r_seg += 1
        
    chart2 = wb.add_chart({'type': 'pie'})
    chart2.add_series({
        'name': 'Market Share',
        'categories': ['Analisis Pasar', 2, 0, r_seg-1, 0],
        'values':     ['Analisis Pasar', 2, 1, r_seg-1, 1],
        'data_labels': {'percentage': True, 'category_name': True, 'leader_lines': True, 'separator': '\n'}
    })
    ws2.insert_chart('A15', chart2)
    
    # --- RESTORED: TABEL JALAN ---
    ws2.write('E1', 'TOP 10 AREA (JALAN)', fmt_head)
    ws2.write('E2', 'Nama Jalan', fmt_head)
    ws2.write('F2', 'Total Omzet', fmt_head)
    
    r_geo = 2
    for i, row in geo_sales.iterrows():
        ws2.write(r_geo, 4, row['Jalan'], fmt_num)
        ws2.write(r_geo, 5, row['Netto'], fmt_curr)
        r_geo += 1
        
    ws2.merge_range('E15:H18', rekomendasi_txt, fmt_txt_wrap)

    # --- SHEET 3: EVALUASI TIM ---
    ws3 = wb.add_worksheet('Evaluasi Tim')
    ws3.set_column('A:H', 18)
    
    headers = ['Salesman', 'Target', 'Realisasi', 'Ach %', 'Rate Bonus', 'Total Bonus (Rp)', 'Avg Diskon']
    for i, h in enumerate(headers): ws3.write(0, i, h, fmt_head)
        
    for i, row in df_kpi_team.iterrows():
        ws3.write(i+1, 0, row['Nama Sales'], fmt_num)
        ws3.write(i+1, 1, row['Target'], fmt_curr)
        ws3.write(i+1, 2, row['Realisasi'], fmt_curr)
        ws3.write(i+1, 3, row['Ach %'], fmt_pct)
        ws3.write(i+1, 4, row['Rate Bonus'], fmt_num)
        ws3.write(i+1, 5, row['Total Bonus (Rp)'], fmt_curr)
        ws3.write(i+1, 6, row['Avg Diskon'], fmt_pct)
    
    # --- RESTORED: CHART TARGET VS REALISASI ---
    chart3 = wb.add_chart({'type': 'column'})
    chart3.add_series({
        'name': 'Target',
        'categories': ['Evaluasi Tim', 1, 0, len(df_kpi_team), 0],
        'values':     ['Evaluasi Tim', 1, 1, len(df_kpi_team), 1],
        'fill':       {'color': '#D3D3D3'} # Abu-abu
    })
    chart3.add_series({
        'name': 'Realisasi',
        'categories': ['Evaluasi Tim', 1, 0, len(df_kpi_team), 0],
        'values':     ['Evaluasi Tim', 1, 2, len(df_kpi_team), 2],
        'fill':       {'color': '#2c3e50'} # Biru Tua
    })
    ws3.insert_chart('A12', chart3)

    # --- SHEET 4: PIUTANG (AR) ---
    ws4 = wb.add_worksheet('Piutang (AR)')
    ws4.set_column('A:A', 20)
    ws4.set_column('B:E', 18)
    ws4.set_column('F:F', 20)
    
    ws4.write('A1', 'REKAP PIUTANG PER SALES', fmt_head)
    ws4.write('A2', 'Nama Sales', fmt_head)
    ws4.write('B2', 'Belum JT (-30 s/d 0)', fmt_head)
    ws4.write('C2', '1 s/d 31 Hari', fmt_head)
    ws4.write('D2', '32 s/d 60 Hari', fmt_head)
    ws4.write('E2', '> 60 Hari (Macet)', fmt_head)
    ws4.write('F2', 'TOTAL PIUTANG', fmt_head)
    
    r_piv = 2
    for sales, row in pivot_bucket.iterrows():
        ws4.write(r_piv, 0, sales, fmt_num)
        ws4.write(r_piv, 1, row['-30 Hari - 0 Hari (Belum JT)'], fmt_curr)
        ws4.write(r_piv, 2, row['1 Hari - 31 Hari'], fmt_curr)
        ws4.write(r_piv, 3, row['32 Hari - 60 Hari'], fmt_curr)
        ws4.write(r_piv, 4, row['> 60 Hari (Macet)'], fmt_curr)
        ws4.write(r_piv, 5, row['TOTAL'], fmt_curr)
        r_piv += 1
        
    total_buckets = pivot_bucket[buckets].sum()
    ws4.write('H2', 'Kategori', fmt_head)
    ws4.write('I2', 'Total Nilai', fmt_head)
    r_h = 2
    for cat, val in total_buckets.items():
        ws4.write(r_h, 7, cat, fmt_num)
        ws4.write(r_h, 8, val, fmt_curr)
        r_h += 1
        
    chart_ar = wb.add_chart({'type': 'pie'})
    chart_ar.add_series({
        'name': 'Status Piutang',
        'categories': ['Piutang (AR)', 2, 7, r_h-1, 7],
        'values':     ['Piutang (AR)', 2, 8, r_h-1, 8],
        'data_labels': {'percentage': True, 'category_name': True, 'leader_lines': True, 'separator': '\n'}
    })
    ws4.insert_chart('H10', chart_ar)
    
    ws4.write('A15', 'TOP 10 FAKTUR PIUTANG TERTINGGI', fmt_head)
    ws4.write('A16', 'No Faktur', fmt_head)
    ws4.write('B16', 'Tanggal', fmt_head)
    ws4.write('C16', 'Customer', fmt_head)
    ws4.write('D16', 'Sales', fmt_head)
    ws4.write('E16', 'Sisa Piutang', fmt_head)
    ws4.write('F16', 'Umur (Hari)', fmt_head)
    
    r_top = 16
    for i, row in top_10_ar.iterrows():
        ws4.write(r_top, 0, row['No. Faktur'], fmt_num)
        ws4.write(r_top, 1, row['Tanggal'], fmt_date)
        ws4.write(r_top, 2, row['Nama Pelanggan'], fmt_num)
        ws4.write(r_top, 3, row['Nama Sales'], fmt_num)
        ws4.write(r_top, 4, row['Sisa Piutang'], fmt_curr)
        ws4.write(r_top, 5, row['Umur Hari'], fmt_num)
        r_top += 1

    start_row = r_top + 5
    ws4.write(start_row, 0, 'RINCIAN LENGKAP SEMUA PIUTANG', fmt_head)
    cols_detail = ['No. Faktur', 'Tanggal', 'Nama Pelanggan', 'Nama Sales', 'Tagihan Awal', 'Jumlah Bayar', 'Sisa Piutang', 'Kategori']
    for i, c in enumerate(cols_detail): ws4.write(start_row+1, i, c, fmt_head)
        
    r_det = start_row + 2
    for i, row in df_ar_final.iterrows():
        ws4.write(r_det, 0, row['No. Faktur'], fmt_num)
        ws4.write(r_det, 1, row['Tanggal'], fmt_date)
        ws4.write(r_det, 2, row['Nama Pelanggan'], fmt_num)
        ws4.write(r_det, 3, row['Nama Sales'], fmt_num)
        ws4.write(r_det, 4, row['Tagihan Awal'], fmt_curr)
        ws4.write(r_det, 5, row['Jumlah Bayar'], fmt_curr)
        ws4.write(r_det, 6, row['Sisa Piutang'], fmt_curr)
        ws4.write(r_det, 7, row['Kategori'], fmt_num)
        r_det += 1

    writer.close()
    print("==========================================")
    print(f"LAPORAN SALES V3 SELESAI. File: {output_file}")
    print("==========================================")

if __name__ == "__main__":
    run_analyst_v3()