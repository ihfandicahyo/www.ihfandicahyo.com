import pandas as pd
import glob
import os
import xlsxwriter
from datetime import datetime

def pilih_file_interaktif():
    """
    Scan folder, tampilkan daftar file Data V9, dan minta user memilih.
    """
    # Pola file yang dicari
    pola_file = 'Data_V9_Dummy*.xlsx'
    
    # Ambil semua file yang cocok
    list_files = glob.glob(pola_file)
    
    if not list_files:
        print(f"ERROR: Tidak ditemukan file dengan pola '{pola_file}'.")
        print("Pastikan Anda sudah menjalankan generator data (V5) sebelumnya.")
        return None

    # Urutkan file berdasarkan waktu pembuatan (Terbaru paling atas)
    list_files.sort(key=os.path.getmtime, reverse=True)

    print("\n========================================")
    print("   PILIH FILE UNTUK DIANALISIS")
    print("========================================")
    print(f"Ditemukan {len(list_files)} file data:\n")

    for i, file_path in enumerate(list_files):
        # Dapatkan waktu modifikasi file agar user tahu mana yang baru
        timestamp = os.path.getmtime(file_path)
        waktu_file = datetime.fromtimestamp(timestamp).strftime('%d-%b-%Y %H:%M:%S')
        
        # Tampilkan menu
        print(f"  [{i+1}] {file_path}")
        print(f"      L--> Dibuat pada: {waktu_file}")

    print("\n  [0] Batal / Keluar")
    print("========================================")

    # Loop input user sampai benar
    while True:
        try:
            pilihan = input("Masukkan nomor file (contoh: 1): ")
            
            # Konversi ke integer
            idx = int(pilihan)

            if idx == 0:
                print("Analisa dibatalkan.")
                return None
            
            if 1 <= idx <= len(list_files):
                # User memilih nomor yang valid
                file_terpilih = list_files[idx - 1]
                print(f"\n>>> Memproses File: {file_terpilih} ...")
                return file_terpilih
            else:
                print(f"Nomor tidak valid. Masukkan angka 1 sampai {len(list_files)}.")

        except ValueError:
            print("Input salah. Harap masukkan ANGKA saja.")

def generate_analyst_report():
    # 1. PILIH FILE (Metode Baru)
    input_file = pilih_file_interaktif()
    
    # Jika user batal atau tidak ada file, berhenti
    if not input_file: 
        return

    # 2. LOAD DATA
    try:
        print("Membaca data Excel...")
        df_jual = pd.read_excel(input_file, sheet_name='Penjualan')
        df_bayar = pd.read_excel(input_file, sheet_name='Pembayaran')
        df_saldo = pd.read_excel(input_file, sheet_name='Saldo Awal')
    except Exception as e:
        print(f"Error membaca file: {e}")
        return

    # 3. PERSIAPAN DATA ANALISIS (PANDAS PROCESSING)
    print("Sedang melakukan kalkulasi statistik...")
    
    # A. KPI Utama
    total_omzet = df_jual['Netto'].sum()
    total_transaksi = len(df_jual)
    avg_basket_size = total_omzet / total_transaksi if total_transaksi > 0 else 0
    total_piutang = df_saldo['Sisa Piutang'].sum()
    total_terbayar = df_bayar['Jumlah Bayar'].sum()

    # B. Agregasi Salesman (Performance)
    sales_perf = df_jual.groupby('Nama Sales')[['Netto', 'Qty']].sum().sort_values('Netto', ascending=False)
    
    # C. Agregasi Produk (Top 10)
    prod_perf = df_jual.groupby('Nama Barang')[['Qty', 'Total']].sum().sort_values('Total', ascending=False).head(10)

    # D. Agregasi Metode Bayar
    pay_method = df_bayar.groupby('Metode')['Jumlah Bayar'].sum()

    # E. Agregasi Umur Piutang
    aging_summary = df_saldo.groupby('Kategori Umur Piutang')['Sisa Piutang'].sum()
    # Urutkan aging agar rapi di chart (Custom Sort)
    order_aging = ["-30 Hari", "-25 Hari", "-15 Hari", "0 Hari", "5 Hari", "7 Hari", "30 Hari", "32 Hari", "45 Hari"]
    aging_summary = aging_summary.reindex(order_aging).fillna(0)

    # 4. MEMBUAT FILE LAPORAN (XLSXWRITER)
    # Nama file output ada timestamp agar tidak tertimpa
    output_file = f"Laporan_Eksekutif_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    workbook = writer.book

    # --- FORMAT STYLING ---
    fmt_title = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#1f497d'})
    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#4f81bd', 'font_color': 'white', 'border': 1})
    fmt_currency = workbook.add_format({'num_format': '#,##0', 'border': 1})
    fmt_number = workbook.add_format({'num_format': '#,##0', 'border': 1})
    
    # Format Khusus Dashboard (KPI Cards)
    fmt_card_val = workbook.add_format({'bold': True, 'font_size': 14, 'num_format': '#,##0', 'bg_color': '#f2f2f2', 'align': 'center', 'font_color': '#008000', 'border': 1})
    fmt_card_title = workbook.add_format({'bold': True, 'font_size': 10, 'bg_color': '#f2f2f2', 'align': 'center', 'border': 1})

    # =========================================================================
    # SHEET 1: DASHBOARD (UTAMA)
    # =========================================================================
    ws_dash = workbook.add_worksheet('Dashboard Utama')
    ws_dash.hide_gridlines(2)
    ws_dash.set_column('A:A', 2)
    ws_dash.set_column('B:F', 20)

    ws_dash.write('B2', "Eksekutif DASHBOARD", fmt_title)
    ws_dash.write('B3', f"Analisa Data: {os.path.basename(input_file)}")

    # -- KPI CARDS --
    kpis = [
        ("Total Penjualan (Net)", total_omzet),
        ("Total Transaksi", total_transaksi),
        ("Rata-rata Order", avg_basket_size),
        ("Total Pembayaran Masuk", total_terbayar),
        ("Total Sisa Piutang Awal", total_piutang)
    ]

    row_card = 5
    col_card = 1
    for title, value in kpis:
        ws_dash.merge_range(row_card, col_card, row_card, col_card+1, title, fmt_card_title)
        ws_dash.merge_range(row_card+1, col_card, row_card+2, col_card+1, value, fmt_card_val)
        col_card += 3 

    # -- GRAFIK DASHBOARD (Top 5 Sales) --
    ws_dash.write('B10', "Top 5 Salesman Performance", fmt_header)
    top5_sales = sales_perf.head(5)
    
    r = 10
    ws_dash.write(r, 1, "Salesman", fmt_header)
    ws_dash.write(r, 2, "Omzet", fmt_header)
    
    for name, row in top5_sales.iterrows():
        r += 1
        ws_dash.write(r, 1, name, fmt_number)
        ws_dash.write(r, 2, row['Netto'], fmt_currency)

    chart_dash = workbook.add_chart({'type': 'column'})
    chart_dash.add_series({
        'name': 'Omzet',
        'categories': ['Dashboard Utama', 11, 1, r, 1],
        'values':     ['Dashboard Utama', 11, 2, r, 2],
        'data_labels': {'value': True, 'num_format': '#,##0,, "jt"'},
    })
    chart_dash.set_title({'name': 'Top 5 Salesman Contribution'})
    chart_dash.set_style(10)
    ws_dash.insert_chart('E10', chart_dash, {'x_scale': 1.5, 'y_scale': 1.2})

    # =========================================================================
    # SHEET 2: ANALISA SALES & PRODUK
    # =========================================================================
    ws_sales = workbook.add_worksheet('Analisa Sales & Produk')
    ws_sales.set_column('A:C', 15)
    ws_sales.set_column('E:G', 15)

    ws_sales.write('A1', "Detail Performa Salesman", fmt_title)
    ws_sales.write('A3', "Nama Sales", fmt_header)
    ws_sales.write('B3', "Total Omzet", fmt_header)
    ws_sales.write('C3', "Total Qty", fmt_header)

    r = 3
    for name, row in sales_perf.iterrows():
        r += 1
        ws_sales.write(r, 0, name)
        ws_sales.write(r, 1, row['Netto'], fmt_currency)
        ws_sales.write(r, 2, row['Qty'], fmt_number)
    
    ws_sales.write('E1', "Top 10 Produk Terlaris", fmt_title)
    ws_sales.write('E3', "Nama Barang", fmt_header)
    ws_sales.write('F3', "Qty Terjual", fmt_header)
    ws_sales.write('G3', "Nilai (Rp)", fmt_header)

    r_prod = 3
    for name, row in prod_perf.iterrows():
        r_prod += 1
        ws_sales.write(r_prod, 4, name)
        ws_sales.write(r_prod, 5, row['Qty'], fmt_number)
        ws_sales.write(r_prod, 6, row['Total'], fmt_currency)

    chart_sales = workbook.add_chart({'type': 'bar'})
    chart_sales.add_series({
        'name': 'Total Omzet',
        'categories': ['Analisa Sales & Produk', 4, 0, r, 0],
        'values':     ['Analisa Sales & Produk', 4, 1, r, 1],
        'fill':       {'color': '#109618'},
    })
    chart_sales.set_title({'name': 'Ranking Salesman (Omzet)'})
    ws_sales.insert_chart('A15', chart_sales)

    chart_pie = workbook.add_chart({'type': 'pie'})
    chart_pie.add_series({
        'name': 'Kontribusi Produk',
        'categories': ['Analisa Sales & Produk', 4, 4, r_prod, 4],
        'values':     ['Analisa Sales & Produk', 4, 6, r_prod, 6],
        'data_labels': {'percentage': True},
    })
    chart_pie.set_title({'name': 'Share Omzet per Produk'})
    ws_sales.insert_chart('E15', chart_pie)

    # =========================================================================
    # SHEET 3: ANALISA KEUANGAN (PIUTANG)
    # =========================================================================
    ws_fin = workbook.add_worksheet('Analisa Keuangan')
    ws_fin.set_column('A:B', 20)
    
    ws_fin.write('A1', "Analisa Umur Piutang (Aging)", fmt_title)
    ws_fin.write('A3', "Kategori Umur", fmt_header)
    ws_fin.write('B3', "Total Piutang", fmt_header)

    r_age = 3
    for cat, val in aging_summary.items():
        r_age += 1
        ws_fin.write(r_age, 0, cat)
        ws_fin.write(r_age, 1, val, fmt_currency)

    chart_aging = workbook.add_chart({'type': 'column'})
    chart_aging.add_series({
        'name': 'Sisa Piutang',
        'categories': ['Analisa Keuangan', 4, 0, r_age, 0],
        'values':     ['Analisa Keuangan', 4, 1, r_age, 1],
        'fill':       {'color': '#dc3912'},
    })
    chart_aging.set_title({'name': 'Distribusi Umur Piutang'})
    chart_aging.set_y_axis({'name': 'Nilai Rupiah'})
    ws_fin.insert_chart('D3', chart_aging, {'x_scale': 1.5})

    ws_fin.write('A20', "Metode Pembayaran", fmt_title)
    ws_fin.write('A22', "Metode", fmt_header)
    ws_fin.write('B22', "Total Masuk", fmt_header)
    
    r_pay = 22
    for met, val in pay_method.items():
        r_pay += 1
        ws_fin.write(r_pay, 0, met)
        ws_fin.write(r_pay, 1, val, fmt_currency)
        
    chart_donut = workbook.add_chart({'type': 'doughnut'})
    chart_donut.add_series({
        'name': 'Metode Bayar',
        'categories': ['Analisa Keuangan', 23, 0, r_pay, 0],
        'values':     ['Analisa Keuangan', 23, 1, r_pay, 1],
        'data_labels': {'percentage': True, 'value': True, 'separator': '\n'},
    })
    chart_donut.set_title({'name': 'Proporsi Metode Pembayaran'})
    ws_fin.insert_chart('D20', chart_donut)

    writer.close()
    print("==================================================")
    print("LAPORAN SELESAI DIBUAT!")
    print(f"File Output: {os.path.abspath(output_file)}")
    print("==================================================")

if __name__ == "__main__":
    generate_analyst_report()