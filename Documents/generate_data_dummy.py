import pandas as pd
import random
import os
import math
import numpy as np
from datetime import datetime, timedelta

def generate_dynamic_dummy_v9():
    print("=== GENERATOR DATA DUMMY ===")
    
    # --- INPUT USER ---
    try:
        bulan = int(input("Masukkan bulan (1-12): "))
        tahun = 2025
        total_data = int(input("Masukkan jumlah baris data penjualan: "))
        
        # --- FITUR BARU: KONTROL GAP TARGET ---
        print("\n--- PENGATURAN TARGET ---")
        print("Agar data realistis, target akan dibuat berdasarkan penjualan yang terjadi.")
        persen_ach = float(input("Rata-rata % pencapaian yang diinginkan (misal 95 artinya target sedikit diatas omzet): ")) / 100
        variasi = float(input("Variasi % antar sales (misal 10 artinya pencapaian berkisar +/- 10% dari rata-rata): ")) / 100
        
        # Logika Rasio Pelanggan
        total_pelanggan = int(math.sqrt(total_data) * 3)
        if total_pelanggan < 5: total_pelanggan = 5 
        if total_pelanggan > total_data: total_pelanggan = total_data
        
        jumlah_sales = int(input("\nBerapa jumlah Salesman? "))
        list_sales = [input(f"Masukkan nama Salesman ke-{i+1}: ") for i in range(jumlah_sales)]
            
    except ValueError:
        print("Input harus berupa angka!")
        return

    nama_file = f"Data_V9_Dummy_Bulan_{bulan}.xlsx"
    
    # --- DATA MASTER BARANG ---
    daftar_barang = {
        "BSM 140 ml": 6300, "BSM 450 ml": 15000, "BSMJ 5.5 kg": 109000, 
        "BSMJ 11 kg": 214500, "BSMJ 24 kg": 447000, "BRS 140 ml": 6300, 
        "BRS 400 ml": 13300, "BRSJ 5.5 kg": 109000, "NY 5,5 kg": 103500, 
        "Hiap HongJ 5,5 kg": 101500, "BSA 600 ml": 13850, "RSM 275 ml": 7900, 
        "RSM 450 ml": 13200, "RSM 5.5 kg": 103000, "REP 520 ml": 10500, "MKH-K2": 2700,
        "SCHS-8 gr": 20400, "SCHS-5,7 KG": 76700, "STKC-8 gr": 17800, "STKC-5,7 KG": 67500
    }

    # --- DATABASE NAMA & ALAMAT ---
    prefix_usaha = ["Toko", "Warung", "UD", "CV", "Agen", "Grosir", "Depot", "Kios", "TB", "Rumah Makan", "Catering", "Bakso", "Soto", "Mie Ayam", "Sate"]
    nama_sifat = ["Lancar", "Jaya", "Abadi", "Makmur", "Sentosa", "Berkah", "Barokah", "Rejeki", "Sumber", "Sido", "Maju", "Sari", "Rasa", "Nikmat"]
    nama_orang_jawa = ["Slamet", "Widodo", "Santoso", "Budi", "Hartono", "Sutrisno", "Wahyu", "Agus", "Sri", "Endang", "Bambang", "Yanto", "Eko"]
    suffix_tempat = ["Semarang", "Jaya", "Baru", "Raya", "Putra", "Putri", "Group", "Mandiri", "Tengah", "Timur", "Selatan", "Utara"]
    
    jalan_semarang_real = [
        "Jl. Pandanaran", "Jl. Pemuda", "Jl. Gajah Mada", "Jl. Ahmad Yani", "Jl. Pahlawan", "Jl. MH Thamrin",
        "Jl. Jend. Sudirman", "Jl. Siliwangi", "Jl. Pamularsih Raya", "Jl. Abdulrahman Saleh", "Jl. Kokrosono",
        "Jl. Majapahit", "Jl. Wolter Monginsidi", "Jl. Fatmawati", "Jl. Soekarno Hatta", "Jl. Brigjen Sudiarto",
        "Jl. Setiabudi", "Jl. Ngesrep Timur V", "Jl. Tirto Agung", "Jl. Banjarsari", "Jl. Durian Raya",
        "Jl. Dr. Cipto", "Jl. MT Haryono", "Jl. Mataram", "Jl. Sriwijaya", "Jl. Veteran", "Jl. Kyai Saleh"
    ]

    # --- GENERATE PELANGGAN & SALES MAPPING ---
    pelanggan_data = {}
    pelanggan_sales_map = {}
    
    print("   -> Menciptakan nama pelanggan realistis...")
    tries = 0
    while len(pelanggan_data) < total_pelanggan:
        pola = random.choice([1, 1, 2, 2, 3])
        if pola == 1:
            nama_baru = f"{random.choice(prefix_usaha)} {random.choice(nama_sifat)} {random.choice(nama_orang_jawa)}"
        elif pola == 2:
            nama_baru = f"{random.choice(prefix_usaha)} {random.choice(nama_sifat)} {random.choice(suffix_tempat)}"
        else:
            nama_baru = f"{random.choice(prefix_usaha)} {random.choice(nama_orang_jawa)} {random.choice(suffix_tempat)}"
            
        nama_split = nama_baru.split()
        nama_fix = " ".join(sorted(set(nama_split), key=nama_split.index))

        if nama_fix not in pelanggan_data:
            jalan = random.choice(jalan_semarang_real)
            nomor = random.randint(1, 900)
            if random.random() > 0.7: 
                alamat_fix = f"{jalan} No. {nomor} Kav. {random.randint(1,5)}, Semarang"
            else:
                alamat_fix = f"{jalan} No. {nomor}, Semarang"
            
            pelanggan_data[nama_fix] = alamat_fix
            pelanggan_sales_map[nama_fix] = random.choice(list_sales)
        tries += 1
        if tries > 10000: break
    pelanggan_names = list(pelanggan_data.keys())

    # --- 1. SHEET PENJUALAN ---
    print(f"\nMemulai proses generate {total_data} transaksi...")
    data_penjualan = []
    inv_counter = 1001
    
    for _ in range(total_data):
        tgl = datetime(tahun, bulan, random.randint(1, 28)) 
        cust_nama = random.choice(pelanggan_names)
        cust_alamat = pelanggan_data[cust_nama]
        sales = pelanggan_sales_map[cust_nama]
        
        inv_no = f"INV/SMG/{tahun}/{bulan:02d}/{inv_counter}"
        barang = random.choice(list(daftar_barang.keys()))
        harga = daftar_barang[barang]
        qty = random.randint(1, 50)
        total = harga * qty
        diskon = 0
        if total > 500000: diskon = random.choice([0, 5000, 10000, 25000])
        retur = random.choices([0, harga], weights=[97, 3])[0] 
        netto = total - diskon - retur
        top = random.choice(["14 Hari", "30 Hari"])
        
        data_penjualan.append([tgl, cust_nama, cust_alamat, sales, inv_no, top, barang, qty, harga, total, diskon, retur, netto])
        inv_counter += 1

    df_penjualan = pd.DataFrame(data_penjualan, columns=[
        'Tanggal', 'Nama Pelanggan', 'Alamat', 'Nama Sales', 'No. Faktur', 'TOP', 
        'Nama Barang', 'Qty', 'Harga Satuan', 'Total', 'Diskon', 'Retur', 'Netto'
    ]).sort_values('Tanggal')

    # --- 2. SHEET SALDO AWAL (PIUTANG) ---
    print("Menyusun data Saldo Awal...")
    data_saldo = []
    ref_date = datetime(tahun, bulan, 1)
    kategori_aging_list = ["-30 Hari", "-25 Hari", "-15 Hari", "0 Hari", "5 Hari", "7 Hari", "30 Hari", "32 Hari", "45 Hari"]
    
    for p_nama in pelanggan_names[:int(total_pelanggan * 0.35)]:
        sales_saldo = pelanggan_sales_map[p_nama]
        sisa_piutang = random.randint(5, 150) * 50000
        aging_choice = random.choice(kategori_aging_list)
        days_offset = int(aging_choice.split()[0])
        standard_top = 30 
        invoice_age_days = standard_top + days_offset
        if invoice_age_days < 1: invoice_age_days = 1
        real_inv_date = ref_date - timedelta(days=invoice_age_days)
        no_faktur_lama = f"INV/SMG/{real_inv_date.year}/{real_inv_date.month:02d}/{random.randint(100, 999)}"
        data_saldo.append([real_inv_date, no_faktur_lama, p_nama, sales_saldo, sisa_piutang, aging_choice])

    df_saldo = pd.DataFrame(data_saldo, columns=[
        'Tanggal Faktur', 'No. Faktur Lama', 'Nama Pelanggan', 'Nama Sales', 'Sisa Piutang', 'Kategori Umur Piutang'
    ]).sort_values('Tanggal Faktur')

    # --- 3. SHEET PEMBAYARAN ---
    print("Menyusun data Pembayaran...")
    data_pembayaran = []
    df_pelunasan_current = df_penjualan.sample(frac=0.6)
    for _, row in df_pelunasan_current.iterrows():
        hari_tambah = random.randint(1, 14)
        tgl_b = row['Tanggal'] + timedelta(days=hari_tambah)
        if tgl_b.month == bulan:
            metode = random.choices(["TRANSFER", "TUNAI"], weights=[70, 30])[0]
            data_pembayaran.append([tgl_b, row['Nama Pelanggan'], row['Nama Sales'], row['No. Faktur'], row['Netto'], metode])
    
    for _, row in df_saldo.sample(frac=0.7).iterrows():
        tgl_b = datetime(tahun, bulan, random.randint(1, 28))
        metode = random.choices(["TRANSFER", "TUNAI"], weights=[80, 20])[0]
        data_pembayaran.append([tgl_b, row['Nama Pelanggan'], row['Nama Sales'], row['No. Faktur Lama'], row['Sisa Piutang'], metode])

    df_pembayaran = pd.DataFrame(data_pembayaran, columns=[
        'Tanggal Bayar', 'Nama Pelanggan', 'Nama Sales', 'No. Faktur', 'Jumlah Bayar', 'Metode'
    ]).sort_values('Tanggal Bayar')

    # --- 4. SHEET TARGET SALES REALISTIS ---
    print("Menyusun Target Realistis (Reverse Engineered from Sales)...")
    
    # Hitung Actual Sales dulu
    actual_summary = df_penjualan.groupby(['Nama Sales', 'Nama Barang']).agg({'Qty': 'sum', 'Netto': 'sum'}).reset_index()
    
    raw_target_data = []
    
    # Pastikan semua kombinasi sales & barang ada targetnya, meski penjualan 0
    for sales in list_sales:
        for barang, harga in daftar_barang.items():
            # Cari apakah ada penjualan
            match = actual_summary[(actual_summary['Nama Sales'] == sales) & (actual_summary['Nama Barang'] == barang)]
            
            if not match.empty:
                actual_qty = match.iloc[0]['Qty']
            else:
                actual_qty = 0
            
            # --- LOGIKA TARGET REALISTIS ---
            # Hitung faktor variasi unik untuk item ini (Random Normal Distribution)
            # Agar tidak flat 95% semua, kita beri noise +/- variasi
            noise = np.random.uniform(-variasi, variasi) 
            target_ratio = persen_ach + noise 
            
            # Mencegah ratio 0 atau negatif
            if target_ratio <= 0.1: target_ratio = 0.5

            if actual_qty > 0:
                # Jika ada penjualan: Target = Actual / Ratio
                # Contoh: Jual 100, Ratio 0.9 (90%) -> Target = 111
                target_qty_calc = int(actual_qty / target_ratio)
            else:
                # Jika tidak ada penjualan, buat target kecil dummy (potensi lost sales)
                target_qty_calc = random.randint(5, 20)
            
            # Bulatkan ke puluhan
            target_qty_final = int(round(target_qty_calc / 5) * 5)
            if target_qty_final == 0: target_qty_final = 5
            
            target_omzet = target_qty_final * harga
            
            raw_target_data.append({
                'Nama Barang': barang,
                'Nama Sales': sales,
                'Target Qty': target_qty_final,
                'Target Value': target_omzet
            })

    df_raw_target = pd.DataFrame(raw_target_data)
    
    # Pivot Multi-Index
    df_pivot = df_raw_target.pivot_table(
        index='Nama Barang', columns='Nama Sales', values=['Target Qty', 'Target Value'], aggfunc='sum', fill_value=0
    )
    df_pivot = df_pivot.swaplevel(0, 1, axis=1)
    df_pivot.sort_index(axis=1, level=0, inplace=True)
    df_pivot.loc['GRAND TOTAL'] = df_pivot.sum()

    # --- SAVING ---
    print(f"Menyimpan ke {nama_file}...")
    with pd.ExcelWriter(nama_file, engine='xlsxwriter', datetime_format='[$-id-ID]dd mmm yyyy') as writer:
        df_penjualan.to_excel(writer, sheet_name='Penjualan', index=False)
        df_pembayaran.to_excel(writer, sheet_name='Pembayaran', index=False)
        df_saldo.to_excel(writer, sheet_name='Saldo Awal', index=False)
        df_pivot.to_excel(writer, sheet_name='Target Sales')

        workbook = writer.book
        indo_date_fmt = workbook.add_format({'num_format': '[$-id-ID]dd mmm yyyy', 'align': 'center'})
        num_fmt = workbook.add_format({'num_format': '#,##0'})
        center_fmt = workbook.add_format({'align': 'center'})
        header_sub_fmt = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D3D3D3', 'border': 1, 'font_size': 9})
        total_row_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'num_format': '#,##0'})

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            if sheet_name == 'Target Sales':
                worksheet.set_column(0, 0, 25) 
                col_start = 1
                for i, col_tuple in enumerate(df_pivot.columns):
                    sales_name, type_col = col_tuple
                    if 'Qty' in type_col: worksheet.set_column(col_start + i, col_start + i, 12, num_fmt)
                    else: worksheet.set_column(col_start + i, col_start + i, 18, num_fmt)
                worksheet.set_row(len(df_pivot) + 2, None, total_row_fmt)
            else:
                if sheet_name == 'Penjualan': df_ref = df_penjualan
                elif sheet_name == 'Pembayaran': df_ref = df_pembayaran
                elif sheet_name == 'Saldo Awal': df_ref = df_saldo
                for i, col in enumerate(df_ref.columns):
                    max_len = max(df_ref[col].astype(str).map(len).max(), len(str(col))) + 3
                    if "Tanggal" in col: worksheet.set_column(i, i, 18, indo_date_fmt)
                    elif any(x in col for x in ["Total", "Harga", "Netto", "Piutang", "Bayar", "Diskon", "Retur"]): worksheet.set_column(i, i, 15, num_fmt)
                    elif any(x in col for x in ["TOP", "Qty", "Umur"]): worksheet.set_column(i, i, 15, center_fmt)
                    else: worksheet.set_column(i, i, max_len)

    print(f"\nSUKSES! File: {os.path.abspath(nama_file)}")

if __name__ == "__main__":
    generate_dynamic_dummy_v9()