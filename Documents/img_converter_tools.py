import os
from PIL import Image

def konversi_gambar_otomatis():
    current_folder = os.path.dirname(os.path.abspath(__file__))
    output_folder_name = "Hasil_Konversi"
    output_folder = os.path.join(current_folder, output_folder_name)
    
    INPUT_EXTENSIONS = ('.jpg', '.jpeg', '.png', '.webp', '.bmp', '.tiff', '.gif')

    print(f"--- KONVERTER GAMBAR OTOMATIS ---")
    print(f"Lokasi Scan: {current_folder}")
    
    files = [f for f in os.listdir(current_folder) 
             if f.lower().endswith(INPUT_EXTENSIONS)]
    
    if not files:
        print("\n[!] Tidak ada file gambar ditemukan di folder ini.")
        input("Tekan Enter untuk keluar...")
        return

    print(f"Ditemukan {len(files)} gambar yang bisa dikonversi.")
    
    formats = {
        '1': {'ext': '.jpg', 'format': 'JPEG', 'name': 'JPG (Cocok untuk foto)'},
        '2': {'ext': '.png', 'format': 'PNG', 'name': 'PNG (Mendukung transparansi)'},
        '3': {'ext': '.webp', 'format': 'WEBP', 'name': 'WEBP (Format web modern, ringan)'},
        '4': {'ext': '.bmp', 'format': 'BMP', 'name': 'BMP (Bitmap standar)'},
        '5': {'ext': '.gif', 'format': 'GIF', 'name': 'GIF (Gambar diam/animasi)'}
    }

    print("\nIngin diubah ke format apa?")
    for k, v in formats.items():
        print(f" [{k}] {v['name']}")

    choice = input("\nMasukkan nomor pilihan (1-5): ").strip()

    if choice not in formats:
        print("[!] Pilihan tidak valid.")
        return

    target = formats[choice]
    print(f"\nMemulai konversi ke format {target['format']}...")
    
    os.makedirs(output_folder, exist_ok=True)

    success_count = 0
    
    for filename in files:
        input_path = os.path.join(current_folder, filename)
        
        filename_no_ext = os.path.splitext(filename)[0]
        new_filename = f"{filename_no_ext}{target['ext']}"
        output_path = os.path.join(output_folder, new_filename)

        try:
            with Image.open(input_path) as img:
                if target['format'] == 'JPEG':
                    if img.mode in ('RGBA', 'P'):
                        img = img.convert('RGB')
                        
                if target['format'] == 'GIF':
                     img.save(output_path, target['format'])
                else:
                    img.save(output_path, target['format'])
                
                print(f"[OK] {filename} -> {new_filename}")
                success_count += 1

        except Exception as e:
            print(f"[ERROR] Gagal mengonversi {filename}: {e}")

    print("-" * 30)
    print(f"Selesai! {success_count} dari {len(files)} berhasil dikonversi.")
    print(f"Cek folder: {output_folder}")
    input("Tekan Enter untuk menutup...")

if __name__ == "__main__":
    konversi_gambar_otomatis()