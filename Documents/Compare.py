import json
import xml.etree.ElementTree as ET
import re
import datetime
import html
from difflib import SequenceMatcher
import glob

BATAS_KEMIRIPAN = 0.65

def cari_file_xml():
    file_xml = glob.glob('Blogger/Blogs/**/*.xml', recursive=True)
    file_xml.extend(glob.glob('Blogger/Blogs/**/*.atom', recursive=True))
    return [f for f in file_xml if 'theme-layouts' not in f.lower()]

def cari_file_json():
    return glob.glob('Facebook/**/your_posts_*.json', recursive=True)

def bersihkan_teks(teks):
    if not teks:
        return ""
    teks = html.unescape(teks)
    teks = re.sub(r'<[^>]+>', ' ', teks)
    teks = teks.lower()
    teks = re.sub(r'[^a-z0-9\s]', '', teks)
    teks = re.sub(r'\s+', ' ', teks).strip()
    return teks

def muat_data_blogger(daftar_file):
    teks_blogger = []
    for nama_file in daftar_file:
        print(f"--> Membaca data Blogger dari: {nama_file}")
        try:
            tree = ET.parse(nama_file)
            root = tree.getroot()
            
            for entry in root.iter():
                if entry.tag.endswith('entry'):
                    for child in entry.iter():
                        if child.tag.endswith('content'):
                            konten_kotor = "".join(child.itertext())
                            if konten_kotor:
                                teks_bersih = bersihkan_teks(konten_kotor)
                                if len(teks_bersih) > 10:
                                    teks_blogger.append(teks_bersih)
                            break
        except Exception as e:
            print(f"--> Error membaca XML Blogger {nama_file}: {e}")
    return teks_blogger

def muat_data_facebook(daftar_file):
    postingan_fb = []
    for nama_file in daftar_file:
        print(f"--> Membaca data Facebook dari: {nama_file}")
        try:
            with open(nama_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            for item in data:
                judul_aktivitas = item.get('title', '').lower()
                if 'mengomentari' in judul_aktivitas or 'membalas' in judul_aktivitas:
                    continue

                if 'data' in item:
                    for d in item['data']:
                        if 'post' in d and isinstance(d['post'], str):
                            teks_asli = d['post']
                            teks_bersih = bersihkan_teks(teks_asli)
                            
                            if len(teks_bersih) > 10: 
                                postingan_fb.append({
                                    'teks_asli': teks_asli,
                                    'teks_bersih': teks_bersih,
                                    'tanggal': item.get('timestamp', 0) 
                                })
        except Exception as e:
            print(f"--> Error membaca JSON Facebook {nama_file}: {e}")
    return postingan_fb

def periksa_kemiripan(teks_fb, daftar_teks_blogger):
    for teks_blog in daftar_teks_blogger:
        if not teks_blog or not teks_fb:
            continue
            
        if teks_fb in teks_blog or teks_blog in teks_fb:
            return True
            
        rasio = SequenceMatcher(None, teks_fb, teks_blog).ratio()
        if rasio >= BATAS_KEMIRIPAN:
            return True
    return False

def jalankan_komparasi():
    daftar_file_xml = cari_file_xml()
    daftar_file_json = cari_file_json()
    
    if not daftar_file_xml:
        print("--> Data XML Blogger tidak ditemukan.")
        return
        
    if not daftar_file_json:
        print("--> Data JSON Facebook tidak ditemukan.")
        return

    daftar_blogger = muat_data_blogger(daftar_file_xml)
    print(f"--> Berhasil memuat {len(daftar_blogger)} postingan dari Blogger.")
    
    daftar_fb = muat_data_facebook(daftar_file_json)
    print(f"--> Berhasil memuat {len(daftar_fb)} postingan dari Facebook.")
    
    if not daftar_fb:
        print("--> Data postingan Facebook kosong.")
        return

    print("--> Membandingkan data...")
    
    postingan_baru = []
    for fb in daftar_fb:
        sudah_ada = periksa_kemiripan(fb['teks_bersih'], daftar_blogger)
        if not sudah_ada:
            postingan_baru.append(fb)
            
    file_hasil = 'daftar_belum_diposting.txt'
    print(f"--> Menyimpan hasil: Ditemukan {len(postingan_baru)} postingan yang belum ada di Blogger.")
    
    with open(file_hasil, 'w', encoding='utf-8') as file_out:
        for index, item in enumerate(postingan_baru, 1):
            tanggal_baca = datetime.datetime.fromtimestamp(item['tanggal']).strftime('%Y-%m-%d %H:%M:%S')
            file_out.write(f"[{index}] Tanggal Asli FB: {tanggal_baca}\n")
            file_out.write(f"Teks Postingan:\n{item['teks_asli']}\n")
            file_out.write("-" * 50 + "\n\n")
            
    print(f"--> Selesai! Silakan buka file '{file_hasil}'.")

if __name__ == '__main__':
    jalankan_komparasi()
