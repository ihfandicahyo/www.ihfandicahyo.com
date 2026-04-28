import requests

url = input("--> masukkan alamat situs: ")
if not url.startswith("http"):
    url = "https://" + url

try:
    response = requests.get(url)
    with open("source.html", "w", encoding="utf-8") as file:
        file.write(response.text)
    print("--> Selesai: source.html telah dibuat")
except Exception as e:
    print("--> Gagal: " + str(e))
