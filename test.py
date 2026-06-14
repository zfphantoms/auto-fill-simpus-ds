from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import re

# Konfigurasi opsi Chrome
options = Options()
options.add_experimental_option("detach", True)  # Agar browser tidak langsung tertutup

# Inisialisasi driver Chrome
# Pastikan chromedriver.exe ada di folder yang sama dengan script ini

driver = webdriver.Chrome(options=options)
driver.maximize_window()

# Buka website SIMPUS Deli Serdang
driver.get("https://dinkesds-simpus.deliserdangkab.go.id")

# Tunggu beberapa detik agar halaman terlihat
print("Website sudah dibuka. Silakan login secara manual, isi captcha, dan klik LOGIN.")
time.sleep(10)

# Tunggu sampai user login dan halaman beralih ke /home
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

try:
    WebDriverWait(driver, 300).until(
        lambda d: d.current_url.startswith("https://dinkesds-simpus.deliserdangkab.go.id/home")
    )
    print("Berhasil login. Silakan lakukan sortir dan klik tombol Detail secara manual.")
    # Tunggu sampai tab baru (laporan fingerprint) terbuka setelah klik Detail manual
    print("Menunggu tab baru laporan fingerprint pasien dibuka...")
    WebDriverWait(driver, 600).until(lambda d: len(d.window_handles) > 1)
    driver.switch_to.window(driver.window_handles[-1])
    print("Berpindah ke tab laporan fingerprint pasien.")
    # Tunggu tabel muncul
    WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.XPATH, '//table'))
    )
    # Ambil semua NIK dari kolom NIK (asumsi kolom ke-3)
    rows = driver.find_elements(By.XPATH, '//table//tr[position()>1]')
    nik_list = []
    for row in rows:
        cols = row.find_elements(By.TAG_NAME, 'td')
        if len(cols) >= 3:
            nik = cols[2].text.strip()
            if nik:
                nik_list.append(nik)
    print(f"Ditemukan {len(nik_list)} NIK.")
    # Simpan ke Excel
    import pandas as pd
    df_nik = pd.DataFrame({'NIK': nik_list})
    df_nik.to_excel('hasil_nik.xlsx', index=False)
    print("Semua NIK berhasil disimpan ke hasil_nik.xlsx")
except Exception as e:
    print(f"Gagal mendeteksi halaman home atau ekstrak NIK: {e}")
# (Browser tetap terbuka karena detach=True)
