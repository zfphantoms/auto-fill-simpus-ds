from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook
from datetime import datetime
import time
import pandas as pd

# === KONFIGURASI ===
URL_LOGIN = "https://dinkesds-simpus.deliserdangkab.go.id/"
URL_PASIEN = "https://dinkesds-simpus.deliserdangkab.go.id/pasien"
EXCEL_PATH_SISWA = "SD KATOLIK utk puskesmas.xlsx"
EXCEL_PATH_KEC = "List_Kecamatan.xlsx"

# === INPUT USER ===
NOMOR_SISWA_AWAL = int(input("Masukkan nomor urut siswa awal: "))
NOMOR_SISWA_AKHIR = int(input("Masukkan nomor urut siswa akhir: "))

if NOMOR_SISWA_AKHIR < NOMOR_SISWA_AWAL:
    raise ValueError("❌ Nomor siswa akhir harus lebih besar atau sama dengan awal.")

ROW_TARGET = NOMOR_SISWA_AWAL + 1  # Baris di Excel = nomor urut siswa + 1 (karena header)
JUMLAH_DATA = NOMOR_SISWA_AKHIR - NOMOR_SISWA_AWAL + 1

# === LOAD DATA REFERENSI KECAMATAN ===
df_ref = pd.read_excel(EXCEL_PATH_KEC)
df_ref.columns = df_ref.columns.str.upper()
dict_kec_to_kab = dict(zip(df_ref["KEC"].str.upper(), df_ref["KAB/KOTA"]))
dict_kec_to_prov = dict(zip(df_ref["KEC"].str.upper(), df_ref["PROVINSI"]))

# === BUKA BROWSER ===
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=chrome_options)
driver.maximize_window()
driver.get(URL_LOGIN)

print("✅ Silakan login secara manual (isi CAPTCHA & klik LOGIN)")
input("🔐 Tekan ENTER setelah berhasil login...")

# === BUKA HALAMAN PASIEN ===
driver.get(URL_PASIEN)
time.sleep(2)

# === LOAD DATA SISWA ===
wb = load_workbook(EXCEL_PATH_SISWA)
sheet = wb.active

# === LOOP DATA SISWA ===
for i in range(JUMLAH_DATA):
    row = sheet[ROW_TARGET + i]

    nama = row[1].value
    jk = row[3].value
    tempat_lahir = row[5].value
    tanggal_lahir_raw = row[6].value
    nik = str(row[7].value)
    agama = row[8].value
    alamat = row[9].value
    kecamatan = row[12].value  # Kolom ke-13 (index 12)

    # Format tanggal lahir
    try:
        if isinstance(tanggal_lahir_raw, datetime):
            tanggal_lahir = tanggal_lahir_raw.strftime("%m/%d/%Y")
        else:
            tanggal_obj = datetime.strptime(str(tanggal_lahir_raw), "%Y-%m-%d")
            tanggal_lahir = tanggal_obj.strftime("%m/%d/%Y")
    except Exception:
        print(f"⚠️ Gagal format tanggal untuk {nama}: {tanggal_lahir_raw}")
        tanggal_lahir = ""

    # Ambil Provinsi & Kabupaten dari kecamatan
    kec_upper = str(kecamatan).strip().upper()
    provinsi = dict_kec_to_prov.get(kec_upper, "")
    kabupaten = dict_kec_to_kab.get(kec_upper, "")

    # === FORMAT KHUSUS KECAMATAN ===
    kecamatan_raw = str(kecamatan).strip()
    if kecamatan_raw == "Stm Hilir":
        kecamatan_clean = "SINEMBAH TANJUNG MUDA HILIR"
    elif kecamatan_raw == "Stm Hulu":
        kecamatan_clean = "SINEMBAH TANJUNG MUDA HULU"
    else:
        kec_upper = kecamatan_raw.upper()
        kecamatan_clean = kec_upper
        if kecamatan_clean.startswith("KEC. "):
            kecamatan_clean = kecamatan_clean[5:]
        if "SIBIRU-BIRU" in kecamatan_clean:
            kecamatan_clean = "BIRU-BIRU"

    print(f"\n🟢 Silakan klik tombol TAMBAH DATA untuk: {nama}")
    input("➡️ Setelah form terbuka dan dropdown wilayah muncul, tekan ENTER untuk isi otomatis...")

    try:
        # === ISI FORM ===
        driver.find_element(By.XPATH, '//input[@placeholder="Nama"]').clear()
        driver.find_element(By.XPATH, '//input[@placeholder="Nama"]').send_keys(nama)

        Select(driver.find_element(By.XPATH, '//select[./option[contains(text(), "Jenis Kelamin")]]')) \
            .select_by_visible_text("Perempuan" if jk.upper() == "P" else "Laki-Laki")

        driver.find_element(By.XPATH, '//input[@placeholder="Tempat Lahir"]').clear()
        driver.find_element(By.XPATH, '//input[@placeholder="Tempat Lahir"]').send_keys(tempat_lahir)

        driver.find_element(By.XPATH, '//input[@placeholder="Tanggal Lahir"]').clear()
        driver.find_element(By.XPATH, '//input[@placeholder="Tanggal Lahir"]').send_keys(tanggal_lahir)

        driver.find_element(By.XPATH, '//input[@placeholder="NIK"]').clear()
        driver.find_element(By.XPATH, '//input[@placeholder="NIK"]').send_keys(nik)

        Select(driver.find_element(By.XPATH, '//select[./option[contains(text(), "Agama")]]')) \
            .select_by_visible_text(agama.strip().upper())

        driver.find_element(By.XPATH, '//input[@placeholder="Alamat"]').clear()
        driver.find_element(By.XPATH, '//input[@placeholder="Alamat"]').send_keys(alamat)

        # === PROVINSI ===
        prov_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//label[contains(text(), "Provinsi")]/following-sibling::div//input'))
        )
        prov_input.click()
        time.sleep(0.5)
        prov_input.clear()

        prov_input.send_keys(provinsi)
        time.sleep(1.5)
        # === KASUS KHUSUS: Jika provinsi adalah "RIAU", tekan ARROW_DOWN dua kali
        if provinsi.strip().upper() == "RIAU":
            prov_input.send_keys(Keys.ARROW_DOWN)
            prov_input.send_keys(Keys.ARROW_DOWN)
        else:
            prov_input.send_keys(Keys.ARROW_DOWN)

        prov_input.send_keys(Keys.ENTER)

        # === KABUPATEN/KOTA ===
        kab_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//label[contains(text(), "Kabupaten")]/following-sibling::div//input'))
        )
        kab_input.click()
        time.sleep(0.5)
        kab_input.clear()
        kab_input.send_keys(kabupaten)
        time.sleep(1.5)
        kab_input.send_keys(Keys.ARROW_DOWN)
        kab_input.send_keys(Keys.ENTER)

        # === KECAMATAN ===
        kec_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//label[contains(text(), "Kecamatan")]/following-sibling::div//input'))
        )
        kec_input.click()
        time.sleep(0.5)
        kec_input.clear()
        kec_input.send_keys(kecamatan_clean)
        time.sleep(1.5)
        kec_input.send_keys(Keys.ARROW_DOWN)
        kec_input.send_keys(Keys.ENTER)

        print("✅ Data berhasil diisi. Silakan klik tombol TAMBAH secara manual.")

    except Exception as e:
        print(f"❌ Gagal input: {nama} → {e}")
        continue
