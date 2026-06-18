from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time
import pandas as pd
import re
import sys
import os
import tkinter as tk
from tkinter import filedialog

from wilayah_loader import load_wilayah_csv, build_wilayah_index, resolve_wilayah

# === PILIH FILE EXCEL (via File Explorer) ===
_root = tk.Tk()
_root.withdraw()
_root.attributes('-topmost', True)

print("\n📂 Silakan pilih file Excel di jendela yang muncul...")
EXCEL_PATH_SISWA = filedialog.askopenfilename(
    title="Pilih file Excel data siswa",
    filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
    initialdir="."
)
_root.destroy()

if not EXCEL_PATH_SISWA:
    print("❌ Tidak ada file yang dipilih. Keluar.")
    sys.exit(1)

print(f"✅ File yang dipilih: {EXCEL_PATH_SISWA}\n")

if EXCEL_PATH_SISWA.lower().endswith('.xls'):
    print("ℹ️ Format .xls terdeteksi. Mengonversi ke format .xlsx...")
    try:
        import xlrd
    except ImportError:
        print("❌ Library 'xlrd' belum terpasang. Menginstal otomatis...")
        import subprocess
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "xlrd"])
            import xlrd
            print("✅ 'xlrd' berhasil dipasang.")
        except Exception as inst_err:
            print(f"❌ Gagal memasang 'xlrd' otomatis: {inst_err}\nSilakan jalankan perintah berikut manual di terminal:\npip install xlrd")
            sys.exit(1)
    
    try:
        TEMP_XLSX_PATH = os.path.splitext(EXCEL_PATH_SISWA)[0] + "_temp_converted.xlsx"
        print(f"ℹ️ Mengonversi {EXCEL_PATH_SISWA} -> {TEMP_XLSX_PATH} ...")
        xls_data = pd.read_excel(EXCEL_PATH_SISWA, sheet_name=None, header=None)
        with pd.ExcelWriter(TEMP_XLSX_PATH, engine='openpyxl') as writer:
            for sheet_name, df_sheet in xls_data.items():
                df_sheet.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        EXCEL_PATH_SISWA = TEMP_XLSX_PATH
        print("✅ Konversi ke .xlsx selesai.")
    except Exception as conv_err:
        print(f"❌ Gagal mengonversi file .xls: {conv_err}")
        sys.exit(1)

input_xlsx = EXCEL_PATH_SISWA

# === INPUT RENTANG NOMOR ===
try:
    NOMOR_AWAL = int(input("Masukkan nomor urut data awal (1-based, misal 1): "))
    NOMOR_AKHIR = int(input("Masukkan nomor urut data akhir: "))
    if NOMOR_AKHIR < NOMOR_AWAL:
        raise ValueError("Nomor akhir harus >= awal.")
except ValueError as e:
    print(f"❌ Input tidak valid: {e}")
    sys.exit(1)

# Membuka browser Chrome
options = webdriver.ChromeOptions()
options.add_experimental_option('detach', True)  # Agar browser tidak langsung tertutup

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

login_url = 'https://dinkesds-simpus.deliserdangkab.go.id/'
home_url_fragment = '/home'
pasien_url = 'https://dinkesds-simpus.deliserdangkab.go.id/pasien'

# Referensi wilayah (harus disediakan)
wilayah_csv = 'wilayah_indonesia.csv'

# Nama kolom default (akan di-resolve secara fleksibel saat baca Excel)
NAME_COLUMN = 'NAMA LENGKAP'
GENDER_COLUMN = 'JENIS KELAMIN'
BIRTHPLACE_COLUMN = 'TEMPAT LAHIR'
BIRTHDATE_COLUMN = 'TANGGAL LAHIR'
NIK_COLUMN = 'NIK/KITAS/KITAP'
KK_COLUMN = 'NO KK'
ADDRESS_COLUMN = 'ALAMAT TEMPAT TINGGAL'
KECAMATAN_COLUMN = 'NAMA KECAMATAN'
DESA_COLUMN = 'NAMA DESA'

BPJS_COLUMN = 'SIPP_Nomor Kartu'
BPJS_STATUS_COLUMN = 'SIPP_Status Kepesertaan'
SEGMENT_COLUMN = 'SIPP_Segmen Peserta'


def _norm_col(s: str) -> str:
    """Normalisasi header kolom agar matching tidak peduli huruf besar/kecil & spasi."""
    return ' '.join(str(s or '').strip().lower().replace('\u00a0', ' ').split())


def _pick_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    """Ambil nama kolom asli di df yang cocok dengan salah satu alias (case-insensitive)."""
    wanted = {_norm_col(a) for a in aliases}
    for c in df.columns:
        if _norm_col(c) in wanted:
            return str(c)
    return None


def resolve_input_columns(df: pd.DataFrame) -> dict[str, str | None]:
    """Resolve kolom input dari Excel dengan banyak alias.

    Semua matching tidak peduli huruf besar/kecil.
    """
    return {
        'name': _pick_column(df, ["nama anggota keluarga", "nama", "nama siswa", "nama peserta", "nama lengkap", "NAMA LENGKAP"]),
        'gender': _pick_column(df, ["jk", "l p", "l/p", "jenis kelamin", "kelamin", "GENDER"]),
        'birthplace': _pick_column(df, ["tempat lahir", "tmpt lahir", "kota lahir", "TMP LAHIR"]),
        'birthdate': _pick_column(df, ["tanggal lahir", "tgl lahir", "lahir", "TTL"]),
        'nik': _pick_column(df, ["nik", "no ktp", "nomor ktp", "no. ktp", "no nik", "no. nik", "nik/kitas/kitap", "nomor induk kependudukan"]),
        'kk': _pick_column(df, ["kode keluarga", "no kk", "no. kk", "nomor kk", "kk", "nomor kartu keluarga", "kartu keluarga", "no kartu keluarga", "no. kartu keluarga", "no.kk"]),
        'address': _pick_column(df, ["alamat", "alamat domisili", "alamat rumah", "alamat tempat tinggal", "alamat lengkap"]),
        'kecamatan': _pick_column(df, ["kecamatan", "kec", "kecamatan domisili", "nama kecamatan"]),
        'desa': _pick_column(df, ["kelurahan", "desa", "desa/kel", "desa/kelurahan", "desa/kel.", "nama desa", "desa kelurahan", "kelurahan desa"]),
        # fallback TEMPAT LAHIR jika tidak ada
        'kab_kota_fallback': _pick_column(df, ['KAB/KOTA', 'KABUPATEN/KOTA', 'KABUPATEN', 'KOTA', 'KABUPATEN KOTA']),
    }


def clean_digits(value: str) -> str | None:
    s = (value or '').strip()
    if not s or s.lower() == 'nan':
        return None
    s = ''.join(ch for ch in s if ch.isdigit())
    return s if s else None


def map_bpjs_status_to_option(value: str) -> str:
    v = (value or '').strip().lower()
    return 'Aktif' if v == 'aktif' else 'Non Aktif'


def map_segmen_to_option(value: str) -> str:
    v = (value or '').strip().upper()
    if v in ['PBPU DAN BP PEMERINTAH DAERAH', 'PBI JAMINAN KESEHATAN']:
        return 'PBI'
    return 'Non PBI'


def map_detail_segmen_bpjs_pbi(sipp_segmen_peserta: str) -> str | None:
    """Mapping sumber `SIPP_Segmen Peserta` -> pilihan Detail Segmen BPJS saat Segmen BPJS=PBI."""
    v = (sipp_segmen_peserta or '').strip().upper()
    if v == 'PBPU DAN BP PEMERINTAH DAERAH':
        return 'PBI APBD'
    if v == 'PBI JAMINAN KESEHATAN':
        return 'PBI JK (APBN)'
    return None


def detect_pensiun_keyword(row: pd.Series, *, name_col: str) -> bool:
    """Cek semua kolom selain nama: jika ada kata kunci pensiun/pensiunan maka True.

    Catatan: pengecekan dilakukan case-insensitive karena sumber data bisa bervariasi.
    """
    keywords = ['pensiun', 'pensiunan']
    for col_name, val in row.items():
        if str(col_name).strip() == name_col:
            continue
        s = str(val or '').strip().lower()
        if not s or s == 'nan':
            continue
        if any(k in s for k in keywords):
            return True
    return False


def row_has_ppu_keywords_case_insensitive(row: pd.Series, *, name_col: str) -> bool:
    """Deteksi PPU berbasis kata kunci case-insensitive pada semua kolom selain nama.

    Kata kunci: 'PNS', 'ASN', 'TNI', 'POLRI' (tidak peduli huruf besar/kecil).
    """
    keywords = ['pns', 'asn', 'tni', 'polri']
    for col_name, val in row.items():
        if str(col_name).strip() == name_col:
            continue
        s = str(val or '').strip().lower()
        if not s or s == 'nan':
            continue
        if any(k in s for k in keywords):
            return True
    return False


def map_detail_segmen_bpjs_non_pbi(
    sipp_segmen_peserta: str,
    row_has_pensiun: bool,
    row_has_ppu_kw: bool,
) -> str:
    """Mapping Detail Segmen BPJS saat Segmen BPJS=Non PBI.

    Aturan:
    - Jika ada kata kunci pensiun/pensiunan pada salah satu kolom (selain nama) => 'Pensiunan'
    - Jika ada kata kunci PPU (case-insensitive) pada salah satu kolom (selain nama):
      'PNS' / 'ASN' / 'TNI' / 'POLRI' => 'PPU'
    - Selain itu => 'PBPU (Mandiri)'
    """
    if row_has_pensiun:
        return 'Pensiunan'

    if row_has_ppu_kw:
        return 'PPU'

    # fallback: jika kolom segmen peserta ada dari sumber (case-insensitive), tetap bisa dianggap PPU
    v = (sipp_segmen_peserta or '').strip().lower()
    if any(k in v for k in ['pns', 'asn', 'tni', 'polri']):
        return 'PPU'

    return 'PBPU (Mandiri)'


def _norm(s: str) -> str:
    return ' '.join((s or '').strip().lower().split())


def _clean_kecamatan_input_for_lookup(kec_raw: str) -> str:
    """Normalisasi kecamatan untuk proses lookup kamus.

    Mengadopsi kasus-kasus ambigu dari `auto_input_simpus_fix.py`.
    """
    k = str(kec_raw or '').strip()
    if not k:
        return ''

    low = k.lower().strip()

    # Variasi STM Hilir/Hulu
    if low in {'stm hilir', 's.t.m hilir', 'st m hilir', 'stmhilir'}:
        return 'SINEMBAH TANJUNG MUDA HILIR'
    if low in {'stm hulu', 's.t.m hulu', 'st m hulu', 'stmhulu'}:
        return 'SINEMBAH TANJUNG MUDA HULU'

    # Variasi SIBIRU-BIRU
    up = str(k).strip().upper().replace('-', ' ')
    up = re.sub(r"\s+", " ", up).strip()
    if 'SIBIRU BIRU' in up or 'SI BIRU' in up:
        return 'BIRU BIRU'

    return k


def _clean_desa_kelurahan_for_lookup(name: str) -> str:
    """Normalisasi desa/kelurahan untuk lookup kamus.

    Mengadopsi pendekatan dari `auto_input_simpus_fix.py`:
    - buang kata DESA/KEL/KELURAHAN/DS.
    - buang tanda baca umum
    - rapikan spasi

    Catatan: ini untuk lookup saja, bukan untuk tampilan UI.
    """
    s = str(name or '').strip()
    if not s:
        return ''

    # Alias spesifik untuk kasus yang sering muncul di sumber data
    # (sumber menyebut lebih detail, kamus hanya punya bentuk lebih singkat)
    #
    # NOTE:
    # Jangan terlalu agresif memendekkan nama desa/kel.
    # Contoh nyata: SIMPUS/kamus memakai 'Klambir Lima Kebun' (Hamparan Perak),
    # tapi sumber Excel bisa menulis 'KELAMBIR LIMA KEBUN'.
    # Kalau dipaksa jadi 'KELAMBIR' justru salah kecamatan (bisa nyasar ke Pantai Labu).
    alias_map = {
        # ejaan sumber -> ejaan kamus
        'KELAMBIR LIMA KEBUN': 'KLAMBIR LIMA KEBUN',
    }
    up0 = re.sub(r"\s+", " ", s.replace('\u00a0', ' ')).strip().upper()
    if up0 in alias_map:
        s = alias_map[up0]

    # rapikan NBSP
    s = s.replace('\u00a0', ' ')

    # buang prefix yang sering muncul
    s = re.sub(r"\bDESA\s*KEL\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\bDESA/KEL\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\bDESA\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\bKELURAHAN\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\bKEL\.?\b", " ", s, flags=re.IGNORECASE)

    # ganti '-' jadi spasi, buang karakter aneh
    s = s.replace('-', ' ')
    s = re.sub(r"[^0-9a-zA-Z ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()

    return s


def _kecamatan_value_for_ui(kec_raw_or_resolved: str) -> str:
    """Nilai yang diketik ke UI untuk field Kecamatan.

    Khusus SIBIRU-BIRU: UI lebih mudah jika diketik 'BIRU' (bukan 'BIRU BIRU').
    """
    v = str(_clean_kecamatan_input_for_lookup(kec_raw_or_resolved) or '').strip()
    if not v:
        return v

    up = v.upper()
    if 'BIRU BIRU' in up:
        return 'BIRU'

    return v


def _norm_wilayah_key(s: str) -> str:
    """Normalisasi ringan untuk kecamatan/desa sebelum lookup kamus.

    Mengatasi variasi umum: spasi ganda, NBSP, prefix 'KEL/DESA', kapitalisasi.
    """
    s = (s or '').strip()
    if not s:
        return s
    s = s.replace('\u00a0', ' ')
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"^(KEL\.?|KELURAHAN|DESA)\s+", "", s, flags=re.IGNORECASE).strip()
    return s


_ROMAN_TO_INT: dict[str, int] = {
    'i': 1,
    'ii': 2,
    'iii': 3,
    'iv': 4,
    'v': 5,
    'vi': 6,
    'vii': 7,
    'viii': 8,
    'ix': 9,
    'x': 10,
}
_INTWORD_TO_INT: dict[str, int] = {
    'satu': 1,
    'dua': 2,
    'tiga': 3,
    'empat': 4,
    'lima': 5,
    'enam': 6,
    'tujuh': 7,
    'delapan': 8,
    'sembilan': 9,
    'sepuluh': 10,
}


def _expand_numeric_variants(name: str) -> list[str]:
    """Buat variasi nama dengan konversi angka romawi <-> kata.

    Contoh: 'KELAMBIR LIMA KEBUN' <-> 'KELAMBIR V KEBUN'.
    """
    base = _norm_wilayah_key(name)
    if not base:
        return []

    tokens = base.split()
    variants: set[str] = {base}

    # roman -> kata
    for i, t in enumerate(tokens):
        t_clean = re.sub(r"[^0-9a-zA-Z]", "", t).lower()
        if t_clean in _ROMAN_TO_INT:
            n = _ROMAN_TO_INT[t_clean]
            word = next((w for w, nn in _INTWORD_TO_INT.items() if nn == n), None)
            if word:
                new_tokens = tokens[:]
                new_tokens[i] = word
                variants.add(' '.join(new_tokens))

    # kata -> roman
    for i, t in enumerate(tokens):
        t_clean = re.sub(r"[^0-9a-zA-Z]", "", t).lower()
        if t_clean in _INTWORD_TO_INT:
            n = _INTWORD_TO_INT[t_clean]
            roman = next((r for r, nn in _ROMAN_TO_INT.items() if nn == n), None)
            if roman:
                new_tokens = tokens[:]
                new_tokens[i] = roman.upper()
                variants.add(' '.join(new_tokens))

    return sorted(variants)


def resolve_wilayah_with_fallback(idx: dict, kecamatan: str, desa: str):
    """Resolve wilayah dengan beberapa kandidat normalisasi.

    Urutan:
    1) coba apa adanya
    2) coba normalisasi prefix/spasi
    3) coba variasi angka (LIMA <-> V, dst)

    Catatan: kecamatan juga akan dinormalisasi untuk kasus STM/SIBIRU-BIRU.
    """
    kecamatan = _clean_kecamatan_input_for_lookup(kecamatan)
    desa = _clean_desa_kelurahan_for_lookup(desa)

    kec_raw = (kecamatan or '').strip()
    desa_raw = (desa or '').strip()

    # 1) raw
    cands = resolve_wilayah(idx, kec_raw, desa_raw)
    if cands:
        return cands

    # 2) normalisasi ringan
    kec_norm = _norm_wilayah_key(kec_raw)
    desa_norm = _norm_wilayah_key(desa_raw)
    cands = resolve_wilayah(idx, kec_norm, desa_norm)
    if cands:
        return cands

    # 2b) variasi desa/kel: beberapa desa di kamus punya awalan kata (contoh: 'DAGANG KELAMBIR')
    # Jika sumber hanya tulis 'KELAMBIR', coba juga 'DAGANG KELAMBIR'.
    if _norm(desa_norm) == 'kelambir':
        for alt in ['Dagang Kelambir']:
            cands = resolve_wilayah(idx, kec_norm or kec_raw, alt)
            if cands:
                return cands

    # 3) variasi numeric
    kec_variants = _expand_numeric_variants(kec_raw) or [kec_norm or kec_raw]
    desa_variants = _expand_numeric_variants(desa_raw) or [desa_norm or desa_raw]

    tried: set[tuple[str, str]] = set()
    for k in kec_variants:
        for d in desa_variants:
            kk = _norm_wilayah_key(k)
            dd = _norm_wilayah_key(d)
            key = (_norm(kk), _norm(dd))
            if key in tried:
                continue
            tried.add(key)
            cands = resolve_wilayah(idx, kk, dd)
            if cands:
                return cands

    return []


def choose_mui_autocomplete(label_text: str, value_to_choose: str) -> None:
    """Pilih nilai pada MUI Autocomplete.

    Disesuaikan dengan struktur SIMPUS pada modal `#modalTambahData`.

    Strategi:
    - Scope pencarian hanya di dalam modal `#modalTambahData`
    - Cari input berdasarkan label text (label diikuti div lalu input)
    - Ketik value
    - Jika list option muncul: pilih yang *exact match* (case-insensitive) terhadap value_to_choose.
      Kalau tidak ada exact, fallback ARROWDOWN+ENTER.

    Ini penting untuk kasus seperti 'RIAU' (dropdown menampilkan 'KEPULAUAN RIAU' dan 'RIAU').
    """
    from selenium.webdriver.common.keys import Keys
    from selenium.common.exceptions import StaleElementReferenceException

    if not value_to_choose:
        raise ValueError(f"Nilai target kosong untuk field {label_text}")

    modal = wait.until(EC.presence_of_element_located((By.ID, "modalTambahData")))

    xpath_input = (
        f".//label[contains(normalize-space(), '{label_text}')]/following-sibling::div//input[@role='combobox' or @type='text' or @type='search' or not(@type)]"
    )

    target_norm = _norm(value_to_choose)

    last_err: Exception | None = None
    for _ in range(3):
        try:
            input_el = WebDriverWait(modal, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_input)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_el)
            input_el.click()
            time.sleep(0.2)
            input_el.clear()
            time.sleep(0.2)
            input_el.send_keys(value_to_choose)
            time.sleep(0.6)

            # Coba pilih exact match dari listbox jika ada
            try:
                listbox = WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.XPATH, "//ul[@role='listbox']"))
                )
                options = listbox.find_elements(By.XPATH, ".//li[@role='option']")
                exact = [o for o in options if _norm(o.text) == target_norm]
                if exact:
                    exact[0].click()
                    time.sleep(0.4)
                    return
            except Exception:
                # listbox tidak muncul cepat, lanjut fallback
                pass

            # Fallback: pilih item pertama
            input_el.send_keys(Keys.ARROW_DOWN)
            input_el.send_keys(Keys.ENTER)
            time.sleep(0.6)
            return
        except StaleElementReferenceException as e:
            last_err = e
            time.sleep(0.5)
        except Exception as e:
            last_err = e
            time.sleep(0.5)

    raise RuntimeError(f"Gagal set MUI Autocomplete '{label_text}' -> '{value_to_choose}': {last_err}")


def load_wilayah_dict() -> tuple[dict, list]:
    records = load_wilayah_csv(wilayah_csv)
    idx = build_wilayah_index(records)
    return idx, records


# Load referensi wilayah sekali
WILAYAH_IDX, _WILAYAH_RECORDS = load_wilayah_dict()


def clean_digits(value: str) -> str | None:
    s = (value or '').strip()
    if not s or s.lower() == 'nan':
        return None
    s = ''.join(ch for ch in s if ch.isdigit())
    return s if s else None


def map_bpjs_status_to_option(value: str) -> str:
    v = (value or '').strip().lower()
    return 'Aktif' if v == 'aktif' else 'Non Aktif'


def map_segmen_to_option(value: str) -> str:
    v = (value or '').strip().upper()
    if v in ['PBPU DAN BP PEMERINTAH DAERAH', 'PBI JAMINAN KESEHATAN']:
        return 'PBI'
    return 'Non PBI'


def map_detail_segmen_bpjs_pbi(sipp_segmen_peserta: str) -> str | None:
    """Mapping sumber `SIPP_Segmen Peserta` -> pilihan Detail Segmen BPJS saat Segmen BPJS=PBI."""
    v = (sipp_segmen_peserta or '').strip().upper()
    if v == 'PBPU DAN BP PEMERINTAH DAERAH':
        return 'PBI APBD'
    if v == 'PBI JAMINAN KESEHATAN':
        return 'PBI JK (APBN)'
    return None


def detect_pensiun_keyword(row: pd.Series, *, name_col: str) -> bool:
    """Cek semua kolom selain nama: jika ada kata kunci pensiun/pensiunan maka True.

    Catatan: pengecekan dilakukan case-insensitive karena sumber data bisa bervariasi.
    """
    keywords = ['pensiun', 'pensiunan']
    for col_name, val in row.items():
        if str(col_name).strip() == name_col:
            continue
        s = str(val or '').strip().lower()
        if not s or s == 'nan':
            continue
        if any(k in s for k in keywords):
            return True
    return False


def row_has_ppu_keywords_case_insensitive(row: pd.Series, *, name_col: str) -> bool:
    """Deteksi PPU berbasis kata kunci case-insensitive pada semua kolom selain nama.

    Kata kunci: 'PNS', 'ASN', 'TNI', 'POLRI' (tidak peduli huruf besar/kecil).
    """
    keywords = ['pns', 'asn', 'tni', 'polri']
    for col_name, val in row.items():
        if str(col_name).strip() == name_col:
            continue
        s = str(val or '').strip().lower()
        if not s or s == 'nan':
            continue
        if any(k in s for k in keywords):
            return True
    return False


def map_detail_segmen_bpjs_non_pbi(
    sipp_segmen_peserta: str,
    row_has_pensiun: bool,
    row_has_ppu_kw: bool,
) -> str:
    """Mapping Detail Segmen BPJS saat Segmen BPJS=Non PBI.

    Aturan:
    - Jika ada kata kunci pensiun/pensiunan pada salah satu kolom (selain nama) => 'Pensiunan'
    - Jika ada kata kunci PPU (case-insensitive) pada salah satu kolom (selain nama):
      'PNS' / 'ASN' / 'TNI' / 'POLRI' => 'PPU'
    - Selain itu => 'PBPU (Mandiri)'
    """
    if row_has_pensiun:
        return 'Pensiunan'

    if row_has_ppu_kw:
        return 'PPU'

    # fallback: jika kolom segmen peserta ada dari sumber (case-insensitive), tetap bisa dianggap PPU
    v = (sipp_segmen_peserta or '').strip().lower()
    if any(k in v for k in ['pns', 'asn', 'tni', 'polri']):
        return 'PPU'

    return 'PBPU (Mandiri)'


def get_first_row_values(path: str) -> dict:
    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    cols = resolve_input_columns(df)

    # Wajib ada: nama, nik, kecamatan, desa, tanggal lahir, jenis kelamin, alamat, kk, bpjs, status bpjs, segmen
    required_keys = [
        ('NAMA', 'name'),
        ('TANGGAL LAHIR', 'birthdate'),
        ('JENIS KELAMIN', 'gender'),
        ('NIK', 'nik'),
        ('NO KK', 'kk'),
        ('ALAMAT', 'address'),
        ('KECAMATAN', 'kecamatan'),
        ('DESA/KELURAHAN', 'desa'),
    ]

    missing = [label for label, key in required_keys if not cols.get(key)]
    # TEMPAT LAHIR boleh kosong: akan diisi dari Kab/Kota jika kolom Tempat Lahir tidak ada atau nilainya kosong
    # BPJS columns (output dari open_sipp) tetap wajib
    if BPJS_COLUMN not in df.columns:
        missing.append(BPJS_COLUMN)
    if BPJS_STATUS_COLUMN not in df.columns:
        missing.append(BPJS_STATUS_COLUMN)
    if SEGMENT_COLUMN not in df.columns:
        missing.append(SEGMENT_COLUMN)

    if missing:
        raise ValueError(f"Kolom tidak ditemukan di {path}: {missing}")

    name_col = cols['name']

    for _, r in df.iterrows():
        name = str(r.get(name_col) or '').strip()
        if not name or name.lower() == 'nan':
            continue

        row_has_pensiun = detect_pensiun_keyword(r, name_col=name_col)
        row_has_ppu_kw = row_has_ppu_keywords_case_insensitive(r, name_col=name_col)

        # TEMPAT LAHIR fallback: pakai Kab/Kota jika kolom Tempat Lahir tidak ada atau nilainya kosong
        birthplace_col = cols.get('birthplace')
        birthplace_val = str(r.get(birthplace_col) or '').strip() if birthplace_col else ''
        if (not birthplace_val) or birthplace_val.lower() == 'nan':
            kab_col = cols.get('kab_kota_fallback')
            if kab_col:
                birthplace_val = str(r.get(kab_col) or '').strip()

        return {
            'name': name,
            'gender': str(r.get(cols['gender']) or '').strip(),
            'birthplace': birthplace_val,
            'birthdate': str(r.get(cols['birthdate']) or '').strip(),
            'nik': str(r.get(cols['nik']) or '').strip(),
            'kk': str(r.get(cols['kk']) or '').strip(),
            'address': str(r.get(cols['address']) or '').strip(),
            'nama_kecamatan': str(r.get(cols['kecamatan']) or '').strip(),
            'nama_desa': str(r.get(cols['desa']) or '').strip(),
            'bpjs': str(r.get(BPJS_COLUMN) or '').strip(),
            'bpjs_status': str(r.get(BPJS_STATUS_COLUMN) or '').strip(),
            'segmen': str(r.get(SEGMENT_COLUMN) or '').strip(),
            'row_has_pensiun': row_has_pensiun,
            'row_has_ppu_kw': row_has_ppu_kw,
        }

    raise ValueError('Tidak ada baris data valid di Excel.')


def map_gender_to_option(value: str) -> str | None:
    v = (value or '').strip().lower()
    if not v or v == 'nan':
        return None
    # Dari excel sering berupa 1/2 atau L/P
    if v in ['1', 'l', 'lk', 'laki', 'laki-laki', 'male']:
        return 'Laki-Laki'
    if v in ['2', 'p', 'pr', 'perempuan', 'female', 'wanita']:
        return 'Perempuan'
    # Jika sudah persis
    if 'laki' in v:
        return 'Laki-Laki'
    if 'perempuan' in v or 'wanita' in v:
        return 'Perempuan'
    return None


def normalize_date_yyyy_mm_dd(value: str) -> str | None:
    """Konversi value tanggal (mis. 'YYYY-MM-DD' atau 'YYYY-MM-DD 00:00:00') ke 'YYYY-MM-DD'."""
    s = (value or '').strip()
    if not s or s.lower() == 'nan':
        return None
    if len(s) >= 10:
        return s[:10]
    return None


def set_date_input_js(date_input, yyyy_mm_dd: str) -> None:
    """Set <input type=date> secara stabil lintas format (mm/dd vs dd/mm) menggunakan JS.

    Catatan: beberapa browser/PC bisa geser 1 hari karena timezone saat pakai new Date(y,m,d).
    Untuk mencegah off-by-one, kita set pakai UTC.
    """
    driver.execute_script(
        """
        const el = arguments[0];
        const v = arguments[1];
        if (!v) return;
        const parts = v.split('-');
        const y = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10) - 1;
        const d = parseInt(parts[2], 10);

        // Buat tanggal pada UTC tengah hari untuk menghindari pergeseran timezone
        const utcDate = new Date(Date.UTC(y, m, d, 12, 0, 0));
        el.valueAsDate = utcDate;

        // fallback kalau browser tidak konsisten dengan valueAsDate
        if (el.value !== v) {
            el.value = v;
        }

        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        """,
        date_input,
        yyyy_mm_dd,
    )


def _strip_admin_prefix_kab_kota(name: str) -> str:
    """Return core Kab/Kota name without leading admin prefixes.

    Examples:
    - 'KAB. DELI SERDANG' -> 'DELI SERDANG'
    - 'KABUPATEN DELI SERDANG' -> 'DELI SERDANG'
    - 'KOTA MEDAN' -> 'MEDAN'
    """
    s = (name or '').strip()
    if not s:
        return s

    up = s.upper().strip()
    prefixes = [
        'KABUPATEN ',
        'KAB. ',
        'KAB ',
        'KOTA ADMINISTRASI ',
        'KOTA ',
    ]
    for p in prefixes:
        if up.startswith(p):
            return s[len(p) :].strip()
    return s


def _normalize_provinsi_for_ui(name: str) -> str:
    """Normalize province name for SIMPUS autocomplete."""
    s = (name or '').strip().strip('"').strip("'").strip()
    if not s:
        return s

    up = s.upper().strip()

    # Common aliases
    if up in {'KEPRI', 'KEP. RIAU', 'KEPULAUAN RIAU'}:
        return 'KEPULAUAN RIAU'

    # Ensure 'RIAU' remains exactly 'RIAU'
    if up == 'RIAU':
        return 'RIAU'

    return s


def _normalize_desa_kelurahan_for_ui(name: str) -> str:
    """Normalize desa/kelurahan name for SIMPUS autocomplete.

    Handles known spelling variations.
    """
    s = (name or '').strip()
    if not s:
        return s

    up = s.upper().strip()

    # Known aliases (extend as needed)
    aliases = {
        'KLUMPANG KEBUN': 'KLUMPANG KEBON',
        # SIMPUS memakai 'KEBON' (bukan 'KEBUN')
        'KLAMBIR LIMA KEBUN': 'KLAMBIR LIMA KEBON',
        # SIMPUS memakai bentuk 'SUNGAI' (bukan 'SEI')
        'SEI BAHARU': 'SUNGAI BAHARU',
    }
    if up in aliases:
        return aliases[up]

    # Kasus khusus: di kamus ada campuran penulisan 'Tandem' vs 'Tandam',
    # dan SIMPUS menampilkan bentuk:
    # - 'KAMPUNG TANDAM HULU SATU'
    # - 'TANDAM HULU DUA'
    # - 'TANDAM HILIR SATU'
    # - 'TANDAM HILIR DUA'
    #
    # Maka kita ubah *yang diketik ke UI* ke bentuk yang mendekati opsi SIMPUS.
    # (Pemilihan akhir tetap dilakukan oleh choose_mui_autocomplete.)
    if up.startswith('TANDEM ') or up.startswith('TANDAM '):
        # Samakan dasar ejaan
        base = up.replace('TANDEM', 'TANDAM', 1)

        # Ubah angka romawi ke kata
        base = re.sub(r"\bI\b", "SATU", base)
        base = re.sub(r"\bII\b", "DUA", base)
        base = re.sub(r"\bIII\b", "TIGA", base)

        # Khusus HULU SATU: di SIMPUS biasanya pakai 'KAMPUNG TANDAM HULU SATU'
        if base == 'TANDAM HULU SATU':
            return 'KAMPUNG TANDAM HULU SATU'

        return base

    return s


def iter_valid_rows(path: str, start_num: int, end_num: int):
    """Yield dict of row values for each valid person row in the Excel."""
    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    cols = resolve_input_columns(df)

    # Wajib ada: nama, nik, kecamatan, desa, tanggal lahir, jenis kelamin, alamat, kk, bpjs, status bpjs, segmen
    required_keys = [
        ('NAMA', 'name'),
        ('TANGGAL LAHIR', 'birthdate'),
        ('JENIS KELAMIN', 'gender'),
        ('NIK', 'nik'),
        ('NO KK', 'kk'),
        ('ALAMAT', 'address'),
        ('KECAMATAN', 'kecamatan'),
        ('DESA/KELURAHAN', 'desa'),
    ]

    missing = [label for label, key in required_keys if not cols.get(key)]
    # TEMPAT LAHIR boleh kosong: akan diisi dari Kab/Kota jika kolom Tempat Lahir tidak ada atau nilainya kosong
    # BPJS columns (output dari open_sipp) tetap wajib
    if BPJS_COLUMN not in df.columns:
        missing.append(BPJS_COLUMN)
    if BPJS_STATUS_COLUMN not in df.columns:
        missing.append(BPJS_STATUS_COLUMN)
    if SEGMENT_COLUMN not in df.columns:
        missing.append(SEGMENT_COLUMN)

    if missing:
        raise ValueError(f"Kolom tidak ditemukan di {path}: {missing}")

    name_col = cols['name']

    # start_num is 1-based data index (row 1 is df.iloc[0])
    df_sliced = df.iloc[start_num - 1 : end_num]

    for _, r in df_sliced.iterrows():
        name = str(r.get(name_col) or '').strip()
        if not name or name.lower() == 'nan':
            continue

        row_has_pensiun = detect_pensiun_keyword(r, name_col=name_col)
        row_has_ppu_kw = row_has_ppu_keywords_case_insensitive(r, name_col=name_col)

        # TEMPAT LAHIR fallback: pakai Kab/Kota jika kolom Tempat Lahir tidak ada atau nilainya kosong
        birthplace_col = cols.get('birthplace')
        birthplace_val = str(r.get(birthplace_col) or '').strip() if birthplace_col else ''
        if (not birthplace_val) or birthplace_val.lower() == 'nan':
            kab_col = cols.get('kab_kota_fallback')
            if kab_col:
                birthplace_val = str(r.get(kab_col) or '').strip()

        yield {
            'name': name,
            'gender': str(r.get(cols['gender']) or '').strip(),
            'birthplace': birthplace_val,
            'birthdate': str(r.get(cols['birthdate']) or '').strip(),
            'nik': str(r.get(cols['nik']) or '').strip(),
            'kk': str(r.get(cols['kk']) or '').strip(),
            'address': str(r.get(cols['address']) or '').strip(),
            'nama_kecamatan': str(r.get(cols['kecamatan']) or '').strip(),
            'nama_desa': str(r.get(cols['desa']) or '').strip(),
            'bpjs': str(r.get(BPJS_COLUMN) or '').strip(),
            'bpjs_status': str(r.get(BPJS_STATUS_COLUMN) or '').strip(),
            'segmen': str(r.get(SEGMENT_COLUMN) or '').strip(),
            'row_has_pensiun': row_has_pensiun,
            'row_has_ppu_kw': row_has_ppu_kw,
        }


def open_tambah_data_modal() -> None:
    tambah_btn = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//button[contains(@class,'btn') and contains(@data-bs-target,'#modalTambahData') and normalize-space()='TAMBAH DATA']",
            )
        )
    )
    tambah_btn.click()
    wait.until(EC.visibility_of_element_located((By.ID, 'modalTambahData')))


def submit_tambah_and_wait_close() -> None:
    # Klik tombol submit TAMBAH di dalam modal
    submit_btn = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//div[@id='modalTambahData']//button[@type='submit' and contains(@class,'btn') and normalize-space()='TAMBAH']",
            )
        )
    )
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", submit_btn)
    submit_btn.click()

    # Tunggu modal tertutup (hilang)
    wait.until(EC.invisibility_of_element_located((By.ID, 'modalTambahData')))
    time.sleep(0.5)


def close_modal_if_open() -> None:
    try:
        modal = driver.find_element(By.ID, 'modalTambahData')
        if not modal.is_displayed():
            return
    except Exception:
        return

    # Coba klik button Close atau icon X
    for xp in [
        "//div[@id='modalTambahData']//button[contains(.,'CLOSE') or contains(.,'Close') or contains(.,'TUTUP') or contains(.,'Tutup')]",
        "//div[@id='modalTambahData']//button[contains(@class,'btn-close') or @aria-label='Close']",
    ]:
        try:
            btn = driver.find_element(By.XPATH, xp)
            btn.click()
            wait.until(EC.invisibility_of_element_located((By.ID, 'modalTambahData')))
            time.sleep(0.3)
            return
        except Exception:
            pass


# Buka halaman login
driver.get(login_url)

print('Silakan login manual di browser. Script akan mendeteksi saat sudah masuk /home, lalu buka /pasien.')

# Tunggu login terdeteksi lewat URL
timeout_seconds = 300  # 5 menit
poll_interval_seconds = 1
start = time.time()

try:
    while True:
        current_url = driver.current_url or ''
        if home_url_fragment in current_url:
            break
        if time.time() - start > timeout_seconds:
            raise TimeoutError('Timeout: belum terdeteksi login (belum masuk /home).')
        time.sleep(poll_interval_seconds)

    # Setelah login terdeteksi, buka halaman pasien
    driver.get(pasien_url)

    failed: list[dict] = []
    success_count = 0

    with open('failed-log.txt', 'a', encoding='utf-8') as f_log:
        f_log.write(f"\n\n--- Sesi Baru: {time.strftime('%Y-%m-%d %H:%M:%S')} ---\n")

    for idx, row_values in enumerate(iter_valid_rows(input_xlsx, NOMOR_AWAL, NOMOR_AKHIR), start=NOMOR_AWAL):
        try:
            print(f"\n=== INPUT {idx} ===")
            open_tambah_data_modal()

            # Tunggu modal muncul dan field siap
            nama_value = row_values['name']
            gender_value = row_values['gender']
            tempat_lahir_value = row_values['birthplace']
            tanggal_lahir_value = normalize_date_yyyy_mm_dd(row_values['birthdate'])
            nik_value = clean_digits(row_values.get('nik'))
            kk_value = clean_digits(row_values.get('kk'))
            bpjs_value = clean_digits(row_values.get('bpjs'))
            bpjs_status_value = row_values.get('bpjs_status')
            bpjs_status_option = map_bpjs_status_to_option(bpjs_status_value)
            segmen_value = row_values.get('segmen')
            segmen_option = map_segmen_to_option(segmen_value)
            row_has_pensiun = bool(row_values.get('row_has_pensiun'))
            row_has_ppu_kw = bool(row_values.get('row_has_ppu_kw'))
            alamat_value = str(row_values.get('address') or '').strip()
            nama_kecamatan_value = str(row_values.get('nama_kecamatan') or '').strip()
            nama_desa_value = str(row_values.get('nama_desa') or '').strip()

            # Hitung opsi Detail Segmen BPJS (dibutuhkan saat pilih dropdown)
            detail_segmen_pbi_option = map_detail_segmen_bpjs_pbi(segmen_value)
            detail_segmen_non_pbi_option = map_detail_segmen_bpjs_non_pbi(
                segmen_value,
                row_has_pensiun=row_has_pensiun,
                row_has_ppu_kw=row_has_ppu_kw,
            )

            # Terapkan normalisasi khusus kecamatan (STM / SIBIRU-BIRU)
            nama_kecamatan_value = _clean_kecamatan_input_for_lookup(nama_kecamatan_value)
            # Terapkan normalisasi desa/kelurahan (DESA/KEL/KELURAHAN, tanda baca)
            nama_desa_value = _clean_desa_kelurahan_for_lookup(nama_desa_value)

            # Resolve prov/kab/kec/kel dari kamus wilayah
            candidates = resolve_wilayah_with_fallback(WILAYAH_IDX, nama_kecamatan_value, nama_desa_value)
            if not candidates:
                raise ValueError(
                    f"Wilayah tidak ditemukan di kamus untuk kecamatan='{nama_kecamatan_value}', desa='{nama_desa_value}'.",
                )
            if len(candidates) > 1:
                preview = '; '.join(
                    [
                        f"{c.provinsi} / {c.kabupaten_kota} / {c.kecamatan} / {c.desa_kelurahan}"
                        for c in candidates[:5]
                    ]
                )
                raise ValueError(
                    "Wilayah ambigu (lebih dari 1 kandidat). "
                    f"kecamatan='{nama_kecamatan_value}', desa='{nama_desa_value}'. Kandidat (contoh): {preview}",
                )

            wilayah = candidates[0]
            prov_value = wilayah.provinsi
            kab_value = wilayah.kabupaten_kota
            kec_value = wilayah.kecamatan
            kel_value = wilayah.desa_kelurahan

            prov_value_for_ui = _normalize_provinsi_for_ui(prov_value)
            kab_value_for_ui = _strip_admin_prefix_kab_kota(kab_value)
            kel_value_for_ui = _normalize_desa_kelurahan_for_ui(kel_value)

            # Untuk UI, beberapa kecamatan lebih stabil pakai nilai khusus (mis. SIBIRU-BIRU -> 'BIRU')
            kec_value_for_ui = _kecamatan_value_for_ui(kec_value)

            # Log ringkas untuk user: identitas + wilayah hasil resolve
            print(
                "[AUTOFILL] "
                f"NIK={nik_value or '-'} | NAMA={nama_value} | "
                f"PROVINSI={prov_value} | KAB/KOTA={kab_value} | KECAMATAN={kec_value} | DESA/KEL={kel_value}"
            )

            # NOTE: Wilayah + alamat akan diisi belakangan (setelah field identitas/BPJS), agar lebih stabil.

            nama_input = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//div[@id='modalTambahData']//input[@type='text' and @placeholder='Nama' and (not(@disabled) or @disabled='false')]",
                    )
                )
            )
            nama_input.click(); nama_input.clear(); nama_input.send_keys(nama_value)

            # Set Jenis Kelamin
            gender_option = map_gender_to_option(gender_value)
            if gender_option:
                gender_select_el = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//div[@id='modalTambahData']//select[contains(@class,'form-control') and .//option[normalize-space()='Jenis Kelamin']]",
                        )
                    )
                )
                Select(gender_select_el).select_by_visible_text(gender_option)

            # Tempat Lahir
            if tempat_lahir_value and tempat_lahir_value.lower() != 'nan':
                tempat_lahir_input = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//div[@id='modalTambahData']//input[@type='text' and @placeholder='Tempat Lahir' and (not(@disabled) or @disabled='false')]",
                        )
                    )
                )
                tempat_lahir_input.click(); tempat_lahir_input.clear(); tempat_lahir_input.send_keys(tempat_lahir_value)

            # Tanggal Lahir
            if tanggal_lahir_value:
                tanggal_lahir_input = wait.until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            "//div[@id='modalTambahData']//input[@type='date' and (@placeholder='Tanggal Lahir' or contains(@placeholder,'Tanggal'))]",
                        )
                    )
                )
                set_date_input_js(tanggal_lahir_input, tanggal_lahir_value)

            # NIK
            if nik_value:
                nik_input = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//div[@id='modalTambahData']//input[@placeholder='NIK' and (@type='number' or @inputmode='numeric') ]",
                        )
                    )
                )
                nik_input.click(); nik_input.clear(); nik_input.send_keys(nik_value)

            # No. KK
            if kk_value:
                kk_input = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//div[@id='modalTambahData']//input[@placeholder='No. KK' and (@type='number' or @inputmode='numeric') ]",
                        )
                    )
                )
                kk_input.click(); kk_input.clear(); kk_input.send_keys(kk_value)

            # No. BPJS
            if bpjs_value:
                bpjs_input = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//div[@id='modalTambahData']//input[@placeholder='No. BPJS' and @type='text']",
                        )
                    )
                )
                bpjs_input.click(); bpjs_input.clear(); bpjs_input.send_keys(bpjs_value)

            # Status BPJS
            status_bpjs_select_el = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//div[@id='modalTambahData']//select[contains(@class,'form-control') and .//option[normalize-space()='Status BPJS']]",
                    )
                )
            )
            Select(status_bpjs_select_el).select_by_visible_text(bpjs_status_option)

            # Segmen BPJS
            segmen_select_el = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//div[@id='modalTambahData']//select[contains(@class,'form-control') and .//option[normalize-space()='Segmen BPJS']]",
                    )
                )
            )
            Select(segmen_select_el).select_by_visible_text(segmen_option)

            # Detail Segmen BPJS
            detail_segmen_select_el = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//div[@id='modalTambahData']//select[contains(@class,'form-control') and .//option[normalize-space()='Detail Segmen BPJS']]",
                    )
                )
            )
            if segmen_option == 'PBI':
                if detail_segmen_pbi_option:
                    Select(detail_segmen_select_el).select_by_visible_text(detail_segmen_pbi_option)
            else:
                Select(detail_segmen_select_el).select_by_visible_text(detail_segmen_non_pbi_option)

            # Alamat
            if alamat_value and alamat_value.lower() != 'nan':
                alamat_input = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//div[@id='modalTambahData']//div[contains(@class,'col-md-12')][.//p[contains(@class,'example-form-small') and normalize-space()='Tulis alamat lengkap']]//input[@type='text' and @placeholder='Alamat' and contains(@class,'form-control')]",
                        )
                    )
                )
                alamat_input.click(); alamat_input.clear(); alamat_input.send_keys(alamat_value)

            # Wilayah (akhir)
            choose_mui_autocomplete('Provinsi', prov_value_for_ui)
            choose_mui_autocomplete('Kabupaten / Kota', kab_value_for_ui)
            choose_mui_autocomplete('Kecamatan', kec_value_for_ui)
            choose_mui_autocomplete('Kelurahan', kel_value_for_ui)

            # Submit TAMBAH
            submit_tambah_and_wait_close()
            success_count += 1
            print(f"✅ Sukses submit: {nama_value}")
            with open('last-success-log.txt', 'w', encoding='utf-8') as f_succ:
                f_succ.write(f"Nama: {nama_value}\nNIK: {nik_value or '-'}\nTanggal/Waktu: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        except Exception as e:
            err_msg = str(e)
            failed.append({'index': idx, 'name': row_values.get('name'), 'nik': row_values.get('nik'), 'error': err_msg})
            print(f"❌ Gagal input index={idx} nama={row_values.get('name')}: {err_msg}")
            with open('failed-log.txt', 'a', encoding='utf-8') as f_log:
                f_log.write(f"Index={idx} | NIK={row_values.get('nik') or '-'} | NAMA={row_values.get('name')} | ERROR: {err_msg}\n")
            close_modal_if_open()
            continue

    print(f"\nSELESAI. sukses={success_count}, gagal={len(failed)}")
    if failed:
        for f in failed[:20]:
            print(f"- gagal index={f['index']} nik={f.get('nik')} nama={f.get('name')}: {f.get('error')}")

    print('Tekan Ctrl+C untuk keluar.')

    while True:
        time.sleep(1)
except KeyboardInterrupt:
    driver.quit()
except Exception as e:
    print(f'Error: {e}')
    driver.quit()
