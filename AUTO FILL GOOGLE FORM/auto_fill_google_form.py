from __future__ import annotations

import argparse
import sys
import subprocess
import os
from typing import Iterable

# BASE URL Form
BASE_URL = "https://docs.google.com/forms/d/e/1FAIpQLSdtftu4p2UWL0OpuwuhGELEdMTbR5WND1JMQelJEtmLPhnLpQ/viewform"

# Query parameters berisi seluruh pasangan entry ID dan nilainya sesuai urutan terbaru
PARAMS = (
    "?entry.31303650=YA"
    "&entry.39879045=YA"
    "&entry.149992713=YA"
    "&entry.263451059=YA"
    "&entry.287515313=YA"
    "&entry.416892484=YA"
    "&entry.552278983=YA"
    "&entry.670052259=YA"
    "&entry.675244400=42320"
    "&entry.698978356=YA"
    "&entry.789089821=YA"
    "&entry.842992950=YA"
    "&entry.938147661=YA"
    "&entry.1007991142=YA"
    "&entry.1286771032=Gideon+Christ+Gilberio+Ginting"
    "&entry.1441840568=YA"
    "&entry.1453158074=YA"
    "&entry.1458868878=YA"
    "&entry.1501790328=082121451169"
    "&entry.1535995396=YA"
    "&entry.1749244526=YA"
    "&entry.1906607431=YA"
    "&entry.1942516585=YA"
)

PREFILLED_URL = BASE_URL + PARAMS


def open_chrome_with_profile(url: str, profile_dir: str) -> bool:
    """Mencari lokasi Chrome dan membukanya dengan profil tertentu."""
    # Daftar kemungkinan lokasi instalasi Chrome di Windows
    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe")
    ]
    
    # Cari path yang valid
    chrome_exe = next((path for path in chrome_paths if os.path.exists(path)), None)
    
    if not chrome_exe:
        return False
        
    try:
        # Menjalankan Chrome dengan parameter profil
        subprocess.Popen([chrome_exe, f'--profile-directory={profile_dir}', url])
        return True
    except Exception:
        return False


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Buka Google Form dengan data terisi otomatis.")
    parser.parse_args(list(argv) if argv is not None else None)

    print("=" * 60)
    print("GOOGLE FORM AUTO-FILL LINK")
    print("=" * 60)

    # Menargetkan profil kedua (Gideon C. G. Ginting)
    target_profile = "Profile 1" 
    
    success = open_chrome_with_profile(PREFILLED_URL, target_profile)
    
    if success:
        print(f"\n[V] Form berhasil dibuka di Chrome menggunakan profil Gideon C. G. Ginting!")
        print("[!] Catatan: Silakan cek Terminal ID (42320), unggah bukti foto, lalu klik 'Kirim'.")
    else:
        print("\n[!] Gagal membuka Chrome secara otomatis.")
        print("Silakan salin URL di bawah ini secara manual:\n")
        print(PREFILLED_URL)

    print("-" * 60)
    input("Tekan Enter di sini jika sudah selesai untuk menutup script...")
    return 0


if __name__ == "__main__":
    sys.exit(main())