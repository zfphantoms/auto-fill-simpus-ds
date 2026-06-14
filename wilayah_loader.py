import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple


@dataclass(frozen=True)
class WilayahRecord:
    provinsi: str
    kabupaten_kota: str
    kecamatan: str
    desa_kelurahan: str


def _norm(s: str) -> str:
    return ' '.join((s or '').strip().lower().split())


def load_wilayah_csv(path: str | Path) -> List[WilayahRecord]:
    """Load referensi wilayah dari CSV.

    CSV harus punya header minimal:
    - provinsi
    - kabupaten_kota
    - kecamatan
    - desa_kelurahan

    Kolom boleh punya variasi nama (akan dicoba beberapa alias umum).
    """
    path = Path(path)
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        headers = {h.strip(): h for h in (reader.fieldnames or [])}

        def pick(*aliases: str) -> Optional[str]:
            for a in aliases:
                for h in headers:
                    if _norm(h) == _norm(a):
                        return headers[h]
            return None

        h_prov = pick('provinsi', 'nama_provinsi')
        h_kab = pick('kabupaten_kota', 'kabupaten/kota', 'kab_kota', 'nama_kabupaten_kota')
        h_kec = pick('kecamatan', 'nama_kecamatan')
        h_desa = pick('desa_kelurahan', 'kelurahan', 'desa', 'nama_desa', 'nama_kelurahan', 'nama_desa_kelurahan')

        missing = [
            ('provinsi', h_prov),
            ('kabupaten_kota', h_kab),
            ('kecamatan', h_kec),
            ('desa_kelurahan', h_desa),
        ]
        missing = [name for name, h in missing if h is None]
        if missing:
            raise ValueError(f"Header CSV wilayah tidak lengkap. Kolom hilang: {missing}")

        out: List[WilayahRecord] = []
        for row in reader:
            prov = (row.get(h_prov) or '').strip()
            kab = (row.get(h_kab) or '').strip()
            kec = (row.get(h_kec) or '').strip()
            desa = (row.get(h_desa) or '').strip()
            if not (prov and kab and kec and desa):
                continue
            out.append(WilayahRecord(provinsi=prov, kabupaten_kota=kab, kecamatan=kec, desa_kelurahan=desa))
        return out


def build_wilayah_index(records: Iterable[WilayahRecord]) -> Dict[Tuple[str, str], List[WilayahRecord]]:
    """Index by (kecamatan_norm, desa_norm) -> list of candidates."""
    idx: Dict[Tuple[str, str], List[WilayahRecord]] = {}
    for r in records:
        key = (_norm(r.kecamatan), _norm(r.desa_kelurahan))
        idx.setdefault(key, []).append(r)
    return idx


def resolve_wilayah(
    idx: Dict[Tuple[str, str], List[WilayahRecord]],
    nama_kecamatan: str,
    nama_desa: str,
) -> List[WilayahRecord]:
    """Return candidate wilayah for a given (kecamatan, desa)."""
    return idx.get((_norm(nama_kecamatan), _norm(nama_desa)), [])
