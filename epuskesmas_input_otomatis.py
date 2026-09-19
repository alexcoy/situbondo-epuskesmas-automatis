# -*- coding: utf-8 -*-
"""
=============================================================================
 OTOMASI INPUT PASIEN BPJS -> EPUSKESMAS (situbondo.epuskesmas.id)
=============================================================================
Cara pakai singkat (lihat PANDUAN.md untuk detail lengkap):

    1. pip install -r requirements.txt
    2. playwright install chromium
    3. Sesuaikan bagian "KONFIGURASI" di bawah (path Excel, dsb).
    4. Jalankan:  python epuskesmas_input_otomatis.py
    5. Browser akan terbuka -> login MANUAL (termasuk captcha) -> tekan ENTER
       di terminal saat sudah masuk ke halaman utama/dashboard.
    6. Script akan memproses pasien satu per satu secara otomatis.

CATATAN PENTING:
    Saya (Claude) tidak punya akses internet untuk membuka situs
    situbondo.epuskesmas.id secara langsung, sehingga nama tombol / field
    di bawah ini ditulis berdasarkan istilah yang Anda berikan
    (mis. "Pendaftaran", "Simpan Pendaftaran", "Kondisi Stabil", dst).
    Kemungkinan besar akan ada 1-2 selector yang perlu disesuaikan pada
    percobaan pertama. Jalankan dulu dengan DEBUG_MODE = True dan
    JUMLAH_PASIEN_TES = 2 sebelum menjalankan untuk 470 data sekaligus.

    Saat terjadi error dan DEBUG_MODE=True, script akan membuka
    "Playwright Inspector" agar Anda bisa lihat & perbaiki selector secara
    langsung, lalu klik "Resume" untuk melanjutkan. Beri tahu saya selector
    mana yang salah supaya saya perbaiki di scriptnya.
=============================================================================
"""

import sys
import time
import traceback
from datetime import datetime

import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout


# =============================================================================
# KONFIGURASI — SESUAIKAN BAGIAN INI
# =============================================================================

URL_LOGIN = "https://situbondo.epuskesmas.id/home"

EXCEL_INPUT = "data_pasien_bulan_ini.xlsx"          # <-- ganti nama file Excel Anda
EXCEL_OUTPUT = "hasil_input_{tanggal}.xlsx"          # laporan hasil otomatis dibuat

# Mode aman untuk uji coba pertama kali:
DEBUG_MODE = True          # True = browser lambat + pause otomatis saat error
JUMLAH_PASIEN_TES = 2       # None = proses semua data; angka = hanya proses N baris dulu
SLOW_MO_MS = 300            # jeda antar aksi (ms) supaya bisa diawasi

WAIT_TIMEOUT = 15000        # timeout tunggu elemen muncul (ms)

# --- Nilai TETAP (sama untuk semua pasien), sesuai instruksi Anda ---
NILAI_TETAP = {
    "dokter": "REVANI YUNI NAILUVAR",
    "lama_sakit_angka": "1",
    "lama_sakit_satuan": "Hari",
    "alergi": "Tidak Ada",
    "sistole": "120",
    "diastole": "80",
    "nadi": "80",
    "nafas": "20",
    "lingkar_perut": "45",
    "edukasi": "anjurkan pasien istirahat yang cukup",
    "prognosa": "Bonam",
    "skrining_visual": "Kondisi Stabil",
    "instalasi": "Rawat Jalan",
    "poli_ruang": "KLASTER 3 UMUM",
    "status_pulang": "Berobat Jalan",
}

# --- Nama kolom di file Excel Anda (ubah di sini jika header Excel berbeda) ---
KOLOM = {
    "nama": "Nama Pasien",
    "nik": "NIK",
    "no_bpjs": "No. Asuransi",
    "jenis_kelamin": "Jenis Kelamin",   # isi "L" atau "P"
    "berat_badan": "Berat Badan (Kg)",
    "tinggi_badan": "Tinggi Badan (cm)",
    "diagnosa": "Diagnosa",             # dipakai utk Keluhan Utama & ICD-X
    "tenaga_medis": "Tenaga Medis",     # dipakai utk Perawat/Bidan
}


# =============================================================================
# BACA DATA EXCEL
# =============================================================================

def cari_baris_header(ws):
    """Cari baris header dengan mendeteksi baris yang mengandung 'Nama Pasien' dan 'NIK'."""
    for r in range(1, ws.max_row + 1):
        nilai_baris = [str(ws.cell(row=r, column=c).value or "") for c in range(1, ws.max_column + 1)]
        if any("Nama Pasien" in v for v in nilai_baris) and any(v.strip() == "NIK" for v in nilai_baris):
            return r
    raise RuntimeError("Baris header (berisi 'Nama Pasien' dan 'NIK') tidak ditemukan di Excel.")


def baca_data_pasien(path_excel):
    wb = openpyxl.load_workbook(path_excel, data_only=True)
    ws = wb.active
    baris_header = cari_baris_header(ws)

    header_ke_kolom = {}
    for c in range(1, ws.max_column + 1):
        nilai = ws.cell(row=baris_header, column=c).value
        if nilai:
            header_ke_kolom[str(nilai).strip()] = c

    for key, nama_header in KOLOM.items():
        if nama_header not in header_ke_kolom:
            raise RuntimeError(
                f"Kolom '{nama_header}' (untuk '{key}') tidak ditemukan di Excel. "
                f"Kolom yang tersedia: {list(header_ke_kolom.keys())}"
            )

    daftar_pasien = []
    for r in range(baris_header + 1, ws.max_row + 1):
        nama = ws.cell(row=r, column=header_ke_kolom[KOLOM["nama"]]).value
        if not nama:
            continue  # baris kosong, lewati
        pasien = {"_baris_excel": r}
        for key, nama_header in KOLOM.items():
            pasien[key] = ws.cell(row=r, column=header_ke_kolom[nama_header]).value
        daftar_pasien.append(pasien)

    return wb, ws, daftar_pasien, header_ke_kolom


def simpan_hasil(wb, ws, header_ke_kolom, hasil_per_baris, path_output):
    kolom_status = ws.max_column + 1
    kolom_keterangan = ws.max_column + 2
    baris_header = cari_baris_header(ws)
    ws.cell(row=baris_header, column=kolom_status, value="Status Input Otomatis")
    ws.cell(row=baris_header, column=kolom_keterangan, value="Keterangan")

    for baris, (status, keterangan) in hasil_per_baris.items():
        ws.cell(row=baris, column=kolom_status, value=status)
        ws.cell(row=baris, column=kolom_keterangan, value=keterangan)

    wb.save(path_output)
    print(f"\n📄 Laporan hasil disimpan ke: {path_output}")


# =============================================================================
# LANGKAH-LANGKAH DI WEBSITE (SESUAI ALUR YANG ANDA BERIKAN)
# =============================================================================

def bersihkan_teks(nilai):
    if nilai is None:
        return ""
    return str(nilai).strip()


def cari_dan_pilih_pasien(page, pasien):
    """Langkah 3a-3c: cari berdasarkan NIK, fallback ke No. BPJS, lalu double-click nama."""
    nik = bersihkan_teks(pasien["nik"])
    no_bpjs = bersihkan_teks(pasien["no_bpjs"])
    nama = bersihkan_teks(pasien["nama"])

    kotak_pencarian = page.get_by_placeholder("Cari").first  # SESUAIKAN jika placeholder beda

    for kata_kunci in [nik, no_bpjs]:
        if not kata_kunci:
            continue
        kotak_pencarian.fill("")
        kotak_pencarian.fill(kata_kunci)
        page.wait_for_timeout(1200)  # tunggu hasil pencarian (autocomplete/grid)

        baris_hasil = page.locator(f"tr:has-text('{nama}')").first
        try:
            baris_hasil.wait_for(state="visible", timeout=4000)
            baris_hasil.dblclick()
            return True
        except PWTimeout:
            continue  # coba kata kunci berikutnya (NIK gagal -> coba No. BPJS)

    return False  # tidak ditemukan sama sekali


def isi_pendaftaran(page, pasien):
    """Langkah 3d-3k."""
    page.get_by_text("Pendaftaran", exact=True).first.click()
    page.wait_for_timeout(800)

    if bersihkan_teks(pasien["jenis_kelamin"]).upper().startswith("P"):
        page.get_by_text("Tidak", exact=True).first.click()  # status hamil/bersalin/nifas

    page.get_by_label("Tinggi Badan").fill(bersihkan_teks(pasien["tinggi_badan"]))
    page.get_by_label("Berat Badan").fill(bersihkan_teks(pasien["berat_badan"]))

    page.get_by_text(NILAI_TETAP["skrining_visual"], exact=False).first.click()
    page.get_by_text(NILAI_TETAP["instalasi"], exact=False).first.click()
    page.get_by_text(NILAI_TETAP["poli_ruang"], exact=False).first.click()

    page.get_by_role("button", name="Simpan Pendaftaran").click()
    page.wait_for_timeout(1500)


def isi_pelayanan(page, pasien):
    """Langkah 4a-4c."""
    page.get_by_text("Pelayanan Medis", exact=False).first.click()
    page.wait_for_timeout(800)

    nama = bersihkan_teks(pasien["nama"])
    page.locator(f"tr:has-text('{nama}')").first.dblclick()
    page.wait_for_timeout(800)

    tombol_ok = page.get_by_role("button", name="OK")
    if tombol_ok.count() > 0:
        tombol_ok.first.click()
    page.wait_for_timeout(500)


def isi_anamnesa(page, pasien):
    """Langkah 5a-5o."""
    page.get_by_text("Anamnesa", exact=True).first.click()
    page.wait_for_timeout(800)

    page.get_by_label("Dokter").fill(NILAI_TETAP["dokter"])
    page.get_by_label("Perawat").or_(page.get_by_label("Bidan")).first.fill(
        bersihkan_teks(pasien["tenaga_medis"])
    )
    page.get_by_label("Keluhan Utama").fill(bersihkan_teks(pasien["diagnosa"]))

    # Lama sakit (angka + satuan)
    page.get_by_label("Lama Sakit").fill(NILAI_TETAP["lama_sakit_angka"])

    page.get_by_text(NILAI_TETAP["alergi"], exact=False).first.click()

    page.get_by_label("Sistole").fill(NILAI_TETAP["sistole"])
    page.get_by_label("Diastole").fill(NILAI_TETAP["diastole"])
    page.get_by_label("Lingkar Perut").fill(NILAI_TETAP["lingkar_perut"])
    page.get_by_label("Nadi").fill(NILAI_TETAP["nadi"])
    page.get_by_label("Nafas").fill(NILAI_TETAP["nafas"])
    page.get_by_label("Edukasi").fill(NILAI_TETAP["edukasi"])

    page.get_by_role("button", name="Update").click()
    page.wait_for_timeout(2500)  # "tunggu beberapa saat" sesuai instruksi


def isi_diagnosa(page, pasien):
    """Langkah 6a-6e."""
    page.get_by_text("Diagnosa", exact=True).first.click()
    page.wait_for_timeout(800)

    icd = bersihkan_teks(pasien["diagnosa"])
    kotak_icd = page.get_by_label("ICD").or_(page.get_by_placeholder("ICD"))
    kotak_icd.first.fill(icd)
    page.wait_for_timeout(1000)
    page.locator(f"li:has-text('{icd[:20]}')").first.click()  # pilih dari dropdown autocomplete

    page.get_by_label("Prognosa").select_option(label=NILAI_TETAP["prognosa"])

    page.get_by_role("button", name="Simpan").click()
    page.wait_for_timeout(1000)

    page.mouse.wheel(0, -2000)  # geser ke atas


def pasien_pulang(page, pasien):
    """Langkah 7a-7c."""
    page.get_by_label("Status Pulang").select_option(label=NILAI_TETAP["status_pulang"])

    kotak_centang = page.get_by_label("Rencana Kontrol")
    if kotak_centang.is_checked():
        kotak_centang.uncheck()

    page.get_by_role("button", name="Selesai").click()
    page.wait_for_timeout(1000)


def proses_satu_pasien(page, pasien):
    ditemukan = cari_dan_pilih_pasien(page, pasien)
    if not ditemukan:
        return "GAGAL", "Pasien tidak ditemukan (NIK maupun No. BPJS tidak cocok)"

    isi_pendaftaran(page, pasien)
    isi_pelayanan(page, pasien)
    isi_anamnesa(page, pasien)
    isi_diagnosa(page, pasien)
    pasien_pulang(page, pasien)

    return "SUKSES", ""


# =============================================================================
# PROGRAM UTAMA
# =============================================================================

def main():
    print("=" * 70)
    print(" OTOMASI INPUT PASIEN BPJS -> EPUSKESMAS")
    print("=" * 70)

    wb, ws, daftar_pasien, header_ke_kolom = baca_data_pasien(EXCEL_INPUT)
    print(f"✔ Ditemukan {len(daftar_pasien)} pasien di file Excel.")

    if JUMLAH_PASIEN_TES:
        daftar_pasien = daftar_pasien[:JUMLAH_PASIEN_TES]
        print(f"⚠ MODE UJI COBA: hanya memproses {len(daftar_pasien)} pasien pertama.")

    hasil_per_baris = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=SLOW_MO_MS if DEBUG_MODE else 0)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(WAIT_TIMEOUT)

        page.goto(URL_LOGIN)
        print("\n>>> Silakan LOGIN MANUAL di browser (termasuk captcha).")
        input(">>> Setelah berhasil login dan berada di halaman utama, tekan ENTER di sini untuk lanjut...\n")

        jumlah_sukses = 0
        jumlah_gagal = 0

        for i, pasien in enumerate(daftar_pasien, start=1):
            nama = bersihkan_teks(pasien["nama"])
            print(f"\n[{i}/{len(daftar_pasien)}] Memproses: {nama} (NIK: {pasien['nik']})")

            try:
                status, keterangan = proses_satu_pasien(page, pasien)
            except Exception as e:
                status = "GAGAL"
                keterangan = f"Error tak terduga: {e}"
                print(f"   ❌ {keterangan}")
                traceback.print_exc()
                if DEBUG_MODE:
                    print("   ⏸  DEBUG_MODE aktif: Inspector dibuka. Perbaiki manual lalu klik Resume.")
                    page.pause()

            hasil_per_baris[pasien["_baris_excel"]] = (status, keterangan)

            if status == "SUKSES":
                jumlah_sukses += 1
                print("   ✅ Berhasil.")
            else:
                jumlah_gagal += 1
                print(f"   ❌ Gagal: {keterangan}")

            # Kembali ke halaman pencarian/pendaftaran untuk pasien berikutnya
            try:
                page.goto(URL_LOGIN)
                page.wait_for_timeout(1000)
            except Exception:
                pass

        browser.close()

    tanggal = datetime.now().strftime("%Y%m%d_%H%M%S")
    path_output = EXCEL_OUTPUT.format(tanggal=tanggal)
    simpan_hasil(wb, ws, header_ke_kolom, hasil_per_baris, path_output)

    print("\n" + "=" * 70)
    print(" LAPORAN AKHIR")
    print("=" * 70)
    print(f" Total diproses : {len(daftar_pasien)}")
    print(f" Sukses         : {jumlah_sukses}")
    print(f" Gagal          : {jumlah_gagal}")
    print("=" * 70)


if __name__ == "__main__":
    main()
