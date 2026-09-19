# Panduan Otomasi Input Pasien BPJS ke ePuskesmas

Script ini menjalankan otomatis 10 langkah yang sudah Anda tentukan:
baca Excel → login manual → cari pasien → pendaftaran → pelayanan →
anamnesa → diagnosa → pasien pulang → catat sukses/gagal → laporan akhir.

## 1. Instalasi (sekali saja)

Install Python 3.9+ dulu jika belum ada (https://python.org), lalu di folder ini jalankan:

```
pip install -r requirements.txt
playwright install chromium
```

## 2. Siapkan file Excel

- Taruh file Excel bulanan Anda di folder yang sama dengan script ini.
- Buka `epuskesmas_input_otomatis.py`, cari baris:
  ```
  EXCEL_INPUT = "data_pasien_bulan_ini.xlsx"
  ```
  Ganti dengan nama file Excel Anda yang sebenarnya.
- Pastikan Excel punya kolom-kolom ini (nama harus persis sama, atau
  sesuaikan di bagian `KOLOM = {...}` di script):
  `Nama Pasien`, `NIK`, `No. Asuransi`, `Jenis Kelamin`,
  `Berat Badan (Kg)`, `Tinggi Badan (cm)`, `Diagnosa`, `Tenaga Medis`.

## 3. WAJIB: Uji coba dulu dengan 2 data

Jangan langsung jalankan untuk 470 data. Di script, pastikan:

```
DEBUG_MODE = True
JUMLAH_PASIEN_TES = 2
```

Lalu jalankan:

```
python epuskesmas_input_otomatis.py
```

- Browser Chrome akan terbuka otomatis.
- **Login manual** (isi username/password + captcha) seperti biasa.
- Setelah masuk ke halaman utama, kembali ke jendela terminal/command
  prompt lalu **tekan ENTER**.
- Script akan mulai memproses 2 pasien pertama sambil Anda awasi.

### Jika ada langkah yang salah / macet

Karena situsnya tidak bisa saya cek langsung, kemungkinan ada nama
tombol/field yang sedikit berbeda dari tebakan saya. Kalau script macet:

- Jika `DEBUG_MODE = True`, jendela **Playwright Inspector** akan terbuka
  otomatis saat error. Anda bisa lihat pesan errornya.
- **Screenshot pesan error tersebut dan kirim ke saya**, sebutkan juga
  sedang di langkah nomor berapa (misalnya "macet di langkah anamnesa,
  field Dokter"). Saya akan perbaiki baris kodenya.
- Setelah diperbaiki, coba lagi dari langkah 3 (2 data dulu).

## 4. Setelah uji coba 2 data berhasil semua

Ubah:

```
DEBUG_MODE = False
JUMLAH_PASIEN_TES = None
```

Lalu jalankan lagi untuk memproses seluruh data (470 pasien).
Proses ini akan memakan waktu — perkirakan beberapa detik per pasien,
jadi untuk 470 data bisa 1-3 jam tergantung kecepatan situs. Biarkan
komputer & internet tetap menyala selama proses berjalan.

## 5. Hasil

Setelah selesai (atau jika dihentikan di tengah jalan), file laporan
`hasil_YYYYMMDD_HHMMSS.xlsx` akan dibuat di folder yang sama — berisi
data asli + kolom tambahan "Status Input Otomatis" dan "Keterangan"
untuk setiap baris pasien.

## Nilai tetap yang dipakai untuk SEMUA pasien

| Field | Nilai |
|---|---|
| Dokter | REVANI YUNI NAILUVAR |
| Perawat/Bidan | (diambil dari kolom "Tenaga Medis" Excel) |
| Keluhan Utama | (diambil dari kolom "Diagnosa" Excel) |
| Lama Sakit | 1 Hari |
| Alergi | Tidak Ada |
| Sistole | 120 |
| Diastole | 80 |
| Nadi | 80 |
| Nafas | 20 |
| Lingkar Perut | 45 |
| Edukasi | anjurkan pasien istirahat yang cukup |
| Prognosa | Bonam |
| Skrining Visual | Kondisi Stabil |
| Instalasi | Rawat Jalan |
| Poli/Ruang | KLASTER 3 UMUM |
| Status Pulang | Berobat Jalan |
| Rencana Kontrol | (selalu di-uncheck) |

Kalau ada nilai yang perlu diubah, cari bagian `NILAI_TETAP = {...}`
di file `epuskesmas_input_otomatis.py` dan ubah nilainya di situ.

## Penting soal keamanan

- Script ini **tidak menyimpan username/password Anda** di mana pun —
  Anda login langsung di jendela browser yang terbuka.
- Jalankan script ini hanya di komputer Anda sendiri yang aman.
