# TA ML — analisis petir pada tower transmisi

Repo ini berisi **kode dan notebook tanpa data** untuk cleaning, analisis, dan evaluasi
prediksi `Count` petir tahunan pada 617 tower. `Count` adalah observasi petir di area
sekitar tower, bukan sambaran langsung ke struktur dan bukan jumlah gangguan PLN.
Notebook utama ada di [`notebooks/01_analisis.ipynb`](notebooks/01_analisis.ipynb):
termasuk section tersendiri untuk **Cleaning dan Validasi Data**. Logika cleaning
dan model ada di `src/ta_ml/` agar notebook tetap singkat dan bisa dijalankan ulang.

## Persiapan lokal

Gunakan Python 3.12 dan venv. Di PowerShell:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev,notebook]"
```

Simpan tiga workbook asli berikut di `data/` lokal (folder ini diabaikan Git):

- `DATA PETIR 2021-2025 UIP3B KAL -statistik (1).xlsx`
- `Data Pendukung Proteksi Petir UIP3B Kal 2026.xlsx`
- `workbook_mitigasi_r22g14_numer - SHARE.xlsx`

Jika file berada di folder lain, tetapkan `TA_ML_DATA_DIR` ke folder itu, misalnya
`$env:TA_ML_DATA_DIR = 'C:\lokasi\workbook'`. Kode tidak menyalin atau mengubah
workbook mentah. Hasil lokal ditulis ke `data/output/`, atau ke lokasi pada
`TA_ML_OUTPUT_DIR` saat notebook dijalankan.

## Jalankan

```powershell
python -m unittest discover -s tests -v
ruff check src tests
ruff format --check src tests
python -m ta_ml.cli --output-dir data/output
jupyter lab notebooks/01_analisis.ipynb
```

Jika workbook berada di folder lain, tambahkan `--data-dir $env:TA_ML_DATA_DIR`.
CLI menolak
menimpa hasil lama kecuali diberi `--overwrite`; notebook sengaja membangun ulang
hasil lokal agar seluruh section konsisten dalam satu eksekusi.

Empat keluaran lokal:

| File | Isi |
|---|---|
| `ta_ml_clean_v2.xlsx` | 3.702 baris tower–tahun, registry 617 tower, audit koordinat 2020, provenance sumber |
| `ta_ml_model_panel.csv` | 3.085 pasangan fitur tahun sebelumnya → `Count` tahun target; `split` development/test |
| `evaluation_report.json` | semua kandidat, validasi 2023–2024, model terpilih, hasil 2025, analisis PLN |
| `test_2025_predictions.csv` | prediksi dan nilai nyata 2025 untuk 617 tower, termasuk skala PLN dan penanda anomali |

## Protokol analisis

- `Count` 2020 dipakai sebagai riwayat awal. Dari 617 nama 2020 yang tidak sesuai GPS,
  609 nilai dipasangkan ulang memakai koordinat unik; delapan sisanya tetap tidak
  ditugaskan. Nama asal dan keputusan tersimpan pada audit, bukan ditimpa di raw.
- Kandidat algoritma adalah `PoissonRegressor` dan
  `HistGradientBoostingRegressor(loss="poisson")`, masing-masing pada empat kelompok
  fitur bertahap. Baseline sederhana menggunakan `Count` tahun sebelumnya.
- Pilih kombinasi algoritma–fitur melalui dua validasi berurutan: latih target
  2021–2022 → validasi 2023; latih 2021–2023 → validasi 2024. Rata-rata MAE
  adalah ukuran utama, rata-rata Poisson deviance pembeda berikutnya.
  Setelah pilihan dikunci, latih ulang pada target 2021–2024 dan nilai pada 2025.
  Imputasi dan transformasi dipelajari hanya dari bagian latih tiap putaran.
- Fitur hanya memakai informasi yang tersedia sebelum tahun target. Statistik
  petir tahun target, `Density` tahun target, serta hasil PLN tidak masuk fitur.
  Grounding/proteksi dari file berlabel 2026 disimpan untuk audit dan analisis
  sensitivitas, tetapi bukan kandidat model historis karena tanggal ukur tiap
  nilainya belum terbukti. Geometri diperlakukan relatif tetap dan keterbatasan
  asumsi ini dijelaskan di notebook.
- Hasil 2025 adalah **evaluasi retrospektif yang direvisi**: tahun ini sudah
  pernah dilihat pada iterasi sebelumnya, sehingga bukan *untouched blind test*.
  Tidak ada klaim prediksi 2026 dalam repo ini.
- Hubungan sumber `Density = round(Count / 3.1273, 2)` berlaku untuk seluruh
  3.694 nilai `Count` yang terpetakan, tetapi kolom sumber `Area` bernilai 2.
  Ketidaksesuaian definisi itu didokumentasikan, bukan diperbaiki diam-diam.
  Prediksi `Count` dikonversi ke skala PLN `P_strike_annual` untuk analisis
  keselarasan. Dua `Ng` workbook yang jatuh ke default akibat format koma
  dikecualikan dari ringkasan, tetapi tetap ditandai dalam file prediksi.
  Workbook PLN memakai kepadatan petir 2025 sebagai input, jadi bukan peramal
  independen 2025.

Tes integrasi memakai workbook asli lokal, tanpa dataset atau fixture sintetis.
Karena repo tidak membawa data, tes tidak bisa dijalankan di clone yang belum
diberi akses ke workbook tersebut. Sebelum commit notebook, pastikan semua
cell output dan execution count kosong; hasil eksekusi hanya boleh tinggal lokal.
