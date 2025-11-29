# Analisis Regresi Linier: Kota Tangerang

Repositori ini berisi analisis **pengaruh jumlah penduduk terhadap jumlah angkatan kerja di Kota Tangerang** menggunakan metode **regresi linier**.  
Analisis dilakukan dengan Python dan didukung oleh beberapa file Excel, PDF laporan, serta grafik hasil regresi. :contentReference[oaicite:0]{index=0}

---

## Tujuan Proyek

1. Mengolah data jumlah penduduk dan jumlah angkatan kerja Kota Tangerang.
2. Menyusun model **regresi linier sederhana**:
   \[
   Y = a + bX
   \]
   di mana:
   - \( X \) = jumlah penduduk  
   - \( Y \) = jumlah angkatan kerja
3. Mengukur seberapa kuat hubungan antara kedua variabel tersebut (korelasi & koefisien determinasi).
4. Menyajikan hasil dalam bentuk:
   - Tabel perhitungan,
   - Grafik garis regresi,
   - Ringkasan laporan dalam format PDF/Excel. :contentReference[oaicite:1]{index=1}

---

## Struktur Repo

Kurang lebih struktur utama repo ini adalah sebagai berikut: :contentReference[oaicite:2]{index=2}

```text
regresi-linier/
├── Analisis_Regresi_Linier.pdf        # Laporan analisis (PDF)
├── Data Cleansing.xlsx                # Hasil pembersihan data
├── Regresi_Output_Data_dan_Grafik.xlsx# Output regresi + grafik
├── Tabel_Persamaan_Regresi_dari_Data_Cleansing.xlsx
├── data_cleansing_update.py           # Script pembersihan data
├── persamaan_regresi_update.py        # Script perhitungan persamaan regresi
├── regresi_linier_update.py           # Script utama analisis & visualisasi
├── grafik_regresi.png                 # Grafik hasil regresi
└── datasheet/                         # Folder berisi data sumber (CSV/Excel)
