import pandas as pd

# Baca file Excel dari Data Cleansing.xlsx sheet summary
file_path = "Data Cleansing.xlsx"
df = pd.read_excel(file_path, sheet_name="Summary")

# Tampilkan struktur data untuk memastikan kolom yang tersedia
print("Kolom yang tersedia:")
print(df.columns.tolist())
print("\nBeberapa baris pertama:")
print(df.head())

# Asumsi kolom yang dibutuhkan (sesuaikan dengan struktur data yang sebenarnya)
# Biasanya dalam data cleansing ada kolom seperti: Tahun, Jumlah_Penduduk, Angkatan_Kerja
# Sesuaikan nama kolom dengan yang ada di file Data Cleansing.xlsx

try:
    # Coba beberapa kemungkinan nama kolom
    possible_year_cols = ['Tahun', 'Year', 'tahun', 'TAHUN']
    possible_population_cols = ['Jumlah Penduduk (X)', 'Jumlah Penduduk', 'Jumlah_Penduduk', 'Population', 'Penduduk', 'JUMLAH_PENDUDUK']
    possible_workforce_cols = ['Jumlah Angkatan Kerja (Y)', 'Angkatan Kerja', 'Angkatan_Kerja', 'Workforce', 'ANGKATAN_KERJA']
    
    year_col = None
    population_col = None
    workforce_col = None
    
    # Cari kolom tahun
    for col in possible_year_cols:
        if col in df.columns:
            year_col = col
            break
    
    # Cari kolom jumlah penduduk
    for col in possible_population_cols:
        if col in df.columns:
            population_col = col
            break
    
    # Cari kolom angkatan kerja
    for col in possible_workforce_cols:
        if col in df.columns:
            workforce_col = col
            break
    
    if not all([year_col, population_col, workforce_col]):
        print("Error: Tidak dapat menemukan kolom yang diperlukan.")
        print(f"Kolom tahun ditemukan: {year_col}")
        print(f"Kolom penduduk ditemukan: {population_col}")
        print(f"Kolom angkatan kerja ditemukan: {workforce_col}")
        print("\nSilakan sesuaikan nama kolom dalam script dengan nama kolom yang sebenarnya.")
        exit()
    
    # Ambil data yang diperlukan dan bersihkan
    df_clean = df[[year_col, population_col, workforce_col]].dropna()
    
    # Rename kolom untuk konsistensi
    df_clean = df_clean.rename(columns={
        year_col: 'Tahun',
        population_col: 'Jumlah Penduduk', 
        workforce_col: 'Angkatan Kerja'
    })
    
    # Pastikan data numerik
    df_clean['Jumlah Penduduk'] = pd.to_numeric(df_clean['Jumlah Penduduk'], errors='coerce')
    df_clean['Angkatan Kerja'] = pd.to_numeric(df_clean['Angkatan Kerja'], errors='coerce')
    df_clean = df_clean.dropna()
    
    print(f"\nData yang akan digunakan untuk regresi:")
    print(df_clean)
    
    # Hitung x^2, y^2, xy untuk persamaan regresi
    df_clean["x"] = df_clean["Jumlah Penduduk"]
    df_clean["y"] = df_clean["Angkatan Kerja"]
    df_clean["x²"] = df_clean["x"] ** 2
    df_clean["y²"] = df_clean["y"] ** 2
    df_clean["xy"] = df_clean["x"] * df_clean["y"]
    
    # Tambahkan baris jumlah
    sum_row = pd.DataFrame({
        "Tahun": ["Jumlah"],
        "x": [df_clean["x"].sum()],
        "y": [df_clean["y"].sum()],
        "x²": [df_clean["x²"].sum()],
        "y²": [df_clean["y²"].sum()],
        "xy": [df_clean["xy"].sum()]
    })
    
    df_final = pd.concat([df_clean[["Tahun", "x", "y", "x²", "y²", "xy"]], sum_row], ignore_index=True)
    
    # Format ribuan (gunakan titik sebagai pemisah ribuan)
    def format_ribuan(val):
        if isinstance(val, (int, float)) and val == val:  # val == val untuk check NaN
            return f"{int(val):,}".replace(",", ".")
        return val
    
    df_final = df_final.applymap(format_ribuan)
    
    # Hitung koefisien regresi
    n = len(df_clean)
    sum_x = df_clean["x"].sum()
    sum_y = df_clean["y"].sum()
    sum_x2 = df_clean["x²"].sum()
    sum_xy = df_clean["xy"].sum()
    
    # Rumus: a = (Σy)(Σx²) - (Σx)(Σxy) / n(Σx²) - (Σx)²
    # Rumus: b = n(Σxy) - (Σx)(Σy) / n(Σx²) - (Σx)²
    denominator = n * sum_x2 - sum_x**2
    a = (sum_y * sum_x2 - sum_x * sum_xy) / denominator
    b = (n * sum_xy - sum_x * sum_y) / denominator
    
    print(f"\nPersamaan Regresi: y = {a:.2f} + {b:.6f}x")
    print(f"Dimana:")
    print(f"y = Angkatan Kerja")
    print(f"x = Jumlah Penduduk")
    
    # Simpan ke Excel
    output_path = "Tabel_Persamaan_Regresi_dari_Data_Cleansing.xlsx"
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name="Persamaan Regresi")
        
        # Tambahkan sheet untuk hasil persamaan
        result_df = pd.DataFrame({
            'Keterangan': ['Konstanta (a)', 'Koefisien (b)', 'Persamaan Regresi'],
            'Nilai': [f"{a:.2f}", f"{b:.6f}", f"y = {a:.2f} + {b:.6f}x"]
        })
        result_df.to_excel(writer, index=False, sheet_name="Hasil Regresi")
    
    print(f"\nFile berhasil disimpan ke: {output_path}")
    print("Sheet 'Persamaan Regresi' berisi tabel perhitungan")
    print("Sheet 'Hasil Regresi' berisi persamaan final")

except Exception as e:
    print(f"Error: {e}")
    print("\nPastikan:")
    print("1. File 'Data Cleansing.xlsx' ada di direktori yang sama")
    print("2. Sheet 'summary' tersedia dalam file")
    print("3. Kolom yang diperlukan tersedia (Tahun, Jumlah Penduduk, Angkatan Kerja)")