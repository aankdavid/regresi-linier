import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import os

print("MEMULAI ANALISIS REGRESI...")
print("="*50)

# === Baca Data dari File Data Cleansing.xlsx ===
try:
    # Baca data dari sheet Summary
    input_file = "Data Cleansing.xlsx"
    
    if not os.path.exists(input_file):
        print(f"ERROR: File '{input_file}' tidak ditemukan!")
        print("Pastikan Anda sudah menjalankan script data cleansing terlebih dahulu.")
        exit()
    
    # Baca data dari sheet Summary
    df_source = pd.read_excel(input_file, sheet_name="Summary")
    print(f"Data berhasil dibaca dari '{input_file}'")
    print("Data yang dibaca:")
    print(df_source)
    
    # Ekstrak data ke variabel
    tahun = df_source['Tahun'].tolist()
    X = df_source['Jumlah Penduduk (X)'].tolist()  # Data Penduduk
    Y = df_source['Jumlah Angkatan Kerja (Y)'].tolist()  # Data Angkatan Kerja
    n = len(X)
    
    print(f"\nData yang akan dianalisis:")
    print(f"Tahun: {tahun}")
    print(f"Jumlah Penduduk (X): {X}")
    print(f"Jumlah Angkatan Kerja (Y): {Y}")
    print(f"Jumlah data: {n}")
    
except Exception as e:
    print(f"ERROR membaca file: {str(e)}")
    print("Menggunakan data dummy untuk testing...")
    
    # Data dummy jika file tidak bisa dibaca
    tahun = [2020, 2021, 2022, 2023, 2024]
    X = [1742604, 1771092, 1834962, 1912679, 1927815]
    Y = [5552172, 5698344, 5940618, 5516656, 5797923]
    n = len(X)

print("\n" + "="*50)
print("MENGHITUNG REGRESI LINEAR...")

# === Hitung Regresi nilai a, b, dan Y=a+bX ===
sum_X = sum(X)
sum_Y = sum(Y)
sum_XY = sum([X[i]*Y[i] for i in range(n)])
sum_X2 = sum([X[i]**2 for i in range(n)])

# Hitung koefisien regresi
b = (n * sum_XY - sum_X * sum_Y) / (n * sum_X2 - sum_X**2)
a = (sum_Y - b * sum_X) / n

# Hitung prediksi
y_pred = [a + b * x for x in X]

# Tampilkan hasil perhitungan
print(f"Koefisien a (intercept): {a:,.2f}")
print(f"Koefisien b (slope): {b:.6f}")
print(f"Persamaan regresi: Y = {a:,.2f} + {b:.6f}X")

# Hitung koefisien korelasi (R)
mean_X = sum_X / n
mean_Y = sum_Y / n
numerator = sum([(X[i] - mean_X) * (Y[i] - mean_Y) for i in range(n)])
denominator_X = sum([(X[i] - mean_X)**2 for i in range(n)])
denominator_Y = sum([(Y[i] - mean_Y)**2 for i in range(n)])
r = numerator / (denominator_X * denominator_Y)**0.5
r_squared = r**2

print(f"Koefisien korelasi (r): {r:.4f}")
print(f"Koefisien determinasi (r²): {r_squared:.4f}")
print(f"Akurasi model: {r_squared*100:.2f}%")

print("\n" + "="*50)
print("MEMBUAT GRAFIK...")

# === Simpan Gambar Grafik ===
plt.figure(figsize=(10, 6))
plt.plot(tahun, Y, 'bo-', label="Data Aktual", markersize=8, linewidth=2)
plt.plot(tahun, y_pred, 'r--', label="Regresi Linear", linewidth=2)

# Tambahkan informasi pada grafik
plt.xlabel("Tahun", fontsize=12)
plt.ylabel("Jumlah Angkatan Kerja", fontsize=12)
plt.title(f"Regresi Linear: Penduduk vs Angkatan Kerja Kota Tangerang\nY = {a:,.0f} + {b:.6f}X (r² = {r_squared:.4f})", fontsize=14)
plt.grid(True, alpha=0.3)
plt.legend(fontsize=11)

# Format y-axis untuk menampilkan angka dalam jutaan
plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1e6:.1f}M'))

# Tambahkan nilai pada setiap titik
for i, (t, actual, pred) in enumerate(zip(tahun, Y, y_pred)):
    plt.annotate(f'{actual/1e6:.1f}M', (t, actual), textcoords="offset points", xytext=(0,10), ha='center', fontsize=9)
    plt.annotate(f'{pred/1e6:.1f}M', (t, pred), textcoords="offset points", xytext=(0,-15), ha='center', fontsize=9, color='red')

plt.tight_layout()
plt.savefig("grafik_regresi.png", dpi=300, bbox_inches='tight')
plt.close()

print("Grafik berhasil disimpan sebagai 'grafik_regresi.png'")

print("\n" + "="*50)
print("MENYIMPAN HASIL KE EXCEL...")

# === Simpan Data ke Excel ===
df_result = pd.DataFrame({
    "Tahun": tahun,
    "Jumlah Penduduk (X)": X,
    "Angkatan Kerja Aktual (Y)": Y,
    "Angkatan Kerja Prediksi": [round(y) for y in y_pred],
    "Selisih (Aktual - Prediksi)": [round(Y[i] - y_pred[i]) for i in range(n)],
    "Persentase Error (%)": [round(abs(Y[i] - y_pred[i])/Y[i]*100, 2) for i in range(n)]
})

# Tambahkan informasi regresi
df_info = pd.DataFrame({
    "Parameter": ["Intercept (a)", "Slope (b)", "Koefisien Korelasi (r)", "Koefisien Determinasi (r²)", "Persamaan Regresi"],
    "Nilai": [f"{a:,.2f}", f"{b:.6f}", f"{r:.4f}", f"{r_squared:.4f}", f"Y = {a:,.2f} + {b:.6f}X"]
})

excel_path = "Regresi_Output_Data_dan_Grafik.xlsx"

try:
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_result.to_excel(writer, index=False, sheet_name="Data dan Hasil")
        df_info.to_excel(writer, index=False, sheet_name="Info Regresi")
    
    print(f"Data berhasil disimpan ke '{excel_path}'")
    
    # Tambah Grafik ke Excel
    wb = load_workbook(excel_path)
    ws = wb["Data dan Hasil"]
    
    # Tambahkan grafik
    img = ExcelImage("grafik_regresi.png")
    img.anchor = "H2"  # Posisi grafik
    img.width = 600    # Lebar grafik
    img.height = 360   # Tinggi grafik
    ws.add_image(img)
    
    wb.save(excel_path)
    print("Grafik berhasil ditambahkan ke Excel")
    
except Exception as e:
    print(f"ERROR menyimpan Excel: {str(e)}")

print("\n" + "="*50)
print("MEMBUAT PDF...")

# === Simpan Grafik dalam PDF ===
pdf_path = "Analisis_Regresi_Linier.pdf"

try:
    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4
    
    # Judul
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width / 2, height - 50, "ANALISIS REGRESI LINEAR")
    c.drawCentredString(width / 2, height - 70, "Penduduk vs Angkatan Kerja Kota Tangerang")
    
    # Grafik
    c.drawImage("grafik_regresi.png", 50, height / 2 - 50, width=500, height=300, preserveAspectRatio=True)
    
    # Informasi regresi
    c.setFont("Helvetica", 12)
    y_pos = height / 2 - 100
    c.drawString(50, y_pos, f"Persamaan Regresi: Y = {a:,.2f} + {b:.6f}X")
    c.drawString(50, y_pos - 20, f"Koefisien Korelasi (r): {r:.4f}")
    c.drawString(50, y_pos - 40, f"Koefisien Determinasi (r²): {r_squared:.4f}")
    c.drawString(50, y_pos - 60, f"Akurasi Model: {r_squared*100:.2f}%")
    
    # Tabel data
    c.setFont("Helvetica-Bold", 10)
    c.drawString(50, y_pos - 100, "Data dan Prediksi:")
    
    c.setFont("Helvetica", 9)
    header_y = y_pos - 120
    c.drawString(50, header_y, "Tahun")
    c.drawString(100, header_y, "Penduduk")
    c.drawString(180, header_y, "Aktual")
    c.drawString(250, header_y, "Prediksi")
    c.drawString(320, header_y, "Error %")
    
    for i, (t, x, actual, pred) in enumerate(zip(tahun, X, Y, y_pred)):
        row_y = header_y - 20 - (i * 15)
        error_pct = abs(actual - pred) / actual * 100
        c.drawString(50, row_y, str(t))
        c.drawString(100, row_y, f"{x:,}")
        c.drawString(180, row_y, f"{actual:,}")
        c.drawString(250, row_y, f"{pred:,.0f}")
        c.drawString(320, row_y, f"{error_pct:.2f}%")
    
    c.save()
    print(f"PDF berhasil disimpan sebagai '{pdf_path}'")
    
except Exception as e:
    print(f"ERROR membuat PDF: {str(e)}")

print("\n" + "="*60)
print("RINGKASAN HASIL ANALISIS")
print("="*60)
print(f"File input: Data Cleansing.xlsx")
print(f"Jumlah data: {n} tahun ({min(tahun)}-{max(tahun)})")
print(f"Persamaan regresi: Y = {a:,.2f} + {b:.6f}X")
print(f"Koefisien korelasi (r): {r:.4f}")
print(f"Koefisien determinasi (r²): {r_squared:.4f}")
print(f"Akurasi model: {r_squared*100:.2f}%")
print("\nFile output yang dihasilkan:")
print(f"1. {excel_path}")
print(f"2. {pdf_path}")
print(f"3. grafik_regresi.png")
print("\nPROSES SELESAI!")
print("="*60)