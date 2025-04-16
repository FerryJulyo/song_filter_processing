import pandas as pd
import glob

# Ambil semua file Excel dalam folder
folder_path = glob.glob(r"D:/Top Hit/Proses/*.xlsx")

# Inisialisasi dictionary untuk menyimpan data
data_dict = {}

# Debugging logs
total_pengguna_debug = 0
total_baris_diproses = 0

# Loop melalui semua file Excel
def process_excel_files():
    global total_pengguna_debug, total_baris_diproses

    for file in folder_path:
        xls = pd.ExcelFile(file)
        total_pengguna_file = 0  # Untuk melacak total pengguna di setiap file
        print(f"\n📂 Memproses file: {file}")

        for sheet_name in xls.sheet_names:
            if sheet_name.lower() in ["total", "sheet1"]:
                continue  # Lewati sheet yang tidak perlu

            df = pd.read_excel(xls, sheet_name=sheet_name)

            # Bersihkan kolom
            df.columns = df.columns.str.strip()
            df = df.dropna(subset=["Jumlah Pengguna"])

            # **Debug: Cek jumlah baris yang valid**
            total_baris_diproses += len(df)

            # Hindari perhitungan jika "Judul Lagu" atau "Penyanyi" berisi "Total"
            df = df[(df["Judul Lagu"].str.lower() != "total") & (df["Penyanyi"].str.lower() != "total")]

            # Konversi "Jumlah Pengguna" ke integer
            df["Jumlah Pengguna"] = pd.to_numeric(df["Jumlah Pengguna"], errors='coerce').fillna(0).astype(int)

            total_pengguna_sheet = df["Jumlah Pengguna"].sum()  # Hitung total pengguna di sheet
            total_pengguna_file += total_pengguna_sheet  # Tambah ke total file
            print(f"📄 Sheet: {sheet_name} → Total Pengguna: {total_pengguna_sheet}")

            for _, row in df.iterrows():
                judul_lagu = str(row["Judul Lagu"]).strip()
                penyanyi = str(row["Penyanyi"]).strip()
                jumlah_pengguna = row["Jumlah Pengguna"]
                kategori = sheet_name  # Ambil kategori dari sheet_name

                key = (judul_lagu, penyanyi, kategori)
                data_dict[key] = data_dict.get(key, 0) + jumlah_pengguna

                # **Debug: Tambahkan jumlah pengguna ke total debug**
                total_pengguna_debug += jumlah_pengguna

        print(f"📊 Total pengguna untuk file ini: {total_pengguna_file}")

# Jalankan fungsi untuk memproses file
process_excel_files()

# Konversi hasil ke DataFrame dan urutkan berdasarkan jumlah pengguna
result_df = pd.DataFrame([(k[0], k[1], v, k[2]) for k, v in data_dict.items()], 
                         columns=["Judul Lagu", "Penyanyi", "Jumlah Pengguna", "Kategori"])

result_df = result_df.sort_values(by="Jumlah Pengguna", ascending=False)

# Simpan hasil ke file Excel
output_path = r"D:/Top Hit/Proses/hasil_akumulasi.xlsx"
result_df.to_excel(output_path, index=False)

# Debugging output
print("\n✅ Proses selesai! Hasil disimpan di", output_path)
print("🔢 Total pengguna akhir (dari DataFrame):", result_df["Jumlah Pengguna"].sum())
print("🔍 Total pengguna akhir (dari debug per baris):", total_pengguna_debug)
print("📌 Total baris diproses:", total_baris_diproses)


file_input = r"D:/Top Hit/Proses/hasil_akumulasi.xlsx"
file_output = r"D:/Top Hit/Proses/hasil_akhir.xlsx"

df = pd.read_excel(file_input)

# Pastikan kolom 'Jumlah Pengguna' adalah numerik
df["Jumlah Pengguna"] = pd.to_numeric(df["Jumlah Pengguna"], errors="coerce")

# Gabungkan kategori Indonesia Pop dan Indonesia Daerah menjadi Indonesia
df["Kategori"] = df["Kategori"].replace({"Indonesia Pop": "Indonesia", "Indonesia Daerah": "Indonesia"})

# Buat writer untuk menulis ke file Excel baru
with pd.ExcelWriter(file_output, engine="xlsxwriter") as writer:
    kategori_sheets = {}  # Menyimpan jumlah sheet dengan nama yang sama

    # Loop berdasarkan kategori
    for kategori, data in df.groupby("Kategori"):
        sheet_name = kategori[:30]  # Maksimal 30 karakter untuk nama sheet

        # Jika nama sheet sudah ada, tambahkan angka untuk menghindari duplikasi
        if sheet_name in kategori_sheets:
            kategori_sheets[sheet_name] += 1
            sheet_name = f"{sheet_name} ({kategori_sheets[sheet_name]})"
        else:
            kategori_sheets[sheet_name] = 1

        # Urutkan data berdasarkan 'Jumlah Pengguna' (descending), lalu 'Judul Lagu', lalu 'Penyanyi'
        data = data.sort_values(by=["Jumlah Pengguna", "Judul Lagu", "Penyanyi"], ascending=[False, True, True])

        # Simpan data ke sheet yang sesuai
        data.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Data telah disimpan ke {file_output}")