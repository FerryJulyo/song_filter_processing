import pandas as pd

# Baca file Excel (gantilah 'data.xlsx' dengan nama file yang sesuai)
file_input = "data.xlsx"
file_output = "hasil.xlsx"

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
