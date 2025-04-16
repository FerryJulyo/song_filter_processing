import pandas as pd

# Baca file hasil sebelumnya
df = pd.read_excel(r"D:/Top Hit/Proses/data_lagu_dengan_composer.xlsx")

# Ganti NaN composer jadi string kosong
df['Penyanyi'] = df['Penyanyi'].fillna('').replace('', 'Tanpa Penyanyi')

# Group by Composer dan jumlahkan Jumlah Pengguna
df_grouped = df.groupby('Penyanyi', as_index=False)['Jumlah Pengguna'].sum()

# Urutkan berdasarkan Jumlah Pengguna (descending)
df_grouped = df_grouped.sort_values(by='Jumlah Pengguna', ascending=False)

# Simpan ke file baru
df_grouped.to_excel(r"D:/Top Hit/Proses/grouped_by_singer.xlsx", index=False)
