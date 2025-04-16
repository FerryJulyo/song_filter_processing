import pandas as pd

# Load XLSX
df_xlsx = pd.read_excel(r"D:/Top Hit/Proses/hasil_akhir.xlsx", sheet_name='Indonesia')

# Ganti penyanyi "Agnez Mo" menjadi "Agnes Monica"
df_xlsx['Penyanyi'] = df_xlsx['Penyanyi'].replace({'Agnez Mo': 'Agnes Monica'})

# Hapus data setelah tanda "-"
df_xlsx['Penyanyi'] = df_xlsx['Penyanyi'].astype(str).str.replace(r'-.*$', '', regex=True).str.strip()

# Normalisasi
df_xlsx['Judul Normal'] = df_xlsx['Judul Lagu'].str.lower().str.strip()
df_xlsx['Penyanyi Normal'] = df_xlsx['Penyanyi'].str.lower().str.strip()

# Load CSV
df_csv = pd.read_csv(r"D:/Top Hit/Proses/Master_VOD.csv", sep=None, engine='python')
df_csv['Judul Normal'] = df_csv['Song'].astype(str).str.lower().str.strip()
df_csv['Sing1 Normal'] = df_csv['Sing1'].astype(str).str.lower().str.strip()
df_csv['COMPOSER1'] = df_csv['csong'].astype(str).str.replace(r'&.*$', '', regex=True).str.strip()

# Merge berdasarkan judul
df_merged = pd.merge(df_xlsx, df_csv, on='Judul Normal', how='left')

# Ambil composer jika penyanyi cocok LIKE Sing1
def ambil_composer(row):
    if pd.notnull(row['Sing1 Normal']) and pd.notnull(row['Penyanyi Normal']):
        if row['Penyanyi Normal'] in row['Sing1 Normal']:
            return row['COMPOSER1']
    return ''

df_merged['Composer Valid'] = df_merged.apply(ambil_composer, axis=1)

# Ambil composer pertama yang valid
composer_map = df_merged[df_merged['Composer Valid'] != ''].groupby(
    ['Judul Lagu', 'Penyanyi'], as_index=False
).first()[['Judul Lagu', 'Penyanyi', 'Composer Valid']]

# Gabungkan ke data awal
df_final = pd.merge(
    df_xlsx,
    composer_map,
    on=['Judul Lagu', 'Penyanyi'],
    how='left'
)

df_final = df_final.rename(columns={'Composer Valid': 'COMPOSER1'})

# Simpan hasil akhir
df_final[['Judul Lagu', 'Penyanyi', 'Jumlah Pengguna', 'Kategori', 'COMPOSER1']].to_excel(
    r"D:/Top Hit/Proses/data_lagu_dengan_composer.xlsx", index=False
)
