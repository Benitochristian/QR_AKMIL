import pandas as pd
import os
from pathlib import Path
from jinja2 import Template
import qrcode

# Path file Excel
excel_file = "NEW AKMIL 2025.xlsx"

# Baca sheet Master Data
xls = pd.ExcelFile(excel_file)
df_master = pd.read_excel(xls, sheet_name='Master Data')

# Bersihkan data
df_cleaned = df_master.iloc[2:].reset_index(drop=True)
df_cleaned.columns = df_master.iloc[0]
df_cleaned = df_cleaned.iloc[1:].reset_index(drop=True)

# Pilih kolom penting
selected_columns = [
    "Nama", "NRP", "Primary key", "BODY SPEC", "SIZE PDH",
    "SIZE KEMEJA", "SIZE CELANA", "Keterangan", "PREMAN"
]
available_columns = [col for col in selected_columns if col in df_cleaned.columns]
df_selected = df_cleaned[available_columns]
sample_df = df_selected.head()  # Ambil 5 sample

# Template HTML
html_template = Template("""
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <title>{{ nama }}</title>
    <style>
        body { font-family: Arial, sans-serif; background: #fff; color: #000; padding: 20px; }
        .container { border: 2px solid #000; padding: 20px; border-radius: 10px; max-width: 500px; margin: auto; }
        h2 { text-align: center; }
    </style>
</head>
<body>
    <div class="container">
        <h2>Data Personal Jas PDU</h2>
        <p><strong>Nama:</strong> {{ nama }}</p>
        <p><strong>NRP:</strong> {{ nrp }}</p>
        <p><strong>Size Kemeja:</strong> {{ size_kemeja }}</p>
        <p><strong>Size Celana:</strong> {{ size_celana }}</p>
        <p><strong>Keterangan:</strong> {{ keterangan or "-" }}</p>
    </div>
</body>
</html>
""")

# Buat folder output
output_folder = Path("output")
html_folder = output_folder / "html"
qr_folder = output_folder / "qr"
html_folder.mkdir(parents=True, exist_ok=True)
qr_folder.mkdir(parents=True, exist_ok=True)

# Generate HTML & QR
for _, row in sample_df.iterrows():
    nama = str(row.get("Nama", "")).strip()
    nrp = str(row.get("NRP", "")).strip().replace(".", "")
    filename = f"{nrp}.html"
    html_path = html_folder / filename

    # Render HTML
    html_content = html_template.render(
        nama=nama,
        nrp=row.get("NRP", ""),
        size_kemeja=row.get("SIZE KEMEJA", ""),
        size_celana=row.get("SIZE CELANA", ""),
        keterangan=row.get("Keterangan", "")
    )

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    # QR Code
    qr_url = f"http://localhost/{filename}"  # bisa diubah ke link hosting
    qr_img = qrcode.make(qr_url)
    qr_img.save(qr_folder / f"{nrp}.png")

print("✅ Selesai! File HTML dan QR code ada di folder 'output'")
