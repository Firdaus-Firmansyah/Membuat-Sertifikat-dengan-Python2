import openpyxl
from docxtpl import DocxTemplate
import datetime

# Path file Excel dan Word
excel = r"C:\FIRDAUS\LOCAL DISK D\Belajar Python\Python Sertifikat\sertifikat_nyoba.xlsx"
word_template = r"C:\FIRDAUS\LOCAL DISK D\Belajar Python\Python Sertifikat\Sertifikat.docx"

# Memuat workbook dan sheet
load = openpyxl.load_workbook(excel)
sheet = load.active

# Menampilkan nama sheet
print(sheet.title)

# Mengambil semua nilai dalam sheet sebagai list
get_values = list(sheet.values)  # Pastikan ini ada sebelum digunakan
print(get_values)  # Menampilkan semua nilai

# Memuat template Word
doc = DocxTemplate(word_template)

# Iterasi untuk setiap baris data kecuali header
for value_tuple in get_values[1:]:  # Menghindari header, mulai dari baris kedua
    if len(value_tuple) >= 5:  # Memastikan ada 5 kolom data
        if all(value_tuple):  # Skip baris kosong
            # Memeriksa apakah kolom TANGGAL adalah string atau datetime
            tanggal = value_tuple[3]
            if isinstance(tanggal, str):
                try:
                    # Jika tanggal berupa string, konversi ke datetime
                    tanggal = datetime.datetime.strptime(tanggal, '%Y-%m-%d')  # Gantilah format ini sesuai format di Excel Anda
                except ValueError:
                    # Jika tidak bisa dikonversi, lanjutkan dengan tanggal default atau kosong
                    tanggal = None
            elif isinstance(tanggal, datetime.datetime):
                # Jika sudah berupa datetime, langsung gunakan
                pass
            
            # Render dokumen Word dengan data
            doc.render({
                "JUDUL": value_tuple[0],          # Kolom 1 (Judul)
                "TAHUN": value_tuple[1],          # Kolom 2 (Tahun)
                "NAMA": value_tuple[2],           # Kolom 3 (Peserta)
                "TANGGAL": tanggal.strftime('%d-%m-%Y') if tanggal else '',  # Format tanggal jika valid
                "PENYELENGGARA": value_tuple[4],  # Kolom 5 (Penyelenggara)
            })

            # Membuat nama file berdasarkan data
            doc_name = f"Sertifikat_{value_tuple[0]}_{value_tuple[2]}.docx"  # Nama berdasarkan Judul dan Peserta
            doc.save(doc_name)
            print(f"File {doc_name} berhasil dibuat.")
        else:
            print(f"Baris data {value_tuple} tidak ")