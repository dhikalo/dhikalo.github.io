import os
import sys

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("Library 'pypdf' belum terinstal. Silakan jalankan perintah ini di terminal:")
    print("pip install pypdf")
    sys.exit(1)

def split_pdf(input_pdf_path, output_folder):
    # Pastikan folder output sudah ada, jika belum maka buat foldernya
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Membaca file PDF asli
    try:
        reader = PdfReader(input_pdf_path)
    except FileNotFoundError:
        print(f"Error: File '{input_pdf_path}' tidak ditemukan.")
        return

    # Mengambil nama dasar file (contoh: 'dokumen.pdf' menjadi 'dokumen')
    base_name = os.path.splitext(os.path.basename(input_pdf_path))[0]

    # Looping untuk setiap halaman di dalam PDF
    total_pages = len(reader.pages)
    for i in range(total_pages):
        writer = PdfWriter()
        
        # Mengambil satu halaman dan menambahkannya ke writer
        page = reader.pages[i]
        writer.add_page(page)
        
        # Menentukan nama file output untuk halaman tersebut
        output_filename = f"{base_name}_halaman_{i + 1}.pdf"
        output_filepath = os.path.join(output_folder, output_filename)
        
        # Menyimpan halaman tunggal tersebut menjadi file PDF baru
        with open(output_filepath, "wb") as output_pdf:
            writer.write(output_pdf)
            
        print(f"Berhasil membuat: {output_filepath}")

    print(f"\nSelesai! {total_pages} halaman telah dipisahkan ke dalam folder '{output_folder}'.")

if __name__ == "__main__":
    # Ganti 'dokumen_asli.pdf' dengan nama file PDF yang ingin Anda pecah
    # Jika file PDF berada di folder yang sama dengan script ini, cukup tulis nama filenya saja.
    file_pdf_asli = "AB26_01.pdf" 
    
    # Folder tempat file PDF per halaman akan disimpan
    folder_tujuan = "hasil_pecahan"     

    print("Memulai proses pemisahan PDF...")
    split_pdf(file_pdf_asli, folder_tujuan)
