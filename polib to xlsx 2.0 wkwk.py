import polib
from openpyxl import Workbook
from tkinter import Tk, filedialog, messagebox, Toplevel, Label
from tkinter.ttk import Progressbar
import os

# Fungsi untuk memperbarui dialog progress bar
def update_progress(progress_window, progress_bar, progress_label, progress):
    progress_bar['value'] = progress
    progress_label.config(text=f"Proses: {progress:.2f}%")
    progress_window.update_idletasks()

# Fungsi untuk membaca file .po dan menulis ke file .xlsx
def convert_po_to_xlsx(po_file_path):
    # Membaca file .po menggunakan polib
    po = polib.pofile(po_file_path)

    # Membuat workbook Excel baru
    workbook = Workbook()
    
    # Mendapatkan nama file untuk output berdasarkan input
    file_name = os.path.splitext(os.path.basename(po_file_path))[0]
    xlsx_file_path = f"{file_name}.xlsx"

    # Sheet saat ini
    sheet = workbook.active
    sheet.title = "Sheet1"
    
    # Menambahkan header ke file Excel
    sheet.append(["Comment Line", "msgctxt", "msgid", "msgstr"])
    
    # Pengaturan batas baris per sheet
    max_rows_per_sheet = 3000
    row_count = 1  # Dimulai dari 1 karena header dihitung
    sheet_number = 1

    # Memproses setiap entri di file .po
    total_entries = len(po)
    for idx, entry in enumerate(po):
        # Mengambil komentar, msgctxt, msgid, dan msgstr
        comments = "\n".join(entry.comment.splitlines()) if entry.comment else ""
        msgctxt = entry.msgctxt or ""
        msgid = entry.msgid or ""
        msgstr = entry.msgstr or ""

        # Menambah baris baru ke sheet
        sheet.append([comments, msgctxt, msgid, msgstr])
        row_count += 1

        # Jika melebihi batas baris, buat sheet baru
        if row_count > max_rows_per_sheet:
            sheet_number += 1
            sheet = workbook.create_sheet(title=f"Sheet{sheet_number}")
            sheet.append(["Comment Line", "msgctxt", "msgid", "msgstr"])
            row_count = 1

    # Menyimpan file Excel
    workbook.save(xlsx_file_path)

# Fungsi utama untuk menjalankan konversi dengan dialog pemilihan folder/file
def main():
    # Membuat jendela Tkinter
    root = Tk()
    root.withdraw()  # Sembunyikan jendela utama

    # Dialog pemilihan beberapa file PO
    po_file_paths = filedialog.askopenfilenames(
        title="Pilih File PO",
        filetypes=[("PO files", "*.po")]
    )

    if po_file_paths:
        # Membuat dialog progress bar
        progress_window = Toplevel()
        progress_window.title("Konversi sedang berlangsung...")
        progress_label = Label(progress_window, text="Proses: 0%")
        progress_label.pack(pady=10)
        progress_bar = Progressbar(progress_window, orient="horizontal", length=300, mode='determinate')
        progress_bar.pack(pady=10)

        total_files = len(po_file_paths)

        # Proses setiap file PO yang dipilih
        for i, po_file_path in enumerate(po_file_paths):
            # Konversi file PO ke XLSX
            convert_po_to_xlsx(po_file_path)

            # Menghitung persentase dan memperbarui progress bar
            progress = ((i + 1) / total_files) * 100
            update_progress(progress_window, progress_bar, progress_label, progress)

        # Menutup dialog progress bar
        progress_window.destroy()
        
        # Tampilkan pesan DONE setelah selesai
        messagebox.showinfo("Info", f"DONE: Konversi selesai untuk {total_files} file!")

if __name__ == "__main__":
    main()
