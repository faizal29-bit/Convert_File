# LocalConvert 🚀

Sistem konversi file lokal — mirip iLovePDF tapi **berjalan di komputer sendiri**.
Tanpa batas file, tanpa internet, tanpa data keluar dari perangkatmu.

## Fitur Konversi

| Dari | Ke | Mesin |
|------|-----|-------|
| 🖼️ Gambar | PDF | img2pdf |
| 🖼️ Gambar | Word (DOCX) | python-docx |
| 🖼️ Gambar | PPT (PPTX) | python-pptx |
| 📝 Word | PDF | LibreOffice |
| 📄 PDF | Word | pdf2docx |
| 📊 PPT | PDF | LibreOffice |
| 📄 PDF | PPT | pdf2image + python-pptx |
| 📝 Word | PPT | python-pptx |
| 📗 Excel | PDF | LibreOffice |

## Instalasi

### Windows

**1. Install Python 3.10+**
https://www.python.org/downloads/

**2. Install LibreOffice** *(untuk Word/PPT/Excel → PDF)*
https://www.libreoffice.org/download/libreoffice/

**3. Install Poppler** *(untuk PDF → PPT)*
- Download: https://github.com/oschwartz10612/poppler-windows/releases
- Ekstrak ke `C:\poppler\`
- Tambahkan `C:\poppler\bin` ke PATH sistem, atau biarkan di sana — app akan mendeteksi otomatis.

**4. Jalankan installer**
```
install_windows.bat
```

**5. Jalankan aplikasi**
```
start_windows.bat
```

---

### Ubuntu / Debian

```bash
# Install dependensi sistem
sudo apt update
sudo apt install python3 python3-pip python3-venv libreoffice poppler-utils

# Install app
bash install_linux.sh

# Jalankan
bash start_linux.sh
```

### macOS

```bash
# Install dependensi sistem
brew install libreoffice poppler

# Install app
bash install_linux.sh

# Jalankan
bash start_linux.sh
```

---

## Cara Pakai

1. Buka browser → `http://localhost:5000`
2. Pilih jenis konversi dari grid
3. Drag & drop file atau klik "Pilih File"
4. Klik **Konversi Sekarang**
5. Unduh hasilnya

## Catatan Teknis

- **Tanpa limit**: Konversi berapa file sekaligus pun, tidak ada batasan
- **Temp files**: File sementara disimpan di temp OS, otomatis dihapus setelah 2 jam
- **Thread-safe**: Server mendukung beberapa konversi paralel
- **Port**: Default `5000`, ubah di `app.py` baris terakhir

## Troubleshooting

| Masalah | Solusi |
|---------|--------|
| "LibreOffice tidak ditemukan" | Install LibreOffice dan pastikan `soffice` ada di PATH |
| "PDF → PPT gagal" | Install Poppler (`pdftoppm` harus ada di PATH) |
| Port 5000 bentrok | Ubah port di `app.py`: `app.run(port=XXXX)` |
| Konversi lambat | Normal untuk file besar. LibreOffice butuh waktu startup |

## Dependensi

```
flask            — Web framework
img2pdf          — Konversi gambar ke PDF lossless
Pillow           — Manipulasi gambar
python-docx      — Baca/tulis Word
pdf2docx         — Konversi PDF ke Word
python-pptx      — Buat dan edit PPT
pdf2image        — Render halaman PDF jadi gambar
LibreOffice      — Konversi Office ↔ PDF (eksternal)
Poppler          — Render PDF (eksternal, untuk pdf2image)
```
