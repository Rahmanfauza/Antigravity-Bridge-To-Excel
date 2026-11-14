# ExcelAI Bridge

Aplikasi desktop ringan untuk menghubungkan Microsoft Excel yang sedang terbuka dengan **Antigravity** secara *real-time* melalui HTTP server lokal (`http://127.0.0.1:5050`) dan `xlwings`.

## Fitur Utama
- **Real-Time Bridge**: Membaca, menulis, menambah sheet, membuat chart/PivotTable, dan menjalankan kode `xlwings` langsung ke workbook yang sedang aktif.
- **Auto-Load Add-in**: Otomatis memuat Add-in Excel yang terpasang (seperti `Terbilang.xlam`) saat membuka Excel.
- **Multi-File & Multi-Format**: Mendukung `.xlsx`, `.xlsm`, `.xls`, `.csv`, `.tsv`, dan `.ods`, serta otomatis menutup `Book1` kosong.
- **Auto-Save**: Perubahan otomatis disimpan ke file (`wb.save(wb.fullname)`).

## Cara Menggunakan
1. Jalankan `run.bat` (atau `python app.py` di terminal).
2. Klik **Pilih & Buka File Excel**, lalu pilih file yang ingin dikerjakan.
3. Biarkan jendela **ExcelAI Bridge** dan **Excel** tetap terbuka.
4. Beri instruksi ke Antigravity—perintah akan dikirim ke `http://127.0.0.1:5050` dan hasilnya langsung muncul di Excel Anda.

## Endpoint HTTP Lokal (`http://127.0.0.1:5050`)
- `GET /status` — Daftar workbook yang terbuka (`name`, `path`, `sheets`).
- `POST /read` — `{"book": "...", "sheet": "...", "range": "A1:I20"}`
- `POST /write` — `{"book": "...", "sheet": "...", "cell": "F4", "value": 6000000, "save": true}`
- `POST /add_sheet` — `{"book": "...", "name": "..."}`
- `POST /run_code` — `{"book": "...", "code": "...", "save": true}`

## Catatan & Troubleshooting
- **File Read-Only**: Jika file berstatus *Read-Only*, klik kanan file → **Properties** → hilangkan centang **Read-only** (dan centang **Unblock** jika dari internet), atau simpan dengan **Save As (`F12`)**.
- **Undo (`Ctrl+Z`)**: Perubahan melalui `xlwings`/COM otomatis mengosongkan riwayat Undo di Excel, jadi pastikan cek kembali sebelum mengeksekusi perubahan besar.
- **Excel Tidak Muncul / Terkunci**: Jika semua jendela Excel sudah ditutup namun file masih terkunci, buka **Task Manager (`Ctrl+Shift+Esc`)** dan tutup proses **Microsoft Excel** yang tersisa di background.
