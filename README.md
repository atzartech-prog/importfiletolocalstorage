# StorageLearn: Panduan Memproses & Menyimpan Upload File di LocalStorage

Proyek ini adalah aplikasi web interaktif berbasis **HTML, CSS (Vanilla), dan JavaScript (ES6)** untuk mempelajari mekanisme unggah file (upload), pemrosesan data di sisi klien (client-side parsing), dan penyimpanannya ke dalam **HTML5 LocalStorage** menggunakan pustaka **SheetJS (XLSX)**.

---

## 💡 Konsep Utama LocalStorage & Pembatasan File

**LocalStorage** adalah fitur penyimpanan web HTML5 yang memungkinkan situs web menyimpan data berformat key-value langsung di dalam browser pengguna. Karakteristik penting LocalStorage meliputi:
1. **Hanya Mendukung Tipe Data String (Teks):** Anda tidak dapat menyimpan objek biner mentah (seperti file gambar asli, file PDF, atau file Excel `.xlsx`) secara langsung. Semua data wajib dikonversi atau diparsing menjadi string terlebih dahulu.
2. **Kapasitas Terbatas (±5 MB per Domain):** Penyimpanan dibatasi sekitar 5 Megabyte per domain/origin. Menyimpan file dengan ukuran besar sangat tidak disarankan karena akan menghabiskan kuota memori ini.
3. **Synchronous & Single-Threaded:** Operasi penulisan/pembacaan LocalStorage berjalan secara synchronous pada thread utama browser. Jika data yang diolah terlalu besar (misal gambar Base64 resolusi tinggi), browser dapat mengalami *lag* atau *freeze* (thread blocking).

---

## 📂 Tipe-tipe File yang Bisa Di-upload & Cara Pemrosesannya

Berikut adalah penjelasan jenis file yang bisa diproses untuk keperluan pembaruan data aplikasi berbasis HTML5/JS:

| Tipe File | Format Ekstensi | Metode Konversi / Parsing | Cara Menyimpan ke LocalStorage | Contoh Penggunaan |
| :--- | :--- | :--- | :--- | :--- |
| **JSON** | `.json` | Dibaca sebagai teks biasa, divalidasi ke Object JS via `JSON.parse()`. | Dimampatkan kembali dengan `JSON.stringify(data)` | Pembaruan data konfigurasi aplikasi, daftar user, katalog produk. |
| **Excel** | `.xlsx`, `.xls` | Dibaca sebagai `ArrayBuffer` biner, kemudian diparsing oleh pustaka **SheetJS** menjadi array objek. | Dimampatkan menjadi JSON string dengan `JSON.stringify(data)` | Mengimpor laporan stok, data absensi karyawan, database barang masal. |
| **CSV** | `.csv` | Dibaca sebagai teks biasa, dipisah per baris (`\n`) dan kolom (`,` atau `;`) ke format Array Objek. | Dimampatkan menjadi JSON string dengan `JSON.stringify(data)` | Sinkronisasi data ekspor dari database eksternal. |
| **Gambar** | `.png`, `.jpg`, `.webp`, `.gif` | Di-encode secara biner menjadi string panjang berformat **Base64 DataURL** menggunakan `FileReader.readAsDataURL()`. | Disimpan langsung sebagai string teks Base64 mentah | Menyimpan logo kustom buatan user, foto profil lokal, tanda tangan digital. |
| **Teks Plain** | `.txt` | Dibaca sebagai teks biasa menggunakan `FileReader.readAsText()`. | Disimpan langsung sebagai teks string mentah | Catatan catatan ringkas, logs aktivitas lokal, data teks konfigurasi. |

---

## 🧪 Mekanisme Parsing Excel Menggunakan SheetJS

Karena file Excel (`.xlsx` atau `.xls`) adalah file biner terkompresi (berbasis XML zip), browser tidak bisa langsung membaca baris dan kolom di dalamnya secara native. Kami menggunakan pustaka **SheetJS** (di-import lewat CDN) untuk menyelesaikan masalah ini.

### Alur Kerja Kode (Workflow):
1. **Menggunakan `FileReader`:**
   Membaca file yang diupload sebagai biner array:
   ```javascript
   const reader = new FileReader();
   reader.readAsArrayBuffer(file);
   ```
2. **Membaca Buku Kerja (Workbook):**
   Setelah file dimuat, pustaka `XLSX` memparse array biner tersebut:
   ```javascript
   reader.onload = function(e) {
       const data = new Uint8Array(e.target.result);
       const workbook = XLSX.read(data, { type: 'array' });
       // Mengambil daftar nama sheet: workbook.SheetNames
   };
   ```
3. **Mengubah Sheet ke Array JSON:**
   Mengambil lembar kerja tertentu dan mengubahnya menjadi objek JavaScript terstruktur (baris-kolom menjadi property-value):
   ```javascript
   const sheetName = workbook.SheetNames[0]; // ambil sheet pertama
   const worksheet = workbook.Sheets[sheetName];
   const jsonData = XLSX.utils.sheet_to_json(worksheet);
   ```
4. **Penyimpanan:**
   Mengubah objek array tersebut menjadi string JSON dan menyimpannya:
   ```javascript
   localStorage.setItem('key_name', JSON.stringify(jsonData));
   ```

---

## 🛠️ Struktur Proyek

Aplikasi dirancang sebagai aplikasi satu halaman (*Single-Page Application*) dengan struktur file sebagai berikut:
- **`index.html`**: Antarmuka dashboard utama (Glassmorphism theme) yang menampilkan file dropzone, monitor memori LocalStorage, tabs panduan belajar, visualizer data, dan modal explorer.
- **`style.css`**: Lembar gaya vanilla CSS modern yang mengimplementasikan variabel HSL, tema gelap, efek transparansi kaca (*backdrop-filter blur*), transisi halus, indikator muatan data, dan *responsive layout*.
- **`app.js`**: Otak logika aplikasi berisi logika *drag & drop*, integrasi API *FileReader*, integrasi pustaka *SheetJS*, perhitungan memori LocalStorage secara *real-time*, parsing CSV, dan rendering visualisasi data (tabel, gambar, teks) dari LocalStorage.
- **`README.md`**: Dokumentasi panduan ini.

---

## 🚀 Cara Menjalankan Aplikasi Secara Lokal

Anda dapat langsung menjalankan aplikasi ini tanpa compiler atau instalasi server yang rumit karena aplikasi ini murni berjalan di sisi client.

### Metode 1: Menggunakan Browser Langsung
1. Buka folder proyek ini di komputer Anda.
2. Klik ganda file `index.html` untuk membukanya di browser favorit Anda (Chrome, Edge, Firefox, Safari).

### Metode 2: Menggunakan Local Development Server (Rekomendasi)
Agar fitur modul atau resource eksternal berjalan optimal tanpa batasan protokol `file://`, Anda bisa menggunakan server lokal sederhana:

**Menggunakan Node.js (`http-server` / `live-server`):**
```bash
# Menggunakan npx live-server
npx live-server

# Atau menggunakan python jika sudah terinstal
python -m http.server 8000
```
Lalu buka alamat browser `http://localhost:8000` atau URL yang diberikan di terminal Anda.

---

## ⚠️ Rekomendasi Penggunaan di Produksi
Jika Anda membangun aplikasi web sesungguhnya:
- **Batasi penyimpanan Base64:** Hindari menyimpan gambar > 500KB di LocalStorage. Gunakan **IndexedDB** jika aplikasi Anda memerlukan penyimpanan file biner berukuran besar (misal 5MB - 500MB).
- **Enkripsi Data Sensitif:** Jangan menyimpan password atau data rahasia user di LocalStorage tanpa enkripsi karena LocalStorage dapat dibaca oleh skrip JavaScript apa pun di origin yang sama (rentan serangan XSS).
- **Try-Catch Block:** Selalu bungkus kode penulisan ke LocalStorage dalam blok `try-catch` untuk mengantisipasi error `QuotaExceededError` jika penyimpanan penuh.
