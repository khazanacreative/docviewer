# Excel & CSV Document Viewer - Antigravity

Aplikasi katalog penampil dokumen berbasis web-app mandiri yang berfungsi dengan cara mengunggah (upload) file Excel (`.xlsx`, `.xls`, atau `.csv`). Aplikasi ini dirancang agar sangat ringan (antigravity), responsif di layar HP/Desktop, serta dioptimalkan khusus untuk membaca teks panjang hasil salinan dari WhatsApp.

## 🚀 Fitur Utama

1. **Local Parsing Instan (< 1 detik)**: Menggunakan library SheetJS untuk memproses file Excel secara penuh di sisi klien (client-side) tanpa memerlukan database atau server eksternal. Privasi data Anda terjamin 100%.
2. **Pencarian Real-Time**: Kolom pencarian dinamis yang menyaring daftar judul dokumen secara instan saat Anda mengetik kata kunci.
3. **Format WhatsApp Akurat**: Mempertahankan format asli spasi, baris baru/enter, serta memproses format teks khusus WhatsApp (`*tebal*`, `_miring_`, `~coret~`, dan link URL otomatis) menjadi visual HTML yang rapi.
4. **Layout Responsif (Sidebar-to-Content)**:
   - **Desktop**: Layout dua kolom (Sidebar navigasi 30% dan Panel baca utama 70%).
   - **Mobile (HP)**: Navigasi geser/transisi master-detail yang mulus, memaksimalkan ruang baca layar HP dengan tombol "Kembali" untuk mempermudah navigasi.
5. **Aksesibilitas Cepat**:
   - Tombol **"Salin Teks"** sekali klik dengan indikator keberhasilan visual.
   - Fitur **Cache Lokal** (Local Storage) untuk mengingat berkas terakhir yang diunggah secara otomatis saat halaman disegarkan (refresh).
   - Tombol **"File Demo"** untuk menguji aplikasi secara langsung dengan satu klik.

## 📁 Struktur Kolom Berkas Excel

Agar dokumen terbaca dengan benar, pastikan berkas Excel/CSV Anda memiliki struktur kolom berikut pada sheet pertama:
* **Kolom A (Kolom ke-1)**: `Judul` (Nama atau tajuk singkat dokumen).
* **Kolom B (Kolom ke-2)**: `Keterangan` (Deskripsi pendek, tanggal, atau info tambahan).
* **Kolom C (Kolom ke-3)**: `Konten` (Teks panjang hasil copy-paste WhatsApp, termasuk spasi, emoji, dan baris baru/enter).

*Catatan: Baris pertama (header) berupa label kolom seperti "Judul", "Keterangan", "Konten" akan dideteksi secara otomatis dan dilewati oleh sistem.*

## 🛠️ Cara Menjalankan Aplikasi

Aplikasi ini menggunakan **Vite** sebagai dev server lokal. Ikuti langkah berikut untuk menjalankannya:

1. **Instal Dependensi**:
   Buka terminal di direktori proyek ini dan jalankan:
   ```bash
   npm install
   ```

2. **Jalankan Dev Server**:
   Jalankan perintah berikut untuk memulai server lokal:
   ```bash
   npm run dev
   ```

3. **Akses di Browser**:
   Buka peramban (browser) Anda dan akses alamat lokal yang ditampilkan di terminal (biasanya `http://localhost:5173`).

---
*Dibuat dengan dedikasi tinggi untuk performa, estetika visual premium, dan pengalaman membaca terbaik.*
