# 25+ Rumus Excel Paling Penting untuk Produktivitas 🚀

Rangkuman ini berisi panduan lengkap penggunaan 25 rumus wajib Excel berdasarkan video tutorial dari kanal YouTube Ignasius Ryan. Materi dibagi menjadi 5 bagian utama berdasarkan fungsi dan studi kasusnya.

---

## 📋 Daftar Isi Materi

### 🔹 Bagian 1: Formula Dasar & Kalkulasi Kriteria
Digunakan untuk menghitung total, rata-rata, logika kelulusan, peringkat, dan kalkulasi dengan kriteria tertentu.
*   **SUM**: Menghitung total nilai dari sekumpulan data.
    > `=SUM(Bulan1:Bulan3)`
*   **AVERAGE**: Menghitung rata-rata nilai.
    > `=AVERAGE(Bulan1:Bulan3)`
*   **IF**: Menghasilkan nilai tertentu berdasarkan uji logika (misal: menentukan Kelulusan jika omset > 30 juta).
    > `=IF(Total_Omset > 30000000; "Lulus"; "Gagal")`
*   **RANK**: Menentukan peringkat atau ranking suatu nilai di dalam kelompok data.
    > `=RANK(Sel_Nilai; Rentang_Kelompok_Data)` *(Gunakan F4 untuk mengunci rentang data)*.
*   **COUNT**: Menghitung jumlah sel yang berisi data berupa angka.
    > `=COUNT(Rentang_Angka)`
*   **COUNTA**: Menghitung jumlah sel yang tidak kosong (bisa berupa teks maupun angka).
    > `=COUNTA(Rentang_Teks)`
*   **MAX**: Mencari nilai tertinggi dari sekumpulan data.
    > `=MAX(Rentang_Data)`
*   **MIN**: Mencari nilai terendah dari sekumpulan data.
    > `=MIN(Rentang_Data)`
*   **SUMIF**: Menjumlahkan nilai hanya jika memenuhi satu kriteria tertentu (misal: total omset khusus karyawan Laki-laki).
    > `=SUMIF(Rentang_Kriteria; "L"; Rentang_Nilai_Yang_Dijumlahkan)`.
*   **COUNTIF**: Menghitung jumlah sel berdasarkan kriteria tertentu (misal: menghitung berapa pegawai yang Lulus).
    > `=COUNTIF(Rentang_Data; "Lulus")`.

---

### 🔹 Bagian 2: Manipulasi & Pembersihan Teks
Digunakan untuk merapikan format penulisan teks atau data string yang berantakan.
*   **TRIM**: Menghapus spasi berlebih di awal, tengah, atau akhir teks.
    > `=TRIM(Sel_Teks)`
*   **LOWER**: Mengubah semua huruf menjadi huruf kecil.
    > `=LOWER(Sel_Teks)`
*   **UPPER**: Mengubah semua huruf menjadi huruf kapital/besar.
    > `=UPPER(Sel_Teks)`
*   **PROPER**: Mengubah huruf pertama di setiap kata menjadi huruf kapital.
    > `=PROPER(Sel_Teks)`

---

### 🔹 Bagian 3: Ekstraksi & Penggabungan Karakter
Digunakan untuk memanipulasi teks, membuat ID unik, atau mengambil sebagian karakter dari kode tertentu.
*   **CONCATENATE**: Menggabungkan teks dari beberapa sel menjadi satu.
    *   *Contoh menggabungkan dengan tanda strip (-):* `=CONCATENATE(Divisi; "-"; Nama; "-"; Tahun)`.
*   **LEN**: Menghitung jumlah total karakter (termasuk spasi dan tanda baca) dalam satu sel.
    > `=LEN(Sel_Teks)`
*   **LEFT**: Mengambil sejumlah karakter dari sisi paling kiri teks.
    > `=LEFT(Sel_Teks; Jumlah_Karakter)`
*   **MID**: Mengambil karakter dari posisi tengah teks.
    > `=MID(Sel_Teks; Posisi_Mulai; Jumlah_Karakter)`
*   **RIGHT**: Mengambil sejumlah karakter dari sisi paling kanan teks.
    > `=RIGHT(Sel_Teks; Jumlah_Karakter)`

---

### 🔹 Bagian 4: Pembulatan Angka & Kelipatan
Membagi teknik pembulatan menjadi dua kategori: desimal matematika dan kelipatan logistik.

#### 1. Pembulatan Nilai Desimal
*   **ROUNDDOWN**: Memaksa pembulatan angka ke bawah sesuai digit desimal yang ditentukan.
    > `=ROUNDDOWN(Angka; 0)` *(Angka 0 berarti tanpa koma)*.
*   **ROUNDUP**: Memaksa pembulatan angka ke atas.
    > `=ROUNDUP(Angka; 0)`
*   **ROUND**: Membulatkan angka secara otomatis mengikuti aturan matematika terdekat (>= 0.5 naik, < 0.5 turun).
    > `=ROUND(Angka; 0)`

#### 2. Pembulatan Nilai Kelipatan (Studi Kasus Order Stok Grosir)
*   **FLOOR**: Membulatkan angka ke bawah ke kelipatan terdekat yang ditentukan.
    > `=FLOOR(Jumlah_Stok; 100)` *(Membulatkan ke bawah kelipatan 100)*.
*   **CEILING**: Membulatkan angka ke atas ke kelipatan terdekat yang ditentukan.
    > `=CEILING(Jumlah_Stok; 100)`
*   **MROUND**: Membulatkan angka ke bawah atau ke atas ke kelipatan terdekat secara otomatis.
    > `=MROUND(Jumlah_Stok; 100)`

---

### 🔹 Bagian 5: Rumus Pencarian (LookUp) & Penanganan Error
Teknik krusial untuk menghubungkan satu tabel utama dengan tabel referensi data lainnya.

*   **VLOOKUP**: Mencari data secara vertikal dari tabel referensi (judul kolom berada di atas).
    > `=VLOOKUP(Nilai_Kunci; Rentang_Tabel_Referensi; Nomor_Kolom; 0)` *(Jangan lupa kunci tabel referensi dengan F4)*.
*   **HLOOKUP**: Mencari data secara horizontal dari tabel referensi (judul baris berada di samping kiri).
    > `=HLOOKUP(Nilai_Kunci; Rentang_Tabel_Referensi; Nomor_Baris; 0)`.
*   **IFERROR (Bonus Profesional)**: Menyamarkan atau mengganti tampilan bawaan error Excel seperti `#N/A` atau `#VALUE!` menjadi teks kustom yang rapi.
    > `=IFERROR(Rumus_VLOOKUP_Anda; "Data Tidak Tersedia")`.

---
*Catatan: File latihan untuk materi di atas dapat diakses pada tautan resmi di deskripsi video YouTube Ignasius Ryan.*
