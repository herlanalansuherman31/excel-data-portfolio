# Master Data Cleaning Essentials on Excel 📊

Rangkuman ini berisi 10 teknik utama pembersihan data (*data cleaning*) di Microsoft Excel berdasarkan tutorial dari kanal YouTube Career Principles.

---

## 📋 Daftar Isi
1. [Autofit Rows and Columns](#1-autofit-rows-and-columns)
2. [Find & Replace (Wildcard)](#2-find--replace-wildcard)
3. [Lower & Upper Case](#3-lower--upper-case)
4. [Trim & Proper](#4-trim--proper)
5. [Text to Columns](#5-text-to-columns)
6. [Removing Duplicates](#6-removing-duplicates)
7. [Filling Empty Cells](#7-filling-empty-cells)
8. [Handling Error dengan IFERROR](#8-handling-error-dengan-iferror)
9. [Header Formatting](#9-header-formatting)
10. [Menghilangkan Gridlines](#10-menghilangkan-gridlines)

---

## 🛠️ 10 Langkah Pembersihan Data

### 1. Autofit Rows and Columns
Memperbaiki tampilan sel yang terlalu sempit (ditandai dengan error `###`) atau baris yang terlalu lebar secara otomatis.
* **Langkah:** Blok semua data (`Ctrl + A`) -> Pilih menu **Format** -> **Autofit Row Height** & **Autofit Column Width**.
* **Shortcut:** 
  * Kolom: `Alt + H + O + I`
  * Baris: `Alt + H + O + A`

### 2. Find & Replace (Wildcard)
Menghapus bagian teks tertentu di dalam tanda kurung secara massal menggunakan simbol asterisk (`*`) sebagai *wildcard*.
* **Langkah:** `Ctrl + H` -> Pada kolom *Find*, ketik `(*)` -> Kosongkan kolom *Replace dengan* -> Klik **Replace All**.

### 3. Lower & Upper Case
Mengubah format huruf teks menjadi huruf kecil semua atau huruf besar semua.
* **Rumus:** `=LOWER(Cell)` atau `=UPPER(Cell)`.
* **Tips Penting:** Setelah mengubah teks, salin (*Copy*) lalu tempel sebagai nilai (**Paste as Values** / `Alt + H + V + V`) agar data asli bisa dihapus tanpa merusak rumus.

### 4. Trim & Proper
Kombinasi fungsi untuk menghapus spasi ganda/berlebih sekaligus membuat huruf kapital hanya pada awal kata.
* **Rumus:** `=TRIM(PROPER(Cell))`

### 5. Text to Columns
Memisahkan satu kolom data menjadi beberapa kolom berdasarkan tanda pembatas tertentu (misal: tanda *underscore* `_`).
* **Langkah:** Blok kolom -> Masuk ke tab **Data** -> Pilih **Text to Columns** -> Pilih **Delimited** -> Centang **Other** dan isi dengan `_` -> Tentukan lokasi tujuan (*Destination*) -> **Finish**.

### 6. Removing Duplicates
Menghapus baris data yang terduplikasi secara otomatis untuk memastikan validitas data.
* **Langkah:** `Ctrl + A` -> Masuk ke tab **Data** -> Klik ikon **Remove Duplicates** -> Pastikan opsi *My data has headers* dicentang -> Klik **OK**.

### 7. Filling Empty Cells
Mengisi semua sel kosong secara massal dengan teks tertentu (misalnya "N/A") tanpa perlu mengetiknya satu per satu.
* **Langkah:** `Ctrl + A` -> Tab **Home** -> **Find & Select** -> **Go To Special** -> Pilih **Blanks** -> Klik **OK**. 
* **Eksekusi:** Ketik teks `N/A` pada formula bar, lalu tekan **`Ctrl + Enter`** (bukan Enter biasa).

### 8. Handling Error dengan IFERROR
Menyembunyikan pesan error (seperti `#VALUE!`) akibat pembagian angka dengan teks, dan menggantinya dengan tampilan yang lebih rapi.
* **Rumus:** `=IFERROR(Rumus_Utama; "N/A")`

### 9. Header Formatting
Memberikan visualisasi yang jelas pada baris judul data agar terlihat profesional.
* **Langkah:** Blok baris header (`Ctrl + Shift + Panah Kanan`) -> Tebalkan teks (`Ctrl + B`) -> Ubah warna latar belakang menjadi biru tua dan teks menjadi putih.

### 10. Menghilangkan Gridlines
Menghapus garis bantu bawaan Excel agar tampilan data akhir terlihat lebih bersih dan siap dipresentasikan.
* **Langkah:** Masuk ke tab **View** -> Hapus centang pada pilihan **Gridlines**.
* **Shortcut:** `Alt + W + V + G`

---
*Catatan: Tutorial ini disarikan dari video Career Principles (2026).*
