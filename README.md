# Material Tracker

**Material Tracker** adalah aplikasi desktop berbasis Python (Qt) untuk mencatat, melacak, dan mengelola log material, aksesori, dan perangkat ONT di lingkungan gudang/teknisi, khususnya untuk tim CKT Purwokerto.

## Fitur

- **Pencatatan Pengambilan Material**: Catat pengambilan kabel, aksesori, dan serial number ONT dengan mudah.
- **Manajemen Stock**: Tambah, lihat, dan hapus data stok masuk serta pantau distribusi material ke teknisi.
- **Import/Export CSV**: Fitur impor data dari file CSV dan ekspor laporan ke file CSV.
- **Laporan Telegram**: Tampilkan laporan dari berbagai sumber dalam bentuk tab (MyRepublic, Asianet, Oxygen).
- **Pengaturan Master Data**: Kelola daftar Divisi, Tim, Material, dan Aksesoris melalui dialog pengaturan.
- **Tema Terang & Gelap**: Pilihan tampilan antarmuka terang/gelap.
- **Penyimpanan Lokal Otomatis**: Semua data disimpan di folder `~/.material_tracker` (tidak butuh database eksternal).

## Instalasi & Menjalankan

### 1. Clone Repository

```sh
git clone https://github.com/endans/material-tracker.git
cd material-tracker
```

### 2. Install Dependencies

Pastikan Anda sudah menginstall Python 3.7+ dan pip.

```sh
pip install -r requirements.txt
```

### 3. Jalankan Aplikasi

```sh
python main.py
```

> Jika ingin membuat shortcut atau icon, Anda dapat menambahkannya secara manual sesuai OS masing-masing.

## Struktur Data

- Semua data tersimpan dalam file `.json` di folder `~/.material_tracker/` pada home user.
- Laporan Telegram dapat di-load dari file CSV pada folder `~/Reports/`.

## Penggunaan

1. **Pengambilan**: Input pengambilan material/ONT, pilih atau tambahkan Divisi/Tim, dan simpan data.
2. **Resume**: Lihat rekap stok, histori pengambilan material, serta status ONT (terpakai/kosong).
3. **Laporan**: Tampilkan laporan dari file CSV (MyRepublic, Asianet, Oxygen).
4. **Pengaturan**: Kelola daftar Divisi, Tim, Material, dan Aksesori dari menu Preferences > Pengaturan.

## Catatan

- Data **tidak** terhubung ke server/cloud, hanya lokal.
- Untuk fitur laporan, pastikan file CSV sesuai format yang didukung (lihat contoh di aplikasi).
- Gunakan fitur Import untuk menambah data secara masal dari CSV.

## Lisensi

MIT License

---

**Developed for CKT Purwokerto**  
https://github.com/endans/material-tracker