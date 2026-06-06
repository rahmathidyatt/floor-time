# Brighton Floor Time Schedule Generator

Aplikasi web berbasis Python Streamlit untuk membantu Staff Operasional Brighton membuat jadwal Floor Time bulanan secara fleksibel, rapi, tervalidasi, dan siap dipakai untuk kebutuhan operasional.

## Fitur Utama

- Input daftar agen fleksibel, jumlah agen bisa berubah setiap bulan.
- Generate jadwal per bulan dan otomatis dikelompokkan menjadi jadwal mingguan.
- Minggu terakhir otomatis diteruskan sampai hari Sabtu walaupun masuk bulan berikutnya, sehingga tidak ada jadwal yang menggantung hanya 1-2 hari.
- Pilih tanggal merah atau tanggal kantor tutup sebelum jadwal digenerate.
- Setiap agen hanya mendapat 1 jadwal per minggu.
- Randomisasi menyeluruh dengan opsi kode audit agar hasil bisa diulang bila diperlukan.
- Request urgent agen untuk tanggal dan shift tertentu.
- Pilihan orientasi poster: Portrait atau Landscape.
- Preview poster mingguan berupa gambar stabil, bukan HTML/CSS yang mudah berubah di browser.
- Export hanya dalam format Excel, gambar PNG, dan PDF.

## Cara Instalasi

Pastikan Python sudah terinstall. Disarankan Python 3.10 atau lebih baru.

```bash
pip install -r requirements.txt
```

## Cara Menjalankan

```bash
streamlit run app.py
```

Setelah itu browser akan membuka aplikasi secara otomatis. Bila tidak terbuka, lihat URL yang muncul di terminal, biasanya:

```text
http://localhost:8501
```

## Format Input Agen

Masukkan satu agen per baris, contoh:

```text
PAULA (DUCC)
AZWAR (IBEX)
MONIKA (NDEP)
```

Atau upload CSV/Excel. Bila memakai CSV/Excel, aplikasi akan mencari kolom bernama `Agen`, `Nama`, atau `Nama Agen`. Jika tidak ada, kolom pertama akan digunakan.

## Aturan Jadwal

- Senin sampai Jumat memiliki shift:
  - Pagi: 08.00-12.30
  - Siang: 12.30-17.00
- Sabtu memiliki shift:
  - Pagi: 08.00-11.30
  - Siang: 11.30-15.00
- Minggu otomatis tidak dijadwalkan.
- Tanggal merah yang dipilih tidak akan dijadwalkan.
- Agen tidak boleh double dalam minggu yang sama.
- Request urgent diprioritaskan, tetapi tetap mengikuti aturan 1 agen hanya 1 jadwal per minggu.

## Output

Aplikasi menyediakan 3 output utama:

1. **Excel** untuk arsip dan pengecekan data detail.
2. **Gambar PNG** dalam bentuk ZIP karena setiap minggu dibuat menjadi poster terpisah.
3. **PDF** multi-page yang berisi seluruh poster mingguan.

## Catatan Tampilan

- Logo Brighton pada poster menggunakan huruf **O** berwarna putih.
- Credit **Created by rh** hanya tampil di halaman aplikasi bagian bawah tengah, bukan pada output Excel, PNG, maupun PDF.

## Catatan Operasional

Aktifkan opsi **Gunakan kode audit** bila ingin hasil generate bisa diulang dengan input yang sama. Matikan opsi tersebut bila ingin hasil random berubah setiap kali generate.

Jika ingin mengubah bentuk poster, pilih **Portrait** atau **Landscape** di sidebar sebelum menekan tombol Generate.


## Catatan deploy Streamlit

Versi ini memakai pencarian font sistem yang lebih luas agar tampilan poster tetap normal saat dijalankan di Streamlit Community Cloud.
