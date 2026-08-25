# Brighton Floor Time Scheduler

Aplikasi internal berbasis **Python + Streamlit** untuk menyiapkan data agent bulan sebelumnya, membuat jadwal Floor Time bulanan secara adil, menangani tanggal merah dengan keterangan hari libur, memvalidasi hasil, dan mengekspor poster mingguan ke Excel, PNG, dan PDF.

## Pembaruan utama

- Desain poster baru mengikuti gaya Brighton yang lebih clean, formal, dan corporate.
- Poster **tidak menampilkan Minggu 1/2/3/4/5/6**; hanya rentang tanggal.
- Icon kalender berulang pada setiap hari dihilangkan.
- Hari libur memakai aksen merah dan dapat diberi nama, misalnya `Maulid Nabi Muhammad SAW`.
- Hari libur tidak memiliki shift dan tidak memiliki agent.
- Data agent dibaca dari Excel/CSV dengan format utama `Nama`, `Jabatan`, `Office`, dan `Total`; kolom `Office` otomatis diringkas menjadi **Unit** dari kata kedua (contoh `Brighton Priority Cibubur, Bogor` → `Priority`).
- Tersedia filter **Minimum Total Kehadiran**; default 5 sehingga kehadiran 1–4 tidak masuk kandidat jadwal.
- Agen yang lolos filter dapat dikeluarkan khusus untuk bulan berjalan dengan menghilangkan centang **Masuk Jadwal**.
- Bulan referensi otomatis adalah bulan sebelum bulan jadwal, termasuk Desember → Januari.
- Batas Floor Time per agent per minggu dapat dipilih **1x, 2x, atau 3x**.
- Agent boleh muncul beberapa kali dalam seminggu pada **hari yang berbeda**, tetapi tidak pernah dua kali pada tanggal yang sama.
- Algoritma mengutamakan pemerataan total assignment, rotasi shift, jarak hari, dan keragaman Business Unit.
- Bila agent sedikit, sistem dapat mengulang agent sesuai batas mingguan; bila kapasitas tidak cukup, slot tetap tidak dipaksa dengan duplikasi ilegal. Notifikasi kapasitas massal tidak ditampilkan di UI agar tetap clean.
- Tab **Jadwal** menampilkan daftar bulanan berjudul `FLOOR TIME [BULAN]` sebelum poster mingguan. Kolom **Nama / Jabatan / Unit / Total Kehadiran** dapat dipilih dan Total Kehadiran dapat diurutkan default, terbesar→terkecil, atau terkecil→terbesar.
- Poster memakai ukuran font stabil. Jika agent banyak, tinggi card dan canvas bertambah ke bawah.
- Jika satu shift sangat padat, daftar agent otomatis dapat dibagi menjadi dua kolom tanpa mengecilkan font.
- Cross-month week didukung, misalnya `31 Agustus 2026 - 05 September 2026`.
- Export tersedia dalam Excel, ZIP PNG, dan PDF multi-page.
- Perubahan konfigurasi setelah generate akan menonaktifkan hasil lama agar tidak menampilkan schedule stale.

## Menjalankan aplikasi

Disarankan Python 3.10 atau lebih baru.

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Perintah `python -m streamlit` direkomendasikan khususnya pada Windows karena beberapa konfigurasi Application Control memblokir launcher `streamlit.exe` walaupun modul Streamlit sendiri dapat dijalankan.

Browser biasanya membuka:

```text
http://localhost:8501
```

## Perbaikan kompatibilitas tampilan Windows

Versi ini memaksa theme Streamlit **light** melalui `.streamlit/config.toml` agar teks widget tidak lagi berwarna terang di atas background putih. Renderer PNG/PDF juga mencari font scalable secara cross-platform (Segoe UI/Arial/Aptos pada Windows, serta fallback macOS/Linux) sehingga export tidak jatuh ke bitmap font kecil. Tidak ada file font yang dibundel di project.

## Format data agent

Data utama dimasukkan melalui upload CSV/Excel. Format yang direkomendasikan:

| No | Nama | Jabatan | Office | Total |
|---:|---|---|---|---:|
| 1 | TONI (ACYS) | Business Manager | Brighton Priority Cibubur, Bogor | 13 |
| 2 | VIRA (AUTU) | Business Manager | Brighton Warrior Cibubur, Bogor | 9 |

- `No` opsional dan akan diabaikan oleh sistem.
- `Nama` wajib. Format `NAMA (KODE)` otomatis memisahkan nama dan kode agent.
- `Jabatan` menjadi informasi profil agent.
- `Office` tetap dibaca dari file sumber, tetapi tampilan **Unit** mengambil kata kedua dari nilai Office, misalnya `Brighton Champion Cibubur, Bogor` menjadi `Champion`.
- `Total` dibaca sebagai total kehadiran bulan sebelumnya.
- Alias format lama seperti `Agent`, `Kehadiran`, dan `Business Unit` tetap dibaca untuk backward compatibility.

Setelah upload, tentukan **Minimum Total Kehadiran** lalu gunakan kolom **Masuk Jadwal** untuk memilih agen aktif. **Total Agen Bertugas** selalu mengikuti hasil filter dan pilihan tersebut.

## Hari libur

Pada tab **Konfigurasi & Generate**:

1. Pilih tanggal libur.
2. Isi nama hari libur.
3. Generate jadwal.

Output hanya menampilkan hari/tanggal dan nama peringatan dengan warna merah. Tidak ada icon X dan tidak ada tulisan besar `TUTUP` atau `KANTOR TUTUP`.

## Aturan scheduling

- Hari kerja default: Senin–Sabtu.
- Senin–Jumat:
  - 08.00–12.30
  - 12.30–17.00
- Sabtu:
  - 08.00–11.30
  - 11.30–15.00
- Agent tidak boleh muncul dua kali pada tanggal yang sama.
- Maksimum assignment per agent per minggu mengikuti pilihan 1x/2x/3x.
- Hari libur tidak dihitung sebagai slot aktif.
- Request khusus tetap tunduk pada aturan tanggal unik dan batas mingguan.
- Jika seed/kode audit sama dan data serta konfigurasi sama, hasil scheduling konsisten.

## Pengujian

Unit test dapat dijalankan dengan:

```bash
python -m unittest -v test_floor_time.py
```

<<<<<<< HEAD
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

## Catatan Operasional

Aktifkan opsi **Gunakan kode audit** bila ingin hasil generate bisa diulang dengan input yang sama. Matikan opsi tersebut bila ingin hasil random berubah setiap kali generate.

Jika ingin mengubah bentuk poster, pilih **Portrait** atau **Landscape** di sidebar sebelum menekan tombol Generate.



=======
Test mencakup cross-month, cross-year, import format Excel baru, filter kehadiran dan pengecualian manual, ekstraksi Unit, sorting daftar bulanan, hari libur, anti-double, agent sedikit dengan repeat pada hari berbeda, insufficient capacity, dynamic poster height, dan dynamic monthly roster height.
>>>>>>> 2c63294 (Update latest Floor Time Scheduler)
