# Brighton Floor Time Scheduler

Aplikasi internal berbasis **Python + Streamlit** untuk menyiapkan data agent bulan sebelumnya, membuat jadwal Floor Time bulanan secara adil, menangani tanggal merah dengan keterangan hari libur, memvalidasi hasil, dan mengekspor poster mingguan ke Excel, PNG, dan PDF.

## Pembaruan utama

- Desain poster mengikuti referensi corporate premium: teks hitam kontras tinggi, aksen emas Brighton, hierarki tegas, card dengan shadow halus, serta time band yang bersih.
- Header poster dibuat lebih rapat, ornamen sudut dihilangkan, dan rentang tanggal tampil tanpa bingkai agar fokus langsung menuju jadwal.
- Hari Sabtu memakai latar kuning lembut untuk menandai jam operasional khusus. Catatan hanya mengikuti isian pengguna; tidak ada keterangan tambahan otomatis.
- Font daftar agen tidak diperkecil saat jumlah agen bertambah; card/canvas bertambah tinggi berdasarkan hasil pembungkusan nama dan setiap shift tetap memakai satu kolom nama.
- Poster **tidak menampilkan Minggu 1/2/3/4/5/6**; hanya rentang tanggal.
- Icon kalender berulang pada setiap hari dihilangkan.
- Hari libur memakai aksen merah dan dapat diberi nama, misalnya `Maulid Nabi Muhammad SAW`.
- Hari libur tidak memiliki shift dan tidak memiliki agent.
- Data agent dibaca dari Excel/CSV dengan format utama `Nama`, `Jabatan`, `Office`, dan `Total`; kolom `Office` otomatis diringkas menjadi **Unit** dari kata kedua (contoh `Brighton Priority Cibubur, Bogor` → `Priority`).
- Tersedia filter **Minimum Total Kehadiran**; default 5 sehingga kehadiran 1–4 tidak masuk kandidat jadwal.
- Agen yang lolos filter dapat dikeluarkan khusus untuk bulan berjalan dengan menghilangkan centang **Masuk Jadwal**.
- Bulan referensi otomatis adalah bulan sebelum bulan jadwal, termasuk Desember → Januari.
- Target Floor Time setiap agen per minggu dapat dipilih **1x, 2x, atau 3x**.
- Agent boleh muncul beberapa kali dalam seminggu pada **hari yang berbeda**, tetapi tidak pernah dua kali pada tanggal yang sama.
- Algoritma mengutamakan pemerataan total assignment, rotasi shift, jarak hari, dan keragaman Business Unit.
- Kapasitas shift otomatis diperluas untuk memenuhi target semua agen. Jika hari aktif kurang dari target, setiap agen dijadwalkan sebanyak hari aktif dan peringatan ditampilkan.
- Tab **Jadwal** menampilkan daftar bulanan berjudul `DAFTAR AGEN FLOOR TIME` sebelum poster mingguan. Kolom **Nama / Jabatan / Unit / Total Kehadiran** dapat dipilih dan Total Kehadiran dapat diurutkan default, terbesar→terkecil, atau terkecil→terbesar.
- Daftar bulanan otomatis dibagi maksimal **20 agen per halaman**. Daftar 50 agen menjadi 3 halaman sehingga tulisan tidak dipaksa mengecil dan lebih nyaman dibaca dari ponsel.
- Informasi **jumlah agen dan nomor halaman** ditempatkan pada footer daftar bulanan; bar kuning dekoratif di bagian bawah dihapus.
- Poster memakai ukuran font stabil. Jika agent banyak, tinggi card dan canvas bertambah ke bawah.
- Tinggi area catatan mengikuti jumlah dan panjang catatan, sehingga tidak menyisakan ruang kosong berlebihan.
- Jika satu shift sangat padat, tinggi daftar bertambah tanpa membagi nama ke kolom sempit.
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

Versi ini memaksa theme Streamlit **light** melalui `.streamlit/config.toml` agar teks widget tidak lagi berwarna terang di atas background putih. Renderer PNG/PDF juga mencari font scalable secara cross-platform (Segoe UI/Arial/Aptos pada Windows, serta fallback macOS/Linux) sehingga export tidak jatuh ke bitmap font kecil. Font DejaVu Sans regular dan bold disertakan dalam assets/fonts dan wajib ikut diunggah.

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

Output hanya menampilkan hari/tanggal dan nama peringatan; teks tetap hitam dengan latar aksen hari libur agar konsisten dengan standar keterbacaan. Tidak ada icon X dan tidak ada tulisan besar `TUTUP` atau `KANTOR TUTUP`.

## Aturan scheduling

- Hari kerja default: Senin–Sabtu.
- Senin–Jumat:
  - 08.00–12.30
  - 12.30–17.00
- Sabtu:
  - 08.00–11.30
  - 11.30–15.00
- Agent tidak boleh muncul dua kali pada tanggal yang sama.
- Setiap agen memperoleh tepat 1x/2x/3x per minggu sesuai pilihan, selama jumlah hari aktif mencukupi.
- Hari libur tidak dihitung sebagai slot aktif.
- Request khusus tetap tunduk pada aturan tanggal unik dan batas mingguan.
- Jika seed/kode audit sama dan data serta konfigurasi sama, hasil scheduling konsisten.

## Pengujian

Unit test dapat dijalankan dengan:

```bash
python -m unittest -v test_floor_time.py
```

Test mencakup cross-month, cross-year, import format Excel baru, filter kehadiran dan pengecualian manual, ekstraksi Unit, sorting dan pagination daftar bulanan, catatan sesuai isian pengguna, hari libur, anti-double, agent sedikit dengan repeat pada hari berbeda, insufficient capacity, serta dynamic poster height.

## Revisi keterbacaan 50 agen

- Seluruh teks pada poster dan daftar bulanan menggunakan warna hitam agar kontras tinggi.
- Huruf **o** pada logo Brighton tetap kuning sebagai identitas logo.
- Ukuran nama agen, jam shift, hari, tanggal, catatan, dan tabel bulanan diperbesar.
- Ukuran font tidak pernah diturunkan ketika jumlah agen bertambah; tinggi card/canvas bertambah mengikuti isi.
- Setiap shift tetap memakai satu kolom; tinggi poster mengikuti jumlah agen dengan font tetap 24 px tebal.
- Hari Sabtu menggunakan latar kuning lembut karena jam operasionalnya berbeda.
- Daftar agen bulanan memakai font 23 px, baris lega, dan pagination 20 agen per halaman agar mudah dibaca oleh agen senior atau pengguna berkacamata.

## Perbaikan keterbacaan terbaru
- Jarak logo dan nama kantor dihitung dari batas huruf dengan ruang 18 px.
- Nama agen dan jam mingguan berukuran 24 px; nama menggunakan huruf tebal.
- Tinggi baris dan poster otomatis mengikuti isi tanpa mengecilkan huruf.
- Catatan kosong tidak menampilkan kotak CATATAN.

## Pembagian merata
49 agen pada target 2x menghasilkan 98 penugasan per minggu; target 3x menghasilkan 147 penugasan. Request khusus dihitung sebagai bagian dari target. Satu agen tidak mendapat dua shift pada tanggal yang sama.

## Perbaikan font untuk deploy
Renderer PNG/PDF memakai assets/fonts/DejaVuSans.ttf dan DejaVuSans-Bold.ttf secara langsung, relatif terhadap app.py. Font sistem operasi tidak lagi menentukan hasil. Kedua font beserta LICENSE.txt wajib ikut disalin dan di-push ke GitHub. Jika aset hilang/rusak, aplikasi memberi pesan jelas, bukan diam-diam memakai font tanpa bold. Footer memakai titik tengah dengan font bawaan: `50 AGEN • HALAMAN 2/3`.

```powershell
git add app.py assets/fonts README.md test_floor_time.py
git commit -m "Fix deployed poster fonts and roster footer"
git push origin main
```
Setelah deploy selesai, generate ulang jadwal untuk membuat gambar baru.

## Nama panjang
Kolom shift dan lebar poster mingguan mengikuti nama lengkap terpanjang, termasuk kode agen, agar tetap satu baris pada ukuran font 24 px. Footer menggunakan titik tengah •.
