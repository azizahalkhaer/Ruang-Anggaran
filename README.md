# Ruang Anggaran — Dashboard RAB & Generator TOR

Demo aplikasi berbasis Streamlit untuk membaca RAB, menampilkan ringkasan dan grafik, membandingkan perubahan anggaran, serta membuat draft TOR dalam format Word.

## Jalankan demo Streamlit

Gunakan Python 3.9 atau lebih baru:

```sh
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Buka alamat lokal yang ditampilkan Streamlit (biasanya `http://localhost:8501`). Untuk menjalankan aplikasi lama berbasis Express, gunakan `npm install` lalu `npm start`; versi Streamlit tidak membutuhkan server Express.

## Fitur

- Login dan signup hanya menggunakan email dan kata sandi; pemeriksaan kartu, selfie, dan OCR tidak digunakan. Akun dan hash kata sandi disimpan di SQLite lokal; akun yang sudah dibuat tetap dapat dipakai login dari jaringan mana pun.
- Signup dibatasi ke jaringan kantor melalui rentang IP/CIDR yang dikonfigurasi admin di `SIGNUP_ALLOWED_CIDRS`. Browser tidak mengizinkan aplikasi web membaca nama Wi-Fi/SSID, jadi nama jaringan seperti `BALMON LT2 01-5G` tidak bisa dibedakan langsung oleh aplikasi. Untuk mengecualikan `BALMON LOBBY TAMU`, masukkan hanya rentang IP/VLAN SSID internal Balmon yang diizinkan dan jangan masukkan rentang tamu. Bila belum dikonfigurasi, signup ditolak (fail closed); login akun yang sudah ada tetap tersedia.
- Unggah Excel (`.xlsx`, `.xls`, `.xlsm`), CSV, Word (`.docx`), atau PDF berbasis teks hingga 20 MB. PDF hasil pindai perlu OCR terlebih dahulu. Format Word `.doc` lama belum didukung pada versi Streamlit.
- Unggah bersama PDF **Rincian Kertas Kerja Satker** dan **Laporan FA Detail** untuk memasangkan kode RO. Dashboard mengambil pagu penuh dan pagu terblokir dari Rincian Kertas Kerja (baris `Jumlah Biaya` bertanda `*`), serta pagu efektif, realisasi kumulatif, dan sisa dari Laporan FA. Baris yang tidak punya pasangan tetap ditampilkan dengan statusnya.
- Preprocessing memilih header uraian dan kolom anggaran, membersihkan baris kosong/tidak bernilai, serta menampilkan baris hasil olahan.
- Grafik batang horizontal untuk membandingkan pagu dan sisa per kategori, komposisi pagu, temuan otomatis, dan pemeriksaan rincian `Semula` → `Menjadi`, termasuk kecukupan `Sisa Anggaran` terhadap `Selisih`.
- Generator TOR berbasis aturan lokal; untuk pasangan PDF, pilih kode RO dan masukkan ke metadata TOR bersama ringkasan pagu/realisasi. Dokumen Word menggunakan format A4, Times New Roman 11, spasi 1,5, paragraf rata kiri-kanan, dan heading biru tua. Hasilnya draft yang perlu ditinjau.
- Logo dan palet warna Balmon, DJID, serta Komdigi.

## Deploy untuk demo

Untuk Streamlit Community Cloud, push folder aplikasi ke GitHub, pilih `streamlit_app.py` sebagai file utama, lalu deploy. Pastikan `requirements.txt`, folder `image/`, dan file `.streamlit/config.toml` ikut masuk ke repository. Untuk mengaktifkan signup, tambahkan daftar rentang IP/VLAN jaringan internal kantor sebagai `SIGNUP_ALLOWED_CIDRS` di secrets/config aplikasi, misalnya `10.20.30.0/24,198.51.100.7/32` (ganti dengan rentang jaringan kantor yang sebenarnya; jangan gunakan contoh ini apa adanya). Proxy tepercaya harus menimpa header `X-Forwarded-For` dengan IP klien dan tidak boleh membiarkan klien menyuplai header tersebut sendiri. SSID kantor dan SSID tamu harus berada di rentang jaringan berbeda agar pengecualian jaringan tamu dapat ditegakkan.

## Batasan privasi dan keamanan

File RAB diunggah ke proses server Streamlit untuk diparsing dan tidak sengaja disimpan ke disk atau dikirim ke layanan AI eksternal. Akun demo disimpan dalam SQLite lokal di server, terpisah dari session browser, sehingga pengguna dapat menutup browser dan login kembali. Pada hosting dengan penyimpanan sementara seperti Streamlit Community Cloud, akun tetap ada selama file database tersedia, tetapi dapat hilang saat server dibuat ulang atau aplikasi dipindah/redeploy; untuk penyimpanan akun yang tahan restart gunakan database persisten terkelola. Pembatasan signup mengandalkan alamat IP yang diteruskan proxy tepercaya, bukan SSID; jika proxy/rentang jaringan salah konfigurasi, pembatasan tidak menjamin perangkat benar-benar menggunakan Wi-Fi internal. Ini bukan sistem autentikasi atau akses pegawai sungguhan. Jangan gunakan aplikasi demo untuk melindungi data rahasia atau memublikasikan dokumen internal. Analisis dan isi TOR dibuat dengan aturan lokal, bukan AI eksternal. Periksa kembali perhitungan, dasar hukum, keluaran, jadwal, dan persetujuan sebelum penggunaan resmi.
