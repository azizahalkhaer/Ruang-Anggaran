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

- Signup dan login demo dengan pemeriksaan awal wajah pada foto kartu dan selfie serta OCR nomor pegawai. Model wajah/OCR dijalankan di browser; hasil cocok/tidak cocok saja yang diteruskan ke Streamlit. Akun dan hash kata sandi disimpan di SQLite lokal sehingga tetap tersedia setelah browser ditutup.
- Unggah Excel (`.xlsx`, `.xls`, `.xlsm`), CSV, Word (`.docx`), atau PDF berbasis teks hingga 20 MB. PDF hasil pindai perlu OCR terlebih dahulu. Format Word `.doc` lama belum didukung pada versi Streamlit.
- Preprocessing memilih header uraian dan kolom anggaran, membersihkan baris kosong/tidak bernilai, serta menampilkan baris hasil olahan.
- Grafik batang horizontal untuk membandingkan pagu dan sisa per kategori, komposisi pagu, temuan otomatis, dan pemeriksaan rincian `Semula` → `Menjadi`, termasuk kecukupan `Sisa Anggaran` terhadap `Selisih`.
- Generator TOR berbasis aturan lokal dengan ringkasan, isi lengkap yang dapat dibuka, dan unduhan `.docx`. Hasilnya draft yang perlu ditinjau.
- Logo dan palet warna Balmon, DJID, serta Komdigi.

## Deploy untuk demo

Untuk Streamlit Community Cloud, push folder aplikasi ke GitHub, pilih `streamlit_app.py` sebagai file utama, lalu deploy. Pastikan `requirements.txt`, folder `identity_component/`, folder `image/`, dan file `.streamlit/config.toml` ikut masuk ke repository. Gunakan HTTPS agar kamera dapat meminta izin browser. Jangan commit database `.data/accounts.sqlite3`, data RAB, atau foto kartu/selfie.

## Batasan privasi dan keamanan

File RAB diunggah ke proses server Streamlit untuk diparsing dan tidak sengaja disimpan ke disk atau dikirim ke layanan AI eksternal. Foto kartu dan selfie hanya diproses di browser, tetapi aplikasi demo perlu menjalankan JavaScript/model pada perangkat pengguna. Verifikasi tidak membuktikan keaslian kartu, liveness, atau status pegawai. Akun demo disimpan dalam SQLite lokal di server, terpisah dari session browser, sehingga pengguna dapat menutup browser dan login kembali. Pada hosting dengan penyimpanan sementara seperti Streamlit Community Cloud, akun tetap ada selama file database tersedia, tetapi dapat hilang saat server dibuat ulang atau aplikasi dipindah/redeploy; untuk penyimpanan akun yang tahan restart gunakan database persisten terkelola. Ini bukan sistem autentikasi atau akses pegawai sungguhan. Jangan gunakan aplikasi demo untuk melindungi data rahasia atau memublikasikan dokumen internal. Analisis dan isi TOR dibuat dengan aturan lokal, bukan AI eksternal. Periksa kembali perhitungan, dasar hukum, keluaran, jadwal, dan persetujuan sebelum penggunaan resmi.
