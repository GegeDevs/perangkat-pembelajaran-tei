---
name: perangkat-pembelajaran-tei
description: Buat perangkat pembelajaran TEI lengkap dan tervalidasi.
version: 1.0.0
author: Gege Desembri (GegeDevs), Hermes Agent
license: MIT
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [pendidikan, TEI, SMK, perangkat-pembelajaran, DOCX, PDF]
    related_skills: [docx, pdf]
---

# Perangkat Pembelajaran TEI

Gunakan skill ini ketika guru meminta perangkat pembelajaran Dasar-Dasar Teknik Elektronika atau Teknik Elektronika Industri dalam format Kurikulum Merdeka. Workflow ini menghasilkan perangkat yang konsisten, dapat diedit, dan diverifikasi secara nyata.

## Kapan Digunakan

- Saat membuat RPP/modul ajar/bahan ajar/LKPD untuk kelas X–XII TEI.
- Saat menyusun perangkat supervisi pembelajaran satu pertemuan.
- Saat membutuhkan DOCX dan PDF dengan kunci jawaban terpisah.

## Susunan Dokumen Standar

Hasil minimal terdiri dari lima dokumen:

1. **RPP** — identitas, CP/TP, model PBL, langkah pembelajaran, asesmen, pengayaan/remedial, pengesahan.
2. **Bahan Ajar** — tujuan, teori dasar, materi per bab, tabel referensi, ilustrasi, contoh, rangkuman, keselamatan, latihan.
3. **Modul Ajar** — informasi umum, komponen inti, kegiatan PBL, asesmen, diferensiasi, refleksi, lampiran.
4. **LKPD** — identitas, tujuan, masalah, kegiatan pengamatan/perhitungan, tabel jawaban, presentasi, refleksi; **tanpa kunci jawaban**.
5. **Kunci Jawaban** — file terpisah dari LKPD, berisi jawaban, langkah perhitungan, rubrik/pedoman skor.

Setiap dokumen disediakan dalam `.docx` dan `.pdf`. Paket ZIP memuat seluruh dokumen dan media pendukung.

## Pola Bahan Ajar

Gunakan struktur bab berikut agar materi tidak bercampur:

- Pendahuluan — tujuan pembelajaran.
- **Bab I — Teori Dasar Kelistrikan**: Hukum Ohm, daya, Kirchhoff, muatan/energi kapasitor, konstanta waktu RC.
- **Bab II — Resistor**: pengertian/fungsi, ilustrasi, kode warna, tabel warna dengan latar sel sesuai warna, contoh pembacaan dan gambar resistor.
- **Bab III — Kapasitor**: pengertian yang diuraikan, struktur pelat-dielektrik, fungsi, satuan, polaritas, tabel kode 3 digit, ilustrasi polar/non-polar dan simbol.
- **Bab IV — Rangkaian Seri dan Paralel**: rumus R/C dan ilustrasi yang koneksinya tidak terputus.
- Contoh soal, rangkuman/keselamatan, latihan mandiri.

Gunakan Bahasa Indonesia, A4, Arial bila tersedia atau Liberation Sans sebagai fallback kompatibel. Untuk dokumen sekolah user ini, gunakan identitas guru/sekolah dari permintaan aktif dan jangan mengarang data.

## Prosedur

1. **Kumpulkan spesifikasi**: sekolah, guru, kelas/fase, mapel, tanggal, durasi, materi, CP/TP, model, dan format keluaran. Jika referensi ATP tersedia, baca dan kutip elemen/TP yang sesuai.
2. **Siapkan workspace** di `~/hermes/<nama-proyek>/`. Simpan generator, media, DOCX, PDF, dan ZIP di sana.
3. **Bangun generator** dengan `python-docx` dan fungsi layout yang konsisten. Gunakan gambar PNG yang dibuat dengan Python; jangan mengandalkan screenshot.
4. **Pisahkan konten**: LKPD tidak boleh memuat bagian “Kunci Jawaban Guru”; buat kunci sebagai dokumen sendiri.
5. **Buat media pendukung**: resistor bergelang, gambar contoh gelang untuk tabel, kapasitor polar/non-polar dengan simbol, dan rangkaian seri-paralel. Periksa visual dengan `vision_analyze` bila tersedia.
6. **Generate DOCX**, lalu konversi ke PDF memakai LibreOffice (`soffice`). Jangan menyatakan selesai sebelum kedua format benar-benar ada.
7. **Validasi** setiap DOCX dengan `docx_validate.py`, baca teks/struktur dengan `docx_read.py`, periksa PDF jumlah halaman/teks, dan uji ZIP dengan `ZipFile.testzip()`.
8. **Salin deliverable Telegram** ke `${HERMES_HOME:-$HOME/.hermes}/document_cache/` sebelum memakai directive `MEDIA:`.
9. **Laporkan singkat**: file yang berubah, validasi, dan hal yang masih tersisa.

## Aturan Visual dan Isi

- Tabel gelang warna harus memakai background warna pada kolom Warna; gunakan teks putih pada warna gelap.
- Tabel contoh gelang harus menyertakan gambar resistor, bukan hanya nama urutan warna.
- Simbol kapasitor polar menampilkan tanda `+`; kapasitor non-polar tidak menampilkan polaritas.
- Kabel ilustrasi kapasitor harus menyentuh pelat pertama dan keluar dari pelat kedua.
- Label komponen dan rumus tidak boleh bertumpuk; beri ruang vertikal yang cukup.
- Pertahankan font sekolah: Arial atau fallback metriks kompatibel.
- Jangan memasukkan Induktor jika materi aktif hanya Resistor dan Kapasitor.
- Jangan menyimpan kredensial, token, password, atau API key di artefak; gunakan `[REDACTED]` bila perlu.

## Verifikasi

Checklist selesai:

- Lima DOCX ada dan lolos `docx_validate.py`.
- Lima PDF ada dan dapat dibaca.
- Bahan Ajar memuat Bab I–IV, teori relevan, tabel warna, gambar contoh resistor, dan gambar kapasitor polar/non-polar.
- LKPD tidak memuat kunci jawaban.
- Kunci jawaban tersedia sebagai DOCX/PDF terpisah.
- Media tertanam di DOCX yang relevan.
- ZIP lolos `testzip()` dan berisi semua DOCX, PDF, serta PNG.
- File yang dikirim berasal dari `document_cache`, bukan path workspace yang mungkin diblokir.

## Pitfall

- Jangan menyebut paket final sebelum generator, konversi PDF, validasi DOCX, dan uji ZIP selesai.
- Jangan menaruh jawaban guru di akhir LKPD.
- Jangan hanya mengubah PNG tanpa regenerasi DOCX/PDF; media lama tetap tertanam.
- Jangan menggunakan `/root/...` langsung pada `MEDIA:`; salin ke `document_cache` terlebih dahulu.
