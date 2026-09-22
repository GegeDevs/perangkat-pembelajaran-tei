# Inventaris Referensi yang Dipakai

Audit ini mencatat artefak dan dependensi yang digunakan saat menyusun perangkat pembelajaran TEI.

## Sumber kurikulum

- `references/kurikulum/REFERENSI-DTEI-Fase-E.md` — dasar DTEI kelas X; Elemen 9 (komponen aktif/pasif) dan Elemen 11 (konsep dasar kelistrikan) relevan untuk materi resistor/kapasitor.
- `references/kurikulum/REFERENSI-TEI-Fase-F-PSE-AKD.md` — referensi lanjutan TEI Fase F, terutama PSE dan AKD. Tidak perlu digunakan untuk kelas X kecuali permintaan aktif memerlukannya.

## Skrip dan library

- `scripts/generate_perangkat.py` — generator proyek yang telah dipakai dan perlu diadaptasi sebelum reuse.
- `scripts/vendor/lib_cv.py` — fungsi layout DOCX dan konversi PDF yang diimpor generator dari `cv-ats-format`.
- `references/workflow-scripts.md` — prosedur eksekusi, konversi, validasi, embedding media, ZIP, dan pengiriman.

## Dependensi eksternal

- `python-docx` — pembuatan DOCX.
- `soffice`/LibreOffice — konversi DOCX ke PDF.
- `docx_validate.py` dari skill `docx` — validasi struktur DOCX.
- `ZipFile` dari standard library — uji integritas ZIP.
- `vision_analyze` — pemeriksaan visual opsional.

## Media yang dibuat generator

- resistor bergelang warna;
- contoh pembacaan gelang resistor;
- kapasitor polar/non-polar dan simbol;
- rangkaian resistor/kapasitor seri-paralel.

## Batasan yang harus diperbaiki saat reuse

- Ubah `OUT` dari path absolut proyek lama menjadi direktori proyek aktif.
- Ubah interpreter/shebang dan impor library sesuai mesin baru.
- Periksa font; Arial dapat diganti Liberation Sans atau font kompatibel.
- Jangan membawa DOCX/PDF/PNG hasil supervisi lama sebagai template tanpa meninjau identitas, tanggal, CP/TP, dan materi.
- Jangan menyalin PDF ATP asli ke repo publik bila lisensi/izin distribusinya belum dipastikan; referensi ringkas yang sudah ada tidak memuat kredensial.
