# Referensi Skrip dan Workflow

## Generator utama

`../scripts/generate_perangkat.py` adalah generator Python yang membuat:

- ilustrasi resistor dan gelang warna;
- ilustrasi contoh pembacaan resistor;
- ilustrasi kapasitor polar/non-polar beserta simbol;
- ilustrasi rangkaian seri/paralel;
- RPP, Bahan Ajar, Modul Ajar, LKPD;
- Kunci Jawaban terpisah.

Skrip memakai `python-docx` dan library layout dari skill `cv-ats-format`. Sesuaikan konstanta identitas, CP/TP, tanggal, materi, dan direktori output sebelum digunakan ulang. Jangan menyalin identitas proyek lama tanpa memeriksanya.

## Environment

Gunakan environment yang memiliki `python-docx`, LibreOffice, dan validator DOCX. Contoh pada lingkungan proyek:

```sh
cd ~/hermes/<nama-proyek>
~/hermes/.venv-pdf/bin/python generate_perangkat.py
```

Jika generator menyediakan fungsi utama tertentu, jalankan entry point yang ada di bagian akhir skrip. Jangan menganggap nama fungsi tanpa membaca skrip versi yang dipakai.

## Konversi DOCX ke PDF

```sh
mkdir -p pdf
soffice --headless --convert-to pdf --outdir pdf *.docx
```

Pastikan jumlah PDF sesuai jumlah DOCX dan buka/ekstrak teks PDF untuk pemeriksaan isi.

## Validasi DOCX

```sh
for f in *.docx; do
  ~/hermes/.venv-pdf/bin/python \
    ~/.hermes/skills/productivity/docx/scripts/docx_validate.py "$f" || exit 1
done
```

Periksa juga bahwa gambar yang diwajibkan benar-benar tertanam, bukan hanya ada sebagai file PNG:

```sh
python3 - <<'PY'
from zipfile import ZipFile
from pathlib import Path
p=Path('BahanAjar_ResistorKapasitor_X-TEI_GegeDesembri.docx')
with ZipFile(p) as z:
    media=[x for x in z.namelist() if x.startswith('word/media/')]
    print('embedded media:', media)
    assert len(media) >= 3
PY
```

## Pembuatan dan uji ZIP

```sh
python3 - <<'PY'
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
root=Path('.')
out=Path('PerangkatPembelajaran_TEI.zip')
files=[p for p in root.iterdir() if p.suffix.lower() in {'.docx','.pdf','.png'}]
with ZipFile(out,'w',ZIP_DEFLATED) as z:
    for p in files: z.write(p,p.name)
with ZipFile(out) as z:
    assert z.testzip() is None
    print('files:', len(z.namelist()))
PY
```

## Pengiriman Telegram

Salin file final ke cache sebelum mengirim:

```sh
cp PerangkatPembelajaran_TEI.zip ~/.hermes/document_cache/
```

Gunakan path cache pada directive `MEDIA:`. Jangan mengirim kredensial, token, password, API key, atau file konfigurasi rahasia. Bila contoh membutuhkan rahasia, tulis `[REDACTED]`.

## Catatan portabilitas

- Gunakan `Path` dan `~/hermes` untuk workspace.
- Arial boleh diganti Liberation Sans jika Arial tidak tersedia.
- Hindari hardcode path `/root/hermes/...` saat membagikan skrip; ubah `OUT` menjadi argumen atau konstanta proyek.
- Tinjau semua identitas sekolah/guru sebelum regenerasi.
