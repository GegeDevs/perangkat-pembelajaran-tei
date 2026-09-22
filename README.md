# perangkat-pembelajaran-tei

Skill Hermes Agent untuk menyusun perangkat pembelajaran TEI Kurikulum Merdeka dalam DOCX/PDF, dengan LKPD dan kunci jawaban terpisah.

## Instalasi Hermes

Salin `skills/productivity/perangkat-pembelajaran-tei` ke direktori skills Hermes, atau instal dari repo sesuai workflow skill manager.

## Cakupan

- RPP
- Bahan Ajar bertingkat bab
- Modul Ajar
- LKPD tanpa kunci jawaban
- Kunci Jawaban terpisah
- Ilustrasi komponen dan rangkaian
- Validasi DOCX, PDF, media, dan ZIP

## Isi skill

- Generator: `skills/productivity/perangkat-pembelajaran-tei/scripts/generate_perangkat.py`
- Library layout: `scripts/vendor/lib_cv.py`
- Workflow: `references/workflow-scripts.md`
- Inventaris audit: `references/inventory.md`
- Referensi kurikulum DTEI Fase E dan TEI Fase F: `references/kurikulum/`

Generator adalah salinan proyek referensi; tinjau `OUT`, identitas, CP/TP, tanggal, dan materi sebelum digunakan ulang.
