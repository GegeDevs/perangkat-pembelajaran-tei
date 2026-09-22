#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Shared library for CV ATS generation.
Import this instead of duplicating functions across scripts.

Usage:
    from lib_cv import (
        set_doc_style, add_section_title, add_text, add_bullet,
        add_rich_paragraph, count_pdf_pages, convert_docx_to_pdf,
        save_version, slugify,
        process_photo, add_photo_header,
    )
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image, ImageOps
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor
from pypdf import PdfReader


# ── HELPERS ──

def slugify(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9]+", "_", name.strip().lower()).strip("_")


def detect_env() -> str:
    """Detect running environment for PDF converter selection."""
    if shutil.which("soffice"):
        return "native"
    if shutil.which("proot-distro"):
        return "proot"
    # Coba cek apakah soffice ada di PATH sistem (kadang tanpa lokasi)
    try:
        subprocess.run(["soffice", "--headless", "--version"],
                       capture_output=True, timeout=3)
        return "native"
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return "unknown"


def _get_best_font() -> str:
    """Return best available serif font, fallback jika Times New Roman tidak ada."""
    candidates = ["Times New Roman", "Tinos", "Liberation Serif", "DejaVu Serif", "serif"]
    for font in candidates:
        if shutil.which("fc-match"):
            try:
                result = subprocess.run(
                    ["fc-match", "-s", font],
                    capture_output=True, text=True, timeout=2
                )
                if font.lower() in result.stdout.lower():
                    return font
            except subprocess.TimeoutExpired:
                continue
        # Fallback: coba return langsung, python-docx akan pakai font name apa pun
    return "Times New Roman"  # fallback, python-docx tetap set nama ini


def set_doc_style(doc: Document, page_w_cm: float, page_h_cm: float,
                  margin_cm: float = 2.0, body_size: int = 12) -> None:
    style = doc.styles["Normal"]
    font_name = _get_best_font()
    style.font.name = font_name
    style.font.size = Pt(body_size)
    for section in doc.sections:
        section.top_margin = Cm(margin_cm)
        section.bottom_margin = Cm(margin_cm)
        section.left_margin = Cm(margin_cm)
        section.right_margin = Cm(margin_cm)
        section.page_width = Cm(page_w_cm)
        section.page_height = Cm(page_h_cm)


def add_section_title(doc: Document, title: str,
                      space_before: int = 6, space_after: int = 0,
                      font_size: int = 12):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(
        '<w:pBdr %s>'
        '  <w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        '</w:pBdr>' % nsdecls("w")
    ))
    run = p.add_run(title)
    run.font.name = "Times New Roman"
    run.font.size = Pt(font_size)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)


def add_text(doc: Document, text: str, *, bold=False, italic=False, size=12,
             align=None, space_before=0, space_after=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)


def add_bullet(doc: Document, text: str, *,
               space_before=0.5, space_after=0.5, font_size=12):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    # Add tab stop at text position (0.5cm from margin edge)
    pPr = p._p.get_or_add_pPr()
    tabs_elem = OxmlElement('w:tabs')
    tab_elem = OxmlElement('w:tab')
    tab_elem.set(qn('w:val'), 'left')
    # 0.5cm in twips: 0.5 / 2.54 * 1440
    tab_elem.set(qn('w:pos'), str(int(0.5 / 2.54 * 1440)))
    tabs_elem.append(tab_elem)
    pPr.append(tabs_elem)
    run = p.add_run("\u2022\t" + text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(0, 0, 0)


def add_rich_paragraph(doc: Document, parts: list, *,
                       align=None, space_before=0, space_after=0,
                       font_size=12):
    """Add paragraph with mixed (text, bold, italic) runs."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if align is not None:
        p.alignment = align
    for text, bold, italic in parts:
        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = RGBColor(0, 0, 0)


def add_work_header(doc: Document, title_text: str, period_text: str, *,
                    body_size: int = 12, space_before: int = 4) -> None:
    """Add work experience header with title left-aligned (bold) and period right-aligned (italic) on same line.

    Creates a single paragraph with the title bold on the left and the period italic
    on the right, separated by a right-aligned tab stop.
    """
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(0)

    # Calculate tab stop position (text width from left margin edge)
    section = doc.sections[0]
    text_width_emu = section.page_width - section.left_margin - section.right_margin
    # Convert EMU to twips: 1 EMU = 1/914400 inch, 1 twip = 1/1440 inch
    tab_pos_twips = text_width_emu * 1440 // 914400

    # Add right-aligned tab stop via oxml
    pPr = p._p.get_or_add_pPr()
    tabs_elem = pPr.find(qn('w:tabs'))
    if tabs_elem is None:
        tabs_elem = OxmlElement('w:tabs')
        pPr.append(tabs_elem)
    tab_elem = OxmlElement('w:tab')
    tab_elem.set(qn('w:val'), 'right')
    tab_elem.set(qn('w:pos'), str(tab_pos_twips))
    tabs_elem.append(tab_elem)

    # Title run (bold)
    run = p.add_run(title_text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(body_size)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)

    # Tab character
    p.add_run("\t")

    # Period run (italic)
    run_period = p.add_run(period_text)
    run_period.font.name = "Times New Roman"
    run_period.font.size = Pt(body_size)
    run_period.font.italic = True
    run_period.font.color.rgb = RGBColor(0, 0, 0)


def count_pdf_pages(pdf_path: Path) -> int:
    return len(PdfReader(str(pdf_path)).pages)


def convert_docx_to_pdf(docx_path: Path, outdir: Path) -> Path:
    """Convert DOCX to PDF using best available engine."""
    env = detect_env()

    if env == "native":
        cmd = ["soffice", "--headless", "--convert-to", "pdf",
               "--outdir", str(outdir), str(docx_path)]
    elif env == "proot":
        cmd = ["proot-distro", "login", "alpine", "--",
               "soffice", "--headless", "--convert-to", "pdf",
               "--outdir", str(outdir), str(docx_path)]
    else:
        # Fallback: coba langsung
        cmd = ["soffice", "--headless", "--convert-to", "pdf",
               "--outdir", str(outdir), str(docx_path)]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or result.stdout.strip())
    except FileNotFoundError:
        raise RuntimeError(
            "LibreOffice tidak ditemukan.\n"
            "  VPS: sudo apt install libreoffice-writer\n"
            "  Termux: proot-distro install alpine && "
            "apk add libreoffice-writer\n"
            "  Atau set environment: export CV_PDF_ENGINE=soffice"
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("Konversi PDF timeout (60s)")

    return outdir / (docx_path.stem + ".pdf")


def save_version(doc: Document, base_name: str, suffix: str, workspace: Path,
                 output_dir: Path | None = None) -> tuple[Path, Path]:
    out = output_dir or Path(os.environ.get("CV_OUTPUT_DIR", str(workspace)))
    docx_path = out / f"{base_name}_{suffix}.docx"
    pdf_path = out / f"{base_name}_{suffix}.pdf"
    out.mkdir(parents=True, exist_ok=True)
    for p in (docx_path, pdf_path):
        if p.exists():
            p.unlink()
    doc.save(str(docx_path))
    generated_pdf = convert_docx_to_pdf(docx_path, out)
    if generated_pdf != pdf_path:
        if pdf_path.exists():
            pdf_path.unlink()
        generated_pdf.rename(pdf_path)
    return docx_path, pdf_path


# ── FOTO HELPERS (untuk Varian B) ──

def process_photo(src: Path) -> Path:
    """Crop to 3:4 portrait, resize to 3x4 cm, add black thin border.
    Preserves RGBA (transparency) if present. Saves as PNG if has alpha, else JPEG.
    Returns temp path (auto-deleted on gc).
    """
    img = Image.open(src)
    has_alpha = img.mode in ("RGBA", "LA", "P")
    if img.mode == "P":
        img = img.convert("RGBA" if img.info.get("transparency") else "RGB")
        has_alpha = img.mode == "RGBA"

    target_ratio = 3 / 4
    w, h = img.size
    current_ratio = w / h

    if current_ratio > target_ratio:
        new_w = int(h * target_ratio)
        left = (w - new_w) // 2
        img = img.crop((left, 0, left + new_w, h))
    elif current_ratio < target_ratio:
        new_h = int(w / target_ratio)
        top = (h - new_h) // 2
        img = img.crop((0, top, w, top + new_h))

    dpi = 150
    target_w = int(3 / 2.54 * dpi)
    target_h = int(4 / 2.54 * dpi)
    img = img.resize((target_w, target_h), Image.LANCZOS)

    # Border: use black, on alpha channel use (0,0,0,255)
    border_color = (0, 0, 0, 255) if has_alpha else "black"
    img = ImageOps.expand(img, border=2, fill=border_color)

    if has_alpha:
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        img.save(tmp.name, "PNG")
    else:
        tmp = tempfile.NamedTemporaryFile(suffix=".jpg", delete=False)
        img.save(tmp.name, "JPEG", quality=95)
    return Path(tmp.name)


def apply_bg_color(photo_path: Path, color: str = "red") -> Path:
    """Replace transparent background with solid color. Returns new temp PNG path."""
    img = Image.open(photo_path)
    if img.mode != "RGBA":
        return photo_path  # nothing to replace

    # Parse color name to RGBA
    color_map = {
        "red": (255, 0, 0, 255),
        "blue": (0, 0, 255, 255),
        "green": (0, 128, 0, 255),
        "white": (255, 255, 255, 255),
        "black": (0, 0, 0, 255),
        "yellow": (255, 255, 0, 255),
        "gray": (128, 128, 128, 255),
        "grey": (128, 128, 128, 255),
        "navy": (0, 0, 128, 255),
        "darkred": (139, 0, 0, 255),
        "maroon": (128, 0, 0, 255),
    }
    bg_rgba = color_map.get(color.lower())
    if not bg_rgba:
        # Try hex color
        try:
            from PIL.ImageColor import getrgb
            rgb = getrgb(color)
            bg_rgba = tuple(rgb) + (255,) if len(rgb) == 3 else rgb
        except Exception:
            print(f"⚠  Warna '{color}' tidak dikenal, pakai merah.")
            bg_rgba = (255, 0, 0, 255)

    bg = Image.new("RGBA", img.size, bg_rgba)
    comp = Image.alpha_composite(bg, img)

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    comp.save(tmp.name, "PNG")
    return Path(tmp.name)


def remove_photo_background(src: Path) -> Path:
    """Remove background from photo using rembg (u2net). Returns PNG with transparency.
    Falls back to original if rembg is not installed.
    """
    try:
        from rembg import remove as remove_bg
    except ImportError:
        print("⚠  rembg tidak terinstall. Lewati remove background.")
        return src

    print("  Menghapus background foto...")
    with open(src, "rb") as f:
        input_data = f.read()
    output_data = remove_bg(input_data)

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmp.write(output_data)
    tmp.flush()
    return Path(tmp.name)


def _remove_cell_borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def _remove_cell_margins(cell):
    """Zero out cell margins (tcMar) so content can use full cell width."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    # Hapus tcMar lama jika ada
    for old in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(old)
    tcMar = OxmlElement("w:tcMar")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)


def _set_cell_vertical_alignment(cell, align="top"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    # Hapus vAlign lama jika ada
    for old in tcPr.findall(qn("w:vAlign")):
        tcPr.remove(old)
    val = OxmlElement("w:vAlign")
    val.set(qn("w:val"), align)
    tcPr.append(val)


def add_photo_header(doc: Document, page_w_cm: float, foto_path: Path,
                     nama: str, alamat: str, kontak: str,
                     name_size: int = 14, info_size: int = 12,
                     margin_cm: float = 2.0) -> None:
    """Tambahkan header dengan foto kiri, teks di kanan (Varian B)."""
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    available_w = page_w_cm - 2 * margin_cm

    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")

    # Hapus tblW lama (type="auto") dan ganti dengan fixed width
    for old_tblW in tblPr.findall(qn("w:tblW")):
        tblPr.remove(old_tblW)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(int(available_w / 2.54 * 1440)))  # dxa
    tblW.set(qn("w:type"), "dxa")
    tblPr.insert(0, tblW)

    # Update tblGrid: set kolom sesuai lebar cell
    left_dxa = int(3.8 / 2.54 * 1440)      # 2154 dxa
    right_dxa = int((available_w - 3.8) / 2.54 * 1440)  # 8617 dxa
    tblGrid = tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        grid_cols = tblGrid.findall(qn("w:gridCol"))
        if len(grid_cols) >= 2:
            grid_cols[0].set(qn("w:w"), str(left_dxa))
            grid_cols[1].set(qn("w:w"), str(right_dxa))
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        tblBorders.append(el)
    tblPr.append(tblBorders)

    left_cell = table.cell(0, 0)
    right_cell = table.cell(0, 1)
    left_cell.width = Cm(3.8)
    right_cell.width = Cm(available_w - 3.8)

    _remove_cell_borders(left_cell)
    _remove_cell_borders(right_cell)
    _set_cell_vertical_alignment(left_cell, "center")
    _set_cell_vertical_alignment(right_cell, "center")

    # Photo
    p_photo = left_cell.paragraphs[0]
    p_photo.paragraph_format.space_before = Pt(0)
    p_photo.paragraph_format.space_after = Pt(0)
    run_photo = p_photo.add_run()
    run_photo.add_picture(str(foto_path), width=Cm(3.0), height=Cm(4.0))

    # Nama
    p_name = right_cell.paragraphs[0]
    p_name.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(0)
    run_name = p_name.add_run(nama)
    run_name.font.name = "Times New Roman"
    run_name.font.size = Pt(name_size)
    run_name.font.bold = True
    run_name.font.color.rgb = RGBColor(0, 0, 0)

    # Info baris: tabel bersarang (label + value) biar titik dua lurus vertikal
    parts = kontak.split("|") if kontak else []
    telp = parts[0].strip() if len(parts) > 0 else ""
    email = parts[1].strip() if len(parts) > 1 else ""

    # Build data rows: (label, value)
    info_data = []
    if alamat:
        info_data.append(("Alamat", alamat))
    if email:
        info_data.append(("Email", email))
    if telp:
        info_data.append(("Telp", telp))

    if info_data:
        nrows = len(info_data)
        info_table = right_cell.add_table(rows=nrows, cols=3)
        info_table.autofit = False
        cell_w_cm = right_cell.width / 914400 * 2.54
        label_w = Cm(2.5)
        colon_w = Cm(1.2)
        val_w = Cm(cell_w_cm - 2.5 - 1.2)

        itbl = info_table._tbl
        itblPr = itbl.tblPr if itbl.tblPr is not None else OxmlElement("w:tblPr")
        for old in itblPr.findall(qn("w:tblW")):
            itblPr.remove(old)
        itblW = OxmlElement("w:tblW")
        itblW.set(qn("w:w"), str(int(cell_w_cm / 2.54 * 1440)))
        itblW.set(qn("w:type"), "dxa")
        itblPr.insert(0, itblW)
        itblLayout = OxmlElement("w:tblLayout")
        itblLayout.set(qn("w:type"), "fixed")
        itblPr.append(itblLayout)
        itblBorders = OxmlElement("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{side}")
            el.set(qn("w:val"), "none")
            el.set(qn("w:sz"), "0")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
            itblBorders.append(el)
        itblPr.append(itblBorders)
        itblGrid = OxmlElement("w:tblGrid")
        for w_cm in (2.5, 1.2, cell_w_cm - 2.5 - 1.2):
            gc = OxmlElement("w:gridCol")
            gc.set(qn("w:w"), str(int(w_cm / 2.54 * 1440)))
            itblGrid.append(gc)
        itbl.insert(1, itblGrid)

        info_table.columns[0].width = label_w
        info_table.columns[1].width = colon_w
        info_table.columns[2].width = val_w

        for idx, (label, val) in enumerate(info_data):
            row = info_table.rows[idx]
            # Label cell — right-aligned
            cell_label = row.cells[0]
            _remove_cell_borders(cell_label)
            _remove_cell_margins(cell_label)
            _set_cell_vertical_alignment(cell_label, "center")
            cell_label.width = label_w
            p_label = cell_label.paragraphs[0]
            p_label.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_label.paragraph_format.space_before = Pt(1)
            p_label.paragraph_format.space_after = Pt(1)
            r_label = p_label.add_run(label)
            r_label.font.name = "Times New Roman"
            r_label.font.size = Pt(info_size)
            r_label.font.color.rgb = RGBColor(0, 0, 0)

            # Colon cell — center-aligned
            cell_colon = row.cells[1]
            _remove_cell_borders(cell_colon)
            _remove_cell_margins(cell_colon)
            _set_cell_vertical_alignment(cell_colon, "center")
            cell_colon.width = colon_w
            p_colon = cell_colon.paragraphs[0]
            p_colon.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_colon.paragraph_format.space_before = Pt(1)
            p_colon.paragraph_format.space_after = Pt(1)
            r_colon = p_colon.add_run(":")
            r_colon.font.name = "Times New Roman"
            r_colon.font.size = Pt(info_size)
            r_colon.font.color.rgb = RGBColor(0, 0, 0)

            # Value cell — left-aligned
            cell_val = row.cells[2]
            _remove_cell_borders(cell_val)
            _remove_cell_margins(cell_val)
            _set_cell_vertical_alignment(cell_val, "center")
            cell_val.width = val_w
            p_val = cell_val.paragraphs[0]
            p_val.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_val.paragraph_format.space_before = Pt(1)
            p_val.paragraph_format.space_after = Pt(1)
            r_val = p_val.add_run(val)
            r_val.font.name = "Times New Roman"
            r_val.font.size = Pt(info_size)
            r_val.font.color.rgb = RGBColor(0, 0, 0)
