"""
Aplikasi Generate Laporan Evaluasi Perkuliahan (CSV to PDF)
Format PDF sesuai dengan hasil.pdf
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate
import zipfile
import os

st.set_page_config(page_title="CSV to PDF - Laporan Evaluasi", page_icon="📄", layout="wide")

st.title("📄 Generator Laporan Evaluasi Perkuliahan")
st.markdown("Upload CSV hasil evaluasi dan generate PDF per Dosen + Mata Kuliah + Kelas")


# ============== QUESTION CATEGORIES ==============
# Mapping: category name -> list of question keywords (partial match)
CATEGORIES = {
    "Keandalan (Reliability)": [
        "Penjelasan sistem perkuliahan",
        "Ketepatan dosen hadir",
        "Kemampuan dosen dalam menyampaikan materi",
        "Penguasaaan dosen terhadap materi",
        "Pemberian umpan balik dari dosen",
    ],
    "Daya Tanggap (Responsiveness)": [
        "Kesigapan dosen dalam merespon kebutuhan mahasiswa (konsultasi) di dalam kelas",
        "Kesigapan dosen dalam merespon kebutuhan mahasiswa (konsultasi) di luar kelas",
        "Kegairahan dosen dalam mengajar",
        "Kemampuan dosen dalam menumbuhkan minat",
        "Kemampuan dosen dalam menumbuhkan suasana",
    ],
    "Kepastian (Assurance)": [
        "Kemampuan dosen dalam menggunakan metode pengajaran",
        "Relevansi materi kuliah",
        "Kemampuan dosen menggunakan media pembelajaran",
        "Ketepatan standar penilaian",
        "Kesesuaian materi perkuliahan dengan UTS",
    ],
    "Empathy": [
        "Perhatian dosen terhadap kemajuan",
        "Kesediaan dosen untuk membantu",
        "Pemberian masukan/pujian",
        "Kemampuan dosen berinteraksi sosial",
        "Kematangan emosional",
    ],
    "Tangible": [
        "Penggunaan bahasa saat pengajaran",
        "Intonasi dan kejelasan suara",
        "Penampilan dosen di kelas",
        "Sarana (alat bantu) pembelajaran",
        "Media ajar (buku, modul",
    ],
}

SCORE_MAP = {
    "Baik Sekali": 4,
    "baik sekali": 4,
    "Baik": 3,
    "baik": 3,
    "Cukup": 2,
    "cukup": 2,
    "Kurang": 1,
    "kurang": 1,
}


def get_score(val):
    return SCORE_MAP.get(str(val).strip(), 0)


def find_question_col(df_cols, keyword):
    """Find column that contains keyword"""
    for col in df_cols:
        if keyword.lower() in col.lower():
            return col
    return None


def get_category_cols(df_cols, category_keywords):
    """Return list of (col_name, short_label) for a category"""
    result = []
    for kw in category_keywords:
        col = find_question_col(df_cols, kw)
        if col:
            # Clean label: remove leading number like "1. " or "1) "
            label = col.strip().strip('"')
            import re
            label = re.sub(r'^\d+[\.\)]\s*', '', label)
            result.append((col, label))
    return result


def create_hline(width=17*cm, thickness=1):
    t = Table([['']], colWidths=[width])
    t.setStyle(TableStyle([
        ('BOTTOMPADDING', (0, 0), (0, 0), 0),
        ('TOPPADDING', (0, 0), (0, 0), 0),
        ('LINEBELOW', (0, 0), (0, 0), thickness, colors.black),
    ]))
    return t


def create_pdf_for_group(dosen, mata_kuliah, kelas, df_group, output_buffer,
                          semester="", judul="", prodi="", teknik="", platform=""):
    """Buat PDF sesuai format hasil.pdf"""

    BLUE = colors.HexColor("#4472C4")

    # ---- FUNGSI GAMBAR KOP DI SETIAP HALAMAN ----
    script_dir = os.path.dirname(os.path.abspath(__file__))
    logo_path  = os.path.join(script_dir, "logo.png")
    PAGE_W, PAGE_H = A4
    ML = 1.5*cm   # margin kiri
    MR = 1.5*cm   # margin kanan
    MT = 1.5*cm   # margin atas
    LINE_W = PAGE_W - ML - MR   # lebar garis = lebar konten

    def draw_kop(c, doc):
        c.saveState()

        # ── Logo (jika ada) ────────────────────────────────────────────────
        logo_w = logo_h = 2.2*cm
        text_x = ML  # default tanpa logo

        if os.path.exists(logo_path):
            logo_x = ML
            logo_y = PAGE_H - MT - logo_h
            c.drawImage(logo_path, logo_x, logo_y,
                        width=logo_w, height=logo_h,
                        preserveAspectRatio=True, mask='auto')
            text_x = ML + logo_w + 0.3*cm

        # ── Teks kop ────────────────────────────────────────────────────────
        c.setFillColor(BLUE)

        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(PAGE_W / 2, PAGE_H - MT - 0.7*cm,
                            "SEKOLAH TINGGI BAHASA ASING")

        c.setFont("Helvetica-Bold", 18)
        c.drawCentredString(PAGE_W / 2, PAGE_H - MT - 1.35*cm,
                            "Y  A  P  A  R  I")

        c.setFont("Helvetica-BoldOblique", 8)
        c.drawCentredString(PAGE_W / 2, PAGE_H - MT - 1.85*cm,
                            "Program Studi : Bhs. INGGRIS – Bhs. PRANCIS – Bhs. JERMAN – Bhs. JEPANG")

        c.setFont("Helvetica", 7.5)
        c.drawCentredString(PAGE_W / 2, PAGE_H - MT - 2.25*cm,
                            "Kampus : Jl. Cihampelas 194 Bandung 40131  |  Telp. (022) 2035426  |  WA 087822127474")
        c.drawCentredString(PAGE_W / 2, PAGE_H - MT - 2.6*cm,
                            "Website : www.stbayapariaba.ac.id  –  E-mail : info@stba.ac.id")

        # ── Dua garis biru mepet (tebal atas, tipis bawah) ─────────────────
        c.setStrokeColor(BLUE)
        y_top = PAGE_H - MT - 3.0*cm
        y_bot = y_top - 3          # jarak 3pt antar garis

        c.setLineWidth(2.5)
        c.line(ML, y_top, PAGE_W - MR, y_top)

        c.setLineWidth(1.0)
        c.line(ML, y_bot, PAGE_W - MR, y_bot)

        c.restoreState()

    # ── Hitung tinggi kop agar konten mulai di bawah garis ─────────────────
    KOP_HEIGHT = MT + 3.1*cm   # ruang yang dipakai kop dari atas halaman

    pdf = BaseDocTemplate(
        output_buffer,
        pagesize=A4,
        rightMargin=MR,
        leftMargin=ML,
        topMargin=KOP_HEIGHT,     # konten mulai di bawah kop
        bottomMargin=1.5*cm
    )

    frame = Frame(
        ML, 1.5*cm,
        LINE_W,
        PAGE_H - KOP_HEIGHT - 1.5*cm,
        id='normal'
    )
    pdf.addPageTemplates([
        PageTemplate(id='kop', frames=[frame], onPage=draw_kop)
    ])

    elements = []
    styles = getSampleStyleSheet()

    # ---- STYLES ----
    def ps(name, **kwargs):
        return ParagraphStyle(name, parent=styles['Normal'], **kwargs)

    title_style = ps('Title', fontSize=11, fontName='Helvetica-Bold', alignment=1, spaceAfter=0.2*cm)
    label_style = ps('Label', fontSize=9, fontName='Helvetica-Bold')
    value_style = ps('Value', fontSize=9, fontName='Helvetica')
    cat_style   = ps('Cat', fontSize=9, fontName='Helvetica-Bold')
    item_style  = ps('Item', fontSize=8, fontName='Helvetica', leading=10)
    avg_bold    = ps('AvgBold', fontSize=8, fontName='Helvetica-Bold')

    # ---- TITLE ----
    doc_title = judul if judul else "KUESIONER PROSES PEMBELAJARAN"
    elements.append(Paragraph(doc_title, title_style))
    elements.append(Spacer(1, 0.15*cm))

    # ---- INFO TABLE (2-column layout) ----
    def info_row(lbl, val):
        return [Paragraph(lbl, label_style), Paragraph(f": {val}", value_style)]

    info_left = [
        info_row("NAMA DOSEN", dosen),
        info_row("PRODI", prodi if prodi else "-"),
        info_row("SEMESTER", semester if semester else "-"),
        info_row("KELAS", kelas),
    ]
    info_right = [
        info_row("MATA KULIAH", mata_kuliah),
        info_row("TEKNIK PEMBELAJARAN", teknik if teknik else "-"),
        # info_row("", ""),
        info_row("PLATFORM ONLINE", platform if platform else "-"),
    ]

    # Build left and right as sub-tables
    # Column widths disesuaikan agar label sejajar dengan tabel utama
    left_t = Table(info_left, colWidths=[3*cm, 5*cm])
    left_t.setStyle(TableStyle([
        ('LEFTPADDING', (0,0),(-1,-1), 0),
        ('RIGHTPADDING', (0,0),(-1,-1), 2),
        ('TOPPADDING', (0,0),(-1,-1), 1),
        ('BOTTOMPADDING', (0,0),(-1,-1), 1),
        ('VALIGN', (0,0),(-1,-1), 'TOP'),
    ]))
    right_t = Table(info_right, colWidths=[4*cm, 4.5*cm])
    right_t.setStyle(TableStyle([
        ('LEFTPADDING', (0,0),(-1,-1), 0),
        ('RIGHTPADDING', (0,0),(-1,-1), 2),
        ('TOPPADDING', (0,0),(-1,-1), 1),
        ('BOTTOMPADDING', (0,0),(-1,-1), 1),
        ('VALIGN', (0,0),(-1,-1), 'TOP'),
    ]))

    outer_info = Table([[left_t, right_t]], colWidths=[8.5*cm, 8.5*cm])
    outer_info.setStyle(TableStyle([
        ('VALIGN', (0,0),(-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0),(-1,-1), 0),
        ('RIGHTPADDING', (0,0),(-1,-1), 0),
        ('TOPPADDING', (0,0),(-1,-1), 0),
        ('BOTTOMPADDING', (0,0),(-1,-1), 0),
    ]))
    elements.append(outer_info)
    elements.append(Spacer(1, 0.3*cm))

    # ---- PERNYATAAN KUESIONER HEADER ----
    elements.append(Paragraph("PERNYATAAN KUESIONER", label_style))
    elements.append(Spacer(1, 0.2*cm))

    # ---- BUILD MAIN TABLE ----
    # Columns: Pernyataan | 1 | 2 | 3 | 4 | Jumlah | Rata-rata
    # Col widths total = 17cm
    COL_W = [9.5*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.7*cm, 1.9*cm]

    # shared header row (shown once as a "floating" merged header)
    def make_header_row():
        return [
            Paragraph("", cat_style),
            Paragraph("Nilai Responden", ps('VH', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
            "", "", "",
            Paragraph("Jumlah", ps('VH', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
            Paragraph("Rata-rata", ps('VH', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
        ]

    def make_sub_header():
        return [
            "",
            Paragraph("1", ps('SH', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
            Paragraph("2", ps('SH', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
            Paragraph("3", ps('SH', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
            Paragraph("4", ps('SH', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
            "",
            "",
        ]

    all_category_avg = []

    for cat_name, cat_keywords in CATEGORIES.items():
        cat_cols = get_category_cols(df_group.columns.tolist(), cat_keywords)
        if not cat_cols:
            continue

        n = len(df_group)
        table_data = []

        # Category header rows
        # Row 1: category name + "Nilai Responden" merged + Jumlah + Rata-rata
        table_data.append([
            Paragraph(cat_name, cat_style),
            Paragraph("Nilai Responden", ps('VH2', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
            "", "", "",
            Paragraph("Jumlah", ps('JH', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
            Paragraph("Rata-rata", ps('RH', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
        ])

        # Row 2: sub-header 1 2 3 4
        table_data.append([
            "",
            Paragraph("1", ps('S1', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
            Paragraph("2", ps('S2', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
            Paragraph("3", ps('S3', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
            Paragraph("4", ps('S4', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
            "",
            "",
        ])

        cat_avgs = []
        item_rows_start = 2  # after 2 header rows

        for col, label in cat_cols:
            ratings = df_group[col].fillna("Cukup")
            count_1 = sum(1 for r in ratings if get_score(r) == 1)
            count_2 = sum(1 for r in ratings if get_score(r) == 2)
            count_3 = sum(1 for r in ratings if get_score(r) == 3)
            count_4 = sum(1 for r in ratings if get_score(r) == 4)
            total = n
            raw_scores = [get_score(r) for r in ratings]
            avg = sum(raw_scores) / len(raw_scores) if raw_scores else 0
            cat_avgs.append(avg)

            table_data.append([
                Paragraph(f" - {label}", item_style),
                Paragraph(str(count_1), ps('c1', fontSize=8, alignment=1)),
                Paragraph(str(count_2), ps('c2', fontSize=8, alignment=1)),
                Paragraph(str(count_3), ps('c3', fontSize=8, alignment=1)),
                Paragraph(str(count_4), ps('c4', fontSize=8, alignment=1)),
                Paragraph(str(total), ps('cj', fontSize=8, alignment=1)),
                Paragraph(f"{avg:.2f}", ps('cr', fontSize=8, alignment=1)),
            ])

        # Category average row
        cat_overall = sum(cat_avgs) / len(cat_avgs) if cat_avgs else 0
        all_category_avg.append(cat_overall)
        table_data.append([
            Paragraph(f"Rata-rata {cat_name}", avg_bold),
            "", "", "", "", "",
            Paragraph(f"{cat_overall:.2f}", ps('cavg', fontSize=8, fontName='Helvetica-Bold', alignment=1)),
        ])

        # Build table
        t = Table(table_data, colWidths=COL_W)

        n_rows = len(table_data)
        avg_row_idx = n_rows - 1

        ts = TableStyle([
            # Outer box
            ('BOX', (0,0), (-1,-1), 0.8, colors.black),
            # Grid for all
            ('GRID', (0,0), (-1,-1), 0.4, colors.black),
            # Category header row (row 0)
            ('SPAN', (1, 0), (4, 0)),  # merge "Nilai Responden" cols
            ('BACKGROUND', (0, 0), (-1, 1), colors.white),
            ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
            # Category name bold + border bottom thick
            ('LINEBELOW', (0,1), (-1,1), 0.8, colors.black),
            # Merge first col rows 0-1
            ('SPAN', (0, 0), (0, 1)),
            ('SPAN', (5, 0), (5, 1)),  # Jumlah span rows 0-1
            ('SPAN', (6, 0), (6, 1)),  # Rata-rata span rows 0-1
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (1,0), (6,-1), 'CENTER'),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('LEFTPADDING', (0,0), (-1,-1), 3),
            ('RIGHTPADDING', (0,0), (-1,-1), 3),
            # Average row styling
            ('SPAN', (1, avg_row_idx), (5, avg_row_idx)),
            ('BACKGROUND', (0, avg_row_idx), (-1, avg_row_idx), colors.white),
            ('FONTNAME', (0, avg_row_idx), (-1, avg_row_idx), 'Helvetica-Bold'),
            ('LINEABOVE', (0, avg_row_idx), (-1, avg_row_idx), 0.8, colors.black),
        ])
        t.setStyle(ts)
        elements.append(t)
        elements.append(Spacer(1, 0.15*cm))

    # ---- TOTAL ROW ----
    grand_avg = sum(all_category_avg) / len(all_category_avg) if all_category_avg else 0
    total_data = [
        [
            Paragraph("Total", ps('TotalLbl', fontSize=9, fontName='Helvetica-Bold')),
            "", "", "", "", "",
            Paragraph(f"{grand_avg:.2f}", ps('TotalVal', fontSize=9, fontName='Helvetica-Bold', alignment=1)),
        ]
    ]
    total_t = Table(total_data, colWidths=COL_W)
    total_t.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.8, colors.black),
        ('SPAN', (0,0), (5,0)),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ('ALIGN', (6,0), (6,0), 'CENTER'),
    ]))
    elements.append(total_t)

    # ---- PAGE 2: KOMENTAR ----
    elements.append(PageBreak())
    elements.append(Paragraph("Komentar", ps('KomTitle', fontSize=11, fontName='Helvetica-Bold', spaceAfter=0.3*cm)))

    # Cari kolom yang mengandung "komentar"
    komentar_col = None
    for col in df_group.columns:
        if "komentar" in col.lower():
            komentar_col = col
            break

    if komentar_col:
        for idx, k in enumerate(df_group[komentar_col], 1):
            if pd.notna(k) and str(k).strip():
                k_str = str(k).strip()
            else:
                k_str = "-"
            elements.append(Paragraph(f"{idx}. {k_str}", ps(f'K{idx}', fontSize=9, leading=12)))
    else:
        elements.append(Paragraph("Tidak ada kolom komentar ditemukan.", ps('NoKom', fontSize=9)))

    pdf.build(elements)
    output_buffer.seek(0)
    return output_buffer


# ============== STREAMLIT UI ==============

uploaded_file = st.file_uploader(
    "Upload file CSV evaluasi",
    type=["csv"],
    help="Format CSV sesuai template Google Form evaluasi"
)

if uploaded_file:
    # Try baca CSV dengan beberapa encoding dan separator
    encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252']
    separators = [',', ';', '\t', '|']  # comma, semicolon, tab, pipe
    df = None
    last_error = None

    # Reset file pointer untuk setiap percobaan
    uploaded_file.seek(0)
    content = uploaded_file.read()

    for enc in encodings:
        for sep in separators:
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding=enc, sep=sep)
                if len(df.columns) > 1:  # Valid jika lebih dari 1 kolom
                    st.success(f"✅ CSV berhasil dimuat! (Encoding: {enc}, Separator: '{sep}') Total {len(df)} baris data")
                    break
            except Exception:
                continue
        if df is not None and len(df.columns) > 1:
            break

    if df is None or len(df.columns) <= 1:
        st.error("❌ Gagal membaca CSV. Format tidak dikenali.")
        st.info("Pastikan file CSV menggunakan pemisah koma (,) atau semicolon (;)")

    with st.expander("👁️ Preview Data", expanded=False):
        st.dataframe(df, use_container_width=True, height=300)

    # Parse semester dari nama file
    filename = uploaded_file.name
    import re
    default_semester = ""
    default_judul = "KUESIONER PROSES PEMBELAJARAN"
    default_prodi = ""
    smt_num = "1"  # default
    smt_type = ""
    year_str = ""

    # Pattern 1: Smt_Ganjil atau Smt_Genap
    smt_type_match = re.search(r'Smt_(Ganjil|Genap)', filename, re.IGNORECASE)
    if smt_type_match:
        smt_type = smt_type_match.group(1).capitalize()
        if smt_type == "Ganjil":
            smt_num = "1"
        else:
            smt_num = "2"

    # Pattern 2: sem1_2526 atau sem1_25/26 (untuk tahun)
    sem_match = re.search(r'sem(\d+)[_\s]*(\d{2})(\d{2})', filename, re.IGNORECASE)
    if sem_match:
        smt_num = sem_match.group(1)
        year1 = "20" + sem_match.group(2)  # 25 -> 2025
        year2 = "20" + sem_match.group(3)  # 26 -> 2026
        year_str = f"{year1}/{year2}"

    # Pattern 3: Extract prodi (antara Smt_Ganjil/Genap dan sem)
    # Contoh: PBM_Smt_Ganjil_Perancis_sem1_2526
    prodi_match = re.search(r'Smt_(?:Ganjil|Genap)_([^_]+)(?:_sem|$)', filename, re.IGNORECASE)
    if prodi_match:
        default_prodi = prodi_match.group(1).capitalize()

    # Build semester & judul
    default_semester = smt_num

    if smt_type and year_str:
        default_judul = f"KUESIONER PROSES PEMBELAJARAN SEMESTER {smt_type.upper()} {year_str}"
    elif smt_type:
        default_judul = f"KUESIONER PROSES PEMBELAJARAN SEMESTER {smt_type.upper()}"
    elif year_str:
        default_judul = f"KUESIONER PROSES PEMBELAJARAN {year_str}"

    # Input field untuk edit
    st.markdown("---")
    col_a, col_b, col_c = st.columns([1, 1, 1])
    with col_a:
        input_prodi = st.text_input("Prodi", value=default_prodi, placeholder="cth: Perancis")
    with col_b:
        input_semester = st.text_input("Semester", value=default_semester, placeholder="cth: 1")
    with col_c:
        input_judul = st.text_input("Judul Dokumen", value=default_judul, placeholder="Judul laporan")
    st.markdown("---")

    # Try to detect extra info columns
    def get_col(df, keyword):
        for c in df.columns:
            if keyword.lower() in c.lower():
                return c
        return None

    dosen_col = get_col(df, "nama dosen")
    mk_col    = get_col(df, "mata kuliah")
    kelas_col = get_col(df, "kelas")
    prodi_col = get_col(df, "prodi")
    sem_col   = get_col(df, "semester")
    teknik_col= get_col(df, "teknik")
    platform_col = get_col(df, "platform")

    if not all([dosen_col, mk_col, kelas_col]):
        st.error("❌ Kolom 'Nama Dosen', 'Nama Mata Kuliah', atau 'Kelas' tidak ditemukan dalam CSV!")
    else:
        df["_dosen"] = df[dosen_col].str.strip()
        df["_mk"]    = df[mk_col].str.strip()
        df["_kelas"] = df[kelas_col].str.strip()
        groups = df.groupby(["_dosen", "_mk", "_kelas"])

        st.info(f"📊 Akan dibuat **{len(groups)} PDF** (per dosen + mata kuliah + kelas)")

        with st.expander("📁 Daftar PDF yang akan dibuat", expanded=True):
            for (dosen, mk, kelas), gdf in groups:
                st.markdown(f"**{dosen}** - *{mk}* ({kelas}) - {len(gdf)} responden")

        col1, col2 = st.columns([1, 1])

        with col1:
            generate_zip = st.button("📦 Generate & Download ZIP (Semua PDF)", type="primary", use_container_width=True)
        with col2:
            generate_individual = st.button("📄 Pilih Individual", use_container_width=True)

        def get_extra(gdf, col, trim_parens=False, unique_items=False):
            if col and col in gdf.columns:
                vals = gdf[col].dropna().unique()
                result = ", ".join(str(v) for v in vals) if len(vals) else ""
                if trim_parens:
                    import re
                    result = re.sub(r'\s*\(.*?\)', '', result).strip()
                if unique_items:
                    # Split by comma, get unique items, rejoin
                    items = [item.strip() for item in result.split(',') if item.strip()]
                    seen = set()
                    unique_items_list = []
                    for item in items:
                        if item not in seen:
                            seen.add(item)
                            unique_items_list.append(item)
                    result = ", ".join(unique_items_list)
                return result
            return ""

        if generate_zip:
            st.info("⏳ Sedang membuat PDF...")
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for (dosen, mk, kelas), gdf in groups:
                    buf = BytesIO()
                    create_pdf_for_group(
                        dosen, mk, kelas, gdf, buf,
                        semester=input_semester,
                        judul=input_judul,
                        prodi=input_prodi if input_prodi else get_extra(gdf, prodi_col),
                        teknik=get_extra(gdf, teknik_col, trim_parens=True, unique_items=True),
                        platform=get_extra(gdf, platform_col, unique_items=True),
                    )
                    safe = lambda s: "".join(c if c.isalnum() or c in (' ','-','_') else '_' for c in s)
                    fname = f"{safe(dosen)}_{safe(mk)}_{safe(kelas)}.pdf"
                    zf.writestr(fname, buf.getvalue())

            zip_buffer.seek(0)
            st.success(f"✅ Berhasil membuat {len(groups)} PDF!")
            st.download_button(
                label="⬇️ Download ZIP",
                data=zip_buffer.getvalue(),
                file_name=f"laporan_evaluasi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                use_container_width=True
            )

        if generate_individual:
            st.subheader("📄 Generate Individual PDF")
            group_list = list(groups)
            group_options = [
                f"{d} - {m} ({k}) - {len(g)} responden"
                for (d, m, k), g in group_list
            ]
            selected = st.selectbox("Pilih kombinasi:", group_options)
            if selected:
                idx = group_options.index(selected)
                (dosen, mk, kelas), gdf = group_list[idx]
                if st.button(f"Generate PDF: {selected}", type="primary"):
                    buf = BytesIO()
                    create_pdf_for_group(
                        dosen, mk, kelas, gdf, buf,
                        semester=input_semester,
                        judul=input_judul,
                        prodi=input_prodi if input_prodi else get_extra(gdf, prodi_col),
                        teknik=get_extra(gdf, teknik_col, trim_parens=True, unique_items=True),
                        platform=get_extra(gdf, platform_col, unique_items=True),
                    )
                    st.success("✅ PDF berhasil dibuat!")
                    safe = lambda s: "".join(c if c.isalnum() or c in (' ','-','_') else '_' for c in s)
                    fname = f"{safe(dosen)}_{safe(mk)}_{safe(kelas)}.pdf"
                    st.download_button(
                        label="⬇️ Download PDF",
                        data=buf.getvalue(),
                        file_name=fname,
                        mime="application/pdf",
                        use_container_width=True
                    )
else:
    st.info("👆 Upload file CSV untuk memulai")

st.markdown("---")
st.markdown("""
<style>.footer { text-align: center; color: #888; font-size: 0.9em; }</style>
<div class="footer">Generator Laporan Evaluasi Perkuliahan · Built with Streamlit + ReportLab</div>
""", unsafe_allow_html=True)
