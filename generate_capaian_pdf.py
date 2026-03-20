#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_capaian_pdf.py — PDF Laporan Capaian IKI Tim LIM
BMKG · Stasiun Meteorologi Kelas III Sangia Nibandera Kolaka (WAWP)

Penggunaan:
    python generate_capaian_pdf.py [file.json] [--q 1] [--year 2026]

    Jika file.json tidak diberikan, script mencari file LIM_data_*.json terbaru.

Dependensi:
    pip install reportlab Pillow
"""

import sys, os, glob, json, base64, io, argparse
from datetime import datetime, timezone, timedelta

# ── Dependency check ─────────────────────────────────────────────────────────
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm, mm
    from reportlab.lib import colors
    from reportlab.platypus import (
        BaseDocTemplate, Frame, PageTemplate, NextPageTemplate,
        Image as RLImage, Paragraph, Spacer, Table, TableStyle,
        KeepTogether, PageBreak
    )
    from reportlab.platypus.flowables import HRFlowable
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.pdfbase import pdfmetrics
except ImportError:
    print("ERROR: ReportLab belum terinstal.")
    print("Jalankan: pip install reportlab Pillow")
    sys.exit(1)

try:
    from PIL import Image as PILImage
except ImportError:
    print("ERROR: Pillow belum terinstal.")
    print("Jalankan: pip install Pillow")
    sys.exit(1)

# ── KOP IMAGE (base64 embedded) ───────────────────────────────────────────────
# KOP_B64 akan diisi oleh script setup atau diambil dari JSON state
KOP_B64 = None  # fallback: coba ambil dari JSON state

# ── CONSTANTS ─────────────────────────────────────────────────────────────────
PW, PH = A4          # 595.27 x 841.89 pts
ML  = 2.5 * cm       # left margin
MR  = 2.0 * cm       # right margin
MT  = 1.5 * cm       # top margin (content, after KOP on page 1)
MB  = 2.5 * cm       # bottom margin
CW  = PW - ML - MR   # content width ≈ 467.7 pts

KOP_H_RATIO = 210 / 1400   # KOP image height/width ratio (1400x210px)
KOP_PRINT_H = PW * KOP_H_RATIO  # ≈ 89.3 pts ≈ 3.15 cm

# Column proportions (exact from DOCX source)
COL_PCTS = [8.3, 51.6, 10.4, 7.4, 7.4, 7.4, 7.4]
COLS = [CW * p / 100 for p in COL_PCTS]

# Colors
C_HEADER_BG   = colors.HexColor('#9CC2E5')
C_HEADER_TXT  = colors.black
C_ODD_ROW     = colors.white
C_EVEN_ROW    = colors.HexColor('#F2F2F2')
C_BORDER      = colors.black

# Month names
MONTHS_ID = ['','Januari','Februari','Maret','April','Mei','Juni',
              'Juli','Agustus','September','Oktober','November','Desember']

Q_LABELS = {1:'Triwulan I', 2:'Triwulan II', 3:'Triwulan III', 4:'Triwulan IV'}
M_LABELS = {
    1: ['Januari',  'Februari', 'Maret'],
    2: ['April',    'Mei',      'Juni'],
    3: ['Juli',     'Agustus',  'September'],
    4: ['Oktober',  'November', 'Desember'],
}

# LIM IKI data
LIM_IKI = [
    {'no':'1.1.1',  'indi':'Persentase akurasi informasi METAR',                                                             'target':'100%'},
    {'no':'1.1.2',  'indi':'Persentase akurasi informasi TAF',                                                               'target':'97,5%'},
    {'no':'1.1.3',  'indi':'Persentase penyediaan informasi Flight Documentation',                                           'target':'100%'},
    {'no':'1.1.4',  'indi':'Persentase akurasi informasi AD WRNG',                                                           'target':'97,5%'},
    {'no':'1.1.5',  'indi':'Indeks penyampaian hasil analisis fenomena meteorologi signifikan dan atau cuaca ekstrem secara lengkap', 'target':'4 Indeks'},
    {'no':'1.1.6',  'indi':'Indeks penyelesaian laporan operasional melalui Portal Laporan Stasiun Meteorologi Penerbangan', 'target':'4 Indeks'},
    {'no':'1.1.7',  'indi':'Hasil evaluasi akurasi informasi METAR',                                                         'target':'3 Lap.'},
    {'no':'1.1.8',  'indi':'Hasil evaluasi akurasi informasi TAF',                                                           'target':'3 Lap.'},
    {'no':'1.1.9',  'indi':'Hasil evaluasi penyediaan Flight Documentation',                                                 'target':'3 Lap.'},
    {'no':'1.1.10', 'indi':'Hasil monitoring data statistik penerbangan',                                                     'target':'3 Lap.'},
    {'no':'1.1.11', 'indi':'Hasil evaluasi akurasi informasi AD WRNG',                                                        'target':'3 Lap.'},
    {'no':'1.1.12', 'indi':'Hasil penyusunan dokumen ACS',                                                                    'target':'3 Lap.'},
    {'no':'1.1.13', 'indi':'Jumlah produk informasi Kolaka Monthly Weather Report yang disebarluaskan ke stakeholder',        'target':'3 Dok.'},
    {'no':'1.1.14', 'indi':'Jumlah produk informasi Monthly Aviation Weather Summary yang disebarluaskan ke stakeholder',    'target':'3 Dok.'},
    {'no':'1.1.15', 'indi':'Jumlah penyampaian produk informasi Catatan Iklim Kolaka ke stakeholder',                        'target':'1 Dok.'},
    {'no':'1.1.16', 'indi':'Hasil rekapitulasi data parameter cuaca ekstrem',                                                 'target':'3 Lap.'},
    {'no':'1.2.1',  'indi':'Jumlah pelaksanaan kegiatan sosialisasi informasi meteorologi penerbangan',                       'target':'4 Lap.'},
    {'no':'1.2.2',  'indi':'Persentase hasil survey pemahaman masyarakat pengguna informasi meteorologi penerbangan',         'target':'87%'},
    {'no':'1.3.1',  'indi':'Jumlah pelaksanaan kegiatan Survey Kepuasan Masyarakat',                                          'target':'2 Lap.'},
    {'no':'1.3.2',  'indi':'Nilai IKM layanan informasi meteorologi penerbangan',                                             'target':'3,8 SL'},
]

# ── UTILITIES ─────────────────────────────────────────────────────────────────
def wita_now():
    return datetime.now(timezone(timedelta(hours=8)))

def fmt_date_id(dt):
    return f"{dt.day} {MONTHS_ID[dt.month]} {dt.year}"

def find_json():
    for pat in ['LIM_data_*.json', 'LIM_WAWP_*.json', 'lim_*.json']:
        files = glob.glob(pat)
        if files:
            return max(files, key=os.path.getmtime)
    return None

def load_state(path):
    with open(path, 'r', encoding='utf-8') as f:
        raw = json.load(f)
    if 'state' in raw:
        return raw['state']
    if '_meta' in raw:
        raw.pop('_meta', None)
    return raw

def get_cap(capaian, year, q, iki_no, field):
    key = f"{year}-Q{q}"
    return (capaian.get(key, {}).get(iki_no, {}) or {}).get(field, '') or ''

def shorten_muhammad(name):
    if name.startswith('Muhammad '):
        return 'Muh.' + name[9:]
    return name

# ── STYLES ────────────────────────────────────────────────────────────────────
def make_styles():
    return {
        'title': ParagraphStyle('title',
            fontName='Helvetica-Bold', fontSize=12,
            alignment=TA_CENTER, leading=16, spaceAfter=0),

        'cell_no': ParagraphStyle('cell_no',
            fontName='Helvetica', fontSize=9,
            alignment=TA_CENTER, leading=11),

        'cell_indi': ParagraphStyle('cell_indi',
            fontName='Helvetica', fontSize=9.5,
            alignment=TA_LEFT, leading=12),

        'cell_mid': ParagraphStyle('cell_mid',
            fontName='Helvetica', fontSize=9.5,
            alignment=TA_CENTER, leading=11),

        'cell_hdr': ParagraphStyle('cell_hdr',
            fontName='Helvetica-Bold', fontSize=9.5,
            alignment=TA_CENTER, leading=12),

        'ttd_city': ParagraphStyle('ttd_city',
            fontName='Helvetica', fontSize=11,
            alignment=TA_LEFT, leading=14),

        'ttd_role': ParagraphStyle('ttd_role',
            fontName='Helvetica', fontSize=11,
            alignment=TA_LEFT, leading=14),

        'ttd_name': ParagraphStyle('ttd_name',
            fontName='Helvetica-Bold', fontSize=11,
            alignment=TA_LEFT, leading=14),

        'ttd_nip': ParagraphStyle('ttd_nip',
            fontName='Helvetica', fontSize=10,
            alignment=TA_LEFT, leading=12),
    }

# ── KOP ───────────────────────────────────────────────────────────────────────
def make_kop_image(b64_data):
    """Return an RLImage of the KOP spanning full page width."""
    img_bytes = base64.b64decode(b64_data)
    img_io = io.BytesIO(img_bytes)
    kop = RLImage(img_io, width=PW, height=KOP_PRINT_H)
    kop.hAlign = 'LEFT'
    return kop

# ── TABLE ─────────────────────────────────────────────────────────────────────
def make_table(capaian, year, q, styles):
    mLabels = M_LABELS.get(q, ['M1', 'M2', 'M3'])
    qShort  = f'Q{q}'

    P = Paragraph  # shorthand

    # Header row 1: merged cols 0+1 rowspan 2, TARGET rowspan 2, CAPAIAN span 4
    # ReportLab SPAN uses (col, row) notation — (0,0) to (1,1) = merge cols 0&1 rows 0&1
    hdr1 = [
        P('INDIKATOR<br/>KINERJA', styles['cell_hdr']),
        '',   # merged
        P('TARGET', styles['cell_hdr']),
        P('CAPAIAN', styles['cell_hdr']),
        '',   # merged
        '',   # merged
        '',   # merged
    ]
    hdr2 = [
        '',   # merged (rowspan from row 0)
        '',   # merged
        '',   # merged (rowspan from row 0)
        P(mLabels[0][:3].upper(), styles['cell_hdr']),
        P(mLabels[1][:3].upper(), styles['cell_hdr']),
        P(mLabels[2][:3].upper(), styles['cell_hdr']),
        P(qShort, styles['cell_hdr']),
    ]

    data = [hdr1, hdr2]

    for idx, iki in enumerate(LIM_IKI):
        m1 = get_cap(capaian, year, q, iki['no'], 'm1')
        m2 = get_cap(capaian, year, q, iki['no'], 'm2')
        m3 = get_cap(capaian, year, q, iki['no'], 'm3')
        qv = get_cap(capaian, year, q, iki['no'], 'q')
        bg = C_ODD_ROW if idx % 2 == 0 else C_EVEN_ROW

        data.append([
            P(iki['no'],    styles['cell_no']),
            P(iki['indi'],  styles['cell_indi']),
            P(iki['target'],styles['cell_mid']),
            P(m1,           styles['cell_mid']),
            P(m2,           styles['cell_mid']),
            P(m3,           styles['cell_mid']),
            P(f'<b>{qv}</b>' if qv else '', styles['cell_mid']),
        ])

    n_data = len(data)

    tbl = Table(data, colWidths=COLS, repeatRows=2)

    style_cmds = [
        # ── Borders ──
        ('GRID',        (0, 0), (-1, -1), 0.5, C_BORDER),
        ('BOX',         (0, 0), (-1, -1), 0.75, C_BORDER),

        # ── Header background ──
        ('BACKGROUND',  (0, 0), (-1, 1),  C_HEADER_BG),

        # ── Header spans ──
        # Cols 0+1, rows 0-1 → merge into one "INDIKATOR KINERJA" cell
        ('SPAN',        (0, 0), (1, 1)),
        # Col 2, rows 0-1 → "TARGET"
        ('SPAN',        (2, 0), (2, 1)),
        # Cols 3-6, row 0 → "CAPAIAN"
        ('SPAN',        (3, 0), (6, 0)),

        # ── Header text alignment ──
        ('VALIGN',      (0, 0), (-1, 1), 'MIDDLE'),
        ('ALIGN',       (0, 0), (-1, 1), 'CENTER'),
        ('FONTNAME',    (0, 0), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE',    (0, 0), (-1, 1), 9.5),

        # ── Data rows ──
        ('VALIGN',      (0, 2), (-1, -1), 'MIDDLE'),
        ('FONTNAME',    (0, 2), (-1, -1), 'Helvetica'),
        ('FONTSIZE',    (0, 2), (-1, -1), 9.5),
        ('TOPPADDING',  (0, 2), (-1, -1), 3),
        ('BOTTOMPADDING',(0,2), (-1, -1), 3),
        ('LEFTPADDING', (1, 2), (1, -1),  5),
        ('TOPPADDING',  (0, 0), (-1, 1),  4),
        ('BOTTOMPADDING',(0,0), (-1, 1),  4),
    ]

    # Alternating row backgrounds
    for i in range(len(LIM_IKI)):
        row = i + 2
        bg = C_ODD_ROW if i % 2 == 0 else C_EVEN_ROW
        style_cmds.append(('BACKGROUND', (0, row), (-1, row), bg))

    tbl.setStyle(TableStyle(style_cmds))
    return tbl

# ── SIGNATURE ─────────────────────────────────────────────────────────────────
def make_signature(nama, nip, report_date, styles):
    """Right-aligned signature block matching reference exactly."""
    name_short = shorten_muhammad(nama)

    sig_data = [
        [Paragraph(f'Kolaka,&nbsp;&nbsp; {report_date}', styles['ttd_city'])],
        [Paragraph('Ketua Tim,', styles['ttd_role'])],
        [Spacer(1, 46)],
        [Paragraph(f'<u>{name_short}</u>', styles['ttd_name'])],
        [Paragraph(f'NIP. {nip}', styles['ttd_nip'])],
    ]
    sig_col_w = 200
    sig_table = Table(sig_data, colWidths=[sig_col_w])
    sig_table.setStyle(TableStyle([
        ('ALIGN',        (0, 0), (-1, -1), 'LEFT'),
        ('TOPPADDING',   (0, 0), (-1, -1), 1),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 1),
        ('LEFTPADDING',  (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
    ]))

    # Wrap in a right-aligning outer table
    outer = Table([[sig_table]], colWidths=[CW])
    outer.setStyle(TableStyle([
        ('ALIGN',       (0, 0), (-1, -1), 'RIGHT'),
        ('TOPPADDING',  (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING',(0,0), (-1, -1), 0),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',(0, 0), (-1, -1), 0),
    ]))
    return outer

# ── PAGE TEMPLATES ────────────────────────────────────────────────────────────
class LIMDocTemplate(BaseDocTemplate):
    """Custom doc with KOP on page 1 only."""

    def __init__(self, filename, kop_b64, **kwargs):
        super().__init__(filename, **kwargs)
        self.kop_b64 = kop_b64
        self._build_templates()

    def _build_templates(self):
        # Page 1: smaller content frame (KOP takes top)
        kop_h = KOP_PRINT_H
        gap   = 10   # pts gap between KOP rule and content
        frame1 = Frame(
            ML,                          # x
            MB,                          # y
            CW,                          # width
            PH - kop_h - 2 - gap - MB,   # height (2pts for rule)
            leftPadding=0, rightPadding=0,
            topPadding=0, bottomPadding=0,
            id='page1_frame'
        )
        # Pages 2+: full content frame with small top gap
        frame_rest = Frame(
            ML, MB, CW,
            PH - 15*mm - MB,
            leftPadding=0, rightPadding=0,
            topPadding=0, bottomPadding=0,
            id='rest_frame'
        )

        def page1_bg(canvas, doc):
            canvas.saveState()
            # Draw KOP full bleed using ImageReader (supports BytesIO)
            try:
                from reportlab.lib.utils import ImageReader
                kop_bytes = base64.b64decode(doc.kop_b64)
                kop_io = io.BytesIO(kop_bytes)
                ir = ImageReader(kop_io)
                canvas.drawImage(
                    ir,
                    0, PH - KOP_PRINT_H,
                    width=PW,
                    height=KOP_PRINT_H,
                    preserveAspectRatio=True
                )
            except Exception as e:
                # KOP unavailable — draw placeholder
                canvas.setFillColor(colors.HexColor('#CCCCCC'))
                canvas.rect(0, PH - KOP_PRINT_H, PW, KOP_PRINT_H, fill=1, stroke=0)
            # Black rule below KOP
            canvas.setStrokeColor(colors.black)
            canvas.setLineWidth(1.5)
            canvas.line(0, PH - KOP_PRINT_H - 1.5, PW, PH - KOP_PRINT_H - 1.5)
            # Footer page number
            canvas.setFont('Helvetica', 8)
            canvas.setFillColor(colors.grey)
            canvas.drawCentredString(PW/2, MB/2, f'Halaman {doc.page}')
            canvas.restoreState()

        def rest_bg(canvas, doc):
            canvas.saveState()
            canvas.setFont('Helvetica', 8)
            canvas.setFillColor(colors.grey)
            canvas.drawCentredString(PW/2, MB/2, f'Halaman {doc.page}')
            canvas.restoreState()

        self.addPageTemplates([
            PageTemplate(id='FirstPage',  frames=[frame1], onPage=page1_bg),
            PageTemplate(id='LaterPages', frames=[frame_rest], onPage=rest_bg),
        ])

# ── MAIN BUILD ────────────────────────────────────────────────────────────────
def build_pdf(state, year, q, output_path, kop_b64):
    capaian  = state.get('capaian', {})
    prof     = state.get('profile', {})
    nama     = prof.get('nama', 'Muhammad Subhan Al Zibrah, S.Tr.Met.')
    nip      = prof.get('nip',  '200009262023021004')

    # Resolve report date
    cap_key  = f"{year}-Q{q}"
    rep_date = state.get('capaianReportDate', {}).get(cap_key, '')
    if not rep_date:
        rep_date = fmt_date_id(wita_now())

    styles  = make_styles()
    ql      = Q_LABELS.get(q, f'Triwulan {q}')
    doc     = LIMDocTemplate(
        output_path, kop_b64=kop_b64,
        pagesize=A4,
        leftMargin=ML, rightMargin=MR,
        topMargin=MT, bottomMargin=MB,
        title=f'Laporan Capaian IKI LIM — {ql} {year}',
        author=nama,
    )

    story = [NextPageTemplate('FirstPage')]

    # Titles
    story.append(Spacer(1, 8))
    for line in [
        'LAPORAN EVALUASI CAPAIAN INDIKATOR KINERJA INDIVIDU',
        'TIM LAYANAN INFORMASI METEOROLOGI',
        f'{ql.upper()} TAHUN {year}',
    ]:
        story.append(Paragraph(line, styles['title']))
        story.append(Spacer(1, 2))
    story.append(Spacer(1, 8))

    # Table (continues to next page automatically, header repeats)
    story.append(NextPageTemplate('LaterPages'))
    story.append(make_table(capaian, year, q, styles))

    # Signature
    story.append(Spacer(1, 6))
    story.append(make_signature(nama, nip, rep_date, styles))

    doc.build(story)
    return output_path


# ── CLI ───────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description='Generate PDF Laporan Capaian IKI LIM dari file JSON ekspor')
    parser.add_argument('json_file', nargs='?', help='Path ke file JSON ekspor')
    parser.add_argument('--q',    type=int, default=0,    help='Triwulan (1-4), default: auto dari bulan WITA')
    parser.add_argument('--year', type=int, default=0,    help='Tahun, default: tahun WITA saat ini')
    parser.add_argument('--out',  type=str, default='',   help='Path output PDF')
    parser.add_argument('--kop',  type=str, default='',   help='Path gambar KOP (PNG/JPG), opsional')
    args = parser.parse_args()

    # Find JSON
    json_path = args.json_file
    if not json_path:
        json_path = find_json()
        if not json_path:
            print("ERROR: Tidak ada file JSON ditemukan.")
            print("Ekspor data dari aplikasi terlebih dahulu, lalu jalankan:")
            print("  python generate_capaian_pdf.py LIM_data_YYYY-MM-DD.json")
            sys.exit(1)
        print(f"Menggunakan: {json_path}")

    if not os.path.exists(json_path):
        print(f"ERROR: File tidak ditemukan: {json_path}")
        sys.exit(1)

    state = load_state(json_path)

    # Year & quarter
    now = wita_now()
    year = args.year or now.year
    if args.q:
        q = args.q
    else:
        # Auto-detect from current month
        q = (now.month - 1) // 3 + 1
    print(f"Periode: {Q_LABELS.get(q, q)} {year}")

    # KOP — try: CLI arg → state kop_img → same folder kop.png/jpg
    kop_b64 = None
    if args.kop and os.path.exists(args.kop):
        from PIL import Image as PI
        img = PI.open(args.kop).convert('RGB')
        buf = io.BytesIO()
        w, h = img.size
        nw, nh = 1400, int(h * 1400 / w)
        img.resize((nw, nh), PI.LANCZOS).save(buf, format='JPEG', quality=94)
        kop_b64 = base64.b64encode(buf.getvalue()).decode()
        print(f"KOP: {args.kop} ({nw}x{nh}px)")
    elif 'KOP_IMG_B64' in state:
        raw = state['KOP_IMG_B64']
        if raw.startswith('data:'):
            raw = raw.split(',', 1)[1]
        kop_b64 = raw
        print("KOP: dari state JSON")
    else:
        # Look for kop image next to json file
        base_dir = os.path.dirname(os.path.abspath(json_path))
        for fname in ['Kop.png','Kop.jpg','kop.png','kop.jpg','Kop__1_.png']:
            candidate = os.path.join(base_dir, fname)
            if os.path.exists(candidate):
                from PIL import Image as PI
                img = PI.open(candidate).convert('RGB')
                buf = io.BytesIO()
                w, h = img.size
                nw, nh = 1400, int(h * 1400 / w)
                img.resize((nw, nh), PI.LANCZOS).save(buf, format='JPEG', quality=94)
                kop_b64 = base64.b64encode(buf.getvalue()).decode()
                print(f"KOP: {candidate}")
                break

    if not kop_b64:
        print("PERINGATAN: Gambar KOP tidak ditemukan. PDF akan dibuat tanpa KOP.")
        print("Letakkan file 'Kop.png' di folder yang sama, atau gunakan --kop path/ke/kop.png")

    # Output path
    out = args.out
    if not out:
        base_dir = os.path.dirname(os.path.abspath(json_path))
        ts = now.strftime('%Y%m%d_%H%M%S')
        out = os.path.join(base_dir, f"LIM_Capaian_Q{q}_{year}_{ts}.pdf")

    print(f"Membuat PDF: {out}")
    build_pdf(state, year, q, out, kop_b64)
    print(f"✅ Selesai: {out}")

    # Auto-open
    try:
        import subprocess, platform
        if platform.system() == 'Windows':
            os.startfile(out)
        elif platform.system() == 'Darwin':
            subprocess.run(['open', out])
        else:
            subprocess.run(['xdg-open', out])
    except Exception:
        pass


if __name__ == '__main__':
    main()
