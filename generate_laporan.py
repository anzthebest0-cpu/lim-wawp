#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_laporan.py — Pembuat PDF Laporan Kinerja Ketua Tim LIM
BMKG · Stasiun Meteorologi Kelas III Sangia Nibandera Kolaka (WAWP)

Penggunaan:
  python generate_laporan.py [path_ke_file_json]

  Jika path tidak disertakan, skrip akan mencari file LIM_data_*.json
  atau LIM_WAWP_*.json terbaru di folder yang sama.

Dependensi:
  pip install reportlab Pillow

Output:
  LIM_Laporan_YYYYMM_<timestamp>.pdf  (di folder yang sama dengan file JSON)
"""

import sys
import json
import os
import glob
import base64
import io
from datetime import datetime, timezone, timedelta

# ── Cek dependensi ────────────────────────────────────────────────────────────
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm, mm
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        HRFlowable, KeepTogether, PageBreak
    )
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.graphics.shapes import Drawing, Rect, String
except ImportError:
    print("=" * 60)
    print("ERROR: ReportLab belum terinstal.")
    print("Jalankan perintah berikut lalu coba lagi:")
    print("  pip install reportlab Pillow")
    print("=" * 60)
    sys.exit(1)

# ── KONSTANTA WARNA ───────────────────────────────────────────────────────────
C_TEAL       = colors.HexColor('#0F6E56')
C_TEAL_MID   = colors.HexColor('#1D9E75')
C_TEAL_LIGHT = colors.HexColor('#E1F5EE')
C_TEAL_PALE  = colors.HexColor('#F0FAF6')
C_BLUE       = colors.HexColor('#185FA5')
C_BLUE_LIGHT = colors.HexColor('#E6F1FB')
C_AMBER      = colors.HexColor('#854F0B')
C_AMBER_LT   = colors.HexColor('#FEF3CD')
C_RED        = colors.HexColor('#A32D2D')
C_RED_LIGHT  = colors.HexColor('#FCEBEB')
C_GRAY_LT    = colors.HexColor('#F7F6F2')
C_GRAY_BD    = colors.HexColor('#D3D1C7')
C_GRAY_MID   = colors.HexColor('#888780')
C_BLACK      = colors.black
C_WHITE      = colors.white

# ── KONSTANTA INSTITUSI ───────────────────────────────────────────────────────
NAMA_BMKG     = "BADAN METEOROLOGI, KLIMATOLOGI, DAN GEOFISIKA"
NAMA_STASIUN  = "STASIUN METEOROLOGI KELAS III SANGIA NIBANDERA KOLAKA"
ICAO          = "WAWP"
ALAMAT        = "Jalan Protokol No. 1, Pomalaa, Kolaka 93562"
KONTAK        = "Telp: (0405) 2401622 · Email: stamet.kolaka@bmkg.go.id"
KEPALA        = "Danu Triatmoko"
NIP_KEPALA    = "—"

# ── DEFAULT PROFIL (override dari JSON jika ada) ──────────────────────────────
DEFAULT_NAMA    = "Muhammad Subhan Al Zibrah, S.Tr.Met."
DEFAULT_NIP     = "200009262023021004"
DEFAULT_JABATAN = "Ketua Tim Unit Layanan Informasi Meteorologi"
DEFAULT_PANGKAT = "Penata Muda / III a"

# ── UTILITAS TANGGAL ──────────────────────────────────────────────────────────
BULAN_ID = [
    '', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
]

def wita_now():
    """Waktu sekarang dalam zona waktu WITA (UTC+8)."""
    return datetime.now(timezone(timedelta(hours=8)))

def fmt_tanggal(dt):
    return f"{dt.day} {BULAN_ID[dt.month]} {dt.year}"

def fmt_bulan_tahun(tahun, bulan):
    return f"{BULAN_ID[bulan]} {tahun}"

def cadence_label(c):
    return {'harian': 'Harian', 'mingguan': 'Mingguan', 'bulanan': 'Bulanan',
            'triwulanan': 'Triwulanan', 'tahunan': 'Tahunan'}.get(c, c.title())

# ── BACA DAN VALIDASI JSON ─────────────────────────────────────────────────────
def find_json_file():
    patterns = ['LIM_data_*.json', 'LIM_WAWP_*.json', 'lim_*.json']
    candidates = []
    for p in patterns:
        candidates.extend(glob.glob(p))
    if not candidates:
        return None
    return max(candidates, key=os.path.getmtime)

def load_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        raw = json.load(f)

    # Support both old format (flat) and new schema-versioned format
    if 'state' in raw and 'meta' in raw:
        # New schema-versioned format
        return raw['state'], raw['meta']
    elif '_meta' in raw:
        # Mid-version format
        meta = raw.pop('_meta', {})
        return raw, meta
    else:
        # Legacy flat format
        return raw, {}

def extract_state(state, meta):
    """Normalisasi state dari berbagai versi JSON export."""
    profile = state.get('profile', {})
    nama    = profile.get('nama', DEFAULT_NAMA)
    nip     = profile.get('nip', DEFAULT_NIP)
    jabatan = profile.get('jabatan', DEFAULT_JABATAN)
    pangkat = profile.get('pangkat', DEFAULT_PANGKAT)

    period  = state.get('period', 'Periode Aktif')
    period_note = state.get('periodNote', '')

    items_state = state.get('items', {})
    reports_state = state.get('reports', {})
    archives = state.get('archives', [])
    daily_history = state.get('dailyHistory', [])

    # Custom tasks override
    custom_tasks = state.get('customTasks', [])
    custom_reports = state.get('customReports', [])

    export_time = meta.get('exportedAt') or meta.get('exported_at', '')
    if export_time:
        try:
            dt = datetime.fromisoformat(export_time.replace('Z', '+00:00'))
            export_time_fmt = dt.astimezone(timezone(timedelta(hours=8))).strftime('%d %B %Y, %H:%M WITA')
        except Exception:
            export_time_fmt = export_time
    else:
        export_time_fmt = fmt_tanggal(wita_now())

    return {
        'nama': nama, 'nip': nip, 'jabatan': jabatan, 'pangkat': pangkat,
        'period': period, 'period_note': period_note,
        'items': items_state, 'reports': reports_state,
        'archives': archives, 'daily_history': daily_history,
        'custom_tasks': custom_tasks, 'custom_reports': custom_reports,
        'export_time': export_time_fmt,
        'print_time': fmt_tanggal(wita_now()),
        'print_wita': wita_now().strftime('%d %B %Y, %H:%M WITA'),
    }

# ── DEFINISI DEFAULT ITEMS (sama dengan HTML) ─────────────────────────────────
DEFAULT_ITEMS = [
    {'id':'h1','cadence':'harian','label':'Pantau akurasi & ketepatan waktu METAR','pills':[{'t':'sup','v':'METAR 100%'}]},
    {'id':'h2','cadence':'harian','label':'Pantau akurasi TAF & verifikasi harian','pills':[{'t':'sup','v':'TAF 97,5%'}]},
    {'id':'h3','cadence':'harian','label':'Pantau kelengkapan Flight Documentation di INA-SIAM','pills':[{'t':'sup','v':'Flight Doc 100%'}]},
    {'id':'h4','cadence':'harian','label':'Pantau penerbitan AD WRNG sesuai standar','pills':[{'t':'sup','v':'AD WRNG 97,5%'}]},
    {'id':'h5','cadence':'harian','label':'Tangani umpan balik stakeholder eksternal','pills':[{'t':'sup','v':'IKM ≥ 3,8'}]},
    {'id':'w1','cadence':'mingguan','label':'Tinjau progres Monthly Aviation Weather Summary','pills':[{'t':'own','v':'IKI personal: 12 dok/tahun'}]},
    {'id':'w2','cadence':'mingguan','label':'Tinjau rekap deviasi & insiden operasional','pills':[{'t':'sup','v':'Basis evaluasi bulanan'}]},
    {'id':'w3','cadence':'mingguan','label':'Eskalasi ke Kepala Stasiun jika ada isu kritis','pills':[{'t':'collab','v':'LAKjIP target: BB'}]},
    {'id':'m1','cadence':'bulanan','label':'Pastikan laporan operasional Portal dikirim tepat waktu','pills':[{'t':'own','v':'IKI personal: 4 indeks/tahun'}]},
    {'id':'m2','cadence':'bulanan','label':'Terbitkan Monthly Aviation Weather Summary bulan ini','pills':[{'t':'own','v':'IKI co-owner: 12 dok/tahun'}]},
    {'id':'m3','cadence':'bulanan','label':'Pimpin rapat evaluasi internal & buat notulen','pills':[{'t':'collab','v':'IKI Tim Kerja: 4 laporan/tahun'}]},
    {'id':'m4','cadence':'bulanan','label':'Susun laporan perkembangan layanan ke Kepala Stasiun','pills':[{'t':'sup','v':'Basis: laporan evaluasi tim'}]},
    {'id':'m5','cadence':'bulanan','label':'Pantau progres Kolaka Monthly Weather Report','pills':[{'t':'sup','v':'Supervisi: 12 dok/tahun'}]},
    {'id':'m6','cadence':'bulanan','label':'Verifikasi tindak lanjut action items rapat bulan lalu','pills':[{'t':'collab','v':'IKPA target: 97'}]},
    {'id':'q1','cadence':'triwulanan','label':'Submit laporan operasional via Portal (1 indeks per triwulan)','pills':[{'t':'own','v':'IKI personal: indeks ke-1/2/3/4'}]},
    {'id':'q2','cadence':'triwulanan','label':'Pastikan laporan monitoring & evaluasi Tim Kerja tersusun','pills':[{'t':'collab','v':'IKI lintas tim: 4 laporan/tahun'}]},
    {'id':'q3','cadence':'triwulanan','label':'Laksanakan sosialisasi informasi meteorologi penerbangan','pills':[{'t':'sup','v':'Supervisi: 4 laporan triwulanan'}]},
    {'id':'q4','cadence':'triwulanan','label':'Pantau penyelesaian analisis fenomena meteorologi signifikan','pills':[{'t':'sup','v':'Supervisi: 4 indeks/tahun'}]},
    {'id':'a1','cadence':'tahunan','label':'Koordinasikan penerbitan Catatan Iklim Kolaka','pills':[{'t':'sup','v':'Supervisi: 1 dok/tahun'}]},
    {'id':'a2','cadence':'tahunan','label':'Pastikan pelaksanaan Survey Kepuasan Masyarakat (2x setahun)','pills':[{'t':'sup','v':'IKM target: ≥ 3,8 Skala Likert'}]},
    {'id':'a3','cadence':'tahunan','label':'Siapkan kontribusi untuk dokumen LAKjIP satker','pills':[{'t':'collab','v':'IKI lintas tim: LAKjIP BB'}]},
    {'id':'a4','cadence':'tahunan','label':'Siapkan dokumen evaluasi anggaran untuk Ravalnas','pills':[{'t':'collab','v':'IKI lintas tim: 1 dokumen'}]},
    {'id':'a5','cadence':'tahunan','label':'Susun inovasi layanan publik bersama tim','pills':[{'t':'own','v':'Tugas SK: inovasi tahunan'}]},
    {'id':'a6','cadence':'tahunan','label':'Dukung pencapaian nilai IKPA satker ≥ 97','pills':[{'t':'collab','v':'IKI lintas tim: IKPA 97'}]},
]

DEFAULT_REPORTS = [
    {'id':'rpt_metar',    'name':'Laporan Evaluasi Akurasi METAR',               'pic':'Satriawan N. Atsidiqi, S.Tr, M.Si'},
    {'id':'rpt_taf',      'name':'Laporan Evaluasi Akurasi TAF',                  'pic':'Anwar Budi Nugroho, S.Tr.Met.'},
    {'id':'rpt_flightdoc','name':'Laporan Evaluasi Flight Documentation',          'pic':'Hijrah K. Musgamy, S.Si, M.Si'},
    {'id':'rpt_adwrng',   'name':'Laporan Evaluasi Akurasi AD WRNG',              'pic':'Adi Kusuma Nugraha, S.Tr.Met'},
    {'id':'rpt_stat',     'name':'Laporan Monitoring Data Statistik Penerbangan', 'pic':'M. Figo Ramadhan, S.Tr.Met.'},
    {'id':'rpt_acs',      'name':'Laporan Penyusunan Dokumen ACS',                'pic':'Safinatunnajah Dinda Putri, S.Tr. Met.'},
    {'id':'rpt_mwr',      'name':'Kolaka Monthly Weather Report',                 'pic':'Hijrah, Safinatunnajah, Dwi, Rainy, Faisal, Yasser'},
    {'id':'rpt_maws',     'name':'Monthly Aviation Weather Summary',              'pic':'Subhan, Adi, Yasser, Figo'},
    {'id':'rpt_ekstrem',  'name':'Rekapitulasi Data Parameter Cuaca Ekstrem',     'pic':'Hijrah K. Musgamy, S.Si, M.Si'},
]

# ── STYLE SETUP ───────────────────────────────────────────────────────────────
def build_styles():
    base = getSampleStyleSheet()
    styles = {}

    styles['kop_org'] = ParagraphStyle('kop_org',
        fontName='Helvetica-Bold', fontSize=12, textColor=C_TEAL,
        alignment=TA_LEFT, leading=15, spaceAfter=1)

    styles['kop_sub'] = ParagraphStyle('kop_sub',
        fontName='Helvetica', fontSize=9, textColor=colors.HexColor('#5A5955'),
        alignment=TA_LEFT, leading=12)

    styles['kop_addr'] = ParagraphStyle('kop_addr',
        fontName='Helvetica', fontSize=8, textColor=colors.HexColor('#9A9890'),
        alignment=TA_LEFT, leading=11)

    styles['doc_title'] = ParagraphStyle('doc_title',
        fontName='Helvetica-Bold', fontSize=12, textColor=C_BLACK,
        alignment=TA_CENTER, leading=16, spaceBefore=4, spaceAfter=2)

    styles['doc_sub'] = ParagraphStyle('doc_sub',
        fontName='Helvetica', fontSize=9, textColor=colors.HexColor('#5A5955'),
        alignment=TA_CENTER, leading=12, spaceAfter=2)

    styles['section_title'] = ParagraphStyle('section_title',
        fontName='Helvetica-Bold', fontSize=9, textColor=C_TEAL,
        leading=12, spaceBefore=10, spaceAfter=4,
        borderPadding=(0, 0, 2, 0))

    styles['cell_label'] = ParagraphStyle('cell_label',
        fontName='Helvetica', fontSize=8, textColor=colors.HexColor('#5A5955'),
        leading=11)

    styles['cell_value'] = ParagraphStyle('cell_value',
        fontName='Helvetica', fontSize=8, textColor=C_BLACK, leading=11)

    styles['cell_bold'] = ParagraphStyle('cell_bold',
        fontName='Helvetica-Bold', fontSize=8, textColor=C_BLACK, leading=11)

    styles['status_done']    = ParagraphStyle('status_done',    fontName='Helvetica-Bold', fontSize=8, textColor=C_TEAL_MID, leading=11)
    styles['status_pending'] = ParagraphStyle('status_pending', fontName='Helvetica',      fontSize=8, textColor=C_GRAY_MID, leading=11)
    styles['status_flagged'] = ParagraphStyle('status_flagged', fontName='Helvetica-Bold', fontSize=8, textColor=C_RED,      leading=11)

    styles['note_text'] = ParagraphStyle('note_text',
        fontName='Helvetica', fontSize=8, textColor=colors.HexColor('#5A5955'),
        leading=12, alignment=TA_JUSTIFY)

    styles['footer'] = ParagraphStyle('footer',
        fontName='Helvetica', fontSize=7, textColor=C_GRAY_MID,
        alignment=TA_CENTER, leading=10)

    return styles

# ── HEADER / KOP SURAT ────────────────────────────────────────────────────────
def build_kop(styles, logo_data=None):
    elements = []

    # Logo + teks institusi dalam tabel
    if logo_data:
        try:
            img_bytes = base64.b64decode(logo_data)
            img_io = io.BytesIO(img_bytes)
            from reportlab.platypus import Image as RLImage
            logo_img = RLImage(img_io, width=1.4*cm, height=1.4*cm)
            logo_cell = logo_img
        except Exception:
            logo_cell = ''
    else:
        logo_cell = ''

    org_text = [
        Paragraph(NAMA_BMKG, styles['kop_org']),
        Paragraph(NAMA_STASIUN, styles['kop_sub']),
        Paragraph(f"{ALAMAT}  ·  {KONTAK}", styles['kop_addr']),
    ]

    kop_table = Table(
        [[logo_cell, org_text]],
        colWidths=[1.8*cm, 15.5*cm],
        hAlign='LEFT'
    )
    kop_table.setStyle(TableStyle([
        ('VALIGN',    (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING',  (0,0), (0,0), 0),
        ('RIGHTPADDING', (0,0), (0,0), 8),
        ('TOPPADDING',   (0,0), (-1,-1), 0),
        ('BOTTOMPADDING',(0,0), (-1,-1), 0),
    ]))
    elements.append(kop_table)
    elements.append(HRFlowable(width='100%', thickness=2, color=C_TEAL, spaceAfter=6))
    return elements

# ── JUDUL DOKUMEN ─────────────────────────────────────────────────────────────
def build_title(styles, data):
    elements = []
    elements.append(Paragraph(
        'LAPORAN KINERJA KETUA TIM UNIT LAYANAN INFORMASI METEOROLOGI',
        styles['doc_title']
    ))
    elements.append(Paragraph(
        f"Berdasarkan SK KEP.20/KPUM/I/2026 dan Cascading IKU/IKI 2026",
        styles['doc_sub']
    ))
    elements.append(Paragraph(
        f"Periode: {data['period']}  ·  Dicetak: {data['print_wita']}",
        styles['doc_sub']
    ))
    elements.append(HRFlowable(width='100%', thickness=0.5, color=C_GRAY_BD, spaceAfter=8))
    return elements

# ── IDENTITAS PEJABAT ─────────────────────────────────────────────────────────
def build_identitas(styles, data):
    elements = []
    elements.append(Paragraph('A. IDENTITAS PEJABAT YANG BERTANGGUNG JAWAB', styles['section_title']))

    rows = [
        ['Nama',            ':', data['nama']],
        ['NIP',             ':', data['nip']],
        ['Jabatan',         ':', data['jabatan']],
        ['Pangkat/Golongan',':', data['pangkat']],
        ['Unit Kerja',      ':', f"{NAMA_STASIUN} ({ICAO})"],
        ['Periode Laporan', ':', data['period'] + (f" — {data['period_note']}" if data['period_note'] else '')],
        ['Data Diekspor',   ':', data['export_time']],
    ]

    cell_data = []
    for r in rows:
        cell_data.append([
            Paragraph(r[0], styles['cell_label']),
            Paragraph(r[1], styles['cell_label']),
            Paragraph(r[2], styles['cell_value']),
        ])

    t = Table(cell_data, colWidths=[3.5*cm, 0.4*cm, 13.4*cm], hAlign='LEFT')
    t.setStyle(TableStyle([
        ('VALIGN',    (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING',   (0,0), (-1,-1), 2),
        ('BOTTOMPADDING',(0,0), (-1,-1), 2),
        ('LEFTPADDING',  (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    elements.append(t)
    return elements

# ── RINGKASAN CAPAIAN ─────────────────────────────────────────────────────────
def build_ringkasan(styles, items_list, items_state):
    elements = []
    elements.append(Paragraph('B. RINGKASAN CAPAIAN CHECKLIST', styles['section_title']))

    CADENCES = ['harian','mingguan','bulanan','triwulanan','tahunan']
    total_done = sum(1 for it in items_list if items_state.get(it['id'],{}).get('done', False))
    total_all  = len(items_list)
    flagged    = sum(1 for it in items_list if items_state.get(it['id'],{}).get('flagged', False))
    pct = round(total_done / total_all * 100) if total_all else 0

    # Summary stats row
    stat_data = [[
        [Paragraph(f"{pct}%", ParagraphStyle('pct', fontName='Helvetica-Bold', fontSize=20,
                   textColor=C_TEAL_MID if pct >= 80 else C_AMBER if pct >= 50 else C_RED, alignment=TA_CENTER)),
         Paragraph('Total Capaian', styles['cell_label'])],

        [Paragraph(str(total_done), ParagraphStyle('sn', fontName='Helvetica-Bold', fontSize=18,
                   textColor=C_TEAL_MID, alignment=TA_CENTER)),
         Paragraph('Tugas Selesai', styles['cell_label'])],

        [Paragraph(str(total_all - total_done), ParagraphStyle('sn', fontName='Helvetica-Bold', fontSize=18,
                   textColor=C_BLACK, alignment=TA_CENTER)),
         Paragraph('Masih Tertunda', styles['cell_label'])],

        [Paragraph(str(flagged), ParagraphStyle('sn', fontName='Helvetica-Bold', fontSize=18,
                   textColor=C_RED if flagged > 0 else C_GRAY_MID, alignment=TA_CENTER)),
         Paragraph('Ditandai Merah', styles['cell_label'])],
    ]]

    def make_stat_cell(content):
        inner = Table([[c] for c in content], colWidths=[4.1*cm])
        inner.setStyle(TableStyle([
            ('ALIGN',    (0,0), (-1,-1), 'CENTER'),
            ('VALIGN',   (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING',   (0,0), (-1,-1), 4),
            ('BOTTOMPADDING',(0,0), (-1,-1), 4),
        ]))
        return inner

    stat_table = Table(
        [[make_stat_cell(c) for c in stat_data[0]]],
        colWidths=[4.2*cm]*4,
        hAlign='LEFT'
    )
    stat_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), C_GRAY_LT),
        ('GRID', (0,0), (-1,-1), 0.5, C_GRAY_BD),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    elements.append(stat_table)
    elements.append(Spacer(1, 4))

    # Per-cadence breakdown
    cad_data = [['Kadence', 'Selesai', 'Total', 'Capaian']]
    for c in CADENCES:
        citems = [it for it in items_list if it['cadence'] == c]
        cdone  = sum(1 for it in citems if items_state.get(it['id'],{}).get('done', False))
        cpct   = round(cdone / len(citems) * 100) if citems else 0
        cad_data.append([cadence_label(c), str(cdone), str(len(citems)), f"{cpct}%"])

    cad_table = Table(cad_data, colWidths=[4*cm, 2.5*cm, 2.5*cm, 2.5*cm], hAlign='LEFT')
    cad_table.setStyle(TableStyle([
        ('FONTNAME',    (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTNAME',    (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE',    (0,0), (-1,-1), 8),
        ('BACKGROUND',  (0,0), (-1,0), C_TEAL_LIGHT),
        ('TEXTCOLOR',   (0,0), (-1,0), C_TEAL),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [C_WHITE, C_GRAY_LT]),
        ('GRID',        (0,0), (-1,-1), 0.5, C_GRAY_BD),
        ('ALIGN',       (1,0), (-1,-1), 'CENTER'),
        ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',  (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(cad_table)
    return elements

# ── TABEL CHECKLIST PER KADENCE ───────────────────────────────────────────────
def build_checklist_tables(styles, items_list, items_state):
    elements = []
    elements.append(Paragraph('C. DETAIL TUGAS PER KADENCE', styles['section_title']))

    CADENCES = ['harian','mingguan','bulanan','triwulanan','tahunan']

    for c in CADENCES:
        citems = [it for it in items_list if it['cadence'] == c]
        if not citems:
            continue

        cdone = sum(1 for it in citems if items_state.get(it['id'],{}).get('done', False))
        header_text = f"{cadence_label(c).upper()}  —  {cdone}/{len(citems)} selesai"

        col_rows = [['No.', 'Uraian Tugas', 'Status', 'IKI / Target', 'Catatan']]
        for i, it in enumerate(citems, 1):
            st = items_state.get(it['id'], {})
            done    = st.get('done', False)
            flagged = st.get('flagged', False)
            note    = st.get('note', '') or '—'
            pills   = ', '.join(p['v'] for p in it.get('pills', []))

            if flagged:
                status_p = Paragraph('! Perlu Perhatian', styles['status_flagged'])
            elif done:
                status_p = Paragraph('✓ Selesai', styles['status_done'])
            else:
                status_p = Paragraph('— Belum', styles['status_pending'])

            col_rows.append([
                Paragraph(str(i), styles['cell_label']),
                Paragraph(it['label'], styles['cell_value']),
                status_p,
                Paragraph(pills, styles['cell_label']),
                Paragraph(note[:80], styles['cell_label']),
            ])

        tbl = Table(col_rows, colWidths=[0.7*cm, 6.5*cm, 2.2*cm, 3.8*cm, 4.1*cm], hAlign='LEFT')
        tbl.setStyle(TableStyle([
            ('FONTNAME',    (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTNAME',    (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE',    (0,0), (-1,-1), 7.5),
            ('BACKGROUND',  (0,0), (-1,0), C_TEAL_LIGHT),
            ('TEXTCOLOR',   (0,0), (-1,0), C_TEAL),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [C_WHITE, C_GRAY_LT]),
            ('GRID',        (0,0), (-1,-1), 0.5, C_GRAY_BD),
            ('VALIGN',      (0,0), (-1,-1), 'TOP'),
            ('ALIGN',       (0,0), (0,-1), 'CENTER'),
            ('TOPPADDING',  (0,0), (-1,-1), 3),
            ('BOTTOMPADDING',(0,0),(-1,-1), 3),
            ('LEFTPADDING', (0,0), (-1,-1), 4),
            ('RIGHTPADDING',(0,0), (-1,-1), 4),
        ]))

        sub_header = Table(
            [[Paragraph(header_text, ParagraphStyle('ch', fontName='Helvetica-Bold',
               fontSize=8.5, textColor=C_TEAL, leading=12))]],
            colWidths=[17.3*cm],
        )
        sub_header.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), C_TEAL_PALE),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING',(0,0),(-1,-1), 4),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('GRID', (0,0), (-1,-1), 0.5, C_TEAL_LIGHT),
        ]))

        elements.append(KeepTogether([sub_header, tbl, Spacer(1, 4)]))

    return elements

# ── TABEL LAPORAN BULANAN ─────────────────────────────────────────────────────
def build_laporan_bulanan(styles, reports_list, reports_state):
    elements = []
    elements.append(Paragraph('D. KOMPILASI LAPORAN BULANAN TIM', styles['section_title']))

    done    = sum(1 for r in reports_list if reports_state.get(r['id'], False))
    total   = len(reports_list)
    missing = [r for r in reports_list if not reports_state.get(r['id'], False)]

    summary_p = Paragraph(
        f"Status: <b>{done}/{total}</b> laporan sudah masuk." +
        (f" Belum masuk dari: {', '.join(r['pic'].split(',')[0].strip() for r in missing)}." if missing else " Semua laporan sudah lengkap."),
        ParagraphStyle('rpt_sum', fontName='Helvetica', fontSize=8,
                       textColor=C_RED if missing else C_TEAL_MID, leading=12,
                       spaceAfter=4)
    )
    elements.append(summary_p)

    rpt_data = [['No.', 'Nama Laporan', 'PIC', 'Status']]
    for i, r in enumerate(reports_list, 1):
        is_done = reports_state.get(r['id'], False)
        status_p = Paragraph('✓ Masuk' if is_done else '— Belum Masuk',
                             styles['status_done'] if is_done else styles['status_pending'])
        rpt_data.append([
            Paragraph(str(i), styles['cell_label']),
            Paragraph(r['name'], styles['cell_value']),
            Paragraph(r['pic'], styles['cell_label']),
            status_p,
        ])

    t = Table(rpt_data, colWidths=[0.7*cm, 7*cm, 6.5*cm, 3.1*cm], hAlign='LEFT')
    t.setStyle(TableStyle([
        ('FONTNAME',  (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',  (0,0), (-1,-1), 8),
        ('BACKGROUND',(0,0), (-1,0), C_TEAL_LIGHT),
        ('TEXTCOLOR', (0,0), (-1,0), C_TEAL),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [C_WHITE, C_GRAY_LT]),
        ('GRID',      (0,0), (-1,-1), 0.5, C_GRAY_BD),
        ('VALIGN',    (0,0), (-1,-1), 'TOP'),
        ('ALIGN',     (0,0), (0,-1), 'CENTER'),
        ('TOPPADDING',(0,0), (-1,-1), 3),
        ('BOTTOMPADDING',(0,0),(-1,-1), 3),
        ('LEFTPADDING',(0,0),(-1,-1), 4),
    ]))
    elements.append(t)
    return elements

# ── CATATAN & TANDA TANGAN ────────────────────────────────────────────────────
def build_catatan_dan_ttd(styles, data):
    elements = []
    elements.append(Paragraph('E. CATATAN DAN KETERANGAN TAMBAHAN', styles['section_title']))
    elements.append(Paragraph(
        '(Catatan, temuan, atau keterangan tambahan yang perlu dilaporkan kepada Kepala Stasiun)',
        ParagraphStyle('note_placeholder', fontName='Helvetica', fontSize=8,
                       textColor=C_GRAY_MID, leading=12)
    ))

    # Kotak catatan kosong
    note_box = Table(
        [['']],
        colWidths=[17.3*cm],
        rowHeights=[2.5*cm],
    )
    note_box.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, C_GRAY_BD),
        ('BACKGROUND', (0,0), (-1,-1), C_GRAY_LT),
    ]))
    elements.append(note_box)
    elements.append(Spacer(1, 14))

    # Tanda tangan
    ttd_data = [[
        [Paragraph('Mengetahui,', styles['cell_label']),
         Paragraph('Kepala Stasiun Meteorologi Sangia Nibandera Kolaka', styles['cell_label']),
         Spacer(1, 50),
         Paragraph(KEPALA, styles['cell_bold']),
         Paragraph(f"NIP. {NIP_KEPALA}", styles['cell_label'])],

        [Paragraph(f"Kolaka, {data['print_time']}", styles['cell_label']),
         Paragraph('Ketua Tim Unit Layanan Informasi Meteorologi', styles['cell_label']),
         Spacer(1, 50),
         Paragraph(data['nama'], styles['cell_bold']),
         Paragraph(f"NIP. {data['nip']}", styles['cell_label'])],
    ]]

    ttd_table = Table(ttd_data, colWidths=[8.5*cm, 8.8*cm], hAlign='LEFT')
    ttd_table.setStyle(TableStyle([
        ('VALIGN',    (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING',(0,0), (-1,-1), 0),
        ('BOTTOMPADDING',(0,0),(-1,-1), 0),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING',(0,0), (-1,-1), 0),
    ]))
    elements.append(KeepTogether([ttd_table]))
    return elements

# ── FOOTER ────────────────────────────────────────────────────────────────────
def on_page(canvas, doc, data, styles):
    canvas.saveState()
    page_w, page_h = A4
    footer_text = (
        f"{NAMA_STASIUN} ({ICAO})  ·  "
        f"Laporan Kinerja Ketua Tim LIM  ·  "
        f"{data['period']}  ·  "
        f"Hal. {canvas.getPageNumber()}"
    )
    canvas.setFont('Helvetica', 7)
    canvas.setFillColor(C_GRAY_MID)
    canvas.drawCentredString(page_w / 2, 1.2*cm, footer_text)
    canvas.setStrokeColor(C_GRAY_BD)
    canvas.setLineWidth(0.5)
    canvas.line(2*cm, 1.5*cm, page_w - 2*cm, 1.5*cm)
    canvas.restoreState()

# ── MAIN ──────────────────────────────────────────────────────────────────────
def generate_pdf(json_path, output_path=None):
    print(f"Membaca: {json_path}")
    state_raw, meta = load_json(json_path)
    data = extract_state(state_raw, meta)

    # Tentukan item dan laporan yang digunakan
    items_list   = data['custom_tasks']   if data['custom_tasks']   else DEFAULT_ITEMS
    reports_list = data['custom_reports'] if data['custom_reports'] else DEFAULT_REPORTS
    items_state  = data['items']
    reports_state = data['reports']

    # Nama file output
    if not output_path:
        now = wita_now()
        base = os.path.dirname(json_path)
        fname = f"LIM_Laporan_{now.strftime('%Y%m')}_{now.strftime('%H%M%S')}.pdf"
        output_path = os.path.join(base, fname) if base else fname

    # Build PDF
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=1.8*cm, bottomMargin=2.2*cm,
        title=f"Laporan Kinerja Ketua Tim LIM — {data['period']}",
        author=data['nama'],
        subject="Laporan Kinerja BMKG WAWP",
    )

    styles = build_styles()

    # Coba ambil logo dari JSON jika ada
    logo_data = None
    if 'BMKG_LOGO_B64' in state_raw:
        raw_logo = state_raw['BMKG_LOGO_B64']
        if raw_logo.startswith('data:'):
            raw_logo = raw_logo.split(',', 1)[1]
        logo_data = raw_logo

    story = []
    story += build_kop(styles, logo_data)
    story += build_title(styles, data)
    story += build_identitas(styles, data)
    story.append(Spacer(1, 6))
    story += build_ringkasan(styles, items_list, items_state)
    story.append(Spacer(1, 6))
    story += build_checklist_tables(styles, items_list, items_state)
    story.append(Spacer(1, 6))
    story += build_laporan_bulanan(styles, reports_list, reports_state)
    story.append(Spacer(1, 6))
    story += build_catatan_dan_ttd(styles, data)

    doc.build(
        story,
        onFirstPage=lambda c, d: on_page(c, d, data, styles),
        onLaterPages=lambda c, d: on_page(c, d, data, styles),
    )

    print(f"✅ PDF berhasil dibuat: {output_path}")
    return output_path

def main():
    # Cari file JSON
    if len(sys.argv) >= 2:
        json_path = sys.argv[1]
        if not os.path.exists(json_path):
            print(f"ERROR: File tidak ditemukan: {json_path}")
            sys.exit(1)
    else:
        json_path = find_json_file()
        if not json_path:
            print("ERROR: Tidak ada file JSON yang ditemukan di folder ini.")
            print("Ekspor data dari aplikasi terlebih dahulu, lalu jalankan:")
            print("  python generate_laporan.py LIM_data_YYYY-MM-DD.json")
            sys.exit(1)
        print(f"Menggunakan file terbaru: {json_path}")

    output = sys.argv[2] if len(sys.argv) >= 3 else None
    result = generate_pdf(json_path, output)

    # Coba buka PDF secara otomatis
    try:
        import subprocess, platform
        if platform.system() == 'Windows':
            os.startfile(result)
        elif platform.system() == 'Darwin':
            subprocess.run(['open', result])
        else:
            subprocess.run(['xdg-open', result])
    except Exception:
        pass

if __name__ == '__main__':
    main()
