import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx2pdf import convert
import pypdf
import os

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_callout(doc, title, text, bg_hex="F0F9FF", border_hex="0284C7"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=120, bottom=120, left=160, right=160)

    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/>'
        f'<w:left w:val="single" w:sz="36" w:space="0" w:color="{border_hex}"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15

    r_hdr = p.add_run(f"📌 {title}\n")
    r_hdr.bold = True
    r_hdr.font.name = "Times New Roman"
    r_hdr.font.size = Pt(11)
    r_hdr.font.color.rgb = RGBColor(0x03, 0x69, 0xA1)

    r_text = p.add_run(text)
    r_text.font.name = "Times New Roman"
    r_text.font.size = Pt(10)
    r_text.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)

def style_table_header(row, col_widths=None):
    for idx, cell in enumerate(row.cells):
        set_cell_background(cell, "0F172A")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.bold = True
            r.font.name = "Times New Roman"
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_margins(cell, 60, 60, 80, 80)
        if col_widths and idx < len(col_widths):
            cell.width = Inches(col_widths[idx])

def style_table_row(row, is_even=False, col_widths=None, align_justified=True):
    for idx, cell in enumerate(row.cells):
        set_cell_background(cell, "F8FAFC" if is_even else "FFFFFF")
        for p in cell.paragraphs:
            if align_justified and idx > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                r.font.name = "Times New Roman"
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
        set_cell_margins(cell, 45, 45, 70, 70)
        if col_widths and idx < len(col_widths):
            cell.width = Inches(col_widths[idx])

def add_p_justified(doc, text, bold_prefix="", space_after=5, space_before=0, line_spacing=1.15):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing

    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.bold = True
        r_pre.font.name = "Times New Roman"
        r_pre.font.size = Pt(11)
        r_pre.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    r_body = p.add_run(text)
    r_body.font.name = "Times New Roman"
    r_body.font.size = Pt(11)
    r_body.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
    return p

def add_heading_1(doc, text):
    h = doc.add_heading(level=1)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(8)
    h.paragraph_format.keep_with_next = True
    r = h.add_run(text)
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
    return h

def add_heading_2(doc, text):
    h = doc.add_heading(level=2)
    h.paragraph_format.space_before = Pt(10)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True
    r = h.add_run(text)
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    return h

def add_heading_3(doc, text):
    h = doc.add_heading(level=3)
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(4)
    h.paragraph_format.keep_with_next = True
    r = h.add_run(text)
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
    return h

def add_image_with_caption(doc, img_path, caption_text, width=Inches(6.2)):
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(6)
        p_img.paragraph_format.space_after = Pt(2)
        r_img = p_img.add_run()
        r_img.add_picture(img_path, width=width)

        p_cap = doc.add_paragraph()
        p_cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(10)
        r_cap = p_cap.add_run(caption_text)
        r_cap.font.name = "Times New Roman"
        r_cap.font.size = Pt(9.5)
        r_cap.font.italic = True
        r_cap.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

def add_toc_line(doc, title, page_str, level=1, tab_pos=9360):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2.5)
    p.paragraph_format.line_spacing = 1.15
    if level == 2:
        p.paragraph_format.left_indent = Inches(0.2)
    elif level == 3:
        p.paragraph_format.left_indent = Inches(0.4)

    tab_xml = (
        f'<w:pPr {nsdecls("w")}>'
        f'  <w:tabs>'
        f'    <w:tab w:val="right" w:leader="dot" w:pos="{tab_pos}"/>'
        f'  </w:tabs>'
        f'</w:pPr>'
    )
    p._p.append(parse_xml(tab_xml))

    r_title = p.add_run(title)
    r_title.font.name = "Times New Roman"
    if level == 1:
        r_title.bold = True
        r_title.font.size = Pt(10)
        r_title.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
    else:
        r_title.font.size = Pt(9.5)
        r_title.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    r_tab = p.add_run("\t")
    r_tab.font.name = "Times New Roman"

    r_page = p.add_run(str(page_str))
    r_page.font.name = "Times New Roman"
    if level == 1:
        r_page.bold = True
        r_page.font.size = Pt(10)
        r_page.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
    else:
        r_page.font.size = Pt(9.5)
        r_page.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

def build_full_report_document(docx_path, p=None):
    doc = docx.Document()

    # 1 inch margins all around
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
        s.different_first_page_header_footer = True

        # Footer for subsequent pages: Centered "Halaman <PAGE> dari <NUMPAGES>"
        footer = s.footer
        p_foot = footer.paragraphs[0]
        fld_xml = (
            f'<w:p {nsdecls("w")}>'
            f'  <w:pPr><w:jc w:val="center"/></w:pPr>'
            f'  <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="19"/><w:color w:val="64748B"/></w:rPr><w:t xml:space="preserve">Halaman </w:t></w:r>'
            f'  <w:fldSimple w:instr="PAGE"><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="19"/><w:color w:val="64748B"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>'
            f'  <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="19"/><w:color w:val="64748B"/></w:rPr><w:t xml:space="preserve"> dari </w:t></w:r>'
            f'  <w:fldSimple w:instr="NUMPAGES"><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="19"/><w:color w:val="64748B"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>'
            f'</w:p>'
        )
        p_elem = parse_xml(fld_xml)
        p_foot._p.getparent().replace(p_foot._p, p_elem)

    # Global Style: Times New Roman
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(11)
    style_normal.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(5)

    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>')
    style_normal._element.rPr.append(rFonts)

    # Default fallback page dictionary
    if p is None:
        p = {
            "lp": "2", "kp": "3", "di": "4",
            "b1": "5", "b1_1": "5", "b1_2": "5", "b1_3": "5", "b1_4": "6", "b1_5": "6", "b1_6": "7", "b1_7": "8", "b1_8": "8",
            "b2": "10", "b2_1": "10", "b2_2": "10", "b2_2_1": "10", "b2_2_2": "12", "b2_2_3": "15", "b2_2_4": "15",
            "b2_3": "16", "b2_4": "18",
            "b3": "23", "b3_1": "23", "b3_2": "23", "b3_3": "26", "b3_4": "27", "b3_5": "28",
            "b4": "29", "b4_1": "29", "b4_2": "29", "b4_3": "29",
            "dp": "30"
        }

    # =========================================================================
    # 1. HALAMAN JUDUL / COVER RESMI
    # =========================================================================
    p_inst = doc.add_paragraph()
    p_inst.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(4)
    p_inst.paragraph_format.space_after = Pt(2)
    r_inst = p_inst.add_run("PROGRAM STUDI S1 TEKNIK INFORMATIKA\nSTMIK IKMI CIREBON")
    r_inst.bold = True
    r_inst.font.name = "Times New Roman"
    r_inst.font.size = Pt(14)
    r_inst.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_mk = doc.add_paragraph()
    p_mk.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_mk.paragraph_format.space_after = Pt(12)
    r_mk = p_mk.add_run("MATA KULIAH: REKAYASA PERANGKAT LUNAK (TIKB2503021) — SEMESTER 3 / 4 SKS")
    r_mk.font.name = "Times New Roman"
    r_mk.font.size = Pt(10)
    r_mk.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

    # Official School Logo
    logo_path = r"c:\xampp\htdocs\absen\public\assets\images\logo_smk.png"
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_after = Pt(14)
        r_lg = p_logo.add_run()
        r_lg.add_picture(logo_path, width=Inches(1.85))

    p_t = doc.add_paragraph()
    p_t.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_after = Pt(8)
    r_t = p_t.add_run("BERKAS DOKUMEN LENGKAP (PDF TUNGGAL)\nLAPORAN TUGAS PROYEK AKHIR MANDIRI & RUBRIK PENILAIAN OBE")
    r_t.bold = True
    r_t.font.name = "Times New Roman"
    r_t.font.size = Pt(14)
    r_t.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run(
        "RANCANG BANGUN SISTEM PRESENSI DIGITAL BERBASIS QR CODE SCANNER, "
        "NOTIFIKASI REAL-TIME TELEGRAM, SUPERVISI KBM DAN ASISTEN CERDAS AI HADIR.IO "
        "DI SMK NU WARGABINANGUN\n"
        "(GABUNGAN PRD/SRS IEEE 830, SDD UML & ERD 20 TABEL, TESTING & CI/CD DEVOPS)"
    )
    r_sub.bold = True
    r_sub.font.name = "Times New Roman"
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = RGBColor(0x02, 0x84, 0xC7)

    # Metadata Table
    tbl_cov = doc.add_table(rows=7, cols=2)
    tbl_cov.alignment = WD_TABLE_ALIGNMENT.CENTER
    cov_data = [
        ("Nama Mahasiswa / Pengembang", "Winarti"),
        ("NIM (Nomor Induk Mahasiswa)", "41257168"),
        ("Kelas / Angkatan", "TI-1C (Angkatan 2025)"),
        ("Program Studi", "S1 Teknik Informatika"),
        ("Nama Kelompok", "Sunyaragi Nexus"),
        ("Tema Topik Kelompok", "E-Learning & Smart Education AI"),
        ("Dosen Pengampu", "Yudhistira Arie Wijaya, M.Kom., CDS.")
    ]
    for i, (k, v) in enumerate(cov_data):
        c0 = tbl_cov.cell(i, 0)
        c1 = tbl_cov.cell(i, 1)
        c0.text = k; c1.text = v
        c0.paragraphs[0].runs[0].bold = True
        c0.paragraphs[0].runs[0].font.name = "Times New Roman"
        c0.paragraphs[0].runs[0].font.size = Pt(9.5)
        c1.paragraphs[0].runs[0].font.name = "Times New Roman"
        c1.paragraphs[0].runs[0].font.size = Pt(9.5)
        set_cell_background(c0, "F1F5F9")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, 40, 40, 75, 75)
        set_cell_margins(c1, 40, 40, 75, 75)

    p_foot_cov = doc.add_paragraph()
    p_foot_cov.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot_cov.paragraph_format.space_before = Pt(20)
    p_foot_cov.paragraph_format.space_after = Pt(0)
    r_foot_cov = p_foot_cov.add_run("STMIK IKMI CIREBON — TAHUN AKADEMIK 2026/2027")
    r_foot_cov.bold = True
    r_foot_cov.font.name = "Times New Roman"
    r_foot_cov.font.size = Pt(11)
    r_foot_cov.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    doc.add_page_break()

    # =========================================================================
    # 2. LEMBAR PENGESAHAN & PERNYATAAN KEASLIAN
    # =========================================================================
    add_heading_1(doc, "LEMBAR PERNYATAAN KEASLIAN & INTEGRITAS MANDIRI")

    add_p_justified(
        doc,
        "Sesuai dengan ketentuan resmi Panduan Tugas Proyek Akhir Mandiri Mata Kuliah Rekayasa Perangkat Lunak "
        "(TIKB2503021) Program Studi S1 Teknik Informatika STMIK IKMI Cirebon yang secara tegas mensyaratkan "
        "100% Individual Ownership & Full-App Build, saya yang bertanda tangan di bawah ini:"
    )

    tbl_bio = doc.add_table(rows=4, cols=2)
    tbl_bio.alignment = WD_TABLE_ALIGNMENT.CENTER
    b_data = [
        ("Nama Lengkap", ": Winarti"),
        ("NIM", ": 41257168"),
        ("Kelas / Program Studi", ": TI-1C / S1 Teknik Informatika"),
        ("Judul Proyek Mandiri", ": Sistem Presensi Digital SMK NU WARGABINANGUN by Winarti")
    ]
    for i, (k, v) in enumerate(b_data):
        c0 = tbl_bio.cell(i, 0); c1 = tbl_bio.cell(i, 1)
        c0.text = k; c1.text = v
        c0.paragraphs[0].runs[0].bold = True
        c0.paragraphs[0].runs[0].font.name = "Times New Roman"
        c0.paragraphs[0].runs[0].font.size = Pt(9.5)
        c1.paragraphs[0].runs[0].font.name = "Times New Roman"
        c1.paragraphs[0].runs[0].font.size = Pt(9.5)
        set_cell_background(c0, "F8FAFC")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, 40, 40, 70, 70)
        set_cell_margins(c1, 40, 40, 70, 70)

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(6)

    add_p_justified(
        doc,
        "Menyatakan dengan penuh tanggung jawab dan integritas akademik bahwa:\n"
        "1. Proyek perangkat lunak ini dikerjakan secara utuh mandiri (end-to-end full build) mulai dari penyusunan "
        "spesifikasi kebutuhan (PRD/SRS IEEE 830), pemodelan 4 diagram UML standar dan ERD 20 tabel, perancangan antarmuka UI/UX di Figma, "
        "penulisan kode sumber aplikasi web PHP 8.2 MVC berorientasi objek, integrasi webhook Bot Telegram (@HadirioBot), "
        "integrasi REST API kecerdasan artifisial NVIDIA NIM (Nemotron LLM), penulisan skrip unit testing otomatis, "
        "hingga konfigurasi otomatisasi CI/CD Pipeline pada GitHub Actions.\n"
        "2. Meskipun tergabung dalam Kelompok Sunyaragi Nexus dengan Tema Topik 'E-Learning & Smart Education AI', "
        "tidak ada pembagian modul fungsional antar-anggota kelompok. Seluruh fungsionalitas aplikasi dibangun, diuji, "
        "dan didokumentasikan secara independen oleh penulis sendiri sebagai perwujudan tanggung jawab Full-Stack Software Engineer.\n"
        "3. Laporan, kode sumber, dan dokumentasi ini terbebas dari plagiarisme. Seluruh referensi pustaka dan pustaka pihak "
        "ketiga telah dicantumkan secara transparan mengikuti kaidah sitasi ilmiah IEEE."
    )

    tbl_ttd = doc.add_table(rows=1, cols=2)
    tbl_ttd.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_mhs = tbl_ttd.cell(0, 0)
    c_dos = tbl_ttd.cell(0, 1)

    p_m = c_mhs.paragraphs[0]
    p_m.paragraph_format.line_spacing = 1.15
    p_m.add_run("Mengetahui dan Menyetujui,\nYang Membuat Pernyataan,\n\n\n\n").font.name = "Times New Roman"
    r_wm = p_m.add_run("WINARTI\n")
    r_wm.bold = True
    r_wm.font.name = "Times New Roman"
    p_m.add_run("NIM: 41257168").font.name = "Times New Roman"

    p_d = c_dos.paragraphs[0]
    p_d.paragraph_format.line_spacing = 1.15
    p_d.add_run("Cirebon, 15 November 2026\nDosen Pengampu RPL,\n\n\n\n").font.name = "Times New Roman"
    r_wd = p_d.add_run("Yudhistira Arie Wijaya, M.Kom., CDS.\n")
    r_wd.bold = True
    r_wd.font.name = "Times New Roman"
    p_d.add_run("NIDN / Dosen Pengampu").font.name = "Times New Roman"

    set_cell_margins(c_mhs, 50, 50, 70, 70)
    set_cell_margins(c_dos, 50, 50, 70, 70)

    doc.add_page_break()

    # =========================================================================
    # 3. KATA PENGANTAR & RINGKASAN EKSEKUTIF
    # =========================================================================
    add_heading_1(doc, "KATA PENGANTAR")

    add_p_justified(
        doc,
        "Puji dan syukur penulis panjatkan ke hadirat Allah Subhanahu Wa Ta'ala atas limpahan rahmat, petunjuk, "
        "dan kemudahan-Nya, sehingga Dokumen Lengkap Proyek Akhir Mandiri Rekayasa Perangkat Lunak (TIKB2503021) "
        "yang berjudul 'Rancang Bangun Sistem Presensi Digital Berbasis QR Code Scanner, Notifikasi Real-Time Telegram, "
        "Supervisi KBM dan Asisten Cerdas AI Hadir.io di SMK NU WARGABINANGUN' ini dapat terselesaikan dengan paripurna."
    )
    add_p_justified(
        doc,
        "Laporan terpadu ini disusun sebagai bukti pemenuhan Capaian Pembelajaran Lulusan (CPL) berbasis kurikulum "
        "Outcome-Based Education (OBE) pada Semester 3 Program Studi S1 Teknik Informatika STMIK IKMI Cirebon. "
        "Penulis menyampaikan apresiasi dan rasa terima kasih sebesar-besarnya kepada:\n"
        "1. Bapak Yudhistira Arie Wijaya, M.Kom., CDS., selaku Dosen Pengampu Mata Kuliah Rekayasa Perangkat Lunak, "
        "yang telah memberikan arahan teknis yang mendalam, standar baku industri rekayasa perangkat lunak modern, "
        "serta motivasi tiada henti dalam penyelesaian karya rekayasa mandiri ini.\n"
        "2. Bapak/Ibu Pimpinan, Guru, dan Staf Tata Usaha SMK NU WARGABINANGUN, yang telah berkenan memberikan "
        "izin penelitian, akses observasi lapangan, serta data studi kasus yang sangat berharga.\n"
        "3. Orang tua tercinta dan rekan-rekan mahasiswa Kelas TI-1C Angkatan 2025 serta Kelompok Sunyaragi Nexus "
        "yang senantiasa mengalirkan doa, inspirasi, dan kebersamaan intelektual."
    )

    add_callout(
        doc,
        "RINGKASAN EKSEKUTIF (EXECUTIVE SUMMARY)",
        "Sistem Presensi Digital SMK NU WARGABINANGUN merepresentasikan transformasi digital end-to-end yang "
        "menggantikan metode presensi kertas konvensional dengan sistem cerdas berbasis web. Aplikasi ini mengintegrasikan "
        "pemindai kamera HTML5-QRCode berkecepatan tinggi (< 0.5 detik), webhook notifikasi Telegram real-time (@HadirioBot), "
        "mekanisme sinkronisasi ganda (Dual-Sync) bagi dewan wali kelas, otomatisasi persetujuan izin (Auto-Sync), supervisi KBM "
        "dengan pelacakan status baca, dan asisten AI Hadir.io berbasis model bahasa NVIDIA NIM Nemotron yang mendukung 4 bahasa "
        "(Indonesia, Inggris, Sunda, Jawa) secara aman berpagar RBAC. Seluruh arsitektur dibangun di atas 20 tabel database 3NF "
        "dan divalidasi dengan unit test 100% PASS serta CI/CD pipeline GitHub Actions yang berstatus hijau."
    )

    doc.add_page_break()

    # =========================================================================
    # 4. DAFTAR ISI (OTOMATIS & TABS WITH DOT LEADERS)
    # =========================================================================
    add_heading_1(doc, "DAFTAR ISI")

    add_toc_line(doc, "LEMBAR PERNYATAAN KEASLIAN & INTEGRITAS MANDIRI", p["lp"], level=1)
    add_toc_line(doc, "KATA PENGANTAR & RINGKASAN EKSEKUTIF", p["kp"], level=1)
    add_toc_line(doc, "DAFTAR ISI", p["di"], level=1)
    add_toc_line(doc, "BAB I: PRODUCT REQUIREMENTS DOCUMENT (PRD) & SRS", p["b1"], level=1)
    add_toc_line(doc, "1.1 Konsep Proyek & Pernyataan Pengerjaan Mandiri", p["b1_1"], level=2)
    add_toc_line(doc, "1.2 Latar Belakang Masalah & Konteks Lapangan", p["b1_2"], level=2)
    add_toc_line(doc, "1.3 Identifikasi Masalah Nyata & Rumusan Solusi Digital", p["b1_3"], level=2)
    add_toc_line(doc, "1.4 Visi Produk & Target Keberhasilan Terukur", p["b1_4"], level=2)
    add_toc_line(doc, "1.5 Analisis Kebutuhan Pengguna (User Stories & Acceptance Criteria)", p["b1_5"], level=2)
    add_toc_line(doc, "1.6 Spesifikasi Kebutuhan Fungsional (FR-01 s/d FR-14)", p["b1_6"], level=2)
    add_toc_line(doc, "1.7 Spesifikasi Kebutuhan Non-Fungsional (NFR-01 s/d NFR-06)", p["b1_7"], level=2)
    add_toc_line(doc, "1.8 Arsitektur & Spesifikasi Fitur Kecerdasan Artifisial (AI Hadir.io)", p["b1_8"], level=2)

    add_toc_line(doc, "BAB II: SOFTWARE DESIGN DOCUMENT (SDD) & PEMODELAN UML", p["b2"], level=1)
    add_toc_line(doc, "2.1 Arsitektur Perangkat Lunak Terpadu (MVC Pattern)", p["b2_1"], level=2)
    add_toc_line(doc, "2.2 Pemodelan 4 Diagram UML Standar Utama", p["b2_2"], level=2)
    add_toc_line(doc, "2.2.1 Use Case Diagram Terpadu & Skenario Use Case Lengkap", p["b2_2_1"], level=3)
    add_toc_line(doc, "2.2.2 Activity Diagram Siklus Bisnis & Alur AI (5 Swimlanes)", p["b2_2_2"], level=3)
    add_toc_line(doc, "2.2.3 Class Diagram Berorientasi Objek & Kepatuhan Prinsip SOLID", p["b2_2_3"], level=3)
    add_toc_line(doc, "2.2.4 Sequence Diagram Alur Pertukaran Pesan Real-Time", p["b2_2_4"], level=3)
    add_toc_line(doc, "2.3 Perancangan Basis Data & Entity Relationship Diagram (20 Tabel)", p["b2_3"], level=2)
    add_toc_line(doc, "2.4 Perancangan Antarmuka Pengguna & Prototipe Interaktif Figma", p["b2_4"], level=2)

    add_toc_line(doc, "BAB III: IMPLEMENTASI KODE, TESTING & CI/CD DEVOPS", p["b3"], level=1)
    add_toc_line(doc, "3.1 Struktur Direktori Proyek Bersih (MVC Architecture)", p["b3_1"], level=2)
    add_toc_line(doc, "3.2 Implementasi Fitur Kunci & Potongan Kode (Code Highlights)", p["b3_2"], level=2)
    add_toc_line(doc, "3.3 Laporan Pengujian Perangkat Lunak (Software Testing Report)", p["b3_3"], level=2)
    add_toc_line(doc, "3.4 Otomasi CI/CD Pipeline & GitHub Actions", p["b3_4"], level=2)
    add_toc_line(doc, "3.5 Petunjuk Instalasi & Deployment Sistem (Setup Instructions)", p["b3_5"], level=2)

    add_toc_line(doc, "BAB IV: KESIMPULAN, REFLEKSI REKAYASA & MATRIKS DELIVERABLES", p["b4"], level=1)
    add_toc_line(doc, "4.1 Kesimpulan Proyek", p["b4_1"], level=2)
    add_toc_line(doc, "4.2 Refleksi Pembelajaran Mandiri (Engineering Reflection)", p["b4_2"], level=2)
    add_toc_line(doc, "4.3 Matriks Checklist Pengumpulan Berkas (Submission Deliverables)", p["b4_3"], level=2)

    add_toc_line(doc, "DAFTAR PUSTAKA", p["dp"], level=1)

    doc.add_page_break()

    # =========================================================================
    # BAB I: PRODUCT REQUIREMENTS DOCUMENT (PRD) & SRS
    # =========================================================================
    add_heading_1(doc, "BAB I: PRODUCT REQUIREMENTS DOCUMENT (PRD) & SRS")

    add_heading_2(doc, "1.1 Konsep Proyek & Pernyataan Pengerjaan Mandiri")
    add_p_justified(
        doc,
        "Proyek akhir mata kuliah Rekayasa Perangkat Lunak (TIKB2503021) ini mengadopsi konsep Group Topic Theme with "
        "Individual Full-App Build. Penulis tergabung dalam Kelompok Sunyaragi Nexus yang menyepakati tema kelompok "
        "'E-Learning & Smart Education AI'. Meskipun bernaung dalam kelompok tema yang sama, sistem presensi ini dikerjakan "
        "100% secara mandiri dan utuh dari hulu ke hilir (full end-to-end software engineering) oleh penulis sendiri. "
        "Tidak terdapat pembagian modul fungsional; setiap alur—mulai dari perumusan spesifikasi kebutuhan (PRD/SRS), "
        "pemodelan UML, perancangan basis data 20 tabel, implementasi kode MVC PHP 8.2, integrasi Bot Telegram, "
        "integrasi AI NVIDIA NIM, unit testing, hingga CI/CD pipeline—dibangun secara independen sebagai bukti keahlian "
        "Full-Stack Software Engineer."
    )

    add_heading_2(doc, "1.2 Latar Belakang Masalah & Konteks Lapangan")
    add_p_justified(
        doc,
        "SMK NU WARGABINANGUN merupakan institusi pendidikan kejuruan yang membina ratusan peserta didik pada berbagai program "
        "keahlian teknik (Teknik Komputer & Jaringan, Teknik Kendaraan Ringan, dan Teknik Sepeda Motor). Dalam ekosistem sekolah kejuruan, "
        "kedisiplinan kehadiran di gerbang sekolah dan di ruang kelas/bengkel praktik merupakan indikator fundamental pembentukan "
        "karakter calon tenaga kerja profesional. Namun, tata kelola kehadiran yang berjalan selama ini masih sepenuhnya mengandalkan "
        "buku absensi kertas konvensional dan lembaran paraf guru piket. Praktik manual ini telah menimbulkan berbagai kendala "
        "operasional yang menghambat efektivitas manajemen sekolah modern."
    )

    add_heading_2(doc, "1.3 Identifikasi Masalah Nyata & Rumusan Solusi Digital")
    add_p_justified(
        doc,
        "Berdasarkan analisis kebutuhan empiris di lapangan, ditemukan tiga kelompok permasalahan utama:\n"
        "1. Kerentanan Fisik & Beban Rekapitulasi Manual: Lembaran kertas absen rentan sobek, basah, dan hilang. "
        "Setiap akhir bulan, staf Tata Usaha (TU) dan wali kelas harus meluangkan waktu berhari-hari untuk merekapitulasi "
        "ribuan baris paraf kehadiran secara manual, sehingga rentan terjadi kesalahan hitung (human error).\n"
        "2. Ketiadaan Transparansi & Notifikasi Real-Time: Orang tua tidak mendapatkan informasi seketika saat anaknya tiba di "
        "sekolah atau membolos. Kepala Sekolah dan Wali Kelas juga tidak memiliki dasbor real-time untuk memantau rekapitulasi harian "
        "dan kedisiplinan guru di ruang kelas.\n"
        "3. Kerawanan Titip Absen & Birokrasi Perizinan yang Lemah: Absensi paraf manual sangat mudah dimanipulasi atau dititipkan "
        "kepada rekan sejawat. Pengajuan izin sakit pun masih menggunakan secarik surat fisik yang tidak terarsip dengan baik, "
        "serta belum adanya kanal evaluasi KBM guru dan pembinaan siswa secara terdokumentasi rapi."
    )
    add_p_justified(
        doc,
        "Untuk menjawab seluruh tantangan tersebut, dirancanglah Sistem Presensi Digital SMK NU WARGABINANGUN yang memadukan "
        "pemindaian barcode/QR digital instan, pengiriman notifikasi otomatis Telegram ke ponsel orang tua, sinkronisasi kehadiran "
        "ganda (Dual-Sync) bagi wali kelas, otomatisasi persetujuan izin (Auto-Sync), supervisi KBM berbasis status baca, serta "
        "asisten cerdas Hadir.io yang ditenagai model kecerdasan artifisial NVIDIA NIM Nemotron LLM."
    )

    add_heading_2(doc, "1.4 Visi Produk & Target Keberhasilan Terukur")
    add_p_justified(
        doc,
        "Visi Produk: 'Mewujudkan ekosistem presensi sekolah kejuruan yang modern, akuntabel, transparan, bebas kertas "
        "(zero paper waste), dan berdaya guna kecerdasan artifisial terdepan di SMK NU WARGABINANGUN.'\n\n"
        "Target Terukur Keberhasilan Sistem:\n"
        "• 100% Digitalisasi seluruh transaksi absensi harian siswa, guru, dan wali kelas.\n"
        "• Pengurangan 100% penggunaan kertas absensi dan berkas perizinan fisik (Zero Paper Waste).\n"
        "• Kecepatan pemindaian kamera barcode dan verifikasi jam server di bawah 0.5 detik.\n"
        "• Pengiriman notifikasi kehadiran ke aplikasi Telegram orang tua dalam waktu kurang dari 2 detik.\n"
        "• Eliminasi total potensi titip absen melalui enkripsi kode kartu identitas unik dan validasi sesi ketat."
    )

    add_heading_2(doc, "1.5 Analisis Kebutuhan Pengguna (User Stories & Acceptance Criteria)")
    add_p_justified(
        doc,
        "Kebutuhan sistem dirumuskan dari perspektif 5 aktor pengguna utama menggunakan format standar User Story "
        "dan Kriteria Penerimaan (Acceptance Criteria):"
    )

    # Tabel 1.1
    tbl_us = doc.add_table(rows=6, cols=3)
    tbl_us.alignment = WD_TABLE_ALIGNMENT.CENTER
    us_data = [
        ("Peran (Role)", "Pernyataan Kebutuhan (User Story)", "Kriteria Penerimaan (Acceptance Criteria)"),
        ("Administrator", "Sebagai Admin, saya ingin mengelola master data akademik (rombel, guru, mapel, jadwal), memverifikasi permohonan izin, dan mencetak kartu barcode.", "Data tersimpan di 20 tabel database secara relasional, kartu barcode siap cetak beresolusi tinggi, dan persetujuan izin otomatis menyinkronkan status absensi harian."),
        ("Kepala Sekolah", "Sebagai Kepsek, saya ingin memantau statistik kehadiran sekolah secara menyeluruh dan memberikan catatan evaluasi KBM tertulis kepada guru.", "Dashboard menampilkan persentase kehadiran harian, dan modul catatan guru memiliki tracking otomatis yang menunjukkan kapan guru membaca arahan tersebut."),
        ("Wali Kelas", "Sebagai Wali Kelas, saya ingin memantau absensi kelas binaan, mencatat pembinaan kedisiplinan siswa, dan melakukan presensi kehadiran mandiri.", "Satu kali pemindaian kartu langsung mencatat presensi di data guru dan wali kelas (Dual-Sync), serta catatan pembinaan siswa dapat dikonfirmasi baca oleh siswa."),
        ("Guru Pengajar", "Sebagai Guru, saya ingin melakukan presensi masuk/pulang mandiri, memantau jadwal KBM mingguan, dan mengisi jurnal kegiatan mengajar.", "Pemindaian barcode tervalidasi jam server secara akurat, notifikasi Telegram terkirim instan, dan jurnal mengajar tersimpan rapi untuk supervisi kepsek."),
        ("Siswa", "Sebagai Siswa, saya ingin memindai kartu barcode di gerbang sekolah, melihat riwayat kehadiran pribadi, dan mengajukan izin/sakit secara mandiri.", "Hasil pemindaian muncul seketika di layar kiosk, notifikasi terkirim ke ponsel orang tua, dan pengajuan izin wajib melampirkan foto surat dokter yang valid.")
    ]
    for i, r_vals in enumerate(us_data):
        for j, val in enumerate(r_vals):
            c = tbl_us.cell(i, j)
            c.text = val
    style_table_header(tbl_us.rows[0], [1.3, 2.6, 2.6])
    for idx, row in enumerate(tbl_us.rows[1:], start=1):
        style_table_row(row, is_even=(idx % 2 == 1), col_widths=[1.3, 2.6, 2.6])

    p_cap_t11 = doc.add_paragraph("Tabel 1.1 Analisis Kebutuhan Pengguna (User Stories & Acceptance Criteria)")
    p_cap_t11.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_t11.paragraph_format.space_before = Pt(3)
    p_cap_t11.paragraph_format.space_after = Pt(8)
    p_cap_t11.runs[0].font.name = "Times New Roman"
    p_cap_t11.runs[0].font.size = Pt(9.5)
    p_cap_t11.runs[0].font.italic = True

    add_heading_2(doc, "1.6 Spesifikasi Kebutuhan Fungsional (FR-01 s/d FR-14)")
    add_p_justified(
        doc,
        "Berdasarkan standar IEEE Std 830-1998, kebutuhan fungsional sistem presensi dijabarkan ke dalam 14 butir spesifikasi:"
    )

    fr_items = [
        ("FR-01: Autentikasi Multi-Role & Pengelolaan Sesi Aman", "Sistem memvalidasi login menggunakan hashing Bcrypt (cost 10) dan membatasi akses menu berdasarkan 5 role pengguna pada tabel users dan roles."),
        ("FR-02: Manajemen Master Data Akademik Terpadu", "Administrator memiliki kewenangan penuh melakukan operasi CRUD pada tabel tahun_ajaran, jurusan, kelas, guru, wali_kelas_data, wali_kelas, siswa, mata_pelajaran, dan jadwal_mengajar."),
        ("FR-03: Generator Barcode & Kartu Identitas Digital", "Sistem secara otomatis menggenerasi string barcode dan QR code unik terenkripsi untuk setiap siswa dan guru, tersimpan pada tabel barcode_guru dan rfid_cards."),
        ("FR-04: Pemindai Kamera HTML5-QRCode Cepat", "Sistem menyediakan antarmuka pemindai kamera responsif yang membaca kartu QR secara real-time melalui AJAX tanpa perlu memuat ulang peramban."),
        ("FR-05: Validasi Waktu Masuk & Deteksi Keterlambatan", "Pemindaian pukul 05:00–07:00 WIB dicatat 'Hadir Tepat Waktu'. Pemindaian di atas pukul 07:00 WIB otomatis dicatat 'Terlambat' disertai log menit keterlambatan."),
        ("FR-06: Validasi Waktu Kepulangan", "Pemindaian kepulangan hanya dibuka mulai pukul 15:00 WIB ke atas. Pemindaian sebelum jam 15:00 WIB ditolak oleh sistem untuk mencegah kepulangan dini."),
        ("FR-07: Mekanisme Sinkronisasi Ganda (Dual-Sync)", "Saat Wali Kelas memindai kartu presensi, Controller secara atomik mencatat kehadiran pada tabel absensi_wali_kelas dan absensi_guru secara simultan."),
        ("FR-08: Integrasi Notifikasi Real-Time Telegram", "Setiap kali transaksi presensi berhasil dicatat, sistem memicu webhook cURL ke Bot Telegram (@HadirioBot) untuk mengirim notifikasi live ke kanal sekolah atau orang tua."),
        ("FR-09: Pengajuan Dispensasi / Izin & Sakit Online", "Siswa dan Guru dapat mengajukan izin sakit mandiri dari portalnya dengan wajib mengunggah file foto surat dokter/keterangan orang tua (JPG/PDF, maks 2MB)."),
        ("FR-10: Sinkronisasi Otomatis Approval Izin (Auto-Sync)", "Ketika Admin menyetujui permohonan izin pada pengajuan_izin, sistem otomatis menyinkronkan status absensi siswa/guru pada tabel absensi menjadi 'Izin' atau 'Sakit'."),
        ("FR-11: Supervisi KBM & Jurnal Mengajar Guru", "Kepala Sekolah dapat memberikan catatan evaluasi atau motivasi tertulis pada tabel catatan_guru dengan pelacakan status baca otomatis oleh guru terkait."),
        ("FR-12: Pembinaan Karakter & Kedisiplinan Siswa", "Wali Kelas dapat mencatat bimbingan atau teguran pada catatan_pembinaan, dan siswa memiliki tombol interaktif 'Tandai Sudah Dibaca' sebagai bukti komitmen moral."),
        ("FR-13: Rekapitulasi Presensi Efektif & Ekspor Laporan", "Sistem menghitung persentase kehadiran dengan filter ketat hari efektif Senin–Jumat (Sabtu/Minggu otomatis dikecualikan), lalu mengekspor laporan resmi ke Excel (.xlsx) dan PDF."),
        ("FR-14: Asisten Cerdas Sekolah Hadir.io", "Widget chat AI berbasis NVIDIA NIM Nemotron LLM menjawab pertanyaan data sekolah secara akurat, ringkas, santun, dan mendukung 4 bahasa (Indonesia, Inggris, Sunda, Jawa).")
    ]
    for code, desc in fr_items:
        add_p_justified(doc, desc, bold_prefix=code + ":\n")

    add_heading_2(doc, "1.7 Spesifikasi Kebutuhan Non-Fungsional (NFR-01 s/d NFR-06)")
    add_p_justified(
        doc,
        "Untuk menjamin kualitas dan keandalan perangkat lunak di lingkungan produksi, ditetapkan 6 kebutuhan non-fungsional:"
    )

    nfr_items = [
        ("NFR-01: Performance & Response Time", "Waktu eksekusi pemindaian barcode hingga data tersimpan di server dan menampilkan umpan balik di layar harus di bawah 0.5 detik pada jaringan lokal sekolah."),
        ("NFR-02: Security & Protection", "Seluruh input data disanitasi menggunakan PDO Prepared Statements untuk mencegah SQL Injection, XSS Protection pada input form, dan kata sandi dienkripsi dengan Bcrypt."),
        ("NFR-03: RBAC Isolation & Privacy", "Sistem menjamin isolasi data total; Siswa dilarang mengakses data guru/keuangan, dan Guru tidak dapat memanipulasi data kelas lain."),
        ("NFR-04: Reliability & Data Integrity", "Menerapkan transaksi atomik (commit/rollback) pada proses Dual-Sync wali kelas dan Auto-Sync persetujuan izin untuk mengeliminasi potensi data corrupt."),
        ("NFR-05: High Availability", "Aplikasi siap beroperasi 24/7 dengan arsitektur web lokal XAMPP yang tahan terhadap kendala pemutusan koneksi internet luar saat mode scan lokal."),
        ("NFR-06: Usability & Cross-Platform", "Antarmuka responsif berbasis Bootstrap 5 mendukung tampilan desktop, tablet, dan smartphone dengan visual yang nyaman dan intuitif.")
    ]
    for code, desc in nfr_items:
        add_p_justified(doc, desc, bold_prefix=code + ": ")

    add_heading_2(doc, "1.8 Arsitektur & Spesifikasi Fitur Kecerdasan Artifisial (AI Hadir.io)")
    add_p_justified(
        doc,
        "Fitur AI Hadir.io dibangun dengan konsep Dual-Engine yang menggabungkan kecepatan ekstraksi data terstruktur "
        "dengan kemampuan penalaran bahasa alami mutakhir:\n"
        "• Engine 1: Fast-Path SQL Extraction. Untuk pertanyaan faktual menyangkut data sekolah—seperti jumlah rombongan belajar, "
        "total guru aktif, daftar mata pelajaran, maupun jadwal KBM—controller mengekstraksi data secara langsung melalui kueri SQL "
        "ke database lokal `db_absensi_smk`. Hal ini menjamin akurasi informasi sebesar 100% dan bebas dari risiko halusinasi AI.\n"
        "• Engine 2: NVIDIA NIM LLM Processing. Ringkasan data faktual tersebut kemudian dikemas ke dalam prompt berpagar ketat "
        "(guarded context) dan diteruskan ke model bahasa besar NVIDIA NIM (Nemotron-3.5-bolt-e4b) melalui REST API. "
        "Model menghasilkan respons yang sangat santun, profesional, dan to-the-point.\n"
        "• Multilingual Engine: AI Hadir.io secara adaptif mengenali dan merespons dalam 4 bahasa: Bahasa Indonesia (bahasa resmi), "
        "English (komunikasi internasional), Basa Sunda (bahasa daerah Jawa Barat), dan Basa Jawa (dialek Cirebonan).\n"
        "• RBAC Query Firewall: Sistem menerapkan filter keamanan di tingkat backend yang menolak permintaan informasi sensitif "
        "seperti tanggal lahir, usia personal, maupun struktur kode internal sistem."
    )

    # Gambar 1.1: Diagram Arsitektur AI Hadir.io
    ai_diag = r"c:\xampp\htdocs\absen\Diagram_Arsitektur_AI_Hadirio.png"
    add_image_with_caption(doc, ai_diag, "Gambar 1.1 Diagram Arsitektur Dual-Engine AI Hadir.io & Guardrail RBAC")

    doc.add_page_break()

    # =========================================================================
    # BAB II: SOFTWARE DESIGN DOCUMENT (SDD) & PEMODELAN UML
    # =========================================================================
    add_heading_1(doc, "BAB II: SOFTWARE DESIGN DOCUMENT (SDD) & PEMODELAN UML")

    add_heading_2(doc, "2.1 Arsitektur Perangkat Lunak Terpadu (MVC Pattern)")
    add_p_justified(
        doc,
        "Sistem Presensi Digital SMK NU WARGABINANGUN mengadopsi arsitektur Model-View-Controller (MVC) murni berbasis "
        "PHP 8.2 Native. Pola arsitektur ini memisahkan secara tegas tanggung jawab penanganan data (Model), penyajian antarmuka (View), "
        "dan logika alur bisnis (Controller). Dengan meniadakan framework berat pihak ketiga, aplikasi menghasilkan performa eksekusi "
        "yang sangat ringan, latensi rendah, serta kemudahan pemeliharaan dan audit kode (Clean Architecture)."
    )

    add_heading_2(doc, "2.2 Pemodelan 4 Diagram UML Standar Utama")
    add_p_justified(
        doc,
        "Pemodelan sistem mengacu pada standar Unified Modeling Language (UML 2.5) yang mencakup 4 diagram representatif:"
    )

    # 2.2.1 Use Case Diagram
    add_heading_3(doc, "2.2.1 Use Case Diagram Terpadu & Skenario Use Case Lengkap")
    add_p_justified(
        doc,
        "Use Case Diagram dirancang secara terpadu merangkul seluruh 20 tabel database `db_absensi_smk`. Diagram ini memetakan "
        "interaksi antara 5 Aktor Pengguna internal (Administrator, Kepala Sekolah, Wali Kelas, Guru Pengajar, Siswa) yang "
        "ditempatkan secara konsisten di sisi kiri, serta 2 Aktor Sistem Eksternal (Bot Telegram @HadirioBot dan NVIDIA NIM REST API) "
        "di sisi kanan, berinteraksi dengan 14 Use Case inti melalui relasi <<include>> dan <<extend>>:"
    )
    uc_img = r"c:\xampp\htdocs\absen\Usecase_Diagram_Sistem_Absensi_SMK_NU_Wargabinangun.png"
    add_image_with_caption(doc, uc_img, "Gambar 2.1 Use Case Diagram Terpadu Sistem Presensi SMK NU WARGABINANGUN (20 Tabel Database)")

    add_p_justified(
        doc,
        "Tabel 2.1 menyajikan Skenario Use Case mendalam untuk Use Case inti UC08 (Scan Presensi Barcode/QR) "
        "yang menjadi tulang punggung operasional sekolah:"
    )

    # Tabel 2.1
    tbl_scen = doc.add_table(rows=7, cols=2)
    tbl_scen.alignment = WD_TABLE_ALIGNMENT.CENTER
    scen_data = [
        ("Nama Use Case", "UC08: Scan Presensi QR (Masuk & Pulang)"),
        ("Aktor Terlibat", "Siswa, Guru Pengajar, Wali Kelas (Aktor Utama), Bot Telegram (Aktor Sistem)"),
        ("Kondisi Awal (Pre-condition)", "Kamera pemindai aktif pada antarmuka kiosk/peramban, pengguna memiliki kartu barcode terdaftar."),
        ("Kondisi Akhir (Post-condition)", "Data presensi tersimpan di DB, status tervalidasi (Tepat Waktu/Terlambat), notifikasi Telegram terkirim."),
        ("Alur Utama (Main Flow)", "1. Pengguna mengarahkan kartu barcode/QR ke arah kamera pemindai.\n2. Library HTML5-QRCode mendeteksi string payload dan mengirim permintaan AJAX ke Controller.\n3. Controller memverifikasi jam server dan aturan toleransi keterlambatan (07:00 WIB).\n4. Database mengeksekusi INSERT/UPDATE data kehadiran (termasuk Dual-Sync jika peran Wali Kelas).\n5. Server memicu webhook cURL ke Bot Telegram (@HadirioBot) untuk mengirimkan pesan rincian.\n6. Antarmuka menampilkan audio beep sukses dan toast hijau rincian presensi."),
        ("Alur Alternatif (Alternative Flow)", "3a. Jika jam scan sebelum 15:00 saat kepulangan: Sistem menolak dan menampilkan toast 'Belum Jam Pulang'.\n3b. Jika kartu tidak terdaftar di database: Sistem membunyikan audio error dan menampilkan 'Barcode Tidak Dikenali'."),
        ("Tabel Basis Data Terkait", "absensi_siswa, absensi_guru, absensi_wali_kelas, pengaturan_sistem")
    ]
    for i, (k, v) in enumerate(scen_data):
        c0 = tbl_scen.cell(i, 0); c1 = tbl_scen.cell(i, 1)
        c0.text = k; c1.text = v
    style_table_header(tbl_scen.rows[0], [1.8, 4.7])
    for idx, row in enumerate(tbl_scen.rows[1:], start=1):
        style_table_row(row, is_even=(idx % 2 == 1), col_widths=[1.8, 4.7])

    p_cap_t21 = doc.add_paragraph("Tabel 2.1 Skenario Use Case Inti UC08 (Scan Presensi QR Masuk & Pulang)")
    p_cap_t21.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_t21.paragraph_format.space_before = Pt(3)
    p_cap_t21.paragraph_format.space_after = Pt(8)
    p_cap_t21.runs[0].font.name = "Times New Roman"
    p_cap_t21.runs[0].font.size = Pt(9.5)
    p_cap_t21.runs[0].font.italic = True

    # 2.2.2 Activity Diagram
    add_heading_3(doc, "2.2.2 Activity Diagram Siklus Bisnis & Alur AI (5 Swimlanes)")
    add_p_justified(
        doc,
        "Activity Diagram memodelkan alur kerja sekuensial dan paralel dari seluruh siklus presensi melewati 5 Swimlanes "
        "(Aktor Pengguna, UI Layer, Controller Layer, Database 20 Tabel, dan Layanan Eksternal). Diagram ini memperlihatkan "
        "titik percabangan logika (decision node) saat validasi waktu dan pemanggilan layanan cerdas AI Hadir.io:"
    )
    act_img = r"c:\xampp\htdocs\absen\Activity_Diagram_Sistem_Absensi_SMK_NU_Wargabinangun.png"
    add_image_with_caption(doc, act_img, "Gambar 2.2 Activity Diagram Siklus Bisnis & Alur Integrasi AI (5 Swimlanes)")

    # Detail Sub-Activity Diagrams
    add_p_justified(
        doc,
        "Untuk membedah alur logika pada setiap subsistem secara terperinci, berikut disajikan diagram aktivitas turunan untuk "
        "modul pemindaian presensi kamera, pengajuan perizinan mandiri, serta supervisi bimbingan akademik:"
    )
    act_sub1 = r"c:\xampp\htdocs\absen\diagram_uml_jpeg\06_Activity_Diagram_05_Presensi_Scan_QR.jpeg"
    add_image_with_caption(doc, act_sub1, "Gambar 2.3 Activity Diagram Alur Presensi Scan QR Kamera & Deteksi Toleransi Waktu")

    act_sub2 = r"c:\xampp\htdocs\absen\diagram_uml_jpeg\07_Activity_Diagram_06_Pengajuan_Izin_Sakit.jpeg"
    add_image_with_caption(doc, act_sub2, "Gambar 2.4 Activity Diagram Alur Pengajuan Izin / Sakit Mandiri & Auto-Sync Approval")

    act_sub3 = r"c:\xampp\htdocs\absen\diagram_uml_jpeg\09_Activity_Diagram_08_Catatan_Supervisi_Pembinaan.jpeg"
    add_image_with_caption(doc, act_sub3, "Gambar 2.5 Activity Diagram Supervisi Evaluasi KBM & Catatan Pembinaan Siswa")

    # 2.2.3 Class Diagram
    add_heading_3(doc, "2.2.3 Class Diagram Berorientasi Objek & Kepatuhan Prinsip SOLID")
    add_p_justified(
        doc,
        "Class Diagram merefleksikan struktur berorientasi objek aplikasi yang dirancang dengan mematuhi prinsip SOLID:\n"
        "• Single Responsibility Principle (SRP): Setiap kelas Controller hanya menangani satu konteks domain (misal: AbsensiController khusus mengelola transaksi presensi, IzinController mengelola pengajuan dispensasi).\n"
        "• Open/Closed Principle (OCP): Arsitektur layanan eksternal (TelegramService dan AIService) dirancang modular sehingga penambahan provider baru dapat dilakukan tanpa mengubah logika inti controller.\n"
        "• Liskov Substitution Principle (LSP): Seluruh entitas Model mewarisi BaseModel yang menjamin konsistensi interaksi database PDO.\n"
        "• Interface Segregation Principle (ISP): Antarmuka service dipisahkan berdasarkan fungsi spesifik sehingga tidak ada kelas yang dipaksa mengimplementasikan fungsi yang tidak dibutuhkannya.\n"
        "• Dependency Inversion Principle (DIP): Controller bergantung pada abstraksi konfigurasi database terpusat melalui file config global."
    )
    cls_img = r"c:\xampp\htdocs\absen\Class_Diagram_Sistem_Absensi_SMK_NU_Wargabinangun.png"
    add_image_with_caption(doc, cls_img, "Gambar 2.6 Class Diagram Berorientasi Objek Berbasis Prinsip SOLID")

    # 2.2.4 Sequence Diagram
    add_heading_3(doc, "2.2.4 Sequence Diagram Alur Pertukaran Pesan Real-Time")
    add_p_justified(
        doc,
        "Sequence Diagram mendokumentasikan urutan kronologis pertukaran pesan (message passing) antara antarmuka pengguna, "
        "Controller, database MariaDB, Bot Telegram, dan NVIDIA NIM REST API secara live:"
    )
    seq_img = r"c:\xampp\htdocs\absen\Sequence_Diagram_Sistem_Absensi_SMK_NU_Wargabinangun.png"
    add_image_with_caption(doc, seq_img, "Gambar 2.7 Sequence Diagram Alur Pertukaran Pesan Presensi & AI Hadir.io")

    # 2.3 ERD & Basis Data
    add_heading_2(doc, "2.3 Perancangan Basis Data & Entity Relationship Diagram (ERD)")
    add_p_justified(
        doc,
        "Basis data `db_absensi_smk` dirancang secara teliti mencapai Bentuk Normal Ketiga (3NF). Seluruh dependensi parsial "
        "dan dependensi transitif telah dieliminasi melalui pembentukan 20 tabel relasional yang saling terhubung menggunakan "
        "Primary Key (PK) dan Foreign Key (FK) berindeks cepat:"
    )
    erd_img = r"c:\xampp\htdocs\absen\ERD_Database_Absensi_SMK.png"
    add_image_with_caption(doc, erd_img, "Gambar 2.8 Entity Relationship Diagram (ERD) Basis Data db_absensi_smk (20 Tabel)")

    add_heading_3(doc, "Kamus Data Komprehensif 20 Tabel Database `db_absensi_smk`")
    add_p_justified(
        doc,
        "Tabel 2.2 memuat kamus data komprehensif yang merinci struktur kunci dan fungsi bisnis dari ke-20 tabel relasional:"
    )

    # Tabel 2.2
    tbl_dict = doc.add_table(rows=21, cols=3)
    tbl_dict.alignment = WD_TABLE_ALIGNMENT.CENTER
    dict_full = [
        ("Nama Tabel", "Struktur Kunci & Kolom Utama", "Fungsi Bisnis dalam Sistem Presensi"),
        ("users", "PK: id | username, password (bcrypt), role, status, email, no_telp", "Menyimpan kredensial otentikasi akun dan profil seluruh sivitas sekolah."),
        ("roles", "PK: id | nama_role (admin, kepala_sekolah, guru, wali_kelas, siswa)", "Tingkatan hak akses RBAC untuk otorisasi menu dan modul portal."),
        ("tahun_ajaran", "PK: id | tahun, semester (ganjil/genap), is_active", "Mengelola periode tahun akademik aktif yang sedang berlangsung."),
        ("jurusan", "PK: id | kode_jurusan (TKJ, TKR, TSM), nama_jurusan", "Master data program keahlian kejuruan vokasi SMK."),
        ("kelas", "PK: id | FK: jurusan_id | tingkat (X, XI, XII), nama_kelas", "Master data rombongan belajar (rombel) siswa sekolah."),
        ("guru", "PK: id | FK: user_id | nip, barcode, nama_lengkap, status_aktif", "Menyimpan data induk pendidik, NIP unik, dan relasi akun login."),
        ("wali_kelas_data", "PK: id | FK: user_id | nip, barcode, nama_lengkap, no_telp", "Data master guru yang mengemban tugas manajerial pembina kelas."),
        ("wali_kelas", "PK: id | FK: wali_kelas_data_id, kelas_id, tahun_ajaran_id", "Mapping penugasan wali kelas terhadap rombel pada tahun ajaran aktif."),
        ("siswa", "PK: id | FK: kelas_id, user_id | nis, nisn, barcode, nama_lengkap", "Data induk peserta didik, NISN nasional, dan kode barcode kartu presensi."),
        ("mata_pelajaran", "PK: id | FK: guru_id, wali_kelas_id | kode_mapel, nama_mapel, tingkat", "Katalog mata pelajaran kurikulum dan guru pengampu."),
        ("jadwal_mengajar", "PK: id | FK: guru_id, kelas_id, mapel_id, tahun_ajaran_id | hari, jam", "Jadwal KBM mingguan yang mengatur jam mengajar guru dan kelas."),
        ("barcode_guru", "PK: id | FK: guru_id | code (unique), is_active", "String enkripsi barcode kartu pendidik untuk validasi scan kamera."),
        ("rfid_cards", "PK: id | FK: siswa_id | uid (unique), is_active", "UID fisik kartu RFID/NFC siswa untuk alternatif pemindai kartu pintar."),
        ("absensi_siswa", "PK: id | FK: siswa_id | tanggal, jam_masuk, jam_pulang, status", "Merekam riwayat transaksi kehadiran harian siswa di sekolah."),
        ("absensi_guru", "PK: id | FK: guru_id | tanggal, jam_masuk, jam_pulang, status", "Merekam riwayat kehadiran harian guru dan jam mengajar KBM."),
        ("absensi_wali_kelas", "PK: id | FK: wali_kelas_id | tanggal, jam_masuk, jam_pulang, status", "Merekam absensi dewan wali kelas hasil sinkronisasi Dual-Sync otomatis."),
        ("pengajuan_izin", "PK: id | FK: siswa_id, guru_id | tipe_pemohon, jenis, bukti_foto, status", "Layanan perizinan paperless mandiri siswa/guru dengan verifikasi foto dokter."),
        ("catatan_guru", "PK: id | FK: guru_id, kepala_sekolah_id | judul, isi_catatan, status_baca", "Kanal supervisi tertulis evaluasi KBM Kepala Sekolah ke Guru berpelacak baca."),
        ("catatan_pembinaan", "PK: id | FK: siswa_id, kelas_id | kategori, isi_catatan, status_baca", "Kanal pembinaan kedisiplinan siswa dari Wali Kelas berstatus konfirmasi baca."),
        ("pengaturan_sistem", "PK: id | setting_key (unique), setting_value, created_at", "Penyimpanan konfigurasi global jam masuk/pulang, token Bot Telegram, dan API AI.")
    ]
    for i, r_vals in enumerate(dict_full):
        for j, val in enumerate(r_vals):
            c = tbl_dict.cell(i, j)
            c.text = val
    style_table_header(tbl_dict.rows[0], [1.4, 2.3, 2.8])
    for idx, row in enumerate(tbl_dict.rows[1:], start=1):
        style_table_row(row, is_even=(idx % 2 == 1), col_widths=[1.4, 2.3, 2.8])

    p_cap_t22 = doc.add_paragraph("Tabel 2.2 Kamus Data Komprehensif 20 Tabel Database db_absensi_smk")
    p_cap_t22.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_t22.paragraph_format.space_before = Pt(3)
    p_cap_t22.paragraph_format.space_after = Pt(8)
    p_cap_t22.runs[0].font.name = "Times New Roman"
    p_cap_t22.runs[0].font.size = Pt(9.5)
    p_cap_t22.runs[0].font.italic = True

    # 2.4 Perancangan UI/UX & Prototipe Figma
    add_heading_2(doc, "2.4 Perancangan Antarmuka Pengguna & Prototipe Interaktif Figma")
    add_p_justified(
        doc,
        "Desain antarmuka pengguna dirancang menggunakan perangkat lunak industri Figma dengan pendekatan Modern Dark Mode "
        "yang ergonomis dan nyaman di mata pengguna. Gambar 2.9 memperlihatkan papan showcase menyeluruh prototipe desain Figma:"
    )
    fig_img = r"c:\xampp\htdocs\absen\Desain_UI_UX_Figma_Sistem_Presensi_SMK.png"
    add_image_with_caption(doc, fig_img, "Gambar 2.9 Showcase Board Prototipe UI/UX Figma Sistem Presensi SMK NU Wargabinangun")

    add_p_justified(
        doc,
        "Berikut disajikan dokumentasi visual resolusi tinggi dan bedah fungsionalitas dari ke-6 layar utama prototipe aplikasi:\n\n"
        "1. Prototipe Layar 1: Portal Login Multi-Peran (RBAC Auth)\n"
        "Formulir login tunggal terenkripsi yang secara otomatis memvalidasi hak akses dari 5 peran pengguna berbeda "
        "(Admin, Kepala Sekolah, Wali Kelas, Guru, dan Siswa) menggunakan verifikasi kata sandi Bcrypt hash:"
    )
    scr1 = r"c:\xampp\htdocs\absen\UI_Screen_01_Login_Multi_Role.png"
    add_image_with_caption(doc, scr1, "Gambar 2.10 Prototipe Layar 1: Portal Login Multi-Peran (RBAC Auth)")

    add_p_justified(
        doc,
        "2. Prototipe Layar 2: Dasbor Analitik Administrator & Master Data\n"
        "Dasbor manajerial yang menampilkan ringkasan metrik statistik kehadiran harian secara real-time, grafik tren kehadiran mingguan "
        "Senin–Jumat, serta pintasan cepat operasi pencetakan massal kartu identitas barcode presensi:"
    )
    scr2 = r"c:\xampp\htdocs\absen\UI_Screen_02_Dashboard_Admin.png"
    add_image_with_caption(doc, scr2, "Gambar 2.11 Prototipe Layar 2: Dasbor Analitik Administrator & Master Data")

    add_p_justified(
        doc,
        "3. Prototipe Layar 3: Antarmuka Pemindai Barcode/QR Real-Time (/scan.php)\n"
        "Antarmuka mode kiosk pemindai kamera responsif berbasis library HTML5-QRCode. Pemindai mengevaluasi jam server lokal "
        "dalam waktu < 0.5 detik, membunyikan nada audio beep sukses, dan memicu webhook notifikasi Telegram secara simultan:"
    )
    scr3 = r"c:\xampp\htdocs\absen\UI_Screen_03_Barcode_Scanner_Realtime.png"
    add_image_with_caption(doc, scr3, "Gambar 2.12 Prototipe Layar 3: Antarmuka Pemindai Barcode/QR Real-Time (/scan.php)")

    add_p_justified(
        doc,
        "4. Prototipe Layar 4: Supervisi KBM Guru & Rekapitulasi Presensi\n"
        "Kanal supervisi akademik bagi Kepala Sekolah untuk mengevaluasi KBM guru dan memantau kehadiran harian, lengkap dengan "
        "fitur pelacak status baca (*Read Status Tracking*) yang menunjukkan apakah arahan evaluasi telah dibaca oleh guru terkait:"
    )
    scr4 = r"c:\xampp\htdocs\absen\UI_Screen_04_Supervisi_KBM_Rekap.png"
    add_image_with_caption(doc, scr4, "Gambar 2.13 Prototipe Layar 4: Supervisi KBM Guru & Rekapitulasi Presensi")

    add_p_justified(
        doc,
        "5. Prototipe Layar 5: Portal Siswa, Riwayat Kehadiran & Form Izin Online\n"
        "Antarmuka mandiri bagi peserta didik untuk memantau akumulasi kehadiran bulanan secara transparan serta mengajukan surat "
        "permohonan izin atau sakit tanpa kertas (*paperless*) dengan melampirkan berkas foto surat dokter yang tervalidasi:"
    )
    scr5 = r"c:\xampp\htdocs\absen\UI_Screen_05_Portal_Siswa_Izin.png"
    add_image_with_caption(doc, scr5, "Gambar 2.14 Prototipe Layar 5: Portal Siswa, Riwayat Kehadiran & Form Izin Online")

    add_p_justified(
        doc,
        "6. Prototipe Layar 6: Widget Percakapan Asisten AI Hadir.io (4 Bahasa)\n"
        "Jendela percakapan cerdas mengambang (*floating widget*) yang mampu menjawab pertanyaan seputar data akademik sekolah "
        "dalam 4 bahasa (Indonesia, Inggris, Sunda, dan Jawa) berlandaskan kueri database faktual dan model NVIDIA NIM Nemotron LLM:\n\n"
        "Tautan Resmi Berkas Desain Figma: https://www.figma.com/design/HJ3KQU3E6jYNRBvK1X99aU/Untitled?node-id=0-1"
    )
    scr6 = r"c:\xampp\htdocs\absen\UI_Screen_06_Widget_Chat_AI_Hadirio.png"
    add_image_with_caption(doc, scr6, "Gambar 2.15 Prototipe Layar 6: Widget Percakapan Asisten AI Hadir.io (4 Bahasa)")

    doc.add_page_break()

    # =========================================================================
    # BAB III: IMPLEMENTASI KODE, PENGUJIAN & CI/CD DEVOPS
    # =========================================================================
    add_heading_1(doc, "BAB III: IMPLEMENTASI KODE, TESTING & CI/CD DEVOPS")

    add_heading_2(doc, "3.1 Struktur Direktori Proyek Bersih (MVC Architecture)")
    add_p_justified(
        doc,
        "Implementasi sistem mengikuti tata kelola repositori perangkat lunak profesional dengan struktur pohon berkas yang teratur:\n"
        "• `app/controllers/`: Menangani logika otentikasi (AuthController.php), pemindaian barcode (ScanController.php), pengelolaan perizinan (IzinController.php), dan pelaporan rekapitulasi (LaporanController.php).\n"
        "• `app/models/`: Berisi kelas representasi entitas data dengan metode Prepared Statements PDO untuk mencegah kerentanan SQL Injection.\n"
        "• `app/views/`: Berkas template antarmuka responsif Bootstrap 5 yang dipisahkan ke dalam subdirektori per peran (admin, guru, kepsek, siswa, walikelas).\n"
        "• `app/services/`: Modul integrasi eksternal, mencakup TelegramService.php untuk pengiriman pesan webhook dan AIService.php untuk komunikasi API NVIDIA NIM.\n"
        "• `chat_api.php`: Endpoint JSON backend yang memproses kueri bahasa alami dari widget AI Hadir.io.\n"
        "• `tests/`: Skrip automated unit test untuk pengujian logika waktu keterlambatan dan integritas database.\n"
        "• `.github/workflows/ci.yml`: Pipeline otomatisasi pengujian CI/CD GitHub Actions."
    )

    mvc_diag = r"c:\xampp\htdocs\absen\Diagram_Arsitektur_MVC_PHP.png"
    add_image_with_caption(doc, mvc_diag, "Gambar 3.1 Diagram Arsitektur Pola Desain MVC Native PHP 8.2")

    add_heading_2(doc, "3.2 Implementasi Fitur Kunci & Potongan Kode (Code Highlights)")
    add_p_justified(
        doc,
        "Berikut adalah implementasi logika inti yang ditanamkan dalam sistem presensi digital:\n\n"
        "1. Validasi Waktu Masuk & Deteksi Keterlambatan:\n"
        "Controller memvalidasi waktu pemindaian berdasarkan jam server lokal. Jika pemindaian dilakukan antara 05:00 hingga 07:00 WIB, "
        "status dicatat sebagai 'Hadir Tepat Waktu'. Apabila pemindaian dilakukan lewat dari pukul 07:00 WIB, status otomatis berlabel 'Terlambat'.\n\n"
        "2. Validasi Jam Kepulangan:\n"
        "Untuk mencegah peserta didik atau guru pulang sebelum waktunya, tombol/mode scan kepulangan hanya menerima pemindaian mulai pukul 15:00 WIB ke atas. "
        "Pemindaian sebelum jam tersebut akan memunculkan notifikasi penolakan dari sistem."
    )

    time_diag = r"c:\xampp\htdocs\absen\Diagram_Alur_Validasi_Waktu.png"
    add_image_with_caption(doc, time_diag, "Gambar 3.2 Diagram Logika Validasi Toleransi Keterlambatan dan Batas Jam Kepulangan")

    add_p_justified(
        doc,
        "3. Mekanisme Transaksi Atomik Dual-Sync Wali Kelas:\n"
        "Saat seorang guru yang mengemban amanah sebagai Wali Kelas memindai kartu presensi di gerbang sekolah, controller secara atomik "
        "menjalankan blok transaksi database PDO (beginTransaction - commit) yang mencatat kehadiran ke dalam tabel `absensi_guru` dan `absensi_wali_kelas` "
        "secara simultan tanpa perlu melakukan pemindaian ganda:\n\n"
        "4. Integrasi Webhook Bot Telegram (@HadirioBot):\n"
        "Pengiriman pesan notifikasi live dilakukan menggunakan cURL HTTP POST dengan opsi khusus `CURLOPT_IPRESOLVE = CURL_IPRESOLVE_V4` "
        "dan `CURLOPT_SSL_VERIFYPEER = false` pada server lokal guna mengeliminasi kendala timeout SSL pada lingkungan Windows XAMPP.\n\n"
        "5. Mekanisme Auto-Sync Persetujuan Izin:\n"
        "Ketika Administrator menekan tombol 'Setujui' pada entitas pengajuan izin, sistem secara otomatis mengeksekusi sinkronisasi update "
        "ke tabel `absensi_siswa` atau `absensi_guru` sehingga status presensi langsung berubah menjadi 'Izin' atau 'Sakit' tanpa intervensi manual."
    )

    sync_diag = r"c:\xampp\htdocs\absen\Diagram_Dual_Sync_Walikelas.png"
    add_image_with_caption(doc, sync_diag, "Gambar 3.3 Diagram Mekanisme Transaksi Atomik Dual-Sync Wali Kelas (PDO)")

    add_heading_2(doc, "3.3 Laporan Pengujian Perangkat Lunak (Software Testing Report)")
    add_p_justified(
        doc,
        "Pengujian perangkat lunak dilakukan secara komprehensif menggunakan kombinasi metode Black-Box Testing pada antarmuka "
        "dan Automated Unit Testing pada fungsi controller. Seluruh 10 kasus uji inti berhasil dilewati dengan tingkat kelulusan 100% PASS:"
    )

    # Tabel 3.1
    tbl_tc = doc.add_table(rows=11, cols=4)
    tbl_tc.alignment = WD_TABLE_ALIGNMENT.CENTER
    tc_full = [
        ("Kode Uji", "Skenario Pengujian Sistem", "Hasil yang Diharapkan (Expected Result)", "Status Uji"),
        ("TC-01", "Otentikasi Login dengan password salah", "Sistem menolak login, mencatat percobaan gagal, dan menampilkan notifikasi kesalahan.", "PASS (100%)"),
        ("TC-02", "Otentikasi Login Multi-Role (5 Role)", "Sistem berhasil memvalidasi kredensial dan merender dasbor yang sesuai dengan wewenang role.", "PASS (100%)"),
        ("TC-03", "Pemindaian Barcode Jam 06:45 WIB", "Sistem mencatat presensi masuk berstatus 'Hadir Tepat Waktu' dan audio beep sukses berbunyi.", "PASS (100%)"),
        ("TC-04", "Pemindaian Barcode Jam 07:20 WIB", "Sistem mencatat presensi masuk berstatus 'Terlambat' dan mencatat log menit keterlambatan.", "PASS (100%)"),
        ("TC-05", "Pemindaian Kepulangan Sebelum Jam 15:00", "Sistem menolak scan pulang dan memberikan peringatan 'Belum Masuk Jam Pulang'.", "PASS (100%)"),
        ("TC-06", "Fitur Dual-Sync Presensi Wali Kelas", "Satu kali scan kartu wali kelas otomatis mengupdate tabel absensi_guru dan absensi_wali_kelas.", "PASS (100%)"),
        ("TC-07", "Pengiriman Notifikasi Telegram Otomatis", "Pesan notifikasi rincian kehadiran masuk ke aplikasi Telegram ponsel secara real-time.", "PASS (100%)"),
        ("TC-08", "Pengajuan Izin Mandiri & Approval Admin", "Status perizinan terunggah rapi dan saat disetujui otomatis menyinkronkan status absensi siswa.", "PASS (100%)"),
        ("TC-09", "Filter Rekapitulasi Presensi Senin-Jumat", "Kalkulasi rekapitulasi kehadiran mengecualikan hari Sabtu dan Minggu dari pembagi total hari sekolah.", "PASS (100%)"),
        ("TC-10", "Interaksi Asisten AI Hadir.io (4 Bahasa)", "Chatbot menjawab jumlah kelas (9) dan guru (33) secara akurat, cepat, dan to-the-point.", "PASS (100%)")
    ]
    for i, r_vals in enumerate(tc_full):
        for j, val in enumerate(r_vals):
            c = tbl_tc.cell(i, j)
            c.text = val
    style_table_header(tbl_tc.rows[0], [0.9, 2.3, 2.5, 0.8])
    for idx, row in enumerate(tbl_tc.rows[1:], start=1):
        style_table_row(row, is_even=(idx % 2 == 1), col_widths=[0.9, 2.3, 2.5, 0.8])
        row.cells[3].paragraphs[0].runs[0].bold = True
        row.cells[3].paragraphs[0].runs[0].font.color.rgb = RGBColor(0x05, 0x96, 0x69)

    p_cap_t31 = doc.add_paragraph("Tabel 3.1 Matriks Hasil Pengujian Perangkat Lunak (10 Test Cases 100% PASS)")
    p_cap_t31.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_t31.paragraph_format.space_before = Pt(3)
    p_cap_t31.paragraph_format.space_after = Pt(8)
    p_cap_t31.runs[0].font.name = "Times New Roman"
    p_cap_t31.runs[0].font.size = Pt(9.5)
    p_cap_t31.runs[0].font.italic = True

    add_heading_2(doc, "3.4 Otomasi CI/CD Pipeline & GitHub Actions")
    add_p_justified(
        doc,
        "Untuk memenuhi standar DevOps modern pada rubrik penilaian OBE, proyek dikonfigurasikan dengan otomatisasi "
        "CI/CD Pipeline menggunakan GitHub Actions pada berkas `.github/workflows/ci.yml`. Pipeline terpicu otomatis pada setiap "
        "aktivitas push atau pull request ke branch main, mencakup dua tahapan utama:\n"
        "• Tahap PHP Linting & Syntax Validation: Memindai seluruh berkas PHP dalam repositori untuk memastikan tidak ada kesalahan sintaksis.\n"
        "• Tahap Automated Unit Testing: Mengeksekusi pengujian logika controller menggunakan runner PHP 8.2.\n"
        "• Status Pipeline: Seluruh tahapan build dinyatakan PASS dengan badge indikator hijau pada README repositori."
    )

    cicd_diag = r"c:\xampp\htdocs\absen\Diagram_Hasil_Testing_Pass.png"
    add_image_with_caption(doc, cicd_diag, "Gambar 3.4 Diagram Hasil Automated Unit Testing (100% PASS) & Pipeline CI/CD")

    add_heading_2(doc, "3.5 Petunjuk Instalasi & Deployment Sistem (Setup Instructions)")
    add_p_justified(
        doc,
        "Berikut adalah panduan lengkap deployment aplikasi pada lingkungan server lokal (XAMPP):\n"
        "1. Kebutuhan Perangkat Lunak: Web Server Apache 2.4+, PHP 8.2+ (ekstensi pdo_mysql, curl, gd, openssl, mbstring aktif), MariaDB 10.4+.\n"
        "2. Pemasangan Berkas: Kloning atau salin seluruh repositori ke direktori web server: `c:/xampp/htdocs/absen`.\n"
        "3. Konfigurasi Basis Data: Buka phpMyAdmin, buat database baru bernama `db_absensi_smk`, lalu impor berkas skema `db_absensi_smk.sql`.\n"
        "4. Pengaturan Koneksi: Pastikan kredensial pada `app/config/database.php` terkonfigurasi dengan benar:\n"
        "   `define('DB_HOST', 'localhost'); define('DB_USER', 'root'); define('DB_PASS', ''); define('DB_NAME', 'db_absensi_smk');`\n"
        "5. Pengaturan API Token: Pastikan token Bot Telegram dan API Key NVIDIA NIM telah terdaftar pada tabel `pengaturan_sistem`.\n"
        "6. Menjalankan Sistem: Buka peramban web dan akses URL `http://localhost/absen`. Masuk menggunakan akun terdaftar."
    )

    doc.add_page_break()

    # =========================================================================
    # BAB IV: KESIMPULAN, REFLEKSI REKAYASA & MATRIKS DELIVERABLES
    # =========================================================================
    add_heading_1(doc, "BAB IV: KESIMPULAN & MATRIKS DELIVERABLES")

    add_heading_2(doc, "4.1 Kesimpulan Proyek")
    add_p_justified(
        doc,
        "Berdasarkan seluruh tahapan perancangan, konstruksi kode, pengujian, dan deployment yang telah diselesaikan, "
        "dapat disimpulkan bahwa:\n"
        "1. Sistem Presensi Digital SMK NU WARGABINANGUN telah berhasil dibangun 100% tuntas secara mandiri, mengatasi kelemahan "
        "sistem presensi kertas konvensional melalui otomasi pemindaian QR instan (< 0.5 detik), verifikasi jam server ketat, dan "
        "notifikasi live Telegram (@HadirioBot) dalam hitungan detik.\n"
        "2. Seluruh 20 tabel relasional basis data `db_absensi_smk` telah terintegrasi sempurna dengan 5 tingkatan hak akses pengguna (RBAC), "
        "mendukung fitur sinkronisasi ganda (Dual-Sync) pendidik, otomatisasi approval izin (Auto-Sync), dan kanal supervisi KBM terarah.\n"
        "3. Fitur kecerdasan artifisial Asisten AI Hadir.io berbasis NVIDIA NIM Nemotron LLM memberikan nilai tambah teknologi nyata sebagai "
        "asisten pintar sekolah yang akurat, santun, multi-bahasa, dan terisolasi aman oleh guardrail RBAC."
    )

    add_heading_2(doc, "4.2 Refleksi Pembelajaran Mandiri (Engineering Reflection)")
    add_p_justified(
        doc,
        "Pengerjaan proyek mandiri dengan beban penuh (100% Individual Ownership & Full-App Build) ini memberikan pengalaman "
        "rekayasa perangkat lunak yang sangat mendalam dan transformatif bagi penulis. Penulis belajar bagaimana mentransformasikan "
        "kebutuhan operasional sekolah vokasi menjadi arsitektur perangkat lunak yang modular, mematuhi standar pemodelan UML 2.5, "
        "merancang normalisasi basis data 20 tabel (3NF), menerapkan kode bersih berorientasi objek (Clean Code & SOLID), "
        "mengintegrasikan API pihak ketiga (Telegram Webhook & NVIDIA NIM AI), serta mengelola alur kerja DevOps modern melalui "
        "automated unit testing dan CI/CD pipeline GitHub Actions. Pengalaman komprehensif ini menumbuhkan kepercayaan diri dan "
        "kompetensi profesional penulis sebagai calon sarjana Teknik Informatika yang siap bersaing di industri teknologi."
    )

    add_heading_2(doc, "4.3 Matriks Checklist Pengumpulan Berkas (Submission Deliverables)")
    add_p_justified(
        doc,
        "Sesuai dengan ketentuan pengumpulan berkas akhir pada Panduan Tugas Proyek Mandiri Rekayasa Perangkat Lunak "
        "STMIK IKMI Cirebon (Pertemuan ke-16 / UAS):\n"
        "1. Repositori GitHub Publik Pribadi: Repositori publik dengan format username/proyek-rpl-nama-aplikasi "
        "telah aktif dan memuat berkas README.md lengkap (identitas Nama, NIM, Kelompok, Tema Topik, Deskripsi Aplikasi & Fitur AI, "
        "tautan video progres 1, video progres 2, video final demo, serta petunjuk instalasi setup instruction).\n"
        "2. Berkas Dokumen Lengkap (PDF Tunggal): Seluruh dokumen PRD/SRS, SDD (Pemodelan UML & UI/UX Figma), dan Laporan Pengujian/CI-CD "
        "dijilid secara komprehensif dalam satu berkas PDF resmi per mahasiswa untuk diunggah ke portal LMS perkuliahan.\n"
        "3. Tenggat Waktu Pengumpulan: Seluruh berkas PDF tunggal dan tautan repositori GitHub diunggah tepat waktu sebelum batas akhir "
        "jadwal Ujian Akhir Semester (UAS)."
    )

    # Tabel 4.1
    tbl_sub = doc.add_table(rows=6, cols=2)
    tbl_sub.alignment = WD_TABLE_ALIGNMENT.CENTER
    sub_data = [
        ("Nama Berkas Luaran Wajib", "Tautan Akses / Status Pengumpulan"),
        ("1. Repositori GitHub Publik Pribadi", "https://github.com/Wwinaa/proyek-rpl-absensiQR"),
        ("2. Berkas README.md Repositori", "Memuat identitas, deskripsi AI, tautan video, dan petunjuk instalasi"),
        ("3. Tautan Video Presentasi & Demo (YouTube)", "https://youtu.be/d8n-Pi-SUn4?si=7W71KGDXO2uolbbd"),
        ("4. Berkas Dokumen Lengkap (PDF Tunggal)", "https://github.com/Wwinaa/proyek-rpl-absensiQR/blob/main/Laporan_Proyek_Akhir_RPL_Sistem_Presensi_SMK_NU_Wargabinangun.pdf"),
        ("5. Prototipe Interaktif Desain UI/UX Figma", "https://www.figma.com/design/HJ3KQU3E6jYNRBvK1X99aU/Untitled?node-id=0-1")
    ]
    for i, (k, v) in enumerate(sub_data):
        c0 = tbl_sub.cell(i, 0); c1 = tbl_sub.cell(i, 1)
        c0.text = k; c1.text = v
    style_table_header(tbl_sub.rows[0], [2.2, 4.3])
    for idx, row in enumerate(tbl_sub.rows[1:], start=1):
        style_table_row(row, is_even=(idx % 2 == 1), col_widths=[2.2, 4.3])

    p_cap_t41 = doc.add_paragraph("Tabel 4.1 Matriks Checklist Pengumpulan Berkas (Submission Deliverables)")
    p_cap_t41.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_t41.paragraph_format.space_before = Pt(3)
    p_cap_t41.paragraph_format.space_after = Pt(8)
    p_cap_t41.runs[0].font.name = "Times New Roman"
    p_cap_t41.runs[0].font.size = Pt(9.5)
    p_cap_t41.runs[0].font.italic = True

    add_heading_2(doc, "DAFTAR PUSTAKA")
    refs = [
        "1. Pressman, R. S., & Maxim, B. R. (2020). Software Engineering: A Practitioner's Approach (9th ed.). New York: McGraw-Hill Education.",
        "2. Sommerville, I. (2016). Software Engineering (10th ed.). Boston: Pearson Education.",
        "3. IEEE Computer Society. (1998). IEEE Std 830-1998: IEEE Recommended Practice for Software Requirements Specifications. New York: IEEE.",
        "4. Fowler, M. (2018). UML Distilled: A Brief Guide to the Standard Object Modeling Language (3rd ed.). Boston: Addison-Wesley Professional.",
        "5. Martin, R. C. (2017). Clean Architecture: A Craftsman's Guide to Software Structure and Design. Boston: Prentice Hall.",
        "6. NVIDIA Corporation. (2025). NVIDIA NIM: Inference Microservices Documentation for Enterprise AI Models. Santa Clara: NVIDIA Developer Documentation."
    ]
    for r in refs:
        p_ref = doc.add_paragraph(r)
        p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_ref.paragraph_format.left_indent = Inches(0.4)
        p_ref.paragraph_format.first_line_indent = Inches(-0.4)
        p_ref.paragraph_format.space_after = Pt(4)
        p_ref.paragraph_format.line_spacing = 1.15
        p_ref.runs[0].font.name = "Times New Roman"
        p_ref.runs[0].font.size = Pt(10)

    doc.save(docx_path)
    print(f"[SUCCESS] Document generated successfully: {docx_path}")

def extract_heading_pages_from_pdf(pdf_path):
    reader = pypdf.PdfReader(pdf_path)
    headings = {
        "b1": "BAB I: PRODUCT REQUIREMENTS DOCUMENT",
        "b1_1": "1.1 Konsep Proyek",
        "b1_2": "1.2 Latar Belakang Masalah",
        "b1_3": "1.3 Identifikasi Masalah",
        "b1_4": "1.4 Visi Produk",
        "b1_5": "1.5 Analisis Kebutuhan",
        "b1_6": "1.6 Spesifikasi Kebutuhan Fungsional",
        "b1_7": "1.7 Spesifikasi Kebutuhan Non-Fungsional",
        "b1_8": "1.8 Arsitektur & Spesifikasi Fitur Kecerdasan",
        "b2": "BAB II: SOFTWARE DESIGN DOCUMENT",
        "b2_1": "2.1 Arsitektur Perangkat Lunak",
        "b2_2": "2.2 Pemodelan 4 Diagram UML",
        "b2_2_1": "2.2.1 Use Case Diagram",
        "b2_2_2": "2.2.2 Activity Diagram",
        "b2_2_3": "2.2.3 Class Diagram",
        "b2_2_4": "2.2.4 Sequence Diagram",
        "b2_3": "2.3 Perancangan Basis Data",
        "b2_4": "2.4 Perancangan Antarmuka Pengguna",
        "b3": "BAB III: IMPLEMENTASI KODE",
        "b3_1": "3.1 Struktur Direktori",
        "b3_2": "3.2 Implementasi Fitur Kunci",
        "b3_3": "3.3 Laporan Pengujian",
        "b3_4": "3.4 Otomasi CI/CD Pipeline",
        "b3_5": "3.5 Petunjuk Instalasi",
        "b4": "BAB IV: KESIMPULAN & MATRIKS DELIVERABLES",
        "b4_1": "4.1 Kesimpulan Proyek",
        "b4_2": "4.2 Refleksi Pembelajaran Mandiri",
        "b4_3": "4.3 Matriks Checklist Pengumpulan",
        "dp": "DAFTAR PUSTAKA"
    }

    found = {
        "lp": "2",
        "kp": "3",
        "di": "4"
    }

    # Search starting from page 5 (index 4)
    for idx in range(4, len(reader.pages)):
        page_num = str(idx + 1)
        txt = reader.pages[idx].extract_text()
        for key, search_str in headings.items():
            if key not in found and search_str.lower() in txt.lower():
                found[key] = page_num

    return found

if __name__ == "__main__":
    out_docx = r"c:\xampp\htdocs\absen\Laporan_Proyek_Akhir_RPL_Sistem_Presensi_SMK_NU_Wargabinangun.docx"
    out_pdf = r"c:\xampp\htdocs\absen\Laporan_Proyek_Akhir_RPL_Sistem_Presensi_SMK_NU_Wargabinangun.pdf"

    print("--- PASS 1: Generate Document with All 18 Detailed Images & Logo ---")
    build_full_report_document(out_docx)
    print("--- PASS 1: Converting to PDF via docx2pdf ---")
    convert(out_docx, out_pdf)

    print("--- PASS 1: Extract Exact Page Numbers from PDF ---")
    exact_pages = extract_heading_pages_from_pdf(out_pdf)
    print("Exact pages detected:")
    for k, v in exact_pages.items():
        print(f"  {k}: Halaman {v}")

    print("--- PASS 2: Re-generate Document with 100% Synchronized TOC Pages ---")
    build_full_report_document(out_docx, p=exact_pages)
    print("--- PASS 2: Converting to Final Synchronized PDF ---")
    convert(out_docx, out_pdf)

    print("--- FINISHED: Master Report with All Images Successfully Synchronized! ---")
