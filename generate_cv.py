#!/usr/bin/env python3
"""Generator profesjonalnego CV (.docx) dla Ernesta Kurdziela.

Nowoczesny, dwukolumnowy układ z ciemnym paskiem bocznym i akcentem
kolorystycznym. Uruchomienie:  python3 generate_cv.py
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------------------------------
# Paleta kolorów (nowoczesny, "zdrowotno-sportowy" motyw)
# ---------------------------------------------------------------------------
SIDEBAR_BG = "16424A"   # głęboki petrol / ciemny teal
ACCENT = "1DA398"       # żywy teal – kolor akcentu
DARK_TEXT = "23313A"    # ciemny grafit – tekst główny
GREY_TEXT = "5C6B73"    # szary – tekst pomocniczy
SIDEBAR_TEXT = "FFFFFF"
SIDEBAR_MUTED = "CFE3E1"
LINE_GREY = "D9DEE1"

FONT = "Calibri"

# ---------------------------------------------------------------------------
# Pomocnicze funkcje XML
# ---------------------------------------------------------------------------

def set_cell_background(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for tag, val in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = OxmlElement(f"w:{tag}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)


def remove_table_borders(table):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        borders.append(el)
    tbl_pr.append(borders)


def set_col_widths(table, widths):
    table.autofit = False
    table.allow_autofit = False
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width


def add_bottom_border(paragraph, color_hex, size=6):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color_hex)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def style_run(run, size=10.5, color=DARK_TEXT, bold=False, italic=False, caps=False, spacing=None, font=FONT):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    if caps:
        run.font.all_caps = True
    if spacing is not None:
        rpr = run._element.get_or_add_rPr()
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(spacing))
        rpr.append(sp)


def set_para_spacing(paragraph, before=0, after=0, line=None):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = line


# ---------------------------------------------------------------------------
# Sekcje paska bocznego
# ---------------------------------------------------------------------------

def sidebar_heading(cell, text):
    p = cell.add_paragraph()
    set_para_spacing(p, before=12, after=4)
    r = p.add_run(text)
    style_run(r, size=11, color="FFFFFF", bold=True, caps=True, spacing=30)
    add_bottom_border(p, ACCENT, size=8)


def sidebar_line(cell, text, muted=False, bold=False, before=0, after=2, size=10):
    p = cell.add_paragraph()
    set_para_spacing(p, before=before, after=after, line=1.05)
    r = p.add_run(text)
    style_run(r, size=size, color=(SIDEBAR_MUTED if muted else SIDEBAR_TEXT), bold=bold)
    return p


def sidebar_bullet(cell, text):
    p = cell.add_paragraph()
    set_para_spacing(p, before=0, after=3, line=1.05)
    p.paragraph_format.left_indent = Pt(10)
    p.paragraph_format.first_line_indent = Pt(-10)
    r = p.add_run("•  ")
    style_run(r, size=10, color=ACCENT, bold=True)
    r2 = p.add_run(text)
    style_run(r2, size=10, color=SIDEBAR_TEXT)


# ---------------------------------------------------------------------------
# Sekcje kolumny głównej
# ---------------------------------------------------------------------------

def main_heading(cell, text):
    p = cell.add_paragraph()
    set_para_spacing(p, before=4, after=1)
    r = p.add_run(text)
    style_run(r, size=11.5, color=SIDEBAR_BG, bold=True, caps=True, spacing=8)
    add_bottom_border(p, ACCENT, size=10)


def experience_entry(cell, title, org, period, bullets, place=None):
    p = cell.add_paragraph()
    set_para_spacing(p, before=2, after=0, line=1.0)
    r = p.add_run(title)
    style_run(r, size=10.5, color=DARK_TEXT, bold=True)

    p2 = cell.add_paragraph()
    set_para_spacing(p2, before=0, after=1, line=1.0)
    org_txt = org if not place else f"{org}  •  {place}"
    r2 = p2.add_run(org_txt)
    style_run(r2, size=9, color=ACCENT, bold=True)
    r3 = p2.add_run(f"    {period}")
    style_run(r3, size=8.5, color=GREY_TEXT, italic=True)

    for b in bullets:
        pb = cell.add_paragraph()
        set_para_spacing(pb, before=0, after=1, line=1.0)
        pb.paragraph_format.left_indent = Pt(12)
        pb.paragraph_format.first_line_indent = Pt(-12)
        rb = pb.add_run("–  ")
        style_run(rb, size=9, color=ACCENT, bold=True)
        rb2 = pb.add_run(b)
        style_run(rb2, size=9, color=DARK_TEXT)


def education_entry(cell, title, org, period):
    p = cell.add_paragraph()
    set_para_spacing(p, before=3, after=0, line=1.0)
    r = p.add_run(title)
    style_run(r, size=10.5, color=DARK_TEXT, bold=True)
    p2 = cell.add_paragraph()
    set_para_spacing(p2, before=0, after=1, line=1.0)
    r2 = p2.add_run(org)
    style_run(r2, size=9, color=ACCENT, bold=True)
    r3 = p2.add_run(f"    {period}")
    style_run(r3, size=8.5, color=GREY_TEXT, italic=True)


# ---------------------------------------------------------------------------
# Budowa dokumentu
# ---------------------------------------------------------------------------

def build():
    doc = Document()

    # Ustawienia strony (A4, wąskie marginesy)
    section = doc.sections[0]
    section.page_height = Mm(297)
    section.page_width = Mm(210)
    section.top_margin = Cm(0)
    section.bottom_margin = Cm(0)
    section.left_margin = Cm(0)
    section.right_margin = Cm(0)

    # Domyślna czcionka
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string(DARK_TEXT)

    # Tabela układu: 1 wiersz, 2 kolumny (pasek boczny + treść)
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    remove_table_borders(table)

    left = table.cell(0, 0)
    right = table.cell(0, 1)

    set_col_widths(table, [Cm(6.6), Cm(14.4)])
    set_cell_background(left, SIDEBAR_BG)
    set_cell_margins(left, top=220, start=340, bottom=100, end=340)
    set_cell_margins(right, top=220, start=400, bottom=100, end=400)

    # -----------------------------------------------------------------
    # PASEK BOCZNY
    # -----------------------------------------------------------------
    # Miejsce na zdjęcie (placeholder – zdjęcie zostanie dodane później)
    photo = left.paragraphs[0]
    photo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(photo, before=6, after=0)
    rp = photo.add_run("[ MIEJSCE NA ZDJĘCIE ]")
    style_run(rp, size=9, color=SIDEBAR_MUTED, italic=True)
    photo_note = left.add_paragraph()
    photo_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(photo_note, before=0, after=6)
    rpn = photo_note.add_run("(do uzupełnienia)")
    style_run(rpn, size=8, color=SIDEBAR_MUTED, italic=True)

    # KONTAKT
    sidebar_heading(left, "Kontakt")
    sidebar_line(left, "Telefon", muted=True, size=8.5, before=4, after=0)
    sidebar_line(left, "574 239 485", size=10.5, after=4)
    sidebar_line(left, "E-mail", muted=True, size=8.5, after=0)
    sidebar_line(left, "ernestkurdziel@gmail.com", size=10, after=4)
    sidebar_line(left, "Lokalizacja", muted=True, size=8.5, after=0)
    sidebar_line(left, "Cieszacin Wielki, k. Jarosławia", size=10, after=4)

    # UMIEJĘTNOŚCI
    sidebar_heading(left, "Umiejętności")
    for s in [
        "Praca z pacjentami i klientami",
        "Indywidualne, empatyczne podejście",
        "Komunikatywność i łatwość nawiązywania kontaktów",
        "Podstawy treningu i rehabilitacji",
        "Kinesiotaping",
        "Praca zespołowa i pod presją czasu",
        "Prawo jazdy kat. B",
    ]:
        sidebar_bullet(left, s)

    # JĘZYKI
    sidebar_heading(left, "Języki")
    sidebar_line(left, "Angielski", bold=True, size=10.5, before=4, after=0)
    sidebar_line(left, "poziom dobry (B2)", muted=True, size=9.5, after=4)

    # KURSY I CERTYFIKATY
    sidebar_heading(left, "Kursy i certyfikaty")
    for name, status in [
        ("Kurs kinesiotapingu", "ukończony"),
        ("Masaż tkanek głębokich", "w trakcie"),
        ("Kurs trenera personalnego", "w trakcie"),
    ]:
        p = left.add_paragraph()
        set_para_spacing(p, before=2, after=2, line=1.05)
        p.paragraph_format.left_indent = Pt(10)
        p.paragraph_format.first_line_indent = Pt(-10)
        rb = p.add_run("•  ")
        style_run(rb, size=10, color=ACCENT, bold=True)
        rn2 = p.add_run(name)
        style_run(rn2, size=10, color=SIDEBAR_TEXT, bold=True)
        rs = p.add_run(f"  ({status})")
        style_run(rs, size=9, color=SIDEBAR_MUTED, italic=True)

    # ZAINTERESOWANIA
    sidebar_heading(left, "Zainteresowania")
    for h in [
        "Siłownia i trening",
        "Fizjoterapia i rehabilitacja",
        "Muzyka",
        "Nowe technologie i sprzęt komputerowy",
    ]:
        sidebar_bullet(left, h)

    # -----------------------------------------------------------------
    # KOLUMNA GŁÓWNA
    # -----------------------------------------------------------------
    name_p = right.paragraphs[0]
    set_para_spacing(name_p, before=2, after=0)
    rn = name_p.add_run("ERNEST KURDZIEL")
    style_run(rn, size=21, color=SIDEBAR_BG, bold=True, spacing=20)

    title_p = right.add_paragraph()
    set_para_spacing(title_p, before=0, after=2)
    rt = title_p.add_run("Trener personalny  •  Fizjoterapeuta (w trakcie studiów)")
    style_run(rt, size=11, color=ACCENT, bold=True)

    # PODSUMOWANIE
    main_heading(right, "Podsumowanie zawodowe")
    summ = right.add_paragraph()
    set_para_spacing(summ, before=1, after=1, line=1.0)
    summ.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    rsum = summ.add_run(
        "Student fizjoterapii z pasją do sportu, treningu siłowego i rehabilitacji. Na co dzień pracuję "
        "jako przedstawiciel ds. merchandisingu, jednocześnie konsekwentnie rozwijając się w kierunku pracy "
        "z ciałem. Podczas praktyk kierunkowych wypracowałem indywidualne i empatyczne podejście do pacjenta. "
        "Poszukuję możliwości rozwoju jako trener personalny, łącząc wiedzę z fizjoterapii z treningiem."
    )
    style_run(rsum, size=9, color=DARK_TEXT)

    # DOŚWIADCZENIE
    main_heading(right, "Doświadczenie zawodowe")
    experience_entry(
        right,
        "Przedstawiciel ds. merchandisingu",
        "OEX Cursor",
        "2024 – obecnie",
        [
            "Dbanie o ekspozycję i dostępność produktów w punktach sprzedaży.",
            "Bieżący kontakt z klientami i obsługą punktów sprzedaży.",
        ],
    )
    experience_entry(
        right,
        "Praktyki zawodowe – fizjoterapia",
        "Praktyki kierunkowe (w ramach studiów)",
        "2024 – obecnie",
        [
            "Praca z pacjentami pod okiem doświadczonych specjalistów.",
            "Nauka technik terapeutycznych i indywidualnego podejścia do osób.",
        ],
    )
    experience_entry(
        right,
        "Żołnierz",
        "Wojsko Polskie",
        "2023 – 2024",
        [
            "Praca wymagająca dyscypliny, odpowiedzialności i pracy w zespole.",
            "Utrzymanie wysokiej sprawności i kondycji fizycznej.",
        ],
    )
    experience_entry(
        right,
        "Pracownik magazynu – praca sezonowa",
        "Kompletowanie zamówień, praca za granicą",
        "wakacje 2018 – 2022",
        [
            "Kompletowanie zamówień w międzynarodowym środowisku pracy.",
        ],
    )

    # WYKSZTAŁCENIE
    main_heading(right, "Wykształcenie")
    education_entry(
        right,
        "Fizjoterapia – studia (w toku)",
        "Wyższa Szkoła Informatyki i Zarządzania w Rzeszowie",
        "2024 – obecnie",
    )
    education_entry(
        right,
        "Technik informatyk",
        "Zespół Szkół Innowacyjnych w Jarosławiu",
        "2018 – 2022",
    )

    # KLAUZULA RODO (stopka)
    rodo = right.add_paragraph()
    set_para_spacing(rodo, before=3, after=0, line=0.95)
    rodo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    rr = rodo.add_run(
        "Wyrażam zgodę na przetwarzanie moich danych osobowych zawartych w niniejszym CV na potrzeby "
        "procesu rekrutacji, zgodnie z RODO (Rozporządzenie UE 2016/679 z dnia 27.04.2016 r.)."
    )
    style_run(rr, size=7, color=GREY_TEXT, italic=True)

    doc.save("CV_Ernest_Kurdziel.docx")
    print("Zapisano: CV_Ernest_Kurdziel.docx")


if __name__ == "__main__":
    build()
