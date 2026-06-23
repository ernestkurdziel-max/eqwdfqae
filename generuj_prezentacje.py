"""
Generator prezentacji:
"Opracowanie indywidualnego programu zabiegów fizykalnych
w ramach balneoklimatologii i odnowy biologicznej"

Przypadek kliniczny: Pacjent z zespołem cieśni nadgarstka.
Autor: Ernest Kurdziel, nr albumu 72714.

Uruchomienie:
    pip install python-pptx Pillow
    python generuj_prezentacje.py
"""

import os
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ----------------------------------------------------------------------------
# Paleta kolorów (motyw medyczny – turkus / granat)
# ----------------------------------------------------------------------------
NAVY = RGBColor(0x0F, 0x3D, 0x5C)        # ciemny granat – tytuły
TEAL = RGBColor(0x12, 0x9A, 0x9C)        # turkus – akcent
LIGHT_TEAL = RGBColor(0xE3, 0xF4, 0xF4)  # jasny turkus – tło bloków
GRAY = RGBColor(0x44, 0x4B, 0x52)        # tekst
LIGHT_GRAY = RGBColor(0xF2, 0xF5, 0xF7)  # tło slajdu
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT = RGBColor(0xF2, 0xA1, 0x3B)      # pomarańcz – wyróżnienie

IMG = "obrazy"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H

BLANK = prs.slide_layouts[6]


# ----------------------------------------------------------------------------
# Funkcje pomocnicze
# ----------------------------------------------------------------------------
def add_slide():
    return prs.slides.add_slide(BLANK)


def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def set_fill_opacity(shape, opacity_pct):
    """Ustawia przezroczystość wypełnienia (opacity_pct: 0–100 = krycie)."""
    spPr = shape._element.spPr
    solidFill = spPr.find(qn("a:solidFill"))
    srgbClr = solidFill.find(qn("a:srgbClr"))
    for a in srgbClr.findall(qn("a:alpha")):
        srgbClr.remove(a)
    alpha = srgbClr.makeelement(qn("a:alpha"),
                                {"val": str(int(opacity_pct * 1000))})
    srgbClr.append(alpha)


def add_rect(slide, x, y, w, h, color, opacity=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    if opacity is not None:
        set_fill_opacity(shape, opacity)
    return shape


def add_round_rect(slide, x, y, w, h, color, opacity=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    if opacity is not None:
        set_fill_opacity(shape, opacity)
    return shape


def add_text(slide, x, y, w, h, text, size=18, color=GRAY, bold=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False,
             line_spacing=1.0):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing:
        p.line_spacing = line_spacing
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.name = "Calibri"
    return tb


def add_bullets(slide, x, y, w, h, items, size=18, color=GRAY,
                bullet_color=TEAL, space_after=8, line_spacing=1.05):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for item in items:
        text, level = (item if isinstance(item, tuple) else (item, 0))
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.level = level
        p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = "•  " if level == 0 else "–  "
        r.font.size = Pt(size)
        r.font.bold = True
        r.font.color.rgb = bullet_color if level == 0 else ACCENT
        r.font.name = "Calibri"
        r2 = p.add_run()
        r2.text = text
        r2.font.size = Pt(size if level == 0 else size - 1)
        r2.font.color.rgb = color
        r2.font.name = "Calibri"
    return tb


def add_image_cover(slide, path, x, y, w, h, border=True):
    """Wstawia obraz wypełniający ramkę (jak background-size: cover)."""
    iw, ih = Image.open(path).size
    target = (w / h)
    ratio = iw / ih
    pic = slide.shapes.add_picture(path, Inches(x), Inches(y),
                                   Inches(w), Inches(h))
    if ratio > target:
        crop = (1 - target / ratio) / 2
        pic.crop_left = crop
        pic.crop_right = crop
    else:
        crop = (1 - ratio / target) / 2
        pic.crop_top = crop
        pic.crop_bottom = crop
    if border:
        pic.line.color.rgb = WHITE
        pic.line.width = Pt(2)
    pic.shadow.inherit = False
    return pic


def image_caption(slide, x, y, w, text):
    """Pasek podpisu na dole obrazu (półprzezroczysty granat)."""
    band = add_rect(slide, Inches(x), Inches(y - 0.55), Inches(w),
                    Inches(0.55), NAVY, opacity=78)
    add_text(slide, Inches(x + 0.12), Inches(y - 0.5), Inches(w - 0.24),
             Inches(0.45), text, size=11, color=WHITE, bold=True,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)


def content_header(slide, number, title, subtitle=None):
    set_bg(slide, LIGHT_GRAY)
    add_rect(slide, 0, 0, Inches(0.22), SLIDE_H, TEAL)
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.5), Inches(0.42),
                                  Inches(0.82), Inches(0.82))
    circ.fill.solid()
    circ.fill.fore_color.rgb = NAVY
    circ.line.fill.background()
    circ.shadow.inherit = False
    cp = circ.text_frame.paragraphs[0]
    cp.alignment = PP_ALIGN.CENTER
    cr = cp.add_run()
    cr.text = str(number)
    cr.font.size = Pt(28)
    cr.font.bold = True
    cr.font.color.rgb = WHITE
    cr.font.name = "Calibri"
    add_text(slide, Inches(1.5), Inches(0.38), Inches(11.4), Inches(0.7),
             title, size=29, color=NAVY, bold=True)
    if subtitle:
        add_text(slide, Inches(1.52), Inches(0.98), Inches(11.4), Inches(0.4),
                 subtitle, size=14, color=TEAL, italic=True)
    add_rect(slide, Inches(1.5), Inches(1.42), Inches(11.3), Pt(2), TEAL)


def treatment_table(slide, x, y, w, header_color, rows, row_h=0.82):
    c1 = w * 0.33
    c2 = w - c1
    hh = 0.5
    add_rect(slide, Inches(x), Inches(y), Inches(c1), Inches(hh),
             header_color)
    add_rect(slide, Inches(x + c1), Inches(y), Inches(c2), Inches(hh),
             header_color)
    add_text(slide, Inches(x + 0.12), Inches(y), Inches(c1 - 0.2),
             Inches(hh), "Zabieg", 13, WHITE, True,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(x + c1 + 0.12), Inches(y), Inches(c2 - 0.2),
             Inches(hh), "Działanie i cel terapeutyczny", 13, WHITE, True,
             anchor=MSO_ANCHOR.MIDDLE)
    yy = y + hh
    for i, (z, d) in enumerate(rows):
        bg = WHITE if i % 2 == 0 else LIGHT_TEAL
        add_rect(slide, Inches(x), Inches(yy), Inches(w), Inches(row_h), bg)
        add_text(slide, Inches(x + 0.12), Inches(yy), Inches(c1 - 0.2),
                 Inches(row_h), z, 11.5, NAVY, True, anchor=MSO_ANCHOR.MIDDLE,
                 line_spacing=1.0)
        add_text(slide, Inches(x + c1 + 0.12), Inches(yy), Inches(c2 - 0.24),
                 Inches(row_h), d, 10.5, GRAY, anchor=MSO_ANCHOR.MIDDLE,
                 line_spacing=1.0)
        yy += row_h


# ============================================================================
# SLAJD 1 – Tytułowy (tło fotograficzne)
# ============================================================================
s = add_slide()
set_bg(s, NAVY)
add_image_cover(s, os.path.join(IMG, "tlo_tytul.png"), 0, 0, 13.333, 7.5,
                border=False)
# przyciemnienie po lewej dla czytelności tekstu
add_rect(s, 0, 0, Inches(7.8), SLIDE_H, NAVY, opacity=62)
add_rect(s, 0, Inches(3.05), Inches(7.8), Inches(0.05), TEAL, opacity=90)

add_text(s, Inches(0.7), Inches(0.55), Inches(6.9), Inches(0.4),
         "BALNEOKLIMATOLOGIA I ODNOWA BIOLOGICZNA", size=14, color=TEAL,
         bold=True)
add_text(s, Inches(0.68), Inches(1.35), Inches(7.0), Inches(1.6),
         "Indywidualny program\nzabiegów fizykalnych", size=37, color=WHITE,
         bold=True, line_spacing=1.0)
add_text(s, Inches(0.7), Inches(3.2), Inches(7.0), Inches(0.9),
         "Przypadek kliniczny: pacjent z zespołem cieśni nadgarstka",
         size=19, color=ACCENT, bold=True, line_spacing=1.05)
add_text(s, Inches(0.7), Inches(4.25), Inches(7.0), Inches(0.7),
         "Opracowanie indywidualnego programu w ramach balneoterapii "
         "i odnowy biologicznej", size=13, color=RGBColor(0xD9, 0xEA, 0xEA),
         italic=True, line_spacing=1.05)

# karta autora
add_round_rect(s, Inches(0.7), Inches(5.55), Inches(6.0), Inches(1.4),
               WHITE, opacity=92)
add_rect(s, Inches(0.7), Inches(5.55), Inches(0.13), Inches(1.4), ACCENT)
add_text(s, Inches(1.0), Inches(5.72), Inches(5.6), Inches(0.45),
         "Autor:  Ernest Kurdziel", size=18, color=NAVY, bold=True)
add_text(s, Inches(1.0), Inches(6.2), Inches(5.6), Inches(0.4),
         "Nr albumu:  72714", size=15, color=GRAY)
add_text(s, Inches(1.0), Inches(6.55), Inches(5.6), Inches(0.4),
         "Balneoklimatologia i odnowa biologiczna", size=12, color=TEAL,
         italic=True)


# ============================================================================
# SLAJD 2 – Plan prezentacji
# ============================================================================
s = add_slide()
content_header(s, "i", "Plan prezentacji")
left_items = [
    "1.  Charakterystyka pacjenta",
    "2.  Analiza wskazań i przeciwwskazań",
    "3.  Dobór zabiegów fizykalnych",
    "4.  Opracowanie programu terapii",
]
right_items = [
    "5.  Efekty terapeutyczne",
    "6.  Elementy odnowy biologicznej",
    "7.  Podsumowanie i wnioski",
    "8.  Bibliografia",
]
y0 = 2.0
for i, txt in enumerate(left_items):
    add_round_rect(s, Inches(1.5), Inches(y0 + i * 1.15), Inches(5.3),
                   Inches(0.9), WHITE)
    add_text(s, Inches(1.7), Inches(y0 + i * 1.15), Inches(5.0), Inches(0.9),
             txt, size=17, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
for i, txt in enumerate(right_items):
    add_round_rect(s, Inches(7.05), Inches(y0 + i * 1.15), Inches(5.3),
                   Inches(0.9), WHITE)
    add_text(s, Inches(7.25), Inches(y0 + i * 1.15), Inches(5.0), Inches(0.9),
             txt, size=17, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)


# ============================================================================
# SLAJD 3 – Charakterystyka pacjenta (+ anatomia)
# ============================================================================
s = add_slide()
content_header(s, 1, "Charakterystyka pacjenta",
               "Dane podstawowe, rozpoznanie i dolegliwości")

card = add_round_rect(s, Inches(0.55), Inches(1.7), Inches(3.95), Inches(5.0),
                      WHITE)
add_text(s, Inches(0.75), Inches(1.85), Inches(3.6), Inches(0.45),
         "Dane podstawowe", size=17, color=TEAL, bold=True)
data = [
    ("Wiek:", "39 lat"),
    ("Płeć:", "mężczyzna"),
    ("Zawód:", "informatyk"),
    ("Praca:", "ponad 8 h dziennie przy komputerze"),
    ("Tryb życia:", "siedzący, niska aktywność fizyczna"),
]
yy = 2.4
for label, val in data:
    add_text(s, Inches(0.75), Inches(yy), Inches(3.55), Inches(0.35),
             label, size=14, color=NAVY, bold=True)
    add_text(s, Inches(0.75), Inches(yy + 0.3), Inches(3.55), Inches(0.5),
             val, size=13, color=GRAY, line_spacing=1.0)
    yy += 0.84

add_text(s, Inches(4.75), Inches(1.7), Inches(4.55), Inches(0.4),
         "Rozpoznanie", size=17, color=TEAL, bold=True)
add_bullets(s, Inches(4.75), Inches(2.12), Inches(4.6), Inches(1.0), [
    "Zespół cieśni nadgarstka – ucisk nerwu pośrodkowego w kanale "
    "nadgarstka prawej ręki.",
], size=13.5)
add_text(s, Inches(4.75), Inches(3.25), Inches(4.55), Inches(0.4),
         "Dolegliwości i ograniczenia", size=17, color=TEAL, bold=True)
add_bullets(s, Inches(4.75), Inches(3.7), Inches(4.6), Inches(3.0), [
    "Drętwienie i mrowienie palców prawej dłoni.",
    "Nasilenie objawów w nocy i przy pracy na klawiaturze.",
    "Ból nadgarstka, osłabienie siły chwytu.",
    "Trudności w precyzyjnych czynnościach manualnych.",
    "Przeciążenie wynikające ze statycznej pozycji.",
], size=13, space_after=7)

add_image_cover(s, os.path.join(IMG, "anatomia.png"), 9.55, 1.7, 3.3, 5.0)
image_caption(s, 9.55, 6.7, 3.3, "Kanał nadgarstka – ucisk nerwu pośrodkowego")


# ============================================================================
# SLAJD 4 – Analiza wskazań
# ============================================================================
s = add_slide()
content_header(s, 2, "Analiza wskazań",
               "Wskazania do zastosowania zabiegów fizykalnych")
add_bullets(s, Inches(1.5), Inches(1.75), Inches(11.2), Inches(3.0), [
    "Zespół cieśni nadgarstka o łagodnym / umiarkowanym nasileniu "
    "(bez wskazań do leczenia operacyjnego).",
    "Dolegliwości bólowe i parestezje (drętwienie, mrowienie).",
    "Stan przeciążeniowy ścięgien i pochewek w okolicy nadgarstka.",
    "Osłabienie siły mięśniowej i obniżona sprawność manualna ręki.",
    "Przewlekłe napięcie mięśniowe obręczy barkowej i przedramienia.",
    "Profilaktyka wtórna – zapobieganie nawrotom i progresji objawów.",
], size=17, space_after=11)
add_round_rect(s, Inches(1.5), Inches(5.75), Inches(11.3), Inches(1.2),
               LIGHT_TEAL)
add_text(s, Inches(1.75), Inches(5.88), Inches(10.8), Inches(0.4),
         "Uzasadnienie wyboru metod", size=15, color=TEAL, bold=True)
add_text(s, Inches(1.75), Inches(6.26), Inches(10.8), Inches(0.6),
         "Metody fizykalne działają przeciwbólowo, przeciwzapalnie i "
         "przeciwobrzękowo, poprawiają ukrwienie oraz zmniejszają ucisk na "
         "nerw pośrodkowy – stanowią leczenie zachowawcze pierwszego wyboru.",
         size=13, color=GRAY, line_spacing=1.05)


# ============================================================================
# SLAJD 5 – Przeciwwskazania
# ============================================================================
s = add_slide()
content_header(s, 2, "Przeciwwskazania",
               "Przeciwwskazania ogólne i miejscowe")
add_round_rect(s, Inches(1.5), Inches(1.75), Inches(5.45), Inches(4.95),
               WHITE)
add_rect(s, Inches(1.5), Inches(1.75), Inches(5.45), Inches(0.6), NAVY)
add_text(s, Inches(1.7), Inches(1.75), Inches(5.1), Inches(0.6),
         "Przeciwwskazania ogólne", size=16, color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_bullets(s, Inches(1.75), Inches(2.55), Inches(5.0), Inches(4.0), [
    "Choroba nowotworowa (czynna).",
    "Stany gorączkowe i ostre infekcje.",
    "Niewydolność krążenia i ciężkie choroby serca.",
    "Skłonność do krwawień, leczenie przeciwkrzepliwe.",
    "Ciąża (dla wybranych zabiegów).",
    "Rozrusznik serca (elektro- i magnetoterapia).",
], size=14, space_after=10)
add_round_rect(s, Inches(7.15), Inches(1.75), Inches(5.45), Inches(4.95),
               WHITE)
add_rect(s, Inches(7.15), Inches(1.75), Inches(5.45), Inches(0.6), TEAL)
add_text(s, Inches(7.35), Inches(1.75), Inches(5.1), Inches(0.6),
         "Przeciwwskazania miejscowe", size=16, color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
add_bullets(s, Inches(7.4), Inches(2.55), Inches(5.0), Inches(4.0), [
    "Ostry stan zapalny i obrzęk okolicy nadgarstka.",
    "Uszkodzenia i zmiany skórne, rany, infekcje skóry.",
    "Świeży uraz, złamanie w obrębie ręki.",
    "Zaawansowany zanik mięśni kłębu kciuka "
    "(konsultacja chirurgiczna).",
    "Zakrzepica żył kończyny.",
], size=14, space_after=10)


# ============================================================================
# SLAJD 6 – Balneoterapia: kąpiele lecznicze (+ foto)
# ============================================================================
s = add_slide()
content_header(s, 3, "Dobór zabiegów – balneoterapia (1/4)",
               "Kąpiele lecznicze i wodolecznictwo")
rows = [
    ("Kąpiele solankowe",
     "Sole mineralne działają przeciwzapalnie i przeciwbólowo, zmniejszają "
     "obrzęk i ucisk na nerw pośrodkowy."),
    ("Kąpiele kwasowęglowe (CO₂)",
     "Rozszerzają naczynia, poprawiają mikrokrążenie i ukrwienie tkanek "
     "przedramienia."),
    ("Kąpiele perełkowe",
     "Delikatny masaż pęcherzykami powietrza – rozluźnienie mięśni "
     "i efekt relaksacyjny."),
    ("Kąpiel wirowa kończyny",
     "Masaż wirowy wody + ciepło: rozluźnienie, poprawa krążenia "
     "i trofiki tkanek."),
    ("Kąpiele naprzemienne",
     "Cieplno-zimne „gimnastyka naczyń” – poprawa mikrokrążenia, "
     "redukcja zastoju i obrzęku."),
    ("Masaż podwodny / hydromasaż",
     "Strumień wody pod ciśnieniem rozluźnia mięśnie i poprawia "
     "drenaż żylno-limfatyczny."),
]
treatment_table(s, 0.55, 1.7, 7.8, NAVY, rows)
add_image_cover(s, os.path.join(IMG, "balneoterapia.png"), 8.6, 1.7, 4.25,
                5.05)
image_caption(s, 8.6, 6.75, 4.25, "Kąpiel wirowa kończyny górnej")


# ============================================================================
# SLAJD 7 – Peloidoterapia, termoterapia i balneoklimatologia (+ foto)
# ============================================================================
s = add_slide()
content_header(s, 3, "Dobór zabiegów – peloidy i klimat (2/4)",
               "Peloidoterapia, termoterapia i balneoklimatologia")
rows = [
    ("„Rękawica” / okłady borowinowe",
     "Głębokie przegrzanie + substancje biologicznie czynne: działanie "
     "przeciwzapalne i regenerujące."),
    ("Fangoterapia (fango)",
     "Termoterapia mułem mineralnym – zmniejsza ból i napięcie tkanek "
     "okołostawowych."),
    ("Okłady parafinowe",
     "Równomierne, długotrwałe ciepło rozluźnia tkanki i zwiększa "
     "elastyczność ścięgien."),
    ("Inhalacje solankowe / tężnie",
     "Aerozol solankowy – działanie ogólnoustrojowe i regeneracyjne "
     "(element klimatoterapii)."),
    ("Subterranoterapia / groty solne",
     "Mikroklimat bogaty w jod i sole – działanie relaksacyjne "
     "i przeciwzapalne."),
    ("Klimatoterapia (helio-, aeroterapia)",
     "Słońce, świeże powietrze i ruch – poprawa odporności, nastroju "
     "i samopoczucia."),
]
treatment_table(s, 0.55, 1.7, 7.8, TEAL, rows)
add_image_cover(s, os.path.join(IMG, "borowina.png"), 8.6, 1.7, 4.25, 5.05)
image_caption(s, 8.6, 6.75, 4.25, "Okład borowinowy (peloidoterapia) ręki")


# ============================================================================
# SLAJD 8 – Fizykoterapia i kinezyterapia (+ foto)
# ============================================================================
s = add_slide()
content_header(s, 3, "Dobór zabiegów – fizykoterapia (3/4)",
               "Fizykoterapia, masaż i kinezyterapia")
rows = [
    ("Krioterapia miejscowa",
     "Schłodzenie tkanek – efekt przeciwbólowy i przeciwobrzękowy "
     "w fazie zaostrzenia."),
    ("Ultradźwięki / fonoforeza",
     "Mikromasaż tkankowy, działanie przeciwbólowe i przeciwzapalne."),
    ("Laseroterapia / magnetoterapia",
     "Stymulacja regeneracji nerwu, poprawa przewodnictwa, "
     "redukcja parestezji."),
    ("Jonoforeza",
     "Wprowadzanie leków przez skórę – miejscowe działanie "
     "przeciwzapalne i przeciwbólowe."),
    ("Masaż przedramienia i barku",
     "Rozluźnienie mięśni, poprawa krążenia i drenażu, redukcja "
     "przeciążeń wtórnych."),
    ("Kinezyterapia – ćwiczenia nerwu",
     "Mobilizacja nerwu i ścięgien, wzmacnianie, poprawa siły chwytu "
     "i sprawności ręki."),
]
treatment_table(s, 0.55, 1.7, 7.8, NAVY, rows)
add_image_cover(s, os.path.join(IMG, "masaz.png"), 8.6, 1.7, 4.25, 5.05)
image_caption(s, 8.6, 6.75, 4.25, "Masaż ręki i przedramienia")


# ============================================================================
# SLAJD 9 – Zabiegi odnowy biologicznej (+ foto)
# ============================================================================
s = add_slide()
content_header(s, 3, "Dobór zabiegów – odnowa biologiczna (4/4)",
               "Zabiegi regeneracyjne i relaksacyjne")
rows = [
    ("Sauna fińska / na podczerwień",
     "Przegrzanie ciała – rozluźnienie mięśni, detoksykacja "
     "i ogólna regeneracja."),
    ("Łaźnia parowa",
     "Wilgotne ciepło – poprawa krążenia, odprężenie i nawilżenie "
     "skóry."),
    ("Jacuzzi / hydromasaż całego ciała",
     "Masaż wodno-powietrzny – redukcja napięcia mięśniowego "
     "i stresu."),
    ("Masaż relaksacyjny / aromaterapeutyczny",
     "Rozluźnienie, poprawa nastroju, jakości snu i samopoczucia."),
    ("Krioterapia ogólnoustrojowa",
     "Krótka ekspozycja na zimno – działanie przeciwbólowe, "
     "przeciwzapalne i regeneracyjne."),
    ("Aromaterapia i muzykoterapia",
     "Wspomaganie relaksu i redukcja napięcia psychicznego."),
]
treatment_table(s, 0.55, 1.7, 7.8, TEAL, rows)
add_image_cover(s, os.path.join(IMG, "odnowa.png"), 8.6, 1.7, 4.25, 5.05)
image_caption(s, 8.6, 6.75, 4.25, "Sauna i strefa relaksu (odnowa biologiczna)")


# ============================================================================
# SLAJD 10 – Program terapii (plan tygodniowy)
# ============================================================================
s = add_slide()
content_header(s, 4, "Opracowanie programu terapii",
               "Przykładowy plan tygodniowy (cykl 3 tygodni / 15 zabiegów)")
days = [
    ("Pon.", "Kąpiel wirowa ręki (15 min)\nUltradźwięki (5 min)\n"
             "Ćwiczenia ślizgowe nerwu"),
    ("Wt.", "Okład borowinowy (20 min)\nMagnetoterapia (15 min)\n"
            "Sauna / relaks"),
    ("Śr.", "Kąpiel solankowa (15 min)\nMasaż przedramienia (15 min)\n"
            "Inhalacja solankowa"),
    ("Czw.", "Laseroterapia\nKinezyterapia – wzmacnianie\n"
             "Ćwiczenia rozciągające"),
    ("Pt.", "Kąpiele naprzemienne\nKrioterapia (wg potrzeby)\n"
            "Aromaterapia / relaksacja"),
]
n = len(days)
total_w = 11.3
gap = 0.2
cw = (total_w - gap * (n - 1)) / n
x = 1.5
y = 1.95
for d, content in days:
    add_rect(s, Inches(x), Inches(y), Inches(cw), Inches(0.55), NAVY)
    add_text(s, Inches(x), Inches(y), Inches(cw), Inches(0.55),
             d, size=15, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
    add_round_rect(s, Inches(x), Inches(y + 0.62), Inches(cw), Inches(2.95),
                   WHITE)
    tb = s.shapes.add_textbox(Inches(x + 0.05), Inches(y + 0.75),
                              Inches(cw - 0.1), Inches(2.7))
    tf = tb.text_frame
    tf.word_wrap = True
    for j, line in enumerate(content.split("\n")):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.space_after = Pt(7)
        r = p.add_run()
        r.text = line
        r.font.size = Pt(10.5)
        r.font.color.rgb = GRAY
        r.font.name = "Calibri"
    x += cw + gap

add_round_rect(s, Inches(1.5), Inches(5.75), Inches(11.3), Inches(1.2),
               LIGHT_TEAL)
add_text(s, Inches(1.75), Inches(5.85), Inches(10.8), Inches(0.4),
         "Zasady realizacji programu", size=14, color=TEAL, bold=True)
add_bullets(s, Inches(1.75), Inches(6.22), Inches(10.8), Inches(0.7), [
    "Częstotliwość: 5 dni w tygodniu; weekend – odpoczynek i autoterapia. "
    "Czas serii: ok. 3 tygodnie (15 zabiegów).",
    "Kolejność: zabiegi cieplne / rozluźniające → terapia właściwa → "
    "kinezyterapia. Intensywność dostosowana do reakcji pacjenta.",
], size=11.5, space_after=4)


# ============================================================================
# SLAJD 11 – Efekty terapeutyczne
# ============================================================================
s = add_slide()
content_header(s, 5, "Efekty terapeutyczne",
               "Przewidywane rezultaty zastosowanej terapii")
cards = [
    ("Zmniejszenie bólu", "Redukcja dolegliwości bólowych nadgarstka "
     "i parestezji (drętwienia, mrowienia) palców."),
    ("Lepsze ukrwienie", "Poprawa mikrokrążenia i trofiki tkanek, "
     "zmniejszenie obrzęku i ucisku na nerw."),
    ("Większa sprawność", "Wzrost siły chwytu i precyzji ruchów, "
     "poprawa zakresu ruchomości nadgarstka."),
    ("Rozluźnienie", "Zmniejszenie napięcia mięśni przedramienia "
     "i obręczy barkowej, ustąpienie przeciążeń."),
    ("Lepszy sen", "Ograniczenie nocnych dolegliwości bólowych "
     "poprawia jakość snu i regenerację."),
    ("Lepsze samopoczucie", "Ogólna regeneracja, redukcja stresu "
     "i poprawa komfortu pracy oraz życia."),
]
cw, ch = 3.6, 1.95
gx, gy = 0.25, 0.25
x0, y0 = 1.5, 1.95
for i, (t, d) in enumerate(cards):
    col, row = i % 3, i // 3
    x = x0 + col * (cw + gx)
    y = y0 + row * (ch + gy)
    add_round_rect(s, Inches(x), Inches(y), Inches(cw), Inches(ch), WHITE)
    add_rect(s, Inches(x), Inches(y), Inches(0.12), Inches(ch), ACCENT)
    add_text(s, Inches(x + 0.3), Inches(y + 0.15), Inches(cw - 0.4),
             Inches(0.5), t, size=16, color=NAVY, bold=True)
    add_text(s, Inches(x + 0.3), Inches(y + 0.65), Inches(cw - 0.45),
             Inches(1.2), d, size=12, color=GRAY, line_spacing=1.05)


# ============================================================================
# SLAJD 12 – Elementy odnowy biologicznej (styl życia, + foto)
# ============================================================================
s = add_slide()
content_header(s, 6, "Elementy odnowy biologicznej",
               "Dodatkowe formy regeneracji i ich znaczenie")
items = [
    ("Ergonomia i higiena pracy",
     "Prawidłowe stanowisko, podkładka pod nadgarstek, przerwy "
     "i mikroćwiczenia co 45–60 min."),
    ("Aktywność fizyczna",
     "Pływanie, spacery, ćwiczenia ogólnousprawniające – "
     "przeciwdziałanie skutkom pracy siedzącej."),
    ("Techniki relaksacyjne",
     "Trening relaksacyjny, ćwiczenia oddechowe i rozciągające – "
     "redukcja napięcia i stresu."),
    ("Dieta i nawodnienie",
     "Dieta przeciwzapalna, wit. z grupy B; nawodnienie wspiera "
     "regenerację tkanki nerwowej."),
    ("Higiena snu i orteza",
     "Regularny sen, orteza nocna stabilizująca nadgarstek "
     "w pozycji neutralnej."),
    ("Edukacja zdrowotna",
     "Świadomość czynników ryzyka i profilaktyka nawrotów – "
     "warunek trwałości efektów."),
]
cw, ch = 3.95, 1.45
gx, gy = 0.2, 0.2
x0, y0 = 0.55, 1.75
for i, (t, d) in enumerate(items):
    col, row = i % 2, i // 2
    x = x0 + col * (cw + gx)
    y = y0 + row * (ch + gy)
    add_round_rect(s, Inches(x), Inches(y), Inches(cw), Inches(ch), WHITE)
    num = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.16),
                             Inches(y + 0.27), Inches(0.55), Inches(0.55))
    num.fill.solid()
    num.fill.fore_color.rgb = TEAL
    num.line.fill.background()
    num.shadow.inherit = False
    npp = num.text_frame.paragraphs[0]
    npp.alignment = PP_ALIGN.CENTER
    nr = npp.add_run()
    nr.text = str(i + 1)
    nr.font.size = Pt(18)
    nr.font.bold = True
    nr.font.color.rgb = WHITE
    add_text(s, Inches(x + 0.85), Inches(y + 0.12), Inches(cw - 1.0),
             Inches(0.45), t, size=13.5, color=NAVY, bold=True)
    add_text(s, Inches(x + 0.85), Inches(y + 0.52), Inches(cw - 1.0),
             Inches(0.85), d, size=10.5, color=GRAY, line_spacing=1.0)

add_image_cover(s, os.path.join(IMG, "lifestyle.png"), 8.95, 1.75, 3.9, 4.95)
image_caption(s, 8.95, 6.7, 3.9, "Zdrowy styl życia i ergonomia pracy")


# ============================================================================
# SLAJD 13 – Podsumowanie i wnioski
# ============================================================================
s = add_slide()
content_header(s, 7, "Podsumowanie i wnioski",
               "Ocena skuteczności i możliwości modyfikacji terapii")
add_text(s, Inches(1.5), Inches(1.75), Inches(11.2), Inches(0.45),
         "Ocena skuteczności programu", size=18, color=TEAL, bold=True)
add_bullets(s, Inches(1.5), Inches(2.25), Inches(11.2), Inches(2.0), [
    "Kompleksowe leczenie zachowawcze (balneoterapia + fizykoterapia + "
    "kinezyterapia + odnowa biologiczna) skutecznie zmniejsza objawy "
    "łagodnego i umiarkowanego zespołu cieśni nadgarstka.",
    "Połączenie zabiegów miejscowych z modyfikacją stylu życia i ergonomii "
    "pracy daje trwalsze efekty niż samo leczenie objawowe.",
    "Wczesne wdrożenie terapii zmniejsza ryzyko progresji i potrzeby "
    "leczenia operacyjnego.",
], size=15, space_after=9)
add_text(s, Inches(1.5), Inches(5.0), Inches(11.2), Inches(0.45),
         "Możliwości modyfikacji terapii", size=18, color=TEAL, bold=True)
add_bullets(s, Inches(1.5), Inches(5.5), Inches(11.2), Inches(1.6), [
    "Program należy monitorować i dostosowywać do reakcji oraz postępów "
    "pacjenta (intensywność, dobór i liczba zabiegów).",
    "Brak poprawy lub nasilenie objawów / zanik mięśni → konsultacja "
    "neurologiczna i chirurgiczna (EMG, ewentualne leczenie operacyjne).",
], size=15, space_after=9)


# ============================================================================
# SLAJD 14 – Bibliografia
# ============================================================================
s = add_slide()
content_header(s, 8, "Bibliografia",
               "Literatura podstawowa i uzupełniająca")
add_text(s, Inches(1.5), Inches(1.75), Inches(11.2), Inches(0.4),
         "Literatura podstawowa", size=16, color=TEAL, bold=True)
add_bullets(s, Inches(1.5), Inches(2.2), Inches(11.2), Inches(2.4), [
    "Straburzyński G., Straburzyńska-Lupa A.: Medycyna fizykalna. "
    "PZWL, Warszawa 2000.",
    "Ponikowska I., Ferson D.: Nowoczesna Medycyna Uzdrowiskowa. "
    "Medi Press, Warszawa 2009.",
    "Kasprzak W.: Fizjoterapia kliniczna. PZWL, Warszawa 2011.",
    "Kasprzak W., Mańkowska A.: Fizykoterapia, medycyna uzdrowiskowa "
    "i SPA. PZWL, Warszawa 2010.",
], size=14, space_after=8)
add_text(s, Inches(1.5), Inches(4.7), Inches(11.2), Inches(0.4),
         "Literatura uzupełniająca", size=16, color=TEAL, bold=True)
add_bullets(s, Inches(1.5), Inches(5.15), Inches(11.2), Inches(1.5), [
    "Ponikowska J.: Kompendium balneologii: rekomendacje krajowego "
    "konsultanta. Toruń 2012.",
    "Kochański J. W.: Balneologia i hydroterapia. Wydawnictwo AWF, "
    "Wrocław 2002.",
], size=14, space_after=8)


# ============================================================================
# SLAJD 15 – Dziękuję (tło fotograficzne)
# ============================================================================
s = add_slide()
set_bg(s, NAVY)
add_image_cover(s, os.path.join(IMG, "tlo_koniec.png"), 0, 0, 13.333, 7.5,
                border=False)
add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY, opacity=45)
add_rect(s, Inches(4.16), Inches(4.25), Inches(5.0), Inches(0.05), TEAL)
add_text(s, Inches(0.8), Inches(2.85), Inches(11.7), Inches(1.0),
         "Dziękuję za uwagę", size=42, color=WHITE, bold=True,
         align=PP_ALIGN.CENTER)
add_text(s, Inches(0.8), Inches(4.4), Inches(11.7), Inches(0.6),
         "Indywidualny program zabiegów fizykalnych – "
         "zespół cieśni nadgarstka", size=16, color=RGBColor(0xD9, 0xEA, 0xEA),
         align=PP_ALIGN.CENTER, italic=True)
add_text(s, Inches(0.8), Inches(5.15), Inches(11.7), Inches(0.5),
         "Ernest Kurdziel  •  nr albumu 72714", size=14, color=ACCENT,
         align=PP_ALIGN.CENTER, bold=True)


# ----------------------------------------------------------------------------
out = "Zespol_ciesni_nadgarstka_prezentacja.pptx"
prs.save(out)
print(f"Zapisano: {out}")
print(f"Liczba slajdów: {len(prs.slides._sldIdLst)}")
