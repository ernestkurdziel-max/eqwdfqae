# -*- coding: utf-8 -*-
"""
Generator prezentacji: Mózgowe porażenie dziecięce – metody specjalne fizjoterapii.
Tworzy plik .pptx (16:9) z jednolitą, profesjonalną szatą graficzną.
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(BASE, "assets")

# --- Paleta kolorów ---
TEAL   = RGBColor(0x0E, 0x6B, 0x7A)
TEAL2  = RGBColor(0x14, 0x8A, 0x9C)
SKY    = RGBColor(0x4F, 0xB3, 0xC4)
MINT   = RGBColor(0xE4, 0xF4, 0xF2)
MINT2  = RGBColor(0xCF, 0xEA, 0xE7)
CORAL  = RGBColor(0xF0, 0x7B, 0x63)
DARK   = RGBColor(0x21, 0x3A, 0x40)
GRAY   = RGBColor(0x4B, 0x5C, 0x61)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

FONT   = "Calibri"
FONT_H = "Calibri"

EMU_IN = 914400
SW, SH = 13.333, 7.5

prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]


def img(name):
    return os.path.join(ASSETS, name)


def add_slide():
    return prs.slides.add_slide(BLANK)


def rect(slide, l, t, w, h, color, shape=MSO_SHAPE.RECTANGLE, line=None, line_w=0):
    sp = slide.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp


def textbox(slide, l, t, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    return tf


def set_run(r, text, size, color, bold=False, italic=False, font=FONT):
    r.text = text
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = font


def para(tf, first=False):
    return tf.paragraphs[0] if first and not tf.paragraphs[0].runs else tf.add_paragraph()


def bullet(tf, text, level=0, size=18, color=DARK, bold=False, glyph_color=CORAL,
           space_after=8, first=False, glyph=True):
    p = tf.paragraphs[0] if (first and not tf.paragraphs[0].runs) else tf.add_paragraph()
    p.space_after = Pt(space_after)
    p.space_before = Pt(0)
    p.line_spacing = 1.05
    if glyph:
        g = "▸  " if level == 0 else "–  "
        r0 = p.add_run()
        set_run(r0, g, size, glyph_color if level == 0 else SKY, bold=True)
    r = p.add_run()
    set_run(r, text, size, color, bold=bold)
    if level == 1:
        p.level = 1
    return p


def footer(slide, idx, note="Mózgowe porażenie dziecięce · Metody specjalne fizjoterapii"):
    rect(slide, 0, SH - 0.32, SW, 0.32, TEAL)
    tf = textbox(slide, 0.55, SH - 0.31, 9.5, 0.3, MSO_ANCHOR.MIDDLE)
    r = tf.paragraphs[0].add_run()
    set_run(r, note, 9, MINT)
    tf2 = textbox(slide, SW - 1.4, SH - 0.31, 0.85, 0.3, MSO_ANCHOR.MIDDLE)
    tf2.paragraphs[0].alignment = PP_ALIGN.RIGHT
    r2 = tf2.paragraphs[0].add_run()
    set_run(r2, str(idx), 10, MINT, bold=True)


def cropped_pic(slide, name, l, t, w, h, target_ratio=4/3, frame=True):
    """Wstaw obraz przycięty do zadanej proporcji, w ozdobnej ramce."""
    if frame:
        fr = rect(slide, l - 0.09, t - 0.09, w + 0.18, h + 0.18, TEAL,
                  shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        fr.adjustments[0] = 0.045
    pic = slide.shapes.add_picture(img(name), Inches(l), Inches(t), Inches(w), Inches(h))
    # źródło 3:2 (1.5). Przytnij boki do target_ratio.
    src = 1.5
    if target_ratio < src:
        crop = (1 - target_ratio / src) / 2
        pic.crop_left = crop
        pic.crop_right = crop
    return pic


def header(slide, kicker, title, idx):
    """Nagłówek slajdu treściowego."""
    rect(slide, 0, 0, SW, 1.32, TEAL)
    rect(slide, 0, 1.32, SW, 0.09, CORAL)
    # akcent boczny
    rect(slide, 0, 0, 0.22, 1.32, CORAL)
    if kicker:
        tf = textbox(slide, 0.55, 0.2, 12, 0.35)
        r = tf.paragraphs[0].add_run()
        set_run(r, kicker.upper(), 12, MINT, bold=True)
    tf2 = textbox(slide, 0.55, 0.52, 12.2, 0.75, MSO_ANCHOR.MIDDLE)
    r2 = tf2.paragraphs[0].add_run()
    set_run(r2, title, 30, WHITE, bold=True, font=FONT_H)
    footer(slide, idx)


# ============================================================
# SLAJD 1 — TYTUŁOWY
# ============================================================
s = add_slide()
s.shapes.add_picture(img("bg_title.png"), 0, 0, Inches(SW), Inches(SH))
# półprzezroczysty panel pod tekstem dla czytelności
panel = rect(s, 0.0, 2.3, 7.7, 3.25, WHITE, shape=MSO_SHAPE.RECTANGLE)
panel.fill.fore_color.rgb = WHITE
panel.fill.transparency = 0  # solid
# pasek akcentu
rect(s, 0.7, 2.62, 1.6, 0.12, CORAL)
tf = textbox(s, 0.7, 2.8, 7.2, 1.7)
r = tf.paragraphs[0].add_run()
set_run(r, "Mózgowe porażenie dziecięce", 40, TEAL, bold=True, font=FONT_H)
p = tf.add_paragraph()
p.space_before = Pt(6)
r = p.add_run()
set_run(r, "Metody specjalne fizjoterapii u pacjenta pediatrycznego", 20, GRAY, bold=False)
p2 = tf.add_paragraph()
p2.space_before = Pt(10)
r = p2.add_run()
set_run(r, "Objawy i problemy funkcjonalne oraz przegląd wybranych metod: "
           "NDT-Bobath · PNF · metoda Peto · Ruch Rozwijający W. Sherborne · "
           "integracja sensoryczna", 13, GRAY, italic=True)
# dolny pasek tematu
rect(s, 0, SH - 0.5, SW, 0.5, TEAL)
tf = textbox(s, 0.7, SH - 0.49, 12, 0.48, MSO_ANCHOR.MIDDLE)
r = tf.paragraphs[0].add_run()
set_run(r, "Przedmiot: Metody specjalne fizjoterapii   |   Praca projektowa — przegląd literatury", 12, MINT, bold=True)

# ============================================================
# SLAJD 2 — CEL I ZAKRES PRACY
# ============================================================
s = add_slide()
header(s, "Wprowadzenie", "Cel i zakres pracy", 2)
tf = textbox(s, 0.7, 1.75, 6.9, 5.0)
bullet(tf, "Przedstawienie mózgowego porażenia dziecięcego (MPD) jako najczęstszej "
           "przyczyny niepełnosprawności ruchowej wieku rozwojowego.", size=18, first=True)
bullet(tf, "Omówienie objawów oraz problemów funkcjonalnych dziecka z MPD.", size=18)
bullet(tf, "Przegląd literatury (książki i artykuły naukowe) dotyczącej metod "
           "usprawniania stosowanych w tej jednostce chorobowej.", size=18)
bullet(tf, "Charakterystyka wybranych metod specjalnych fizjoterapii:", size=18)
bullet(tf, "NDT-Bobath, PNF, metoda Peto (nauczanie kierowane),", level=1, size=16, space_after=3)
bullet(tf, "Ruch Rozwijający Weroniki Sherborne, integracja sensoryczna (SI).", level=1, size=16)
bullet(tf, "Podkreślenie znaczenia wczesnej, kompleksowej i zespołowej rehabilitacji.", size=18)
cropped_pic(s, "team.png", 8.05, 2.15, 4.6, 3.45)
# nota terminologiczna
note = rect(s, 8.05, 5.85, 4.6, 0.85, MINT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
note.adjustments[0] = 0.1
ntf = note.text_frame
ntf.word_wrap = True
ntf.margin_left = Inches(0.15); ntf.margin_right = Inches(0.15)
ntf.margin_top = Inches(0.08); ntf.margin_bottom = Inches(0.08)
r = ntf.paragraphs[0].add_run()
set_run(r, "MPD (mózgowe / dziecięce porażenie mózgowe, ang. cerebral palsy) — "
           "jednostka, w której stosuje się wszystkie omawiane metody.", 11, TEAL, italic=True)

# ============================================================
# SLAJD 3 — DEFINICJA
# ============================================================
s = add_slide()
header(s, "Charakterystyka choroby", "Czym jest mózgowe porażenie dziecięce?", 3)
tf = textbox(s, 0.7, 1.75, 6.9, 5.0)
bullet(tf, "Grupa trwałych zaburzeń rozwoju ruchu i postawy, powodujących "
           "ograniczenie aktywności.", size=18, first=True)
bullet(tf, "Przyczyną jest niepostępujące uszkodzenie lub zaburzenie rozwoju "
           "niedojrzałego mózgu (okres płodowy, okołoporodowy lub wczesnodziecięcy).", size=18)
bullet(tf, "Uszkodzenie ma charakter statyczny, lecz jego obraz kliniczny zmienia się "
           "wraz z dojrzewaniem dziecka.", size=18)
bullet(tf, "Zaburzeniom ruchowym często towarzyszą zaburzenia: czucia, percepcji, "
           "poznawcze, komunikacji, zachowania oraz padaczka.", size=18)
bullet(tf, "To zespół objawów (nie jedna choroba) o różnorodnym obrazie i nasileniu.", size=18)
cropped_pic(s, "brain.png", 8.05, 2.3, 4.6, 3.45)

# ============================================================
# SLAJD 4 — ETIOLOGIA
# ============================================================
s = add_slide()
header(s, "Charakterystyka choroby", "Etiologia — przyczyny i czynniki ryzyka", 4)
cards = [
    ("Czynniki przedporodowe", ["Infekcje wewnątrzmaciczne (TORCH)",
                                 "Niedotlenienie i niewydolność łożyska",
                                 "Wady rozwojowe OUN", "Wcześniactwo, ciąża mnoga",
                                 "Czynniki genetyczne"]),
    ("Czynniki okołoporodowe", ["Zamartwica i niedotlenienie",
                                 "Uraz okołoporodowy", "Niska masa urodzeniowa",
                                 "Krwawienia dokomorowe", "Ciężka żółtaczka (kernicterus)"]),
    ("Czynniki poporodowe", ["Zapalenie opon mózgowo-rdzeniowych",
                             "Urazy czaszkowo-mózgowe", "Niedotlenienie (np. utonięcie)",
                             "Udar, zaburzenia metaboliczne"]),
]
x = 0.7
for title, items in cards:
    card = rect(s, x, 1.9, 3.95, 4.35, MINT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    card.adjustments[0] = 0.06
    rect(s, x, 1.9, 3.95, 0.62, TEAL2, shape=MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.12
    ht = textbox(s, x + 0.15, 1.98, 3.65, 0.5, MSO_ANCHOR.MIDDLE)
    ht.paragraphs[0].alignment = PP_ALIGN.CENTER
    r = ht.paragraphs[0].add_run(); set_run(r, title, 15, WHITE, bold=True)
    bt = textbox(s, x + 0.28, 2.72, 3.5, 3.4)
    for i, it in enumerate(items):
        bullet(bt, it, size=13.5, space_after=7, first=(i == 0), glyph_color=CORAL)
    x += 4.13

# ============================================================
# SLAJD 5 — EPIDEMIOLOGIA
# ============================================================
s = add_slide()
header(s, "Charakterystyka choroby", "Epidemiologia", 5)
stats = [("2–3", "przypadki na 1000\nżywych urodzeń"),
         ("~1", "najczęstsza przyczyna\nniepełnosprawności ruchowej u dzieci"),
         ("40–60%", "dzieci z MPD ma\nniepełnosprawność intelektualną*")]
x = 0.7
for big, small in stats:
    card = rect(s, x, 1.95, 3.95, 2.1, MINT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    card.adjustments[0] = 0.08
    bt = textbox(s, x + 0.2, 2.1, 3.55, 1.8, MSO_ANCHOR.MIDDLE)
    bt.paragraphs[0].alignment = PP_ALIGN.CENTER
    r = bt.paragraphs[0].add_run(); set_run(r, big, 40, TEAL, bold=True)
    p = bt.add_paragraph(); p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); set_run(r, small, 13, GRAY)
    x += 4.13
tf = textbox(s, 0.7, 4.4, 11.9, 2.2)
bullet(tf, "Częstość występowania MPD utrzymuje się na stałym poziomie mimo postępu "
           "opieki perinatalnej — dzięki niej przeżywają dzieci coraz bardziej "
           "niedojrzałe i obciążone.", size=17, first=True)
bullet(tf, "Ryzyko rośnie znacząco u wcześniaków i noworodków z bardzo małą masą urodzeniową.", size=17)
bullet(tf, "Postacie i nasilenie objawów są bardzo zróżnicowane — od dyskretnych zaburzeń "
           "chodu po ciężką, wielonarządową niepełnosprawność.", size=17)
tf2 = textbox(s, 0.7, 6.7, 11.9, 0.3)
r = tf2.paragraphs[0].add_run()
set_run(r, "* Dane orientacyjne; częstość zaburzeń towarzyszących zależy od postaci i ciężkości MPD.", 10, GRAY, italic=True)

# ============================================================
# SLAJD 6 — KLASYFIKACJA
# ============================================================
s = add_slide()
header(s, "Charakterystyka choroby", "Klasyfikacja postaci klinicznych", 6)
rows = [
    ("Postać", "Charakterystyka", True),
    ("Spastyczna (najczęstsza)", "Wzmożone napięcie mięśni; diplegia, hemiplegia, tetraplegia (kwadriplegia)", False),
    ("Dyskinetyczna", "Ruchy mimowolne, zmienne napięcie — atetoza, dystonia, pląsawica", False),
    ("Ataktyczna", "Zaburzenia równowagi i koordynacji, drżenie zamiarowe", False),
    ("Mieszana", "Współwystępowanie cech kilku postaci jednocześnie", False),
]
top = 1.95
tbl_w = 8.9
gx = 0.7
rh = 0.72
for i, (a, b, hd) in enumerate(rows):
    y = top + i * rh
    col = TEAL if hd else (MINT if i % 2 else MINT2)
    rect(s, gx, y, 3.4, rh, col)
    rect(s, gx + 3.4, y, tbl_w - 3.4, rh, col)
    t1 = textbox(s, gx + 0.15, y, 3.15, rh, MSO_ANCHOR.MIDDLE)
    r = t1.paragraphs[0].add_run(); set_run(r, a, 14, WHITE if hd else TEAL, bold=True)
    t2 = textbox(s, gx + 3.55, y, tbl_w - 3.7, rh, MSO_ANCHOR.MIDDLE)
    r = t2.paragraphs[0].add_run(); set_run(r, b, 13, WHITE if hd else DARK, bold=hd)
# panel GMFCS
side = rect(s, 9.85, 1.95, 2.8, rh * 5, TEAL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
side.adjustments[0] = 0.05
st = textbox(s, 10.05, 2.15, 2.4, rh * 5 - 0.4)
r = st.paragraphs[0].add_run(); set_run(r, "GMFCS", 18, WHITE, bold=True)
p = st.add_paragraph(); p.space_before = Pt(6)
r = p.add_run(); set_run(r, "Skala funkcji motoryki dużej (5 poziomów) opisuje samodzielność w lokomocji.", 12, MINT)
p = st.add_paragraph(); p.space_before = Pt(8)
r = p.add_run(); set_run(r, "I – chodzi bez ograniczeń\nV – zależny, wózek", 12, MINT, italic=True)
tf = textbox(s, 0.7, 5.75, 8.9, 1.0)
bullet(tf, "Do oceny funkcji stosuje się także MACS (kończyny górne) oraz CFCS (komunikacja).",
       size=14, first=True)

# ============================================================
# SLAJD 7 — OBJAWY RUCHOWE
# ============================================================
s = add_slide()
header(s, "Objawy", "Objawy ruchowe (podstawowe)", 7)
tf = textbox(s, 0.7, 1.75, 6.9, 5.0)
bullet(tf, "Nieprawidłowe napięcie mięśniowe — spastyczność, wiotkość lub napięcie zmienne.", size=17, first=True)
bullet(tf, "Przetrwałe odruchy pierwotne i osłabione reakcje prostowania oraz równowagi.", size=17)
bullet(tf, "Opóźnienie rozwoju ruchowego — opóźnione osiąganie kamieni milowych.", size=17)
bullet(tf, "Patologiczne, stereotypowe wzorce ruchowe i postawy.", size=17)
bullet(tf, "Zaburzenia kontroli postawy, równowagi i koordynacji.", size=17)
bullet(tf, "Osłabienie siły mięśniowej i szybka męczliwość.", size=17)
bullet(tf, "Zaburzenia selektywnej kontroli ruchu (trudność w izolowanych ruchach).", size=17)
cropped_pic(s, "symptoms.png", 8.05, 2.3, 4.6, 3.45)

# ============================================================
# SLAJD 8 — OBJAWY TOWARZYSZĄCE
# ============================================================
s = add_slide()
header(s, "Objawy", "Zaburzenia współistniejące", 8)
tf = textbox(s, 0.7, 1.8, 5.9, 5.0)
bullet(tf, "Padaczka (nawet u ok. 30–50% dzieci).", size=17, first=True)
bullet(tf, "Niepełnosprawność intelektualna o różnym nasileniu.", size=17)
bullet(tf, "Zaburzenia wzroku (zez, niedowidzenie) i słuchu.", size=17)
bullet(tf, "Zaburzenia mowy i komunikacji (dyzartria).", size=17)
bullet(tf, "Zaburzenia karmienia, połykania, refluks, ślinotok.", size=17)
tf2 = textbox(s, 6.8, 1.8, 5.85, 5.0)
bullet(tf2, "Zaburzenia czucia i integracji sensorycznej.", size=17, first=True)
bullet(tf2, "Zaburzenia emocjonalne i zachowania.", size=17)
bullet(tf2, "Przewlekły ból (np. mięśniowo-szkieletowy).", size=17)
bullet(tf2, "Zaburzenia oddychania i częste infekcje dróg oddechowych.", size=17)
bullet(tf2, "Osteoporoza, zaburzenia odżywiania i wzrastania.", size=17)
info = rect(s, 6.8, 5.15, 5.85, 1.3, MINT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
info.adjustments[0] = 0.06
itf = info.text_frame; itf.word_wrap = True
itf.margin_left = Inches(0.2); itf.margin_right = Inches(0.2)
itf.margin_top = Inches(0.12)
r = itf.paragraphs[0].add_run()
set_run(r, "Wielość zaburzeń sprawia, że MPD wymaga opieki interdyscyplinarnej, "
           "a fizjoterapia jest tylko jednym z jej elementów.", 13, TEAL, italic=True)

# ============================================================
# SLAJD 9 — PROBLEMY FUNKCJONALNE
# ============================================================
s = add_slide()
header(s, "Problemy funkcjonalne", "Konsekwencje w codziennym funkcjonowaniu", 9)
tf = textbox(s, 0.7, 1.75, 6.9, 5.0)
bullet(tf, "Lokomocja: trudności lub brak samodzielnego chodu, potrzeba zaopatrzenia "
           "ortopedycznego (ortezy, balkonik, wózek).", size=16.5, first=True)
bullet(tf, "Samoobsługa: ograniczenia w ubieraniu, jedzeniu, higienie.", size=16.5)
bullet(tf, "Manipulacja: utrudniony chwyt i precyzyjne ruchy rąk.", size=16.5)
bullet(tf, "Komunikacja i nauka: trudności w mowie, pisaniu, edukacji szkolnej.", size=16.5)
bullet(tf, "Uczestnictwo: ograniczenie zabawy, kontaktów rówieśniczych i aktywności społecznej.", size=16.5)
bullet(tf, "Powikłania wtórne: przykurcze, zwichnięcie stawu biodrowego, skolioza, "
           "deformacje stóp.", size=16.5)
cropped_pic(s, "team.png", 8.05, 2.3, 4.6, 3.45)

# ============================================================
# SLAJD 10 — DIAGNOSTYKA
# ============================================================
s = add_slide()
header(s, "Postępowanie", "Diagnostyka i ocena funkcjonalna", 10)
tf = textbox(s, 0.7, 1.8, 5.9, 5.0)
bullet(tf, "Wywiad — przebieg ciąży, porodu i rozwoju dziecka.", size=17, first=True)
bullet(tf, "Badanie neurologiczne i ocena napięcia mięśniowego oraz odruchów.", size=17)
bullet(tf, "Ocena rozwoju psychoruchowego i wzorców ruchowych.", size=17)
bullet(tf, "Neuroobrazowanie — MRI/USG przezciemiączkowe mózgowia.", size=17)
bullet(tf, "Konsultacje: okulistyczna, laryngologiczna, logopedyczna, psychologiczna.", size=17)
tf2 = textbox(s, 6.8, 1.95, 5.85, 0.5)
r = tf2.paragraphs[0].add_run()
set_run(r, "Wybrane skale oceny funkcjonalnej:", 16, TEAL, bold=True)
scales = [("GMFCS", "poziom funkcji motoryki dużej"),
          ("GMFM", "pomiar funkcji motorycznych"),
          ("MACS", "funkcja manualna kończyn górnych"),
          ("Skala Ashworth", "ocena napięcia / spastyczności")]
y = 2.55
for a, b in scales:
    card = rect(s, 6.8, y, 5.85, 0.82, MINT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    card.adjustments[0] = 0.14
    ct = textbox(s, 7.0, y, 5.5, 0.82, MSO_ANCHOR.MIDDLE)
    r = ct.paragraphs[0].add_run(); set_run(r, a + " — ", 15, TEAL, bold=True)
    r = ct.paragraphs[0].add_run(); set_run(r, b, 14, DARK)
    y += 0.98

# ============================================================
# SLAJD 11 — ZNACZENIE FIZJOTERAPII
# ============================================================
s = add_slide()
header(s, "Fizjoterapia", "Znaczenie i zasady rehabilitacji", 11)
tf = textbox(s, 0.7, 1.75, 6.9, 5.0)
bullet(tf, "Wczesna interwencja — wykorzystanie neuroplastyczności rozwijającego się mózgu.", size=17, first=True)
bullet(tf, "Kompleksowość i ciągłość — terapia dostosowana do wieku i potrzeb dziecka.", size=17)
bullet(tf, "Podejście funkcjonalne — cele osadzone w codziennych czynnościach i zabawie.", size=17)
bullet(tf, "Zapobieganie powikłaniom wtórnym (przykurczom, deformacjom).", size=17)
bullet(tf, "Praca zespołu interdyscyplinarnego z aktywnym udziałem rodziny.", size=17)
bullet(tf, "Łączenie różnych metod — dobór indywidualny, a nie jedna „recepta”.", size=17)
cropped_pic(s, "team.png", 8.05, 2.3, 4.6, 3.45)

# ============================================================
# SLAJD 12 — PRZERYWNIK SEKCJI
# ============================================================
s = add_slide()
s.shapes.add_picture(img("bg_section.png"), 0, 0, Inches(SW), Inches(SH))
rect(s, 0.0, 3.0, 0.32, 1.6, CORAL)
tf = textbox(s, 0.75, 2.85, 6.5, 2.0, MSO_ANCHOR.MIDDLE)
r = tf.paragraphs[0].add_run()
set_run(r, "CZĘŚĆ II", 16, CORAL, bold=True)
p = tf.add_paragraph(); p.space_before = Pt(6)
r = p.add_run()
set_run(r, "Wybrane metody specjalne\nfizjoterapii w MPD", 34, TEAL, bold=True, font=FONT_H)
p = tf.add_paragraph(); p.space_before = Pt(10)
r = p.add_run()
set_run(r, "NDT-Bobath · PNF · metoda Peto · Ruch Rozwijający W. Sherborne · integracja sensoryczna",
        15, GRAY, italic=True)
footer(s, 12)


# ---- Wspólny szablon slajdu metody ----
def method_slide(idx, kicker, title, image, lead, points, note=None):
    s = add_slide()
    header(s, kicker, title, idx)
    # panel wiodący
    lp = rect(s, 0.7, 1.7, 6.95, 0.95, MINT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    lp.adjustments[0] = 0.08
    lt = textbox(s, 0.92, 1.72, 6.55, 0.9, MSO_ANCHOR.MIDDLE)
    r = lt.paragraphs[0].add_run(); set_run(r, lead, 14, TEAL, bold=True, italic=True)
    tf = textbox(s, 0.7, 2.85, 6.95, 3.9)
    for i, (txt, lvl) in enumerate(points):
        bullet(tf, txt, level=lvl, size=15.5 if lvl == 0 else 14,
               space_after=6 if lvl == 0 else 3, first=(i == 0))
    cropped_pic(s, image, 8.05, 2.35, 4.6, 3.45)
    if note:
        nb = rect(s, 8.05, 5.95, 4.6, 0.78, MINT2, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        nb.adjustments[0] = 0.12
        ntf = nb.text_frame; ntf.word_wrap = True
        ntf.margin_left = Inches(0.15); ntf.margin_right = Inches(0.15)
        ntf.margin_top = Inches(0.06)
        r = ntf.paragraphs[0].add_run(); set_run(r, note, 11, TEAL, italic=True)
    return s

# SLAJD 13 — NDT-Bobath
method_slide(
    13, "Metoda", "NDT-Bobath (koncepcja neurorozwojowa)", "ndt_bobath.png",
    "Twórcy: Berta i Karel Bobath. Podstawa: neurofizjologiczna koncepcja rozwoju.",
    [("Hamowanie nieprawidłowych wzorców ruchu i postawy.", 0),
     ("Normalizacja napięcia mięśniowego (obniżanie lub podwyższanie).", 0),
     ("Torowanie i utrwalanie prawidłowych wzorców ruchowych.", 0),
     ("Prowadzenie ruchu z tzw. punktów kluczowych kontroli.", 0),
     ("Włączanie prawidłowych wzorców w codzienne czynności — także w pielęgnację "
      "prowadzoną przez rodziców (handling).", 0),
     ("Terapia całościowa, dostosowana indywidualnie do dziecka.", 0)],
    note="Najczęściej stosowana metoda u niemowląt i dzieci z MPD.")

# SLAJD 14 — PNF
method_slide(
    14, "Metoda", "PNF — torowanie nerwowo-mięśniowe", "pnf.png",
    "Proprioceptive Neuromuscular Facilitation — wykorzystanie proprioceptorów.",
    [("Ruch w naturalnych, trójpłaszczyznowych wzorcach (diagonalnych, spiralnych).", 0),
     ("Torowanie ruchu przez bodźce: dotyk, opór, rozciąganie, trakcję/aproksymację.", 0),
     ("Praca na częściach ciała silniejszych, by torować słabsze (irradiacja).", 0),
     ("Podejście funkcjonalne, pozytywne — bazuje na możliwościach pacjenta.", 0),
     ("Wykorzystywana m.in. w treningu chodu, kontroli tułowia i stabilizacji.", 0),
     ("Cele ustalane w porozumieniu z pacjentem i rodziną.", 0)],
    note="Bezbolesna, motywująca; oparta na mocnych stronach dziecka.")

# SLAJD 15 — Peto
method_slide(
    15, "Metoda", "Metoda Peto — nauczanie kierowane", "peto.png",
    "Twórca: András Petö. System łączący rehabilitację z edukacją (konduktywny).",
    [("Zajęcia grupowe prowadzone przez wykwalifikowanego terapeutę („dyrygenta”).", 0),
     ("Rytmiczna intencja — słowne zapowiadanie i rytmizowanie wykonywanych ruchów.", 0),
     ("Specjalne meble ze szczeblami (drabinki, plintusy) ułatwiające chwyt i pozycje.", 0),
     ("Stały, uporządkowany plan dnia i powtarzalne zadania.", 0),
     ("Cel: maksymalna samodzielność w czynnościach dnia codziennego.", 0),
     ("Rozwija wolę działania, motywację i samokontrolę dziecka.", 0)],
    note="Dla dzieci w wieku przedszkolnym i szkolnym.")

# SLAJD 16 — Sherborne
method_slide(
    16, "Metoda", "Ruch Rozwijający Weroniki Sherborne", "sherborne.png",
    "Metoda wspomagająca rozwój psychoruchowy oparta na ruchu, dotyku i relacji.",
    [("Rozwijanie świadomości własnego ciała i przestrzeni.", 0),
     ("Budowanie poczucia bezpieczeństwa, zaufania i relacji z drugą osobą.", 0),
     ("Ćwiczenia: „z”, „przeciwko” i „razem” z partnerem (rodzicem, terapeutą).", 0),
     ("Naturalne formy ruchu wywodzące się z dziecięcej zabawy (baraszkowania).", 0),
     ("Poprawa integracji sensomotorycznej i kontaktów społecznych.", 0),
     ("Metoda uzupełniająca — wspiera terapię prowadzoną innymi metodami.", 0)],
    note="Prowadzona indywidualnie lub w małych grupach.")

# SLAJD 17 — Integracja sensoryczna
method_slide(
    17, "Metoda", "Integracja sensoryczna (SI)", "sensory.png",
    "Twórczyni: A. Jean Ayres. Poprawa przetwarzania bodźców zmysłowych w OUN.",
    [("Stymulacja układów: przedsionkowego, proprioceptywnego i dotykowego.", 0),
     ("Dostarczanie kontrolowanej „diety sensorycznej” podczas ukierunkowanej zabawy.", 0),
     ("Poprawa kontroli postawy, równowagi, planowania ruchu (praksji).", 0),
     ("Specjalistyczny sprzęt: podwieszane huśtawki, platformy, deski równoważne.", 0),
     ("Wymaga diagnozy i prowadzenia przez certyfikowanego terapeutę SI.", 0),
     ("Metoda uzupełniająca kompleksową rehabilitację dziecka z MPD.", 0)],
    note="Wspiera lepsze reagowanie dziecka na otoczenie.")

# ============================================================
# SLAJD 18 — PORÓWNANIE METOD
# ============================================================
s = add_slide()
header(s, "Podsumowanie metod", "Zestawienie omawianych metod", 18)
comp = [
    ("Metoda", "Główny cel / charakter", True),
    ("NDT-Bobath", "Normalizacja napięcia i wzorców ruchowych; podejście neurorozwojowe", False),
    ("PNF", "Torowanie ruchu we wzorcach; praca na mocnych stronach", False),
    ("Metoda Peto", "Usamodzielnienie; połączenie terapii z edukacją (grupowo)", False),
    ("W. Sherborne", "Świadomość ciała, relacja i więź poprzez ruch", False),
    ("Integracja sensor. (SI)", "Poprawa przetwarzania bodźców zmysłowych", False),
]
top = 1.95; rh = 0.72; gx = 0.7; c1 = 3.6; c2 = 8.9
for i, (a, b, hd) in enumerate(comp):
    y = top + i * rh
    col = TEAL if hd else (MINT if i % 2 else MINT2)
    rect(s, gx, y, c1, rh, col)
    rect(s, gx + c1, y, c2, rh, col)
    t1 = textbox(s, gx + 0.15, y, c1 - 0.3, rh, MSO_ANCHOR.MIDDLE)
    r = t1.paragraphs[0].add_run(); set_run(r, a, 14, WHITE if hd else TEAL, bold=True)
    t2 = textbox(s, gx + c1 + 0.18, y, c2 - 0.35, rh, MSO_ANCHOR.MIDDLE)
    r = t2.paragraphs[0].add_run(); set_run(r, b, 13.5, WHITE if hd else DARK, bold=hd)
tf = textbox(s, 0.7, top + 6 * rh + 0.1, 11.9, 0.8)
bullet(tf, "Metody wzajemnie się uzupełniają — o wyborze decydują wiek, postać i nasilenie "
           "MPD oraz indywidualne cele terapii.", size=14, first=True)

# ============================================================
# SLAJD 19 — WNIOSKI
# ============================================================
s = add_slide()
header(s, "Podsumowanie", "Wnioski", 19)
tf = textbox(s, 0.7, 1.8, 11.9, 5.0)
bullet(tf, "MPD to najczęstsza przyczyna niepełnosprawności ruchowej wieku rozwojowego "
           "o bardzo zróżnicowanym obrazie klinicznym.", size=18, first=True)
bullet(tf, "Wczesne rozpoznanie i szybkie wdrożenie rehabilitacji poprawiają rokowanie "
           "funkcjonalne dziecka.", size=18)
bullet(tf, "Nie istnieje jedna uniwersalna metoda — skuteczna terapia łączy różne podejścia, "
           "dobrane indywidualnie.", size=18)
bullet(tf, "NDT-Bobath i PNF stanowią podstawę usprawniania neurorozwojowego, a metoda Peto, "
           "Ruch Rozwijający Sherborne i SI cennie je uzupełniają.", size=18)
bullet(tf, "Kluczowe znaczenie mają: kompleksowość, ciągłość, praca zespołowa oraz aktywne "
           "zaangażowanie rodziny.", size=18)
bullet(tf, "Cel nadrzędny: maksymalna samodzielność, uczestnictwo i jakość życia dziecka.", size=18)

# ============================================================
# SLAJD 20–21 — BIBLIOGRAFIA
# ============================================================
books = [
    "Michałowicz R. (red.): Mózgowe porażenie dziecięce. Wydawnictwo Lekarskie PZWL, Warszawa 2001.",
    "Levitt S.: Rehabilitacja w porażeniu mózgowym i zaburzeniach ruchu. Wydawnictwo Lekarskie PZWL, Warszawa 2007.",
    "Matyja M., Domagalska M.: Podstawy usprawniania neurorozwojowego według Berty i Karela Bobathów. Wyd. AWF, Katowice 2005.",
    "Borkowska M., Szwiling Z.: Metoda NDT-Bobath. Poradnik dla rodziców. Wydawnictwo Lekarskie PZWL, Warszawa 2011.",
    "Sherborne W.: Ruch rozwijający dla dzieci. Wydawnictwo Naukowe PWN, Warszawa 2012.",
]
books2 = [
    "Adler S.S., Beckers D., Buck M.: PNF w praktyce. Ilustrowany przewodnik. DB Publishing, 2014.",
    "Ayres A.J.: Dziecko a integracja sensoryczna. Wydawnictwo Harmonia, Gdańsk 2015.",
]
articles = [
    "Kułak W., Sobaniec W.: Mózgowe porażenie dziecięce — standardy postępowania. Standardy Medyczne 2004; 1(1): 96–99.",
    "Bagnowska K., Falkowski M.: Wybrane metody usprawniania dzieci z mózgowym porażeniem dziecięcym. Nowa Pediatria 2013; 3: 130–135.",
    "Kwolek A., Majka M., Pabis M.: Rehabilitacja dzieci z porażeniem mózgowym — problemy, aktualne kierunki. Ortopedia Traumatologia Rehabilitacja 2001; 3(4): 499–507.",
]


def biblio_item(tf, n, text, first=False):
    p = tf.paragraphs[0] if (first and not tf.paragraphs[0].runs) else tf.add_paragraph()
    p.space_after = Pt(10); p.line_spacing = 1.05
    r = p.add_run(); set_run(r, f"{n}.  ", 15, CORAL, bold=True)
    r = p.add_run(); set_run(r, text, 14.5, DARK)


# Slajd 20 — książki
s = add_slide()
header(s, "Bibliografia (1/2)", "Książki i podręczniki", 20)
tf = textbox(s, 0.75, 1.85, 11.85, 5.0)
lbl = tf.paragraphs[0].add_run(); set_run(lbl, "Pozycje książkowe", 16, TEAL, bold=True)
n = 1
for b in books + books2:
    biblio_item(tf, n, b)
    n += 1

# Slajd 21 — artykuły
s = add_slide()
header(s, "Bibliografia (2/2)", "Artykuły naukowe i źródła internetowe", 21)
tf = textbox(s, 0.75, 1.85, 11.85, 4.2)
lbl = tf.paragraphs[0].add_run()
set_run(lbl, "Artykuły naukowe (dostępne m.in. w bazach PubMed / Google Scholar)", 16, TEAL, bold=True)
for a in articles:
    biblio_item(tf, n, a)
    n += 1
note = rect(s, 0.75, 5.9, 11.85, 0.85, MINT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
note.adjustments[0] = 0.1
ntf = note.text_frame; ntf.word_wrap = True
ntf.margin_left = Inches(0.2); ntf.margin_right = Inches(0.2); ntf.margin_top = Inches(0.1)
r = ntf.paragraphs[0].add_run()
set_run(r, "Uwaga: przed cytowaniem w pracy zaleca się zweryfikowanie pełnych danych "
           "bibliograficznych (rok, tom, strony) w katalogu biblioteki oraz w bazach naukowych.",
        11.5, TEAL, italic=True)

# ============================================================
# SLAJD 22 — ZAKOŃCZENIE
# ============================================================
s = add_slide()
s.shapes.add_picture(img("bg_closing.png"), 0, 0, Inches(SW), Inches(SH))
tf = textbox(s, 1.5, 2.55, 10.3, 1.3, MSO_ANCHOR.MIDDLE)
tf.paragraphs[0].alignment = PP_ALIGN.CENTER
r = tf.paragraphs[0].add_run()
set_run(r, "Dziękuję za uwagę", 44, TEAL, bold=True, font=FONT_H)
tf2 = textbox(s, 2.0, 4.15, 9.3, 1.2, MSO_ANCHOR.TOP)
tf2.paragraphs[0].alignment = PP_ALIGN.CENTER
r = tf2.paragraphs[0].add_run()
set_run(r, "Wczesna, kompleksowa i zespołowa rehabilitacja daje dziecku z MPD "
           "szansę na większą samodzielność i lepszą jakość życia.", 16, GRAY, italic=True)

out = os.path.join(BASE, "Mozgowe_porazenie_dzieciece_metody_fizjoterapii.pptx")
prs.save(out)
print("Zapisano:", out)
print("Liczba slajdów:", len(prs.slides._sldIdLst))
