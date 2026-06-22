# -*- coding: utf-8 -*-
"""
Generator prezentacji PowerPoint:
"Zasady planowania i programowania fizjoterapii pacjentow
z chorobami ukladu oddechowego" - na przykladzie mukowiscydozy.

Autor: Ernest Kurdziel (album 72714), prowadzacy: dr Weronika Cyganik.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn


# ---------------------------------------------------------------------------
# Paleta i ustawienia
# ---------------------------------------------------------------------------
PRIMARY = RGBColor(0x15, 0x40, 0x6B)     # gleboki granat
ACCENT = RGBColor(0x1C, 0xA9, 0xA6)      # turkus
ACCENT2 = RGBColor(0xE8, 0x8A, 0x3A)     # cieply akcent (pomarancz)
LIGHT = RGBColor(0xEE, 0xF4, 0xF8)       # jasne tlo
LIGHT2 = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x22, 0x30, 0x3C)        # tekst
GREY = RGBColor(0x5B, 0x6B, 0x79)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

FONT = 'Calibri'
FONT_H = 'Calibri'

SW, SH = Inches(13.333), Inches(7.5)

prs = Presentation()
prs.slide_width = SW
prs.slide_height = SH
BLANK = prs.slide_layouts[6]

_slide_no = 0


# ---------------------------------------------------------------------------
# Funkcje pomocnicze
# ---------------------------------------------------------------------------

def _set_font(run, size, bold=False, color=DARK, name=FONT, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = name
    run.font.color.rgb = color


def rect(slide, x, y, w, h, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp


def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    return tb, tf


def add_footer(slide, label):
    global _slide_no
    _slide_no += 1
    rect(slide, 0, Inches(7.18), SW, Inches(0.32), PRIMARY)
    tb, tf = textbox(slide, Inches(0.4), Inches(7.18), Inches(10.5), Inches(0.32),
                     anchor=MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = 'Fizjoterapia w mukowiscydozie  \u2022  Ernest Kurdziel'
    _set_font(r, 9, color=WHITE, name=FONT)
    tb2, tf2 = textbox(slide, Inches(12.3), Inches(7.18), Inches(0.9), Inches(0.32),
                       anchor=MSO_ANCHOR.MIDDLE)
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.RIGHT
    r2 = p2.add_run()
    r2.text = str(_slide_no)
    _set_font(r2, 10, bold=True, color=WHITE, name=FONT)


def content_header(slide, title, kicker=None):
    """Naglowek slajdu tresci: pasek + tytul + akcent."""
    rect(slide, 0, 0, SW, Inches(1.25), PRIMARY)
    rect(slide, 0, Inches(1.25), SW, Inches(0.08), ACCENT)
    rect(slide, 0, 0, Inches(0.18), Inches(1.25), ACCENT)
    tb, tf = textbox(slide, Inches(0.55), Inches(0.12), Inches(12.2), Inches(1.05),
                     anchor=MSO_ANCHOR.MIDDLE)
    if kicker:
        pk = tf.paragraphs[0]
        rk = pk.add_run()
        rk.text = kicker.upper()
        _set_font(rk, 12, bold=True, color=ACCENT, name=FONT_H)
        p = tf.add_paragraph()
    else:
        p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title
    _set_font(r, 28, bold=True, color=WHITE, name=FONT_H)


def body_frame(slide, top=Inches(1.6), left=Inches(0.7), width=Inches(11.93),
               height=Inches(5.3)):
    tb, tf = textbox(slide, left, top, width, height, anchor=MSO_ANCHOR.TOP)
    return tf


def add_bullet(tf, text, level=0, first=False, bold_lead=None, size=None,
               color=DARK, bullet_color=ACCENT):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.level = level
    p.space_after = Pt(8)
    p.space_before = Pt(2)
    p.line_spacing = 1.08
    if size is None:
        size = 20 if level == 0 else 17
    # marker
    mk = p.add_run()
    mk.text = ('\u25A0  ' if level == 0 else '\u2013  ')
    _set_font(mk, size, bold=True, color=(bullet_color if level == 0 else GREY),
              name=FONT)
    if bold_lead:
        rb = p.add_run()
        rb.text = bold_lead
        _set_font(rb, size, bold=True, color=PRIMARY, name=FONT)
    r = p.add_run()
    r.text = text
    _set_font(r, size, color=color, name=FONT)
    return p


def new_slide(bg=LIGHT):
    slide = prs.slides.add_slide(BLANK)
    rect(slide, 0, 0, SW, SH, bg)
    return slide


def content_slide(title, kicker=None, bg=LIGHT):
    slide = new_slide(bg)
    content_header(slide, title, kicker)
    add_footer(slide, title)
    return slide


# ---------------------------------------------------------------------------
# 1. SLAJD TYTULOWY
# ---------------------------------------------------------------------------

slide = prs.slides.add_slide(BLANK)
rect(slide, 0, 0, SW, SH, PRIMARY)
# dekoracyjne pasy
rect(slide, 0, 0, SW, Inches(0.25), ACCENT)
rect(slide, 0, Inches(7.25), SW, Inches(0.25), ACCENT)
rect(slide, Inches(0.0), Inches(2.55), SW, Inches(0.06), ACCENT)
rect(slide, Inches(0.0), Inches(4.95), SW, Inches(0.06), ACCENT)

# kierunek / uczelnia
tb, tf = textbox(slide, Inches(0.8), Inches(0.7), Inches(11.7), Inches(1.6))
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = 'Uniwersytet Rzeszowski  \u2022  Kierunek FIZJOTERAPIA'
_set_font(r, 18, bold=True, color=WHITE)
p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
r2 = p2.add_run(); r2.text = 'Studia niestacjonarne  \u2022  Zdrowie Publiczne'
_set_font(r2, 14, color=LIGHT)
p3 = tf.add_paragraph(); p3.alignment = PP_ALIGN.CENTER
r3 = p3.add_run(); r3.text = 'Projekt samodzielny'
_set_font(r3, 12, italic=True, color=ACCENT)

# tytul glowny
tb, tf = textbox(slide, Inches(0.9), Inches(2.8), Inches(11.5), Inches(2.0),
                 anchor=MSO_ANCHOR.MIDDLE)
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = 'Zasady planowania i programowania fizjoterapii pacjent\u00f3w z chorobami uk\u0142adu oddechowego'
_set_font(r, 32, bold=True, color=WHITE)
p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER; p2.space_before = Pt(10)
r2 = p2.add_run()
r2.text = 'na przyk\u0142adzie mukowiscydozy (zw\u0142\u00f3knienia torbielowatego)'
_set_font(r2, 20, italic=True, color=ACCENT)

# autor / prowadzacy
tb, tf = textbox(slide, Inches(0.8), Inches(5.25), Inches(11.7), Inches(1.8),
                 anchor=MSO_ANCHOR.MIDDLE)
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = 'Ernest Kurdziel'
_set_font(r, 22, bold=True, color=WHITE)
p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
r2 = p2.add_run(); r2.text = 'Numer albumu: 72714'
_set_font(r2, 14, color=LIGHT)
p3 = tf.add_paragraph(); p3.alignment = PP_ALIGN.CENTER; p3.space_before = Pt(8)
r3 = p3.add_run(); r3.text = 'Prowadz\u0105cy: dr Weronika Cyganik'
_set_font(r3, 16, color=LIGHT)
p4 = tf.add_paragraph(); p4.alignment = PP_ALIGN.CENTER; p4.space_before = Pt(6)
r4 = p4.add_run(); r4.text = 'Rzesz\u00f3w 2024'
_set_font(r4, 13, italic=True, color=ACCENT)


# ---------------------------------------------------------------------------
# 2. PLAN PREZENTACJI
# ---------------------------------------------------------------------------

slide = content_slide('Plan prezentacji', kicker='Agenda')
tf = body_frame(slide)
agenda = [
    'Wprowadzenie \u2013 choroby uk\u0142adu oddechowego i rola fizjoterapii',
    'Charakterystyka mukowiscydozy (etiologia, patofizjologia, objawy)',
    'Metody badania pacjenta z chorob\u0105 uk\u0142adu oddechowego',
    'Zasady planowania i programowania fizjoterapii',
    'Plan terapii \u2013 techniki, trening, przyk\u0142adowy harmonogram',
    'Ocena skuteczno\u015bci, podsumowanie i pi\u015bmiennictwo',
]
for i, t in enumerate(agenda):
    add_bullet(tf, t, first=(i == 0), size=22)


# ---------------------------------------------------------------------------
# 3. WSTEP
# ---------------------------------------------------------------------------

slide = content_slide('Wprowadzenie', kicker='Wst\u0119p')
tf = body_frame(slide)
add_bullet(tf, 'Choroby uk\u0142adu oddechowego to jedna z g\u0142\u00f3wnych przyczyn '
               'chorobowo\u015bci i \u015bmiertelno\u015bci na \u015bwiecie.', first=True)
add_bullet(tf, 'Obok farmakoterapii kluczow\u0105 rol\u0119 odgrywa fizjoterapia: utrzymanie '
               'dro\u017cno\u015bci dr\u00f3g oddechowych, poprawa wentylacji, wzrost tolerancji wysi\u0142ku.')
add_bullet(tf, 'Mukowiscydoza wybrana jako przyk\u0142ad \u2013 choroba przewlek\u0142a i post\u0119puj\u0105ca, '
               'w kt\u00f3rej fizjoterapia jest stosowana codziennie przez ca\u0142e \u017cycie.')
add_bullet(tf, 'Cel pracy: przedstawienie metod badania oraz zasad planowania '
               'i programowania fizjoterapii wraz z przyk\u0142adowym planem terapii.')


# ---------------------------------------------------------------------------
# 4. MUKOWISCYDOZA - DEFINICJA
# ---------------------------------------------------------------------------

slide = content_slide('Mukowiscydoza \u2013 definicja i etiologia',
                      kicker='Charakterystyka choroby')
tf = body_frame(slide)
add_bullet(tf, 'Najcz\u0119stsza choroba genetyczna dziedziczona autosomalnie '
               'recesywnie w populacji rasy bia\u0142ej.', first=True)
add_bullet(tf, 'Przyczyna: mutacje genu CFTR (chromosom 7); najcz\u0119stsza \u2013 F508del.')
add_bullet(tf, 'Bia\u0142ko CFTR = kana\u0142 chlorkowy; jego dysfunkcja zaburza transport '
               'jon\u00f3w Cl\u207b, Na\u207a i wody.')
add_bullet(tf, 'Skutek: produkcja g\u0119stego, lepkiego \u015bluzu.')
add_bullet(tf, 'Choroba wielonarz\u0105dowa: uk\u0142ad oddechowy, trzustka, przew\u00f3d pokarmowy, '
               'w\u0105troba, gruczo\u0142y potowe \u2013 o rokowaniu decyduje choroba p\u0142ucna.')


# ---------------------------------------------------------------------------
# 5. PATOFIZJOLOGIA
# ---------------------------------------------------------------------------

slide = content_slide('Patofizjologia zmian w uk\u0142adzie oddechowym',
                      kicker='Charakterystyka choroby')
tf = body_frame(slide)
add_bullet(tf, 'Odwodnienie warstwy p\u0142ynu okrywaj\u0105cego nab\u0142onek rz\u0119skowy.', first=True)
add_bullet(tf, 'Zag\u0119szczenie \u015bluzu \u2192 upo\u015bledzenie klirensu mukocyliarnego.')
add_bullet(tf, 'Zalegaj\u0105ca wydzielina = po\u017cywka dla bakterii '
               '(S. aureus, P. aeruginosa).')
add_bullet(tf, 'B\u0142\u0119dne ko\u0142o: infekcja \u2192 zapalenie \u2192 uszkodzenie \u015bciany oskrzeli.')
add_bullet(tf, 'Nast\u0119pstwa: rozstrzenie oskrzeli, w\u0142\u00f3knienie, post\u0119puj\u0105ca obturacja.')


# ---------------------------------------------------------------------------
# 6. OBJAWY
# ---------------------------------------------------------------------------

slide = content_slide('Objawy ze strony uk\u0142adu oddechowego',
                      kicker='Obraz kliniczny')
tf = body_frame(slide)
add_bullet(tf, 'Przewlek\u0142y, produktywny kaszel z g\u0119st\u0105 wydzielin\u0105.', first=True)
add_bullet(tf, 'Nawracaj\u0105ce i przewlek\u0142e zaka\u017cenia dolnych dr\u00f3g oddechowych.')
add_bullet(tf, 'Duszno\u015b\u0107 wysi\u0142kowa, a w zaawansowanych stadiach spoczynkowa.')
add_bullet(tf, '\u015awisty, furczenia i trzeszczenia w os\u0142uchiwaniu.')
add_bullet(tf, 'Spadek tolerancji wysi\u0142ku.')
add_bullet(tf, 'Stadia zaawansowane: palce pa\u0142eczkowate, beczkowata klatka piersiowa, '
               'przewlek\u0142a niewydolno\u015b\u0107 oddechowa.')


# ---------------------------------------------------------------------------
# 7. ZNACZENIE FIZJOTERAPII
# ---------------------------------------------------------------------------

slide = content_slide('Znaczenie fizjoterapii w mukowiscydozie',
                      kicker='Dlaczego to wa\u017cne')
tf = body_frame(slide, width=Inches(7.5))
add_bullet(tf, 'Techniki oczyszczania oskrzeli (ACT) \u2013 podstawowy, codzienny '
               'element leczenia.', first=True)
add_bullet(tf, 'Wspomagaj\u0105 usuwanie zalegaj\u0105cej wydzieliny.')
add_bullet(tf, 'Zmniejszaj\u0105 cz\u0119sto\u015b\u0107 zaostrze\u0144.')
add_bullet(tf, 'Poprawiaj\u0105 wentylacj\u0119 p\u0142uc.')
add_bullet(tf, 'W po\u0142\u0105czeniu z treningiem \u2013 lepsza wydolno\u015b\u0107 i jako\u015b\u0107 \u017cycia.')
# karta z haslem
card = rect(slide, Inches(8.5), Inches(2.0), Inches(4.2), Inches(3.6), PRIMARY)
tb, tcf = textbox(slide, Inches(8.7), Inches(2.2), Inches(3.8), Inches(3.2),
                  anchor=MSO_ANCHOR.MIDDLE)
pp = tcf.paragraphs[0]; pp.alignment = PP_ALIGN.CENTER
rr = pp.add_run(); rr.text = 'Fizjoterapia w CF'
_set_font(rr, 18, bold=True, color=ACCENT)
pp2 = tcf.add_paragraph(); pp2.alignment = PP_ALIGN.CENTER; pp2.space_before = Pt(10)
rr2 = pp2.add_run()
rr2.text = 'to nie dodatek do leczenia, lecz jego nieod\u0142\u0105czny, codzienny element \u2013 prowadzony przez ca\u0142e \u017cycie chorego.'
_set_font(rr2, 16, color=WHITE)


# ---------------------------------------------------------------------------
# 8. METODY BADANIA - WYWIAD
# ---------------------------------------------------------------------------

slide = content_slide('Metody badania \u2013 wywiad', kicker='Diagnostyka funkcjonalna')
tf = body_frame(slide)
add_bullet(tf, 'Dane og\u00f3lne: wiek, masa cia\u0142a, wzrost, BMI, aktywno\u015b\u0107 fizyczna.', first=True)
add_bullet(tf, 'Wywiad chorobowy: test potowy, badania genetyczne, zaostrzenia, hospitalizacje.')
add_bullet(tf, 'Objawy: charakter kaszlu, ilo\u015b\u0107 i rodzaj wydzieliny, duszno\u015b\u0107.')
add_bullet(tf, 'Leczenie: bronchodilatatory, mukolityki (dornaza alfa, hipertoniczny NaCl), '
               'antybiotyki, modulatory CFTR.')
add_bullet(tf, 'Dotychczasowa fizjoterapia i jej akceptacja przez pacjenta.')
add_bullet(tf, 'Jako\u015b\u0107 \u017cycia \u2013 kwestionariusz CFQ-R.')


# ---------------------------------------------------------------------------
# 9. BADANIE PRZEDMIOTOWE (4 skladowe - karty)
# ---------------------------------------------------------------------------

slide = content_slide('Badanie przedmiotowe klatki piersiowej',
                      kicker='Diagnostyka funkcjonalna')
cards = [
    ('Ogl\u0105danie', 'Budowa i ruchomo\u015b\u0107 klatki, tor i cz\u0119sto\u015b\u0107 oddech\u00f3w, sinica, '
                    'palce pa\u0142eczkowate, mi\u0119\u015bnie dodatkowe.'),
    ('Palpacja', 'Ruchomo\u015b\u0107 klatki piersiowej, dr\u017cenie g\u0142osowe, lokalizacja bolesno\u015bci.'),
    ('Opukiwanie', 'R\u00f3\u017cnicowanie odg\u0142osu (jawny, st\u0142umiony, b\u0119benkowy) \u2013 niedodma, rozedma.'),
    ('Os\u0142uchiwanie', 'Szmer oddechowy i zjawiska dodatkowe: \u015bwisty, furczenia, trzeszczenia.'),
]
cw, ch = Inches(5.9), Inches(2.45)
gap = Inches(0.25)
x0, y0 = Inches(0.7), Inches(1.7)
positions = [(x0, y0), (x0 + cw + gap, y0),
             (x0, y0 + ch + gap), (x0 + cw + gap, y0 + ch + gap)]
for (cx, cy), (h, body) in zip(positions, cards):
    rect(slide, cx, cy, cw, ch, WHITE, line=RGBColor(0xD5, 0xDE, 0xE6))
    rect(slide, cx, cy, Inches(0.12), ch, ACCENT)
    tb, ctf = textbox(slide, cx + Inches(0.3), cy + Inches(0.15), cw - Inches(0.45),
                      ch - Inches(0.3), anchor=MSO_ANCHOR.TOP)
    ph = ctf.paragraphs[0]
    rh = ph.add_run(); rh.text = h
    _set_font(rh, 18, bold=True, color=PRIMARY)
    pb = ctf.add_paragraph(); pb.space_before = Pt(6); pb.line_spacing = 1.05
    rbb = pb.add_run(); rbb.text = body
    _set_font(rbb, 14, color=DARK)


# ---------------------------------------------------------------------------
# 10. OCENA DUSZNOSCI
# ---------------------------------------------------------------------------

slide = content_slide('Ocena duszno\u015bci i wysi\u0142ku oddechowego',
                      kicker='Skale pomiarowe')
tf = body_frame(slide)
add_bullet(tf, 'Standaryzowane skale umo\u017cliwiaj\u0105 obiektywizacj\u0119 i monitorowanie '
               'post\u0119p\u00f3w terapii.', first=True)
add_bullet(tf, 'skala mMRC \u2013 duszno\u015b\u0107 w aktywno\u015bciach codziennych (0\u20134).',
           level=1, bold_lead='')
add_bullet(tf, 'skala Borga / Borg CR10 \u2013 subiektywna ocena duszno\u015bci i zm\u0119czenia.',
           level=1)
add_bullet(tf, 'wizualna skala analogowa (VAS) \u2013 duszno\u015b\u0107 lub bolesno\u015b\u0107.',
           level=1)


# ---------------------------------------------------------------------------
# 11. BADANIA CZYNNOSCIOWE
# ---------------------------------------------------------------------------

slide = content_slide('Badania czynno\u015bciowe uk\u0142adu oddechowego',
                      kicker='Diagnostyka funkcjonalna')
tf = body_frame(slide)
add_bullet(tf, 'Spirometria \u2013 obecno\u015b\u0107 i nasilenie obturacji (FEV1, FVC, FEV1/FVC); '
               'FEV1 to podstawowy parametr w CF.', first=True, bold_lead='')
add_bullet(tf, 'Pulsoksymetria \u2013 nieinwazyjny pomiar SpO2 w spoczynku i wysi\u0142ku.')
add_bullet(tf, 'Gazometria \u2013 PaO2, PaCO2, r\u00f3wnowaga kwasowo-zasadowa w niewydolno\u015bci '
               'oddechowej.')
add_bullet(tf, 'LCI (lung clearance index) \u2013 wczesny wska\u017anik nier\u00f3wnomierno\u015bci '
               'wentylacji.')


# ---------------------------------------------------------------------------
# 12. WYDOLNOSC FIZYCZNA
# ---------------------------------------------------------------------------

slide = content_slide('Ocena wydolno\u015bci fizycznej',
                      kicker='Diagnostyka funkcjonalna')
tf = body_frame(slide)
add_bullet(tf, '6MWT \u2013 test 6-minutowego marszu: dystans i reakcja (SpO2, t\u0119tno, '
               'duszno\u015b\u0107) na wysi\u0142ek submaksymalny.', first=True)
add_bullet(tf, 'CPET \u2013 sercowo-p\u0142ucny test wysi\u0142kowy: z\u0142oty standard, VO2peak.')
add_bullet(tf, 'Shuttle walk test oraz pr\u00f3by si\u0142y mi\u0119\u015bni oddechowych (PImax, PEmax).')
add_bullet(tf, 'Ocena postawy i narz\u0105du ruchu: kifoza piersiowa, protrakcja bark\u00f3w, '
               'ruchomo\u015b\u0107 i obwody klatki piersiowej.')


# ---------------------------------------------------------------------------
# 13. ZASADY PLANOWANIA
# ---------------------------------------------------------------------------

slide = content_slide('Zasady planowania i programowania',
                      kicker='Metodyka')
tf = body_frame(slide)
add_bullet(tf, 'Indywidualizacja \u2013 wiek, stadium, wydolno\u015b\u0107, preferencje pacjenta.',
           first=True)
add_bullet(tf, 'Systematyczno\u015b\u0107 \u2013 ACT codziennie, cz\u0119sto kilka razy dziennie.')
add_bullet(tf, 'Kompleksowo\u015b\u0107 \u2013 oczyszczanie + trening + inhalacje + edukacja.')
add_bullet(tf, 'Stopniowanie obci\u0105\u017ce\u0144 \u2013 zgodnie z tolerancj\u0105 pacjenta.')
add_bullet(tf, 'Wsp\u00f3\u0142praca interdyscyplinarna \u2013 lekarz, dietetyk, psycholog, '
               'piel\u0119gniarka.')
add_bullet(tf, 'Monitorowanie i modyfikacja \u2013 regularna ocena efekt\u00f3w.')


# ---------------------------------------------------------------------------
# 14. CELE TERAPII (dwie kolumny)
# ---------------------------------------------------------------------------

slide = content_slide('Cele fizjoterapii', kicker='Metodyka')
# kolumna 1
rect(slide, Inches(0.7), Inches(1.7), Inches(5.9), Inches(0.55), ACCENT)
tb, h1 = textbox(slide, Inches(0.85), Inches(1.72), Inches(5.6), Inches(0.5),
                 anchor=MSO_ANCHOR.MIDDLE)
r = h1.paragraphs[0].add_run(); r.text = 'Cele kr\u00f3tkoterminowe'
_set_font(r, 18, bold=True, color=WHITE)
tb, c1 = textbox(slide, Inches(0.8), Inches(2.45), Inches(5.8), Inches(4.4))
for i, t in enumerate([
    'U\u0142atwienie ewakuacji wydzieliny i poprawa dro\u017cno\u015bci dr\u00f3g oddechowych.',
    'Zmniejszenie duszno\u015bci i pracy oddechowej.',
    'Poprawa wentylacji i ruchomo\u015bci klatki piersiowej.',
]):
    add_bullet(c1, t, first=(i == 0), size=17)
# kolumna 2
rect(slide, Inches(6.83), Inches(1.7), Inches(5.8), Inches(0.55), PRIMARY)
tb, h2 = textbox(slide, Inches(6.98), Inches(1.72), Inches(5.5), Inches(0.5),
                 anchor=MSO_ANCHOR.MIDDLE)
r = h2.paragraphs[0].add_run(); r.text = 'Cele d\u0142ugoterminowe'
_set_font(r, 18, bold=True, color=WHITE)
tb, c2 = textbox(slide, Inches(6.93), Inches(2.45), Inches(5.7), Inches(4.4))
for i, t in enumerate([
    'Spowolnienie progresji zmian p\u0142ucnych i mniej zaostrze\u0144.',
    'Utrzymanie / poprawa wydolno\u015bci i tolerancji wysi\u0142ku.',
    'Utrwalenie nawyku samodzielnej, codziennej terapii.',
    'Poprawa jako\u015bci \u017cycia i postawy cia\u0142a.',
]):
    add_bullet(c2, t, first=(i == 0), size=17, bullet_color=PRIMARY)


# ---------------------------------------------------------------------------
# 15. TECHNIKI OCZYSZCZANIA (ACT)
# ---------------------------------------------------------------------------

slide = content_slide('Techniki oczyszczania drzewa oskrzelowego (ACT)',
                      kicker='Plan terapii')
tf = body_frame(slide)
add_bullet(tf, 'Drena\u017c u\u0142o\u017ceniowy \u2013 wykorzystanie grawitacji, cz\u0119sto z oklepywaniem '
               'i wibracjami.', first=True, bold_lead='')
add_bullet(tf, 'ACBT \u2013 cykl aktywnego oddychania (kontrola oddechu, \u0107wiczenia '
               'rozpr\u0119\u017caj\u0105ce, huffing).')
add_bullet(tf, 'Drena\u017c autogeniczny (AD) \u2013 oddychanie na r\u00f3\u017cnych obj\u0119to\u015bciach p\u0142uc.')
add_bullet(tf, 'PEP \u2013 terapia dodatnim ci\u015bnieniem wydechowym (maska / ustnik).')
add_bullet(tf, 'Oscylacyjny PEP \u2013 Flutter, Acapella, RC-Cornet (PEP + oscylacje).')
add_bullet(tf, 'HFCWO \u2013 wysokocz\u0119stotliwo\u015bciowe oscylacje \u015bciany klatki (kamizelka).')


# ---------------------------------------------------------------------------
# 16. TERAPIA INHALACYJNA - KOLEJNOSC (proces)
# ---------------------------------------------------------------------------

slide = content_slide('Terapia inhalacyjna \u2013 kolejno\u015b\u0107',
                      kicker='Plan terapii')
steps = [
    ('1', 'Bronchodilatator', 'lek rozszerzaj\u0105cy oskrzela'),
    ('2', 'Mukolityk', 'hipertoniczny NaCl / dornaza alfa'),
    ('3', 'Techniki ACT', 'oczyszczanie dr\u00f3g oddechowych'),
    ('4', 'Antybiotyk wziewny', 'po skutecznym oczyszczeniu'),
]
n = len(steps)
bw = Inches(2.85); bh = Inches(2.6); gap = Inches(0.25)
total = bw * n + gap * (n - 1)
startx = (SW - total) / 2
y = Inches(2.6)
for i, (num, title, sub) in enumerate(steps):
    bx = startx + i * (bw + gap)
    rect(slide, bx, y, bw, bh, WHITE, line=RGBColor(0xD5, 0xDE, 0xE6))
    rect(slide, bx, y, bw, Inches(0.7), ACCENT)
    tb, ntf = textbox(slide, bx, y, bw, Inches(0.7), anchor=MSO_ANCHOR.MIDDLE)
    pn = ntf.paragraphs[0]; pn.alignment = PP_ALIGN.CENTER
    rn = pn.add_run(); rn.text = 'Krok ' + num
    _set_font(rn, 16, bold=True, color=WHITE)
    tb, btf = textbox(slide, bx + Inches(0.15), y + Inches(0.85), bw - Inches(0.3),
                      bh - Inches(1.0), anchor=MSO_ANCHOR.MIDDLE)
    pt = btf.paragraphs[0]; pt.alignment = PP_ALIGN.CENTER
    rt = pt.add_run(); rt.text = title
    _set_font(rt, 18, bold=True, color=PRIMARY)
    ps = btf.add_paragraph(); ps.alignment = PP_ALIGN.CENTER; ps.space_before = Pt(8)
    rs = ps.add_run(); rs.text = sub
    _set_font(rs, 13, color=GREY)
    if i < n - 1:
        ar = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                    bx + bw + Inches(0.01), y + bh/2 - Inches(0.18),
                                    gap - Inches(0.02), Inches(0.36))
        ar.fill.solid(); ar.fill.fore_color.rgb = PRIMARY
        ar.line.fill.background(); ar.shadow.inherit = False


# ---------------------------------------------------------------------------
# 17. TRENING FIZYCZNY
# ---------------------------------------------------------------------------

slide = content_slide('Trening fizyczny', kicker='Plan terapii')
tf = body_frame(slide)
add_bullet(tf, 'Uzupe\u0142nienie (nie zamiennik!) technik oczyszczania oskrzeli.', first=True)
add_bullet(tf, 'Wytrzyma\u0142o\u015bciowy (aerobowy): marsz, bieg, rower, p\u0142ywanie \u2013 '
               '3\u20135 \u00d7/tydz., 20\u201340 min, 60\u201380% HRmax.')
add_bullet(tf, 'Oporowy (si\u0142owy): 2\u20133 \u00d7/tydz., g\u0142\u00f3wne grupy mi\u0119\u015bniowe i mi\u0119\u015bnie '
               'posturalne.')
add_bullet(tf, 'Trening mi\u0119\u015bni oddechowych (IMT) \u2013 u wybranych pacjent\u00f3w.')
add_bullet(tf, '\u0106wiczenia rozci\u0105gaj\u0105ce i mobilizacja klatki piersiowej \u2013 korekcja postawy.')


# ---------------------------------------------------------------------------
# 18. CWICZENIA ODDECHOWE I EDUKACJA
# ---------------------------------------------------------------------------

slide = content_slide('\u0106wiczenia oddechowe i edukacja', kicker='Plan terapii')
tf = body_frame(slide)
add_bullet(tf, 'Oddychanie przepon\u0105 oraz nauka efektywnego kaszlu i techniki huffing.',
           first=True)
add_bullet(tf, '\u0106wiczenia zwi\u0119kszaj\u0105ce ruchomo\u015b\u0107 klatki piersiowej i pog\u0142\u0119biaj\u0105ce wdech.')
add_bullet(tf, 'Edukacja pacjenta i rodziny: systematyczno\u015b\u0107, higiena sprz\u0119tu.')
add_bullet(tf, 'Rozpoznawanie objaw\u00f3w zaostrzenia i samokontrola.')


# ---------------------------------------------------------------------------
# 19. PRZYKLADOWY PLAN TERAPII (TABELA)
# ---------------------------------------------------------------------------

slide = content_slide('Przyk\u0142adowy plan terapii (posta\u0107 umiarkowana)',
                      kicker='Plan terapii')
rows_data = [
    ('Element terapii', 'Cz\u0119stotliwo\u015b\u0107 / czas', 'Uwagi'),
    ('Inhalacja (bronchodilatator + NaCl)', '2 \u00d7 dziennie, przed ACT',
     'Przygotowanie dr\u00f3g oddechowych'),
    ('Techniki ACT (ACBT / oscylacyjny PEP)', '2 \u00d7 dziennie, 15\u201320 min',
     'Rano i wieczorem'),
    ('Trening aerobowy (rower / marsz)', '3\u20135 \u00d7/tydz., 20\u201340 min',
     'Umiarkowany, kontrola SpO2'),
    ('Trening oporowy', '2\u20133 \u00d7/tydz.', 'Mi\u0119\u015bnie posturalne'),
    ('\u0106wiczenia oddechowe + mobilizacja', 'codziennie, 10\u201315 min',
     'Huffing, rozci\u0105ganie'),
    ('Antybiotyk wziewny (je\u015bli zlecony)', 'wg lekarza, po ACT',
     'Po oczyszczeniu dr\u00f3g'),
]
rows = len(rows_data); cols = 3
tbl_x, tbl_y = Inches(0.7), Inches(1.65)
tbl_w, tbl_h = Inches(11.93), Inches(5.0)
gtable = slide.shapes.add_table(rows, cols, tbl_x, tbl_y, tbl_w, tbl_h).table
gtable.columns[0].width = Inches(5.0)
gtable.columns[1].width = Inches(3.6)
gtable.columns[2].width = Inches(3.33)
gtable.first_row = True
for r_i, row in enumerate(rows_data):
    for c_i, val in enumerate(row):
        cell = gtable.cell(r_i, c_i)
        cell.margin_left = Inches(0.12); cell.margin_right = Inches(0.08)
        cell.margin_top = Inches(0.04); cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf2 = cell.text_frame; tf2.word_wrap = True
        p = tf2.paragraphs[0]
        run = p.add_run(); run.text = val
        if r_i == 0:
            cell.fill.solid(); cell.fill.fore_color.rgb = PRIMARY
            _set_font(run, 14, bold=True, color=WHITE)
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT if r_i % 2 else WHITE
            _set_font(run, 13, color=DARK, bold=(c_i == 0))


# ---------------------------------------------------------------------------
# 20. KRYTERIA OCENY SKUTECZNOSCI
# ---------------------------------------------------------------------------

slide = content_slide('Kryteria oceny skuteczno\u015bci terapii',
                      kicker='Monitorowanie')
tf = body_frame(slide)
add_bullet(tf, 'Parametry czynno\u015bciowe: FEV1, FVC, SpO2.', first=True)
add_bullet(tf, 'Ilo\u015b\u0107 i charakter wydzieliny oraz nasilenie kaszlu.')
add_bullet(tf, 'Tolerancja wysi\u0142ku: 6MWT, VO2peak, skala Borga.')
add_bullet(tf, 'Cz\u0119sto\u015b\u0107 zaostrze\u0144 i hospitalizacji.')
add_bullet(tf, 'Jako\u015b\u0107 \u017cycia (CFQ-R) i przestrzeganie zalece\u0144.')


# ---------------------------------------------------------------------------
# 21. PODSUMOWANIE
# ---------------------------------------------------------------------------

slide = content_slide('Podsumowanie i wnioski', kicker='Zako\u0144czenie')
tf = body_frame(slide)
add_bullet(tf, 'Mukowiscydoza \u2013 przewlek\u0142a, post\u0119puj\u0105ca choroba; fizjoterapia oddechowa '
               'odgrywa kluczow\u0105 rol\u0119.', first=True)
add_bullet(tf, 'Dob\u00f3r technik ACT musi by\u0107 zindywidualizowany \u2013 to sprzyja '
               'przestrzeganiu zalece\u0144.')
add_bullet(tf, 'Trening fizyczny poprawia wydolno\u015b\u0107, si\u0142\u0119 mi\u0119\u015bniow\u0105 i jako\u015b\u0107 \u017cycia.')
add_bullet(tf, 'Edukacja pacjenta i monitorowanie efekt\u00f3w warunkuj\u0105 trwa\u0142e korzy\u015bci '
               'zdrowotne.')


# ---------------------------------------------------------------------------
# 22. BIBLIOGRAFIA
# ---------------------------------------------------------------------------

slide = content_slide('Bibliografia', kicker='Pi\u015bmiennictwo')
tf = body_frame(slide, top=Inches(1.55))
refs = [
    'Elborn J.S. Cystic fibrosis. The Lancet, 2016; 388(10059): 2519\u20132531.',
    'Mazurek H. (red.). Mukowiscydoza. Choroba wieloukladowa. Termedia, Pozna\u0144 2020.',
    'Castellani C. i wsp. ECFS best practice guidelines: 2018 revision. J Cyst Fibros, 2018; 17(2): 153\u2013178.',
    'Flume P.A. i wsp. CF pulmonary guidelines: airway clearance therapies. Respir Care, 2009; 54(4): 522\u2013537.',
    'Pryor J.A., Prasad S.A. Physiotherapy for Respiratory and Cardiac Problems. 4th ed., Elsevier, 2008.',
    'Smyth A.R. i wsp. ECFS Standards of Care: Best Practice guidelines. J Cyst Fibros, 2014; 13(S1): S23\u2013S42.',
    'Radtke T. i wsp. Physical activity and exercise training in CF. Cochrane Database Syst Rev, 2022; CD002768.',
    'Mc Ilwaine M. i wsp. Personalising airway clearance in chronic lung disease. Eur Respir Rev, 2017; 26(143): 160086.',
]
for i, t in enumerate(refs):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.space_after = Pt(7); p.line_spacing = 1.02
    rn = p.add_run(); rn.text = '[{}]  '.format(i + 1)
    _set_font(rn, 13, bold=True, color=ACCENT)
    r = p.add_run(); r.text = t
    _set_font(r, 13, color=DARK)


# ---------------------------------------------------------------------------
# 23. SLAJD KONCOWY
# ---------------------------------------------------------------------------

slide = prs.slides.add_slide(BLANK)
rect(slide, 0, 0, SW, SH, PRIMARY)
rect(slide, 0, Inches(3.35), SW, Inches(0.06), ACCENT)
tb, tf = textbox(slide, Inches(1.0), Inches(2.7), Inches(11.3), Inches(2.2),
                 anchor=MSO_ANCHOR.MIDDLE)
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = 'Dzi\u0119kuj\u0119 za uwag\u0119'
_set_font(r, 40, bold=True, color=WHITE)
p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER; p2.space_before = Pt(14)
r2 = p2.add_run()
r2.text = 'Ernest Kurdziel  \u2022  album 72714  \u2022  prowadz\u0105cy: dr Weronika Cyganik'
_set_font(r2, 16, color=LIGHT)


out = 'Fizjoterapia_mukowiscydoza_prezentacja.pptx'
prs.save(out)
print('Zapisano:', out, '| slajdy:', len(prs.slides._sldIdLst))
