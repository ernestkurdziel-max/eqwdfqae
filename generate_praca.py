# -*- coding: utf-8 -*-
"""
Generator pracy:
"Zasady planowania i programowania fizjoterapii pacjentow
z chorobami ukladu oddechowego" - na przykladzie mukowiscydozy.

Skrypt tworzy dokument Word (.docx) przy uzyciu biblioteki python-docx.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ---------------------------------------------------------------------------
# Funkcje pomocnicze
# ---------------------------------------------------------------------------

def set_cell_background(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tc_pr.append(shd)


def add_page_number_field(paragraph):
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)


def body_paragraph(doc, text, justify=True, first_line_indent=True):
    p = doc.add_paragraph(text)
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_after = Pt(6)
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if first_line_indent:
        pf.first_line_indent = Cm(1.0)
    return p


def bullet(doc, text, bold_lead=None):
    p = doc.add_paragraph(style='List Bullet')
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_after = Pt(3)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if bold_lead:
        r = p.add_run(bold_lead)
        r.bold = True
        p.add_run(text)
    else:
        p.add_run(text)
    return p


def numbered(doc, text):
    p = doc.add_paragraph(style='List Number')
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_after = Pt(3)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run(text)
    return p


# ---------------------------------------------------------------------------
# Dokument i style
# ---------------------------------------------------------------------------

doc = Document()

# Marginesy
for section in doc.sections:
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(2.5)

# Czcionka domyslna
normal = doc.styles['Normal']
normal.font.name = 'Times New Roman'
normal.font.size = Pt(12)
normal._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

# Naglowki
for lvl, size in [('Heading 1', 14), ('Heading 2', 13), ('Heading 3', 12)]:
    st = doc.styles[lvl]
    st.font.name = 'Times New Roman'
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = RGBColor(0x1F, 0x3B, 0x57)


# ---------------------------------------------------------------------------
# STRONA TYTULOWA
# ---------------------------------------------------------------------------

def center(doc, text, size=12, bold=False, italic=False, space_before=0, space_after=6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    return p


center(doc, 'Uczelnia / Wydzia\u0142', 12, space_before=6)
center(doc, 'Kierunek: Fizjoterapia', 12, space_after=24)

for _ in range(3):
    doc.add_paragraph()

center(doc, 'Imi\u0119 i Nazwisko', 12, space_after=2)
center(doc, 'Numer albumu: ........', 11, italic=True, space_after=36)

center(doc, 'Zasady planowania i programowania fizjoterapii '
            'pacjent\u00f3w z chorobami uk\u0142adu oddechowego', 18, bold=True,
       space_before=12, space_after=10)
center(doc, 'na przyk\u0142adzie mukowiscydozy (zw\u0142\u00f3knienia torbielowatego)',
       14, italic=True, space_after=36)

for _ in range(4):
    doc.add_paragraph()

center(doc, 'Praca zaliczeniowa', 12, italic=True, space_after=2)
center(doc, 'napisana pod kierunkiem', 12, space_after=2)
center(doc, 'dr / dr hab. ........................', 12, space_after=48)

for _ in range(3):
    doc.add_paragraph()

center(doc, 'Miejscowo\u015b\u0107, 2026', 12)

doc.add_page_break()


# ---------------------------------------------------------------------------
# SPIS TRESCI (pole TOC)
# ---------------------------------------------------------------------------

h = doc.add_heading('Spis tre\u015bci', level=1)
h.alignment = WD_ALIGN_PARAGRAPH.LEFT

p = doc.add_paragraph()
run = p.add_run()
fld_begin = OxmlElement('w:fldChar')
fld_begin.set(qn('w:fldCharType'), 'begin')
instr = OxmlElement('w:instrText')
instr.set(qn('xml:space'), 'preserve')
instr.text = 'TOC \\o "1-3" \\h \\z \\u'
fld_sep = OxmlElement('w:fldChar')
fld_sep.set(qn('w:fldCharType'), 'separate')
fld_text = OxmlElement('w:t')
fld_text.text = ('Aby zaktualizowa\u0107 spis tre\u015bci: kliknij prawym przyciskiem '
                 'i wybierz "Aktualizuj pole" (lub naci\u015bnij F9).')
run_t = OxmlElement('w:r')
run_t.append(fld_text)
fld_end = OxmlElement('w:fldChar')
fld_end.set(qn('w:fldCharType'), 'end')
run._r.append(fld_begin)
run._r.append(instr)
run._r.append(fld_sep)
run._r.append(run_t)
run._r.append(fld_end)

doc.add_page_break()


# ---------------------------------------------------------------------------
# STOPKA Z NUMERACJA STRON (od sekcji z trescia)
# ---------------------------------------------------------------------------

footer = doc.sections[0].footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_page_number_field(fp)


# ---------------------------------------------------------------------------
# WSTEP
# ---------------------------------------------------------------------------

doc.add_heading('Wst\u0119p', level=1)

body_paragraph(doc,
    'Choroby uk\u0142adu oddechowego stanowi\u0105 jedn\u0105 z najwa\u017cniejszych przyczyn '
    'chorobowo\u015bci i \u015bmiertelno\u015bci na \u015bwiecie, a ich przebieg w istotny spos\u00f3b '
    'obni\u017ca jako\u015b\u0107 \u017cycia pacjent\u00f3w oraz ogranicza ich wydolno\u015b\u0107 funkcjonaln\u0105. '
    'W post\u0119powaniu terapeutycznym, obok leczenia farmakologicznego, kluczow\u0105 rol\u0119 '
    'odgrywa fizjoterapia, kt\u00f3rej zadaniem jest utrzymanie dro\u017cno\u015bci dr\u00f3g oddechowych, '
    'poprawa wentylacji p\u0142uc, zwi\u0119kszenie tolerancji wysi\u0142ku oraz spowolnienie '
    'post\u0119pu zmian strukturalnych w obr\u0119bie uk\u0142adu oddechowego.')

body_paragraph(doc,
    'Celem niniejszej pracy jest przedstawienie zasad planowania i programowania '
    'fizjoterapii u pacjent\u00f3w z chorobami uk\u0142adu oddechowego na przyk\u0142adzie '
    'mukowiscydozy (zw\u0142\u00f3knienia torbielowatego, ang. cystic fibrosis, CF). '
    'Mukowiscydoza zosta\u0142a wybrana jako jednostka chorobowa ze wzgl\u0119du na jej '
    'przewlek\u0142y, post\u0119puj\u0105cy charakter oraz fakt, \u017ce fizjoterapia stanowi w niej '
    'nieod\u0142\u0105czny i codzienny element leczenia, prowadzony przez ca\u0142e \u017cycie chorego.')

body_paragraph(doc,
    'W pracy om\u00f3wiono charakterystyk\u0119 kliniczn\u0105 mukowiscydozy, metody badania '
    'pacjenta z chorob\u0105 uk\u0142adu oddechowego (w tym wywiad, badanie przedmiotowe, '
    'ocen\u0119 duszno\u015bci oraz badania czynno\u015bciowe i wydolno\u015bciowe), og\u00f3lne zasady '
    'planowania post\u0119powania fizjoterapeutycznego, a nast\u0119pnie zaproponowano '
    'przyk\u0142adowy, kompleksowy plan terapii. Ca\u0142o\u015b\u0107 oparto na aktualnym pi\u015bmiennictwie '
    'naukowym oraz wytycznych towarzystw naukowych.')


# ---------------------------------------------------------------------------
# ROZDZIAL 1 - CEL I ZAKRES
# ---------------------------------------------------------------------------

doc.add_heading('1. Cel i zakres pracy', level=1)

body_paragraph(doc, 'Cele szczeg\u00f3\u0142owe pracy obejmuj\u0105:', first_line_indent=False)
bullet(doc, 'przedstawienie etiopatogenezy i obrazu klinicznego mukowiscydozy '
            'ze szczeg\u00f3lnym uwzgl\u0119dnieniem zmian w uk\u0142adzie oddechowym;')
bullet(doc, 'om\u00f3wienie metod badania (diagnostyki funkcjonalnej) pacjenta '
            'z chorob\u0105 uk\u0142adu oddechowego wykorzystywanych w praktyce fizjoterapeutycznej;')
bullet(doc, 'opis zasad planowania i programowania fizjoterapii, w tym '
            'formu\u0142owania cel\u00f3w kr\u00f3tko- i d\u0142ugoterminowych;')
bullet(doc, 'opracowanie przyk\u0142adowego planu terapii oraz zaprezentowanie '
            'najwa\u017cniejszych technik oczyszczania drzewa oskrzelowego i treningu '
            'fizycznego.')


# ---------------------------------------------------------------------------
# ROZDZIAL 2 - MUKOWISCYDOZA
# ---------------------------------------------------------------------------

doc.add_heading('2. Charakterystyka mukowiscydozy', level=1)

doc.add_heading('2.1. Definicja i etiologia', level=2)
body_paragraph(doc,
    'Mukowiscydoza (zw\u0142\u00f3knienie torbielowate, cystic fibrosis, CF) jest '
    'najcz\u0119stsz\u0105 chorob\u0105 genetyczn\u0105 dziedziczon\u0105 w spos\u00f3b autosomalny recesywny '
    'w populacji rasy bia\u0142ej. Jej przyczyn\u0105 s\u0105 mutacje w genie CFTR '
    '(ang. Cystic Fibrosis Transmembrane Conductance Regulator) zlokalizowanym '
    'na d\u0142ugim ramieniu chromosomu 7. Najcz\u0119stsz\u0105 mutacj\u0105 jest F508del '
    '(delecja fenyloalaniny w pozycji 508). Bia\u0142ko CFTR funkcjonuje jako kana\u0142 '
    'chlorkowy w b\u0142onach komórek nab\u0142onkowych; jego dysfunkcja prowadzi do '
    'zaburzenia transportu jon\u00f3w chlorkowych i sodowych oraz wody, czego '
    'nast\u0119pstwem jest produkcja g\u0119stego, lepkiego \u015bluzu [1, 2].')

body_paragraph(doc,
    'Zaburzenie to ma charakter wielonarz\u0105dowy i dotyczy przede wszystkim '
    'uk\u0142adu oddechowego, trzustki i przewodu pokarmowego, w\u0105troby, gruczo\u0142\u00f3w '
    'potowych oraz uk\u0142adu rozrodczego. Najwi\u0119kszy wp\u0142yw na rokowanie ma jednak '
    'post\u0119puj\u0105ca choroba oskrzelowo-p\u0142ucna, kt\u00f3ra odpowiada za wi\u0119kszo\u015b\u0107 '
    'powik\u0142a\u0144 i zgon\u00f3w [2, 3].')

doc.add_heading('2.2. Patofizjologia zmian w uk\u0142adzie oddechowym', level=2)
body_paragraph(doc,
    'W drogach oddechowych nieprawid\u0142owy transport jon\u00f3w prowadzi do odwodnienia '
    'warstwy p\u0142ynu okrywaj\u0105cego nab\u0142onek rz\u0119skowy. Skutkuje to zag\u0119szczeniem '
    '\u015bluzu i upo\u015bledzeniem transportu \u015bluzowo-rz\u0119skowego (klirensu '
    'mukocyliarnego). Zalegaj\u0105ca, g\u0119sta wydzielina stanowi po\u017cywk\u0119 dla '
    'bakterii, sprzyjaj\u0105c przewlek\u0142ym zaka\u017ceniom (m.in. Staphylococcus aureus, '
    'Pseudomonas aeruginosa). Powtarzaj\u0105ce si\u0119 cykle infekcji i zapalenia '
    'prowadz\u0105 do uszkodzenia \u015bciany oskrzeli, rozwoju rozstrzeni oskrzeli, '
    'w\u0142\u00f3knienia oraz post\u0119puj\u0105cej obturacji [1, 3, 4].')

doc.add_heading('2.3. Objawy ze strony uk\u0142adu oddechowego', level=2)
bullet(doc, 'przewlek\u0142y, produktywny kaszel z odkrztuszaniem g\u0119stej wydzieliny;')
bullet(doc, 'nawracaj\u0105ce i przewlek\u0142e zaka\u017cenia dolnych dr\u00f3g oddechowych;')
bullet(doc, 'duszno\u015b\u0107 wysi\u0142kowa, a w zaawansowanych stadiach r\u00f3wnie\u017c spoczynkowa;')
bullet(doc, '\u015bwisty, furczenia i trzeszczenia w os\u0142uchiwaniu;')
bullet(doc, 'zmniejszenie tolerancji wysi\u0142ku fizycznego;')
bullet(doc, 'w stadiach zaawansowanych \u2013 palce pa\u0142eczkowate, beczkowata klatka '
            'piersiowa, objawy przewlek\u0142ej niewydolno\u015bci oddechowej.')

doc.add_heading('2.4. Znaczenie fizjoterapii w mukowiscydozie', level=2)
body_paragraph(doc,
    'Fizjoterapia, a w szczeg\u00f3lno\u015bci techniki oczyszczania drzewa oskrzelowego '
    '(ang. airway clearance techniques, ACT), stanowi podstawowy i codzienny '
    'element leczenia mukowiscydozy. Jej regularne stosowanie wspomaga usuwanie '
    'zalegaj\u0105cej wydzieliny, zmniejsza cz\u0119sto\u015b\u0107 zaostrze\u0144, poprawia wentylacj\u0119 '
    'oraz \u2013 w po\u0142\u0105czeniu z treningiem fizycznym \u2013 wp\u0142ywa korzystnie na wydolno\u015b\u0107 '
    'i jako\u015b\u0107 \u017cycia chorych [4, 5, 6].')


# ---------------------------------------------------------------------------
# ROZDZIAL 3 - METODY BADANIA
# ---------------------------------------------------------------------------

doc.add_heading('3. Metody badania pacjenta z chorob\u0105 uk\u0142adu oddechowego', level=1)

body_paragraph(doc,
    'Prawid\u0142owo przeprowadzone badanie fizjoterapeutyczne jest punktem wyj\u015bcia '
    'do zaplanowania skutecznej terapii. Powinno mie\u0107 charakter kompleksowy '
    'i obejmowa\u0107 zar\u00f3wno ocen\u0119 podmiotow\u0105 (wywiad), jak i przedmiotow\u0105 oraz '
    'badania czynno\u015bciowe i wydolno\u015bciowe.')

doc.add_heading('3.1. Wywiad (badanie podmiotowe)', level=2)
bullet(doc, 'dane og\u00f3lne: wiek, masa cia\u0142a, wzrost, BMI, styl \u017cycia, '
            'aktywno\u015b\u0107 fizyczna;')
bullet(doc, 'wywiad chorobowy: czas trwania choroby, wynik testu potowego '
            'i bada\u0144 genetycznych, przebyte zaostrzenia i hospitalizacje;')
bullet(doc, 'charakterystyka objaw\u00f3w: kaszel (rodzaj, pora dnia), ilo\u015b\u0107 '
            'i charakter odkrztuszanej wydzieliny, duszno\u015b\u0107;')
bullet(doc, 'leczenie farmakologiczne: leki rozszerzaj\u0105ce oskrzela, mukolityki '
            '(dornaza alfa, hipertoniczny NaCl), antybiotyki, modulatory CFTR;')
bullet(doc, 'dotychczasowa fizjoterapia: stosowane techniki, ich skuteczno\u015b\u0107 '
            'i akceptacja przez pacjenta;')
bullet(doc, 'ocena jako\u015bci \u017cycia z wykorzystaniem kwestionariuszy '
            '(np. CFQ-R \u2013 Cystic Fibrosis Questionnaire-Revised).')

doc.add_heading('3.2. Badanie przedmiotowe', level=2)
body_paragraph(doc, 'Klasyczne badanie fizykalne klatki piersiowej obejmuje '
                    'cztery podstawowe sk\u0142adowe:', first_line_indent=False)
bullet(doc, 'ocen\u0119 budowy i ruchomo\u015bci klatki piersiowej, toru oddechowego, '
            'cz\u0119sto\u015bci oddech\u00f3w, obecno\u015bci sinicy i palc\u00f3w pa\u0142eczkowatych, '
            'wykorzystania dodatkowych mi\u0119\u015bni oddechowych;', bold_lead='Ogl\u0105danie \u2013 ')
bullet(doc, 'ocen\u0119 ruchomo\u015bci klatki piersiowej, dr\u017cenia g\u0142osowego oraz '
            'lokalizacji ewentualnej bolesno\u015bci;', bold_lead='Palpacja (obmacywanie) \u2013 ')
bullet(doc, 'r\u00f3\u017cnicowanie odg\u0142osu opukowego (jawny, st\u0142umiony, b\u0119benkowy), '
            'pomocne w lokalizacji obszar\u00f3w niedodmy lub rozedmy;', bold_lead='Opukiwanie \u2013 ')
bullet(doc, 'ocen\u0119 szmeru oddechowego oraz zjawisk dodatkowych '
            '(\u015bwisty, furczenia, trzeszczenia, rz\u0119\u017cenia), pomocne w lokalizacji '
            'zalegaj\u0105cej wydzieliny.', bold_lead='Os\u0142uchiwanie (auskultacja) \u2013 ')

doc.add_heading('3.3. Ocena duszno\u015bci i wysi\u0142ku oddechowego', level=2)
body_paragraph(doc,
    'Do oceny nasilenia duszno\u015bci wykorzystuje si\u0119 wystandaryzowane skale, '
    'co umo\u017cliwia obiektywizacj\u0119 i monitorowanie post\u0119p\u00f3w terapii:', first_line_indent=False)
bullet(doc, 'skala mMRC (modified Medical Research Council) \u2013 ocena duszno\u015bci '
            'w zwi\u0105zku z aktywno\u015bci\u0105 codzienn\u0105 w zakresie 0\u20134;')
bullet(doc, 'skala Borga (oraz zmodyfikowana skala Borga CR10) \u2013 subiektywna '
            'ocena duszno\u015bci i zm\u0119czenia, m.in. podczas wysi\u0142ku;')
bullet(doc, 'wizualna skala analogowa (VAS) \u2013 ocena duszno\u015bci lub bolesno\u015bci.')

doc.add_heading('3.4. Badania czynno\u015bciowe uk\u0142adu oddechowego', level=2)
bullet(doc, 'pozwala oceni\u0107 obecno\u015b\u0107 i nasilenie obturacji '
            '(FEV1, FVC, wska\u017anik FEV1/FVC) oraz monitorowa\u0107 post\u0119p choroby; '
            'FEV1 jest podstawowym parametrem oceny czynno\u015bci p\u0142uc w CF;',
       bold_lead='Spirometria \u2013 ')
bullet(doc, 'nieinwazyjny pomiar wysycenia hemoglobiny tlenem (SpO2), '
            'wykorzystywany w spoczynku i podczas wysi\u0142ku;', bold_lead='Pulsoksymetria \u2013 ')
bullet(doc, 'ocena pr\u0119\u017cno\u015bci gaz\u00f3w (PaO2, PaCO2) i r\u00f3wnowagi kwasowo-zasadowej '
            'w stanach zaawansowanej niewydolno\u015bci oddechowej;', bold_lead='Gazometria \u2013 ')
bullet(doc, 'wczesny wska\u017anik nier\u00f3wnomierno\u015bci wentylacji (LCI \u2013 lung clearance '
            'index), coraz cz\u0119\u015bciej stosowany w monitorowaniu CF.', bold_lead='Wymywanie azotu / test wielu oddech\u00f3w \u2013 ')

doc.add_heading('3.5. Ocena wydolno\u015bci fizycznej', level=2)
bullet(doc, 'test 6-minutowego marszu (6MWT) \u2013 prosta, powtarzalna pr\u00f3ba '
            'oceniaj\u0105ca dystans i reakcj\u0119 (SpO2, t\u0119tno, duszno\u015b\u0107) na wysi\u0142ek '
            'submaksymalny;')
bullet(doc, 'sercowo-p\u0142ucny test wysi\u0142kowy (CPET) \u2013 z\u0142oty standard oceny '
            'wydolno\u015bci tlenowej (szczytowe poch\u0142anianie tlenu, VO2peak);')
bullet(doc, 'test wahad\u0142owy (shuttle walk test) oraz pr\u00f3by si\u0142y mi\u0119\u015bniowej '
            '(np. si\u0142a mi\u0119\u015bni wdechowych i wydechowych \u2013 PImax, PEmax).')

doc.add_heading('3.6. Ocena postawy i uk\u0142adu mi\u0119\u015bniowo-szkieletowego', level=2)
body_paragraph(doc,
    'U pacjent\u00f3w z przewlek\u0142\u0105 chorob\u0105 oddechow\u0105 cz\u0119sto wyst\u0119puj\u0105 wt\u00f3rne zmiany '
    'w obr\u0119bie narz\u0105du ruchu: nadmierna kifoza piersiowa, protrakcja barków, '
    'os\u0142abienie mi\u0119\u015bni posturalnych, ograniczenie ruchomo\u015bci klatki piersiowej '
    'oraz zwi\u0119kszone napi\u0119cie pomocniczych mi\u0119\u015bni oddechowych. Ocena postawy '
    'cia\u0142a, ruchomo\u015bci kr\u0119gos\u0142upa i obwod\u00f3w klatki piersiowej (np. pomiar '
    'r\u00f3\u017cnicy obwod\u00f3w na wdechu i wydechu) uzupe\u0142nia obraz funkcjonalny pacjenta '
    'i ukierunkowuje program terapii [5, 6].')


# ---------------------------------------------------------------------------
# ROZDZIAL 4 - ZASADY PLANOWANIA
# ---------------------------------------------------------------------------

doc.add_heading('4. Zasady planowania i programowania fizjoterapii', level=1)

doc.add_heading('4.1. Og\u00f3lne zasady post\u0119powania', level=2)
bullet(doc, 'indywidualizacja \u2013 program dostosowany do wieku, stadium choroby, '
            'wydolno\u015bci i preferencji pacjenta;')
bullet(doc, 'systematyczno\u015b\u0107 \u2013 techniki oczyszczania oskrzeli stosowane s\u0105 '
            'codziennie, cz\u0119sto kilka razy dziennie, przez ca\u0142e \u017cycie;')
bullet(doc, 'kompleksowo\u015b\u0107 \u2013 po\u0142\u0105czenie technik oczyszczania, treningu '
            'fizycznego, edukacji i terapii inhalacyjnej;')
bullet(doc, 'stopniowanie obci\u0105\u017ce\u0144 \u2013 zgodnie z zasadami treningu zdrowotnego '
            'i tolerancj\u0105 pacjenta;')
bullet(doc, 'wsp\u00f3\u0142praca interdyscyplinarna \u2013 z lekarzem, dietetykiem, '
            'psychologiem i piel\u0119gniark\u0105;')
bullet(doc, 'monitorowanie i modyfikacja \u2013 regularna ocena efekt\u00f3w '
            'i dostosowywanie programu.')

doc.add_heading('4.2. Cele fizjoterapii', level=2)
body_paragraph(doc, 'Cele kr\u00f3tkoterminowe:', first_line_indent=False)
bullet(doc, 'u\u0142atwienie ewakuacji zalegaj\u0105cej wydzieliny i poprawa dro\u017cno\u015bci '
            'dr\u00f3g oddechowych;')
bullet(doc, 'zmniejszenie duszno\u015bci i pracy oddechowej;')
bullet(doc, 'poprawa wentylacji oraz utrzymanie ruchomo\u015bci klatki piersiowej.')

body_paragraph(doc, 'Cele d\u0142ugoterminowe:', first_line_indent=False)
bullet(doc, 'spowolnienie post\u0119pu zmian oskrzelowo-p\u0142ucnych i zmniejszenie '
            'cz\u0119sto\u015bci zaostrze\u0144;')
bullet(doc, 'utrzymanie lub poprawa wydolno\u015bci fizycznej i tolerancji wysi\u0142ku;')
bullet(doc, 'utrwalenie nawyku samodzielnej, codziennej terapii (samodzielno\u015b\u0107 '
            'pacjenta);')
bullet(doc, 'poprawa jako\u015bci \u017cycia oraz prawid\u0142owej postawy cia\u0142a.')

doc.add_heading('4.3. Dob\u00f3r metod \u2013 zasada \u201eod og\u00f3\u0142u do szczeg\u00f3\u0142u\u201d', level=2)
body_paragraph(doc,
    'Wsp\u00f3\u0142czesne wytyczne podkre\u015blaj\u0105, \u017ce nie istnieje jedna technika '
    'oczyszczania oskrzeli o udowodnionej przewadze nad pozosta\u0142ymi; wyb\u00f3r metody '
    'powinien by\u0107 zindywidualizowany i uwzgl\u0119dnia\u0107 wiek, mo\u017cliwo\u015bci, preferencje '
    'oraz akceptacj\u0119 pacjenta, co decyduje o d\u0142ugoterminowym przestrzeganiu '
    'zalece\u0144 [4, 6, 7]. U ma\u0142ych dzieci stosuje si\u0119 techniki bierne i wspomagane '
    'przez opiekuna, natomiast u starszych dzieci i doros\u0142ych d\u0105\u017cy si\u0119 do technik '
    'umo\u017cliwiaj\u0105cych samodzieln\u0105 terapi\u0119.')


# ---------------------------------------------------------------------------
# ROZDZIAL 5 - PLAN TERAPII
# ---------------------------------------------------------------------------

doc.add_heading('5. Plan terapii \u2013 propozycja post\u0119powania', level=1)

doc.add_heading('5.1. Techniki oczyszczania drzewa oskrzelowego (ACT)', level=2)
bullet(doc, 'wykorzystuje grawitacj\u0119 do przemieszczania wydzieliny z obwodowych '
            'do centralnych dr\u00f3g oddechowych; cz\u0119sto \u0142\u0105czony z oklepywaniem i wibracjami '
            '(obecnie rzadziej stosowany samodzielnie u doros\u0142ych);',
       bold_lead='Drena\u017c u\u0142o\u017ceniowy (pozycyjny) \u2013 ')
bullet(doc, 'cykl aktywnego oddychania \u2013 sekwencja kontroli oddechu, '
            '\u0107wicze\u0144 rozpr\u0119\u017caj\u0105cych klatk\u0119 piersiow\u0105 oraz techniki natywnego '
            'wydechu (huffing); jedna z podstawowych, skutecznych technik samodzielnych;',
       bold_lead='ACBT \u2013 ')
bullet(doc, 'drena\u017c autogeniczny \u2013 kontrolowane oddychanie na r\u00f3\u017cnych '
            'obj\u0119to\u015bciach p\u0142uc, pozwalaj\u0105ce mobilizowa\u0107 i transportowa\u0107 wydzielin\u0119 '
            'bez kaszlu wyniszczaj\u0105cego;', bold_lead='AD \u2013 ')
bullet(doc, 'terapia dodatnim ci\u015bnieniem wydechowym (maska/ustnik PEP) '
            'stabilizuje drogi oddechowe i u\u0142atwia mobilizacj\u0119 wydzieliny;',
       bold_lead='PEP \u2013 ')
bullet(doc, 'urz\u0105dzenia oscylacyjne (Flutter, Acapella, RC-Cornet) \u0142\u0105cz\u0105 '
            'dodatnie ci\u015bnienie wydechowe z oscylacjami, rozlu\u017aniaj\u0105c \u015bluz;',
       bold_lead='Oscylacyjny PEP \u2013 ')
bullet(doc, 'wysokocz\u0119stotliwo\u015bciowe oscylacje \u015bciany klatki piersiowej (kamizelka) '
            'jako alternatywa u wybranych pacjent\u00f3w.', bold_lead='HFCWO \u2013 ')

doc.add_heading('5.2. Terapia inhalacyjna i jej kolejno\u015b\u0107', level=2)
body_paragraph(doc,
    'Zaleca si\u0119 \u0142\u0105czenie fizjoterapii z terapi\u0105 inhalacyjn\u0105 w odpowiedniej '
    'kolejno\u015bci: najpierw lek rozszerzaj\u0105cy oskrzela, nast\u0119pnie \u015brodek '
    'rozrzedzaj\u0105cy wydzielin\u0119 (np. hipertoniczny roztw\u00f3r NaCl lub dornaza alfa), '
    'po czym wykonuje si\u0119 techniki oczyszczania oskrzeli, a antybiotyk wziewny '
    'podaje si\u0119 po skutecznym oczyszczeniu dr\u00f3g oddechowych [4, 7].')

doc.add_heading('5.3. Trening fizyczny', level=2)
body_paragraph(doc,
    'Regularny wysi\u0142ek fizyczny jest istotnym uzupe\u0142nieniem (a nie zamiennikiem) '
    'technik oczyszczania oskrzeli. Poprawia wydolno\u015b\u0107 tlenow\u0105, si\u0142\u0119 mi\u0119\u015bniow\u0105, '
    'g\u0119sto\u015b\u0107 mineraln\u0105 ko\u015bci oraz u\u0142atwia ewakuacj\u0119 wydzieliny [5, 6, 8].',
    first_line_indent=False)
bullet(doc, 'trening wytrzyma\u0142o\u015bciowy (aerobowy): marsz, bieg, jazda na rowerze, '
            'p\u0142ywanie \u2013 3\u20135 razy w tygodniu po 20\u201340 min, intensywno\u015b\u0107 umiarkowana '
            'do du\u017cej (np. 60\u201380% t\u0119tna maksymalnego);')
bullet(doc, 'trening oporowy (si\u0142owy): 2\u20133 razy w tygodniu, g\u0142\u00f3wne grupy '
            'mi\u0119\u015bniowe, ze szczeg\u00f3lnym uwzgl\u0119dnieniem mi\u0119\u015bni posturalnych;')
bullet(doc, 'trening mi\u0119\u015bni oddechowych (IMT) \u2013 u wybranych pacjent\u00f3w '
            'z os\u0142abieniem mi\u0119\u015bni wdechowych;')
bullet(doc, '\u0107wiczenia rozci\u0105gaj\u0105ce i mobilizacja klatki piersiowej \u2013 utrzymanie '
            'ruchomo\u015bci i korekcja postawy.')

doc.add_heading('5.4. \u0106wiczenia oddechowe i edukacja', level=2)
bullet(doc, '\u0107wiczenia toru oddechowego (oddychanie przepon\u0105) oraz nauka '
            'efektywnego, kontrolowanego kaszlu i techniki huffing;')
bullet(doc, '\u0107wiczenia zwi\u0119kszaj\u0105ce ruchomo\u015b\u0107 klatki piersiowej '
            'i pog\u0142\u0119biaj\u0105ce wdech;')
bullet(doc, 'edukacja pacjenta i rodziny: znaczenie systematyczno\u015bci, higiena '
            'sprz\u0119tu, rozpoznawanie objaw\u00f3w zaostrzenia, samokontrola.')

doc.add_heading('5.5. Przyk\u0142adowy plan terapii (pacjent doros\u0142y, posta\u0107 umiarkowana)', level=2)

table = doc.add_table(rows=1, cols=3)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
hdr[0].text = 'Element terapii'
hdr[1].text = 'Cz\u0119stotliwo\u015b\u0107 / czas'
hdr[2].text = 'Uwagi'
for c in hdr:
    c.paragraphs[0].runs[0].font.bold = True
    set_cell_background(c, '1F3B57')
    c.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

plan_rows = [
    ('Inhalacja (bronchodilatator + hipertoniczny NaCl)',
     '2 x dziennie, przed ACT',
     'Przygotowanie dr\u00f3g oddechowych do oczyszczania'),
    ('Techniki oczyszczania oskrzeli (ACBT / oscylacyjny PEP)',
     '2 x dziennie po 15\u201320 min',
     'Rano i wieczorem; modyfikacja w zaostrzeniu'),
    ('Trening aerobowy (rower / marsz)',
     '3\u20135 x w tygodniu, 20\u201340 min',
     'Intensywno\u015b\u0107 umiarkowana, kontrola SpO2'),
    ('Trening oporowy',
     '2\u20133 x w tygodniu',
     'G\u0142\u00f3wne grupy mi\u0119\u015bniowe, mi\u0119\u015bnie posturalne'),
    ('\u0106wiczenia oddechowe i mobilizacja klatki piersiowej',
     'codziennie, 10\u201315 min',
     'Oddychanie przepon\u0105, huffing, rozci\u0105ganie'),
    ('Antybiotyk wziewny (je\u015bli zlecony)',
     'wg zalece\u0144 lekarza, po ACT',
     'Po skutecznym oczyszczeniu dr\u00f3g oddechowych'),
    ('Edukacja i samokontrola',
     'sta\u0142y element',
     'Higiena sprz\u0119tu, rozpoznawanie zaostrze\u0144'),
]
for r0, r1, r2 in plan_rows:
    cells = table.add_row().cells
    cells[0].text = r0
    cells[1].text = r1
    cells[2].text = r2
    for c in cells:
        c.paragraphs[0].runs[0].font.size = Pt(11)

doc.add_paragraph()
body_paragraph(doc,
    'Przedstawiony plan ma charakter ramowy i wymaga indywidualnego dostosowania '
    'do aktualnego stanu pacjenta. W okresie zaostrzenia zwi\u0119ksza si\u0119 cz\u0119stotliwo\u015b\u0107 '
    'technik oczyszczania oskrzeli oraz modyfikuje intensywno\u015b\u0107 treningu, natomiast '
    'w stabilnym okresie choroby nacisk k\u0142adziony jest na utrzymanie wydolno\u015bci '
    'i nawyku regularnej terapii.')

doc.add_heading('5.6. Kryteria oceny skuteczno\u015bci terapii', level=2)
bullet(doc, 'parametry czynno\u015bciowe: FEV1, FVC, SpO2;')
bullet(doc, 'ilo\u015b\u0107 i charakter odkrztuszanej wydzieliny oraz nasilenie kaszlu;')
bullet(doc, 'tolerancja wysi\u0142ku (dystans 6MWT, VO2peak, skala Borga);')
bullet(doc, 'cz\u0119sto\u015b\u0107 zaostrze\u0144 i hospitalizacji;')
bullet(doc, 'jako\u015b\u0107 \u017cycia (kwestionariusz CFQ-R) i przestrzeganie zalece\u0144.')


# ---------------------------------------------------------------------------
# PODSUMOWANIE
# ---------------------------------------------------------------------------

doc.add_heading('6. Podsumowanie i wnioski', level=1)
body_paragraph(doc,
    'Fizjoterapia stanowi nieod\u0142\u0105czny element kompleksowego leczenia pacjent\u00f3w '
    'z chorobami uk\u0142adu oddechowego, a w przypadku mukowiscydozy jest stosowana '
    'codziennie przez ca\u0142e \u017cycie chorego. Skuteczne planowanie i programowanie '
    'terapii musi by\u0107 poprzedzone rzeteln\u0105 ocen\u0105 funkcjonaln\u0105 pacjenta, obejmuj\u0105c\u0105 '
    'wywiad, badanie przedmiotowe oraz badania czynno\u015bciowe i wydolno\u015bciowe.')

numbered(doc, 'Mukowiscydoza jest przewlek\u0142\u0105, post\u0119puj\u0105c\u0105 chorob\u0105, w kt\u00f3rej '
              'fizjoterapia oddechowa odgrywa kluczow\u0105 rol\u0119 w utrzymaniu dro\u017cno\u015bci '
              'dr\u00f3g oddechowych i spowalnianiu progresji zmian p\u0142ucnych.')
numbered(doc, 'Dob\u00f3r technik oczyszczania oskrzeli powinien by\u0107 zindywidualizowany '
              'i uwzgl\u0119dnia\u0107 wiek, stadium choroby oraz preferencje pacjenta, co '
              'sprzyja d\u0142ugoterminowemu przestrzeganiu zalece\u0144.')
numbered(doc, 'Trening fizyczny jest wa\u017cnym uzupe\u0142nieniem terapii, poprawiaj\u0105cym '
              'wydolno\u015b\u0107, si\u0142\u0119 mi\u0119\u015bniow\u0105 i jako\u015b\u0107 \u017cycia.')
numbered(doc, 'Edukacja pacjenta i jego rodziny oraz systematyczne monitorowanie '
              'efekt\u00f3w terapii s\u0105 niezb\u0119dne dla osi\u0105gni\u0119cia trwa\u0142ych korzy\u015bci '
              'zdrowotnych.')


# ---------------------------------------------------------------------------
# BIBLIOGRAFIA
# ---------------------------------------------------------------------------

doc.add_heading('Bibliografia', level=1)

refs = [
    'Elborn J.S.: Cystic fibrosis. The Lancet, 2016; 388(10059): 2519\u20132531.',
    'Mazurek H. (red.): Mukowiscydoza. Choroba wieloukladowa. Termedia, Pozna\u0144 2020.',
    'Castellani C., Duff A.J.A., Bell S.C. i wsp.: ECFS best practice guidelines: '
    'the 2018 revision. Journal of Cystic Fibrosis, 2018; 17(2): 153\u2013178.',
    'Flume P.A., Robinson K.A., O\u2019Sullivan B.P. i wsp.: Cystic fibrosis pulmonary '
    'guidelines: airway clearance therapies. Respiratory Care, 2009; 54(4): 522\u2013537.',
    'Pryor J.A., Prasad S.A.: Physiotherapy for Respiratory and Cardiac Problems: '
    'Adults and Paediatrics. 4th ed., Churchill Livingstone/Elsevier, 2008.',
    'Main E., Denehy L. (red.): Cardiorespiratory Physiotherapy: Adults and '
    'Paediatrics. 5th ed., Elsevier, 2016.',
    'Smyth A.R., Bell S.C., Bojcin S. i wsp.: European Cystic Fibrosis Society '
    'Standards of Care: Best Practice guidelines. Journal of Cystic Fibrosis, '
    '2014; 13 (Suppl 1): S23\u2013S42.',
    'Radtke T., Smith S., Nevitt S.J. i wsp.: Physical activity and exercise '
    'training in cystic fibrosis. Cochrane Database of Systematic Reviews, 2022; '
    'Issue 8: CD002768.',
    'Walicka-Serzysko K., Sands D.: Mukowiscydoza \u2013 wsp\u00f3\u0142czesne zasady leczenia. '
    'Pediatria po Dyplomie, 2019.',
    'Mc Ilwaine M., Bradley J., Elborn J.S., Moran F.: Personalising airway '
    'clearance in chronic lung disease. European Respiratory Review, 2017; '
    '26(143): 160086.',
]
for i, r in enumerate(refs, 1):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_after = Pt(4)
    pf.left_indent = Cm(0.75)
    pf.first_line_indent = Cm(-0.75)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run('[{}] '.format(i)).bold = True
    p.add_run(r)


# ---------------------------------------------------------------------------
# Zapis
# ---------------------------------------------------------------------------

out = 'Fizjoterapia_mukowiscydoza_praca.docx'
doc.save(out)
print('Zapisano:', out)
