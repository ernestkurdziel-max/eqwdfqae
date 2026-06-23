# Prezentacja: Zespół cieśni nadgarstka

Indywidualny program zabiegów fizykalnych w ramach balneoklimatologii
i odnowy biologicznej — przypadek kliniczny pacjenta z zespołem cieśni
nadgarstka (39-letni informatyk).

Autor: **Ernest Kurdziel**, nr albumu **72714**.

## Zawartość

- `Zespol_ciesni_nadgarstka_prezentacja.pptx` — gotowa prezentacja (15 slajdów).
- `Zespol_ciesni_nadgarstka_prezentacja.pdf` — wersja PDF do szybkiego podglądu.
- `generuj_prezentacje.py` — skrypt generujący prezentację.
- `obrazy/` — zdjęcia i tła użyte w prezentacji.

## Struktura prezentacji (15 slajdów)

1. Slajd tytułowy (tło fotograficzne, autor i nr albumu)
2. Plan prezentacji
3. Charakterystyka pacjenta (+ ilustracja anatomiczna)
4. Analiza wskazań
5. Przeciwwskazania (ogólne i miejscowe)
6. Dobór zabiegów – balneoterapia (kąpiele lecznicze i wodolecznictwo)
7. Dobór zabiegów – peloidoterapia, termoterapia i balneoklimatologia
8. Dobór zabiegów – fizykoterapia, masaż i kinezyterapia
9. Dobór zabiegów – zabiegi odnowy biologicznej
10. Opracowanie programu terapii (plan tygodniowy)
11. Efekty terapeutyczne
12. Elementy odnowy biologicznej (styl życia)
13. Podsumowanie i wnioski
14. Bibliografia
15. Slajd końcowy

## Ponowne wygenerowanie

```bash
pip install python-pptx Pillow
python generuj_prezentacje.py
```

Aby utworzyć wersję PDF (wymaga LibreOffice):

```bash
soffice --headless --convert-to pdf Zespol_ciesni_nadgarstka_prezentacja.pptx
```
