# Prezentacja: Zespół cieśni nadgarstka

Indywidualny program zabiegów fizykalnych w ramach balneoklimatologii
i odnowy biologicznej — przypadek kliniczny pacjenta z zespołem cieśni
nadgarstka (39-letni informatyk).

## Zawartość

- `Zespol_ciesni_nadgarstka_prezentacja.pptx` — gotowa prezentacja (13 slajdów).
- `Zespol_ciesni_nadgarstka_prezentacja.pdf` — wersja PDF do szybkiego podglądu.
- `generuj_prezentacje.py` — skrypt generujący prezentację.

## Struktura prezentacji (10–15 slajdów)

1. Slajd tytułowy
2. Plan prezentacji
3. Charakterystyka pacjenta
4. Analiza wskazań
5. Przeciwwskazania (ogólne i miejscowe)
6. Dobór zabiegów fizykalnych (1/2) — balneoterapia i hydroterapia
7. Dobór zabiegów fizykalnych (2/2) — fizykoterapia, masaż, kinezyterapia
8. Opracowanie programu terapii (plan tygodniowy)
9. Efekty terapeutyczne
10. Elementy odnowy biologicznej
11. Podsumowanie i wnioski
12. Bibliografia
13. Slajd końcowy

## Ponowne wygenerowanie

```bash
pip install python-pptx
python generuj_prezentacje.py
```

Aby utworzyć wersję PDF (wymaga LibreOffice):

```bash
soffice --headless --convert-to pdf Zespol_ciesni_nadgarstka_prezentacja.pptx
```
