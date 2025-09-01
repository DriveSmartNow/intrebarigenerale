# Script pentru procesarea întrebărilor DRPCIV din Word

Acest script procesează fișierul Word cu întrebările DRPCIV și extrage toate întrebările, răspunsurile, răspunsurile corecte și imaginile într-un format JSON structurat.

## Rezultate extracție

✅ **Extracție reușită din documentul DRPCIV:**
- **797 întrebări** complete extrase
- **365 imagini** salvate în format optimizat
- **Format JSON** structurat generat
- **Răspunsuri multiple choice** (A, B, C) pentru fiecare întrebare

## Cerințe

- Python 3.7+
- Pachete Python: `python-docx`, `Pillow`, `tqdm`

## Instalare

1. Clonează sau descarcă acest repository
2. Instalează dependențele:

```bash
pip install -r requirements.txt
```

## Utilizare

### Rulare básică

```bash
python extract_questions.py
```

Aceasta va procesa fișierul implicit `473514317-227353061-Intrebari-Drpciv-docx.docx` din directorul curent.

### Rulare cu fișier personalizat

```bash
python extract_questions.py cale/catre/fisier.docx
```

### Specificarea directorului de ieșire

```bash
python extract_questions.py --output-dir /cale/catre/output
```

## Fișiere generate

După rularea scriptului se vor crea:

1. **`intrebari-generale.json`** - Fișierul JSON cu toate întrebările structurate
2. **`imagini/`** - Directorul cu imaginile extrase și optimizate (365 imagini)
3. **`extract_questions.log`** - Fișierul de log cu detalii despre procesare

## Structura JSON generată

```json
{
  "categorii": [
    {
      "nume": "Mediu de învățare general",
      "descriere": "Întrebări generale pentru examenul auto DRPCIV",
      "intrebari": [
        {
          "id": 1,
          "intrebare": "Textul întrebării...",
          "raspunsuri": ["A) Prima variantă", "B) A doua variantă", "C) A treia variantă"]
        }
      ]
    }
  ],
  "total_intrebari": 797,
  "versiune": "1.0",
  "ultima_actualizare": "2025-09-01"
}
```

## Note importante

### Răspunsuri corecte
Documentul sursă nu conține indicatori expliciti pentru răspunsurile corecte. Structura JSON este pregătită să includă câmpul `raspuns_corect` când această informație va fi disponibilă.

### Imagini
- 365 imagini au fost extrase și optimizate
- Imaginile sunt salvate în directorul `imagini/` cu nume format `intrebare_X.jpg`
- Asocierea imaginilor cu întrebările specifice necesită ajustări suplimentare în funcție de structura documentului

## Funcționalități

- **Extracție automată** a întrebărilor din format Word
- **Parsare inteligentă** a pattern-urilor de întrebări și răspunsuri
- **Extracție și optimizare imagini** pentru web
- **Validare** integritate date
- **Logging detaliat** pentru debugging
- **Progress bar** pentru monitorizarea procesării

## Pattern-uri recunoscute

Scriptul recunoaște următoarele formate de întrebări:

- `1. Textul întrebării`
- `1) Textul întrebării`
- `Întrebarea 1: Textul întrebării`

Răspunsuri în format:
- `A. Primul răspuns`
- `B. Al doilea răspuns`
- `C. Al treilea răspuns`

## Optimizări imagini

Imaginile sunt optimizate automat:
- Redimensionare la maximum 800x600 pixeli
- Compresie JPEG la 85% calitate
- Optimizare PNG

## Troubleshooting

### Erori comune

1. **ModuleNotFoundError**: Instalează dependențele cu `pip install -r requirements.txt`
2. **FileNotFoundError**: Verifică că fișierul Word există în locația specificată
3. **PermissionError**: Verifică permisiunile pentru directorul de ieșire

### Logging

Pentru debugging detaliat, verifică fișierul `extract_questions.log` generat automat.

## Dezvoltare

Pentru modificări la script:

1. Pattern-urile de recunoaștere sunt în clasa `DrpcivQuestionExtractor`
2. Logica de parsare este modulară și poate fi extinsă
3. Validarea datelor poate fi îmbunătățită în metoda `_is_question_valid`

## Licență

Acest script este destinat procesării documentelor DRPCIV pentru scopuri educaționale.