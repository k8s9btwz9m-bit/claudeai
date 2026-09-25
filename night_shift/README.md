# R2 Night Shift Performance

`build_sheet.py` genera il file Google Sheets (xlsx da importare) per analizzare i turni notturni R2
(18-3 e 21-6): Dashboard (totale + split per turno, trend 8 settimane), Agenti, Spotcheck
(Handoff, Pending, NPS, Reopen) e Spotcheck_Log.

```
pip install openpyxl
python3 build_sheet.py "R2 Overnight Shift.xlsx" "R2 Night Shift Performance.xlsx"
```

Input: export xlsx del file "R2 Overnight Shift" (fogli `Schedules` e `Teableau`).
Le formule usano funzioni native di Google Sheets (MAP, LAMBDA, FILTER, SORT, MAKEARRAY):
il file va aperto in Google Sheets (File > Importa, oppure Apri con > Fogli Google).

Il file generato contiene dati di test (Fonte = TEST in Turni_Log) da cancellare prima dell'uso reale.
Istruzioni operative nel foglio `Leggimi`. I file xlsx non sono versionati (contengono dati dei dipendenti).
