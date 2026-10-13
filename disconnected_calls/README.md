# Disconnected Calls - All Markets (copia con dettaglio per caso)

`build_copy.py` crea una copia del file Google Sheets "[Foundever Copy] Disconnected Calls - All Markets"
con quanto chiesto dal cliente:

1. riepilogo ogni giorno, weekend compresi, per ogni tab di lingua (`Daily Report`, `Daily Summary`);
2. dettaglio per caso: causa (tecnica: headset, microfono, unmute, transfer...; comportamento), callback, azioni
   (colonne U:AA delle tab São Paulo, Portugal PT, Portugal SP, Seville, Milan; vista unica in `All Cases`);
3. numeri che tornano con la dashboard (`Daily Summary` D:F e Coverage check: ogni caso è contato, anche
   "reviewed, no action taken" e "not yet reviewed");
4. pop-up/endpoint sopra la soglia del 5%: check tecnico in giornata e causa scritta (`Endpoint Check`);
5. recidivi: in processo sì/no e data dell'ultima azione (`Repeat Offenders` + `Process Tracker`).

```
pip install openpyxl
python3 build_copy.py "Disconnected Calls - All Markets.xlsx" "Disconnected Calls - All Markets (copy).xlsx"
```

Input: export xlsx del file (File > Scarica > Microsoft Excel). Output: xlsx da caricare su Google Drive e aprire
con Fogli Google (le formule usano ARRAYFORMULA, LET, FILTER, VSTACK...). Il file viene modificato a livello XML:
le parti non toccate (commenti, attività, immagini, viste filtro, formule Google esistenti) restano identiche.

Dopo l'import: scoprire la tab nascosta `Convex`, selezionare A1 e cliccare "Consenti accesso" (IMPORTRANGE del
roster). Senza il roster RAW non trova cost center e lingua e le tab di lingua restano vuote.

Test: `--lo` scrive le stesse formule in sintassi Excel/LibreOffice (formule matrice, prefissi `_xlfn.`) per
ricalcolare il file con LibreOffice; `--test-inputs file.json` aggiunge input di prova (solo per il file di test).
I file xlsx/json non sono versionati (contengono dati dei dipendenti).
