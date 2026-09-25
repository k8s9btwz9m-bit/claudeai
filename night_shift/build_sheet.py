"""Genera il file "R2 Night Shift Performance" (xlsx da importare in Google Sheets).

Input:  export xlsx del file "R2 Overnight Shift" (fogli Schedules, Teableau).
Output: night_shift_performance.xlsx + reference.json (valori attesi per la verifica).

Le formule usano funzioni native di Google Sheets (ARRAYFORMULA, MAP, LAMBDA,
FILTER, SORT, UNIQUE, HSTACK, MAKEARRAY): il file va aperto in Google Sheets,
non in Excel.
"""
import datetime as dt
import json
import re
import sys
from collections import defaultdict

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

SRC = sys.argv[1] if len(sys.argv) > 1 else "r2b.xlsx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "night_shift_performance.xlsx"

PERIOD_START = dt.date(2026, 9, 13)
PERIOD_END = dt.date(2026, 9, 26)
SHIFTS = [("18:00-03:00", "18-3"), ("21:00-06:00", "21-6")]
EXCLUDED_TIERS = ["Premium Support"]
TARGETS = [  # KPI, target, direzione, fonte
    ("Throughput", 0.58, "Alto", "EMEA H2'26 Targets - Resolutions 2"),
    ("NPS", 70, "Alto", "EMEA H2'26 Targets - Resolutions 2 Italian"),
    ("Handoff", 0.06, "Basso", "EMEA H2'26 Targets - Resolutions 2"),
    ("Pending", 0.28, "Basso", "EMEA H2'26 Targets - Resolutions 2"),
    ("Reopen", 0.10, "Basso", "EMEA H2'26 Targets - Resolutions 2"),
    ("Recontact", 0.33, "Basso", "EMEA H2'26 Targets - Resolutions 2"),
]
MEASURES = [  # etichetta, nome misura Tableau, colonna in Turni_Log
    ("Assegnazioni completate", "cs_ambassador_case_assignments_completed", "P"),
    ("Casi risolti", "cs_cases_solved_excluding_bot_only", "Q"),
    ("Pend events", "cs_ambassador_pend_events", "R"),
    ("Handoff", "cs_ambassador_transfer_handoffs", "S"),
    ("Reopen", "cs_case_reopen", "T"),
    ("Denominatore Reopen", "cs_cases_solved_excluding_bot_only", "U"),
    ("Recontact", "cs_interaction_specialist_recontacts", "V"),
    ("Risposte NPS", "cs_customer_nps_responses", "W"),
    ("Net promoters NPS", "cs_customer_nps_net_promoters", "X"),
]
SPOT_PARAMS = [
    ("Min assegnazioni per valutare Handoff/Pending", 20),
    ("Min denominatore per valutare Reopen", 15),
    ("Min risposte per valutare NPS", 5),
    ("Fattore k (deviazioni standard)", 1),
    ("Casi da campionare per KPI ROSSO", 3),
    ("Casi da campionare per KPI GIALLO", 1),
]
LISTS = {
    "KPI": ["Handoff", "Pending", "NPS", "Reopen"],
    "Esito": ["Corretto", "Migliorabile", "Non corretto"],
    "Root cause": [
        "Procedura non seguita", "Knowledge gap", "Pend non necessario",
        "Handoff evitabile", "Handoff a fine turno", "Comunicazione poco chiara",
        "Utente non risponde", "Limite tool/sistema", "Policy / fuori controllo agente",
        "Nessun problema",
    ],
    "Azione": ["Coaching 1:1", "Feedback scritto", "Calibrazione team", "Nessuna"],
    "Sì/No": ["Sì", "No"],
}
AGENT_ROWS = 150  # righe massime nei fogli Agenti / Spotcheck

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
SEC_FONT = Font(bold=True, size=12, color="1F3864")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="E7E6E6")
RED = PatternFill("solid", fgColor="F4CCCC")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GREEN = PatternFill("solid", fgColor="D9EAD3")
PCT = "0.0%"
NUM1 = "0.0"
DATE = "dd/mm/yyyy"


def header(ws, row, values, col=1):
    for i, v in enumerate(values):
        c = ws.cell(row, col + i, v)
        c.fill, c.font = HDR_FILL, HDR_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def flag_cf(ws, rng, first_cell):
    for text, fill in (("ROSSO", RED), ("GIALLO", YELLOW), ("VERDE", GREEN)):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first_cell}="{text}"'], fill=fill))


def as_date(x):
    if isinstance(x, dt.datetime):
        return x.date()
    if isinstance(x, dt.date):
        return x
    return dt.date.fromisoformat(str(x)[:10])


# --------------------------------------------------------------------------- input
src = openpyxl.load_workbook(SRC, read_only=True)
sched_rows = [r for r in src["Schedules"].iter_rows(values_only=True)]
sched_hdr = sched_rows[0][:16]
sched = [r[:16] for r in sched_rows[1:] if len(r) > 1 and r[1]]
sched_by_ldap = {r[1]: r for r in sched}

tab_raw = [r for r in src["Teableau"].iter_rows(values_only=True)]
tab = [r[:7] for r in tab_raw if len(r) >= 7 and r[0] and r[0] != "Breakdown Selection 1"]
measure_names = {m for _, m, _ in MEASURES}
# per restare leggeri si caricano solo le misure grezze usate dalle formule
MINI = int(__import__("os").environ.get("MINI", "0"))  # >0: solo N agenti di test (validazione)
tab = [r for r in tab if str(r[4]).strip() in measure_names and r[6]]  # gli zeri non servono ai SUMIF

if MINI:
    keep = sorted({r[0] for r in tab})[:MINI]
    tab = [r for r in tab if r[0] in keep]
night_strings = {s for s, _ in SHIFTS}
code_of = dict(SHIFTS)


def is_night_row(r):
    return any(str(x) in night_strings for x in r[8:15]) and r[4] not in EXCLUDED_TIERS


# Schedules di prova: tutti gli agenti con almeno un turno notturno + 15 righe diurne
# (per dimostrare che Nuova_Settimana le scarta)
sched_night = [r for r in sched if any(str(x) in night_strings for x in r[8:15])]
sched_other = [r for r in sched if r not in sched_night][:15]
sched_sample = sched_night + sched_other

# Righe REALI della settimana 27/09-03/10 (stessa logica di Nuova_Settimana)
week_dates = [as_date(x) for x in sched_hdr[8:15]]
log_rows = []
for r in sched_sample:
    if not is_night_row(r):
        continue
    for d, v in zip(week_dates, r[8:15]):
        log_rows.append([d, r[1], r[4], r[5], r[7], str(v) if v is not None else "", "REALE"])

# Righe di TEST: turni fittizi per gli agenti dell'estrazione di prova (13-26/09)
tab_days = defaultdict(set)
for r in tab:
    tab_days[r[0]].add(as_date(r[5]))
test_agents = sorted({r[0] for r in tab})
for i, a in enumerate(test_agents):
    s = sched_by_ldap.get(a)
    tier, tl, sa = (s[4], s[5], s[7]) if s else ("Resolutions 2 SE", "n.d.", "No")
    d = PERIOD_START
    while d <= PERIOD_END:
        if d in tab_days[a]:
            base = 0 if i % 2 == 0 else 1
            if i % 4 == 3 and d.weekday() in (1, 2):  # rotazione mar/mer
                base = 1 - base
            shift = SHIFTS[base][0]
            log_rows.append([d, a, tier, tl, sa, shift, "TEST"])
        d += dt.timedelta(days=1)

# --------------------------------------------------------------------------- reference
def hours(shift):
    try:
        s, e = int(shift[0:2]), int(shift[6:8])
        assert shift[2] == ":" and shift[5] == "-"
    except Exception:
        return 0, 0
    return (24 - s, e) if e <= s else (e - s, 0)


def code(tier, shift):
    if tier in EXCLUDED_TIERS:
        return "Escluso"
    if shift in code_of:
        return code_of[shift]
    return "Altro" if hours(shift) != (0, 0) else "Off"


ref_log = {}
for d, a, tier, tl, sa, shift, src_ in log_rows:
    ref_log[(a, d)] = (code(tier, shift),) + hours(shift)
tab_val = defaultdict(float)
for r in tab:
    tab_val[(r[0], as_date(r[5]), str(r[4]).strip())] += float(r[6])

agg = defaultdict(lambda: defaultdict(float))
agent_agg = defaultdict(lambda: defaultdict(float))
for d, a, tier, tl, sa, shift, src_ in log_rows:
    c, hd, _ = ref_log[(a, d)]
    prev = ref_log.get((a, d - dt.timedelta(days=1)), ("Off", 0, 0))
    att = "Off" if hd == 0 and prev[2] == 0 else (c if hd >= prev[2] else prev[0])
    if att not in code_of.values() or not (PERIOD_START <= d <= PERIOD_END):
        continue
    for lab, m, _ in MEASURES:
        v = tab_val[(a, d, m)]
        agg[att][lab] += v
        agg["Totale"][lab] += v
        agent_agg[a][lab] += v
    agg[att]["giornate"] += 1
    agg["Totale"]["giornate"] += 1


def kpis(s):
    f = lambda n, dd: round(n / dd, 4) if dd else None
    return {
        "Throughput": f(s["Casi risolti"], s["Assegnazioni completate"]),
        "NPS": f(100 * s["Net promoters NPS"], s["Risposte NPS"]),
        "Handoff": f(s["Handoff"], s["Assegnazioni completate"]),
        "Pending": f(s["Pend events"], s["Assegnazioni completate"]),
        "Reopen": f(s["Reopen"], s["Denominatore Reopen"]),
        "Recontact": f(s["Recontact"], s["Assegnazioni completate"]),
        "Assegnazioni": s["Assegnazioni completate"],
        "Giornate": s["giornate"],
    }


reference = {k: kpis(v) for k, v in agg.items()}
reference["agenti"] = len(agent_agg)
reference["log_rows"] = len(log_rows)
reference["reali_rows"] = sum(1 for r in log_rows if r[6] == "REALE")

# --------------------------------------------------------------------------- workbook
wb = openpyxl.Workbook()
ws_readme = wb.active
ws_readme.title = "Leggimi"
ws_dash = wb.create_sheet("Dashboard")
ws_spot = wb.create_sheet("Spotcheck")
ws_ag = wb.create_sheet("Agenti")
ws_log = wb.create_sheet("Spotcheck_Log")
ws_set = wb.create_sheet("Setup")
ws_tab = wb.create_sheet("Tableau")
ws_sch = wb.create_sheet("Schedules")
ws_new = wb.create_sheet("Nuova_Settimana")
ws_tl = wb.create_sheet("Turni_Log")

C1, C2 = "Setup!$C$7", "Setup!$C$8"  # codici turno
NIGHT_CODES = "Setup!$C$7:$C$12"

# ---- Setup
ws = ws_set
ws["B1"] = "SETUP - le celle gialle sono modificabili"
ws["B1"].font = SEC_FONT
ws["B3"], ws["C3"] = "Data inizio periodo", PERIOD_START
ws["B4"], ws["C4"] = "Data fine periodo", PERIOD_END
for c in ("C3", "C4"):
    ws[c].number_format = DATE
    ws[c].fill = INPUT_FILL
ws["D3"] = "Periodo analizzato in Dashboard / Agenti / Spotcheck (date Tableau = giorno di lavoro)"
header(ws, 6, ["Orario in Schedules", "Codice turno"], col=2)
for i in range(6):
    ws.cell(7 + i, 2).fill = ws.cell(7 + i, 3).fill = INPUT_FILL
for i, (s, c) in enumerate(SHIFTS):
    ws.cell(7 + i, 2, s)
    ws.cell(7 + i, 3, c)
ws["D7"] = "I primi due codici (C7, C8) sono i turni confrontati nello split view"
ws["B14"] = "Tier esclusi"
ws["B14"].font = Font(bold=True)
for i in range(4):
    ws.cell(14 + i, 3).fill = INPUT_FILL
for i, t in enumerate(EXCLUDED_TIERS):
    ws.cell(14 + i, 3, t)
header(ws, 19, ["KPI", "Target", "Direzione", "Fonte"], col=2)
for i, (k, t, d, f) in enumerate(TARGETS):
    ws.cell(20 + i, 2, k)
    c = ws.cell(20 + i, 3, t)
    c.fill = INPUT_FILL
    c.number_format = NUM1 if k == "NPS" else PCT
    ws.cell(20 + i, 4, d)
    ws.cell(20 + i, 5, f)
header(ws, 29, ["Campo", "Nome misura in Tableau (Measure Names)"], col=2)
for i, (lab, m, _) in enumerate(MEASURES):
    ws.cell(30 + i, 2, lab)
    c = ws.cell(30 + i, 3, m)
    c.fill = INPUT_FILL
ws["D35"] = ("TEMPORANEO: il dizionario KPI usa cs_ambassador_case_solves. "
             "Appena è nell'estrazione Tableau, scrivilo in C35.")
ws["D35"].font = Font(bold=True, color="C00000")
header(ws, 41, ["Parametri spotcheck", "Valore"], col=2)
for i, (lab, v) in enumerate(SPOT_PARAMS):
    ws.cell(42 + i, 2, lab)
    c = ws.cell(42 + i, 3, v)
    c.fill = INPUT_FILL
header(ws, 2, list(LISTS.keys()), col=7)
for j, vals in enumerate(LISTS.values()):
    for i, v in enumerate(vals):
        ws.cell(3 + i, 7 + j, v).fill = INPUT_FILL
ws["G1"] = "Liste per i menu a tendina di Spotcheck_Log"
ws["G1"].font = Font(bold=True)
widths(ws, {"A": 2, "B": 44, "C": 40, "D": 12, "E": 38, "F": 3, "G": 12, "H": 14, "I": 30, "J": 18, "K": 8})

# ---- Tableau (incolla)
ws = ws_tab
header(ws, 1, ["Breakdown Selection 1", "Breakdown Selection 2", "Breakdown Selection 3",
               "Breakdown Selection 4", "Measure Names", "Column Breakdown Selection",
               "Measure Values", "Chiave (formula - non toccare)"])
ws["H1"].fill = CALC_FILL
ws["H1"].font = Font(bold=True)
for i, r in enumerate(tab, start=2):
    for j, v in enumerate(r, start=1):
        if j in (2, 3, 4):  # colonne "Global" non usate dalle formule
            continue
        ws.cell(i, j, v.date() if isinstance(v, dt.datetime) else v)
ws.column_dimensions["F"].number_format = DATE
ws["H2"] = ('=ARRAYFORMULA(IF(A2:A="","",IFERROR(A2:A&"|"&INT(IF(ISNUMBER(F2:F),F2:F,'
            'DATEVALUE(F2:F)))&"|"&TRIM(E2:E),"")))')
widths(ws, {"A": 24, "E": 42, "F": 14, "G": 12, "H": 60})
ws.freeze_panes = "A2"

# ---- Schedules (incolla)
ws = ws_sch
for j, v in enumerate(sched_hdr, start=1):
    ws.cell(1, j, v)
    if isinstance(v, dt.datetime):
        ws.cell(1, j).number_format = DATE
header(ws, 1, [c.value for c in ws[1]])
for j in range(9, 16):
    ws.cell(1, j).number_format = DATE
for i, r in enumerate(sched_sample, start=2):
    for j, v in enumerate(r, start=1):
        ws.cell(i, j, v)
widths(ws, {"B": 24, "E": 18, "F": 20})
ws.freeze_panes = "C2"

# ---- Nuova_Settimana (formula)
ws = ws_new
header(ws, 1, ["Data", "LDAP", "Tier", "Team Lead", "Special Assignment", "Turno (orario)", "Fonte"])
ws["I1"] = ("Copia A2:G (Incolla speciale > solo valori) in fondo a Turni_Log. "
            "Contiene solo gli agenti con almeno un turno notturno in Schedules, esclusi i tier di Setup.")
ws["I1"].font = Font(italic=True, color="7F7F7F")
ws["A2"] = (
    '=IFERROR(LET(ld,Schedules!B2:B3000,tr,Schedules!E2:E3000,tl,Schedules!F2:F3000,'
    'sa,Schedules!H2:H3000,sh,Schedules!I2:O3000,'
    'dts,MAP(Schedules!I1:O1,LAMBDA(x,IF(ISNUMBER(x),x,DATEVALUE(x)))),'
    'nm,FILTER(Setup!B7:B12,Setup!B7:B12<>""),'
    'keep,BYROW(sh,LAMBDA(r,SUMPRODUCT(--ISNUMBER(MATCH(r,nm,0))))),'
    'idx,FILTER(SEQUENCE(ROWS(ld)),ld<>"",keep>0,COUNTIF(Setup!C14:C17,tr)=0),'
    'MAKEARRAY(ROWS(idx)*7,7,LAMBDA(i,j,LET(r,INDEX(idx,INT((i-1)/7)+1),d,MOD(i-1,7)+1,'
    'CHOOSE(j,INDEX(dts,1,d),INDEX(ld,r),INDEX(tr,r),INDEX(tl,r),INDEX(sa,r),INDEX(sh,r,d),"REALE"))))),'
    '"Nessun turno notturno trovato in Schedules")'
)
ws.column_dimensions["A"].number_format = DATE
widths(ws, {"A": 12, "B": 24, "C": 18, "D": 20, "E": 12, "F": 14, "G": 8})
ws.freeze_panes = "A2"

# ---- Turni_Log
ws = ws_tl
log_hdr = ["Data", "LDAP", "Tier", "Team Lead", "Special Assignment", "Turno (orario)", "Fonte",
           "Chiave", "Settimana (dom)", "Codice turno", "Ore su giorno D", "Ore su giorno D+1",
           "Codice turno D-1", "Ore turno D-1 su D", "Turno attribuito (giorno Tableau)"]
log_hdr += [lab for lab, _, _ in MEASURES] + ["Nel periodo (notte)"]
header(ws, 1, log_hdr)
for j in range(8, len(log_hdr) + 1):
    ws.cell(1, j).fill = CALC_FILL
    ws.cell(1, j).font = Font(bold=True)
for i, r in enumerate(log_rows, start=2):
    for j, v in enumerate(r, start=1):
        ws.cell(i, j, v)
    ws.cell(i, 1).number_format = DATE
ws.column_dimensions["A"].number_format = DATE
ws.column_dimensions["I"].number_format = DATE
VALID = 'REGEXMATCH(F2:F&"","^\\d\\d:\\d\\d-\\d\\d:\\d\\d$")'
ST, EN = "VALUE(LEFT(F2:F,2))", "VALUE(MID(F2:F,7,2))"
ws["H2"] = '=ARRAYFORMULA(IF(B2:B="","",B2:B&"|"&INT(A2:A)))'
ws["I2"] = '=ARRAYFORMULA(IF(B2:B="","",A2:A-WEEKDAY(A2:A)+1))'
ws["J2"] = ('=MAP(B2:B,C2:C,F2:F,LAMBDA(b,c,f,IF(b="","",IF(COUNTIF(Setup!$C$14:$C$17,c)>0,"Escluso",'
            'IFERROR(VLOOKUP(f,Setup!$B$7:$C$12,2,0),IF(REGEXMATCH(f&"","^\\d\\d:\\d\\d-\\d\\d:\\d\\d$"),"Altro","Off"))))))')
ws["K2"] = f'=ARRAYFORMULA(IF(B2:B="","",IF({VALID},IF({EN}<={ST},24-{ST},{EN}-{ST}),0)))'
ws["L2"] = f'=ARRAYFORMULA(IF(B2:B="","",IF({VALID},IF({EN}<={ST},{EN},0),0)))'
ws["M2"] = ('=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B&"|"&(INT(A2:A)-1),'
            'HSTACK(H2:H,J2:J),2,0),"Off")))')
ws["N2"] = ('=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B&"|"&(INT(A2:A)-1),'
            'HSTACK(H2:H,L2:L),2,0),0)))')
ws["O2"] = '=ARRAYFORMULA(IF(B2:B="","",IF((K2:K=0)*(N2:N=0),"Off",IF(K2:K>=N2:N,J2:J,M2:M))))'
IS_NIGHT = f"OR(o={C1},o={C2})"
for i, (lab, m, col) in enumerate(MEASURES):
    ws[f"{col}2"] = (f'=MAP(H2:H,O2:O,LAMBDA(h,o,IF(h="","",IF({IS_NIGHT},'
                     f'SUMIF(Tableau!$H$2:$H,h&"|"&Setup!$C${30 + i},Tableau!$G$2:$G),0))))')
ws["Y2"] = ('=ARRAYFORMULA(IF(B2:B="","",IF((A2:A>=Setup!$C$3)*(A2:A<=Setup!$C$4)*'
            f'(((O2:O={C1})+(O2:O={C2}))>0),1,0)))')
widths(ws, {"A": 11, "B": 24, "C": 17, "D": 18, "E": 10, "F": 13, "G": 8, "H": 30, "I": 11,
            "J": 9, "O": 12})
ws.freeze_panes = "C2"

# ---- Dashboard
ws = ws_dash
ws["B1"] = "R2 NIGHT SHIFT - PERFORMANCE"
ws["B1"].font = Font(bold=True, size=14, color="1F3864")
ws["B2"] = '="Periodo: "&TEXT(Setup!C3,"dd/mm/yyyy")&" - "&TEXT(Setup!C4,"dd/mm/yyyy")&"   (modifica in Setup)"'
header(ws, 4, ["KPI", "Totale notte", "=Setup!C7", "=Setup!C8", '="Δ "&Setup!C7&" vs "&Setup!C8',
               "Target", "Totale vs target"], col=2)
# somme di base (riga 30+)
SUM_ROW = 31
ws.cell(SUM_ROW - 1, 2, "Somme di base (non modificare)").font = SEC_FONT
header(ws, SUM_ROW, ["Misura", "Totale notte", "=Setup!C7", "=Setup!C8"], col=2)
sum_ref = {}
for i, (lab, m, col) in enumerate(MEASURES):
    r = SUM_ROW + 1 + i
    ws.cell(r, 2, lab)
    rng = f"Turni_Log!${col}$2:${col}"
    ws.cell(r, 3, f"=SUMIFS({rng},Turni_Log!$Y$2:$Y,1)")
    ws.cell(r, 4, f"=SUMIFS({rng},Turni_Log!$Y$2:$Y,1,Turni_Log!$O$2:$O,{C1})")
    ws.cell(r, 5, f"=SUMIFS({rng},Turni_Log!$Y$2:$Y,1,Turni_Log!$O$2:$O,{C2})")
    sum_ref[lab] = r
KPI_DEF = {  # num, den, moltiplicatore
    "Throughput": ("Casi risolti", "Assegnazioni completate", ""),
    "NPS": ("Net promoters NPS", "Risposte NPS", "100*"),
    "Handoff": ("Handoff", "Assegnazioni completate", ""),
    "Pending": ("Pend events", "Assegnazioni completate", ""),
    "Reopen": ("Reopen", "Denominatore Reopen", ""),
    "Recontact": ("Recontact", "Assegnazioni completate", ""),
}
kpi_row = {}
for i, (k, t, d, f) in enumerate(TARGETS):
    r = 5 + i
    kpi_row[k] = r
    num, den, mul = KPI_DEF[k]
    ws.cell(r, 2, k).font = Font(bold=True)
    for col in "CDE":
        ws[f"{col}{r}"] = f'=IFERROR({mul}{col}{sum_ref[num]}/{col}{sum_ref[den]},"n.d.")'
    ws[f"F{r}"] = f'=IFERROR(D{r}-E{r},"")'
    ws[f"G{r}"] = f"=Setup!C{20 + i}"
    ws[f"H{r}"] = (f'=IF(ISNUMBER(C{r}),IF(Setup!D{20 + i}="Alto",IF(C{r}>=G{r},"In target","Fuori target"),'
                   f'IF(C{r}<=G{r},"In target","Fuori target")),"")')
    fmt = NUM1 if k == "NPS" else PCT
    for col in "CDEFG":
        ws[f"{col}{r}"].number_format = fmt
ws.conditional_formatting.add("H5:H10", FormulaRule(formula=['H5="Fuori target"'], fill=RED))
ws.conditional_formatting.add("H5:H10", FormulaRule(formula=['H5="In target"'], fill=GREEN))
ws["B12"] = "Volumi"
ws["B12"].font = SEC_FONT
vols = [
    ("Assegnazioni completate", "SUM", "Assegnazioni completate"),
    ("Casi risolti", "SUM", "Casi risolti"),
    ("Risposte NPS", "SUM", "Risposte NPS"),
]
r = 13
for lab, _, src_lab in vols:
    ws.cell(r, 2, lab)
    for col in "CDE":
        ws[f"{col}{r}"] = f"={col}{sum_ref[src_lab]}"
    r += 1
ws.cell(r, 2, "Giornate-agente (turni notturni)")
ws[f"C{r}"] = "=COUNTIF(Turni_Log!$Y$2:$Y,1)"
ws[f"D{r}"] = f"=COUNTIFS(Turni_Log!$Y$2:$Y,1,Turni_Log!$O$2:$O,{C1})"
ws[f"E{r}"] = f"=COUNTIFS(Turni_Log!$Y$2:$Y,1,Turni_Log!$O$2:$O,{C2})"
r += 1
ws.cell(r, 2, "Agenti distinti")
ws[f"C{r}"] = '=IFERROR(ROWS(UNIQUE(FILTER(Turni_Log!$B$2:$B,Turni_Log!$Y$2:$Y=1))),0)'
ws[f"D{r}"] = f'=IFERROR(ROWS(UNIQUE(FILTER(Turni_Log!$B$2:$B,Turni_Log!$Y$2:$Y=1,Turni_Log!$O$2:$O={C1}))),0)'
ws[f"E{r}"] = f'=IFERROR(ROWS(UNIQUE(FILTER(Turni_Log!$B$2:$B,Turni_Log!$Y$2:$Y=1,Turni_Log!$O$2:$O={C2}))),0)'
r += 1
ws.cell(r, 2, "Assegnazioni per giornata-agente")
ws[f"C{r}"] = f'=IFERROR(C13/C16,"n.d.")'
ws[f"D{r}"] = f'=IFERROR(D13/D16,"n.d.")'
ws[f"E{r}"] = f'=IFERROR(E13/E16,"n.d.")'
for col in "CDE":
    ws[f"{col}{r}"].number_format = NUM1
ws["B20"] = ("Nota: una giornata Tableau (giorno D) è attribuita al turno che vi lavora più ore: "
             "18-3 del giorno D (6h) oppure 21-6 del giorno D-1 (6h). "
             "Special Assignment inclusi; tier esclusi da Setup.")
ws["B20"].font = Font(italic=True, color="7F7F7F")

# Trend settimanale
TR = 44
ws.cell(TR - 1, 2, "Trend settimanale (ultime 8 settimane fino alla data fine)").font = SEC_FONT
tr_hdr = ["Settimana (dom)"]
for k, *_ in TARGETS:
    tr_hdr += [f"{k} Tot", f"{k} 18-3", f"{k} 21-6"]
tr_hdr += ["Assegn. Tot", "Assegn. 18-3", "Assegn. 21-6"]
header(ws, TR, tr_hdr, col=2)
W1, W8 = TR + 1, TR + 8
ws.cell(W1, 2, f"=SEQUENCE(8,1,Setup!$C$4-WEEKDAY(Setup!$C$4)+1-49,7)")
for r in range(W1, W8 + 1):
    ws.cell(r, 2).number_format = DATE
WK = f"$B${W1}:$B${W8}"


def s(col, code_ref=None):
    base = f"SUMIFS(Turni_Log!${col}$2:${col},Turni_Log!$I$2:$I,w,Turni_Log!$O$2:$O,"
    if code_ref:
        return base + code_ref + ")"
    return f"({base}{C1})+{base}{C2}))"


c = 3
for k, *_ in TARGETS:
    num, den, mul = KPI_DEF[k]
    ncol = next(cc for lab, _, cc in MEASURES if lab == num)
    dcol = next(cc for lab, _, cc in MEASURES if lab == den)
    for code_ref in (None, C1, C2):
        ws.cell(W1, c, f'=MAP({WK},LAMBDA(w,IFERROR({mul}{s(ncol, code_ref)}/{s(dcol, code_ref)},"")))')
        L = openpyxl.utils.get_column_letter(c)
        ws.column_dimensions[L].number_format = NUM1 if k == "NPS" else PCT
        c += 1
for code_ref in (None, C1, C2):
    ws.cell(W1, c, f"=MAP({WK},LAMBDA(w,{s('P', code_ref)}))")
    c += 1
widths(ws, {"A": 2, "B": 34, "C": 13, "D": 13, "E": 13, "F": 16, "G": 11, "H": 15})
for col in range(9, 30):
    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 11

# ---- Agenti
ws = ws_ag
ag_hdr = ["LDAP", "Team Lead", "Giorni " + SHIFTS[0][1], "Giorni " + SHIFTS[1][1], "Turno prevalente"]
ag_hdr += [lab for lab, _, _ in MEASURES]
ag_hdr += ["Throughput", "NPS", "Handoff", "Pending", "Reopen", "Recontact",
           "Flag Handoff", "Flag Pending", "Flag NPS", "Flag Reopen", "Punteggio", "Casi da campionare"]
header(ws, 1, ag_hdr)
N = AGENT_ROWS + 1
A = f"A2:A{N}"
LOG = "Turni_Log!$B$2:$B"
PER = "Turni_Log!$Y$2:$Y"
ws["A2"] = f'=IFERROR(SORT(UNIQUE(FILTER({LOG},{PER}=1))),"")'
ws["B2"] = f'=MAP({A},LAMBDA(a,IF(a="","",IFERROR(VLOOKUP(a,Turni_Log!$B$2:$D,3,0),""))))'
ws["C2"] = f'=MAP({A},LAMBDA(a,IF(a="","",COUNTIFS({LOG},a,{PER},1,Turni_Log!$O$2:$O,{C1}))))'
ws["D2"] = f'=MAP({A},LAMBDA(a,IF(a="","",COUNTIFS({LOG},a,{PER},1,Turni_Log!$O$2:$O,{C2}))))'
ws["E2"] = (f'=ARRAYFORMULA(IF({A}="","",IF((C2:C{N}>0)*(D2:D{N}>0),"Misto",'
            f'IF(C2:C{N}>0,{C1},{C2}))))')
for i, (lab, m, col) in enumerate(MEASURES):
    L = openpyxl.utils.get_column_letter(6 + i)
    ws[f"{L}2"] = f'=MAP({A},LAMBDA(a,IF(a="","",SUMIFS(Turni_Log!${col}$2:${col},{LOG},a,{PER},1))))'
# F Assegn, G Risolti, H Pend, I Handoff, J Reopen, K Den reopen, L Recontact, M Risp NPS, N Net NPS
g = lambda c: f"{c}2:{c}{N}"
MINA, MINR, MINN = "Setup!$C$42", "Setup!$C$43", "Setup!$C$44"
ws["O2"] = f'=ARRAYFORMULA(IF({A}="","",IF({g("F")}>0,{g("G")}/{g("F")},"n.d.")))'
ws["P2"] = f'=ARRAYFORMULA(IF({A}="","",IF({g("M")}>={MINN},100*{g("N")}/{g("M")},"n.s.")))'
ws["Q2"] = f'=ARRAYFORMULA(IF({A}="","",IF({g("F")}>={MINA},{g("I")}/{g("F")},"n.s.")))'
ws["R2"] = f'=ARRAYFORMULA(IF({A}="","",IF({g("F")}>={MINA},{g("H")}/{g("F")},"n.s.")))'
ws["S2"] = f'=ARRAYFORMULA(IF({A}="","",IF({g("K")}>={MINR},{g("J")}/{g("K")},"n.s.")))'
ws["T2"] = f'=ARRAYFORMULA(IF({A}="","",IF({g("F")}>0,{g("L")}/{g("F")},"n.d.")))'


def flag(valcol, tgt_row, mean_cell, higher_is_worse):
    v = g(valcol)
    tgt = f"Setup!$C${tgt_row}"
    sd = f'IFERROR(STDEV(FILTER({v},ISNUMBER({v}))),0)'
    if higher_is_worse:
        red = f"({v}>{tgt})*({v}>{mean_cell}+Setup!$C$45*{sd})"
        yel = f"{v}>{tgt}"
    else:
        red = f"({v}<{tgt})*({v}<{mean_cell}-Setup!$C$45*{sd})"
        yel = f"{v}<{tgt}"
    return (f'=ARRAYFORMULA(IF({A}="","",IF(ISNUMBER({v}),IF({red},"ROSSO",IF({yel},"GIALLO","VERDE")),"n.s.")))')


ws["U2"] = flag("Q", 22, f"Dashboard!$C${kpi_row['Handoff']}", True)
ws["V2"] = flag("R", 23, f"Dashboard!$C${kpi_row['Pending']}", True)
ws["W2"] = flag("P", 21, f"Dashboard!$C${kpi_row['NPS']}", False)
ws["X2"] = flag("S", 24, f"Dashboard!$C${kpi_row['Reopen']}", True)
cnt = lambda t: "+".join(f'({g(c)}="{t}")' for c in "UVWX")
ws["Y2"] = f'=ARRAYFORMULA(IF({A}="","",3*({cnt("ROSSO")})+({cnt("GIALLO")})))'
ws["Z2"] = (f'=ARRAYFORMULA(IF({A}="","",Setup!$C$46*({cnt("ROSSO")})+Setup!$C$47*({cnt("GIALLO")})))')
for c in "OQRST":
    ws.column_dimensions[c].number_format = PCT
ws.column_dimensions["P"].number_format = NUM1
flag_cf(ws, f"U2:X{N}", "U2")
widths(ws, {"A": 24, "B": 20, "E": 12})
ws.freeze_panes = "B2"

# ---- Spotcheck
ws = ws_spot
ws["A1"] = "SPOTCHECK - focus su Handoff, Pending, NPS e Reopen"
ws["A1"].font = Font(bold=True, size=14, color="1F3864")
ws["A2"] = '="Periodo: "&TEXT(Setup!C3,"dd/mm/yyyy")&" - "&TEXT(Setup!C4,"dd/mm/yyyy")&"  |  ordinato per priorità"'
ws["A3"] = ("ROSSO = fuori target e oltre media notte ± k·dev.std  |  GIALLO = fuori target  |  "
            "VERDE = in target  |  n.s. = volume insufficiente (soglie in Setup). "
            "Registra ogni caso controllato in Spotcheck_Log.")
ws["A3"].font = Font(italic=True, color="7F7F7F")
sp_hdr = ["LDAP", "Team Lead", "Turno", "Assegn.", "Handoff", "Flag", "Pending", "Flag", "NPS",
          "Risp. NPS", "Flag", "Reopen", "Flag", "Punteggio", "Casi da campionare",
          "Spotcheck registrati", "Stato"]
header(ws, 4, sp_hdr)
M = f"{AGENT_ROWS + 1}"
cols = ["A", "B", "E", "F", "Q", "U", "R", "V", "P", "M", "W", "S", "X", "Y", "Z"]
stack = ",".join(f"Agenti!{c}2:{c}{M}" for c in cols)
ws["A5"] = f'=IFERROR(SORT(FILTER(HSTACK({stack}),Agenti!A2:A{M}<>""),14,FALSE,15,FALSE,1,TRUE),"")'
E = 5 + AGENT_ROWS - 1
ws["P5"] = (f'=MAP(A5:A{E},LAMBDA(a,IF(a="","",COUNTIFS(Spotcheck_Log!$B$2:$B,a,'
            f'Spotcheck_Log!$A$2:$A,">="&Setup!$C$3,Spotcheck_Log!$A$2:$A,"<="&Setup!$C$4))))')
ws["Q5"] = (f'=ARRAYFORMULA(IF(A5:A{E}="","",IF(O5:O{E}=0,"-",IF(P5:P{E}>=O5:O{E},"Completato",'
            f'"Da fare ("&(O5:O{E}-P5:P{E})&")"))))')
for c in "EGL":
    ws.column_dimensions[c].number_format = PCT
ws.column_dimensions["I"].number_format = NUM1
for c in "FHKM":
    flag_cf(ws, f"{c}5:{c}{E}", f"{c}5")
ws.conditional_formatting.add(f"Q5:Q{E}", FormulaRule(formula=['LEFT(Q5,7)="Da fare"'], fill=RED))
ws.conditional_formatting.add(f"Q5:Q{E}", FormulaRule(formula=['Q5="Completato"'], fill=GREEN))
# riepilogo root cause
ws["S4"], ws["T4"] = "Root cause (casi del periodo)", "N. casi"
header(ws, 4, ["Root cause (casi del periodo)", "N. casi", "% sul totale"], col=19)
for i in range(len(LISTS["Root cause"])):
    r = 5 + i
    ws[f"S{r}"] = f"=Setup!I{3 + i}"
    ws[f"T{r}"] = (f'=COUNTIFS(Spotcheck_Log!$G$2:$G,S{r},Spotcheck_Log!$A$2:$A,">="&Setup!$C$3,'
                   f'Spotcheck_Log!$A$2:$A,"<="&Setup!$C$4)')
    ws[f"U{r}"] = f'=IFERROR(T{r}/SUM($T$5:$T${4 + len(LISTS["Root cause"])}),0)'
    ws[f"U{r}"].number_format = PCT
widths(ws, {"A": 24, "B": 20, "C": 9, "D": 9, "N": 10, "O": 11, "P": 11, "Q": 14, "R": 3,
            "S": 32, "T": 9, "U": 11})
for c in "EFGHIJKLM":
    ws.column_dimensions[c].width = 9
ws.freeze_panes = "B5"

# ---- Spotcheck_Log
ws = ws_log
log2_hdr = ["Data del caso", "LDAP", "Turno", "KPI", "Case ID / link", "Esito", "Root cause",
            "Azione", "Coaching fatto?", "Data spotcheck", "Eseguito da (TL)", "Note"]
header(ws, 1, log2_hdr)
ws.column_dimensions["A"].number_format = DATE
ws.column_dimensions["J"].number_format = DATE
dv = [
    ("B", f"=Agenti!$A$2:$A${AGENT_ROWS + 1}"),
    ("C", "=Setup!$C$7:$C$12"),
    ("D", f"=Setup!$G$3:$G${2 + len(LISTS['KPI'])}"),
    ("F", f"=Setup!$H$3:$H${2 + len(LISTS['Esito'])}"),
    ("G", f"=Setup!$I$3:$I${2 + len(LISTS['Root cause'])}"),
    ("H", f"=Setup!$J$3:$J${2 + len(LISTS['Azione'])}"),
    ("I", f"=Setup!$K$3:$K${2 + len(LISTS['Sì/No'])}"),
]
for col, formula in dv:
    v = DataValidation(type="list", formula1=formula, allow_blank=True)
    v.add(f"{col}2:{col}2000")
    ws.add_data_validation(v)
widths(ws, {"A": 12, "B": 24, "C": 8, "D": 10, "E": 34, "F": 13, "G": 30, "H": 18, "I": 10,
            "J": 12, "K": 18, "L": 40})
ws.freeze_panes = "A2"

# ---- Leggimi
ws = ws_readme
lines = [
    ("R2 NIGHT SHIFT - PERFORMANCE (18-3 vs 21-6)", "title"),
    ("", None),
    ("A COSA SERVE", "sec"),
    ("Analizza le performance dei turni notturni R2 (18-3 e 21-6) insieme e separati: "
     "Throughput, NPS, Handoff, Pending, Reopen, Recontact. Il foglio Spotcheck dice a ogni TL quali agenti "
     "e quanti casi controllare, con focus su Handoff, Pending, NPS e Reopen.", None),
    ("", None),
    ("PROCEDURA SETTIMANALE", "sec"),
    ("1. Tableau (Dynamic Slicing, breakdown per LDAP, colonna = Day): esporta i dati e incollali in 'Tableau' da A1. "
     "Sostituisci solo A:G, la colonna H è una formula.", None),
    ("   Misure necessarie (nomi in Setup C30:C38): assignments_completed, cases_solved_excluding_bot_only, "
     "pend_events, transfer_handoffs, case_reopen, case_solves, interaction_specialist_recontacts, "
     "nps_responses, nps_net_promoters.", None),
    ("2. Schedules: incolla il foglio Schedules della nuova settimana in 'Schedules' da A1 (stesse colonne).", None),
    ("3. 'Nuova_Settimana' si calcola da sola: copia A2:G e incolla SOLO VALORI in fondo a 'Turni_Log' "
     "(Turni_Log è lo storico: non cancellare le settimane precedenti).", None),
    ("4. In 'Setup' imposta Data inizio / Data fine del periodo da analizzare.", None),
    ("5. Leggi 'Dashboard' e 'Spotcheck'; i TL registrano i controlli in 'Spotcheck_Log'.", None),
    ("", None),
    ("COME VIENE ATTRIBUITO IL TURNO", "sec"),
    ("Tableau ha solo il giorno, mentre i turni passano la mezzanotte. La giornata Tableau D di un agente viene "
     "attribuita al turno che vi lavora più ore: il turno iniziato il giorno D (18-3 = 6h, 21-6 = 3h) oppure "
     "quello iniziato il giorno D-1 (18-3 = 3h, 21-6 = 6h). A parità vince il turno del giorno D. Se il turno "
     "prevalente è diurno la giornata è 'Altro' ed è esclusa.", None),
    ("Special Assignment inclusi. Tier esclusi in Setup (ora: Premium Support).", None),
    ("", None),
    ("DEFINIZIONI KPI (Data Dictionary - Metrics)", "sec"),
    ("Throughput = cs_cases_solved_excluding_bot_only / cs_ambassador_case_assignments_completed", None),
    ("NPS = 100 × cs_customer_nps_net_promoters / cs_customer_nps_responses", None),
    ("Handoff = cs_ambassador_transfer_handoffs / cs_ambassador_case_assignments_completed", None),
    ("Pending = cs_ambassador_pend_events / cs_ambassador_case_assignments_completed", None),
    ("Reopen = cs_case_reopen / cs_ambassador_case_solves", None),
    ("Recontact = cs_interaction_specialist_recontacts / cs_ambassador_case_assignments_completed", None),
    ("I KPI si calcolano sempre come somma numeratori / somma denominatori (mai media di percentuali).", None),
    ("Target: EMEA 2026 H2 CS Delivery Targets, Resolutions 2 (NPS: Resolutions 2 Italian). "
     "Modificabili in Setup.", None),
    ("", None),
    ("SPOTCHECK: COME FUNZIONA", "sec"),
    ("Semaforo per agente su Handoff, Pending, NPS, Reopen: ROSSO = fuori target e anche oltre la media della notte "
     "di k deviazioni standard (outlier vero); GIALLO = solo fuori target; VERDE = in target; n.s. = volume troppo "
     "basso per giudicare (soglie in Setup).", None),
    ("Punteggio = 3 × ROSSI + GIALLI: la lista è ordinata per priorità. "
     "Casi da campionare = 3 per ogni ROSSO + 1 per ogni GIALLO (modificabile).", None),
    ("Ogni caso controllato va registrato in Spotcheck_Log: la colonna 'Stato' del foglio Spotcheck mostra "
     "cosa resta da fare, a destra c'è il riepilogo delle root cause.", None),
    ("", None),
    ("COSA CONTROLLARE NEI CASI", "sec"),
    ("Handoff: era evitabile? Avviene a fine turno (03:00 / 06:00)? Nei notturni è il rischio tipico: "
     "casi passati al turno successivo invece di chiuderli o metterli in pend correttamente.", None),
    ("Pending: il pend era necessario? Il next step è chiaro per l'utente e per chi riprende il caso? "
     "Si ripete sullo stesso caso? (è anche il punto di partenza per misurare gli 'unresponsive users').", None),
    ("NPS: leggi i verbatim dei detractor (0-6) e distingui colpa agente / policy / prodotto.", None),
    ("Reopen: perché il caso è stato riaperto? Soluzione incompleta, comunicazione poco chiara, chiusura "
     "prematura (es. chiuso a fine turno senza conferma utente).", None),
    ("Altre idee: calibrazione settimanale tra TL su 2-3 casi comuni; confronto 18-3 vs 21-6 sulle fasce di "
     "sovrapposizione 21-03; trend a 4 settimane dopo il coaching per vedere se il KPI migliora.", None),
    ("", None),
    ("NOTE APERTE", "sec"),
    ("- Reopen: cs_ambassador_case_solves non è ancora nell'estrazione. Per ora il denominatore è "
     "cs_cases_solved_excluding_bot_only: appena disponibile, cambiare Setup C35.", None),
    ("- Unresponsive users: in sospeso, nessun campo disponibile. La root cause 'Utente non risponde' in "
     "Spotcheck_Log permette già di iniziare a contarli.", None),
    ("- Limiti: le formule leggono fino a 5000 righe di Turni_Log (circa 35-40 settimane) e 40000 righe di "
     "Tableau. Oltre, archivia le settimane vecchie in un altro file.", None),
    ("- DATI DI TEST: Turni_Log contiene righe con Fonte = TEST (turni fittizi assegnati agli agenti "
     "dell'estrazione di prova 13-24/09) per provare le formule. Prima dell'uso reale filtra Fonte = TEST, "
     "cancella quelle righe e sostituisci i dati in 'Tableau'.", None),
]
for i, (t, kind) in enumerate(lines, start=1):
    c = ws.cell(i, 2, t)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if kind == "title":
        c.font = Font(bold=True, size=14, color="1F3864")
    elif kind == "sec":
        c.font = SEC_FONT
widths(ws, {"A": 2, "B": 130})

# Google Sheets non importa gli intervalli aperti (es. B2:B): si chiudono con un limite di righe
OPEN_RANGE = re.compile(r"((?:[A-Za-z_]+!)?)(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3})(?![$0-9A-Za-z_(])")
ROW_LIMIT = {"Tableau": 40000}


def close_ranges(formula, sheet):
    def repl(m):
        target = m.group(1)[:-1] if m.group(1) else sheet
        return f"{m.group(1)}{m.group(2)}:{m.group(3)}{ROW_LIMIT.get(target, 5000)}"
    return OPEN_RANGE.sub(repl, formula)


for sh in wb.worksheets:
    for row in sh.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.value = close_ranges(c.value, sh.title)
# le formule che si espandono richiedono abbastanza righe nella griglia: segnaposto in fondo
for sh, row, col in ((ws_tab, 40001, "J"), (ws_tl, 5001, "AA"), (ws_new, 1101, "I"),
                     (ws_ag, AGENT_ROWS + 2, "AB"), (ws_spot, AGENT_ROWS + 6, "R")):
    cell = sh[f"{col}{row}"]
    cell.value = "fine area formule"
    cell.font = Font(italic=True, color="BFBFBF")
for sh in wb.worksheets:
    for dv in sh.data_validations.dataValidation:
        dv.formula1 = close_ranges(dv.formula1, sh.title)

wb.save(OUT)
with open("reference.json", "w") as fh:
    json.dump(reference, fh, indent=1, default=str)
print(json.dumps(reference, indent=1, default=str))
print("tableau rows", len(tab), "sched rows", len(sched_sample), "log rows", len(log_rows))
