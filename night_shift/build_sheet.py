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

PERIOD_START = dt.date(2026, 9, 20)  # una settimana: il confronto usa la settimana prima (13-19/09)
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
FOCUS = ["Handoff", "Pending", "NPS", "Reopen"]
WEEKDAYS = ["Dom", "Lun", "Mar", "Mer", "Gio", "Ven", "Sab"]  # WEEKDAY() 1..7
MEDALLIA_COLS = [  # campo, intestazione di colonna attesa nell'export Medallia (da verificare)
    ("LDAP agente", "Agent LDAP"),
    ("Data risposta", "Response Date"),
    ("Punteggio (0-10)", "Likelihood to Recommend"),
    ("Verbatim", "Comment"),
    ("Case ID", "Case ID"),
]
DETRACTOR_MAX = 6

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
    d = min(tab_days[a])  # anche la settimana prima del periodo, per il confronto
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

PREV_START = PERIOD_START - (PERIOD_END - PERIOD_START) - dt.timedelta(days=1)
PREV_END = PERIOD_START - dt.timedelta(days=1)
agg = defaultdict(lambda: defaultdict(float))
agg_prev = defaultdict(lambda: defaultdict(float))
agent_agg = defaultdict(lambda: defaultdict(float))
agent_shift = defaultdict(lambda: defaultdict(float))  # (agente, turno) -> misure
load_wd = defaultdict(lambda: [0, 0.0])  # (turno, giorno inizio turno 1=dom) -> [giornate, assegn.]
for d, a, tier, tl, sa, shift, src_ in log_rows:
    c, hd, _ = ref_log[(a, d)]
    prev = ref_log.get((a, d - dt.timedelta(days=1)), ("Off", 0, 0))
    att = "Off" if hd == 0 and prev[2] == 0 else (c if hd >= prev[2] else prev[0])
    if att not in code_of.values():
        continue
    if PREV_START <= d <= PREV_END:
        for lab, m, _ in MEASURES:
            agg_prev[att][lab] += tab_val[(a, d, m)]
            agg_prev["Totale"][lab] += tab_val[(a, d, m)]
        agg_prev[att]["giornate"] += 1
        agg_prev["Totale"]["giornate"] += 1
    if not (PERIOD_START <= d <= PERIOD_END):
        continue
    for lab, m, _ in MEASURES:
        v = tab_val[(a, d, m)]
        agg[att][lab] += v
        agg["Totale"][lab] += v
        agent_agg[a][lab] += v
        agent_shift[(a, att)][lab] += v
    agent_shift[(a, att)]["giornate"] += 1
    agg[att]["giornate"] += 1
    agg["Totale"]["giornate"] += 1
    start = d if hd >= prev[2] else d - dt.timedelta(days=1)
    wd = (start.weekday() + 1) % 7 + 1
    load_wd[(att, wd)][0] += 1
    load_wd[(att, wd)][1] += tab_val[(a, d, MEASURES[0][1])]


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
reference["prev"] = {k: kpis(v) for k, v in agg_prev.items()}
reference["carico_giorno"] = {f"{c} {WEEKDAYS[wd - 1]}": [n, round(x / n, 2)] for (c, wd), (n, x) in sorted(load_wd.items())}
mixed = [a for a in agent_agg if all(agent_shift[(a, c)]["giornate"] for _, c in SHIFTS)]
mix_tot = {c: defaultdict(float) for _, c in SHIFTS}
for a in mixed:
    for _, c in SHIFTS:
        for lab, v in agent_shift[(a, c)].items():
            mix_tot[c][lab] += v
reference["rotazione"] = {"agenti_misti": len(mixed), **{c: kpis(v) for c, v in mix_tot.items()}}
tab_agents_period = {r[0] for r in tab if PERIOD_START <= as_date(r[5]) <= PERIOD_END}
reference["tableau_non_in_log"] = len(tab_agents_period - {r[1] for r in log_rows})
reference["agenti"] = len(agent_agg)
reference["log_rows"] = len(log_rows)
reference["reali_rows"] = sum(1 for r in log_rows if r[6] == "REALE")

# --------------------------------------------------------------------------- workbook
wb = openpyxl.Workbook()
ws_readme = wb.active
ws_readme.title = "Leggimi"
ws_dash = wb.create_sheet("Dashboard")
ws_spot = wb.create_sheet("Spotcheck")
ws_vtl = wb.create_sheet("Vista_TL")
ws_ag = wb.create_sheet("Agenti")
ws_rot = wb.create_sheet("Rotazione")
ws_car = wb.create_sheet("Carico")
ws_str = wb.create_sheet("Rossi_Consecutivi")
ws_coa = wb.create_sheet("Coaching")
ws_det = wb.create_sheet("Detractor_NPS")
ws_log = wb.create_sheet("Spotcheck_Log")
ws_set = wb.create_sheet("Setup")
ws_tab = wb.create_sheet("Tableau")
ws_med = wb.create_sheet("Medallia")
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
ws["B27"], ws["C27"] = "Inizio periodo di confronto (calcolato)", "=C3-(C4-C3+1)"
ws["B28"], ws["C28"] = "Fine periodo di confronto (calcolato)", "=C3-1"
for c in ("C27", "C28"):
    ws[c].number_format = DATE
    ws[c].fill = CALC_FILL
ws["D27"] = "Periodo precedente di pari durata (con 1 settimana = settimana precedente)"
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
               "Measure Values", "Chiave (formula - non toccare)", "Data (formula)", "LDAP|data (formula)"])
for c in ("H1", "I1", "J1"):
    ws[c].fill = CALC_FILL
    ws[c].font = Font(bold=True)
for i, r in enumerate(tab, start=2):
    for j, v in enumerate(r, start=1):
        if j in (2, 3, 4):  # colonne "Global" non usate dalle formule
            continue
        ws.cell(i, j, v.date() if isinstance(v, dt.datetime) else v)
ws.column_dimensions["F"].number_format = DATE
ws["H2"] = ('=ARRAYFORMULA(IF(A2:A="","",IFERROR(A2:A&"|"&INT(IF(ISNUMBER(F2:F),F2:F,'
            'DATEVALUE(F2:F)))&"|"&TRIM(E2:E),"")))')
ws["I2"] = '=ARRAYFORMULA(IF(A2:A="","",IFERROR(INT(IF(ISNUMBER(F2:F),F2:F,DATEVALUE(F2:F))),"")))'
ws["J2"] = '=ARRAYFORMULA(IF(A2:A="","",A2:A&"|"&I2:I))'
ws.column_dimensions["I"].number_format = DATE
widths(ws, {"A": 24, "E": 42, "F": 14, "G": 12, "H": 60, "I": 12, "J": 34})
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
log_hdr += [lab for lab, _, _ in MEASURES] + ["Nel periodo (notte)", "Dati Tableau presenti",
                                               "Data inizio turno", "Giorno inizio turno (1=dom)"]
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
ws["Z2"] = ('=MAP(H2:H,Y2:Y,LAMBDA(h,y,IF(h="","",IF(y=1,IF(COUNTIF(Tableau!$J$2:$J,h)>0,1,0),""))))')
ws["AA2"] = '=ARRAYFORMULA(IF(B2:B="","",IF(K2:K>=N2:N,INT(A2:A),INT(A2:A)-1)))'
ws["AB2"] = '=ARRAYFORMULA(IF(B2:B="","",WEEKDAY(AA2:AA)))'
ws.column_dimensions["AA"].number_format = DATE
widths(ws, {"A": 11, "B": 24, "C": 17, "D": 18, "E": 10, "F": 13, "G": 8, "H": 30, "I": 11,
            "J": 9, "O": 12})
ws.freeze_panes = "C2"

# ---- Dashboard
ws = ws_dash
ws["B1"] = "R2 NIGHT SHIFT - PERFORMANCE"
ws["B1"].font = Font(bold=True, size=14, color="1F3864")
ws["B2"] = ('="Periodo: "&TEXT(Setup!C3,"dd/mm/yyyy")&" - "&TEXT(Setup!C4,"dd/mm/yyyy")&'
            '"   |   confronto con: "&TEXT(Setup!C27,"dd/mm/yyyy")&" - "&TEXT(Setup!C28,"dd/mm/yyyy")&"   (modifica in Setup)"')
header(ws, 4, ["KPI", "Totale notte", "=Setup!C7", "=Setup!C8", '="Δ "&Setup!C7&" vs "&Setup!C8',
               "Target", "Totale vs target"], col=2)
header(ws, 4, ["Prec. Totale", '="Prec. "&Setup!C7', '="Prec. "&Setup!C8', "Δ Totale vs prec.",
               '="Δ "&Setup!C7&" vs prec."', '="Δ "&Setup!C8&" vs prec."', "Andamento (totale)"], col=9)
# somme di base (riga 30+)
SUM_ROW = 31
ws.cell(SUM_ROW - 1, 2, "Somme di base (non modificare)").font = SEC_FONT
header(ws, SUM_ROW, ["Misura", "Totale notte", "=Setup!C7", "=Setup!C8", "Prec. Totale",
                     '="Prec. "&Setup!C7', '="Prec. "&Setup!C8'], col=2)
PA = 'Turni_Log!$A$2:$A,">="&Setup!$C$27,Turni_Log!$A$2:$A,"<="&Setup!$C$28'
sum_ref = {}
for i, (lab, m, col) in enumerate(MEASURES):
    r = SUM_ROW + 1 + i
    ws.cell(r, 2, lab)
    rng = f"Turni_Log!${col}$2:${col}"
    ws.cell(r, 3, f"=SUMIFS({rng},Turni_Log!$Y$2:$Y,1)")
    ws.cell(r, 4, f"=SUMIFS({rng},Turni_Log!$Y$2:$Y,1,Turni_Log!$O$2:$O,{C1})")
    ws.cell(r, 5, f"=SUMIFS({rng},Turni_Log!$Y$2:$Y,1,Turni_Log!$O$2:$O,{C2})")
    ws.cell(r, 6, f"=SUMIFS({rng},{PA})")  # le misure sono già 0 fuori dai turni notturni
    ws.cell(r, 7, f"=SUMIFS({rng},{PA},Turni_Log!$O$2:$O,{C1})")
    ws.cell(r, 8, f"=SUMIFS({rng},{PA},Turni_Log!$O$2:$O,{C2})")
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
    for col, scol in zip("IJK", "FGH"):
        ws[f"{col}{r}"] = f'=IFERROR({mul}{scol}{sum_ref[num]}/{scol}{sum_ref[den]},"n.d.")'
    for col, a, b in zip("LMN", "CDE", "IJK"):
        ws[f"{col}{r}"] = f'=IF(AND(ISNUMBER({a}{r}),ISNUMBER({b}{r})),{a}{r}-{b}{r},"n.d.")'
    ws[f"O{r}"] = (f'=IF(ISNUMBER(L{r}),IF(L{r}=0,"Stabile",IF((Setup!D{20 + i}="Alto")=(L{r}>0),'
                   f'"Migliora","Peggiora")),"n.d.")')
    fmt = NUM1 if k == "NPS" else PCT
    for col in "CDEFGIJKLMN":
        ws[f"{col}{r}"].number_format = fmt
ws.conditional_formatting.add("H5:H10", FormulaRule(formula=['H5="Fuori target"'], fill=RED))
ws.conditional_formatting.add("H5:H10", FormulaRule(formula=['H5="In target"'], fill=GREEN))
ws.conditional_formatting.add("O5:O10", FormulaRule(formula=['O5="Peggiora"'], fill=RED))
ws.conditional_formatting.add("O5:O10", FormulaRule(formula=['O5="Migliora"'], fill=GREEN))
header(ws, 12, ["Prec. Totale", '="Prec. "&Setup!C7', '="Prec. "&Setup!C8'], col=9)
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
    for col, scol in zip("IJK", "FGH"):
        ws[f"{col}{r}"] = f"={scol}{sum_ref[src_lab]}"
    r += 1
ws.cell(r, 2, "Giornate-agente (turni notturni)")
ws[f"C{r}"] = "=COUNTIF(Turni_Log!$Y$2:$Y,1)"
ws[f"D{r}"] = f"=COUNTIFS(Turni_Log!$Y$2:$Y,1,Turni_Log!$O$2:$O,{C1})"
ws[f"E{r}"] = f"=COUNTIFS(Turni_Log!$Y$2:$Y,1,Turni_Log!$O$2:$O,{C2})"
ws[f"I{r}"] = f"=COUNTIFS({PA},Turni_Log!$O$2:$O,{C1})+COUNTIFS({PA},Turni_Log!$O$2:$O,{C2})"
ws[f"J{r}"] = f"=COUNTIFS({PA},Turni_Log!$O$2:$O,{C1})"
ws[f"K{r}"] = f"=COUNTIFS({PA},Turni_Log!$O$2:$O,{C2})"
r += 1
ws.cell(r, 2, "Agenti distinti")
ws[f"C{r}"] = '=IFERROR(ROWS(UNIQUE(FILTER(Turni_Log!$B$2:$B,Turni_Log!$Y$2:$Y=1))),0)'
ws[f"D{r}"] = f'=IFERROR(ROWS(UNIQUE(FILTER(Turni_Log!$B$2:$B,Turni_Log!$Y$2:$Y=1,Turni_Log!$O$2:$O={C1}))),0)'
ws[f"E{r}"] = f'=IFERROR(ROWS(UNIQUE(FILTER(Turni_Log!$B$2:$B,Turni_Log!$Y$2:$Y=1,Turni_Log!$O$2:$O={C2}))),0)'
PF = "Turni_Log!$A$2:$A>=Setup!$C$27,Turni_Log!$A$2:$A<=Setup!$C$28"
ws[f"I{r}"] = (f'=IFERROR(ROWS(UNIQUE(FILTER(Turni_Log!$B$2:$B,{PF},'
               f'((Turni_Log!$O$2:$O={C1})+(Turni_Log!$O$2:$O={C2}))>0))),0)')
ws[f"J{r}"] = f'=IFERROR(ROWS(UNIQUE(FILTER(Turni_Log!$B$2:$B,{PF},Turni_Log!$O$2:$O={C1}))),0)'
ws[f"K{r}"] = f'=IFERROR(ROWS(UNIQUE(FILTER(Turni_Log!$B$2:$B,{PF},Turni_Log!$O$2:$O={C2}))),0)'
r += 1
ws.cell(r, 2, "Assegnazioni per giornata-agente")
ws[f"C{r}"] = f'=IFERROR(C13/C16,"n.d.")'
ws[f"D{r}"] = f'=IFERROR(D13/D16,"n.d.")'
ws[f"E{r}"] = f'=IFERROR(E13/E16,"n.d.")'
for col in "IJK":
    ws[f"{col}{r}"] = f'=IFERROR({col}13/{col}16,"n.d.")'
for col in "CDEIJK":
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
coach_agents = [a for a in test_agents if agent_agg.get(a, {}).get("Assegnazioni completate", 0) >= 20][:2]
for i, (a, k) in enumerate(zip(coach_agents, ["Handoff", "Pending"])):
    ws.append([PERIOD_START - dt.timedelta(days=3 - i), a, SHIFTS[i][1], k, f"TEST-{2000 + i}", "Migliorabile",
               "Handoff evitabile" if k == "Handoff" else "Pend non necessario", "Coaching 1:1", "Sì",
               PERIOD_START, "TL di prova", "TEST - riga di esempio per il foglio Coaching: cancellala"])
    for col in (1, 10):
        ws.cell(ws.max_row, col).number_format = DATE
widths(ws, {"A": 12, "B": 24, "C": 8, "D": 10, "E": 34, "F": 13, "G": 30, "H": 18, "I": 10,
            "J": 12, "K": 18, "L": 40})
ws.freeze_panes = "A2"

# ---- Riferimenti comuni per i fogli di analisi
TL_ = "Turni_Log!"
LOGB, LOGY, LOGO = "Turni_Log!$B$2:$B", "Turni_Log!$Y$2:$Y", "Turni_Log!$O$2:$O"
MCOL = {lab: col for lab, _, col in MEASURES}
MIN_OF = {"Throughput": "0", "Recontact": "0", "Handoff": "Setup!$C$42", "Pending": "Setup!$C$42",
          "Reopen": "Setup!$C$43", "NPS": "Setup!$C$44"}
KPI_NAMES = [k for k, *_ in TARGETS]
TGT_ROW = {k: 20 + i for i, k in enumerate(KPI_NAMES)}
L_ = openpyxl.utils.get_column_letter


def kpi_expr(k, sumf, na='"n.s."'):
    """Espressione KPI (con soglia di volume) data una funzione colonna -> SUMIFS."""
    num, den, mul = KPI_DEF[k]
    return (f'LET(nx,{sumf(MCOL[num])},dx,{sumf(MCOL[den])},'
            f'IF(AND(dx>0,dx>={MIN_OF[k]}),{mul}nx/dx,{na}))')


def fmt_col(ws, col, k, r1, r2):
    for r in range(r1, r2 + 1):
        ws[f"{col}{r}"].number_format = NUM1 if k == "NPS" else PCT


def title(ws, text, note=None):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    if note:
        ws["A2"] = note
        ws["A2"].font = Font(italic=True, color="7F7F7F")


AG_END = AGENT_ROWS + 1  # ultima riga di Agenti

# ---- Rotazione (4): stesso agente, giorni 18-3 vs giorni 21-6
ws = ws_rot
title(ws, "ROTAZIONE - stesso agente sul 18-3 e sul 21-6",
      "Solo agenti che nel periodo hanno lavorato su entrambi i turni. Δ = " + SHIFTS[0][1] + " meno "
      + SHIFTS[1][1] + ". n.s. = volume sotto le soglie di Setup (C42:C44). "
      "La riga 'Totale agenti misti' somma numeratori e denominatori di tutti questi agenti.")
rot_hdr = ["LDAP", "Team Lead", '="Giorni "&Setup!C7', '="Giorni "&Setup!C8',
           '="Assegn./giornata "&Setup!C7', '="Assegn./giornata "&Setup!C8']
for k in KPI_NAMES:
    rot_hdr += [f'="{k} "&Setup!C7', f'="{k} "&Setup!C8', f"Δ {k}"]
header(ws, 4, rot_hdr)
R1, R2 = 6, 6 + AGENT_ROWS - 1
RA = f"A{R1}:A{R2}"
ws["A5"] = "Totale agenti misti"
ws["A5"].font = Font(bold=True)
ws[f"A{R1}"] = (f'=IFERROR(FILTER(Agenti!A2:A{AG_END},Agenti!E2:E{AG_END}="Misto"),'
                '"Nessun agente su entrambi i turni nel periodo")')
ws[f"B{R1}"] = f'=MAP({RA},LAMBDA(a,IF(a="","",IFERROR(VLOOKUP(a,Agenti!$A$2:$B${AG_END},2,0),""))))'
sa_ = lambda col, code: f"SUMIFS(Turni_Log!${col}$2:${col},{LOGB},a,{LOGY},1,{LOGO},{code})"
tot_ = lambda col, code: f"SUM(MAP({RA},LAMBDA(a,IF(a=\"\",0,{sa_(col, code)}))))"
for col, code in (("C", C1), ("D", C2)):
    ws[f"{col}{R1}"] = f'=MAP({RA},LAMBDA(a,IF(a="","",COUNTIFS({LOGB},a,{LOGY},1,{LOGO},{code}))))'
    ws[f"{col}5"] = f"=SUM({col}{R1}:{col}{R2})"
for col, dcol, code in (("E", "C", C1), ("F", "D", C2)):
    ws[f"{col}{R1}"] = (f'=MAP({RA},{dcol}{R1}:{dcol}{R2},LAMBDA(a,n,IF(a="","",IF(n>0,'
                        f'{sa_("P", code)}/n,"-"))))')
    ws[f"{col}5"] = f'=IFERROR({tot_("P", code)}/{dcol}5,"n.d.")'
    for r in range(5, R2 + 1):
        ws[f"{col}{r}"].number_format = NUM1
c = 7
for k in KPI_NAMES:
    num, den, mul = KPI_DEF[k]
    for code in (C1, C2):
        L = L_(c)
        ws[f"{L}{R1}"] = (f'=MAP({RA},LAMBDA(a,IF(a="","",'
                          f'{kpi_expr(k, lambda cc, code=code: sa_(cc, code))})))')
        ws[f"{L}5"] = f'=IFERROR({mul}{tot_(MCOL[num], code)}/{tot_(MCOL[den], code)},"n.d.")'
        fmt_col(ws, L, k, 5, R2)
        c += 1
    L, La, Lb = L_(c), L_(c - 2), L_(c - 1)
    ws[f"{L}5"] = f'=IF(AND(ISNUMBER({La}5),ISNUMBER({Lb}5)),{La}5-{Lb}5,"")'
    ws[f"{L}{R1}"] = (f'=MAP({La}{R1}:{La}{R2},{Lb}{R1}:{Lb}{R2},LAMBDA(x,y,'
                      f'IF(AND(ISNUMBER(x),ISNUMBER(y)),x-y,"")))')
    fmt_col(ws, L, k, 5, R2)
    ws.column_dimensions[L].width = 9
    c += 1
for col in range(1, c):
    ws.cell(5, col).fill = CALC_FILL
    ws.cell(5, col).font = Font(bold=True)
widths(ws, {"A": 26, "B": 20, "C": 9, "D": 9, "E": 11, "F": 11})
ws.freeze_panes = "B6"

# ---- Carico (6): assegnazioni per giornata-agente
ws = ws_car
title(ws, "CARICO - assegnazioni per giornata-agente",
      "Giornata-agente = giorno Tableau attribuito a un turno notturno nel periodo. Giorno / data = giorno di INIZIO "
      "del turno (es. il 21-6 che inizia lunedì è 'Lun'). Settimane = settimana (dom) del giorno Tableau.")
blk = ["Giornate", "Assegn.", "Assegn./giornata"]
car_hdr = [f'="{b} "&Setup!C{7 + j}' for j in range(2) for b in blk] + [f"{b} totale" for b in blk]


# a) per giorno della settimana
ws["A4"] = "Per giorno della settimana"
ws["A4"].font = SEC_FONT
header(ws, 5, ["Giorno"] + car_hdr)
for i, wd in enumerate(WEEKDAYS + ["Totale"]):
    r = 6 + i
    ws[f"A{r}"] = wd
    extra = f",Turni_Log!$AB$2:$AB,{i + 1}" if wd != "Totale" else ""
    for j, code in enumerate((C1, C2, None)):
        cg, ca, cr = (L_(2 + 3 * j + x) for x in range(3))
        if code:
            ws[f"{cg}{r}"] = f"=COUNTIFS({LOGY},1,{LOGO},{code}{extra})"
            ws[f"{ca}{r}"] = f"=SUMIFS(Turni_Log!$P$2:$P,{LOGY},1,{LOGO},{code}{extra})"
        else:
            ws[f"{cg}{r}"] = f"=B{r}+E{r}"
            ws[f"{ca}{r}"] = f"=C{r}+F{r}"
        ws[f"{cr}{r}"] = f'=IFERROR({ca}{r}/{cg}{r},"-")'
        ws[f"{cr}{r}"].number_format = NUM1
    if wd == "Totale":
        for col in range(1, 11):
            ws.cell(r, col).font = Font(bold=True)
            ws.cell(r, col).fill = CALC_FILL


def load_map(ws, r1, r2, keycol, crit):
    """Colonne per chiave (settimana o data) calcolate con MAP; crit = criterio SUMIFS sulla chiave k."""
    K = f"{keycol}{r1}:{keycol}{r2}"
    for j, code in enumerate((C1, C2, None)):
        cg, ca, cr = (L_(ord(keycol) - 64 + 1 + 3 * j + x) for x in range(3))
        if code:
            ws[f"{cg}{r1}"] = f'=MAP({K},LAMBDA(k,IF(k="","",COUNTIFS({LOGY},1,{LOGO},{code},{crit}))))'
            ws[f"{ca}{r1}"] = (f'=MAP({K},LAMBDA(k,IF(k="","",'
                               f'SUMIFS(Turni_Log!$P$2:$P,{LOGY},1,{LOGO},{code},{crit}))))')
        else:
            g1, a1 = L_(ord(keycol) - 64 + 1), L_(ord(keycol) - 64 + 2)
            g2, a2 = L_(ord(keycol) - 64 + 4), L_(ord(keycol) - 64 + 5)
            ws[f"{cg}{r1}"] = f'=ARRAYFORMULA(IF({K}="","",{g1}{r1}:{g1}{r2}+{g2}{r1}:{g2}{r2}))'
            ws[f"{ca}{r1}"] = f'=ARRAYFORMULA(IF({K}="","",{a1}{r1}:{a1}{r2}+{a2}{r1}:{a2}{r2}))'
        ws[f"{cr}{r1}"] = (f'=MAP({K},{cg}{r1}:{cg}{r2},{ca}{r1}:{ca}{r2},LAMBDA(k,g,x,'
                           f'IF(k="","",IF(g>0,x/g,"-"))))')
        for r in range(r1, r2 + 1):
            ws[f"{cr}{r}"].number_format = NUM1


# b) per settimana
ws["A16"] = "Per settimana"
ws["A16"].font = SEC_FONT
header(ws, 17, ["Settimana (dom)"] + car_hdr)
WS_ = lambda x: f"({x}-WEEKDAY({x})+1)"
ws["A18"] = (f"=SEQUENCE(INT(({WS_('Setup!$C$4')}-{WS_('Setup!$C$3')})/7)+1,1,{WS_('Setup!$C$3')},7)")
load_map(ws, 18, 37, "A", "Turni_Log!$I$2:$I,k")
for r in range(18, 38):
    ws[f"A{r}"].number_format = DATE
# c) per giorno (data di inizio turno)
ws["A40"] = "Per giorno (data di inizio turno)"
ws["A40"].font = SEC_FONT
header(ws, 41, ["Data inizio turno"] + car_hdr + ["Giorno"])
ws["A42"] = f'=IFERROR(SORT(UNIQUE(FILTER(Turni_Log!$AA$2:$AA,{LOGY}=1))),"")'
load_map(ws, 42, 141, "A", "Turni_Log!$AA$2:$AA,k")
ws["K42"] = ('=MAP(A42:A141,LAMBDA(k,IF(k="","",CHOOSE(WEEKDAY(k),'
             + ",".join(f'"{w}"' for w in WEEKDAYS) + "))))")
for r in range(42, 142):
    ws[f"A{r}"].number_format = DATE
# d) per agente
ws["M4"] = "Per agente (assegnazioni per giornata)"
ws["M4"].font = SEC_FONT
header(ws, 5, ["LDAP", "Turno prevalente", "Giornate", "Assegn./giornata", '="Assegn./giornata "&Setup!C7',
               '="Assegn./giornata "&Setup!C8'] + [f"{w} (inizio turno)" for w in WEEKDAYS], col=13)
CA1, CA2 = 6, 6 + AGENT_ROWS - 1
MA = f"M{CA1}:M{CA2}"
ws[f"M{CA1}"] = f"=ARRAYFORMULA(Agenti!A2:A{AG_END})"
ws[f"N{CA1}"] = f"=ARRAYFORMULA(Agenti!E2:E{AG_END})"
ws[f"O{CA1}"] = f'=ARRAYFORMULA(IF({MA}="","",Agenti!C2:C{AG_END}+Agenti!D2:D{AG_END}))'
ws[f"P{CA1}"] = (f'=MAP({MA},O{CA1}:O{CA2},LAMBDA(a,n,IF(a="","",IF(n>0,'
                 f'SUMIFS(Turni_Log!$P$2:$P,{LOGB},a,{LOGY},1)/n,"-"))))')
for col, code in (("Q", C1), ("R", C2)):
    ws[f"{col}{CA1}"] = (f'=MAP({MA},LAMBDA(a,IF(a="","",IFERROR(SUMIFS(Turni_Log!$P$2:$P,{LOGB},a,{LOGY},1,'
                         f'{LOGO},{code})/COUNTIFS({LOGB},a,{LOGY},1,{LOGO},{code}),"-"))))')
for i in range(7):
    col = L_(19 + i)
    crit = f"{LOGB},a,{LOGY},1,Turni_Log!$AB$2:$AB,{i + 1}"
    ws[f"{col}{CA1}"] = (f'=MAP({MA},LAMBDA(a,IF(a="","",IFERROR(SUMIFS(Turni_Log!$P$2:$P,{crit})'
                         f'/COUNTIFS({crit}),"-"))))')
for c in range(16, 26):
    for r in range(CA1, CA2 + 1):
        ws.cell(r, c).number_format = NUM1
    ws.column_dimensions[L_(c)].width = 10
widths(ws, {"A": 16, "K": 7, "L": 3, "M": 24, "N": 11, "O": 9})
for c in "BCDEFGHIJ":
    ws.column_dimensions[c].width = 11

# ---- Rossi_Consecutivi (10)
ws = ws_str
title(ws, "KPI IN ROSSO PER SETTIMANE CONSECUTIVE (ultime 8 settimane)",
      "Stessa regola del foglio Spotcheck applicata settimana per settimana (media e dev. std della notte "
      "di quella settimana). Il conteggio parte dall'ultima settimana e si ferma al primo non-ROSSO.")
header(ws, 4, ["LDAP", "Team Lead"] + [f"Settimane ROSSO {k}" for k in FOCUS] + ["Max"])
S1, S2 = 5, 5 + AGENT_ROWS - 1
SA = f"$A${S1}:$A${S2}"
ws[f"A{S1}"] = f"=ARRAYFORMULA(Agenti!A2:A{AG_END})"
ws[f"B{S1}"] = f"=ARRAYFORMULA(Agenti!B2:B{AG_END})"
c = 10
names = ["ra", "rb", "rc", "rd", "re", "rf", "rg", "rh"]
for q, k in enumerate(FOCUS):
    num, den, mul = KPI_DEF[k]
    idx = KPI_NAMES.index(k)
    trend_col = L_(3 + 3 * idx)  # colonna "k Tot" del trend in Dashboard
    ws.cell(3, c, f"{k} - valore settimanale").font = Font(bold=True)
    ws.cell(3, c + 8, f"{k} - flag settimanale").font = Font(bold=True)
    flag_cols = []
    for w in range(8):
        vc, fc = L_(c + w), L_(c + 8 + w)
        for col in (vc, fc):
            ws[f"{col}4"] = f"=Dashboard!$B${W1 + w}"
            ws[f"{col}4"].number_format = "dd/mm"
            ws[f"{col}4"].fill, ws[f"{col}4"].font = HDR_FILL, HDR_FONT
            ws.column_dimensions[col].width = 8
        sw = lambda cc, vc=vc: f"SUMIFS(Turni_Log!${cc}$2:${cc},{LOGB},a,Turni_Log!$I$2:$I,{vc}$4)"
        ws[f"{vc}{S1}"] = f'=MAP({SA},LAMBDA(a,IF(a="","",{kpi_expr(k, sw)})))'
        fmt_col(ws, vc, k, S1, S2)
        v = f"{vc}{S1}:{vc}{S2}"
        mean = f"Dashboard!${trend_col}${W1 + w}"
        tgt = f"Setup!$C${TGT_ROW[k]}"
        sd = f"IFERROR(STDEV(FILTER({v},ISNUMBER({v}))),0)"
        if k == "NPS":
            red, yel = f"({v}<{tgt})*({v}<{mean}-Setup!$C$45*{sd})", f"{v}<{tgt}"
        else:
            red, yel = f"({v}>{tgt})*({v}>{mean}+Setup!$C$45*{sd})", f"{v}>{tgt}"
        ws[f"{fc}{S1}"] = (f'=ARRAYFORMULA(IF({SA}="","",IF(ISNUMBER({v})*ISNUMBER({mean}),'
                           f'IF({red},"ROSSO",IF({yel},"GIALLO","VERDE")),"n.s.")))')
        flag_cols.append(f"{fc}{S1}:{fc}{S2}")
    flag_cf(ws, f"{L_(c + 8)}{S1}:{L_(c + 15)}{S2}", f"{L_(c + 8)}{S1}")
    # striscia: dall'ultima settimana (rh) all'indietro
    expr = "8"
    for j in range(0, 8):
        expr = f'IF({names[j]}<>"ROSSO",{7 - j},{expr})'
    ws[f"{L_(3 + q)}{S1}"] = (f'=MAP({SA},{",".join(flag_cols)},LAMBDA(a,{",".join(names)},'
                              f'IF(a="","",{expr})))')
    c += 17
ws[f"G{S1}"] = f'=MAP({SA},C{S1}:C{S2},D{S1}:D{S2},E{S1}:E{S2},F{S1}:F{S2},LAMBDA(a,h,p,n,o,IF(a="","",MAX(h,p,n,o))))'
ws.conditional_formatting.add(f"C{S1}:G{S2}", FormulaRule(formula=[f"AND(ISNUMBER(C{S1}),C{S1}>=2)"], fill=RED))
ws.conditional_formatting.add(f"C{S1}:G{S2}", FormulaRule(formula=[f"AND(ISNUMBER(C{S1}),C{S1}=1)"], fill=YELLOW))
widths(ws, {"A": 24, "B": 20, "C": 11, "D": 11, "E": 11, "F": 11, "G": 7, "H": 3, "I": 3})
ws.freeze_panes = "C5"

# ---- Coaching (9)
ws = ws_coa
title(ws, "EFFETTO COACHING - KPI 2 settimane prima vs 2 settimane dopo",
      "Righe di Spotcheck_Log con 'Coaching fatto?' = Sì e Data spotcheck compilata. Prima = 14 giorni prima della "
      "Data spotcheck; Dopo = 14 giorni dal giorno successivo. Solo giornate attribuite ai turni notturni. "
      "'(parziale)' = i 14 giorni dopo non sono ancora tutti nell'estrazione Tableau.")
header(ws, 4, ["LDAP", "KPI", "Data coaching", "KPI prima", "Volume prima", "KPI dopo", "Volume dopo",
               "Δ (dopo - prima)", "Esito"])
K1, K2 = 5, 204
KA = f"A{K1}:A{K2}"
SL = "Spotcheck_Log!"
ws[f"A{K1}"] = (f'=IFERROR(SORT(UNIQUE(FILTER(HSTACK({SL}$B$2:$B,{SL}$D$2:$D,{SL}$J$2:$J),'
                f'{SL}$I$2:$I="Sì",{SL}$J$2:$J<>"",{SL}$B$2:$B<>"")),3,FALSE,1,TRUE),'
                '"Nessun coaching registrato in Spotcheck_Log")')


def win(s_, e_, what):
    sw = lambda cc: f'SUMIFS(Turni_Log!${cc}$2:${cc},{LOGB},a,Turni_Log!$A$2:$A,">="&({s_}),Turni_Log!$A$2:$A,"<="&({e_}))'
    num = "SWITCH(kp," + ",".join(f'"{k}",{KPI_DEF[k][2]}{sw(MCOL[KPI_DEF[k][0]])}' for k in KPI_NAMES) + ",0)"
    den = "SWITCH(kp," + ",".join(f'"{k}",{sw(MCOL[KPI_DEF[k][1]])}' for k in KPI_NAMES) + ",0)"
    body = f'LET(nx,{num},dx,{den},IF(dx>0,nx/dx,"n.d."))' if what == "kpi" else den
    return f'=MAP({KA},B{K1}:B{K2},C{K1}:C{K2},LAMBDA(a,kp,dc,IF(OR(a="",NOT(ISNUMBER(dc))),"",{body})))'


ws[f"D{K1}"] = win("dc-14", "dc-1", "kpi")
ws[f"E{K1}"] = win("dc-14", "dc-1", "vol")
ws[f"F{K1}"] = win("dc+1", "dc+14", "kpi")
ws[f"G{K1}"] = win("dc+1", "dc+14", "vol")
ws[f"H{K1}"] = f'=MAP(D{K1}:D{K2},F{K1}:F{K2},LAMBDA(x,y,IF(AND(ISNUMBER(x),ISNUMBER(y)),y-x,"")))'
ws[f"I{K1}"] = (f'=LET(mx,MAX(Tableau!$I$2:$I),MAP({KA},B{K1}:B{K2},C{K1}:C{K2},H{K1}:H{K2},'
                'LAMBDA(a,kp,dc,dl,IF(OR(a="",NOT(ISNUMBER(dc))),"",IF(mx<dc+14,"(parziale) ","")&'
                'IF(dl="","n.d.",IF(dl=0,"Invariato",IF(OR(kp="NPS",kp="Throughput")=(dl>0),'
                '"Migliorato","Peggiorato")))))))')
ws["K4"] = "Riepilogo"
ws["K4"].font = SEC_FONT
for i, (lab, pat) in enumerate((("Migliorati", "*Migliorato"), ("Peggiorati", "*Peggiorato"),
                                ("Invariati", "*Invariato"), ("n.d.", "*n.d."))):
    ws[f"K{5 + i}"] = lab
    ws[f"L{5 + i}"] = f'=COUNTIF(I{K1}:I{K2},"{pat}")'
for r in range(K1, K2 + 1):
    ws[f"C{r}"].number_format = DATE
    for col in "DFH":
        ws[f"{col}{r}"].number_format = "0.00"
ws.conditional_formatting.add(f"I{K1}:I{K2}", FormulaRule(formula=[f'ISNUMBER(SEARCH("Migliorato",I{K1}))'], fill=GREEN))
ws.conditional_formatting.add(f"I{K1}:I{K2}", FormulaRule(formula=[f'ISNUMBER(SEARCH("Peggiorato",I{K1}))'], fill=RED))
ws["A3"] = "KPI: Handoff/Pending/Reopen in rapporto (0.08 = 8%), NPS in punti. Per NPS e Throughput Δ > 0 = meglio."
ws["A3"].font = Font(italic=True, color="7F7F7F")
widths(ws, {"A": 24, "B": 11, "C": 12, "D": 10, "E": 10, "F": 10, "G": 10, "H": 12, "I": 22, "J": 3, "K": 12})
ws.freeze_panes = "A5"

# ---- Medallia (incolla) + mappatura in Setup (15)
ws = ws_med
med_hdr = [h for _, h in MEDALLIA_COLS]
header(ws, 1, med_hdr)
ws["G1"] = ("Incolla qui l'export Medallia da A1 (intestazioni in riga 1). Le colonne usate si impostano in "
            "Setup C64:C68. Le righe attuali sono DATI DI TEST da cancellare.")
ws["G1"].font = Font(italic=True, color="C00000")
ws.column_dimensions["B"].number_format = DATE
med_agents = [a for a in test_agents if a in agent_agg][:8]
for i in range(16):
    a = med_agents[i % len(med_agents)]
    d = PERIOD_START + dt.timedelta(days=i % 5) if i < 12 else PREV_START + dt.timedelta(days=i % 5)
    ws.append([a, d, [2, 9, 5, 10, 0, 6, 8, 3, 7, 10, 4, 1, 2, 9, 6, 0][i], f"TEST - commento di prova {i + 1}",
               f"TEST-{1000 + i}"])
    ws.cell(ws.max_row, 2).number_format = DATE
widths(ws, {"A": 24, "B": 14, "C": 12, "D": 50, "E": 14})
ws.freeze_panes = "A2"

ws = ws_set
header(ws, 63, ["Medallia: campo", "Intestazione colonna nel foglio Medallia"], col=2)
for i, (lab, h) in enumerate(MEDALLIA_COLS):
    ws.cell(64 + i, 2, lab)
    ws.cell(64 + i, 3, h).fill = INPUT_FILL
ws.cell(69, 2, "Detractor = punteggio minore o uguale a")
ws.cell(69, 3, DETRACTOR_MAX).fill = INPUT_FILL
ws["D64"] = ("DA VERIFICARE: scrivi i nomi esatti delle colonne dell'export Medallia. "
             "LDAP anche come email (la parte dopo @ viene ignorata).")
ws["D64"].font = Font(bold=True, color="C00000")
MED_MATCH = [f"MATCH(Setup!$C${64 + i},Medallia!$A$1:$AZ$1,0)" for i in range(5)]

# ---- Detractor_NPS (15)
ws = ws_det
title(ws, "DETRACTOR NPS DEGLI AGENTI NOTTURNI (da Medallia)",
      "Risposte con punteggio ≤ soglia (Setup C69), data nel periodo, agenti presenti in Turni_Log nel periodo. "
      "Usale per scegliere i casi NPS da controllare; 'Già in Spotcheck_Log' cerca il Case ID nella colonna E.")
header(ws, 4, ["Data", "LDAP", "Punteggio", "Case ID", "Verbatim", "Team Lead", "Turno attribuito",
               "Già in Spotcheck_Log?"])
D1, D2 = 5, 504
ws[f"A{D1}"] = (
    '=IFERROR(LET(m,Medallia!$A$2:$AZ$5000,'
    f'l,MAP(INDEX(m,0,{MED_MATCH[0]}),LAMBDA(x,IF(x="","",TRIM(REGEXREPLACE(x&"","@.*",""))))),'
    f'dn,MAP(INDEX(m,0,{MED_MATCH[1]}),LAMBDA(x,IF(ISNUMBER(x),INT(x),IFERROR(INT(DATEVALUE(x)),"")))),'
    f'sn,MAP(INDEX(m,0,{MED_MATCH[2]}),LAMBDA(x,IF(x="","",IFERROR(VALUE(x),"")))),'
    f'vb,INDEX(m,0,{MED_MATCH[3]}),id,INDEX(m,0,{MED_MATCH[4]}),'
    'ok,MAP(l,dn,sn,LAMBDA(a,b,c,IF(AND(a<>"",ISNUMBER(b),ISNUMBER(c)),'
    f'AND(b>=Setup!$C$3,b<=Setup!$C$4,c<=Setup!$C$69,COUNTIF(Agenti!$A$2:$A${AG_END},a)>0),FALSE))),'
    'SORT(FILTER(HSTACK(dn,l,sn,id,vb),ok),3,TRUE,1,TRUE)),'
    '"Nessun detractor nel periodo (se Medallia è pieno, verifica le colonne in Setup C64:C68)")')
DB = f"B{D1}:B{D2}"
ws[f"F{D1}"] = f'=MAP({DB},LAMBDA(l,IF(l="","",IFERROR(VLOOKUP(l,Agenti!$A$2:$B${AG_END},2,0),""))))'
ws[f"G{D1}"] = (f'=MAP(A{D1}:A{D2},{DB},LAMBDA(d,l,IF(l="","",'
                f'IFERROR(VLOOKUP(l&"|"&d,Turni_Log!$H$2:$O,8,0),"n.d."))))')
ws[f"H{D1}"] = (f'=MAP(D{D1}:D{D2},LAMBDA(i,IF(i="","",'
                f'IF(COUNTIF(Spotcheck_Log!$E$2:$E,"*"&i&"*")>0,"Sì","No"))))')
for r in range(D1, D2 + 1):
    ws[f"A{r}"].number_format = DATE
    ws[f"E{r}"].alignment = Alignment(wrap_text=True, vertical="top")
widths(ws, {"A": 12, "B": 24, "C": 10, "D": 16, "E": 70, "F": 20, "G": 12, "H": 12})
ws.freeze_panes = "A5"

# ---- Completamento spotcheck per TL (11) - nel foglio Spotcheck
ws = ws_spot
TLR1, TLR2 = 18, 47
header(ws, 17, ["Completamento per TL", "Agenti da controllare", "Casi da campionare", "Casi registrati",
                "% completamento"], col=19)
SPB, SPO, SPP = f"$B$5:$B${E}", f"$O$5:$O${E}", f"$P$5:$P${E}"
ws[f"S{TLR1}"] = f'=IFERROR(SORT(UNIQUE(FILTER(Agenti!B2:B{AG_END},Agenti!A2:A{AG_END}<>"",Agenti!B2:B{AG_END}<>""))),"")'
TLS = f"S{TLR1}:S{TLR2}"
ws[f"T{TLR1}"] = f'=MAP({TLS},LAMBDA(t,IF(t="","",COUNTIFS({SPB},t,{SPO},">0"))))'
ws[f"U{TLR1}"] = f'=MAP({TLS},LAMBDA(t,IF(t="","",SUMIFS({SPO},{SPB},t))))'
ws[f"V{TLR1}"] = (f'=MAP({TLS},LAMBDA(t,IF(t="","",SUM(MAP({SPB},{SPO},{SPP},'
                  f'LAMBDA(b,o,p,IF(AND(b=t,ISNUMBER(o)),MIN(o,p),0)))))))')
ws[f"W{TLR1}"] = f'=MAP({TLS},U{TLR1}:U{TLR2},V{TLR1}:V{TLR2},LAMBDA(t,u,v,IF(t="","",IF(u>0,v/u,"-"))))'
for r in range(TLR1, TLR2 + 1):
    ws[f"W{r}"].number_format = "0%"
ws.conditional_formatting.add(f"W{TLR1}:W{TLR2}", FormulaRule(formula=[f"AND(ISNUMBER(W{TLR1}),W{TLR1}<1)"], fill=RED))
ws.conditional_formatting.add(f"W{TLR1}:W{TLR2}", FormulaRule(formula=[f"AND(ISNUMBER(W{TLR1}),W{TLR1}>=1)"], fill=GREEN))
ws[f"S{TLR2 + 1}"] = "Casi registrati = casi in Spotcheck_Log nel periodo, al massimo quelli dovuti per agente."
ws[f"S{TLR2 + 1}"].font = Font(italic=True, color="7F7F7F")
widths(ws, {"V": 11, "W": 13})

# ---- Vista_TL (8)
ws = ws_vtl
title(ws, "VISTA TEAM LEAD")
ws["A2"] = "Team Lead (vuoto = tutti):"
ws["A2"].font = Font(bold=True)
ws["B2"].fill = INPUT_FILL
ws["AB1"] = "Elenco TL"
ws["AB2"] = f'=IFERROR(SORT(UNIQUE(FILTER(Agenti!B2:B{AG_END},Agenti!A2:A{AG_END}<>"",Agenti!B2:B{AG_END}<>""))),"")'
v = DataValidation(type="list", formula1="=$AB$2:$AB$60", allow_blank=True)
v.add("B2")
ws.add_data_validation(v)
sel = lambda rng: f'((({rng}=$B$2)+($B$2=""))*({rng}<>""))>0'
ws["A3"] = (f'="Agenti: "&COUNTIFS(Spotcheck!{SPB},IF($B$2="","?*",$B$2))&"   |   Casi da campionare: "&'
            f'IF($B$2="",SUM(Spotcheck!{SPO}),SUMIFS(Spotcheck!{SPO},Spotcheck!{SPB},$B$2))&'
            f'"   |   Registrati: "&SUM(MAP(Spotcheck!{SPB},Spotcheck!{SPO},Spotcheck!{SPP},LAMBDA(b,o,p,'
            f'IF(AND(ISNUMBER(o),IF($B$2="",b<>"",b=$B$2)),MIN(o,p),0))))')
ws["A3"].font = Font(bold=True, color="1F3864")
header(ws, 5, sp_hdr)
V1, V2 = 6, 6 + AGENT_ROWS - 1
ws[f"A{V1}"] = (f'=IFERROR(FILTER(Spotcheck!A5:Q{E},Spotcheck!A5:A{E}<>"",{sel(f"Spotcheck!B5:B{E}")}),'
                '"Nessun agente per questo TL")')
for c_ in "EGL":
    ws.column_dimensions[c_].number_format = PCT
ws.column_dimensions["I"].number_format = NUM1
for c_ in "FHKM":
    flag_cf(ws, f"{c_}{V1}:{c_}{V2}", f"{c_}{V1}")
ws.conditional_formatting.add(f"Q{V1}:Q{V2}", FormulaRule(formula=[f'LEFT(Q{V1},7)="Da fare"'], fill=RED))
ws.conditional_formatting.add(f"Q{V1}:Q{V2}", FormulaRule(formula=[f'Q{V1}="Completato"'], fill=GREEN))
ws["S4"] = "Detractor NPS da leggere (da Medallia)"
ws["S4"].font = SEC_FONT
header(ws, 5, ["Data", "LDAP", "Punteggio", "Case ID", "Verbatim", "Team Lead", "Turno", "Già in log?"], col=19)
ws[f"S{V1}"] = (f'=IFERROR(FILTER(Detractor_NPS!A{D1}:H{D2},Detractor_NPS!B{D1}:B{D2}<>"",'
                f'{sel(f"Detractor_NPS!F{D1}:F{D2}")}),"Nessun detractor")')
ws.column_dimensions["S"].number_format = DATE
widths(ws, {"A": 24, "B": 20, "C": 9, "D": 9, "N": 10, "O": 11, "P": 11, "Q": 14, "R": 3,
            "S": 11, "T": 22, "U": 9, "V": 12, "W": 50, "X": 16, "Y": 9, "Z": 9, "AB": 20})
for c_ in "EFGHIJKLM":
    ws.column_dimensions[c_].width = 9
ws.freeze_panes = "B6"

# ---- Setup: controllo qualità dati (1)
ws = ws_set
ws["B50"] = "CONTROLLO QUALITÀ DATI (si aggiorna da solo)"
ws["B50"].font = Font(bold=True, size=12, color="C00000")
header(ws, 51, ["Controllo", "Valore", "Esito", "Dettaglio"], col=2)
TD = "Tableau!$I$2:$I"
dq = [
    ("Righe incollate in Tableau", "=COUNTA(Tableau!$A$2:$A)", 'IF(C{r}>0,"OK","ATTENZIONE")',
     '"Estrazione Dynamic Slicing per LDAP e Day"'),
    ("Date coperte da Tableau", f'=IFERROR(TEXT(MIN({TD}),"dd/mm/yyyy")&" - "&TEXT(MAX({TD}),"dd/mm/yyyy"),"-")',
     f'IF(AND(MIN({TD})<=$C$27,MAX({TD})>=$C$4),"OK","ATTENZIONE")',
     '"Deve coprire periodo + periodo di confronto ("&TEXT($C$27,"dd/mm")&" - "&TEXT($C$4,"dd/mm")&")"'),
    ("Giorni del periodo senza dati Tableau",
     f'=LET(d,SEQUENCE($C$4-$C$3+1,1,$C$3),ROWS(d)-SUM(MAP(d,LAMBDA(x,IF(COUNTIF({TD},x)>0,1,0)))))',
     'IF(C{r}=0,"OK","ATTENZIONE")',
     f'IFERROR(TEXTJOIN(", ",TRUE,MAP(FILTER(SEQUENCE($C$4-$C$3+1,1,$C$3),MAP(SEQUENCE($C$4-$C$3+1,1,$C$3),'
     f'LAMBDA(x,COUNTIF({TD},x)=0))),LAMBDA(x,TEXT(x,"dd/mm")))),"")'),
    ("Misure mancanti nell'estrazione (Setup C30:C38)",
     '=SUM(MAP($C$30:$C$38,LAMBDA(m,IF(COUNTIF(Tableau!$E$2:$E,m)=0,1,0))))',
     'IF(C{r}=0,"OK","ATTENZIONE")',
     'IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER($C$30:$C$38,MAP($C$30:$C$38,LAMBDA(m,COUNTIF(Tableau!$E$2:$E,m)=0))))),"")'),
    ("Agenti in Tableau (nel periodo) assenti da Turni_Log - ignorati",
     f'=IFERROR(ROWS(FILTER(UNIQUE(FILTER(Tableau!$A$2:$A,{TD}>=$C$3,{TD}<=$C$4)),'
     f'MAP(UNIQUE(FILTER(Tableau!$A$2:$A,{TD}>=$C$3,{TD}<=$C$4)),LAMBDA(a,COUNTIF({LOGB},a)=0)))),0)',
     'IF(C{r}=0,"OK","INFO")',
     f'IFERROR(TEXTJOIN(", ",TRUE,FILTER(UNIQUE(FILTER(Tableau!$A$2:$A,{TD}>=$C$3,{TD}<=$C$4)),'
     f'MAP(UNIQUE(FILTER(Tableau!$A$2:$A,{TD}>=$C$3,{TD}<=$C$4)),LAMBDA(a,COUNTIF({LOGB},a)=0)))),"")'),
    ("Giornate notte in Turni_Log senza righe Tableau", "=COUNTIFS(Turni_Log!$Y$2:$Y,1,Turni_Log!$Z$2:$Z,0)",
     'IF(C{r}=0,"OK","VERIFICA")',
     'IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER(Turni_Log!$B$2:$B,Turni_Log!$Y$2:$Y=1,Turni_Log!$Z$2:$Z=0))),"")'),
    ("Righe duplicate in Turni_Log (stesso LDAP e data)",
     '=COUNTIF(Turni_Log!$H$2:$H,"?*")-IFERROR(ROWS(UNIQUE(FILTER(Turni_Log!$H$2:$H,Turni_Log!$H$2:$H<>""))),0)',
     'IF(C{r}=0,"OK","ATTENZIONE")', '"Le righe doppie raddoppiano i numeri: cancellale da Turni_Log"'),
    ("Righe di TEST in Turni_Log", '=COUNTIF(Turni_Log!$G$2:$G,"TEST")', 'IF(C{r}=0,"OK","ATTENZIONE")',
     '"Dati fittizi: filtra Fonte = TEST e cancellali prima dell\'uso reale"'),
    ("Agenti del periodo senza Team Lead", f'=COUNTIFS(Agenti!$A$2:$A${AG_END},"?*",Agenti!$B$2:$B${AG_END},"")',
     'IF(C{r}=0,"OK","VERIFICA")', '"Servono per Vista_TL e completamento per TL"'),
    ("Colonne Medallia trovate (Setup C64:C68)", "=" + "+".join(f"ISNUMBER({m})" for m in MED_MATCH) + "",
     'IF(COUNTA(Medallia!$A$2:$A)=0,"INFO",IF(C{r}=5,"OK","ATTENZIONE"))',
     '"Su 5. INFO = foglio Medallia vuoto"'),
]
DQ1 = 52
for i, (lab, val, esito, det) in enumerate(dq):
    r = DQ1 + i
    ws.cell(r, 2, lab)
    ws.cell(r, 3, val).alignment = Alignment(horizontal="left")
    ws.cell(r, 4, "=" + esito.format(r=r))
    ws.cell(r, 5, "=" + det)
DQ2 = DQ1 + len(dq) - 1
for t, fill in (("ATTENZIONE", RED), ("VERIFICA", YELLOW), ("INFO", YELLOW), ("OK", GREEN)):
    ws.conditional_formatting.add(f"D{DQ1}:D{DQ2}", FormulaRule(formula=[f'D{DQ1}="{t}"'], fill=fill))
ws["B2"] = f"Controllo qualità dati: riga {DQ1 - 2}   |   Mappatura Medallia: riga 63"
ws["B2"].font = Font(italic=True, color="C00000")

# ---- Dashboard: stato qualità dati + sintesi automatica (13)
ws = ws_dash
DQCNT = f'COUNTIF(Setup!$D${DQ1}:$D${DQ2},"ATTENZIONE")'
ws["B3"] = (f'=IF({DQCNT}>0,"⚠ Qualità dati: "&{DQCNT}&" avvisi - vedi Setup riga {DQ1 - 2}",'
            '"Qualità dati: nessun avviso bloccante")')
ws.conditional_formatting.add("B3", FormulaRule(formula=['LEFT(B3,1)="⚠"'], fill=RED))
ws["B22"] = "Sintesi automatica"
ws["B22"].font = SEC_FONT
lst = lambda cond, none: f'IFERROR(TEXTJOIN(", ",TRUE,FILTER($B$5:$B$10,{cond})),"{none}")'
better = lambda sign: (f'MAP($F$5:$F$10,Setup!$D$20:$D$25,LAMBDA(f,d,IF(ISNUMBER(f),'
                       f'IF(d="Alto",f{sign}0,f{"<" if sign == ">" else ">"}0),FALSE)))')
capped = (f'SUM(MAP(Spotcheck!{SPO},Spotcheck!{SPP},LAMBDA(o,p,IF(ISNUMBER(o),MIN(o,p),0))))')
IN_T, OUT_T = '$H$5:$H$10="In target"', '$H$5:$H$10="Fuori target"'
MIG, PEG = '$O$5:$O$10="Migliora"', '$O$5:$O$10="Peggiora"'
summary = [
    ('="Periodo "&TEXT(Setup!C3,"dd/mm")&"-"&TEXT(Setup!C4,"dd/mm")&": "&TEXT(C13,"#,##0")&" assegnazioni in "&'
     'C16&" giornate-agente ("&C17&" agenti), periodo precedente "&TEXT(I13,"#,##0")&". "&Setup!C7&": "&'
     'TEXT(D18,"0.0")&" assegnazioni per giornata, "&Setup!C8&": "&TEXT(E18,"0.0")&"."'),
    f'="In target: "&{lst(IN_T, "nessun KPI")}&". Fuori target: "&{lst(OUT_T, "nessuno")}&"."',
    (f'="Confronto turni: "&Setup!C7&" meglio di "&Setup!C8&" su "&{lst(better(">"), "nessun KPI")}&'
     f'"; peggio su "&{lst(better("<"), "nessun KPI")}&"."'),
    (f'="Rispetto al periodo precedente migliorano: "&{lst(MIG, "nessun KPI")}&'
     f'"; peggiorano: "&{lst(PEG, "nessun KPI")}&"."'),
    (f'="Spotcheck: "&COUNTIF(Agenti!$Y$2:$Y${AG_END},">=3")&" agenti con almeno un KPI ROSSO, "&'
     f'SUM(Agenti!$Z$2:$Z${AG_END})&" casi da campionare, completamento "&'
     f'IFERROR(TEXT({capped}/SUM(Spotcheck!{SPO}),"0%"),"-")&". Agenti in ROSSO da 2+ settimane consecutive: "&'
     f'COUNTIF(Rossi_Consecutivi!$G${S1}:$G${S2},">=2")&"."'),
]
for i, f in enumerate(summary):
    ws[f"B{23 + i}"] = f
ws.column_dimensions["O"].width = 15
for c_ in "IJKLMN":
    ws.column_dimensions[c_].width = 12

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
    ("4. Medallia: incolla l'export delle risposte NPS in 'Medallia' da A1 (verifica una volta i nomi delle "
     "colonne in Setup C64:C68).", None),
    ("5. In 'Setup' imposta Data inizio / Data fine (di solito una settimana, dom-sab): il confronto usa il periodo "
     "precedente di pari durata. Controlla il box 'CONTROLLO QUALITÀ DATI' in Setup (riga 50).", None),
    ("6. Leggi 'Dashboard' e 'Spotcheck'; ogni TL usa 'Vista_TL' e registra i controlli in 'Spotcheck_Log'.", None),
    ("", None),
    ("FOGLI DI ANALISI", "sec"),
    ("Dashboard: KPI totale / 18-3 / 21-6, confronto con il periodo precedente (Δ e Migliora/Peggiora), volumi, "
     "sintesi automatica in testo, trend 8 settimane. In alto lo stato della qualità dati.", None),
    ("Spotcheck: priorità per agente + completamento per TL (casi dovuti vs registrati) + root cause.", None),
    ("Vista_TL: scegli il TL in B2: i suoi agenti da controllare, lo stato e i detractor NPS da leggere.", None),
    ("Rotazione: per gli agenti che fanno entrambi i turni, KPI nei giorni 18-3 vs nei giorni 21-6 "
     "(stesso agente = confronto più pulito tra i turni).", None),
    ("Carico: assegnazioni per giornata-agente per giorno della settimana, per settimana, per data e per agente.", None),
    ("Rossi_Consecutivi: per agente e KPI (Handoff, Pending, NPS, Reopen) quante settimane di fila è ROSSO; "
     "2+ settimane = problema strutturale, non un caso isolato.", None),
    ("Coaching: per ogni coaching registrato (Coaching fatto? = Sì) il KPI nei 14 giorni prima e nei 14 dopo.", None),
    ("Detractor_NPS: risposte 0-6 di Medallia per gli agenti notturni nel periodo, con verbatim: da qui si "
     "scelgono i casi NPS dello spotcheck.", None),
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
     "cancella quelle righe e sostituisci i dati in 'Tableau'. Sono di test anche le righe di 'Medallia' e le "
     "2 righe TEST di 'Spotcheck_Log'.", None),
    ("- Medallia: la struttura dell'export non è ancora verificata. Se i nomi delle colonne sono diversi, "
     "correggili in Setup C64:C68 (il box qualità dati dice quante colonne vengono trovate).", None),
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
for sh, row, col in ((ws_tab, 40001, "L"), (ws_tl, 5001, "AD"), (ws_new, 1101, "I"),
                     (ws_ag, AGENT_ROWS + 2, "AB"), (ws_spot, AGENT_ROWS + 6, "R"),
                     (ws_vtl, V2 + 2, "AD"), (ws_rot, R2 + 2, "Z"), (ws_car, CA2 + 5, "A"),
                     (ws_str, S2 + 2, "A"), (ws_coa, K2 + 2, "A"), (ws_det, D2 + 2, "A")):
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
