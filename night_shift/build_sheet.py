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
    ("Throughput", 0.58, "High", "EMEA H2'26 Targets - Resolutions 2"),
    ("NPS", 70, "High", "EMEA H2'26 Targets - Resolutions 2 Italian"),
    ("Handoff", 0.06, "Low", "EMEA H2'26 Targets - Resolutions 2"),
    ("Pending", 0.28, "Low", "EMEA H2'26 Targets - Resolutions 2"),
    ("Reopen", 0.10, "Low", "EMEA H2'26 Targets - Resolutions 2"),
    ("Recontact", 0.33, "Low", "EMEA H2'26 Targets - Resolutions 2"),
]
MEASURES = [  # etichetta, nome misura Tableau, colonna in Shift_Log
    ("Completed assignments", "cs_ambassador_case_assignments_completed", "P"),
    ("Solved cases", "cs_cases_solved_excluding_bot_only", "Q"),
    ("Pend events", "cs_ambassador_pend_events", "R"),
    ("Handoff", "cs_ambassador_transfer_handoffs", "S"),
    ("Reopen", "cs_case_reopen", "T"),
    ("Reopen denominator", "cs_cases_solved_excluding_bot_only", "U"),
    ("Recontact", "cs_interaction_specialist_recontacts", "V"),
    ("NPS responses", "cs_customer_nps_responses", "W"),
    ("Net promoters NPS", "cs_customer_nps_net_promoters", "X"),
]
SPOT_PARAMS = [
    ("Min assignments to assess Handoff/Pending", 20),
    ("Min denominator to assess Reopen", 15),
    ("Min responses to assess NPS", 5),
    ("k factor (standard deviations)", 1),
    ("Cases to sample per RED KPI", 3),
    ("Cases to sample per YELLOW KPI", 1),
]
LISTS = {
    "KPI": ["Handoff", "Pending", "NPS", "Reopen"],
    "Outcome": ["Correct", "Could improve", "Incorrect"],
    "Root cause": [
        "Procedure not followed", "Knowledge gap", "Unnecessary pend",
        "Avoidable handoff", "End-of-shift handoff", "Unclear communication",
        "User not responding", "Tool/system limitation", "Policy / outside agent control",
        "No issue",
    ],
    "Action": ["1:1 coaching", "Written feedback", "Team calibration", "None"],
    "Yes/No": ["Yes", "No"],
}
AGENT_ROWS = 150  # righe massime nei fogli Agenti / Spotcheck
FOCUS = ["Handoff", "Pending", "NPS", "Reopen"]
WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]  # WEEKDAY() 1..7
MEDALLIA_COLS = [  # campo, intestazione di colonna attesa nell'export Medallia (da verificare)
    ("Agent LDAP", "Agent LDAP"),
    ("Response date", "Response Date"),
    ("Score (0-10)", "Likelihood to Recommend"),
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
    for text, fill in (("RED", RED), ("YELLOW", YELLOW), ("GREEN", GREEN)):
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
# (per dimostrare che New_Week le scarta)
sched_night = [r for r in sched if any(str(x) in night_strings for x in r[8:15])]
sched_other = [r for r in sched if r not in sched_night][:15]
sched_sample = sched_night + sched_other

# Righe REALI della settimana 27/09-03/10 (stessa logica di New_Week)
week_dates = [as_date(x) for x in sched_hdr[8:15]]
log_rows = []
for r in sched_sample:
    if not is_night_row(r):
        continue
    for d, v in zip(week_dates, r[8:15]):
        log_rows.append([d, r[1], r[4], r[5], r[7], str(v) if v is not None else "", "REAL"])

# Righe di TEST: turni fittizi per gli agenti dell'estrazione di prova (13-26/09)
tab_days = defaultdict(set)
for r in tab:
    tab_days[r[0]].add(as_date(r[5]))
test_agents = sorted({r[0] for r in tab})
for i, a in enumerate(test_agents):
    s = sched_by_ldap.get(a)
    tier, tl, sa = (s[4], s[5], s[7]) if s else ("Resolutions 2 SE", "n/a", "No")
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
        return "Excluded"
    if shift in code_of:
        return code_of[shift]
    return "Other" if hours(shift) != (0, 0) else "Off"


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
            agg_prev["Total"][lab] += tab_val[(a, d, m)]
        agg_prev[att]["giornate"] += 1
        agg_prev["Total"]["giornate"] += 1
    if not (PERIOD_START <= d <= PERIOD_END):
        continue
    for lab, m, _ in MEASURES:
        v = tab_val[(a, d, m)]
        agg[att][lab] += v
        agg["Total"][lab] += v
        agent_agg[a][lab] += v
        agent_shift[(a, att)][lab] += v
    agent_shift[(a, att)]["giornate"] += 1
    agg[att]["giornate"] += 1
    agg["Total"]["giornate"] += 1
    start = d if hd >= prev[2] else d - dt.timedelta(days=1)
    wd = (start.weekday() + 1) % 7 + 1
    load_wd[(att, wd)][0] += 1
    load_wd[(att, wd)][1] += tab_val[(a, d, MEASURES[0][1])]


def kpis(s):
    f = lambda n, dd: round(n / dd, 4) if dd else None
    return {
        "Throughput": f(s["Solved cases"], s["Completed assignments"]),
        "NPS": f(100 * s["Net promoters NPS"], s["NPS responses"]),
        "Handoff": f(s["Handoff"], s["Completed assignments"]),
        "Pending": f(s["Pend events"], s["Completed assignments"]),
        "Reopen": f(s["Reopen"], s["Reopen denominator"]),
        "Recontact": f(s["Recontact"], s["Completed assignments"]),
        "Assegnazioni": s["Completed assignments"],
        "Agent-days": s["giornate"],
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
reference["reali_rows"] = sum(1 for r in log_rows if r[6] == "REAL")

# --------------------------------------------------------------------------- workbook
wb = openpyxl.Workbook()
ws_readme = wb.active
ws_readme.title = "Read_Me"
ws_dash = wb.create_sheet("Dashboard")
ws_spot = wb.create_sheet("Spotcheck")
ws_vtl = wb.create_sheet("TL_View")
ws_ag = wb.create_sheet("Agents")
ws_rot = wb.create_sheet("Rotation")
ws_car = wb.create_sheet("Workload")
ws_str = wb.create_sheet("Red_Streaks")
ws_coa = wb.create_sheet("Coaching")
ws_det = wb.create_sheet("NPS_Detractors")
ws_log = wb.create_sheet("Spotcheck_Log")
ws_set = wb.create_sheet("Setup")
ws_tab = wb.create_sheet("Tableau")
ws_med = wb.create_sheet("Medallia")
ws_sch = wb.create_sheet("Schedules")
ws_new = wb.create_sheet("New_Week")
ws_tl = wb.create_sheet("Shift_Log")

C1, C2 = "Setup!$C$7", "Setup!$C$8"  # codici turno
NIGHT_CODES = "Setup!$C$7:$C$12"

# ---- Setup
ws = ws_set
ws["B1"] = "SETUP - yellow cells can be edited"
ws["B1"].font = SEC_FONT
ws["B3"], ws["C3"] = "Period start date", PERIOD_START
ws["B4"], ws["C4"] = "Period end date", PERIOD_END
for c in ("C3", "C4"):
    ws[c].number_format = DATE
    ws[c].fill = INPUT_FILL
ws["D3"] = "Period analysed in Dashboard / Agents / Spotcheck (Tableau dates = working day)"
header(ws, 6, ["Hours in Schedules", "Shift code"], col=2)
for i in range(6):
    ws.cell(7 + i, 2).fill = ws.cell(7 + i, 3).fill = INPUT_FILL
for i, (s, c) in enumerate(SHIFTS):
    ws.cell(7 + i, 2, s)
    ws.cell(7 + i, 3, c)
ws["D7"] = "The first two codes (C7, C8) are the shifts compared in the split view"
ws["B14"] = "Excluded tiers"
ws["B14"].font = Font(bold=True)
for i in range(4):
    ws.cell(14 + i, 3).fill = INPUT_FILL
for i, t in enumerate(EXCLUDED_TIERS):
    ws.cell(14 + i, 3, t)
header(ws, 19, ["KPI", "Target", "Direction", "Source"], col=2)
for i, (k, t, d, f) in enumerate(TARGETS):
    ws.cell(20 + i, 2, k)
    c = ws.cell(20 + i, 3, t)
    c.fill = INPUT_FILL
    c.number_format = NUM1 if k == "NPS" else PCT
    ws.cell(20 + i, 4, d)
    ws.cell(20 + i, 5, f)
header(ws, 29, ["Field", "Tableau measure name (Measure Names)"], col=2)
for i, (lab, m, _) in enumerate(MEASURES):
    ws.cell(30 + i, 2, lab)
    c = ws.cell(30 + i, 3, m)
    c.fill = INPUT_FILL
ws["D35"] = ("TEMPORARY: the KPI dictionary uses cs_ambassador_case_solves. "
             "As soon as it is in the Tableau extraction, enter it in C35.")
ws["D35"].font = Font(bold=True, color="C00000")
ws["B27"], ws["C27"] = "Comparison period start (calculated)", "=C3-(C4-C3+1)"
ws["B28"], ws["C28"] = "Comparison period end (calculated)", "=C3-1"
for c in ("C27", "C28"):
    ws[c].number_format = DATE
    ws[c].fill = CALC_FILL
ws["D27"] = "Previous period of the same length (1 week = previous week)"
header(ws, 41, ["Spotcheck parameters", "Value"], col=2)
for i, (lab, v) in enumerate(SPOT_PARAMS):
    ws.cell(42 + i, 2, lab)
    c = ws.cell(42 + i, 3, v)
    c.fill = INPUT_FILL
header(ws, 2, list(LISTS.keys()), col=7)
for j, vals in enumerate(LISTS.values()):
    for i, v in enumerate(vals):
        ws.cell(3 + i, 7 + j, v).fill = INPUT_FILL
ws["G1"] = "Lists for the Spotcheck_Log dropdowns"
ws["G1"].font = Font(bold=True)
widths(ws, {"A": 2, "B": 44, "C": 40, "D": 12, "E": 38, "F": 3, "G": 12, "H": 14, "I": 30, "J": 18, "K": 8})

# ---- Tableau (incolla)
ws = ws_tab
header(ws, 1, ["Breakdown Selection 1", "Breakdown Selection 2", "Breakdown Selection 3",
               "Breakdown Selection 4", "Measure Names", "Column Breakdown Selection",
               "Measure Values", "Key (formula - do not edit)", "Date (formula)", "LDAP|date (formula)"])
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

# ---- New_Week (formula)
ws = ws_new
header(ws, 1, ["Date", "LDAP", "Tier", "Team Lead", "Special Assignment", "Shift (hours)", "Source"])
ws["I1"] = ("Copy A2:G (Paste special > values only) at the bottom of Shift_Log. "
            "Contains only agents with at least one night shift in Schedules, excluding the tiers in Setup.")
ws["I1"].font = Font(italic=True, color="7F7F7F")
ws["A2"] = (
    '=IFERROR(LET(ld,Schedules!B2:B3000,tr,Schedules!E2:E3000,tl,Schedules!F2:F3000,'
    'sa,Schedules!H2:H3000,sh,Schedules!I2:O3000,'
    'dts,MAP(Schedules!I1:O1,LAMBDA(x,IF(ISNUMBER(x),x,DATEVALUE(x)))),'
    'nm,FILTER(Setup!B7:B12,Setup!B7:B12<>""),'
    'keep,BYROW(sh,LAMBDA(r,SUMPRODUCT(--ISNUMBER(MATCH(r,nm,0))))),'
    'idx,FILTER(SEQUENCE(ROWS(ld)),ld<>"",keep>0,COUNTIF(Setup!C14:C17,tr)=0),'
    'MAKEARRAY(ROWS(idx)*7,7,LAMBDA(i,j,LET(r,INDEX(idx,INT((i-1)/7)+1),d,MOD(i-1,7)+1,'
    'CHOOSE(j,INDEX(dts,1,d),INDEX(ld,r),INDEX(tr,r),INDEX(tl,r),INDEX(sa,r),INDEX(sh,r,d),"REAL"))))),'
    '"No night shift found in Schedules")'
)
ws.column_dimensions["A"].number_format = DATE
widths(ws, {"A": 12, "B": 24, "C": 18, "D": 20, "E": 12, "F": 14, "G": 8})
ws.freeze_panes = "A2"

# ---- Shift_Log
ws = ws_tl
log_hdr = ["Date", "LDAP", "Tier", "Team Lead", "Special Assignment", "Shift (hours)", "Source",
           "Key", "Week (Sun)", "Shift code", "Hours on day D", "Hours on day D+1",
           "Shift code D-1", "Hours of D-1 shift on D", "Assigned shift (Tableau day)"]
log_hdr += [lab for lab, _, _ in MEASURES] + ["In period (night)", "Tableau data present",
                                               "Shift start date", "Shift start weekday (1=Sun)"]
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
ws["J2"] = ('=MAP(B2:B,C2:C,F2:F,LAMBDA(b,c,f,IF(b="","",IF(COUNTIF(Setup!$C$14:$C$17,c)>0,"Excluded",'
            'IFERROR(VLOOKUP(f,Setup!$B$7:$C$12,2,0),IF(REGEXMATCH(f&"","^\\d\\d:\\d\\d-\\d\\d:\\d\\d$"),"Other","Off"))))))')
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
ws["B2"] = ('="Period: "&TEXT(Setup!C3,"dd/mm/yyyy")&" - "&TEXT(Setup!C4,"dd/mm/yyyy")&'
            '"   |   compared with: "&TEXT(Setup!C27,"dd/mm/yyyy")&" - "&TEXT(Setup!C28,"dd/mm/yyyy")&"   (edit in Setup)"')
header(ws, 4, ["KPI", "Night total", "=Setup!C7", "=Setup!C8", '="Δ "&Setup!C7&" vs "&Setup!C8',
               "Target", "Total vs target"], col=2)
header(ws, 4, ["Prev. total", '="Prev. "&Setup!C7', '="Prev. "&Setup!C8', "Δ Total vs prev.",
               '="Δ "&Setup!C7&" vs prev."', '="Δ "&Setup!C8&" vs prev."', "Trend (total)"], col=9)
# somme di base (riga 30+)
SUM_ROW = 31
ws.cell(SUM_ROW - 1, 2, "Base sums (do not edit)").font = SEC_FONT
header(ws, SUM_ROW, ["Measure", "Night total", "=Setup!C7", "=Setup!C8", "Prev. total",
                     '="Prev. "&Setup!C7', '="Prev. "&Setup!C8'], col=2)
PA = 'Shift_Log!$A$2:$A,">="&Setup!$C$27,Shift_Log!$A$2:$A,"<="&Setup!$C$28'
sum_ref = {}
for i, (lab, m, col) in enumerate(MEASURES):
    r = SUM_ROW + 1 + i
    ws.cell(r, 2, lab)
    rng = f"Shift_Log!${col}$2:${col}"
    ws.cell(r, 3, f"=SUMIFS({rng},Shift_Log!$Y$2:$Y,1)")
    ws.cell(r, 4, f"=SUMIFS({rng},Shift_Log!$Y$2:$Y,1,Shift_Log!$O$2:$O,{C1})")
    ws.cell(r, 5, f"=SUMIFS({rng},Shift_Log!$Y$2:$Y,1,Shift_Log!$O$2:$O,{C2})")
    ws.cell(r, 6, f"=SUMIFS({rng},{PA})")  # le misure sono già 0 fuori dai turni notturni
    ws.cell(r, 7, f"=SUMIFS({rng},{PA},Shift_Log!$O$2:$O,{C1})")
    ws.cell(r, 8, f"=SUMIFS({rng},{PA},Shift_Log!$O$2:$O,{C2})")
    sum_ref[lab] = r
KPI_DEF = {  # num, den, moltiplicatore
    "Throughput": ("Solved cases", "Completed assignments", ""),
    "NPS": ("Net promoters NPS", "NPS responses", "100*"),
    "Handoff": ("Handoff", "Completed assignments", ""),
    "Pending": ("Pend events", "Completed assignments", ""),
    "Reopen": ("Reopen", "Reopen denominator", ""),
    "Recontact": ("Recontact", "Completed assignments", ""),
}
kpi_row = {}
for i, (k, t, d, f) in enumerate(TARGETS):
    r = 5 + i
    kpi_row[k] = r
    num, den, mul = KPI_DEF[k]
    ws.cell(r, 2, k).font = Font(bold=True)
    for col in "CDE":
        ws[f"{col}{r}"] = f'=IFERROR({mul}{col}{sum_ref[num]}/{col}{sum_ref[den]},"n/a")'
    ws[f"F{r}"] = f'=IFERROR(D{r}-E{r},"")'
    ws[f"G{r}"] = f"=Setup!C{20 + i}"
    ws[f"H{r}"] = (f'=IF(ISNUMBER(C{r}),IF(Setup!D{20 + i}="High",IF(C{r}>=G{r},"On target","Off target"),'
                   f'IF(C{r}<=G{r},"On target","Off target")),"")')
    for col, scol in zip("IJK", "FGH"):
        ws[f"{col}{r}"] = f'=IFERROR({mul}{scol}{sum_ref[num]}/{scol}{sum_ref[den]},"n/a")'
    for col, a, b in zip("LMN", "CDE", "IJK"):
        ws[f"{col}{r}"] = f'=IF(AND(ISNUMBER({a}{r}),ISNUMBER({b}{r})),{a}{r}-{b}{r},"n/a")'
    ws[f"O{r}"] = (f'=IF(ISNUMBER(L{r}),IF(L{r}=0,"Stable",IF((Setup!D{20 + i}="High")=(L{r}>0),'
                   f'"Better","Worse")),"n/a")')
    fmt = NUM1 if k == "NPS" else PCT
    for col in "CDEFGIJKLMN":
        ws[f"{col}{r}"].number_format = fmt
ws.conditional_formatting.add("H5:H10", FormulaRule(formula=['H5="Off target"'], fill=RED))
ws.conditional_formatting.add("H5:H10", FormulaRule(formula=['H5="On target"'], fill=GREEN))
ws.conditional_formatting.add("O5:O10", FormulaRule(formula=['O5="Worse"'], fill=RED))
ws.conditional_formatting.add("O5:O10", FormulaRule(formula=['O5="Better"'], fill=GREEN))
header(ws, 12, ["Prev. total", '="Prev. "&Setup!C7', '="Prev. "&Setup!C8'], col=9)
ws["B12"] = "Volumes"
ws["B12"].font = SEC_FONT
vols = [
    ("Completed assignments", "SUM", "Completed assignments"),
    ("Solved cases", "SUM", "Solved cases"),
    ("NPS responses", "SUM", "NPS responses"),
]
r = 13
for lab, _, src_lab in vols:
    ws.cell(r, 2, lab)
    for col in "CDE":
        ws[f"{col}{r}"] = f"={col}{sum_ref[src_lab]}"
    for col, scol in zip("IJK", "FGH"):
        ws[f"{col}{r}"] = f"={scol}{sum_ref[src_lab]}"
    r += 1
ws.cell(r, 2, "Agent-days (night shifts)")
ws[f"C{r}"] = "=COUNTIF(Shift_Log!$Y$2:$Y,1)"
ws[f"D{r}"] = f"=COUNTIFS(Shift_Log!$Y$2:$Y,1,Shift_Log!$O$2:$O,{C1})"
ws[f"E{r}"] = f"=COUNTIFS(Shift_Log!$Y$2:$Y,1,Shift_Log!$O$2:$O,{C2})"
ws[f"I{r}"] = f"=COUNTIFS({PA},Shift_Log!$O$2:$O,{C1})+COUNTIFS({PA},Shift_Log!$O$2:$O,{C2})"
ws[f"J{r}"] = f"=COUNTIFS({PA},Shift_Log!$O$2:$O,{C1})"
ws[f"K{r}"] = f"=COUNTIFS({PA},Shift_Log!$O$2:$O,{C2})"
r += 1
ws.cell(r, 2, "Distinct agents")
ws[f"C{r}"] = '=IFERROR(ROWS(UNIQUE(FILTER(Shift_Log!$B$2:$B,Shift_Log!$Y$2:$Y=1))),0)'
ws[f"D{r}"] = f'=IFERROR(ROWS(UNIQUE(FILTER(Shift_Log!$B$2:$B,Shift_Log!$Y$2:$Y=1,Shift_Log!$O$2:$O={C1}))),0)'
ws[f"E{r}"] = f'=IFERROR(ROWS(UNIQUE(FILTER(Shift_Log!$B$2:$B,Shift_Log!$Y$2:$Y=1,Shift_Log!$O$2:$O={C2}))),0)'
PF = "Shift_Log!$A$2:$A>=Setup!$C$27,Shift_Log!$A$2:$A<=Setup!$C$28"
ws[f"I{r}"] = (f'=IFERROR(ROWS(UNIQUE(FILTER(Shift_Log!$B$2:$B,{PF},'
               f'((Shift_Log!$O$2:$O={C1})+(Shift_Log!$O$2:$O={C2}))>0))),0)')
ws[f"J{r}"] = f'=IFERROR(ROWS(UNIQUE(FILTER(Shift_Log!$B$2:$B,{PF},Shift_Log!$O$2:$O={C1}))),0)'
ws[f"K{r}"] = f'=IFERROR(ROWS(UNIQUE(FILTER(Shift_Log!$B$2:$B,{PF},Shift_Log!$O$2:$O={C2}))),0)'
r += 1
ws.cell(r, 2, "Assignments per agent-day")
ws[f"C{r}"] = f'=IFERROR(C13/C16,"n/a")'
ws[f"D{r}"] = f'=IFERROR(D13/D16,"n/a")'
ws[f"E{r}"] = f'=IFERROR(E13/E16,"n/a")'
for col in "IJK":
    ws[f"{col}{r}"] = f'=IFERROR({col}13/{col}16,"n/a")'
for col in "CDEIJK":
    ws[f"{col}{r}"].number_format = NUM1
ws["B20"] = ("Note: a Tableau day (day D) is assigned to the shift that works the most hours on it: "
             "18-3 of day D (6h) or 21-6 of day D-1 (6h). "
             "Special Assignments included; tiers excluded in Setup.")
ws["B20"].font = Font(italic=True, color="7F7F7F")

# Trend settimanale
TR = 44
ws.cell(TR - 1, 2, "Weekly trend (last 8 weeks up to the end date)").font = SEC_FONT
tr_hdr = ["Week (Sun)"]
for k, *_ in TARGETS:
    tr_hdr += [f"{k} Tot", f"{k} 18-3", f"{k} 21-6"]
tr_hdr += ["Assign. Tot", "Assign. 18-3", "Assign. 21-6"]
header(ws, TR, tr_hdr, col=2)
W1, W8 = TR + 1, TR + 8
ws.cell(W1, 2, f"=SEQUENCE(8,1,Setup!$C$4-WEEKDAY(Setup!$C$4)+1-49,7)")
for r in range(W1, W8 + 1):
    ws.cell(r, 2).number_format = DATE
WK = f"$B${W1}:$B${W8}"


def s(col, code_ref=None):
    base = f"SUMIFS(Shift_Log!${col}$2:${col},Shift_Log!$I$2:$I,w,Shift_Log!$O$2:$O,"
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
ag_hdr = ["LDAP", "Team Lead", "Days " + SHIFTS[0][1], "Days " + SHIFTS[1][1], "Main shift"]
ag_hdr += [lab for lab, _, _ in MEASURES]
ag_hdr += ["Throughput", "NPS", "Handoff", "Pending", "Reopen", "Recontact",
           "Flag Handoff", "Flag Pending", "Flag NPS", "Flag Reopen", "Score", "Cases to sample"]
header(ws, 1, ag_hdr)
N = AGENT_ROWS + 1
A = f"A2:A{N}"
LOG = "Shift_Log!$B$2:$B"
PER = "Shift_Log!$Y$2:$Y"
ws["A2"] = f'=IFERROR(SORT(UNIQUE(FILTER({LOG},{PER}=1))),"")'
ws["B2"] = f'=MAP({A},LAMBDA(a,IF(a="","",IFERROR(VLOOKUP(a,Shift_Log!$B$2:$D,3,0),""))))'
ws["C2"] = f'=MAP({A},LAMBDA(a,IF(a="","",COUNTIFS({LOG},a,{PER},1,Shift_Log!$O$2:$O,{C1}))))'
ws["D2"] = f'=MAP({A},LAMBDA(a,IF(a="","",COUNTIFS({LOG},a,{PER},1,Shift_Log!$O$2:$O,{C2}))))'
ws["E2"] = (f'=ARRAYFORMULA(IF({A}="","",IF((C2:C{N}>0)*(D2:D{N}>0),"Mixed",'
            f'IF(C2:C{N}>0,{C1},{C2}))))')
for i, (lab, m, col) in enumerate(MEASURES):
    L = openpyxl.utils.get_column_letter(6 + i)
    ws[f"{L}2"] = f'=MAP({A},LAMBDA(a,IF(a="","",SUMIFS(Shift_Log!${col}$2:${col},{LOG},a,{PER},1))))'
# F Assegn, G Risolti, H Pend, I Handoff, J Reopen, K Den reopen, L Recontact, M Risp NPS, N Net NPS
g = lambda c: f"{c}2:{c}{N}"
MINA, MINR, MINN = "Setup!$C$42", "Setup!$C$43", "Setup!$C$44"
ws["O2"] = f'=ARRAYFORMULA(IF({A}="","",IF({g("F")}>0,{g("G")}/{g("F")},"n/a")))'
ws["P2"] = f'=ARRAYFORMULA(IF({A}="","",IF({g("M")}>={MINN},100*{g("N")}/{g("M")},"n.s.")))'
ws["Q2"] = f'=ARRAYFORMULA(IF({A}="","",IF({g("F")}>={MINA},{g("I")}/{g("F")},"n.s.")))'
ws["R2"] = f'=ARRAYFORMULA(IF({A}="","",IF({g("F")}>={MINA},{g("H")}/{g("F")},"n.s.")))'
ws["S2"] = f'=ARRAYFORMULA(IF({A}="","",IF({g("K")}>={MINR},{g("J")}/{g("K")},"n.s.")))'
ws["T2"] = f'=ARRAYFORMULA(IF({A}="","",IF({g("F")}>0,{g("L")}/{g("F")},"n/a")))'


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
    return (f'=ARRAYFORMULA(IF({A}="","",IF(ISNUMBER({v}),IF({red},"RED",IF({yel},"YELLOW","GREEN")),"n.s.")))')


ws["U2"] = flag("Q", 22, f"Dashboard!$C${kpi_row['Handoff']}", True)
ws["V2"] = flag("R", 23, f"Dashboard!$C${kpi_row['Pending']}", True)
ws["W2"] = flag("P", 21, f"Dashboard!$C${kpi_row['NPS']}", False)
ws["X2"] = flag("S", 24, f"Dashboard!$C${kpi_row['Reopen']}", True)
cnt = lambda t: "+".join(f'({g(c)}="{t}")' for c in "UVWX")
ws["Y2"] = f'=ARRAYFORMULA(IF({A}="","",3*({cnt("RED")})+({cnt("YELLOW")})))'
ws["Z2"] = (f'=ARRAYFORMULA(IF({A}="","",Setup!$C$46*({cnt("RED")})+Setup!$C$47*({cnt("YELLOW")})))')
for c in "OQRST":
    ws.column_dimensions[c].number_format = PCT
ws.column_dimensions["P"].number_format = NUM1
flag_cf(ws, f"U2:X{N}", "U2")
widths(ws, {"A": 24, "B": 20, "E": 12})
ws.freeze_panes = "B2"

# ---- Spotcheck
ws = ws_spot
ws["A1"] = "SPOTCHECK - focus on Handoff, Pending, NPS and Reopen"
ws["A1"].font = Font(bold=True, size=14, color="1F3864")
ws["A2"] = '="Period: "&TEXT(Setup!C3,"dd/mm/yyyy")&" - "&TEXT(Setup!C4,"dd/mm/yyyy")&"  |  sorted by priority"'
ws["A3"] = ("RED = off target and beyond night average ± k·std.dev  |  YELLOW = off target  |  "
            "GREEN = on target  |  n.s. = not enough volume (thresholds in Setup). "
            "Log every checked case in Spotcheck_Log.")
ws["A3"].font = Font(italic=True, color="7F7F7F")
sp_hdr = ["LDAP", "Team Lead", "Shift", "Assign.", "Handoff", "Flag", "Pending", "Flag", "NPS",
          "NPS resp.", "Flag", "Reopen", "Flag", "Score", "Cases to sample",
          "Spotchecks logged", "Status"]
header(ws, 4, sp_hdr)
M = f"{AGENT_ROWS + 1}"
cols = ["A", "B", "E", "F", "Q", "U", "R", "V", "P", "M", "W", "S", "X", "Y", "Z"]
stack = ",".join(f"Agents!{c}2:{c}{M}" for c in cols)
ws["A5"] = f'=IFERROR(SORT(FILTER(HSTACK({stack}),Agents!A2:A{M}<>""),14,FALSE,15,FALSE,1,TRUE),"")'
E = 5 + AGENT_ROWS - 1
ws["P5"] = (f'=MAP(A5:A{E},LAMBDA(a,IF(a="","",COUNTIFS(Spotcheck_Log!$B$2:$B,a,'
            f'Spotcheck_Log!$A$2:$A,">="&Setup!$C$3,Spotcheck_Log!$A$2:$A,"<="&Setup!$C$4))))')
ws["Q5"] = (f'=ARRAYFORMULA(IF(A5:A{E}="","",IF(O5:O{E}=0,"-",IF(P5:P{E}>=O5:O{E},"Done",'
            f'"To do ("&(O5:O{E}-P5:P{E})&")"))))')
for c in "EGL":
    ws.column_dimensions[c].number_format = PCT
ws.column_dimensions["I"].number_format = NUM1
for c in "FHKM":
    flag_cf(ws, f"{c}5:{c}{E}", f"{c}5")
ws.conditional_formatting.add(f"Q5:Q{E}", FormulaRule(formula=['LEFT(Q5,5)="To do"'], fill=RED))
ws.conditional_formatting.add(f"Q5:Q{E}", FormulaRule(formula=['Q5="Done"'], fill=GREEN))
# riepilogo root cause
ws["S4"], ws["T4"] = "Root cause (period cases)", "No. of cases"
header(ws, 4, ["Root cause (period cases)", "No. of cases", "% of total"], col=19)
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
log2_hdr = ["Case date", "LDAP", "Shift", "KPI", "Case ID / link", "Outcome", "Root cause",
            "Action", "Coaching done?", "Spotcheck date", "Done by (TL)", "Notes"]
header(ws, 1, log2_hdr)
ws.column_dimensions["A"].number_format = DATE
ws.column_dimensions["J"].number_format = DATE
dv = [
    ("B", f"=Agents!$A$2:$A${AGENT_ROWS + 1}"),
    ("C", "=Setup!$C$7:$C$12"),
    ("D", f"=Setup!$G$3:$G${2 + len(LISTS['KPI'])}"),
    ("F", f"=Setup!$H$3:$H${2 + len(LISTS['Outcome'])}"),
    ("G", f"=Setup!$I$3:$I${2 + len(LISTS['Root cause'])}"),
    ("H", f"=Setup!$J$3:$J${2 + len(LISTS['Action'])}"),
    ("I", f"=Setup!$K$3:$K${2 + len(LISTS['Yes/No'])}"),
]
for col, formula in dv:
    v = DataValidation(type="list", formula1=formula, allow_blank=True)
    v.add(f"{col}2:{col}2000")
    ws.add_data_validation(v)
coach_agents = [a for a in test_agents if agent_agg.get(a, {}).get("Completed assignments", 0) >= 20][:2]
for i, (a, k) in enumerate(zip(coach_agents, ["Handoff", "Pending"])):
    ws.append([PERIOD_START - dt.timedelta(days=3 - i), a, SHIFTS[i][1], k, f"TEST-{2000 + i}", "Could improve",
               "Avoidable handoff" if k == "Handoff" else "Unnecessary pend", "1:1 coaching", "Yes",
               PERIOD_START, "Test TL", "TEST - sample row for the Coaching sheet: delete it"])
    for col in (1, 10):
        ws.cell(ws.max_row, col).number_format = DATE
widths(ws, {"A": 12, "B": 24, "C": 8, "D": 10, "E": 34, "F": 13, "G": 30, "H": 18, "I": 10,
            "J": 12, "K": 18, "L": 40})
ws.freeze_panes = "A2"

# ---- Riferimenti comuni per i fogli di analisi
TL_ = "Shift_Log!"
LOGB, LOGY, LOGO = "Shift_Log!$B$2:$B", "Shift_Log!$Y$2:$Y", "Shift_Log!$O$2:$O"
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
title(ws, "ROTATION - same agent on 18-3 and on 21-6",
      "Only agents who worked both shifts in the period. Δ = " + SHIFTS[0][1] + " minus "
      + SHIFTS[1][1] + ". n.s. = volume below the Setup thresholds (C42:C44). "
      "The 'Total mixed agents' row sums numerators and denominators of all these agents.")
rot_hdr = ["LDAP", "Team Lead", '="Days "&Setup!C7', '="Days "&Setup!C8',
           '="Assign./agent-day "&Setup!C7', '="Assign./agent-day "&Setup!C8']
for k in KPI_NAMES:
    rot_hdr += [f'="{k} "&Setup!C7', f'="{k} "&Setup!C8', f"Δ {k}"]
header(ws, 4, rot_hdr)
R1, R2 = 6, 6 + AGENT_ROWS - 1
RA = f"A{R1}:A{R2}"
ws["A5"] = "Total mixed agents"
ws["A5"].font = Font(bold=True)
ws[f"A{R1}"] = (f'=IFERROR(FILTER(Agents!A2:A{AG_END},Agents!E2:E{AG_END}="Mixed"),'
                '"No agent on both shifts in the period")')
ws[f"B{R1}"] = f'=MAP({RA},LAMBDA(a,IF(a="","",IFERROR(VLOOKUP(a,Agents!$A$2:$B${AG_END},2,0),""))))'
sa_ = lambda col, code: f"SUMIFS(Shift_Log!${col}$2:${col},{LOGB},a,{LOGY},1,{LOGO},{code})"
tot_ = lambda col, code: f"SUM(MAP({RA},LAMBDA(a,IF(a=\"\",0,{sa_(col, code)}))))"
for col, code in (("C", C1), ("D", C2)):
    ws[f"{col}{R1}"] = f'=MAP({RA},LAMBDA(a,IF(a="","",COUNTIFS({LOGB},a,{LOGY},1,{LOGO},{code}))))'
    ws[f"{col}5"] = f"=SUM({col}{R1}:{col}{R2})"
for col, dcol, code in (("E", "C", C1), ("F", "D", C2)):
    ws[f"{col}{R1}"] = (f'=MAP({RA},{dcol}{R1}:{dcol}{R2},LAMBDA(a,n,IF(a="","",IF(n>0,'
                        f'{sa_("P", code)}/n,"-"))))')
    ws[f"{col}5"] = f'=IFERROR({tot_("P", code)}/{dcol}5,"n/a")'
    for r in range(5, R2 + 1):
        ws[f"{col}{r}"].number_format = NUM1
c = 7
for k in KPI_NAMES:
    num, den, mul = KPI_DEF[k]
    for code in (C1, C2):
        L = L_(c)
        ws[f"{L}{R1}"] = (f'=MAP({RA},LAMBDA(a,IF(a="","",'
                          f'{kpi_expr(k, lambda cc, code=code: sa_(cc, code))})))')
        ws[f"{L}5"] = f'=IFERROR({mul}{tot_(MCOL[num], code)}/{tot_(MCOL[den], code)},"n/a")'
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
title(ws, "WORKLOAD - assignments per agent-day",
      "Agent-day = Tableau day assigned to a night shift in the period. Weekday / date = the day the shift STARTS "
      "(e.g. a 21-6 starting on Monday is 'Mon'). Weeks = week (Sun) of the Tableau day.")
blk = ["Agent-days", "Assign.", "Assign./agent-day"]
car_hdr = [f'="{b} "&Setup!C{7 + j}' for j in range(2) for b in blk] + [f"{b} total" for b in blk]


# a) per giorno della settimana
ws["A4"] = "By weekday"
ws["A4"].font = SEC_FONT
header(ws, 5, ["Weekday"] + car_hdr)
for i, wd in enumerate(WEEKDAYS + ["Total"]):
    r = 6 + i
    ws[f"A{r}"] = wd
    extra = f",Shift_Log!$AB$2:$AB,{i + 1}" if wd != "Total" else ""
    for j, code in enumerate((C1, C2, None)):
        cg, ca, cr = (L_(2 + 3 * j + x) for x in range(3))
        if code:
            ws[f"{cg}{r}"] = f"=COUNTIFS({LOGY},1,{LOGO},{code}{extra})"
            ws[f"{ca}{r}"] = f"=SUMIFS(Shift_Log!$P$2:$P,{LOGY},1,{LOGO},{code}{extra})"
        else:
            ws[f"{cg}{r}"] = f"=B{r}+E{r}"
            ws[f"{ca}{r}"] = f"=C{r}+F{r}"
        ws[f"{cr}{r}"] = f'=IFERROR({ca}{r}/{cg}{r},"-")'
        ws[f"{cr}{r}"].number_format = NUM1
    if wd == "Total":
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
                               f'SUMIFS(Shift_Log!$P$2:$P,{LOGY},1,{LOGO},{code},{crit}))))')
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
ws["A16"] = "By week"
ws["A16"].font = SEC_FONT
header(ws, 17, ["Week (Sun)"] + car_hdr)
WS_ = lambda x: f"({x}-WEEKDAY({x})+1)"
ws["A18"] = (f"=SEQUENCE(INT(({WS_('Setup!$C$4')}-{WS_('Setup!$C$3')})/7)+1,1,{WS_('Setup!$C$3')},7)")
load_map(ws, 18, 37, "A", "Shift_Log!$I$2:$I,k")
for r in range(18, 38):
    ws[f"A{r}"].number_format = DATE
# c) per giorno (data di inizio turno)
ws["A40"] = "By day (shift start date)"
ws["A40"].font = SEC_FONT
header(ws, 41, ["Shift start date"] + car_hdr + ["Weekday"])
ws["A42"] = f'=IFERROR(SORT(UNIQUE(FILTER(Shift_Log!$AA$2:$AA,{LOGY}=1))),"")'
load_map(ws, 42, 141, "A", "Shift_Log!$AA$2:$AA,k")
ws["K42"] = ('=MAP(A42:A141,LAMBDA(k,IF(k="","",CHOOSE(WEEKDAY(k),'
             + ",".join(f'"{w}"' for w in WEEKDAYS) + "))))")
for r in range(42, 142):
    ws[f"A{r}"].number_format = DATE
# d) per agente
ws["M4"] = "By agent (assignments per agent-day)"
ws["M4"].font = SEC_FONT
header(ws, 5, ["LDAP", "Main shift", "Agent-days", "Assign./agent-day", '="Assign./agent-day "&Setup!C7',
               '="Assign./agent-day "&Setup!C8'] + [f"{w} (shift start)" for w in WEEKDAYS], col=13)
CA1, CA2 = 6, 6 + AGENT_ROWS - 1
MA = f"M{CA1}:M{CA2}"
ws[f"M{CA1}"] = f"=ARRAYFORMULA(Agents!A2:A{AG_END})"
ws[f"N{CA1}"] = f"=ARRAYFORMULA(Agents!E2:E{AG_END})"
ws[f"O{CA1}"] = f'=ARRAYFORMULA(IF({MA}="","",Agents!C2:C{AG_END}+Agents!D2:D{AG_END}))'
ws[f"P{CA1}"] = (f'=MAP({MA},O{CA1}:O{CA2},LAMBDA(a,n,IF(a="","",IF(n>0,'
                 f'SUMIFS(Shift_Log!$P$2:$P,{LOGB},a,{LOGY},1)/n,"-"))))')
for col, code in (("Q", C1), ("R", C2)):
    ws[f"{col}{CA1}"] = (f'=MAP({MA},LAMBDA(a,IF(a="","",IFERROR(SUMIFS(Shift_Log!$P$2:$P,{LOGB},a,{LOGY},1,'
                         f'{LOGO},{code})/COUNTIFS({LOGB},a,{LOGY},1,{LOGO},{code}),"-"))))')
for i in range(7):
    col = L_(19 + i)
    crit = f"{LOGB},a,{LOGY},1,Shift_Log!$AB$2:$AB,{i + 1}"
    ws[f"{col}{CA1}"] = (f'=MAP({MA},LAMBDA(a,IF(a="","",IFERROR(SUMIFS(Shift_Log!$P$2:$P,{crit})'
                         f'/COUNTIFS({crit}),"-"))))')
for c in range(16, 26):
    for r in range(CA1, CA2 + 1):
        ws.cell(r, c).number_format = NUM1
    ws.column_dimensions[L_(c)].width = 10
widths(ws, {"A": 16, "K": 7, "L": 3, "M": 24, "N": 11, "O": 9})
for c in "BCDEFGHIJ":
    ws.column_dimensions[c].width = 11

# ---- Red_Streaks (10)
ws = ws_str
title(ws, "KPIs RED IN CONSECUTIVE WEEKS (last 8 weeks)",
      "Same rule as the Spotcheck sheet applied week by week (night average and std. dev. "
      "of that week). The count starts from the latest week and stops at the first non-RED.")
header(ws, 4, ["LDAP", "Team Lead"] + [f"Weeks RED {k}" for k in FOCUS] + ["Max"])
S1, S2 = 5, 5 + AGENT_ROWS - 1
SA = f"$A${S1}:$A${S2}"
ws[f"A{S1}"] = f"=ARRAYFORMULA(Agents!A2:A{AG_END})"
ws[f"B{S1}"] = f"=ARRAYFORMULA(Agents!B2:B{AG_END})"
c = 10
names = ["ra", "rb", "rc", "rd", "re", "rf", "rg", "rh"]
for q, k in enumerate(FOCUS):
    num, den, mul = KPI_DEF[k]
    idx = KPI_NAMES.index(k)
    trend_col = L_(3 + 3 * idx)  # colonna "k Tot" del trend in Dashboard
    ws.cell(3, c, f"{k} - weekly value").font = Font(bold=True)
    ws.cell(3, c + 8, f"{k} - weekly flag").font = Font(bold=True)
    flag_cols = []
    for w in range(8):
        vc, fc = L_(c + w), L_(c + 8 + w)
        for col in (vc, fc):
            ws[f"{col}4"] = f"=Dashboard!$B${W1 + w}"
            ws[f"{col}4"].number_format = "dd/mm"
            ws[f"{col}4"].fill, ws[f"{col}4"].font = HDR_FILL, HDR_FONT
            ws.column_dimensions[col].width = 8
        sw = lambda cc, vc=vc: f"SUMIFS(Shift_Log!${cc}$2:${cc},{LOGB},a,Shift_Log!$I$2:$I,{vc}$4)"
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
                           f'IF({red},"RED",IF({yel},"YELLOW","GREEN")),"n.s.")))')
        flag_cols.append(f"{fc}{S1}:{fc}{S2}")
    flag_cf(ws, f"{L_(c + 8)}{S1}:{L_(c + 15)}{S2}", f"{L_(c + 8)}{S1}")
    # striscia: dall'ultima settimana (rh) all'indietro
    expr = "8"
    for j in range(0, 8):
        expr = f'IF({names[j]}<>"RED",{7 - j},{expr})'
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
title(ws, "COACHING EFFECT - KPI 2 weeks before vs 2 weeks after",
      "Spotcheck_Log rows with 'Coaching done?' = Yes and Spotcheck date filled in. Before = 14 days before the "
      "Spotcheck date; After = 14 days from the next day. Only days assigned to night shifts. "
      "'(partial)' = the 14 days after are not all in the Tableau extraction yet.")
header(ws, 4, ["LDAP", "KPI", "Coaching date", "KPI before", "Volume before", "KPI after", "Volume after",
               "Δ (after - before)", "Outcome"])
K1, K2 = 5, 204
KA = f"A{K1}:A{K2}"
SL = "Spotcheck_Log!"
ws[f"A{K1}"] = (f'=IFERROR(SORT(UNIQUE(FILTER(HSTACK({SL}$B$2:$B,{SL}$D$2:$D,{SL}$J$2:$J),'
                f'{SL}$I$2:$I="Yes",{SL}$J$2:$J<>"",{SL}$B$2:$B<>"")),3,FALSE,1,TRUE),'
                '"No coaching logged in Spotcheck_Log")')


def win(s_, e_, what):
    sw = lambda cc: f'SUMIFS(Shift_Log!${cc}$2:${cc},{LOGB},a,Shift_Log!$A$2:$A,">="&({s_}),Shift_Log!$A$2:$A,"<="&({e_}))'
    num = "SWITCH(kp," + ",".join(f'"{k}",{KPI_DEF[k][2]}{sw(MCOL[KPI_DEF[k][0]])}' for k in KPI_NAMES) + ",0)"
    den = "SWITCH(kp," + ",".join(f'"{k}",{sw(MCOL[KPI_DEF[k][1]])}' for k in KPI_NAMES) + ",0)"
    body = f'LET(nx,{num},dx,{den},IF(dx>0,nx/dx,"n/a"))' if what == "kpi" else den
    return f'=MAP({KA},B{K1}:B{K2},C{K1}:C{K2},LAMBDA(a,kp,dc,IF(OR(a="",NOT(ISNUMBER(dc))),"",{body})))'


ws[f"D{K1}"] = win("dc-14", "dc-1", "kpi")
ws[f"E{K1}"] = win("dc-14", "dc-1", "vol")
ws[f"F{K1}"] = win("dc+1", "dc+14", "kpi")
ws[f"G{K1}"] = win("dc+1", "dc+14", "vol")
ws[f"H{K1}"] = f'=MAP(D{K1}:D{K2},F{K1}:F{K2},LAMBDA(x,y,IF(AND(ISNUMBER(x),ISNUMBER(y)),y-x,"")))'
ws[f"I{K1}"] = (f'=LET(mx,MAX(Tableau!$I$2:$I),MAP({KA},B{K1}:B{K2},C{K1}:C{K2},H{K1}:H{K2},'
                'LAMBDA(a,kp,dc,dl,IF(OR(a="",NOT(ISNUMBER(dc))),"",IF(mx<dc+14,"(partial) ","")&'
                'IF(dl="","n/a",IF(dl=0,"Unchanged",IF(OR(kp="NPS",kp="Throughput")=(dl>0),'
                '"Improved","Worsened")))))))')
ws["K4"] = "Summary"
ws["K4"].font = SEC_FONT
for i, (lab, pat) in enumerate((("Improved", "*Improved"), ("Worsened", "*Worsened"),
                                ("Unchanged", "*Unchanged"), ("n/a", "*n/a"))):
    ws[f"K{5 + i}"] = lab
    ws[f"L{5 + i}"] = f'=COUNTIF(I{K1}:I{K2},"{pat}")'
for r in range(K1, K2 + 1):
    ws[f"C{r}"].number_format = DATE
    for col in "DFH":
        ws[f"{col}{r}"].number_format = "0.00"
ws.conditional_formatting.add(f"I{K1}:I{K2}", FormulaRule(formula=[f'ISNUMBER(SEARCH("Improved",I{K1}))'], fill=GREEN))
ws.conditional_formatting.add(f"I{K1}:I{K2}", FormulaRule(formula=[f'ISNUMBER(SEARCH("Worsened",I{K1}))'], fill=RED))
ws["A3"] = "KPI: Handoff/Pending/Reopen as ratios (0.08 = 8%), NPS in points. For NPS and Throughput Δ > 0 = better."
ws["A3"].font = Font(italic=True, color="7F7F7F")
widths(ws, {"A": 24, "B": 11, "C": 12, "D": 10, "E": 10, "F": 10, "G": 10, "H": 12, "I": 22, "J": 3, "K": 12})
ws.freeze_panes = "A5"

# ---- Medallia (incolla) + mappatura in Setup (15)
ws = ws_med
med_hdr = [h for _, h in MEDALLIA_COLS]
header(ws, 1, med_hdr)
ws["G1"] = ("Paste the Medallia export here from A1 (headers in row 1). The columns used are set in "
            "Setup C64:C68. The current rows are TEST DATA to delete.")
ws["G1"].font = Font(italic=True, color="C00000")
ws.column_dimensions["B"].number_format = DATE
med_agents = [a for a in test_agents if a in agent_agg][:8]
for i in range(16):
    a = med_agents[i % len(med_agents)]
    d = PERIOD_START + dt.timedelta(days=i % 5) if i < 12 else PREV_START + dt.timedelta(days=i % 5)
    ws.append([a, d, [2, 9, 5, 10, 0, 6, 8, 3, 7, 10, 4, 1, 2, 9, 6, 0][i], f"TEST - sample comment {i + 1}",
               f"TEST-{1000 + i}"])
    ws.cell(ws.max_row, 2).number_format = DATE
widths(ws, {"A": 24, "B": 14, "C": 12, "D": 50, "E": 14})
ws.freeze_panes = "A2"

ws = ws_set
header(ws, 63, ["Medallia: field", "Column header in the Medallia sheet"], col=2)
for i, (lab, h) in enumerate(MEDALLIA_COLS):
    ws.cell(64 + i, 2, lab)
    ws.cell(64 + i, 3, h).fill = INPUT_FILL
ws.cell(69, 2, "Detractor = score less than or equal to")
ws.cell(69, 3, DETRACTOR_MAX).fill = INPUT_FILL
ws["D64"] = ("TO VERIFY: enter the exact column names of the Medallia export. "
             "LDAP can also be an email (the part after @ is ignored).")
ws["D64"].font = Font(bold=True, color="C00000")
MED_MATCH = [f"MATCH(Setup!$C${64 + i},Medallia!$A$1:$AZ$1,0)" for i in range(5)]

# ---- NPS_Detractors (15)
ws = ws_det
title(ws, "NPS DETRACTORS OF NIGHT AGENTS (from Medallia)",
      "Responses with score ≤ threshold (Setup C69), date in the period, agents in Shift_Log in the period. "
      "Use them to pick the NPS cases to check; 'Already in Spotcheck_Log' looks for the Case ID in column E.")
header(ws, 4, ["Date", "LDAP", "Score", "Case ID", "Verbatim", "Team Lead", "Assigned shift",
               "Already in Spotcheck_Log?"])
D1, D2 = 5, 504
ws[f"A{D1}"] = (
    '=IFERROR(LET(m,Medallia!$A$2:$AZ$5000,'
    f'l,MAP(INDEX(m,0,{MED_MATCH[0]}),LAMBDA(x,IF(x="","",TRIM(REGEXREPLACE(x&"","@.*",""))))),'
    f'dn,MAP(INDEX(m,0,{MED_MATCH[1]}),LAMBDA(x,IF(ISNUMBER(x),INT(x),IFERROR(INT(DATEVALUE(x)),"")))),'
    f'sn,MAP(INDEX(m,0,{MED_MATCH[2]}),LAMBDA(x,IF(x="","",IFERROR(VALUE(x),"")))),'
    f'vb,INDEX(m,0,{MED_MATCH[3]}),id,INDEX(m,0,{MED_MATCH[4]}),'
    'ok,MAP(l,dn,sn,LAMBDA(a,b,c,IF(AND(a<>"",ISNUMBER(b),ISNUMBER(c)),'
    f'AND(b>=Setup!$C$3,b<=Setup!$C$4,c<=Setup!$C$69,COUNTIF(Agents!$A$2:$A${AG_END},a)>0),FALSE))),'
    'SORT(FILTER(HSTACK(dn,l,sn,id,vb),ok),3,TRUE,1,TRUE)),'
    '"No detractors in the period (if Medallia has data, check the columns in Setup C64:C68)")')
DB = f"B{D1}:B{D2}"
ws[f"F{D1}"] = f'=MAP({DB},LAMBDA(l,IF(l="","",IFERROR(VLOOKUP(l,Agents!$A$2:$B${AG_END},2,0),""))))'
ws[f"G{D1}"] = (f'=MAP(A{D1}:A{D2},{DB},LAMBDA(d,l,IF(l="","",'
                f'IFERROR(VLOOKUP(l&"|"&d,Shift_Log!$H$2:$O,8,0),"n/a"))))')
ws[f"H{D1}"] = (f'=MAP(D{D1}:D{D2},LAMBDA(i,IF(i="","",'
                f'IF(COUNTIF(Spotcheck_Log!$E$2:$E,"*"&i&"*")>0,"Yes","No"))))')
for r in range(D1, D2 + 1):
    ws[f"A{r}"].number_format = DATE
    ws[f"E{r}"].alignment = Alignment(wrap_text=True, vertical="top")
widths(ws, {"A": 12, "B": 24, "C": 10, "D": 16, "E": 70, "F": 20, "G": 12, "H": 12})
ws.freeze_panes = "A5"

# ---- Completamento spotcheck per TL (11) - nel foglio Spotcheck
ws = ws_spot
TLR1, TLR2 = 18, 47
header(ws, 17, ["Completion by TL", "Agents to check", "Cases to sample", "Cases logged",
                "% completed"], col=19)
SPB, SPO, SPP = f"$B$5:$B${E}", f"$O$5:$O${E}", f"$P$5:$P${E}"
ws[f"S{TLR1}"] = f'=IFERROR(SORT(UNIQUE(FILTER(Agents!B2:B{AG_END},Agents!A2:A{AG_END}<>"",Agents!B2:B{AG_END}<>""))),"")'
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
ws[f"S{TLR2 + 1}"] = "Cases logged = cases in Spotcheck_Log in the period, capped at the cases due per agent."
ws[f"S{TLR2 + 1}"].font = Font(italic=True, color="7F7F7F")
widths(ws, {"V": 11, "W": 13})

# ---- TL_View (8)
ws = ws_vtl
title(ws, "TEAM LEAD VIEW")
ws["A2"] = "Team Lead (empty = all):"
ws["A2"].font = Font(bold=True)
ws["B2"].fill = INPUT_FILL
ws["AB1"] = "TL list"
ws["AB2"] = f'=IFERROR(SORT(UNIQUE(FILTER(Agents!B2:B{AG_END},Agents!A2:A{AG_END}<>"",Agents!B2:B{AG_END}<>""))),"")'
v = DataValidation(type="list", formula1="=$AB$2:$AB$60", allow_blank=True)
v.add("B2")
ws.add_data_validation(v)
sel = lambda rng: f'((({rng}=$B$2)+($B$2=""))*({rng}<>""))>0'
ws["A3"] = (f'="Agents: "&COUNTIFS(Spotcheck!{SPB},IF($B$2="","?*",$B$2))&"   |   Cases to sample: "&'
            f'IF($B$2="",SUM(Spotcheck!{SPO}),SUMIFS(Spotcheck!{SPO},Spotcheck!{SPB},$B$2))&'
            f'"   |   Logged: "&SUM(MAP(Spotcheck!{SPB},Spotcheck!{SPO},Spotcheck!{SPP},LAMBDA(b,o,p,'
            f'IF(AND(ISNUMBER(o),IF($B$2="",b<>"",b=$B$2)),MIN(o,p),0))))')
ws["A3"].font = Font(bold=True, color="1F3864")
header(ws, 5, sp_hdr)
V1, V2 = 6, 6 + AGENT_ROWS - 1
ws[f"A{V1}"] = (f'=IFERROR(FILTER(Spotcheck!A5:Q{E},Spotcheck!A5:A{E}<>"",{sel(f"Spotcheck!B5:B{E}")}),'
                '"No agents for this TL")')
for c_ in "EGL":
    ws.column_dimensions[c_].number_format = PCT
ws.column_dimensions["I"].number_format = NUM1
for c_ in "FHKM":
    flag_cf(ws, f"{c_}{V1}:{c_}{V2}", f"{c_}{V1}")
ws.conditional_formatting.add(f"Q{V1}:Q{V2}", FormulaRule(formula=[f'LEFT(Q{V1},5)="To do"'], fill=RED))
ws.conditional_formatting.add(f"Q{V1}:Q{V2}", FormulaRule(formula=[f'Q{V1}="Done"'], fill=GREEN))
ws["S4"] = "NPS detractors to read (from Medallia)"
ws["S4"].font = SEC_FONT
header(ws, 5, ["Date", "LDAP", "Score", "Case ID", "Verbatim", "Team Lead", "Shift", "Already logged?"], col=19)
ws[f"S{V1}"] = (f'=IFERROR(FILTER(NPS_Detractors!A{D1}:H{D2},NPS_Detractors!B{D1}:B{D2}<>"",'
                f'{sel(f"NPS_Detractors!F{D1}:F{D2}")}),"No detractors")')
for r in range(V1, V2 + 1):
    ws[f"S{r}"].number_format = DATE
widths(ws, {"A": 24, "B": 20, "C": 9, "D": 9, "N": 10, "O": 11, "P": 11, "Q": 14, "R": 3,
            "S": 11, "T": 22, "U": 9, "V": 12, "W": 50, "X": 16, "Y": 9, "Z": 9, "AB": 20})
for c_ in "EFGHIJKLM":
    ws.column_dimensions[c_].width = 9
ws.freeze_panes = "B6"

# ---- Setup: controllo qualità dati (1)
ws = ws_set
ws["B50"] = "DATA QUALITY CHECK (updates automatically)"
ws["B50"].font = Font(bold=True, size=12, color="C00000")
header(ws, 51, ["Check", "Value", "Outcome", "Detail"], col=2)
TD = "Tableau!$I$2:$I"
dq = [
    ("Rows pasted in Tableau", "=COUNTA(Tableau!$A$2:$A)", 'IF(C{r}>0,"OK","WARNING")',
     '"Dynamic Slicing extraction by LDAP and Day"'),
    ("Dates covered by Tableau", f'=IFERROR(TEXT(MIN({TD}),"dd/mm/yyyy")&" - "&TEXT(MAX({TD}),"dd/mm/yyyy"),"-")',
     f'IF(AND(MIN({TD})<=$C$27,MAX({TD})>=$C$4),"OK","WARNING")',
     '"Must cover period + comparison period ("&TEXT($C$27,"dd/mm")&" - "&TEXT($C$4,"dd/mm")&")"'),
    ("Period days without Tableau data",
     f'=LET(d,SEQUENCE($C$4-$C$3+1,1,$C$3),ROWS(d)-SUM(MAP(d,LAMBDA(x,IF(COUNTIF({TD},x)>0,1,0)))))',
     'IF(C{r}=0,"OK","WARNING")',
     f'IFERROR(TEXTJOIN(", ",TRUE,MAP(FILTER(SEQUENCE($C$4-$C$3+1,1,$C$3),MAP(SEQUENCE($C$4-$C$3+1,1,$C$3),'
     f'LAMBDA(x,COUNTIF({TD},x)=0))),LAMBDA(x,TEXT(x,"dd/mm")))),"")'),
    ("Measures missing from the extraction (Setup C30:C38)",
     '=SUM(MAP($C$30:$C$38,LAMBDA(m,IF(COUNTIF(Tableau!$E$2:$E,m)=0,1,0))))',
     'IF(C{r}=0,"OK","WARNING")',
     'IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER($C$30:$C$38,MAP($C$30:$C$38,LAMBDA(m,COUNTIF(Tableau!$E$2:$E,m)=0))))),"")'),
    ("Agents in Tableau (in period) not in Shift_Log - ignored",
     f'=IFERROR(ROWS(FILTER(UNIQUE(FILTER(Tableau!$A$2:$A,{TD}>=$C$3,{TD}<=$C$4)),'
     f'MAP(UNIQUE(FILTER(Tableau!$A$2:$A,{TD}>=$C$3,{TD}<=$C$4)),LAMBDA(a,COUNTIF({LOGB},a)=0)))),0)',
     'IF(C{r}=0,"OK","INFO")',
     f'IFERROR(TEXTJOIN(", ",TRUE,FILTER(UNIQUE(FILTER(Tableau!$A$2:$A,{TD}>=$C$3,{TD}<=$C$4)),'
     f'MAP(UNIQUE(FILTER(Tableau!$A$2:$A,{TD}>=$C$3,{TD}<=$C$4)),LAMBDA(a,COUNTIF({LOGB},a)=0)))),"")'),
    ("Night agent-days in Shift_Log without Tableau rows", "=COUNTIFS(Shift_Log!$Y$2:$Y,1,Shift_Log!$Z$2:$Z,0)",
     'IF(C{r}=0,"OK","CHECK")',
     'IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER(Shift_Log!$B$2:$B,Shift_Log!$Y$2:$Y=1,Shift_Log!$Z$2:$Z=0))),"")'),
    ("Duplicate rows in Shift_Log (same LDAP and date)",
     '=COUNTIF(Shift_Log!$H$2:$H,"?*")-IFERROR(ROWS(UNIQUE(FILTER(Shift_Log!$H$2:$H,Shift_Log!$H$2:$H<>""))),0)',
     'IF(C{r}=0,"OK","WARNING")', '"Duplicate rows double the numbers: delete them from Shift_Log"'),
    ("TEST rows in Shift_Log", '=COUNTIF(Shift_Log!$G$2:$G,"TEST")', 'IF(C{r}=0,"OK","WARNING")',
     '"Dummy data: filter Source = TEST and delete them before real use"'),
    ("Period agents without Team Lead", f'=COUNTIFS(Agents!$A$2:$A${AG_END},"?*",Agents!$B$2:$B${AG_END},"")',
     'IF(C{r}=0,"OK","CHECK")', '"Needed for TL_View and completion by TL"'),
    ("Medallia columns found (Setup C64:C68)", "=" + "+".join(f"ISNUMBER({m})" for m in MED_MATCH) + "",
     'IF(COUNTA(Medallia!$A$2:$A)=0,"INFO",IF(C{r}=5,"OK","WARNING"))',
     '"Out of 5. INFO = Medallia sheet empty"'),
]
DQ1 = 52
for i, (lab, val, esito, det) in enumerate(dq):
    r = DQ1 + i
    ws.cell(r, 2, lab)
    ws.cell(r, 3, val).alignment = Alignment(horizontal="left")
    ws.cell(r, 4, "=" + esito.format(r=r))
    ws.cell(r, 5, "=" + det)
DQ2 = DQ1 + len(dq) - 1
for t, fill in (("WARNING", RED), ("CHECK", YELLOW), ("INFO", YELLOW), ("OK", GREEN)):
    ws.conditional_formatting.add(f"D{DQ1}:D{DQ2}", FormulaRule(formula=[f'D{DQ1}="{t}"'], fill=fill))
ws["B2"] = f"Data quality check: row {DQ1 - 2}   |   Medallia mapping: row 63"
ws["B2"].font = Font(italic=True, color="C00000")

# ---- Dashboard: stato qualità dati + sintesi automatica (13)
ws = ws_dash
DQCNT = f'COUNTIF(Setup!$D${DQ1}:$D${DQ2},"WARNING")'
ws["B3"] = (f'=IF({DQCNT}>0,"⚠ Data quality: "&{DQCNT}&" warnings - see Setup row {DQ1 - 2}",'
            '"Data quality: no blocking warnings")')
ws.conditional_formatting.add("B3", FormulaRule(formula=['LEFT(B3,1)="⚠"'], fill=RED))
ws["B22"] = "Automatic summary"
ws["B22"].font = SEC_FONT
lst = lambda cond, none: f'IFERROR(TEXTJOIN(", ",TRUE,FILTER($B$5:$B$10,{cond})),"{none}")'
better = lambda sign: (f'MAP($F$5:$F$10,Setup!$D$20:$D$25,LAMBDA(f,d,IF(ISNUMBER(f),'
                       f'IF(d="High",f{sign}0,f{"<" if sign == ">" else ">"}0),FALSE)))')
capped = (f'SUM(MAP(Spotcheck!{SPO},Spotcheck!{SPP},LAMBDA(o,p,IF(ISNUMBER(o),MIN(o,p),0))))')
IN_T, OUT_T = '$H$5:$H$10="On target"', '$H$5:$H$10="Off target"'
MIG, PEG = '$O$5:$O$10="Better"', '$O$5:$O$10="Worse"'
summary = [
    ('="Period "&TEXT(Setup!C3,"dd/mm")&"-"&TEXT(Setup!C4,"dd/mm")&": "&TEXT(C13,"#,##0")&" assignments in "&'
     'C16&" agent-days ("&C17&" agents), previous period "&TEXT(I13,"#,##0")&". "&Setup!C7&": "&'
     'TEXT(D18,"0.0")&" assignments per agent-day, "&Setup!C8&": "&TEXT(E18,"0.0")&"."'),
    f'="On target: "&{lst(IN_T, "no KPI")}&". Off target: "&{lst(OUT_T, "none")}&"."',
    (f'="Shift comparison: "&Setup!C7&" better than "&Setup!C8&" on "&{lst(better(">"), "no KPI")}&'
     f'"; worse on "&{lst(better("<"), "no KPI")}&"."'),
    (f'="Vs previous period, improving: "&{lst(MIG, "no KPI")}&'
     f'"; worsening: "&{lst(PEG, "no KPI")}&"."'),
    (f'="Spotcheck: "&COUNTIF(Agents!$Y$2:$Y${AG_END},">=3")&" agents with at least one RED KPI, "&'
     f'SUM(Agents!$Z$2:$Z${AG_END})&" cases to sample, completion "&'
     f'IFERROR(TEXT({capped}/SUM(Spotcheck!{SPO}),"0%"),"-")&". Agents RED for 2+ consecutive weeks: "&'
     f'COUNTIF(Red_Streaks!$G${S1}:$G${S2},">=2")&"."'),
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
    ("PURPOSE", "sec"),
    ("Analyses the performance of the R2 night shifts (18-3 and 21-6), together and split: "
     "Throughput, NPS, Handoff, Pending, Reopen, Recontact. The Spotcheck sheet tells each TL which agents "
     "and how many cases to check, with focus on Handoff, Pending, NPS and Reopen.", None),
    ("", None),
    ("WEEKLY PROCEDURE", "sec"),
    ("1. Tableau (Dynamic Slicing, breakdown by LDAP, column = Day): export the data and paste it into 'Tableau' "
     "from A1. Replace only A:G; columns H:J are formulas.", None),
    ("   Required measures (names in Setup C30:C38): assignments_completed, cases_solved_excluding_bot_only, "
     "pend_events, transfer_handoffs, case_reopen, case_solves, interaction_specialist_recontacts, "
     "nps_responses, nps_net_promoters.", None),
    ("2. Schedules: paste the new week's Schedules sheet into 'Schedules' from A1 (same columns).", None),
    ("3. 'New_Week' calculates itself: copy A2:G and paste VALUES ONLY at the bottom of 'Shift_Log' "
     "(Shift_Log is the history: do not delete previous weeks).", None),
    ("4. Medallia: paste the NPS responses export into 'Medallia' from A1 (check the column names once "
     "in Setup C64:C68).", None),
    ("5. In 'Setup' set the start / end date (usually one week, Sun-Sat): the comparison uses the previous "
     "period of the same length. Check the 'DATA QUALITY CHECK' box in Setup (row 50).", None),
    ("6. Read 'Dashboard' and 'Spotcheck'; each TL uses 'TL_View' and logs the checks in 'Spotcheck_Log'.", None),
    ("", None),
    ("ANALYSIS SHEETS", "sec"),
    ("Dashboard: KPIs total / 18-3 / 21-6, comparison with the previous period (Δ and Better/Worse), volumes, "
     "automatic text summary, 8-week trend. The data quality status is at the top.", None),
    ("Spotcheck: priority by agent + completion by TL (cases due vs logged) + root causes.", None),
    ("TL_View: pick the TL in B2 to see their agents to check, the status and the NPS detractors to read.", None),
    ("Rotation: for agents working both shifts, KPIs on their 18-3 days vs their 21-6 days "
     "(same agent = cleaner comparison between shifts).", None),
    ("Workload: assignments per agent-day by weekday, by week, by date and by agent.", None),
    ("Red_Streaks: for each agent and KPI (Handoff, Pending, NPS, Reopen), how many weeks in a row it is RED; "
     "2+ weeks = structural problem, not an isolated case.", None),
    ("Coaching: for each logged coaching (Coaching done? = Yes), the KPI in the 14 days before and the 14 after.", None),
    ("NPS_Detractors: Medallia 0-6 responses for night agents in the period, with verbatims: use them to "
     "pick the NPS cases for the spotcheck.", None),
    ("", None),
    ("HOW THE SHIFT IS ASSIGNED", "sec"),
    ("Tableau only has the day, while shifts cross midnight. An agent's Tableau day D is assigned to the shift "
     "that works the most hours on it: the shift started on day D (18-3 = 6h, 21-6 = 3h) or the one started on "
     "day D-1 (18-3 = 3h, 21-6 = 6h). On a tie, the day-D shift wins. If the main shift is a day shift, the day "
     "is 'Other' and is excluded.", None),
    ("Special Assignments included. Tiers excluded in Setup (currently: Premium Support).", None),
    ("", None),
    ("KPI DEFINITIONS (Data Dictionary - Metrics)", "sec"),
    ("Throughput = cs_cases_solved_excluding_bot_only / cs_ambassador_case_assignments_completed", None),
    ("NPS = 100 × cs_customer_nps_net_promoters / cs_customer_nps_responses", None),
    ("Handoff = cs_ambassador_transfer_handoffs / cs_ambassador_case_assignments_completed", None),
    ("Pending = cs_ambassador_pend_events / cs_ambassador_case_assignments_completed", None),
    ("Reopen = cs_case_reopen / cs_ambassador_case_solves", None),
    ("Recontact = cs_interaction_specialist_recontacts / cs_ambassador_case_assignments_completed", None),
    ("KPIs are always calculated as sum of numerators / sum of denominators (never an average of percentages).", None),
    ("Targets: EMEA 2026 H2 CS Delivery Targets, Resolutions 2 (NPS: Resolutions 2 Italian). "
     "Editable in Setup.", None),
    ("", None),
    ("SPOTCHECK: HOW IT WORKS", "sec"),
    ("Traffic light per agent on Handoff, Pending, NPS, Reopen: RED = off target and also beyond the night average "
     "by k standard deviations (true outlier); YELLOW = off target only; GREEN = on target; n.s. = volume too "
     "low to judge (thresholds in Setup).", None),
    ("Score = 3 × RED + YELLOW: the list is sorted by priority. "
     "Cases to sample = 3 per RED + 1 per YELLOW (editable).", None),
    ("Every checked case must be logged in Spotcheck_Log: the 'Status' column of the Spotcheck sheet shows "
     "what is left to do; the root cause summary is on the right.", None),
    ("", None),
    ("WHAT TO CHECK IN THE CASES", "sec"),
    ("Handoff: was it avoidable? Does it happen at the end of the shift (03:00 / 06:00)? This is the typical night "
     "risk: cases passed to the next shift instead of being solved or pended correctly.", None),
    ("Pending: was the pend needed? Is the next step clear for the user and for whoever picks up the case? "
     "Does it repeat on the same case? (also the starting point to measure 'unresponsive users').", None),
    ("NPS: read the detractor verbatims (0-6) and separate agent fault / policy / product.", None),
    ("Reopen: why was the case reopened? Incomplete solution, unclear communication, premature closure "
     "(e.g. closed at the end of the shift without user confirmation).", None),
    ("Other ideas: weekly calibration between TLs on 2-3 shared cases; 18-3 vs 21-6 comparison on the "
     "21-03 overlap; 4-week trend after coaching to see whether the KPI improves.", None),
    ("", None),
    ("OPEN NOTES", "sec"),
    ("- Reopen: cs_ambassador_case_solves is not in the extraction yet. For now the denominator is "
     "cs_cases_solved_excluding_bot_only: as soon as it is available, change Setup C35.", None),
    ("- Unresponsive users: on hold, no field available. The 'User not responding' root cause in "
     "Spotcheck_Log already lets you start counting them.", None),
    ("- Limits: formulas read up to 5000 rows of Shift_Log (about 35-40 weeks) and 40000 rows of "
     "Tableau. Beyond that, archive old weeks in another file.", None),
    ("- TEST DATA: Shift_Log contains rows with Source = TEST (dummy shifts assigned to the agents "
     "of the 13-24/09 test extraction) to try the formulas. Before real use, filter Source = TEST, "
     "delete those rows and replace the data in 'Tableau'. The 'Medallia' rows and the 2 TEST rows in "
     "'Spotcheck_Log' are test data too.", None),
    ("- Medallia: the export structure is not verified yet. If the column names differ, "
     "fix them in Setup C64:C68 (the data quality box shows how many columns are found).", None),
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
    cell.value = "end of formula area"
    cell.font = Font(italic=True, color="BFBFBF")
for sh in wb.worksheets:
    for dv in sh.data_validations.dataValidation:
        dv.formula1 = close_ranges(dv.formula1, sh.title)

wb.save(OUT)
with open("reference.json", "w") as fh:
    json.dump(reference, fh, indent=1, default=str)
print(json.dumps(reference, indent=1, default=str))
print("tableau rows", len(tab), "sched rows", len(sched_sample), "log rows", len(log_rows))
