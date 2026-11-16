"""
Round 2: corrects the price-fluctuation sheets of VALIDATION_of_the_Developed_Model.xlsx
  1. PF 9: factor 0.966 -> 0.96 (typo)
  2. PF 3 / PF 8 / PF 9: titles carried the wrong valuation number
  3. PF 8: stale values copied from PF 7 (previous cost of material at site, dates)
  4. Corrections Log section B: wrong Grand Summary cells (my earlier formula pointed at J3/J4; correct cells J7/J9)
  5. New sheet 'PF Model Check': live comparison of the PF sheets with the model (Px of the valued IPC, V-Vna from the model)
  6. Corrections Log: round-2 entries
Judgement calls (Px alignment, non-adjustable series, PF 8 current material at site) are NOT changed – they are shown in 'PF Model Check'.
"""
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

SRC = "/mnt/user-data/uploads/1794743347573_VALIDATION_of_the_Developed_Model.xlsx"
OUT = "/mnt/user-data/outputs/validation_pf/VALIDATION_of_the_Developed_Model_CORRECTED.xlsx"
wb = load_workbook(SRC); wv = load_workbook(SRC, data_only=True)
B = Font(bold=True); HF = PatternFill("solid", fgColor="D9E1F2"); YEL = PatternFill("solid", fgColor="FFF2CC")
PFS = list(range(2, 10))

def rows_of(ws):
    d = {}
    occ = {"cumv": [], "mat": [], "cum_na": [], "date": []}
    for r in range(36, ws.max_row + 1):
        t = str(ws.cell(r, 2).value or "")
        if t == "Date": occ["date"].append(r)
        elif t.startswith("Cumulative value"): occ["cumv"].append(r)
        elif t.startswith("Cost of material"): occ["mat"].append(r)
        elif t.startswith("Cumulative cost"): occ["cum_na"].append(r)
        elif t.startswith("V - Vna ="): d["vvna"] = r
        elif t.startswith("0.96(V - Vna)"): d["factor"] = r
        elif t.startswith("Price Adjustment for"): d["pa"] = r
        elif t.startswith("Price adjustment computation"): d["title"] = r
    d.update({k: v for k, v in occ.items()}); return d
R = {n: rows_of(wb[f"Price Fluctuation {n}"]) for n in PFS}
for n in PFS: assert all(len(R[n][k]) == 2 for k in ("cumv", "mat", "cum_na", "date")) and "pa" in R[n], n

changes = []
def setv(ws, addr, new, why):
    old = ws[addr].value; ws[addr].value = new; ws[addr].fill = YEL
    changes.append((ws.title, addr, old, new, why))

# 1. PF 9 factor
ws9 = wb["Price Fluctuation 9"]; fa = R[9]["factor"]
assert ws9.cell(fa, 4).value == "=0.966*(D56)/100"
setv(ws9, f"D{fa}", "=0.96*(D56)/100", "typo 0.966 -> 0.96 (all other sheets and the row label use 0.96)")
# 2. titles
for n, wrong in ((3, "No 02"), (8, "No 07"), (9, "No 07")):
    ws = wb[f"Price Fluctuation {n}"]; t = R[n]["title"]; old = ws.cell(t, 2).value
    assert old.endswith(wrong), old
    setv(ws, f"B{t}", old[:-len(wrong)] + f"No {n:02d}", f"title said valuation {wrong}")
# 3. PF 8 stale cells (copied from PF 7)
ws8, ws7 = wb["Price Fluctuation 8"], wb["Price Fluctuation 7"]; ws9v = wv["Price Fluctuation 9"]
r8, r7 = R[8], R[7]
assert ws8.cell(r8["mat"][0], 4).value == wv["Price Fluctuation 7"].cell(r7["mat"][0], 4).value   # = PF 7 *previous*
setv(ws8, f"D{r8['mat'][0]}", wv["Price Fluctuation 7"].cell(r7["mat"][1], 4).value, "previous valuation must equal PF 7 current (was PF 7 previous value)")
setv(ws8, f"D{r8['date'][0]}", wv["Price Fluctuation 7"].cell(r7["date"][1], 4).value, "previous date must equal PF 7 current date")
setv(ws8, f"D{r8['date'][1]}", ws9v.cell(R[9]["date"][0], 4).value, "current date duplicated PF 7; PF 9 previous date used (chain)")

# 4. log section B
lg = wb["Corrections Log"]
assert str(lg["C25"].value).startswith("='BOQ Grand Summary'!J3")
setv(lg, "C25", "='BOQ Grand Summary'!J7+'BOQ Grand Summary'!J9", "formula pointed at empty cells (J3/J4); correct cells are J7 and J9")

# 5. PF Model Check sheet
mc = wb.create_sheet("PF Model Check")
mc["A1"] = "PRICE FLUCTUATION – model check (live formulas)"; mc["A1"].font = Font(bold=True, size=12)
mc["A2"] = ("Compares each Price Fluctuation sheet with the model. Px: composition of the IPC whose work is being valued. V - Vna: adjustable work of that IPC from the model "
            "('18 IPC Treatment Summary') plus the change in material at site. PF No. N values the work between IPC N and IPC N+1 (cumulative values of the sheet).")
mc["A2"].alignment = Alignment(wrap_text=True, vertical="top"); mc.merge_cells("A2:K2"); mc.row_dimensions[2].height = 48
mc["A4"] = "Px IPC offset  (0 = Px of IPC N, as in the PF sheets;  1 = Px of IPC N+1, the IPC being valued)"; mc["A4"].font = B
mc["I4"] = 1; mc["I4"].fill = YEL; mc["I4"].font = B; mc["I4"].alignment = Alignment(horizontal="center")
mc["A6"] = "PF No."; mc["A6"].font = B
mc["A7"] = "IPC whose Px is used"; mc["A7"].font = B
for j, n in enumerate(PFS):
    c = 3 + j; mc.cell(6, c, n).font = B; mc.cell(6, c).fill = HF; mc.cell(6, c).alignment = Alignment(horizontal="center")
    mc.cell(7, c, f"={L(c)}6+$I$4").alignment = Alignment(horizontal="center")
codes = [(wv["Price Fluctuation 2"].cell(r, 1).value, wv["Price Fluctuation 2"].cell(r, 2).value) for r in range(2, 36)]
mc["A9"] = "A. Index change (Ixc − Ixb) / Ixb"; mc["A9"].font = B
RA0 = 10
for i, (code, name) in enumerate(codes):
    r = RA0 + i; mc.cell(r, 1, code); mc.cell(r, 2, name)
    for j, n in enumerate(PFS):
        mc.cell(r, 3 + j, f"='Price Fluctuation {n}'!E{2+i}/'Price Fluctuation {n}'!D{2+i}-1").number_format = "0.0000"
RA1 = RA0 + len(codes) - 1
RB_T = RA1 + 2; mc.cell(RB_T, 1, "B. Px (%) of the selected IPC – from '21 Dynamic Px 96%'").font = B
RB0 = RB_T + 1
for i, (code, name) in enumerate(codes):
    r = RB0 + i; mc.cell(r, 1, code); mc.cell(r, 2, name); col = L(4 + i)
    for j in range(len(PFS)):
        mc.cell(r, 3 + j, f"=INDEX('21 Dynamic Px 96%'!{col}$2:{col}$10,MATCH({L(3+j)}$7,'21 Dynamic Px 96%'!$A$2:$A$10,0))").number_format = "0.0000"
RB1 = RB0 + len(codes) - 1
RC_T = RB1 + 2; mc.cell(RC_T, 1, "C. Result").font = B
labels = ["Valued period", "Σ Px(Ixc−Ixb)/Ixb  – as in PF sheet", "Σ Px(Ixc−Ixb)/Ixb  – with Px selected above",
          "V − Vna – as in PF sheet (Rs.)", "V − Vna – from model (Rs.)", "Price adjustment – as in PF sheet (Rs.)",
          "Price adjustment – model-driven (Rs.)", "Difference: model-driven − PF sheet (Rs.)"]
RC0 = RC_T + 1
for k, t in enumerate(labels):
    mc.cell(RC0 + k, 1, t).font = B if k >= 5 else Font()
for j, n in enumerate(PFS):
    c = 3 + j; cl = L(c); s = f"'Price Fluctuation {n}'"; rr = R[n]
    mc.cell(RC0, c, f'="IPC "&{cl}6&" → "&({cl}6+1)').alignment = Alignment(horizontal="center")
    mc.cell(RC0 + 1, c, f"={s}!F36").number_format = "0.0000"
    mc.cell(RC0 + 2, c, f"=SUMPRODUCT({cl}{RB0}:{cl}{RB1},{cl}{RA0}:{cl}{RA1})").number_format = "0.0000"
    mc.cell(RC0 + 3, c, f"={s}!D{rr['vvna']}").number_format = "#,##0"
    mc.cell(RC0 + 4, c, f"=INDEX('18 IPC Treatment Summary'!$E$2:$E$10,MATCH({cl}6+1,'18 IPC Treatment Summary'!$A$2:$A$10,0))+({s}!D{rr['mat'][1]}-{s}!D{rr['mat'][0]})").number_format = "#,##0"
    mc.cell(RC0 + 5, c, f"={s}!D{rr['pa']}").number_format = "#,##0"
    mc.cell(RC0 + 6, c, f"=0.96*{cl}{RC0+4}/100*{cl}{RC0+2}").number_format = "#,##0"
    mc.cell(RC0 + 7, c, f"={cl}{RC0+6}-{cl}{RC0+5}").number_format = "#,##0"
mc.cell(RC0 - 1, 11, "Total").font = B
for k in (5, 6, 7): mc.cell(RC0 + k, 11, f"=SUM(C{RC0+k}:J{RC0+k})").number_format = "#,##0"; mc.cell(RC0 + k, 11).font = B
for k in (5, 6, 7):
    for c in range(1, 12): mc.cell(RC0 + k, c).fill = HF
rn = RC0 + 9
notes = ["Reading the result:",
         "• 'V − Vna as in PF sheet' relies on the cumulative non-adjustable figure (item d). That series is not cumulative (it falls within PF 3, 2.07 M → 0.71 M, and within PF 8, 9.09 M → 0.32 M), so in PF 8 V − Vna (12.77 M) exceeds the work valued in that period (4.00 M).",
         "• Px of IPC N+1 (offset 1) is the composition of the work actually valued by PF N; the PF sheets use IPC N (offset 0). Set cell I4 to 0 to reproduce the PF sheets' Px.",
         "• Σ Px = 96 and the factor 0.96 are both applied (effective 92.2 %). Confirm that this is intended."]
for k, t in enumerate(notes):
    mc.cell(rn + k, 1, t).font = B if k == 0 else Font()
mc.column_dimensions["A"].width = 46; mc.column_dimensions["B"].width = 34
for c in range(3, 12): mc.column_dimensions[L(c)].width = 15
mc.freeze_panes = "C8"

# 6. log round 2
r = lg.max_row + 2
def put(rr, c, v, bold=False, wrap=False, fill=None):
    x = lg.cell(rr, c, v)
    if bold: x.font = B
    if wrap: x.alignment = Alignment(wrap_text=True, vertical="top")
    if fill: x.fill = fill
put(r, 1, "ROUND 2 – price fluctuation sheets (Price Fluctuation 2–9)", True); r += 1
items = [
 ("7", "Corrected (typos / stale cells)",
  "PF 9: factor 0.966 corrected to 0.96. PF 3, PF 8, PF 9: titles carried the wrong valuation number (02, 07, 07). PF 8: previous cost of material at site (754,860.86), previous date and current date were stale copies from PF 7 "
  "and now equal the PF 7 current values / PF 9 previous date. Corrections Log section B pointed at empty cells; fixed. Changed cells are highlighted yellow."),
 ("8", "To confirm – Px alignment",
  "PF N values the work between IPC N and IPC N+1 but uses the Px of IPC N (the previous period). The Px of IPC N+1 is the composition of the work being valued. Shown in 'PF Model Check' (cell I4 switches); not changed."),
 ("9", "To confirm – non-adjustable series (item d)",
  "The 'cumulative cost of non-adjustable elements' is taken from sheet 22 column U (period Vna after discount + variations), which is not cumulative: it falls from 2.07 M to 0.71 M (PF 3) and from 9.09 M to 0.32 M (PF 8). "
  "Consequently V − Vna in PF 8 (12.77 M) is larger than the work valued in that period (4.00 M). Variations are also deducted although V (model C.W.D) does not contain them. Not changed."),
 ("10", "To confirm – PF 8 current material at site",
  "PF 8 current = 176,800 (same as PF 7 current) while PF 9 previous = 380,800; every other valuation chains exactly. 380,800 is probably correct (+204,000 on V − Vna). Not changed."),
 ("11", "To confirm – Σ Px = 96 and factor 0.96",
  "Px are scaled to sum to 96 in sheet 21 and the formula multiplies by 0.96 again (effective 92.2 % of the mix). If the factor is meant to be the only reduction, Px should sum to 100."),
]
for n_, t, d in items:
    put(r, 1, n_, True); put(r, 2, t, True); put(r, 3, d, wrap=True); lg.row_dimensions[r].height = 80; r += 1
wb.save(OUT)
print("saved", OUT)
for c in changes: print(c)
