# -*- coding: utf-8 -*-
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter as gcl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule, CellIsRule
import os, json, datetime as _dt
from data import TRANSIT, RULES, DOCS, SOURCES, INBOX, TODAY

_LJ = os.environ.get("LISTINGS_JSON")
if _LJ:
    OPPS = json.load(open(_LJ))
    for _o in OPPS:                      # JSON dates -> real dates
        for _k in ("deadline", "pub"):
            if isinstance(_o.get(_k), str) and len(_o[_k]) == 10:
                _o[_k] = _dt.date.fromisoformat(_o[_k])
else:
    from data import OPPS

OUT = os.environ.get("WORKBOOK_OUT",
      "/Users/johnpaulroche/Desktop/Apartments/NYC_Affordable_Housing_System.xlsx")

FONT = "Arial"
NAVY   = "1F3864"
BLUE   = "2E5C8A"
LIGHT  = "DCE6F1"
BAND   = "F2F6FB"
GREEN  = "C6EFCE"; GREENF = "006100"
RED    = "FFC7CE"; REDF   = "9C0006"
AMBER  = "FFEB9C"; AMBERF = "9C6500"
GREY   = "808080"
INPUTC = "0000FF"   # blue = editable input
LINKC  = "0563C1"

thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()

def F(sz=10, b=False, color="000000", it=False):
    return Font(name=FONT, size=sz, bold=b, color=color, italic=it)

def title(ws, text, sub=None, span=12):
    ws["A1"] = text
    ws["A1"].font = Font(name=FONT, size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 30
    for c in range(1, span + 1):
        ws.cell(row=1, column=c).fill = PatternFill("solid", fgColor=NAVY)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = F(9, it=True, color="404040")
        ws.row_dimensions[2].height = 26
        ws["A2"].alignment = Alignment(vertical="center", wrap_text=False)

def hdr(ws, row, headers, widths=None, height=42):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BLUE)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BOX
    ws.row_dimensions[row].height = height
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[gcl(i)].width = w

def sect(ws, row, text, span=3):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
    for i in range(1, span + 1):
        ws.cell(row=row, column=i).fill = PatternFill("solid", fgColor=BLUE)
    ws.row_dimensions[row].height = 20

def link(cell, url, text=None):
    if not url:
        return
    cell.value = text or url
    cell.hyperlink = url
    cell.font = Font(name=FONT, size=9, color=LINKC, underline="single")

MONEY  = '$#,##0;($#,##0);"-"'
MONEY2 = '$#,##0.00;($#,##0.00);"-"'
PCT    = '0.0%'
DATE   = 'yyyy-mm-dd'

# =============================================================== 2. INPUTS
ws = wb.active
ws.title = "INPUTS"
title(ws, "MASTER INPUTS  —  edit the blue cells only",
      "Everything in this workbook recalculates from this sheet. Blue = you edit it. Black = calculated, do not type over. "
      "Change salary, household size, rent limits, commute limits or geography and every other sheet updates.", span=4)
ws.column_dimensions["A"].width = 46
ws.column_dimensions["B"].width = 26
ws.column_dimensions["C"].width = 92

IN = {}  # label -> row

def irow(r, label, value, note, fmt=None, editable=True, formula=False):
    ws.cell(row=r, column=1, value=label).font = F(10, b=not editable)
    c = ws.cell(row=r, column=2, value=value)
    c.font = F(11, b=True, color=INPUTC if editable else "000000")
    c.border = BOX
    if editable:
        c.fill = PatternFill("solid", fgColor="FFF2CC")
    else:
        c.fill = PatternFill("solid", fgColor="EDEDED")
    if fmt:
        c.number_format = fmt
    c.alignment = Alignment(horizontal="center")
    n = ws.cell(row=r, column=3, value=note)
    n.font = F(9, color="404040")
    n.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 26
    return r

sect(ws, 4, "FINANCIAL  —  incomes are ANNUAL and GROSS (before tax and before any deductions)")
R_SAL   = irow(5,  "Applicant annual salary (NYC DOE base)", 72000,
   "Gross salary before tax. HPD annualises the GROSS line on your paystubs, not take-home pay.", MONEY)
R_DOE2  = irow(6,  "Applicant additional DOE income", 0,
   "Per-session, coverage, summer school, F-status work. This IS counted — include it or eligibility will shift later.", MONEY)
R_OTH   = irow(7,  "Applicant other income", 0,
   "Any other recurring income. One-off gifts and inheritances are excluded (Handbook 5-4.H) but still disclose them.", MONEY)
R_M2    = irow(8,  "Household member 2 — annual gross income", 0,
   "Brother #2. Only counted when Household size is 2 or more. Siblings qualify as one household automatically.", MONEY)
R_M3    = irow(9,  "Household member 3 — annual gross income", 0,
   "Brother #3. Only counted when Household size is 3.", MONEY)
R_ADJ   = irow(10, "Income adjustments (enter as negative)", 0,
   "WARNING: HPD uses gross income BEFORE any expenses or deductions (Handbook 5-4.A(3)). Travel costs, pre-tax "
   "commuter benefits, 403(b) and taxes do NOT reduce countable income. This cell exists only to model 'what if' — "
   "do not rely on it for a real application.", MONEY)
R_INC   = irow(11, "TOTAL COUNTABLE HOUSEHOLD INCOME",
   f"=B{5}+B{6}+B{7}+IF($B$13>=2,B{8},0)+IF($B$13>=3,B{9},0)+B{10}",
   "This single number drives every eligibility test in the workbook.", MONEY, editable=False)
R_GMI   = irow(12, "Gross monthly household income", "=B11/12",
   "Used for every rent-burden calculation.", MONEY2, editable=False)
R_HH    = irow(13, "Household size  (1, 2 or 3)", 1,
   "Set to 1 = applying alone (your verified situation today). Set to 2 or 3 to apply with your brothers — siblings are "
   "Immediate Family Members and qualify as one household automatically. Occupancy (Handbook 5-7): studio 1-2 people, "
   "1BR 1-3, 2BR 2-5, 3BR 3-7. Changing this switches every listing to that household size's income cap, so enter the "
   "brothers' incomes above at the same time.", None)
R_AS1   = irow(14, "Applicant assets (savings/checking/brokerage)", 0,
   "EXCLUDE pension (NYC TRS), 403(b)/457, life insurance, annuities and college savings — Handbook 5-5.A(4).", MONEY)
R_AS2   = irow(15, "Household member 2 assets", 0, "Counted when household size is 2 or more.", MONEY)
R_AS3   = irow(16, "Household member 3 assets", 0, "Counted when household size is 3.", MONEY)
R_AST   = irow(17, "TOTAL COUNTABLE HOUSEHOLD ASSETS", "=B14+IF($B$13>=2,B15,0)+IF($B$13>=3,B16,0)",
   "Tested against ONE limit for the whole household — the four-person HUD income limit at the unit's AMI.", MONEY, editable=False)
R_BURD  = irow(18, "Target rent burden %", 0.30,
   "The share of gross monthly income you want housing to cost. 30% is the conventional benchmark.", PCT)
R_TGT   = irow(19, "Target monthly housing cost", "=B12*B18",
   "Gross monthly income x target burden. This is your affordability yardstick.", MONEY2, editable=False)
R_PMAX  = irow(20, "Preferred maximum rent", 1800,
   "Soft ceiling. Used for the 'rent saved vs your maximum' column.", MONEY)
R_AMAX  = irow(21, "Absolute maximum rent", 2100,
   "Hard ceiling. Anything above this is classified ABOVE ABSOLUTE MAXIMUM and never reaches APPLY NOW.", MONEY)
R_UTIL  = irow(22, "Default estimated utilities (per month)", 85,
   "Used where a listing does not include utilities. Electric-heat buildings are far higher — override per listing.", MONEY)
R_BORD  = irow(23, "'Borderline' upper burden %", 0.35, "Above target but still tolerable.", PCT)
R_VAFF  = irow(24, "'Very affordable' burden %", 0.25, "At or below this = VERY AFFORDABLE.", PCT)

sect(ws, 26, "HOUSING PREFERENCES")
R_STU  = irow(27, "Studio acceptable?", "YES", "Studio holds 1-2 people.")
R_1BR  = irow(28, "1-Bedroom acceptable?", "YES", "1BR holds 1-3 people — works for two or three brothers.")
R_2BR  = irow(29, "2-Bedroom acceptable?", "YES", "2BR requires a minimum of 2 people. Set NO if applying alone.")
R_MAN  = irow(30, "Manhattan allowed?", "YES", "Primary target.")
R_B25  = irow(31, "Brooklyn — direct 2/5 corridor allowed?", "YES", "Direct ride to Freeman St / 174 St with no transfer.")
R_B34  = irow(32, "Brooklyn — 3/4 lines allowed?", "YES", "One cross-platform transfer to the 2/5 at Franklin Av.")
R_BOT  = irow(33, "Brooklyn — other lines allowed?", "NO", "A/C, G, J/M/Z, L. Set YES to surface Atlantic Chestnut (cheapest eligible units found).")
R_BX   = irow(34, "Bronx allowed?", "NO", "Excluded at your instruction. Work destination only.")
R_QSI  = irow(35, "Queens / Staten Island allowed?", "NO", "")
R_ISL  = irow(36, "Must be physically on Manhattan island?", "YES", "Excludes Roosevelt Island, which is in Manhattan borough but not on the island.")
R_SIDE = irow(37, "Preferred side of Manhattan", "West Side", "Note: the 2 runs down the WEST side (7 Av) but through CENTRAL Harlem on Lenox Ave.")
R_HOODS= irow(38, "Preferred neighbourhoods", "Harlem, UWS, Hell's Kitchen, Chelsea, Morningside Hts",
   "Harlem gives by far the best commute: 135 St is 14 minutes to Freeman St on a direct 2.")

sect(ws, 40, "TRANSIT")
R_WORK = irow(41, "Work address / area", "Crotona Park East, Bronx, NY 10460", "")
R_DEST1= irow(42, "Primary work station", "Freeman St (2/5)", "Southern Blvd at Freeman St, on the eastern edge of Crotona Park East.")
R_DEST2= irow(43, "Secondary work station", "174 St (2/5)", "CAUTION: '174 St' (2/5) is NOT '174-175 Sts' (B/D), which is a mile away on the Grand Concourse.")
R_WALKD= irow(44, "Walk from work station to work (min)", 7, "Adjust once you know the exact school address.")
R_WALKM= irow(45, "Maximum acceptable walk to home station (min)", 12, "")
R_CPREF= irow(46, "Preferred maximum commute (min)", 50, "Door to door, one way.")
R_CABS = irow(47, "Absolute maximum commute (min)", 70, "Above this a listing is marked NO - too long.")
R_WAIT = irow(48, "Average wait at origin station (min)", 4, "The 2 and 5 together give roughly 14 trains/hour into 174 St in the AM.")
R_TPEN = irow(49, "Minutes added per transfer", 4, "Raise this if you dislike transfers; it re-grades every listing.")
R_DIRP = irow(50, "Direct subway strongly preferred?", "YES", "")

sect(ws, 52, "EMPLOYMENT & STATUS")
R_DOE  = irow(53, "NYC DOE employee?", "YES", "")
R_YRS  = irow(54, "Years employed (DOE)", 2, "Second-year teacher. Salary steps matter — model them on SCENARIOS.")
R_MUNI = irow(55, "Municipal employee (paid by City of New York)?", "YES",
   "QUALIFIES for the 10% Municipal Employee / Military Veteran preference. Proof = paystubs issued by the City of New York.")
R_VET  = irow(56, "Military veteran (any household member)?", "NO",
   "If YES for any brother, answer YES to the Housing Connect veteran question and keep a DD-214.")
R_RES  = irow(57, "Current residence", "Westchester County, NY", "")
R_NYC  = irow(58, "NYC resident?", "NO",
   "COSTS YOU. NYC residents are processed before non-residents — including inside the municipal-employee preference pool.")

sect(ws, 60, "SYSTEM")
R_TODAY= irow(61, "Today's date", "=TODAY()", "Drives every deadline countdown.", DATE, editable=False)
R_REV  = irow(62, "Last full source review", TODAY, "Update whenever you work through SOURCES_TO_MONITOR.", DATE)

yn = DataValidation(type="list", formula1='"YES,NO"', allow_blank=True)
ws.add_data_validation(yn)
for r in [R_STU,R_1BR,R_2BR,R_MAN,R_B25,R_B34,R_BOT,R_BX,R_QSI,R_ISL,R_DIRP,R_DOE,R_MUNI,R_VET,R_NYC]:
    yn.add(ws.cell(row=r, column=2))
hhv = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
ws.add_data_validation(hhv); hhv.add(ws.cell(row=R_HH, column=2))
ws.freeze_panes = "A5"

# shorthand refs
I = lambda r: f"INPUTS!$B${r}"
INC, GMI, HH, AST = I(R_INC), I(R_GMI), I(R_HH), I(R_AST)
BURD, PMAX, AMAX, BORD, VAFF = I(R_BURD), I(R_PMAX), I(R_AMAX), I(R_BORD), I(R_VAFF)

# ================================================== 3. ACTIVE_OPPORTUNITIES
COLS = [
 ("id","Listing ID",12),("name","Project name",30),("addr","Address",34),("hood","Neighbourhood",24),
 ("boro","Borough",11),("side","Manhattan side",13),("cb","Community Board",16),
 ("url","Official advertisement",26),("appurl","Direct application link",24),("program","Program",34),
 ("pub","Published",12),("deadline","Deadline",12),("days","Days remaining",11),("open","Open / Closed",14),
 ("unit","Unit type",13),("beds","Bedrooms",9),("ami","AMI tier",14),
 ("minimum","Minimum income",14),("max1","Max income 1 person",14),("max2","Max income 2 people",14),
 ("max3","Max income 3 people",14),("maxapp","Applicable max (by HH size)",15),
 ("inc","Your household income",15),("elig","INCOME ELIGIBLE?",24),
 ("above","$ above minimum",14),("below","$ below maximum",14),
 ("rent","Monthly rent",13),("utilinc","Utilities included",30),("util","Est. utilities",11),
 ("hcost","Est. total housing cost",14),("gmi","Gross monthly income",14),
 ("rburd","Rent % of gross",12),("hburd","Housing cost % of gross",13),
 ("left","$ left after housing",14),("afford","AFFORDABILITY",24),
 ("save","Rent saved vs your max",14),("excep","Exceptional value?",13),
 ("asset","Asset limit",13),("aselig","Asset eligible?",13),
 ("mepref","Municipal employee preference",16),("prefpct","Preference %",11),
 ("station","Nearest useful station",26),("lines","Subway lines",18),("walk","Walk to subway (min)",11),
 ("route","Best work-bound route",56),("direct","Direct ride?",10),("transfers","Transfers",9),
 ("ttype","Transfer type",18),("tdetail","Transfer detail",46),("dest","Destination station",16),
 ("ride","Train time (min)",11),("dd","Door-to-door (min)",12),("grade","TRANSIT GRADE",12),
 ("cok","Commute acceptable?",20),
 ("west","West Side?",10),("corridor","Transit corridor",20),("geog","Geography allowed?",13),
 ("laundry","Laundry",9),("elevator","Elevator",9),("amen","Amenities",46),
 ("pass","PASSES ALL 3 TESTS",16),
 ("status","Application status",20),("applied","Date applied",12),("conf","Confirmation number",16),
 ("log","Log number",13),("verified","Last verified",12),("notes","Notes",80),
 ("qual","Qualifies now",10),("score","Sort score",12),("rank","Apply rank",10),
]
KEY = {k: gcl(i + 1) for i, (k, _, _) in enumerate(COLS)}
def C(k, r): return f"${KEY[k]}{r}"

wa = wb.create_sheet("ACTIVE_OPPORTUNITIES")
title(wa, "ACTIVE OPPORTUNITIES  —  master database",
      "One row per unit-type / AMI-band combination. Every yellow-ish calculated column recalculates from INPUTS. "
      "Published income figures are transcribed from the official advertisement linked in each row — never inferred from the AMI label.",
      span=len(COLS))
HROW = 3
hdr(wa, HROW, [h for _, h, _ in COLS], [w for _, _, w in COLS])
r0 = HROW + 1
n = len(OPPS)
rn = r0 + n - 1

for i, o in enumerate(OPPS):
    r = r0 + i
    def put(k, v, fmt=None, al=None, wrap=False):
        c = wa.cell(row=r, column=list(KEY).index(k) + 1, value=v)
        c.font = F(9)
        c.border = BOX
        c.alignment = Alignment(wrap_text=wrap, vertical="top",
                                horizontal=al or ("center" if fmt in (MONEY, MONEY2, PCT, DATE) else "left"))
        if fmt: c.number_format = fmt
        return c

    put("id", o["id"]); put("name", o["name"], wrap=True); put("addr", o["addr"], wrap=True)
    put("hood", o["hood"], wrap=True); put("boro", o["boro"]); put("side", o["side"]); put("cb", o["cb"])
    link(put("url", ""), o["url"], "Open advertisement")
    if o["appurl"]: link(put("appurl", ""), o["appurl"], "Apply / contact")
    else: put("appurl", "—", al="center")
    put("program", o["program"], wrap=True)
    put("pub", o["pub"] or "—", DATE)
    put("deadline", o["deadline"] if o["deadline"] else "Ongoing", DATE if o["deadline"] else None, al="center")
    put("days", f'=IF(ISNUMBER({C("deadline",r)}),{C("deadline",r)}-{I(R_TODAY)},"")', al="center")
    put("open", f'=IF(NOT(ISNUMBER({C("deadline",r)})),"OPEN (ongoing)",IF({C("days",r)}<0,"CLOSED","OPEN"))', al="center")
    put("unit", o["unit"]); put("beds", o["beds"], al="center"); put("ami", o["ami"])
    put("minimum", o["minimum"] if o["minimum"] is not None else "", MONEY)
    for k in ("max1", "max2", "max3"):
        v = o[k]
        put(k, v if v is not None else "", MONEY if isinstance(v, (int, float)) else None, al="center")
    put("maxapp",
        f'=IF({HH}=1,IF({C("max1",r)}="","",{C("max1",r)}),'
        f'IF({HH}=2,IF({C("max2",r)}="","",{C("max2",r)}),'
        f'IF({C("max3",r)}="","",{C("max3",r)})))', MONEY, al="center")
    put("inc", f"={INC}", MONEY)
    put("elig",
        f'=IF({C("maxapp",r)}="N/A","NO - no unit for this household size",'
        f'IF(AND({C("minimum",r)}<>"",{C("inc",r)}<{C("minimum",r)}),"NO - below minimum",'
        f'IF({C("maxapp",r)}="","REVIEW - limits not published",'
        f'IF({C("inc",r)}>{C("maxapp",r)},"NO - over maximum",'
        f'IF({C("minimum",r)}="","REVIEW - minimum not published","YES")))))', wrap=True)
    put("above", f'=IF(OR({C("minimum",r)}="",NOT(ISNUMBER({C("maxapp",r)}))),"",{C("inc",r)}-{C("minimum",r)})', MONEY)
    put("below", f'=IF(NOT(ISNUMBER({C("maxapp",r)})),"",{C("maxapp",r)}-{C("inc",r)})', MONEY)
    put("rent", o["rent"] if o["rent"] is not None else "", MONEY2)
    put("utilinc", o["utilinc"], wrap=True)
    put("util", o["util"], MONEY)
    put("hcost", f'=IF({C("rent",r)}="","",{C("rent",r)}+{C("util",r)})', MONEY2)
    put("gmi", f"={GMI}", MONEY2)
    put("rburd", f'=IF(OR({C("rent",r)}="",{C("gmi",r)}=0),"",{C("rent",r)}/{C("gmi",r)})', PCT)
    put("hburd", f'=IF(OR({C("hcost",r)}="",{C("gmi",r)}=0),"",{C("hcost",r)}/{C("gmi",r)})', PCT)
    put("left", f'=IF({C("hcost",r)}="","",{C("gmi",r)}-{C("hcost",r)})', MONEY2)
    put("afford",
        f'=IF({C("rent",r)}="","REVIEW - no rent published",'
        f'IF({C("rent",r)}>{AMAX},"ABOVE ABSOLUTE MAXIMUM",'
        f'IF({C("hburd",r)}<={VAFF},"VERY AFFORDABLE",'
        f'IF({C("hburd",r)}<={BURD},"AFFORDABLE",'
        f'IF({C("hburd",r)}<={BORD},"BORDERLINE","ABOVE TARGET")))))', wrap=True)
    put("save", f'=IF({C("rent",r)}="","",{PMAX}-{C("rent",r)})', MONEY2)
    put("excep", f'=IF({C("rent",r)}="","",IF(AND({C("hburd",r)}<={VAFF},{C("elig",r)}="YES"),"YES","—"))', al="center")
    put("asset", o["asset"] if o["asset"] is not None else "", MONEY)
    put("aselig", f'=IF({C("asset",r)}="","REVIEW",IF({AST}<={C("asset",r)},"YES","NO"))', al="center")
    put("mepref", o["mepref"], al="center")
    put("prefpct", o["prefpct"] if o["prefpct"] else "", PCT)
    put("station", o["station"], wrap=True); put("lines", o["lines"])
    put("walk", o["walk"], al="center")
    put("route", o["route"], wrap=True)
    put("direct", o["direct"], al="center"); put("transfers", o["transfers"], al="center")
    put("ttype", o["ttype"]); put("tdetail", o["tdetail"], wrap=True)
    put("dest", o["dest"]); put("ride", o["ride"], al="center")
    put("dd", f'=IF(OR({C("walk",r)}="",{C("ride",r)}=""),"",'
              f'{C("walk",r)}+{I(R_WAIT)}+{C("ride",r)}+{C("transfers",r)}*{I(R_TPEN)}+{I(R_WALKD)})', al="center")
    put("grade",
        f'=IF({C("dd",r)}="","",'
        f'IF(AND({C("direct",r)}="YES",{C("walk",r)}<=8,{C("dd",r)}<=45),"A+",'
        f'IF(AND({C("direct",r)}="YES",{C("dd",r)}<=60),"A",'
        f'IF(AND({C("transfers",r)}=1,OR({C("ttype",r)}="Same-platform",{C("ttype",r)}="Easy same-station"),{C("dd",r)}<=70),"B",'
        f'IF(AND({C("transfers",r)}<=1,{C("dd",r)}<=85),"C","D")))))', al="center")
    put("cok", f'=IF({C("dd",r)}="","",IF({C("dd",r)}<={I(R_CPREF)},"YES - within preferred",'
               f'IF({C("dd",r)}<={I(R_CABS)},"OK - within absolute max","NO - too long")))', wrap=True)
    put("west", "YES" if o["side"] == "West" else "—", al="center")
    put("corridor", o["corridor"])
    put("geog",
        f'=IF({C("corridor",r)}="Manhattan",IF({I(R_MAN)}="YES","YES","NO"),'
        f'IF({C("corridor",r)}="Brooklyn 2/5",IF({I(R_B25)}="YES","YES","NO"),'
        f'IF({C("corridor",r)}="Brooklyn 3/4",IF({I(R_B34)}="YES","YES","NO"),'
        f'IF({C("corridor",r)}="Brooklyn other",IF({I(R_BOT)}="YES","YES","NO"),"NO"))))', al="center")
    put("laundry", o["laundry"], al="center"); put("elevator", o["elevator"], al="center")
    put("amen", o["amen"], wrap=True)
    put("pass",
        f'=IF(AND({C("elig",r)}="YES",{C("geog",r)}="YES",'
        f'OR({C("afford",r)}="VERY AFFORDABLE",{C("afford",r)}="AFFORDABLE",{C("afford",r)}="BORDERLINE"),'
        f'{C("cok",r)}<>"NO - too long"),"YES","NO")', al="center")
    put("status", o["status"], al="center")
    put("applied", "", DATE); put("conf", ""); put("log", "")
    put("verified", TODAY, DATE)
    put("notes", o["notes"], wrap=True)
    put("qual", f'=IF(AND({C("pass",r)}="YES",{C("status",r)}<>"APPLIED",LEFT({C("open",r)},4)="OPEN"),1,0)', al="center")
    put("score",
        f'=IF({C("qual",r)}=0,"",IF(ISNUMBER({C("days",r)}),{C("days",r)},999)*10000'
        f'+IF({C("afford",r)}="VERY AFFORDABLE",1,IF({C("afford",r)}="AFFORDABLE",2,3))*100'
        f'+IF({C("grade",r)}="A+",1,IF({C("grade",r)}="A",2,IF({C("grade",r)}="B",3,IF({C("grade",r)}="C",4,5))))'
        f'+ROW()/10000)')
    put("rank", f'=IF({C("qual",r)}=0,"",SUMPRODUCT((${KEY["qual"]}${r0}:${KEY["qual"]}${rn}=1)*'
                f'(${KEY["score"]}${r0}:${KEY["score"]}${rn}<{C("score",r)}))+1)', al="center")
    wa.row_dimensions[r].height = 58

tab = Table(displayName="tblActive", ref=f"A{HROW}:{gcl(len(COLS))}{rn}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
wa.add_table(tab)
wa.freeze_panes = f"E{r0}"
wa.auto_filter.ref = f"A{HROW}:{gcl(len(COLS))}{rn}"

stat_list = ('"NEW,CHECKING ELIGIBILITY,ELIGIBLE,APPLY NOW,APPLIED,DOCUMENT REQUEST,DOCUMENTS SUBMITTED,'
             'PROCESSING,INTERVIEW,WAITLIST,OFFER,NOT SELECTED,REJECTED,EXPIRED,NOT ELIGIBLE"')
dv = DataValidation(type="list", formula1=stat_list, allow_blank=True)
wa.add_data_validation(dv)
dv.add(f"{KEY['status']}{r0}:{KEY['status']}{rn}")
dvp = DataValidation(type="list", formula1='"YES,NO,UNCERTAIN"', allow_blank=True)
wa.add_data_validation(dvp); dvp.add(f"{KEY['mepref']}{r0}:{KEY['mepref']}{rn}")

def cf(col, rng_from=r0, rng_to=None, rules=()):
    rng_to = rng_to or rn
    ref = f"{KEY[col]}{rng_from}:{KEY[col]}{rng_to}"
    for formula, fill, fontc in rules:
        wa.conditional_formatting.add(ref, FormulaRule(
            formula=[formula.format(c=f"${KEY[col]}{rng_from}")],
            fill=PatternFill("solid", bgColor=fill), font=Font(name=FONT, size=9, bold=True, color=fontc)))

cf("elig", rules=[('LEFT({c},3)="YES"', GREEN, GREENF),
                  ('LEFT({c},6)="REVIEW"', AMBER, AMBERF),
                  ('LEFT({c},2)="NO"', RED, REDF)])
cf("afford", rules=[('OR({c}="VERY AFFORDABLE",{c}="AFFORDABLE")', GREEN, GREENF),
                    ('{c}="BORDERLINE"', AMBER, AMBERF),
                    ('OR({c}="ABOVE TARGET",{c}="ABOVE ABSOLUTE MAXIMUM")', RED, REDF)])
cf("grade", rules=[('OR({c}="A+",{c}="A")', GREEN, GREENF), ('{c}="B"', "D9EAD3", "274E13"),
                   ('{c}="C"', AMBER, AMBERF), ('{c}="D"', RED, REDF)])
cf("pass", rules=[('{c}="YES"', GREEN, GREENF), ('{c}="NO"', "F2F2F2", GREY)])
cf("geog", rules=[('{c}="YES"', GREEN, GREENF), ('{c}="NO"', RED, REDF)])
cf("aselig", rules=[('{c}="YES"', GREEN, GREENF), ('{c}="NO"', RED, REDF)])
cf("excep", rules=[('{c}="YES"', "FFD966", "7F6000")])
cf("mepref", rules=[('{c}="YES"', GREEN, GREENF), ('{c}="UNCERTAIN"', AMBER, AMBERF)])
cf("cok", rules=[('LEFT({c},3)="YES"', GREEN, GREENF), ('LEFT({c},2)="OK"', AMBER, AMBERF),
                 ('LEFT({c},2)="NO"', RED, REDF)])
wa.conditional_formatting.add(f"{KEY['days']}{r0}:{KEY['days']}{rn}",
    CellIsRule(operator="between", formula=["0", "7"], fill=PatternFill("solid", bgColor=RED),
               font=Font(name=FONT, size=9, bold=True, color=REDF)))
wa.conditional_formatting.add(f"{KEY['days']}{r0}:{KEY['days']}{rn}",
    CellIsRule(operator="between", formula=["8", "14"], fill=PatternFill("solid", bgColor=AMBER),
               font=Font(name=FONT, size=9, bold=True, color=AMBERF)))

# ========================================================== 4. APPLY_NOW
wn = wb.create_sheet("APPLY_NOW")
title(wn, "APPLY NOW  —  everything you are eligible for, can afford, and can commute from",
      "Pulled automatically from ACTIVE_OPPORTUNITIES. Sorted by deadline, then affordability, then transit quality — "
      "NOT by any guess at your odds of winning. If a row is blank, nothing else currently qualifies.",
      span=13)
AN = [("Deadline","deadline",13),("Days left","days",10),("Building","name",26),("Neighbourhood","hood",22),
      ("Borough","boro",11),("Unit","unit",13),("AMI","ami",12),("Rent","rent",12),
      ("Rent burden","rburd",11),("Nearest station","station",24),("Route to work","route",52),
      ("Commute (min)","dd",12),("Grade","grade",9),("Muni pref","mepref",11),
      ("Affordability","afford",22),("Apply","appurl",22)]
hdr(wn, 3, [a for a, _, _ in AN], [w for _, _, w in AN])
for k in range(1, 21):
    r = 3 + k
    for j, (_, key, _) in enumerate(AN, start=1):
        col = KEY[key]
        f = (f'=IFERROR(INDEX(ACTIVE_OPPORTUNITIES!${col}${r0}:${col}${rn},'
             f'MATCH({k},ACTIVE_OPPORTUNITIES!${KEY["rank"]}${r0}:${KEY["rank"]}${rn},0)),"")')
        c = wn.cell(row=r, column=j, value=f)
        c.font = F(9); c.border = BOX
        c.alignment = Alignment(wrap_text=(key in ("route", "name", "station", "afford")), vertical="top")
        if key in ("rent",): c.number_format = MONEY2
        if key == "rburd": c.number_format = PCT
        if key == "deadline": c.number_format = DATE
    wn.row_dimensions[r].height = 46
wn.freeze_panes = "A4"
wn.conditional_formatting.add(f"M4:M23", FormulaRule(formula=['OR($M4="A+",$M4="A")'],
    fill=PatternFill("solid", bgColor=GREEN), font=Font(name=FONT, size=9, bold=True, color=GREENF)))
wn.conditional_formatting.add(f"O4:O23", FormulaRule(formula=['OR($O4="VERY AFFORDABLE",$O4="AFFORDABLE")'],
    fill=PatternFill("solid", bgColor=GREEN), font=Font(name=FONT, size=9, bold=True, color=GREENF)))
wn.conditional_formatting.add(f"B4:B23", CellIsRule(operator="between", formula=["0","7"],
    fill=PatternFill("solid", bgColor=RED), font=Font(name=FONT, size=9, bold=True, color=REDF)))

# ======================================================= 5. NEW_LISTINGS
wi = wb.create_sheet("NEW_LISTINGS")
title(wi, "NEW LISTINGS INBOX  —  every newly discovered listing lands here first",
      "Triage sheet. Log anything you find before deciding. Nothing should reach ACTIVE_OPPORTUNITIES without passing through here.",
      span=15)
IH = ["Date discovered","Building","Address","Neighbourhood","Source","Official listing","Deadline",
      "Studio available?","1BR available?","Potential income match?","Potential rent match?",
      "Transit appears viable?","Already in database?","Reviewed?","Result / next action"]
IW = [13,26,28,22,20,22,13,12,12,26,20,18,14,11,60]
hdr(wi, 3, IH, IW)
for i, row in enumerate(INBOX):
    r = 4 + i
    for j, v in enumerate(row, start=1):
        if j == 6:
            c = wi.cell(row=r, column=j); link(c, v, "Open listing")
        else:
            c = wi.cell(row=r, column=j, value=v)
            c.font = F(9)
        c.border = BOX
        c.alignment = Alignment(wrap_text=j in (2,3,4,10,15), vertical="top")
    wi.row_dimensions[r].height = 42
for r in range(4 + len(INBOX), 4 + len(INBOX) + 40):
    for j in range(1, 16):
        wi.cell(row=r, column=j).border = BOX
    wi.row_dimensions[r].height = 22
LAST_I = 4 + len(INBOX) + 39
dvr = DataValidation(type="list", formula1='"Yes,No,Unknown"', allow_blank=True)
wi.add_data_validation(dvr)
for col in ("H","I","M","N"): dvr.add(f"{col}4:{col}{LAST_I}")
wi.freeze_panes = "A4"
wi.auto_filter.ref = f"A3:O{LAST_I}"

# ========================================================== 6. TRACKER
wt = wb.create_sheet("APPLICATION_TRACKER")
title(wt, "APPLICATION TRACKER  —  from submitted to outcome",
      "One row per submitted application. Statuses are a dropdown. Remember: once you are being processed for a unit "
      "you must withdraw your other in-process applications (Marketing Handbook 5-5.C).", span=13)
TH = ["Project","Unit / AMI","Deadline","Application URL","Date applied","Status","Confirmation number",
      "Log number","Last contact","Documents requested","Document deadline","Next action","Next-action date","Notes"]
TW = [28,18,13,24,13,22,18,14,13,34,14,34,14,50]
hdr(wt, 3, TH, TW)
for r in range(4, 64):
    for j in range(1, 15):
        c = wt.cell(row=r, column=j); c.border = BOX; c.font = F(9)
        c.alignment = Alignment(wrap_text=j in (10,12,14), vertical="top")
    wt.cell(row=r, column=3).number_format = DATE
    wt.cell(row=r, column=5).number_format = DATE
    wt.cell(row=r, column=9).number_format = DATE
    wt.cell(row=r, column=11).number_format = DATE
    wt.cell(row=r, column=13).number_format = DATE
    wt.row_dimensions[r].height = 22
dvt = DataValidation(type="list", formula1=stat_list, allow_blank=True)
wt.add_data_validation(dvt); dvt.add("F4:F63")
wt.conditional_formatting.add("F4:F63", FormulaRule(formula=['OR($F4="OFFER",$F4="INTERVIEW")'],
    fill=PatternFill("solid", bgColor=GREEN), font=Font(name=FONT, size=9, bold=True, color=GREENF)))
wt.conditional_formatting.add("F4:F63", FormulaRule(formula=['OR($F4="NOT SELECTED",$F4="REJECTED",$F4="EXPIRED",$F4="NOT ELIGIBLE")'],
    fill=PatternFill("solid", bgColor=RED), font=Font(name=FONT, size=9, bold=True, color=REDF)))
wt.conditional_formatting.add("F4:F63", FormulaRule(formula=['OR($F4="DOCUMENT REQUEST",$F4="APPLY NOW")'],
    fill=PatternFill("solid", bgColor=AMBER), font=Font(name=FONT, size=9, bold=True, color=AMBERF)))
wt.freeze_panes = "A4"

# ========================================================== 7. TRANSIT
wr = wb.create_sheet("TRANSIT_ANALYSIS")
title(wr, "TRANSIT ANALYSIS  —  Manhattan & Brooklyn to Crotona Park East",
      "Train times are MEDIAN SCHEDULED minutes from the MTA GTFS feed (weekday schedule effective 2026-05-26 to 2026-10-31), "
      "AM departures 06:30-08:30, measured station-to-station. Door-to-door adds your walks, wait and transfer penalty from INPUTS.",
      span=14)
TRH = ["Neighbourhood","Station","Lines","Walkability of area to station","Direct 2?","Direct 5?",
       "Transfer required?","Transfer type","Destination","Train min to Freeman St","Train min to 174 St",
       "Door-to-door (min)","Grade","Notes"]
TRW = [26,30,16,30,9,9,12,18,14,12,12,12,8,74]
hdr(wr, 3, TRH, TRW)
for i, t in enumerate(TRANSIT):
    r = 4 + i
    vals = list(t)
    for j, v in enumerate(vals, start=1):
        c = wr.cell(row=r, column=j, value=v)
        c.font = F(9); c.border = BOX
        c.alignment = Alignment(wrap_text=j in (1,4,14), vertical="top",
                                horizontal="center" if j in (5,6,7,10,11,13) else "left")
    # insert door-to-door formula at col 12 (shift notes/grade)
    wr.row_dimensions[r].height = 34
# rebuild columns 12-14 properly (door-to-door formula, grade, notes)
for i, t in enumerate(TRANSIT):
    r = 4 + i
    freeman = t[9]
    wr.cell(row=r, column=12,
            value=f'=8+{I(R_WAIT)}+{freeman}+IF($G{r}="YES",{I(R_TPEN)},0)+{I(R_WALKD)}').font = F(9, b=True)
    wr.cell(row=r, column=12).border = BOX
    wr.cell(row=r, column=12).alignment = Alignment(horizontal="center")
    c = wr.cell(row=r, column=13, value=t[11]); c.font = F(9, b=True); c.border = BOX
    c.alignment = Alignment(horizontal="center")
    c = wr.cell(row=r, column=14, value=t[12]); c.font = F(9); c.border = BOX
    c.alignment = Alignment(wrap_text=True, vertical="top")
wr.freeze_panes = "C4"
wr.auto_filter.ref = f"A3:N{3+len(TRANSIT)}"
wr.conditional_formatting.add(f"M4:M{3+len(TRANSIT)}", FormulaRule(formula=['OR($M4="A+",$M4="A")'],
    fill=PatternFill("solid", bgColor=GREEN), font=Font(name=FONT, size=9, bold=True, color=GREENF)))
wr.conditional_formatting.add(f"M4:M{3+len(TRANSIT)}", FormulaRule(formula=['$M4="B"'],
    fill=PatternFill("solid", bgColor="D9EAD3"), font=Font(name=FONT, size=9, bold=True, color="274E13")))
wr.conditional_formatting.add(f"M4:M{3+len(TRANSIT)}", FormulaRule(formula=['$M4="C"'],
    fill=PatternFill("solid", bgColor=AMBER), font=Font(name=FONT, size=9, bold=True, color=AMBERF)))
wr.conditional_formatting.add(f"M4:M{3+len(TRANSIT)}", FormulaRule(formula=['$M4="D"'],
    fill=PatternFill("solid", bgColor=RED), font=Font(name=FONT, size=9, bold=True, color=REDF)))

r2 = 5 + len(TRANSIT) + 1
wr.cell(row=r2, column=1, value="CORRIDOR & TRANSFER REFERENCE").font = Font(name=FONT, size=12, bold=True, color=NAVY)
CORR = [
 ("2 train corridor (7 Av / Lenox Av)","Direct, no transfer","None",
  "THE backbone. The 2 runs local through the entire Bronx White Plains Road line at all times, so it ALWAYS stops at Freeman St and 174 St. In Manhattan it runs express on 7 Av (96, 72, Times Sq, 34, 14, Chambers) and local on Lenox in Harlem (135, 125, 116, 110). Best single line for this commute."),
 ("5 train corridor (Lexington Av)","Direct, no transfer","None",
  "Also serves Freeman St and 174 St. Confirmed against MTA's station dataset and GTFS: 27 of 31 AM northbound Bronx trips stop at 174 St. Faster than the 2 from Brooklyn because it runs express there. Weekday daytime only for Brooklyn service; late nights it runs solely as a Bronx shuttle between Dyre Av and E 180 St."),
 ("3 -> 2 / 3 -> 5","One transfer","Same-platform",
  "The easiest transfer in the system for this trip. The 2 and 3 share the same express tracks in Manhattan, so you simply step across or stay put. In Brooklyn the 3 meets the 2/4/5 at Franklin Av cross-platform. The 3 itself terminates at Harlem-148 St and never reaches the Bronx."),
 ("1 -> 2","One transfer","Same-platform",
  "At 96 St and 72 St the 1, 2 and 3 share island platforms — a genuine cross-platform step. This is what makes the Upper West Side workable. But from ABOVE 96 St the 1 forces you to ride SOUTH before going north, a real penalty for Hamilton Heights and Washington Heights."),
 ("A/C -> 2 or 5","One transfer","Easy same-station",
  "Best point is Fulton St / Fulton Center, where the A/C meet the 2/3 and 4/5 inside one fare-controlled complex. From DUMBO, High St to Fulton St is just 3 minutes. Also possible at 14 St (A/C/E to 1/2/3 via the passageway) and Park Place/Chambers."),
 ("B/D connections","Two transfers or a bus","Complicated",
  "AVOID. The B/D serve the Grand Concourse in the Bronx, not the White Plains Road line. Critically, '174-175 Sts' on the B/D is a completely different station about a mile west of the 2/5 '174 St'. In Harlem there is NO free transfer between 125 St (A/B/C/D, St Nicholas Av) and 125 St (2/3, Lenox Av)."),
 ("4 -> 5","One transfer","Same-platform",
  "The 4 and 5 share the Lexington Av express tracks through Manhattan, so switching is trivial — often the same platform. But in the Bronx the 4 diverges to Jerome Av (Woodlawn) and never reaches Crotona Park East, so you must be on a 5."),
 ("6 -> 5","One transfer","Same-platform",
  "At 125 St, 86 St, 59 St, Grand Central and 14 St-Union Sq the 4/5/6 share island platforms. A 6 rider in East Harlem or the UES is effectively one easy step from a direct 5. This is what makes The Carolina a 35-minute commute."),
]
hdr(wr, r2 + 1, ["Corridor / connection","Directness","Transfer type","Assessment"], [34,20,20,140], height=22)
for i, (a, b, c_, d) in enumerate(CORR):
    r = r2 + 2 + i
    for j, v in enumerate((a, b, c_, d), start=1):
        cc = wr.cell(row=r, column=j, value=v); cc.font = F(9); cc.border = BOX
        cc.alignment = Alignment(wrap_text=True, vertical="top")
    wr.row_dimensions[r].height = 52

# ========================================================= 8. DOCUMENTS
wd = wb.create_sheet("DOCUMENT_CHECKLIST")
title(wd, "DOCUMENT CHECKLIST",
      "HPD reduced required documents in 2025 for straightforward employed applicants. Keep everything current anyway — "
      "when your log number comes up you typically get a short window to respond.", span=8)
DH = ["Document","Needed at initial application?","Needed if contacted?","How current must it be?",
      "Available?","Last updated","Needs refresh?","Notes"]
DW = [46,15,15,34,11,13,12,86]
hdr(wd, 3, DH, DW)
for i, d in enumerate(DOCS):
    r = 4 + i
    for j, v in enumerate(d, start=1):
        c = wd.cell(row=r, column=j, value=v); c.font = F(9); c.border = BOX
        c.alignment = Alignment(wrap_text=j in (1,4,8), vertical="top",
                                horizontal="center" if j in (2,3,5,6,7) else "left")
    wd.cell(row=r, column=6).number_format = DATE
    wd.row_dimensions[r].height = 44
dvd = DataValidation(type="list", formula1='"Yes,No,In progress"', allow_blank=True)
wd.add_data_validation(dvd); dvd.add(f"E4:E{3+len(DOCS)}"); dvd.add(f"G4:G{3+len(DOCS)}")
wd.freeze_panes = "A4"
wd.conditional_formatting.add(f"E4:E{3+len(DOCS)}", FormulaRule(formula=['$E4="Yes"'],
    fill=PatternFill("solid", bgColor=GREEN), font=Font(name=FONT, size=9, color=GREENF)))
wd.conditional_formatting.add(f"E4:E{3+len(DOCS)}", FormulaRule(formula=['$E4="No"'],
    fill=PatternFill("solid", bgColor=RED), font=Font(name=FONT, size=9, color=REDF)))

# ====================================================== 9. RULES & SOURCES
wru = wb.create_sheet("RULES_AND_SOURCES")
title(wru, "RULES & SOURCES  —  every rule that governs this application, with its citation",
      "Each rule is quoted or closely paraphrased from the official source and linked. Re-verify anything older than "
      "six months, and re-check the AMI chart every spring.", span=7)
RH = ["Rule","What it actually says","Effect on THIS applicant","Official source","Link","Effective date","Last verified"]
RW = [40,80,80,40,22,16,13]
hdr(wru, 3, RH, RW)
for i, rr in enumerate(RULES):
    r = 4 + i
    for j, v in enumerate(rr, start=1):
        if j == 5:
            c = wru.cell(row=r, column=j); link(c, v, "Open source")
        else:
            c = wru.cell(row=r, column=j, value=v); c.font = F(9)
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="top")
    wru.row_dimensions[r].height = 76
wru.freeze_panes = "A4"
wru.auto_filter.ref = f"A3:G{3+len(RULES)}"

# =================================================== 10. SOURCES TO MONITOR
wsm = wb.create_sheet("SOURCES_TO_MONITOR")
title(wsm, "SOURCES TO MONITOR  —  the discovery engine",
      "Housing Connect alone would have surfaced ZERO eligible units for you today. Most actionable inventory is in "
      "re-rentals advertised directly by marketing agents. Work down this list on the stated cadence.", span=7)
SH = ["Source","URL","Type of opportunity","How often it changes","How often to check","Last checked","Notes"]
SW = [40,26,44,36,18,13,90]
hdr(wsm, 3, SH, SW)
for i, s in enumerate(SOURCES):
    r = 4 + i
    for j, v in enumerate(s, start=1):
        if j == 2:
            c = wsm.cell(row=r, column=j); link(c, v, "Open")
        else:
            c = wsm.cell(row=r, column=j, value=v); c.font = F(9)
        c.border = BOX
        c.alignment = Alignment(wrap_text=j in (1,3,4,7), vertical="top")
    wsm.row_dimensions[r].height = 38
wsm.freeze_panes = "A4"
wsm.auto_filter.ref = f"A3:G{3+len(SOURCES)}"

# ====================================================== 11. AMI REFERENCE
wam = wb.create_sheet("AMI_REFERENCE")
title(wam, "2026 AMI REFERENCE  —  New York City",
      "Source: HPD Area Median Income page. 100% AMI = $118,800 for one person and $169,600 for four people. "
      "Individual advertisements govern — always use the published figures, not this table.", span=8)
hdr(wam, 3, ["AMI band","1 person","2 people","3 people","4 people","ASSET LIMIT (rental) = 4-person figure",
             "Indicative studio rent","Eligible at your income?"], [12,14,14,14,14,26,18,26])
BANDS = [20,30,40,50,60,70,80,90,100,110,120,130,165]
B1, B4 = 118800, 169600
for i, b in enumerate(BANDS):
    r = 4 + i
    wam.cell(row=r, column=1, value=f"{b}%").font = F(10, b=True)
    for j, mult in enumerate([1.0, 1.143, 1.286, 1.4286], start=2):
        c = wam.cell(row=r, column=j, value=round(B1 * b / 100 * mult / 10) * 10)
        c.number_format = MONEY; c.font = F(10)
    c = wam.cell(row=r, column=5, value=round(B4 * b / 100))
    c.number_format = MONEY; c.font = F(10)
    c = wam.cell(row=r, column=6, value=f"=E{r}"); c.number_format = MONEY; c.font = F(10, b=True)
    c = wam.cell(row=r, column=7, value=round(B1 * b / 100 * 0.30 / 12 / 5) * 5)
    c.number_format = MONEY; c.font = F(10, it=True)
    c = wam.cell(row=r, column=8,
        value=f'=IF({INC}>INDEX(B{r}:D{r},{HH}),"NO - income above this band\'s cap",'
              f'"YES - within cap (check each listing\'s minimum)")')
    c.font = F(9)
    for j in range(1, 9):
        wam.cell(row=r, column=j).border = BOX
        wam.cell(row=r, column=j).alignment = Alignment(horizontal="center" if j > 1 else "left")
wam.conditional_formatting.add(f"H4:H{3+len(BANDS)}", FormulaRule(formula=['LEFT($H4,3)="YES"'],
    fill=PatternFill("solid", bgColor=GREEN), font=Font(name=FONT, size=9, color=GREENF)))
wam.conditional_formatting.add(f"H4:H{3+len(BANDS)}", FormulaRule(formula=['LEFT($H4,2)="NO"'],
    fill=PatternFill("solid", bgColor=RED), font=Font(name=FONT, size=9, color=REDF)))
nr = 6 + len(BANDS)
for k, txt in enumerate([
  "Household-size multipliers follow the standard HUD convention (1 person = 70% of the 4-person figure, 2 = 80%, 3 = 90%).",
  "Verified against live advertisements: The Lirio's 100% AMI asset limit of $169,600 and 30% limit of $50,880 match this table exactly.",
  "Indicative studio rent is 30% of the band's 1-person income divided by 12 — a sanity check only, never a substitute for the advertised rent.",
  "The asset limit for a rental is always the FOUR-person income limit at the unit's AMI, whatever your household size (Handbook 5-5.A(6)).",
]):
    c = wam.cell(row=nr + k, column=1, value=txt); c.font = F(9, it=True, color="404040")
wam.freeze_panes = "B4"

# ========================================================= 12. SCENARIOS
wsc = wb.create_sheet("SCENARIOS")
title(wsc, "SCENARIOS  —  how many opportunities open up as income and household size change",
      "Counts rows in ACTIVE_OPPORTUNITIES that are in allowed geography and whose published minimum/maximum bracket the "
      "modelled income. Nothing here feeds the rest of the workbook — the master income on INPUTS is what governs.", span=6)
minr = f"ACTIVE_OPPORTUNITIES!${KEY['minimum']}${r0}:${KEY['minimum']}${rn}"
geo  = f"ACTIVE_OPPORTUNITIES!${KEY['geog']}${r0}:${KEY['geog']}${rn}"
maxr = {h: f"ACTIVE_OPPORTUNITIES!${KEY['max'+str(h)]}${r0}:${KEY['max'+str(h)]}${rn}" for h in (1,2,3)}

wsc.cell(row=3, column=1, value="A.  SINGLE APPLICANT — income sensitivity").font = Font(name=FONT, size=12, bold=True, color=NAVY)
hdr(wsc, 4, ["Modelled household income","Opportunities open (HH 1)","Opportunities open (HH 2)","Opportunities open (HH 3)",
             "Comment"], [26,22,22,22,90], height=34)
incomes = [55000,60000,65000,70000,71280,72000,75000,80000,85000,90000,100000,110000,125000,150000]
notes_by_inc = {
 71280:"Exactly the 60% AMI cap for one person — at or below this the cheapest units open up.",
 72000:"Your current estimate. $720 over the 60% cap.",
 75000:"A modest raise changes nothing at 60% AMI but clears some minimums.",
 90000:"Approaching the 80% AMI cap of $95,040.",
 110000:"Above 90% AMI; several low-rent bands close.",
 150000:"Only the 130%/165% bands remain — rents there are market-like.",
}
for i, inc in enumerate(incomes):
    r = 5 + i
    c = wsc.cell(row=r, column=1, value=inc); c.number_format = MONEY; c.font = F(10, b=True); c.border = BOX
    for h in (1, 2, 3):
        f = (f'=SUMPRODUCT(ISNUMBER({minr})*ISNUMBER({maxr[h]})*({geo}="YES")'
             f'*($A{r}>={minr})*($A{r}<={maxr[h]}))')
        cc = wsc.cell(row=r, column=1 + h, value=f); cc.font = F(10); cc.border = BOX
        cc.alignment = Alignment(horizontal="center")
    cc = wsc.cell(row=r, column=5, value=notes_by_inc.get(inc, "")); cc.font = F(9, it=True, color="404040")
    cc.border = BOX; cc.alignment = Alignment(wrap_text=True, vertical="top")
    wsc.row_dimensions[r].height = 22

b0 = 5 + len(incomes) + 2
wsc.cell(row=b0, column=1, value="B.  THE BROTHERS OPTION — what changes with 2 or 3 siblings").font = Font(name=FONT, size=12, bold=True, color=NAVY)
hdr(wsc, b0 + 1, ["Point","Detail"], [34,150], height=22)
BR = [
 ("Do siblings qualify as one household?",
  "YES, automatically. The Marketing Handbook defines 'Immediate Family Member' to include a SIBLING, and two or more Immediate Family Members are a household under Section 5-2(ii)(a). No proof of financial interdependence is needed — unlike a household of friends, who would have to show a shared lease or shared bank accounts."),
 ("What combines?",
  "All brothers' gross incomes add together and are tested against the unit minimum and the household-size maximum. All brothers' assets add together against ONE asset limit (the four-person figure), not one limit each."),
 ("Why this helps most",
  "His binding constraint is usually the MINIMUM income, not the maximum. Combined income clears minimums that lock him out alone — for example 90 Sands' 120% AMI 1-bedroom needs $73,098 and he falls $1,098 short by himself."),
 ("Caps rise too",
  "2-person caps are about 1.143x the 1-person cap and 3-person about 1.286x. At 60% AMI the cap moves from $71,280 to $81,420 (2 people) or $91,620 (3 people) — which turns the $1,134 studio and $1,215 one-bedroom at 90 Sands from near-misses into live options."),
 ("Unit sizes available",
  "Studio holds 1-2 people; 1BR holds 1-3; 2BR requires 2-5; 3BR requires 3-7. Two brothers can take a studio, 1BR or 2BR. Three brothers can take a 1BR, 2BR or 3BR. Where you qualify for more than one size, the Handbook says the household chooses."),
 ("The 10% preference still applies",
  "The addendum requires only that AT LEAST ONE household member be a municipal employee. The DOE brother's status covers all of them."),
 ("What every brother must satisfy",
  "None may own residential property within 100 miles of NYC. Each must surrender any current lease and occupy the new unit as their sole primary residence at least 270 days a year. Every adult signs the asset certification."),
 ("The big warning",
  "Decide the household BEFORE applying. Once the application snapshot is taken, a change in household composition that affects eligibility means REJECTION, unless it is a narrow extenuating circumstance (death, birth, separation, custody order, domestic violence). You cannot add or drop a brother mid-process."),
 ("The risk to watch",
  "Combined income can overshoot. Two brothers at $72,000 and $60,000 total $132,000 — that clears every minimum on the board but breaches the 80% AMI 2-person cap of $108,560. Model the exact combination on INPUTS before committing."),
]
for i, (a, b) in enumerate(BR):
    r = b0 + 2 + i
    ca = wsc.cell(row=r, column=1, value=a); ca.font = F(10, b=True); ca.border = BOX
    ca.alignment = Alignment(wrap_text=True, vertical="top")
    cb = wsc.cell(row=r, column=2, value=b); cb.font = F(9); cb.border = BOX
    cb.alignment = Alignment(wrap_text=True, vertical="top")
    wsc.row_dimensions[r].height = 50

# ========================================================= 1. DASHBOARD
wdb = wb.create_sheet("DASHBOARD", 0)
title(wdb, "NYC AFFORDABLE HOUSING  —  APPLICATION CONTROL PANEL",
      f"Built 18 September 2026.  All figures verified against official HPD / HDC / Housing Connect / MTA sources on that date.  "
      f"Edit INPUTS to change anything.", span=12)
wdb.column_dimensions["A"].width = 46
for col, w in zip("BCDEFGHIJKL", [16,14,16,14,13,22,44,13,10,12,22]):
    wdb.column_dimensions[col].width = w

def kv(r, label, formula, fmt=None, bold=False, note=None):
    c = wdb.cell(row=r, column=1, value=label); c.font = F(10, b=bold)
    v = wdb.cell(row=r, column=2, value=formula)
    v.font = F(11, b=True, color=NAVY); v.alignment = Alignment(horizontal="center")
    v.fill = PatternFill("solid", fgColor=LIGHT); v.border = BOX
    if fmt: v.number_format = fmt
    if note:
        nc = wdb.cell(row=r, column=3, value=note); nc.font = F(9, it=True, color="404040")
    wdb.row_dimensions[r].height = 19
    return r

sect(wdb, 3, "YOUR SITUATION", span=12)
kv(4,  "Total countable household income (gross)", f"={INC}", MONEY, True)
kv(5,  "Gross monthly income", f"={GMI}", MONEY2)
kv(6,  "Household size", f"={HH}", None, True, "1 = alone, 2 or 3 = with brothers (siblings qualify automatically)")
kv(7,  "Target rent burden", f"={BURD}", PCT)
kv(8,  "Target monthly housing cost", f"={I(R_TGT)}", MONEY2)
kv(9,  "Preferred / absolute maximum rent", f'=TEXT({PMAX},"$#,##0")&"  /  "&TEXT({AMAX},"$#,##0")')
kv(10, "Countable household assets", f"={AST}", MONEY)
kv(11, "DOE employee / municipal employee", f'={I(R_DOE)}&"  /  "&{I(R_MUNI)}', None, False,
   "Qualifies for the 10% Municipal Employee / Military Veteran preference")
kv(12, "Current residence", f"={I(R_RES)}", None, False, "Non-NYC: you are processed AFTER all NYC residents")
kv(13, "Work destination", f"={I(R_WORK)}", None, False, "Best stations: Freeman St and 174 St, both on the 2 and the 5")

sect(wdb, 15, "PIPELINE COUNTS  (live)", span=12)
Q = lambda col: f"ACTIVE_OPPORTUNITIES!${KEY[col]}${r0}:${KEY[col]}${rn}"
kv(16, "Listings tracked in the database", f'=COUNTA({Q("id")})')
kv(17, "In allowed geography and currently open",
   f'=SUMPRODUCT(({Q("geog")}="YES")*(LEFT({Q("open")},4)="OPEN"))')
kv(18, "Income eligible", f'=SUMPRODUCT(({Q("geog")}="YES")*({Q("elig")}="YES"))')
kv(19, "Income eligible AND affordable",
   f'=SUMPRODUCT(({Q("geog")}="YES")*({Q("elig")}="YES")*(({Q("afford")}="VERY AFFORDABLE")+({Q("afford")}="AFFORDABLE")+({Q("afford")}="BORDERLINE")))')
kv(20, "Affordable with an A+ or A commute",
   f'=SUMPRODUCT(({Q("geog")}="YES")*({Q("elig")}="YES")*(({Q("grade")}="A+")+({Q("grade")}="A"))*(({Q("afford")}="VERY AFFORDABLE")+({Q("afford")}="AFFORDABLE")+({Q("afford")}="BORDERLINE")))')
kv(21, "Affordable with DIRECT 2 or 5 service",
   f'=SUMPRODUCT(({Q("geog")}="YES")*({Q("elig")}="YES")*({Q("direct")}="YES")*(({Q("afford")}="VERY AFFORDABLE")+({Q("afford")}="AFFORDABLE")+({Q("afford")}="BORDERLINE")))')
kv(22, "PASSES ALL THREE TESTS", f'=SUMPRODUCT(--({Q("pass")}="YES"))', None, True)
kv(23, "Carrying the 10% municipal-employee preference", f'=COUNTIF({Q("mepref")},"YES")')
kv(24, "Needing review (limits not published)", f'=SUMPRODUCT(--(LEFT({Q("elig")},6)="REVIEW"))')
kv(25, "Not yet applied to", f'=SUMPRODUCT(({Q("pass")}="YES")*({Q("status")}<>"APPLIED"))')
kv(26, "Already applied", f'=COUNTIF({Q("status")},"APPLIED")')
kv(27, "Closing within 7 days", f'=SUMPRODUCT(ISNUMBER({Q("days")})*({Q("days")}>=0)*({Q("days")}<=7))')
kv(28, "Closing within 14 days", f'=SUMPRODUCT(ISNUMBER({Q("days")})*({Q("days")}>=0)*({Q("days")}<=14))')
kv(29, "New listings logged in the inbox", f'=COUNTA(NEW_LISTINGS!$B$4:$B${LAST_I})')

sect(wdb, 31, "APPLY NOW  —  open, eligible, affordable, commutable, not yet applied to", span=12)
AN2 = [("Deadline","deadline",13),("Days","days",8),("Building","name",26),("Neighbourhood","hood",20),
       ("Unit","unit",13),("Rent","rent",12),("Burden","rburd",9),("Station","station",22),
       ("Commute","dd",9),("Grade","grade",8),("Muni pref","mepref",10),("Apply","appurl",20)]
hdr(wdb, 32, [a for a, _, _ in AN2], None, height=26)
for j, (_, _, w) in enumerate(AN2, start=1):
    wdb.column_dimensions[gcl(j)].width = w if j > 1 else 46
for k in range(1, 13):
    r = 32 + k
    for j, (_, key, _) in enumerate(AN2, start=1):
        col = KEY[key]
        f = (f'=IFERROR(INDEX(ACTIVE_OPPORTUNITIES!${col}${r0}:${col}${rn},'
             f'MATCH({k},ACTIVE_OPPORTUNITIES!${KEY["rank"]}${r0}:${KEY["rank"]}${rn},0)),"")')
        c = wdb.cell(row=r, column=j, value=f); c.font = F(9); c.border = BOX
        c.alignment = Alignment(wrap_text=key in ("name", "station"), vertical="top")
        if key == "rent": c.number_format = MONEY2
        if key == "rburd": c.number_format = PCT
        if key == "deadline": c.number_format = DATE
    wdb.row_dimensions[r].height = 30
wdb.conditional_formatting.add("J33:J44", FormulaRule(formula=['OR($J33="A+",$J33="A")'],
    fill=PatternFill("solid", bgColor=GREEN), font=Font(name=FONT, size=9, bold=True, color=GREENF)))

hr = 47
sect(wdb, hr, "HEADLINE FINDINGS  (18 September 2026)", span=12)
FIND = [
 "1.  Housing Connect alone is not enough. Today there are 12 open rental lotteries citywide and exactly ONE in Manhattan — The Lirio — and it has no studio or 1-bedroom at all (minimum household size 2). A single applicant watching only Housing Connect would find nothing.",
 "2.  The $720 problem. At $72,000 you sit just above the 60% AMI cap for one person, which is $71,280. That single $720 is what blocks the cheapest units, including a $1,134 studio and a $1,215 one-bedroom at 90 Sands in DUMBO.",
 "3.  Gross, not net. HPD uses gross income before any expenses or deductions. A $60k net figure and travel costs do not reduce it. The only exception in the rule is self-employment income.",
 "4.  Your real eligibility window as a single applicant is roughly 70%-90% AMI, with rent between about $1,500 and $2,000 — which happens to coincide almost exactly with your 30% affordability target of $1,800.",
 "5.  Applying with your brothers is the single biggest lever. Siblings are Immediate Family Members, so they qualify as one household automatically with no proof of financial interdependence. Combined income clears the MINIMUM-income tests that block you alone, and the caps rise to $81,420 (2 people) or $91,620 (3 people) at 60% AMI.",
 "6.  The best commute in Manhattan is Harlem, not the far West Side. From 135 St the 2 reaches Freeman St in 14 scheduled minutes with no transfer. From 125 St on Lexington the 5 takes 16. Hamilton Heights and Washington Heights look close on a map but force a southbound back-track to 96 St.",
 "7.  Both the 2 AND the 5 serve Freeman St and 174 St — confirmed in MTA's station dataset and GTFS schedule. In the AM your direction is reverse-peak, so the 5's rush-hour Bronx express pattern does not skip your stops.",
 "8.  Watch the name trap: '174 St' on the 2/5 is NOT '174-175 Sts' on the B/D, which is about a mile away on the Grand Concourse.",
 "9.  Your Westchester address costs you twice — NYC residents are processed before non-residents, including within the municipal-employee preference pool, and you cannot use the 20% community-board preference anywhere.",
 "10. The 10% municipal-employee preference is real and you qualify, but it does not appear on every lottery. The Lirio advertises it; 1718 Crotona Park East (a 485-x project) advertises no preference at all. Check each advertisement.",
]
for i, t in enumerate(FIND):
    r = hr + 1 + i
    c = wdb.cell(row=r, column=1, value=t); c.font = F(10)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    wdb.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    wdb.row_dimensions[r].height = 32

wr2 = hr + 1 + len(FIND) + 1
sect(wdb, wr2, "WEEKLY ROUTINE", span=12)
ROUT = [
 "DAILY/OFTEN  —  Open Housing Connect, filter to Manhattan and Brooklyn, and log anything new in NEW_LISTINGS.",
 "WEEKLY  —  Work through the HDC re-rentals page and the top marketing agents in SOURCES_TO_MONITOR. This is where the real inventory is.",
 "THEN  —  Transcribe the advertisement's published minimum and maximum income into ACTIVE_OPPORTUNITIES. Never infer them from the AMI label.",
 "THE SHEET DECIDES  —  Can he qualify? Can he afford it? Can he commute from it? Anything that passes appears in APPLY NOW.",
 "AFTER APPLYING  —  Set status to APPLIED, then move the row into APPLICATION_TRACKER with the confirmation and log number.",
 "MONTHLY  —  Re-verify anything in RULES_AND_SOURCES older than six months. Re-check the AMI chart every spring.",
]
for i, t in enumerate(ROUT):
    r = wr2 + 1 + i
    c = wdb.cell(row=r, column=1, value=t); c.font = F(10)
    wdb.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    c.alignment = Alignment(wrap_text=True, vertical="center")
    wdb.row_dimensions[r].height = 20
wdb.freeze_panes = "A4"

for s in wb.worksheets:
    s.sheet_view.showGridLines = False
wb.save(OUT)
print("saved", OUT, "| opportunity rows:", n, "| rows", r0, "-", rn)
