"""Build vendor_risk_assessment.xlsx: third-party risk tiering, due diligence scoring, SOC 2 report
review and onboarding/renewal decisions for 8 fictional vendors. Tiers, scores, ratings, report age,
conclusions, due dates and decisions are formulas; blue cells are inputs and judgments.

Usage: python3 build_vendor_assessment.py ../vendor_risk_assessment.xlsx
"""
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "vendor_risk_assessment.xlsx"
NAVY, BLUE = "1F3864", "0000FF"
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill("solid", start_color=NAVY, end_color=NAVY)
RED = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
ORG = PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD")
AMB = PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")
GRN = PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")
GRY = PatternFill("solid", start_color="E7E6E6", end_color="E7E6E6")


def f(bold=False, color="000000", size=10, italic=False):
    return Font(name="Arial", bold=bold, color=color, size=size, italic=italic)


def head(ws, row, labels, start=1):
    for i, lab in enumerate(labels, start):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill, c.border = f(True, "FFFFFF"), HEAD, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 40


def title(ws, text, sub):
    ws["A1"], ws["A2"] = text, sub
    ws["A1"].font, ws["A2"].font = f(True, NAVY, 14), f(italic=True, color="595959")


def put(ws, r, c, v, inp=False, center=False, bold=False):
    x = ws.cell(row=r, column=c, value=v)
    x.font = f(bold=bold, color=BLUE if inp else "000000")
    x.border = BOX
    x.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center" if center else None)
    return x


# name, service, data, access, criticality, replaceability, last assessed, required actions
V = [
    ("Nimbus Cloud Hosting", "Hosts ERP database backups and a disaster recovery environment", 3, 2, 3, 3, date(2025, 12, 1),
     "Confirm MFA is enforced for all administrative access; get evidence that subcontractors are assessed."),
    ("Brightpath Software", "Cloud billing application used by the revenue cycle team", 3, 2, 3, 2, date(2025, 3, 20),
     "Overdue reassessment. Obtain a bridge letter for the SOC 2 gap, require MFA, and map the 4 unmapped user entity controls."),
    ("Summit IT Services", "Managed help desk and server administration with privileged access", 2, 3, 3, 2, date(2025, 2, 1),
     "Overdue reassessment. Do not renew until MFA, termination procedures and an independent security audit are in place."),
    ("Keystone Legal LLP", "Outside counsel holding confidential contract and employment records", 3, 1, 2, 1, date(2025, 3, 10),
     "Request evidence of encryption of files in transit, and a tested incident notification process."),
    ("Pinecrest Staffing", "Temporary finance and clerical staff with limited system access", 2, 2, 2, 1, date(2024, 9, 1),
     "Overdue reassessment. Require access to be removed at the end of each assignment and MFA for system access."),
    ("Oakmont Audit Partners", "External financial statement auditor with read access to ledgers", 3, 1, 2, 2, date(2026, 1, 15),
     "No action. Confirm incident response testing at next cycle."),
    ("Lighthouse Training Co", "Online compliance training content with no access to company data", 1, 1, 1, 1, date(2024, 6, 1),
     "Collect a signed security attestation at renewal."),
    ("Harvest Food Service", "Cafeteria and catering", 1, 1, 1, 1, date(2025, 5, 1),
     "None required."),
]
NV = len(V)

Q = [  # domain, question, weight
    ("Governance", "A written information security policy approved by management is in place.", 2),
    ("Assurance", "An independent security audit or certification (SOC 2 Type II or ISO 27001) was completed in the last 12 months.", 3),
    ("Access", "Access to our data is limited by role and reviewed at least every six months.", 3),
    ("Access", "Staff access is removed within 24 hours of termination.", 3),
    ("Access", "Multi-factor authentication is required for remote and administrative access.", 3),
    ("Data protection", "Our data is encrypted at rest and in transit.", 3),
    ("Operations", "Vulnerabilities are patched on a defined schedule.", 2),
    ("Change", "Production changes are tested and approved before release.", 2),
    ("Incident response", "An incident response plan was tested in the last 12 months and includes customer notice within 72 hours.", 3),
    ("Continuity", "A business continuity and recovery plan is tested every year.", 2),
    ("Fourth parties", "Subcontractors with access to our data are assessed and bound by equivalent terms.", 2),
    ("People", "Staff complete security awareness training every year.", 1),
]
Y, P, N, X = "Yes", "Partial", "No", "N/A"
R = {  # responses per vendor in question order
    "Nimbus Cloud Hosting":   [Y, Y, Y, Y, P, Y, Y, Y, Y, Y, P, Y],
    "Brightpath Software":    [Y, Y, P, Y, N, Y, Y, P, P, Y, P, Y],
    "Summit IT Services":     [Y, N, P, P, N, Y, Y, P, P, Y, N, Y],
    "Keystone Legal LLP":     [Y, N, Y, Y, Y, Y, P, X, P, Y, Y, Y],
    "Pinecrest Staffing":     [P, N, P, P, P, P, X, X, P, P, X, P],
    "Oakmont Audit Partners": [Y, Y, Y, Y, Y, Y, Y, X, P, Y, X, Y],
    "Lighthouse Training Co": [Y, X, X, X, X, X, X, X, X, X, X, P],
    "Harvest Food Service":   [None] * 12,
}
# report type, period end, opinion, exceptions, relevant exceptions, subservice orgs, CUEC total, CUEC mapped
S = {
    "Nimbus Cloud Hosting":   ("Type II", date(2026, 3, 31), "Unqualified", 1, 0, 1, 8, 8),
    "Brightpath Software":    ("Type II", date(2025, 8, 31), "Unqualified", 2, 1, 2, 10, 6),
}

wb = Workbook()

# ---------------------------------------------------------------- Scope
sc = wb.active
sc.title = "Scope"
title(sc, "Third-Party Risk Assessment", "Fictional healthcare services company and fictional vendors. Judgments are mine and editable.")
info = [
    ("Purpose", "Decide how much diligence each vendor needs, score what they told us, review their SOC 2 report, and decide whether to onboard, renew or condition each one."),
    ("Process", "1) Tiering scores inherent risk from four factors. 2) Questionnaire scores control strength from 12 weighted questions. "
                "3) SOC 2 Review checks the report's age, opinion, exceptions and user entity controls. "
                "4) Decisions combines tier and control strength into residual risk, sets the next reassessment date and gives a decision."),
    ("Inherent vs residual", "Inherent risk comes from what the vendor does for us (data, access, criticality, replaceability). Residual risk is what remains given how strong their controls are."),
    ("Judgment", "Factor scores, questionnaire answers and thresholds are my illustrative judgments. Change them on Scale, Tiering or Questionnaire and every result updates."),
    ("Link to risk register", "The risk register for this company flags critical vendors not reassessed in over 18 months (R-15). Brightpath and Summit are the two overdue Tier 1 vendors here."),
    ("Colour key", "Blue text is an input. Black text is a formula."),
    ("Data", "All vendors and answers are fictional. A real program verifies answers with evidence, not self-reported responses alone."),
]
for i, (k, v) in enumerate(info):
    sc.cell(row=4 + i, column=1, value=k).font = f(True)
    c = sc.cell(row=4 + i, column=2, value=v)
    c.font, c.alignment = f(), Alignment(wrap_text=True, vertical="top")
    sc.cell(row=4 + i, column=1).alignment = Alignment(vertical="top")
sc.column_dimensions["A"].width = 22
sc.column_dimensions["B"].width = 115
sc.sheet_view.showGridLines = False

# ---------------------------------------------------------------- Scale
sl = wb.create_sheet("Scale")
title(sl, "Weights, Thresholds and Definitions", "Edit blue cells. Everything downstream recalculates.")
head(sl, 4, ["Inherent risk factor", "Weight", "1 = low", "2 = medium", "3 = high"])
fac = [("Data sensitivity", 0.35, "No company data", "Internal or business confidential", "Regulated or personal data"),
       ("System access", 0.25, "None", "Limited (one application or read only)", "Privileged or administrative"),
       ("Business criticality", 0.25, "Easy to work around", "Disrupts one function", "Stops core operations"),
       ("Replaceability", 0.15, "Easy to switch", "Takes months to switch", "Very hard to switch")]
for i, (a, w, l1, l2, l3) in enumerate(fac):
    put(sl, 5 + i, 1, a, bold=True)
    put(sl, 5 + i, 2, w, inp=True, center=True).number_format = "0%"
    for j, t in enumerate((l1, l2, l3)):
        put(sl, 5 + i, 3 + j, t)
put(sl, 9, 1, "Total weight (should be 100%)", bold=True)
put(sl, 9, 2, "=SUM(B5:B8)", center=True, bold=True).number_format = "0%"

head(sl, 11, ["Tier", "Minimum score", "Reassess every (months)", "Diligence required"])
tiers = [("Tier 1", 2.4, 12, "Full questionnaire, SOC 2 Type II review and annual reassessment"),
         ("Tier 2", 1.5, 24, "Full questionnaire and evidence on request"),
         ("Tier 3", 0, 36, "Signed attestation; short questionnaire if data is involved")]
for i, (t, mn, mo, d) in enumerate(tiers):
    put(sl, 12 + i, 1, t, bold=True)
    put(sl, 12 + i, 2, mn, inp=True, center=True).number_format = "0.0"
    put(sl, 12 + i, 3, mo, inp=True, center=True)
    put(sl, 12 + i, 4, d, inp=True)

head(sl, 16, ["Control rating", "Minimum score", "Meaning"])
for i, (t, mn, m) in enumerate([("Strong", 0.85, "Controls meet our expectations"),
                                ("Adequate", 0.70, "Gaps exist but are manageable with conditions"),
                                ("Weak", 0.0, "Material gaps in controls")]):
    put(sl, 17 + i, 1, t, bold=True)
    put(sl, 17 + i, 2, mn, inp=True, center=True).number_format = "0%"
    put(sl, 17 + i, 3, m)

head(sl, 21, ["Residual risk", "Strong controls", "Adequate controls", "Weak controls"])
mat = [("Tier 1", "Moderate", "High", "Critical"), ("Tier 2", "Low", "Moderate", "High"), ("Tier 3", "Low", "Low", "Moderate")]
for i, row in enumerate(mat):
    put(sl, 22 + i, 1, row[0], bold=True)
    for j in range(1, 4):
        put(sl, 22 + i, 1 + j, row[j], inp=True, center=True)
put(sl, 26, 1, "SOC 2 report stale after (months)", bold=True)
put(sl, 26, 2, 12, inp=True, center=True)
put(sl, 27, 1, "Report as-of date", bold=True)
put(sl, 27, 2, date(2026, 10, 30), inp=True, center=True).number_format = "yyyy-mm-dd"
sl["C27"] = "Fixed so results are reproducible. Replace with =TODAY() for a live view."
sl["C27"].font = f(italic=True, color="595959")
for col, w in zip("ABCDE", (34, 16, 26, 36, 28)):
    sl.column_dimensions[col].width = w
ASOF, STALE = "Scale!$B$27", "Scale!$B$26"
for word, fill in (("Critical", RED), ("High", ORG), ("Moderate", AMB), ("Low", GRN)):
    sl.conditional_formatting.add("B22:D24", CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))

# ---------------------------------------------------------------- Tiering
tg = wb.create_sheet("Tiering")
title(tg, "Vendor Tiering (Inherent Risk)", "Score each factor 1 to 3 (see Scale). Score and tier are formulas.")
head(tg, 4, ["Vendor", "Service", "Data sensitivity", "System access", "Business criticality", "Replaceability",
             "Inherent score", "Tier", "Reassess every (months)", "Diligence required"])
for i, v in enumerate(V):
    r = 5 + i
    put(tg, r, 1, v[0], inp=True, bold=True)
    put(tg, r, 2, v[1], inp=True)
    for j in range(4):
        put(tg, r, 3 + j, v[2 + j], inp=True, center=True)
    put(tg, r, 7, f"=C{r}*Scale!$B$5+D{r}*Scale!$B$6+E{r}*Scale!$B$7+F{r}*Scale!$B$8", center=True, bold=True).number_format = "0.00"
    put(tg, r, 8, f'=IF(G{r}>=Scale!$B$12,Scale!$A$12,IF(G{r}>=Scale!$B$13,Scale!$A$13,Scale!$A$14))', center=True, bold=True)
    put(tg, r, 9, f"=INDEX(Scale!$C$12:$C$14,MATCH(H{r},Scale!$A$12:$A$14,0))", center=True)
    put(tg, r, 10, f"=INDEX(Scale!$D$12:$D$14,MATCH(H{r},Scale!$A$12:$A$14,0))")
    tg.row_dimensions[r].height = 40
for col, w in zip("ABCDEFGHIJ", (24, 44, 11, 10, 11, 12, 10, 8, 11, 48)):
    tg.column_dimensions[col].width = w
dv13 = DataValidation(type="whole", operator="between", formula1="1", formula2="3", showErrorMessage=True, error="Enter 1, 2 or 3.")
tg.add_data_validation(dv13)
dv13.add(f"C5:F{4 + NV}")
tg.freeze_panes = "B5"
for word, fill in (("Tier 1", RED), ("Tier 2", AMB), ("Tier 3", GRN)):
    tg.conditional_formatting.add(f"H5:H{4 + NV}", CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))

# ---------------------------------------------------------------- Questionnaire
qs = wb.create_sheet("Questionnaire")
title(qs, "Due Diligence Questionnaire", "Answers are Yes, Partial, No or N/A. Control score = weighted share of applicable questions answered Yes (Partial counts half).")
head(qs, 4, ["#", "Domain", "Question", "Weight"])
for j in range(NV):
    c = qs.cell(row=4, column=5 + j, value=f"=Tiering!A{5 + j}")
    c.font, c.fill, c.border = f(True, "FFFFFF"), HEAD, BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
QF, QL = 5, 5 + len(Q) - 1
for i, (dom, q, w) in enumerate(Q):
    r = QF + i
    put(qs, r, 1, i + 1, center=True)
    put(qs, r, 2, dom, inp=True)
    put(qs, r, 3, q, inp=True)
    put(qs, r, 4, w, inp=True, center=True)
    for j, v in enumerate(V):
        put(qs, r, 5 + j, R[v[0]][i], inp=True, center=True)
    qs.row_dimensions[r].height = 42
SC_ROW, RT_ROW, AN_ROW = QL + 2, QL + 3, QL + 4
put(qs, SC_ROW, 3, "Control score", bold=True)
put(qs, RT_ROW, 3, "Control rating", bold=True)
put(qs, AN_ROW, 3, "Questions answered", bold=True)
for j in range(NV):
    col = L(5 + j)
    rng = f"{col}{QF}:{col}{QL}"
    wts = f"$D${QF}:$D${QL}"
    app = f'({rng}<>"N/A")*({rng}<>"")'
    put(qs, SC_ROW, 5 + j,
        f'=IF(SUMPRODUCT({wts},{app})=0,"n/a",SUMPRODUCT({wts},({rng}="Yes")*1+({rng}="Partial")*0.5)/SUMPRODUCT({wts},{app}))',
        center=True, bold=True).number_format = "0%"
    put(qs, RT_ROW, 5 + j,
        f'=IF({col}{SC_ROW}="n/a","Not assessed",IF({col}{SC_ROW}>=Scale!$B$17,Scale!$A$17,IF({col}{SC_ROW}>=Scale!$B$18,Scale!$A$18,Scale!$A$19)))',
        center=True, bold=True)
    put(qs, AN_ROW, 5 + j, f'=SUMPRODUCT({app}*1)', center=True)
for col, w in zip("ABCD", (5, 16, 62, 8)):
    qs.column_dimensions[col].width = w
for j in range(NV):
    qs.column_dimensions[L(5 + j)].width = 14
qs.freeze_panes = "E5"
dvq = DataValidation(type="list", formula1='"Yes,Partial,No,N/A"', allow_blank=True)
qs.add_data_validation(dvq)
dvq.add(f"E{QF}:{L(4 + NV)}{QL}")
rr = f"E{RT_ROW}:{L(4 + NV)}{RT_ROW}"
for word, fill in (("Strong", GRN), ("Adequate", AMB), ("Weak", RED), ("Not assessed", GRY)):
    qs.conditional_formatting.add(rr, CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))
rq = f"E{QF}:{L(4 + NV)}{QL}"
qs.conditional_formatting.add(rq, CellIsRule(operator="equal", formula=['"No"'], fill=RED))
qs.conditional_formatting.add(rq, CellIsRule(operator="equal", formula=['"Partial"'], fill=AMB))

# ---------------------------------------------------------------- SOC 2 Review
so = wb.create_sheet("SOC2 Review")
title(so, "SOC 2 Report Review", "Report age, stale flag, CUEC gap and conclusion are formulas.")
head(so, 4, ["Vendor", "Tier", "Report type", "Period end", "Opinion", "Exceptions noted", "Exceptions relevant to us",
             "Subservice orgs", "User entity controls (CUECs)", "CUECs mapped to our controls", "Report age (months)",
             "Stale?", "CUEC gap", "Conclusion"])
for i, v in enumerate(V):
    r = 5 + i
    rep = S.get(v[0])
    put(so, r, 1, f"=Tiering!A{r}", bold=True)
    put(so, r, 2, f"=Tiering!H{r}", center=True)
    if rep:
        t, pe, op, ex, rel, sub, cu, cm = rep
    else:
        t, pe, op, ex, rel, sub, cu, cm = "None", None, "Not provided", None, None, None, None, None
    put(so, r, 3, t, inp=True, center=True)
    put(so, r, 4, pe, inp=True, center=True).number_format = "yyyy-mm-dd"
    put(so, r, 5, op, inp=True, center=True)
    for col, val in ((6, ex), (7, rel), (8, sub), (9, cu), (10, cm)):
        put(so, r, col, val, inp=True, center=True)
    put(so, r, 11, f'=IF(D{r}="","",ROUND(({ASOF}-D{r})/30.44,0))', center=True)
    put(so, r, 12, f'=IF(K{r}="","",IF(K{r}>{STALE},"Stale","Current"))', center=True)
    put(so, r, 13, f'=IF(I{r}="","",I{r}-J{r})', center=True)
    put(so, r, 14,
        f'=IF(E{r}="Not provided",IF(B{r}="Tier 1","Required but not provided: obtain alternative evidence","Not required"),'
        f'IF(E{r}<>"Unqualified","Escalate: opinion is "&E{r},'
        f'IF(L{r}="Stale","Report is "&K{r}&" months old: request a bridge letter",'
        f'IF(G{r}>0,"Follow up on relevant exceptions",'
        f'IF(M{r}>0,"Map remaining user entity controls","Acceptable")))))')
    so.row_dimensions[r].height = 40
for col, w in zip("ABCDEFGHIJKLMN", (24, 8, 11, 12, 13, 11, 12, 11, 13, 13, 11, 9, 8, 50)):
    so.column_dimensions[col].width = w
so.freeze_panes = "B5"
dvo = DataValidation(type="list", formula1='"Unqualified,Qualified,Adverse,Not provided"')
dvt = DataValidation(type="list", formula1='"Type II,Type I,None"')
so.add_data_validation(dvo)
so.add_data_validation(dvt)
dvo.add(f"E5:E{4 + NV}")
dvt.add(f"C5:C{4 + NV}")
so.conditional_formatting.add(f"L5:L{4 + NV}", CellIsRule(operator="equal", formula=['"Stale"'], fill=RED))
so.conditional_formatting.add(f"N5:N{4 + NV}", FormulaRule(formula=['OR(LEFT(N5,8)="Required",LEFT(N5,8)="Escalate",LEFT(N5,6)="Report")'], fill=RED))

# ---------------------------------------------------------------- Decisions
dc = wb.create_sheet("Decisions")
title(dc, "Vendor Decisions", "Residual risk combines tier and control strength using the matrix on Scale. Everything but the actions is a formula.")
head(dc, 4, ["Vendor", "Tier", "Inherent score", "Control score", "Control rating", "Residual risk", "SOC 2 conclusion",
             "Last assessed", "Next due", "Reassessment", "Decision", "Required actions"])
for i, v in enumerate(V):
    r = 5 + i
    qcol = L(5 + i)
    put(dc, r, 1, f"=Tiering!A{r}", bold=True)
    put(dc, r, 2, f"=Tiering!H{r}", center=True)
    put(dc, r, 3, f"=Tiering!G{r}", center=True).number_format = "0.00"
    put(dc, r, 4, f"=Questionnaire!{qcol}{SC_ROW}", center=True).number_format = "0%"
    put(dc, r, 5, f"=Questionnaire!{qcol}{RT_ROW}", center=True)
    put(dc, r, 6, f'=IFERROR(INDEX(Scale!$B$22:$D$24,MATCH(B{r},Scale!$A$22:$A$24,0),MATCH(E{r},Scale!$A$17:$A$19,0)),"Not assessed")',
        center=True, bold=True)
    put(dc, r, 7, f"='SOC2 Review'!N{r}")
    put(dc, r, 8, v[6], inp=True, center=True).number_format = "yyyy-mm-dd"
    put(dc, r, 9, f"=EDATE(H{r},Tiering!I{r})", center=True).number_format = "yyyy-mm-dd"
    put(dc, r, 10, f'=IF(I{r}<{ASOF},"OVERDUE","Current")', center=True, bold=True)
    put(dc, r, 11,
        f'=IF(F{r}="Not assessed","Approve: Tier 3, attestation only",IF(F{r}="Critical","Do not onboard or renew until critical gaps are fixed",'
        f'IF(F{r}="High","Approve with conditions and a remediation deadline",IF(F{r}="Moderate","Approve and monitor","Approve"))))', bold=True)
    put(dc, r, 12, v[7], inp=True)
    dc.row_dimensions[r].height = 56
for col, w in zip("ABCDEFGHIJKL", (24, 8, 10, 10, 12, 12, 40, 12, 12, 13, 36, 58)):
    dc.column_dimensions[col].width = w
dc.freeze_panes = "B5"
rr = f"F5:F{4 + NV}"
for word, fill in (("Critical", RED), ("High", ORG), ("Moderate", AMB), ("Low", GRN), ("Not assessed", GRY)):
    dc.conditional_formatting.add(rr, CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))
dc.conditional_formatting.add(f"J5:J{4 + NV}", CellIsRule(operator="equal", formula=['"OVERDUE"'], fill=RED))
for word, fill in (("Tier 1", RED), ("Tier 2", AMB), ("Tier 3", GRN)):
    dc.conditional_formatting.add(f"B5:B{4 + NV}", CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))

b = 5 + NV + 2
dc.cell(row=b, column=1, value="Portfolio summary").font = f(True, NAVY, 11)
head(dc, b + 1, ["Measure", "Count"])
summ = [("Vendors assessed", f"=COUNTA(A5:A{4 + NV})"),
        ("Tier 1 vendors", f'=COUNTIF(B5:B{4 + NV},"Tier 1")'),
        ("Critical residual risk", f'=COUNTIF(F5:F{4 + NV},"Critical")'),
        ("High residual risk", f'=COUNTIF(F5:F{4 + NV},"High")'),
        ("Moderate residual risk", f'=COUNTIF(F5:F{4 + NV},"Moderate")'),
        ("Low residual risk", f'=COUNTIF(F5:F{4 + NV},"Low")'),
        ("Overdue reassessments", f'=COUNTIF(J5:J{4 + NV},"OVERDUE")'),
        ("Overdue Tier 1 reassessments", f'=COUNTIFS(J5:J{4 + NV},"OVERDUE",B5:B{4 + NV},"Tier 1")')]
for i, (k, fm) in enumerate(summ):
    put(dc, b + 2 + i, 1, k, bold=True)
    put(dc, b + 2 + i, 2, fm, center=True)

wb._sheets = [wb[n] for n in ["Scope", "Tiering", "Questionnaire", "SOC2 Review", "Decisions", "Scale"]]
wb.active = 4
wb.save(OUT)
print("saved", OUT)
