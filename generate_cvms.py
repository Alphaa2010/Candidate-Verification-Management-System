"""Generate the Candidate Verification Management System workbook.

All candidate details are fictional demonstration data. Verification decisions
must be entered by an authorised human operator.
"""
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = Path(__file__).with_name("Candidate_Verification_Management_System.xlsx")
NAVY = "17365D"
BLUE = "2F75B5"
PALE = "EAF1F8"
LIGHT = "F4F6F8"
WHITE = "FFFFFF"
TEXT = "243447"
GREEN = "E2F0D9"
YELLOW = "FFF2CC"
RED = "FCE4D6"
GREY = "E7E6E6"
TEAL = "DDEBF7"
START_ROW, END_ROW = 2, 201

cases = [
    ["CV-1001", "DEMO-CAND-001", "Forensic Analyst", date(2026, 10, 1), "Aarav Mehta", "Verified", "Verified", "Verified", "Verified", None, date(2026, 10, 7), "Medium", date(2026, 10, 7), "Fictional demonstration record; all checks complete."],
    ["CV-1002", "DEMO-CAND-002", "Risk Advisory Analyst", date(2026, 10, 2), "Diya Nair", "Verified", "Pending", "Verified", "Not Started", None, None, "High", date(2026, 10, 8), "Education evidence awaiting review."],
    ["CV-1003", "DEMO-CAND-003", "Forensic Data Analyst", date(2026, 10, 3), "Kabir Shah", "Verified", "Rejected", "Pending", "Not Started", None, None, "Urgent", date(2026, 10, 9), "Education discrepancy; authorised reviewer follow-up."],
    ["CV-1004", "DEMO-CAND-004", "Investigations Associate", date(2026, 10, 4), "Aarav Mehta", "Not Started", "Not Started", "Not Started", "Not Started", None, None, "Low", date(2026, 10, 4), "Newly assigned fictional case."],
    ["CV-1005", "DEMO-CAND-005", "Compliance Analyst", date(2026, 10, 5), "Mira Iyer", "Verified", "Verified", "Pending", "Verified", None, None, "Medium", date(2026, 10, 10), "Address check in progress."],
    ["CV-1006", "DEMO-CAND-006", "Forensic Analyst", date(2026, 10, 6), "Diya Nair", "Rejected", "Verified", "Verified", "Pending", None, None, "Urgent", date(2026, 10, 11), "Identity exception flagged for manual review."],
    ["CV-1007", "DEMO-CAND-007", "Risk Advisory Analyst", date(2026, 10, 7), "Kabir Shah", "Verified", "Verified", "Verified", "Verified", None, date(2026, 10, 14), "High", date(2026, 10, 14), "Fictional demonstration record; all checks complete."],
    ["CV-1008", "DEMO-CAND-008", "Investigations Associate", date(2026, 10, 8), "Mira Iyer", "Pending", "Not Started", "Pending", "Not Started", None, None, "High", date(2026, 10, 12), "Awaiting candidate documentation."],
    ["CV-1009", "DEMO-CAND-009", "Compliance Analyst", date(2026, 10, 9), "Aarav Mehta", "Verified", "Verified", "Verified", "Verified", None, date(2026, 10, 15), "Low", date(2026, 10, 15), "Fictional demonstration record; all checks complete."],
    ["CV-1010", "DEMO-CAND-010", "Forensic Data Analyst", date(2026, 10, 10), "Diya Nair", "Not Started", "Pending", "Not Started", "Pending", None, None, "Medium", date(2026, 10, 13), "Checks not yet completed."],
    ["CV-1011", "DEMO-CAND-011", "Investigations Associate", date(2026, 10, 11), "Kabir Shah", "Verified", "Verified", "Verified", "Pending", None, None, "Medium", date(2026, 10, 16), "Reference confirmation outstanding."],
    ["CV-1012", "DEMO-CAND-012", "Compliance Analyst", date(2026, 10, 12), "Mira Iyer", "Verified", "Pending", "Pending", "Verified", None, None, "High", date(2026, 10, 17), "Two checks remain in progress."],
]
headers = ["Case ID", "Candidate Reference", "Job Role", "Assigned Date", "Case Owner", "Identity Verification", "Education Verification", "Address Verification", "Reference Verification", "Overall Status", "Completion Date", "Priority", "Last Updated", "Remarks"]

def set_title(ws, cell_range, text):
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    cell.value = text
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.font = Font(name="Aptos Display", size=19, bold=True, color=WHITE)
    cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[cell.row].height = 34

def style_header(cell):
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.font = Font(name="Aptos", bold=True, color=WHITE, size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def build_readme(wb):
    ws = wb.active
    ws.title = "README - Instructions"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 88
    set_title(ws, "B2:C2", "Candidate Verification Management System")
    ws.merge_cells("B3:C3")
    ws["B3"] = "CVMS  |  Excel workflow tracker and management dashboard"
    ws["B3"].font = Font(italic=True, color=BLUE, size=11)
    sections = [
        (5, "Purpose", "Manage candidate verification cases from assignment through completion. Track four checks, ownership, priority, exceptions, and productivity at a glance."),
        (8, "How to add a case", "On VERIFICATION TRACKER, enter a unique Case ID and candidate reference, role, assigned date, owner, priority, and remarks in the next blank table row. Select check statuses from dropdowns. Copy the Overall Status formula into a new row if Excel does not extend it automatically."),
        (11, "Updating verification statuses", "Only an authorised human operator should enter Not Started, Pending, Verified, or Rejected for Identity, Education, Address, and Reference. Update Last Updated when a case changes. Enter a Completion Date when all four checks are verified."),
        (14, "Automatic status calculation", 'A case is Review Required if any check is Rejected. It is Completed when all four checks are Verified. All other populated cases are In Progress. An empty Case ID keeps the Overall Status blank.'),
        (17, "Dashboard", "KPI cards summarize case volume, completion, review workload, and progress. Charts show status, verification outcomes, priority, and case owner. Dashboard sources cover rows 2–201 and update as those tracker records change."),
        (20, "Colour legend", "Verified / Completed = green   |   Pending / In Progress = yellow   |   Rejected / Review Required = red   |   Not Started = grey   |   Input fields = white"),
        (23, "Limitations and privacy", "CVMS is workflow management and reporting only. It does not authenticate identity documents, contact universities or references, search criminal records, or conduct background checks. Results must be entered by an authorised operator. Sample records are fictional. Do not enter unnecessary sensitive personal data; follow your organisation's privacy, retention, access, and review policies."),
    ]
    for row, heading, body in sections:
        ws.cell(row, 2, heading).font = Font(bold=True, color=NAVY, size=12)
        ws.cell(row, 2).fill = PatternFill("solid", fgColor=PALE)
        ws.merge_cells(start_row=row, start_column=3, end_row=row+1, end_column=3)
        ws.cell(row, 3, body).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(row, 3).font = Font(color=TEXT, size=10)
        ws.row_dimensions[row].height = 28
        ws.row_dimensions[row+1].height = 24
    ws["B27"] = "WORKFLOW"
    ws["B27"].font = Font(bold=True, color=NAVY, size=12)
    ws.merge_cells("B28:C28")
    ws["B28"] = "Case Assigned  →  Verification Started  →  Checks Updated  →  Overall Status Calculated  →  Case Completed / Review Required"
    ws["B28"].fill = PatternFill("solid", fgColor=PALE)
    ws["B28"].font = Font(bold=True, color=NAVY, size=11)
    ws["B28"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[28].height = 32
    ws.freeze_panes = "B5"

def build_tracker(wb):
    ws = wb.create_sheet("Verification Tracker")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 34
    for col, name in enumerate(headers, 1):
        c = ws.cell(1, col, name); style_header(c)
    for row_idx in range(START_ROW, END_ROW + 1):
        values = cases[row_idx-2] if row_idx-2 < len(cases) else [None] * len(headers)
        for col, value in enumerate(values, 1):
            c = ws.cell(row_idx, col, value)
            c.font = Font(name="Aptos", size=10, color=TEXT)
            c.alignment = Alignment(vertical="center", wrap_text=(col == 14))
            if row_idx % 2 == 0:
                c.fill = PatternFill("solid", fgColor="F8FAFC")
        ws.cell(row_idx, 10, f'=IF(A{row_idx}="","",IF(COUNTIF(F{row_idx}:I{row_idx},"Rejected")>0,"Review Required",IF(COUNTIF(F{row_idx}:I{row_idx},"Verified")=4,"Completed","In Progress")))')
        for col in (4, 11, 13):
            ws.cell(row_idx, col).number_format = "dd-mmm-yyyy"
        ws.row_dimensions[row_idx].height = 22
    widths = [15, 23, 27, 16, 18, 22, 23, 21, 23, 20, 17, 13, 16, 52]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1, col).column_letter].width = width
    tab = Table(displayName="CVMS_Cases", ref=f"A1:N{END_ROW}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws.add_table(tab)
    check_dv = DataValidation(type="list", formula1='"Not Started,Pending,Verified,Rejected"', allow_blank=True)
    check_dv.errorTitle = "Choose a listed status"; check_dv.error = "Select a status from the dropdown."
    check_dv.promptTitle = "Human-entered result"; check_dv.prompt = "Enter the result recorded by an authorised operator."
    ws.add_data_validation(check_dv)
    for col in "FGHI": check_dv.add(f"{col}{START_ROW}:{col}{END_ROW}")
    priority_dv = DataValidation(type="list", formula1='"Low,Medium,High,Urgent"', allow_blank=True)
    ws.add_data_validation(priority_dv); priority_dv.add(f"L{START_ROW}:L{END_ROW}")
    id_dv = DataValidation(type="custom", formula1=f'=OR(A{START_ROW}="",COUNTIF($A${START_ROW}:$A${END_ROW},A{START_ROW})=1)', allow_blank=True)
    id_dv.errorTitle = "Duplicate Case ID"; id_dv.error = "Case IDs must be unique."
    id_dv.showErrorMessage = True
    ws.add_data_validation(id_dv); id_dv.add(f"A{START_ROW}:A{END_ROW}")
    for col in "FGHI":
        rng = f"{col}{START_ROW}:{col}{END_ROW}"
        for value, color in [("Verified", GREEN), ("Pending", YELLOW), ("Rejected", RED), ("Not Started", GREY)]:
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{col}{START_ROW}="{value}"'], fill=PatternFill("solid", fgColor=color)))
    for value, color in [("Completed", GREEN), ("In Progress", YELLOW), ("Review Required", RED)]:
        ws.conditional_formatting.add(f"J{START_ROW}:J{END_ROW}", FormulaRule(formula=[f'J{START_ROW}="{value}"'], fill=PatternFill("solid", fgColor=color)))
    # The Excel table owns its filter buttons and AutoFilter definition.
    # Do not also attach a worksheet-level AutoFilter to the same range; Excel
    # treats the duplicate overlapping definitions as corrupt table content.

def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard")
    ws.sheet_view.showGridLines = False
    for col in range(1, 15): ws.column_dimensions[ws.cell(1, col).column_letter].width = 15
    set_title(ws, "B2:M2", "Candidate Verification | Management Dashboard")
    ws.merge_cells("B3:M3"); ws["B3"] = "Fictional demonstration data  •  Human-entered verification outcomes  •  Snapshot updates from Verification Tracker rows 2–201"
    ws["B3"].font = Font(color="64748B", italic=True, size=9)
    cards = [("B5:D5", "B6:D7", "Total Cases", f'=COUNTIF(\'Verification Tracker\'!$A$2:$A${END_ROW},"<>")', BLUE),
             ("E5:G5", "E6:G7", "Completed Cases", f'=COUNTIF(\'Verification Tracker\'!$J$2:$J${END_ROW},"Completed")', "548235"),
             ("H5:J5", "H6:J7", "In Progress Cases", f'=COUNTIF(\'Verification Tracker\'!$J$2:$J${END_ROW},"In Progress")', "BF9000"),
             ("K5:M5", "K6:M7", "Review Required", f'=COUNTIF(\'Verification Tracker\'!$J$2:$J${END_ROW},"Review Required")', "C00000"),
             ("B9:D9", "B10:D11", "Not Started Cases", f'=COUNTIFS(\'Verification Tracker\'!$A$2:$A${END_ROW},"<>",\'Verification Tracker\'!$F$2:$F${END_ROW},"Not Started",\'Verification Tracker\'!$G$2:$G${END_ROW},"Not Started",\'Verification Tracker\'!$H$2:$H${END_ROW},"Not Started",\'Verification Tracker\'!$I$2:$I${END_ROW},"Not Started")', "7F7F7F"),
             ("E9:G9", "E10:G11", "Completion Rate", '=IFERROR(E6/B6,0)', BLUE)]
    for head_rng, value_rng, label, formula, color in cards:
        ws.merge_cells(head_rng); ws.merge_cells(value_rng)
        h = ws[head_rng.split(":")[0]]; v = ws[value_rng.split(":")[0]]
        h.value = label; h.fill = PatternFill("solid", fgColor=color); h.font = Font(bold=True, color=WHITE, size=10); h.alignment = Alignment(horizontal="center", vertical="center")
        v.value = formula; v.fill = PatternFill("solid", fgColor=LIGHT); v.font = Font(bold=True, color=NAVY, size=22); v.alignment = Alignment(horizontal="center", vertical="center")
        for row in range(ws[head_rng.split(":")[0]].row, ws[head_rng.split(":")[0]].row+1): ws.row_dimensions[row].height = 23
        if label == "Completion Rate": v.number_format = "0.0%"
    ws.row_dimensions[6].height = 27; ws.row_dimensions[7].height = 18; ws.row_dimensions[10].height = 27; ws.row_dimensions[11].height = 18
    # Chart source tables, kept visible for transparency.
    source_col = 16
    for col in range(source_col, source_col+8): ws.column_dimensions[ws.cell(1, col).column_letter].hidden = True
    def put_table(col, title, labels, formulas):
        ws.cell(1, col, title)
        for i, label in enumerate(labels, 2):
            ws.cell(i, col, label)
            ws.cell(i, col+1, formulas[i-2])
    put_table(16, "Overall Status", ["Completed", "In Progress", "Review Required"], [f'=COUNTIF(\'Verification Tracker\'!$J$2:$J${END_ROW},P2)', f'=COUNTIF(\'Verification Tracker\'!$J$2:$J${END_ROW},P3)', f'=COUNTIF(\'Verification Tracker\'!$J$2:$J${END_ROW},P4)'])
    put_table(19, "Verification Status", ["Not Started", "Pending", "Verified", "Rejected"], [f'=COUNTIF(\'Verification Tracker\'!$F$2:$I${END_ROW},S{i})' for i in range(2, 6)])
    put_table(22, "Priority", ["Low", "Medium", "High", "Urgent"], [f'=COUNTIF(\'Verification Tracker\'!$L$2:$L${END_ROW},V2)', f'=COUNTIF(\'Verification Tracker\'!$L$2:$L${END_ROW},V3)', f'=COUNTIF(\'Verification Tracker\'!$L$2:$L${END_ROW},V4)', f'=COUNTIF(\'Verification Tracker\'!$L$2:$L${END_ROW},V5)'])
    owners = sorted({c[4] for c in cases})
    put_table(25, "Cases by Owner", owners, [f'=COUNTIF(\'Verification Tracker\'!$E$2:$E${END_ROW},Y{i})' for i in range(2, 2+len(owners))])
    def chart(title, col, n, kind="doughnut", anchor="B14"):
        C = DoughnutChart if kind == "doughnut" else BarChart
        ch = C(); ch.title = title; ch.style = 10; ch.height = 7.2; ch.width = 13.5
        # Chart source tables live in hidden helper columns; explicitly include
        # hidden cells so Excel still plots the dashboard series.
        ch.visible_cells_only = False
        ch.add_data(Reference(ws, min_col=col+1, min_row=1, max_row=n+1), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=col, min_row=2, max_row=n+1))
        ch.legend.position = "b"
        if kind == "doughnut": ch.holeSize = 62; ch.dataLabels = DataLabelList(); ch.dataLabels.showPercent = True
        else: ch.type = "bar"; ch.y_axis.title = "Cases"; ch.x_axis.title = "Category"
        ws.add_chart(ch, anchor)
    chart("Overall Case Status Distribution", 16, 3, "doughnut", "B14")
    chart("Verification Status Distribution", 19, 4, "bar", "H14")
    chart("Cases by Priority", 22, 4, "bar", "B29")
    chart("Cases by Case Owner", 25, len(owners), "bar", "H29")
    ws.freeze_panes = "B5"

def build_summary(wb):
    ws = wb.create_sheet("Case Summary")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4; ws.column_dimensions["B"].width = 29; ws.column_dimensions["C"].width = 42; ws.column_dimensions["D"].width = 4
    set_title(ws, "B2:C2", "Case Summary")
    ws["B4"] = "Enter Case ID"; ws["B4"].font = Font(bold=True, color=NAVY)
    ws["C4"] = "CV-1001"; ws["C4"].fill = PatternFill("solid", fgColor=YELLOW); ws["C4"].font = Font(bold=True, color=NAVY)
    dv = DataValidation(type="list", formula1="=CaseIDs", allow_blank=True)
    ws.add_data_validation(dv); dv.add("C4")
    rows = [(6,"Candidate Reference",2),(7,"Job Role",3),(8,"Case Owner",5),(9,"Assigned Date",4),(11,"Identity Verification",6),(12,"Education Verification",7),(13,"Address Verification",8),(14,"Reference Verification",9),(16,"Overall Status",10),(17,"Priority",12),(18,"Completion Date",11),(20,"Remarks",14)]
    for r, label, source_col in rows:
        ws.cell(r, 2, label).font = Font(bold=True, color=NAVY)
        col_letter = ws.cell(1, source_col).column_letter
        ws.cell(r, 3, f'=IFERROR(INDEX(\'Verification Tracker\'!${col_letter}$2:${col_letter}${END_ROW},MATCH($C$4,\'Verification Tracker\'!$A$2:$A${END_ROW},0)),"")')
        ws.cell(r, 3).fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="center")
        if source_col in (4, 11): ws.cell(r, 3).number_format = "dd-mmm-yyyy"
        ws.row_dimensions[r].height = 26 if r != 20 else 54
    ws["B22"] = "Select a Case ID above to view its current tracker details."
    ws["B22"].font = Font(italic=True, color="64748B", size=9)
    ws.freeze_panes = "B4"

def main():
    wb = Workbook()
    build_readme(wb)
    build_tracker(wb)
    build_dashboard(wb)
    build_summary(wb)
    wb.defined_names.add(DefinedName("CaseIDs", attr_text=f"'Verification Tracker'!$A$2:$A${END_ROW}"))
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.active = 2
    wb.save(OUT)
    print(f"Created {OUT}")

if __name__ == "__main__":
    main()
