# Candidate Verification Management System (CVMS)

CVMS is an Excel-based workflow tracker for a verification team. It records case ownership, priority, four human-entered verification outcomes, case status, and summary reporting. It is designed as a fictional demonstration project for an entry-level analyst interview.

## Files

- `Candidate_Verification_Management_System.xlsx` — generated Excel application
- `generate_cvms.py` — reproducible workbook generator
- `requirements.txt` — Python dependency list

## Generate the workbook

Use Python 3.9 or newer, then run:

```powershell
pip install -r requirements.txt
python generate_cvms.py
```

The script writes the workbook next to itself and does not call paid APIs, external services, or databases.

## Use the workbook

1. Open `Candidate_Verification_Management_System.xlsx` in Microsoft Excel.
2. Start at **README - Instructions** for the workflow and colour legend.
3. In **Verification Tracker**, add a unique Case ID and fictional candidate reference, role, assigned date, owner, priority, and remarks in a blank row.
4. Use dropdowns for the four verification checks. The authorised human operator records the actual outcome and updates Last Updated; the workbook does not verify evidence.
5. Review the automatically calculated Overall Status, then check **Dashboard** for workload and **Case Summary** for a single case view.

The tracker includes formulas and dropdowns through row 201. If adding rows beyond that range, extend the table, formulas, validations, and dashboard source ranges in Excel or adjust `END_ROW` in the generator and regenerate.

## Status rules

- Any `Rejected` check → `Review Required`
- Four `Verified` checks → `Completed`
- Otherwise, a populated case → `In Progress`
- Blank Case ID → blank Overall Status

Completion Date is manually entered when appropriate; it does not drive Overall Status.

## Privacy and scope

All sample records, names, roles, IDs, and case owners are fictional. CVMS is workflow management and reporting only. It does not authenticate identity documents, contact universities or references, search criminal records, or perform real background checks. Use only under appropriate organisational access, privacy, retention, and review policies. Avoid unnecessary sensitive personal data.

## Compatibility note

The workbook uses standard Excel formulas (COUNTIF/COUNTIFS, INDEX/MATCH), tables, data validation, conditional formatting, and charts. Formulas are configured for automatic recalculation in Excel. Formula results are not stored by openpyxl, so a compatible spreadsheet application should recalculate them when the workbook is opened.
