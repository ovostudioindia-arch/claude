#!/usr/bin/env python3
"""
Build the 50/30/20 Budget Tracker as an .xlsx file.
Upload the output to Google Drive → Open as Google Sheets.

All formulas use Google Sheets syntax so they work after conversion.
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.chart import BarChart, PieChart, Reference as ChartRef
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.protection import SheetProtection
from copy import copy

# ── Colors ──────────────────────────────────────────────────────────
DARK_SAGE   = 'FF4A5E52'
DUSTY_ROSE  = 'FFC49A8A'
WARM_BEIGE  = 'FFF0EBE3'
SOFT_SAGE   = 'FFD6DDD8'
WHITE_      = 'FFFFFFFF'
DARK_TEXT   = 'FF2C2C2C'
LIGHT_YELLOW = 'FFFFF9C4'
LIGHT_GREEN = 'FFD6F0DC'
LIGHT_RED   = 'FFF5C6C0'
RED_TEXT     = 'FFC0392B'
SAGE_LIGHT  = 'FF8FA898'
BORDER_CLR  = 'FFC8C8C8'

def fill(color):
    return PatternFill(start_color=color, end_color=color, fill_type='solid')

def font(color='FF000000', bold=False, size=10):
    return Font(color=color, bold=bold, size=size)

thin_border_side = Side(style='thin', color=BORDER_CLR)
thin_border = Border(
    left=thin_border_side, right=thin_border_side,
    top=thin_border_side, bottom=thin_border_side
)

CENTER = Alignment(horizontal='center', vertical='center')
LEFT   = Alignment(horizontal='left', vertical='center')
WRAP   = Alignment(horizontal='center', vertical='center', wrap_text=True)

# ── Helpers ─────────────────────────────────────────────────────────
def style_range(ws, row_start, col_start, row_end, col_end,
                bg=None, ft=None, alignment=None, border=None, num_fmt=None):
    for r in range(row_start, row_end + 1):
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=r, column=c)
            if bg:        cell.fill = fill(bg)
            if ft:        cell.font = ft
            if alignment: cell.alignment = alignment
            if border:    cell.border = border
            if num_fmt:   cell.number_format = num_fmt

def merge_and_set(ws, range_str, value, bg=None, ft=None, alignment=CENTER):
    ws.merge_cells(range_str)
    top_left = range_str.split(':')[0]
    cell = ws[top_left]
    cell.value = value
    if bg: cell.fill = fill(bg)
    if ft: cell.font = ft
    cell.alignment = alignment


# ====================================================================
#  REFERENCE TAB
# ====================================================================
def build_reference(wb):
    ws = wb.create_sheet('Reference', 3)

    # A:B — Sub-Category → Category
    cat_lookup = [
        ('Groceries','Expenses'),('Gas','Expenses'),('Health','Expenses'),
        ('Dining Out','Expenses'),('Clothes','Expenses'),('Drugstore','Expenses'),
        ('Transport','Expenses'),('Entertainment','Expenses'),('Other','Expenses'),
        ('Rent/Mortgage','Bills'),('Utilities','Bills'),('Insurance','Bills'),
        ('Phone','Bills'),('Internet','Bills'),
        ('Streaming','Subscriptions'),('Gym','Subscriptions'),
        ('Loan Payment','Debts'),('Credit Card','Debts'),
        ('401k','Savings'),('Savings Transfer','Savings'),
    ]
    for i, (sub, cat) in enumerate(cat_lookup, 1):
        ws.cell(row=i, column=1, value=sub)
        ws.cell(row=i, column=2, value=cat)

    # D:E — Sub-Category → Bucket
    bucket_lookup = [
        ('Groceries','Need'),('Gas','Need'),('Health','Need'),
        ('Rent/Mortgage','Need'),('Utilities','Need'),('Insurance','Need'),
        ('Phone','Need'),('Internet','Need'),
        ('Dining Out','Want'),('Clothes','Want'),('Drugstore','Want'),
        ('Entertainment','Want'),('Transport','Want'),('Streaming','Want'),
        ('Gym','Want'),('Other','Want'),
        ('Loan Payment','Savings / Debts'),('Credit Card','Savings / Debts'),
        ('401k','Savings / Debts'),('Savings Transfer','Savings / Debts'),
    ]
    for i, (sub, bucket) in enumerate(bucket_lookup, 1):
        ws.cell(row=i, column=4, value=sub)
        ws.cell(row=i, column=5, value=bucket)

    # G — Month names
    months = ['January','February','March','April','May','June',
              'July','August','September','October','November','December']
    for i, m in enumerate(months, 1):
        ws.cell(row=i, column=7, value=m)

    ws.sheet_state = 'hidden'
    return ws


# ====================================================================
#  BUDGET SETUP TAB
# ====================================================================
def build_budget_setup(wb):
    ws = wb.create_sheet('Budget Setup', 2)

    # Section 1 — Income & Ratios
    merge_and_set(ws, 'A1:B1', 'INCOME & RATIO SETTINGS',
                  bg=DARK_SAGE, ft=font(WHITE_, bold=True))

    labels = [('Monthly Income', 5000), ('Need %', 50), ('Want %', 30),
              ('Savings/Debts %', 20)]
    for i, (label, val) in enumerate(labels, 2):
        ws.cell(row=i, column=1, value=label)
        c = ws.cell(row=i, column=2, value=val)
        c.fill = fill(LIGHT_YELLOW)
    ws['B2'].number_format = '$#,##0.00'

    ws['A6'] = 'Ratio Check'
    ws['B6'] = '=IF(B3+B4+B5=100,"✓ OK","⚠ Must = 100")'

    # Section 2 — Fixed Monthly Bills
    for col, hdr in [(4,'FIXED ITEM'),(5,'AMOUNT'),(6,'CATEGORY')]:
        c = ws.cell(row=8, column=col, value=hdr)
        c.fill = fill(DUSTY_ROSE); c.font = font(WHITE_, bold=True)

    bills = [
        ('Rent/Mortgage',1200,'Bills'),('Phone',55,'Bills'),
        ('Internet',45,'Bills'),('Utilities',80,'Bills'),
        ('Streaming',45,'Subscriptions'),('Gym',60,'Subscriptions'),
        ('Loan Payment',500,'Debts'),
    ]
    for i, (item, amt, cat) in enumerate(bills, 9):
        ws.cell(row=i, column=4, value=item)
        ws.cell(row=i, column=5, value=amt).number_format = '$#,##0.00'
        ws.cell(row=i, column=6, value=cat)

    # Data validation for column F
    dv = DataValidation(type='list', formula1='"Bills,Subscriptions,Debts"', allow_blank=True)
    dv.error = 'Pick Bills, Subscriptions, or Debts'
    ws.add_data_validation(dv)
    dv.add('F9:F30')

    # Column widths
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 13
    ws.column_dimensions['F'].width = 16

    ws.freeze_panes = 'A2'
    return ws


# ====================================================================
#  TRANSACTIONS TAB
# ====================================================================
def build_transactions(wb):
    ws = wb.create_sheet('Transactions', 1)

    # Headers
    headers = ['DATE','AMOUNT','SUB-CATEGORY','CATEGORY','DESCRIPTION','BUCKET']
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = fill(DARK_SAGE)
        cell.font = font(WHITE_, bold=True)

    # Sub-category dropdown (C2:C200)
    dv = DataValidation(type='list', formula1='=Reference!$A$1:$A$20', allow_blank=True)
    dv.error = 'Pick a sub-category from the list'
    ws.add_data_validation(dv)
    dv.add('C2:C200')

    # Formulas for D (CATEGORY) and F (BUCKET) — rows 2-200
    for r in range(2, 201):
        ws.cell(row=r, column=4).value = f'=IFERROR(VLOOKUP(C{r},Reference!$A:$B,2,0),"")'
        ws.cell(row=r, column=6).value = f'=IFERROR(VLOOKUP(C{r},Reference!$D:$E,2,0),"")'

    # ── Seed data ────────────────────────────────────────────────────
    seed = [
        # JANUARY (20)
        ('1/1/2024',1200,'Rent/Mortgage','Monthly rent'),
        ('1/1/2024',80,'Utilities','Electric bill'),
        ('1/2/2024',100,'Groceries','Weekly shop'),
        ('1/3/2024',45,'Internet',''),
        ('1/4/2024',55,'Phone',''),
        ('1/5/2024',500,'Loan Payment','Car loan'),
        ('1/5/2024',60,'Gym',''),
        ('1/6/2024',45,'Streaming','Netflix + Spotify'),
        ('1/7/2024',150,'Groceries',''),
        ('1/8/2024',90,'Dining Out','Birthday dinner'),
        ('1/10/2024',200,'Clothes','Winter jacket'),
        ('1/12/2024',120,'Gas',''),
        ('1/14/2024',75,'Drugstore',''),
        ('1/15/2024',250,'Savings Transfer','Monthly savings'),
        ('1/16/2024',40,'Entertainment','Cinema'),
        ('1/18/2024',500,'Other','Boat repair'),
        ('1/20/2024',150,'Health','Dentist'),
        ('1/22/2024',50,'Transport','Uber rides'),
        ('1/25/2024',130,'Dining Out',''),
        ('1/28/2024',15,'Clothes',''),
        # FEBRUARY (10)
        ('2/1/2024',1200,'Rent/Mortgage',''),
        ('2/2/2024',80,'Utilities',''),
        ('2/3/2024',120,'Groceries',''),
        ('2/5/2024',45,'Streaming',''),
        ('2/7/2024',200,'Clothes',"Valentine's gift"),
        ('2/10/2024',75,'Dining Out',''),
        ('2/14/2024',500,'Loan Payment',''),
        ('2/15/2024',300,'Savings Transfer',''),
        ('2/18/2024',55,'Phone',''),
        ('2/22/2024',90,'Gas',''),
        # MARCH (10)
        ('3/1/2024',1200,'Rent/Mortgage',''),
        ('3/3/2024',95,'Groceries',''),
        ('3/5/2024',500,'Loan Payment',''),
        ('3/8/2024',60,'Gym',''),
        ('3/10/2024',130,'Clothes',''),
        ('3/12/2024',45,'Streaming',''),
        ('3/15/2024',250,'Savings Transfer',''),
        ('3/18/2024',110,'Dining Out',''),
        ('3/20/2024',80,'Utilities',''),
        ('3/25/2024',55,'Phone',''),
    ]

    from datetime import datetime
    for i, (date_str, amt, subcat, desc) in enumerate(seed, 2):
        dt = datetime.strptime(date_str, '%m/%d/%Y')
        ws.cell(row=i, column=1, value=dt).number_format = 'M/D/YYYY'
        ws.cell(row=i, column=2, value=amt).number_format = '$#,##0.00'
        ws.cell(row=i, column=3, value=subcat)
        ws.cell(row=i, column=5, value=desc)

    # Format columns A and B for empty rows too
    for r in range(42, 201):
        ws.cell(row=r, column=1).number_format = 'M/D/YYYY'
        ws.cell(row=r, column=2).number_format = '$#,##0.00'

    # Alternating row colors
    for r in range(2, 42):
        bg = WHITE_ if r % 2 == 0 else WARM_BEIGE
        style_range(ws, r, 1, r, 6, bg=bg)

    # Column widths
    widths = [13, 12, 18, 14, 24, 17]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Conditional formatting: large transactions
    ws.conditional_formatting.add('B2:B200',
        CellIsRule(operator='greaterThan', formula=['300'],
                   fill=fill(WARM_BEIGE), font=Font(bold=True)))

    ws.freeze_panes = 'A2'

    # Protect formula columns D and F (warning only — not locked hard)
    # We'll just set them to a subtle style to indicate auto-fill
    for r in range(2, 201):
        ws.cell(row=r, column=4).font = Font(color='FF888888', italic=True)
        ws.cell(row=r, column=6).font = Font(color='FF888888', italic=True)

    return ws


# ====================================================================
#  DASHBOARD TAB
# ====================================================================
def build_dashboard(wb):
    ws = wb.active
    ws.title = 'Dashboard'

    # ── ROW 1-2: Title banner ────────────────────────────────────────
    merge_and_set(ws, 'A1:I2', '50/30/20 DASHBOARD',
                  bg=WARM_BEIGE, ft=font(DARK_SAGE, bold=True, size=20))

    # ── ROW 3: Month selector label ──────────────────────────────────
    merge_and_set(ws, 'A3:I3', 'WHICH MONTH DO YOU WANT TO IMPROVE?',
                  bg=DUSTY_ROSE, ft=font(WHITE_, bold=True, size=10))

    # ── ROW 4: Month selector ────────────────────────────────────────
    ws['H4'] = 'SELECT MONTH →'
    ws['H4'].font = font(bold=True)
    ws['H4'].alignment = Alignment(horizontal='right', vertical='center')
    ws['I4'] = 'January'
    ws['I4'].font = font(bold=True, size=11)

    dv = DataValidation(type='list', formula1='=Reference!$G$1:$G$12', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add('I4')

    # ── ROW 6: Helper month number ───────────────────────────────────
    ws['I6'] = '=MATCH(I4,Reference!$G$1:$G$12,0)'
    ws['I6'].font = font('FF999999', size=8)

    # ── ROW 6-13: Budget Summary Table ───────────────────────────────
    # Row 6 headers
    for col, hdr in [(1,'SOURCE'),(3,'BUDGET'),(5,'REAL')]:
        c = ws.cell(row=6, column=col, value=hdr)
        c.fill = fill(DUSTY_ROSE); c.font = font(WHITE_, bold=True)
    # Fill entire row 6 header bg
    style_range(ws, 6, 1, 6, 5, bg=DUSTY_ROSE, ft=font(WHITE_, bold=True))

    # Data rows
    summary_data = [
        (7, 'INCOME',
         "='Budget Setup'!B2",
         "='Budget Setup'!B2"),
        (8, 'BILLS',
         '=IFERROR(SUMIFS(\'Budget Setup\'!E:E,\'Budget Setup\'!F:F,"Bills"),0)',
         '=IFERROR(SUMIFS(Transactions!B:B,Transactions!D:D,"Bills",MONTH(Transactions!A:A),I$6),0)'),
        (9, 'SUBSCRIPTIONS',
         '=IFERROR(SUMIFS(\'Budget Setup\'!E:E,\'Budget Setup\'!F:F,"Subscriptions"),0)',
         '=IFERROR(SUMIFS(Transactions!B:B,Transactions!D:D,"Subscriptions",MONTH(Transactions!A:A),I$6),0)'),
        (10, 'EXPENSES',
         "=IFERROR('Budget Setup'!B2*('Budget Setup'!B3/100)-C8-C9,0)",
         '=IFERROR(SUMIFS(Transactions!B:B,Transactions!D:D,"Expenses",MONTH(Transactions!A:A),I$6),0)'),
        (11, 'SAVINGS',
         "=IFERROR('Budget Setup'!B2*('Budget Setup'!B5/100)*0.5,0)",
         '=IFERROR(SUMIFS(Transactions!B:B,Transactions!D:D,"Savings",MONTH(Transactions!A:A),I$6),0)'),
        (12, 'DEBTS',
         "=IFERROR('Budget Setup'!B2*('Budget Setup'!B5/100)*0.5,0)",
         '=IFERROR(SUMIFS(Transactions!B:B,Transactions!D:D,"Debts",MONTH(Transactions!A:A),I$6),0)'),
        (13, 'AMOUNT LEFT', '', '=IFERROR(E7-SUM(E8:E12),0)'),
    ]
    for row, label, budget_f, real_f in summary_data:
        ws.cell(row=row, column=1, value=label).font = font(bold=(row==13))
        if budget_f:
            ws.cell(row=row, column=3).value = budget_f
        ws.cell(row=row, column=5).value = real_f

    style_range(ws, 7, 3, 13, 3, num_fmt='$#,##0.00')
    style_range(ws, 7, 5, 13, 5, num_fmt='$#,##0.00')
    ws.cell(row=13, column=5).font = font(bold=True)

    # Alternating row colors
    for r in range(7, 14):
        bg = WHITE_ if r % 2 == 1 else WARM_BEIGE
        style_range(ws, r, 1, r, 5, bg=bg)

    # Borders
    style_range(ws, 6, 1, 13, 5, border=thin_border)

    # ── ROW 15-20: Goal Breakdown ────────────────────────────────────
    merge_and_set(ws, 'A15:E15', 'MY GOAL BREAKDOWN',
                  bg=DARK_SAGE, ft=font(WHITE_, bold=True))
    style_range(ws, 16, 1, 16, 5, bg=DUSTY_ROSE, ft=font(WHITE_, bold=True))
    ws.cell(row=16, column=1, value='CATEGORY')
    ws.cell(row=16, column=3, value='%')
    ws.cell(row=16, column=5, value='$')

    goal_rows = [
        (17, 'Need',    "=TEXT('Budget Setup'!B3/100,\"0%\")", "='Budget Setup'!B2*('Budget Setup'!B3/100)"),
        (18, 'Want',    "=TEXT('Budget Setup'!B4/100,\"0%\")", "='Budget Setup'!B2*('Budget Setup'!B4/100)"),
        (19, 'Savings / Debts', "=TEXT('Budget Setup'!B5/100,\"0%\")", "='Budget Setup'!B2*('Budget Setup'!B5/100)"),
        (20, 'TOTAL',   '100%', '=SUM(E17:E19)'),
    ]
    for row, label, pct, dollar in goal_rows:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=3, value=pct)
        ws.cell(row=row, column=5, value=dollar)

    style_range(ws, 17, 5, 20, 5, num_fmt='$#,##0.00')
    style_range(ws, 17, 3, 20, 3, num_fmt='0%')
    for r in range(17, 21):
        bg = WHITE_ if r % 2 == 1 else WARM_BEIGE
        style_range(ws, r, 1, r, 5, bg=bg)
    style_range(ws, 15, 1, 20, 5, border=thin_border)

    # ── ROW 22-27: Real Breakdown ────────────────────────────────────
    merge_and_set(ws, 'A22:E22', 'MY REAL BUDGET BREAKDOWN',
                  bg=DARK_SAGE, ft=font(WHITE_, bold=True))
    style_range(ws, 23, 1, 23, 5, bg=DUSTY_ROSE, ft=font(WHITE_, bold=True))
    ws.cell(row=23, column=1, value='CATEGORY')
    ws.cell(row=23, column=3, value='%')
    ws.cell(row=23, column=5, value='$')

    real_rows = [
        (24, 'Need',
         '=IFERROR(E24/E7,0)',
         '=IFERROR(SUMIFS(Transactions!B:B,Transactions!F:F,"Need",MONTH(Transactions!A:A),I$6),0)'),
        (25, 'Want',
         '=IFERROR(E25/E7,0)',
         '=IFERROR(SUMIFS(Transactions!B:B,Transactions!F:F,"Want",MONTH(Transactions!A:A),I$6),0)'),
        (26, 'Savings / Debts',
         '=IFERROR(E26/E7,0)',
         '=IFERROR(SUMIFS(Transactions!B:B,Transactions!F:F,"Savings / Debts",MONTH(Transactions!A:A),I$6),0)'),
        (27, 'TOTAL', '=IFERROR(SUM(C24:C26),0)', '=SUM(E24:E26)'),
    ]
    for row, label, pct_f, dollar_f in real_rows:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=3).value = pct_f
        ws.cell(row=row, column=5).value = dollar_f

    style_range(ws, 24, 5, 27, 5, num_fmt='$#,##0.00')
    style_range(ws, 24, 3, 27, 3, num_fmt='0%')
    for r in range(24, 28):
        bg = WHITE_ if r % 2 == 0 else WARM_BEIGE
        style_range(ws, r, 1, r, 5, bg=bg)
    style_range(ws, 22, 1, 27, 5, border=thin_border)

    # ── ROW 29+: Transaction Viewer ──────────────────────────────────
    merge_and_set(ws, 'A29:F29', 'TRANSACTION TRACKER',
                  bg=DARK_SAGE, ft=font(WHITE_, bold=True))
    tx_headers = ['DATE','AMOUNT','SUB-CATEGORY','CATEGORY','DESCRIPTION','BUCKET']
    for c, h in enumerate(tx_headers, 1):
        cell = ws.cell(row=30, column=c, value=h)
        cell.fill = fill(DUSTY_ROSE)
        cell.font = font(WHITE_, bold=True)

    # QUERY formula (Google Sheets native — won't work in Excel but will in GSheets)
    ws['A31'] = ('=IFERROR(QUERY(Transactions!A:F,'
                 '"SELECT A,B,C,D,E,F WHERE A IS NOT NULL '
                 'AND MONTH(A)+1 = "&I$6&"'
                 ' ORDER BY A DESC '
                 'LABEL A \'\',B \'\',C \'\',D \'\',E \'\',F \'\'"),"")')

    # ── Column widths ────────────────────────────────────────────────
    col_widths = {'A':22,'B':4,'C':14,'D':4,'E':14,'F':4,'G':18,'H':18,'I':18}
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # ── Conditional Formatting ───────────────────────────────────────
    # Rule 1: Amount Left > 0 → green
    ws.conditional_formatting.add('E13',
        CellIsRule(operator='greaterThan', formula=['0'],
                   fill=fill(LIGHT_GREEN), font=Font(color=DARK_SAGE, bold=True)))
    # Rule 2: Amount Left < 0 → red
    ws.conditional_formatting.add('E13',
        CellIsRule(operator='lessThan', formula=['0'],
                   fill=fill(LIGHT_RED), font=Font(color=RED_TEXT, bold=True)))
    # Rule 3: Real over goal (per row)
    ws.conditional_formatting.add('E24',
        FormulaRule(formula=['E24>E17'], fill=fill(LIGHT_RED)))
    ws.conditional_formatting.add('E25',
        FormulaRule(formula=['E25>E18'], fill=fill(LIGHT_RED)))
    ws.conditional_formatting.add('E26',
        FormulaRule(formula=['E26>E19'], fill=fill(LIGHT_RED)))
    # Rule 4: Budget summary real > budget
    ws.conditional_formatting.add('E8:E12',
        FormulaRule(formula=['E8>C8'], font=Font(color=RED_TEXT)))

    # ── Freeze rows 1-4 ─────────────────────────────────────────────
    ws.freeze_panes = 'A5'

    # ── Charts ───────────────────────────────────────────────────────
    build_charts(ws)

    return ws


def build_charts(ws):
    # For the charts we need to use static values that openpyxl can plot.
    # Google Sheets will recalculate formulas live after upload.
    # We create charts referencing the formula cells — they'll update in GSheets.

    # Chart 1: Grouped Bar — Goal vs Real
    chart1 = BarChart()
    chart1.type = 'col'
    chart1.grouping = 'clustered'
    chart1.title = 'MY GOAL vs MY REAL BUDGET'
    chart1.y_axis.numFmt = '$#,##0'
    chart1.legend.position = 't'
    chart1.width = 16
    chart1.height = 11

    cats = ChartRef(ws, min_col=1, min_row=17, max_row=19)
    goal_data = ChartRef(ws, min_col=5, min_row=16, max_row=19)
    real_data = ChartRef(ws, min_col=5, min_row=23, max_row=26)

    chart1.add_data(goal_data, titles_from_data=True)
    chart1.add_data(real_data, titles_from_data=True)
    chart1.set_categories(cats)

    # Series colors
    from openpyxl.chart.series import SeriesLabel
    chart1.series[0].title = SeriesLabel(v='MY GOAL')
    chart1.series[0].graphicalProperties.solidFill = '4A5E52'
    chart1.series[1].title = SeriesLabel(v='MY BUDGET')
    chart1.series[1].graphicalProperties.solidFill = 'C49A8A'

    ws.add_chart(chart1, 'G6')

    # Chart 2: Donut — Goal Breakdown
    chart2 = PieChart()
    chart2.title = 'MY GOAL BREAKDOWN'
    chart2.width = 13
    chart2.height = 10

    labels2 = ChartRef(ws, min_col=1, min_row=17, max_row=19)
    data2 = ChartRef(ws, min_col=5, min_row=17, max_row=19)
    chart2.add_data(data2)
    chart2.set_categories(labels2)
    chart2.legend.position = 'r'

    # Donut style
    chart2.style = 10
    # Slice colors
    colors = ['4A5E52', 'C49A8A', '8FA898']
    for i, color in enumerate(colors):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = color
        chart2.series[0].data_points.append(pt)

    ws.add_chart(chart2, 'G15')

    # Chart 3: Donut — Real Breakdown
    chart3 = PieChart()
    chart3.title = 'MY REAL BUDGET BREAKDOWN'
    chart3.width = 13
    chart3.height = 10

    labels3 = ChartRef(ws, min_col=1, min_row=24, max_row=26)
    data3 = ChartRef(ws, min_col=5, min_row=24, max_row=26)
    chart3.add_data(data3)
    chart3.set_categories(labels3)
    chart3.legend.position = 'r'

    chart3.style = 10
    for i, color in enumerate(colors):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = color
        chart3.series[0].data_points.append(pt)

    ws.add_chart(chart3, 'G22')


# ====================================================================
#  MAIN
# ====================================================================
def main():
    wb = openpyxl.Workbook()

    # Build in order: Reference → Budget Setup → Transactions → Dashboard
    build_reference(wb)
    build_budget_setup(wb)
    build_transactions(wb)
    build_dashboard(wb)

    # Reorder sheets: Dashboard, Transactions, Budget Setup, Reference
    wb.move_sheet('Dashboard', offset=-3)

    output = '/home/user/claude/Budget_Tracker_50_30_20.xlsx'
    wb.save(output)
    print(f'✅ Saved to {output}')
    print('Upload this file to Google Drive, then open with Google Sheets.')
    print('The QUERY formula and all SUMIFS will work natively in Google Sheets.')

if __name__ == '__main__':
    main()
