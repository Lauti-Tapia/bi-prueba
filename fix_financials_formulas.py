"""
Parche directo de XML para reparar las fórmulas rotas de la pestaña Financials.

Contexto: cuando guardé el workbook con openpyxl se perdieron las Power Query
connections y las definiciones de tabla. Después, al abrir con Excel, las
referencias estructuradas cr83c_financials[...] se "congelaron" a #REF! en el
texto de las fórmulas. Este script NO usa openpyxl: abre el .xlsx como zip,
reescribe sólo xl/worksheets/sheet5.xml (la pestaña Financials) reemplazando
cada celda de fórmula por su versión con rangos directos al sheet _DealFinancial,
y reempaca todo lo demás tal cual.

Columnas en _DealFinancial (confirmadas en sharedStrings):
    A = Deal Name
    B = Deal Number
    C = Date
    D = Property Size
    E = Category
    F = Account Category
    G = Sub-Category
    H = Internal Line Item
    I = Amount
    J = Comments
    K = Loan Amount
Data: A2:K4333 (uso referencias de columna completa por si crece).
"""

import zipfile
import shutil
import re
import os
from datetime import datetime, timedelta

WORKBOOK = "Dynamics Workbook_03.23.2026 1.xlsx"
SHEET_XML = "xl/worksheets/sheet5.xml"  # Financials
DEAL_XML  = "xl/worksheets/sheet4.xml"  # _DealFinancial

# ---------- range refs to _DealFinancial ----------
DEAL_COL = "_DealFinancial!$A:$A"   # Deal Name
DATE_COL = "_DealFinancial!$C:$C"   # Date (Excel serial)
CAT_COL  = "_DealFinancial!$E:$E"   # Category
ILI_COL  = "_DealFinancial!$H:$H"   # Internal Line Item
AMT_COL  = "_DealFinancial!$I:$I"   # Amount
PSZ_COL  = "_DealFinancial!$D:$D"   # Property Size
DEAL_REF = "COVER!$C$2"

# ---------- date picker cells (Financials sheet) ----------
# I2 = "Date:" label, J2 = dropdown, K2 = readable display,
# N2 = numeric effective date used by SUMIFS (hidden column),
# O2..O{N+1} = static list of "yyyy-mm-dd" text values (hidden column).
USER_DATE_CELL = "$J$2"                  # user's dropdown pick (text "yyyy-mm-dd" or blank)
EFF_DATE_CELL  = "Financials!$N$2"       # numeric effective date for SUMIFS criteria
DATE_RANGE_REF = "$A$2:$A$5000"          # range used inside AGGREGATE lookups
DATE_SERIAL_REF = "$C$2:$C$5000"

# NOTE: PBI's $/unit measure multiplies the numerator by [Annual Factor] =
# 12 / DISTINCTCOUNT(DimDate[Year-Month]), where DimDate is CALENDARAUTO()
# scanning *all* date columns across the whole semantic model (~10 years of
# calendar → factor ~0.1). The business agreed that the annualization factor
# is not important for this workbook, so we intentionally drop it and compute
# $/unit as a straight $ / units. Earlier versions of this script injected a
# "Annual Factor" helper cell at J2/K2 — those cells are now cleaned up on
# every run to make the patch idempotent.

# ---------- layout ----------
# Tuples: (excel_row, tag, row_type). Section rows (no formulas) excluded.
LAYOUT = [
    (7,  "ADJ_GPR",                  "subtotal"),
    (8,  "Vacancy",                  "data"),
    (9,  "Concessions",              "data"),
    (10, "Bad Debt Loss",            "data"),
    (11, "EFF_RENTAL",               "subtotal"),
    (12, "Utility Income",           "data"),
    (13, "Other Income",             "data"),
    (14, "EFF_GROSS_REV",            "subtotal"),
    (16, "General & Administrative", "data"),
    (17, "Payroll",                  "data"),
    (18, "Sales & Marketing",        "data"),
    (19, "Utilities",                "data"),
    (20, "Repairs & Maintenance",    "data"),
    (21, "Management Fee",           "data"),
    (22, "Turnover",                 "data"),
    (23, "TOTAL_CTRL",               "subtotal"),
    (24, "Property Taxes",           "data"),
    (25, "Insurance",                "data"),
    (26, "TOTAL_OPEX",               "subtotal"),
    (27, "NOI",                      "final"),
    (29, "CapEx Reserve",            "data"),
    (30, "Debt Service",             "data"),
    (31, "CASHFLOW",                 "final"),
]

ROW_BY_TAG = {tag: row for row, tag, _ in LAYOUT}

# Full layout including section header rows (which carry no formulas — just a
# label spanning the ACT/UW area). Used by build_sheet_xml_from_scratch().
# Tuples: (row, label, row_type). row_type is one of:
#   "section"  -> label only, no formulas
#   "data"     -> SUMIFS on a single Internal Line Item
#   "subtotal" -> SUMIFS over multiple items, or sum of prior rows
#   "final"    -> sum-of-prior-rows for NOI / CashFlow
FULL_LAYOUT = [
    (4,  "Category",                  "header1"),  # top header row
    (5,  None,                        "header2"),  # sub-header ($ / % / $ / %)
    (6,  "Revenue",                   "section"),
    (7,  "Adj. GPR",                  "subtotal"),
    (8,  "Vacancy",                   "data"),
    (9,  "Concessions",               "data"),
    (10, "Bad Debt Loss",             "data"),
    (11, "Effective Rental Income",   "subtotal"),
    (12, "Utility Income",            "data"),
    (13, "Other Income",              "data"),
    (14, "Effective Gross Revenue",   "subtotal"),
    (15, "Operating Expenses",        "section"),
    (16, "General & Administrative",  "data"),
    (17, "Payroll",                   "data"),
    (18, "Sales & Marketing",         "data"),
    (19, "Utilities",                 "data"),
    (20, "Repairs & Maintenance",     "data"),
    (21, "Management Fee",            "data"),
    (22, "Turnover",                  "data"),
    (23, "Total Controllable Expenses","subtotal"),
    (24, "Property Taxes",            "data"),
    (25, "Insurance",                 "data"),
    (26, "Total Operating Expenses",  "subtotal"),
    (27, "Net Operating Income",      "final"),
    (28, "Below NOI",                 "section"),
    (29, "CapEx Reserve",             "data"),
    (30, "Debt Service",              "data"),
    (31, "Cash Flow from Operations", "final"),
]

# Each tag in LAYOUT corresponds to a label in FULL_LAYOUT (the "Internal Line
# Item" in the source table for data rows, or a subtotal/final tag).
# Mapping fix-script tags -> Internal Line Item (only for data rows).
ILI_BY_LABEL = {
    "Vacancy":                  "Vacancy",
    "Concessions":              "Concessions",
    "Bad Debt Loss":            "Bad Debt Loss",
    "Utility Income":           "Utility Income",
    "Other Income":             "Other Income",
    "General & Administrative": "General & Administrative",
    "Payroll":                  "Payroll",
    "Sales & Marketing":        "Sales & Marketing",
    "Utilities":                "Utilities",
    "Repairs & Maintenance":    "Repairs & Maintenance",
    "Management Fee":           "Management Fee",
    "Turnover":                 "Turnover",
    "Property Taxes":           "Property Taxes",
    "Insurance":                "Insurance",
    "CapEx Reserve":            "CapEx Reserve",
    "Debt Service":             "Debt Service",
}

# Subtotal / final labels -> tag used by build_dollar_formula().
TAG_BY_LABEL = {
    "Adj. GPR":                    "ADJ_GPR",
    "Effective Rental Income":     "EFF_RENTAL",
    "Effective Gross Revenue":     "EFF_GROSS_REV",
    "Total Controllable Expenses": "TOTAL_CTRL",
    "Total Operating Expenses":    "TOTAL_OPEX",
    "Net Operating Income":        "NOI",
    "Cash Flow from Operations":   "CASHFLOW",
}

# Labels that use Adj.GPR (row 7) as the %-of-GPR denominator instead of the
# default (Effective Gross Revenue, row 14).
REVENUE_BLOCK_LABELS = {
    "Adj. GPR", "Vacancy", "Concessions", "Bad Debt Loss",
    "Effective Rental Income", "Utility Income", "Other Income",
    "Effective Gross Revenue",
}

# Tags whose % of GPR denominator is Adj.GPR (row 7).
# All others use Effective Gross Revenue (row 14).
REVENUE_BLOCK = {
    "ADJ_GPR", "Vacancy", "Concessions", "Bad Debt Loss",
    "EFF_RENTAL", "Utility Income", "Other Income", "EFF_GROSS_REV",
}


def sumifs_single(category: str, item: str) -> str:
    return (
        f'SUMIFS({AMT_COL},{DEAL_COL},{DEAL_REF},'
        f'{CAT_COL},"{category}",{ILI_COL},"{item}",'
        f'{DATE_COL},{EFF_DATE_CELL})'
    )


def sumifs_multi(category: str, items: list) -> str:
    arr = "{" + ",".join(f'"{i}"' for i in items) + "}"
    return (
        f'SUM(SUMIFS({AMT_COL},{DEAL_COL},{DEAL_REF},'
        f'{CAT_COL},"{category}",{ILI_COL},{arr},'
        f'{DATE_COL},{EFF_DATE_CELL}))'
    )


def maxifs_units() -> str:
    # INDEX+MATCH in place of MAXIFS for compatibility with pre-2019 Excel.
    # All rows of a given deal share the same Property Size, so the first
    # match is sufficient.
    return f'INDEX({PSZ_COL},MATCH({DEAL_REF},{DEAL_COL},0))'


def build_dollar_formula(tag: str, category: str, dollar_col: str) -> str:
    """Return the body (without leading '=') of the $ formula for a given row tag."""
    if tag == "ADJ_GPR":
        return sumifs_multi(category, ["Gross Potential Rent", "Loss to Lease"])
    if tag == "EFF_RENTAL":
        rows = [ROW_BY_TAG[k] for k in ("ADJ_GPR", "Vacancy", "Concessions", "Bad Debt Loss")]
        return "+".join(f"{dollar_col}{r}" for r in rows)
    if tag == "EFF_GROSS_REV":
        rows = [ROW_BY_TAG[k] for k in ("EFF_RENTAL", "Utility Income", "Other Income")]
        return "+".join(f"{dollar_col}{r}" for r in rows)
    if tag == "TOTAL_CTRL":
        items = ["General & Administrative", "Payroll", "Sales & Marketing",
                 "Utilities", "Repairs & Maintenance", "Management Fee", "Turnover"]
        rows = [ROW_BY_TAG[i] for i in items]
        return "+".join(f"{dollar_col}{r}" for r in rows)
    if tag == "TOTAL_OPEX":
        rows = [ROW_BY_TAG[k] for k in ("TOTAL_CTRL", "Property Taxes", "Insurance")]
        return "+".join(f"{dollar_col}{r}" for r in rows)
    if tag == "NOI":
        return f"{dollar_col}{ROW_BY_TAG['EFF_GROSS_REV']}-{dollar_col}{ROW_BY_TAG['TOTAL_OPEX']}"
    if tag == "CASHFLOW":
        return (f"{dollar_col}{ROW_BY_TAG['NOI']}"
                f"-{dollar_col}{ROW_BY_TAG['CapEx Reserve']}"
                f"-{dollar_col}{ROW_BY_TAG['Debt Service']}")
    # plain data row — tag IS the Internal Line Item value
    return sumifs_single(category, tag)


def build_all_cells() -> dict:
    """Return {(col_letter, row): formula_body_without_equals}.

    The $/unit columns (D, G) are intentionally skipped — those cells are
    cleared separately by clear_unit_cells().
    """
    cells = {}
    for row, tag, _ in LAYOUT:
        for side, category in (("CDE", "ACT"), ("FGH", "UW")):
            c_dollar, _c_unit, c_pct = side[0], side[1], side[2]
            cells[(c_dollar, row)] = build_dollar_formula(tag, category, c_dollar)
            denom_row = ROW_BY_TAG["ADJ_GPR"] if tag in REVENUE_BLOCK else ROW_BY_TAG["EFF_GROSS_REV"]
            cells[(c_pct,    row)] = f"IFERROR({c_dollar}{row}/{c_dollar}${denom_row},0)"
    return cells


def xml_escape_formula(f: str) -> str:
    """Escape XML-special chars for inclusion inside <f>...</f> element text."""
    # Only & and < need escaping in element content. Our formulas never contain <.
    return f.replace("&", "&amp;")


def patch_sheet(xml: str, cells: dict) -> tuple:
    """
    For each target cell, replace its <c>...</c> element with one that contains
    only the new formula. Preserves the original style attribute s="...".
    Returns (new_xml, num_replaced).
    """
    count = 0
    for (col, row), formula in cells.items():
        coord = f"{col}{row}"
        # Greedy-safe match of the cell element
        pattern = re.compile(
            rf'<c r="{re.escape(coord)}"([^>]*)(?:/>|>.*?</c>)',
            re.DOTALL,
        )
        escaped = xml_escape_formula(formula)

        def repl(m):
            nonlocal count
            attrs = m.group(1)
            # Drop t="..." (error / string / etc) — we want default numeric type
            attrs = re.sub(r'\s*t="[^"]*"', '', attrs)
            count += 1
            return f'<c r="{coord}"{attrs}><f>{escaped}</f></c>'

        xml, n = pattern.subn(repl, xml, count=1)
    return xml, count


def clear_annual_factor_cells(xml: str) -> str:
    """
    Remove the J2 (label) and K2 (value) helper cells that earlier versions of
    this script injected for the annualization factor. Idempotent: no-op if the
    cells aren't present.
    """
    xml = re.sub(r'<c r="J2"[^>]*(?:/>|>.*?</c>)', '', xml, flags=re.DOTALL)
    xml = re.sub(r'<c r="K2"[^>]*(?:/>|>.*?</c>)', '', xml, flags=re.DOTALL)
    return xml


def clear_unit_cells(xml: str) -> str:
    """
    Empty the $/unit cells on both the ACT (column D) and UW (column G) sides,
    across all table rows (4..31). Includes the "$/unit" sub-header on row 5
    and any merged-continuation cells on rows 4 and 6.
    """
    for col in ("D", "G"):
        for row in range(4, 32):
            coord = f"{col}{row}"
            pattern = re.compile(
                rf'<c r="{re.escape(coord)}"[^>]*(?:/>|>.*?</c>)',
                re.DOTALL,
            )
            xml = pattern.sub(f'<c r="{coord}"/>', xml)
    return xml


def strip_table_styles(xml: str) -> str:
    """
    Remove the s="..." style attribute from every cell inside the Financials
    table area (columns B..H, rows 2..31). Leaves cell values and formulas
    untouched — only the visual style reference is dropped.
    """
    for row in range(2, 32):
        for col in "BCDEFGH":
            coord = f"{col}{row}"
            pattern = re.compile(rf'(<c r="{re.escape(coord)}")(\s[^>]*)?(/>|>)')

            def strip(m):
                attrs = m.group(2) or ""
                attrs = re.sub(r'\s*s="\d+"', "", attrs)
                return f"{m.group(1)}{attrs}{m.group(3)}"

            xml = pattern.sub(strip, xml)
    return xml


def update_cols_hide_units(xml: str, n_dates: int) -> str:
    """
    Rewrite the <cols> element so that:
      - D (4) and G (7): $/unit columns — hidden
      - N (14) .. last-helper: date-picker helpers — hidden
      - I (9), J (10), K (11): visible — these carry the date picker UI
    Drop any default column styles so the table no longer carries
    workbook-level formatting.
    """
    last_helper_col_idx = O_COL_IDX + n_dates - 1  # inclusive
    new_cols = (
        "<cols>"
        '<col min="2" max="2" width="32" customWidth="1"/>'
        '<col min="3" max="3" width="14" customWidth="1"/>'
        '<col min="4" max="4" width="14" hidden="1" customWidth="1"/>'
        '<col min="5" max="6" width="14" customWidth="1"/>'
        '<col min="7" max="7" width="14" hidden="1" customWidth="1"/>'
        '<col min="8" max="8" width="14" customWidth="1"/>'
        '<col min="9" max="9" width="8" customWidth="1"/>'
        '<col min="10" max="10" width="14" customWidth="1"/>'
        '<col min="11" max="11" width="14" customWidth="1"/>'
        f'<col min="14" max="{last_helper_col_idx}" width="14" hidden="1" customWidth="1"/>'
        "</cols>"
    )
    xml, n = re.subn(r"<cols>.*?</cols>", new_cols, xml, count=1, flags=re.DOTALL)
    if n == 0:
        raise SystemExit("ERROR: <cols> element not found in Financials sheet")
    return xml


# ---------- date picker ----------

def extract_unique_dates(deal_xml: str) -> list:
    """Return sorted-desc list of unique Excel-serial dates in _DealFinancial!C."""
    dates = set()
    rows = re.findall(r'<row r="(\d+)"[^>]*>(.*?)</row>', deal_xml, re.DOTALL)
    for rn, content in rows[1:]:  # skip header
        for m in re.finditer(
            r'<c r="([A-Z]+)(\d+)"(?:\s+t="([^"]+)")?[^>]*?'
            r'(?:>(?:<v>([^<]*)</v>)?</c>|/>)',
            content,
        ):
            col, _row, t, v = m.group(1), m.group(2), m.group(3), m.group(4)
            if col == "C" and v is not None and t != "s":
                try:
                    dates.add(int(v))
                except ValueError:
                    pass
    return sorted(dates, reverse=True)


def serial_to_text(n: int) -> str:
    return (datetime(1899, 12, 30) + timedelta(days=int(n))).strftime("%Y-%m-%d")


def _col_letter(n: int) -> str:
    """1-based column index -> Excel column letter (1=A, 27=AA, ...)."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# Column indexes for the date-picker helpers (1-based).
N_COL_IDX = 14   # N
O_COL_IDX = 15   # O = first cell of date list
# Last column of the date list depends on len(date_serials); computed at runtime.


def insert_date_picker(xml: str, date_serials: list) -> str:
    """
    Inject the date picker into row 2 of the Financials sheet:
      B2 (existing) = deal name (=COVER!C2)
      I2            = "Date:" label
      J2            = empty dropdown cell (text type)
      K2            = readable effective date =TEXT($N$2,"yyyy-mm-dd")
      N2            = numeric effective date (what SUMIFS filters by)
      O2..{last}2   = static "yyyy-mm-dd" list backing the dropdown
                      (laid out HORIZONTALLY in row 2 so we don't have to
                      touch other rows — any new <row> inserted mid-sheet
                      breaks Excel's required row-order invariant.)

    The AGGREGATE(14,6,...) in N2 is a legacy-safe replacement for MAXIFS:
    _DealFinancial!C / (A=deal) yields #DIV/0 on non-matching rows; option 6
    tells AGGREGATE to ignore errors and LARGE-k=1 returns the max.
    """
    last_col_idx = O_COL_IDX + len(date_serials) - 1
    helper_cols = [_col_letter(i) for i in range(O_COL_IDX, last_col_idx + 1)]

    # --- Idempotent cleanup -------------------------------------------------
    # Remove any date-picker cells left over from a prior run. This covers
    # I2/J2/K2/N2 and every possible helper column we've ever used, so we
    # recover cleanly even if the number of dates changed.
    cleanup_coords = ["I2", "J2", "K2", "N2"]
    for i in range(O_COL_IDX, O_COL_IDX + 100):  # generous upper bound
        cleanup_coords.append(f"{_col_letter(i)}2")
    for coord in cleanup_coords:
        xml = re.sub(
            rf'<c r="{re.escape(coord)}"[^>]*(?:/>|>.*?</c>)',
            "",
            xml,
            flags=re.DOTALL,
        )

    # Also strip any O-column cells (legacy broken state where dates were
    # written vertically in column O across rows 2..21).
    xml = re.sub(
        r'<c r="O\d+"[^>]*(?:/>|>.*?</c>)',
        "",
        xml,
        flags=re.DOTALL,
    )

    # Delete any now-empty <row> elements left over from that same legacy
    # state. These rows had r-numbers clashing with the real table rows and
    # must go before we can save valid XML.
    xml = re.sub(
        r'<row r="\d+"[^>]*>\s*</row>',
        "",
        xml,
    )

    # --- Build the new row-2 cells -----------------------------------------
    cells = []

    # Label
    cells.append('<c r="I2" t="inlineStr"><is><t>Date:</t></is></c>')

    # Dropdown target — start empty (= use latest)
    cells.append('<c r="J2" t="str"><v></v></c>')

    # Visible effective-date display
    cells.append('<c r="K2" t="str"><f>TEXT($N$2,"yyyy-mm-dd")</f></c>')

    # Numeric effective date used by every SUMIFS.
    # If J2 is blank -> AGGREGATE picks the latest date for the current deal,
    # otherwise -> DATEVALUE of the user's selection.
    n2_formula = (
        f'IF({USER_DATE_CELL}="",'
        f'AGGREGATE(14,6,_DealFinancial!{DATE_SERIAL_REF}/'
        f'(_DealFinancial!{DATE_RANGE_REF}={DEAL_REF}),1),'
        f'DATEVALUE({USER_DATE_CELL}))'
    )
    cells.append(f'<c r="N2"><f>{n2_formula}</f></c>')

    # Horizontal list of date strings, one per column starting at O2.
    for col_letter, serial in zip(helper_cols, date_serials):
        text = serial_to_text(serial)
        cells.append(
            f'<c r="{col_letter}2" t="inlineStr"><is><t>{text}</t></is></c>'
        )

    # --- Inject into row 2 -------------------------------------------------
    row2_pattern = re.compile(r'(<row r="2"[^>]*)( spans="[^"]*")?(>)(.*?)(</row>)', re.DOTALL)
    m = row2_pattern.search(xml)
    if not m:
        raise SystemExit("ERROR: row 2 not found in Financials sheet")
    # Widen the row's "spans" so Excel knows the row extends to the helper cols.
    new_spans = f' spans="2:{last_col_idx}"'
    new_row2 = (
        m.group(1) + new_spans + m.group(3)
        + m.group(4) + "".join(cells) + m.group(5)
    )
    xml = xml[:m.start()] + new_row2 + xml[m.end():]

    # Expose last_col_idx to the caller via a module attribute so the data-
    # validation builder can use the same column range.
    insert_date_picker.last_col_idx = last_col_idx  # type: ignore[attr-defined]
    return xml


def add_data_validation(xml: str, n_dates: int) -> str:
    """
    Add a <dataValidations> element with a list-type validation on J2 pointing
    at the row-2 horizontal range $O$2:${last_col}$2. The element must appear
    after <mergeCells> and before <pageMargins> per the xlsx schema. Idempotent.
    """
    # Drop any prior <dataValidations> so we can re-insert cleanly.
    xml = re.sub(r"<dataValidations\b.*?</dataValidations>", "", xml, flags=re.DOTALL)

    last_col_idx = O_COL_IDX + n_dates - 1
    last_col = _col_letter(last_col_idx)
    dv = (
        '<dataValidations count="1">'
        '<dataValidation type="list" allowBlank="1" showInputMessage="1" '
        'showErrorMessage="1" sqref="J2">'
        f'<formula1>$O$2:${last_col}$2</formula1>'
        '</dataValidation>'
        '</dataValidations>'
    )

    # Insert right before <pageMargins ... />
    xml, n = re.subn(r'(<pageMargins\b)', dv + r'\1', xml, count=1)
    if n == 0:
        raise SystemExit("ERROR: <pageMargins> anchor not found for dataValidations")
    return xml


def xml_escape_text(s: str) -> str:
    """Escape for element content (text inside <t>, <f>, etc.)."""
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_sheet_xml_from_scratch(date_serials: list) -> str:
    """
    Build a complete, valid sheet5.xml (the Financials sheet) from scratch.

    This is used instead of patch_sheet() when the previous sheet state is
    too corrupted to patch in place. It produces a minimal sheet with:
      - row 2: deal-name display (=COVER!C2) + date picker UI + hidden helpers
      - row 4..5: merged header band (Category | ACT | UW + $/%/etc sub-header)
      - row 6..31: labeled rows with SUMIFS / subtotal / final formulas
      - mergeCells for the header band
      - dataValidation list on J2 sourcing from the hidden O2:${last}2 row
      - no styles whatsoever (explicitly stripped to avoid inherited formatting)
    """
    last_col_idx = O_COL_IDX + len(date_serials) - 1
    helper_cols = [_col_letter(i) for i in range(O_COL_IDX, last_col_idx + 1)]
    last_col_letter = _col_letter(last_col_idx)

    # -------- row 2: deal name + date picker ---------------------------------
    row2_cells = []
    # Deal name mirrored from COVER
    row2_cells.append('<c r="B2" t="str"><f>COVER!C2</f></c>')
    # Date-picker label
    row2_cells.append('<c r="I2" t="inlineStr"><is><t>Date:</t></is></c>')
    # Dropdown cell (empty = latest)
    row2_cells.append('<c r="J2" t="str"><v></v></c>')
    # Readable display of the effective date
    row2_cells.append('<c r="K2" t="str"><f>TEXT($N$2,"yyyy-mm-dd")</f></c>')
    # Hidden numeric effective date used by every SUMIFS.
    n2_formula = (
        f'IF({USER_DATE_CELL}="",'
        f'AGGREGATE(14,6,_DealFinancial!{DATE_SERIAL_REF}/'
        f'(_DealFinancial!{DATE_RANGE_REF}={DEAL_REF}),1),'
        f'DATEVALUE({USER_DATE_CELL}))'
    )
    row2_cells.append(f'<c r="N2"><f>{n2_formula}</f></c>')
    # Hidden horizontal list of date strings feeding the dropdown.
    for col_letter, serial in zip(helper_cols, date_serials):
        text = serial_to_text(serial)
        row2_cells.append(
            f'<c r="{col_letter}2" t="inlineStr"><is><t>{text}</t></is></c>'
        )

    # -------- rows 4 & 5: header band ----------------------------------------
    row4_cells = [
        '<c r="B4" t="inlineStr"><is><t>Category</t></is></c>',
        '<c r="C4" t="inlineStr"><is><t>ACT</t></is></c>',
        '<c r="F4" t="inlineStr"><is><t>UW</t></is></c>',
    ]
    row5_cells = [
        '<c r="C5" t="inlineStr"><is><t>$</t></is></c>',
        '<c r="E5" t="inlineStr"><is><t>% of GPR</t></is></c>',
        '<c r="F5" t="inlineStr"><is><t>$</t></is></c>',
        '<c r="H5" t="inlineStr"><is><t>% of GPR</t></is></c>',
    ]

    # -------- rows 6..31: body -----------------------------------------------
    body_rows_xml = []
    for row, label, rtype in FULL_LAYOUT:
        if rtype in ("header1", "header2"):
            continue

        cells = []
        # Label in column B
        if label is not None:
            cells.append(
                f'<c r="B{row}" t="inlineStr"><is><t>{xml_escape_text(label)}</t></is></c>'
            )

        if rtype != "section":
            # Pick the formula category / build dollar formula.
            # Build_dollar_formula() needs the fix-script "tag" — map from label.
            if label in TAG_BY_LABEL:
                tag = TAG_BY_LABEL[label]
            else:
                tag = ILI_BY_LABEL[label]

            # Denominator for % of GPR
            denom_row = 7 if label in REVENUE_BLOCK_LABELS else 14

            for side, category in (("CE", "ACT"), ("FH", "UW")):
                c_dollar, c_pct = side[0], side[1]
                dollar_f = build_dollar_formula(tag, category, c_dollar)
                cells.append(
                    f'<c r="{c_dollar}{row}"><f>{xml_escape_formula(dollar_f)}</f></c>'
                )
                pct_f = f'IFERROR({c_dollar}{row}/{c_dollar}${denom_row},0)'
                cells.append(
                    f'<c r="{c_pct}{row}"><f>{pct_f}</f></c>'
                )

        body_rows_xml.append(
            f'<row r="{row}">{"".join(cells)}</row>'
        )

    # -------- assemble final sheet XML ---------------------------------------
    # <cols>: widen B, keep C/E/F/H normal, hide D/G ($/unit) and N..last.
    cols_xml = (
        "<cols>"
        '<col min="2" max="2" width="32" customWidth="1"/>'
        '<col min="3" max="3" width="14" customWidth="1"/>'
        '<col min="4" max="4" width="14" hidden="1" customWidth="1"/>'
        '<col min="5" max="6" width="14" customWidth="1"/>'
        '<col min="7" max="7" width="14" hidden="1" customWidth="1"/>'
        '<col min="8" max="8" width="14" customWidth="1"/>'
        '<col min="9" max="9" width="8" customWidth="1"/>'
        '<col min="10" max="10" width="14" customWidth="1"/>'
        '<col min="11" max="11" width="14" customWidth="1"/>'
        f'<col min="14" max="{last_col_idx}" width="14" hidden="1" customWidth="1"/>'
        "</cols>"
    )

    sheet_data = (
        "<sheetData>"
        f'<row r="2" spans="2:{last_col_idx}">{"".join(row2_cells)}</row>'
        f'<row r="4" spans="2:8">{"".join(row4_cells)}</row>'
        f'<row r="5" spans="2:8">{"".join(row5_cells)}</row>'
        + "".join(body_rows_xml)
        + "</sheetData>"
    )

    merge_cells = (
        '<mergeCells count="3">'
        '<mergeCell ref="B4:B5"/>'
        '<mergeCell ref="C4:E4"/>'
        '<mergeCell ref="F4:H4"/>'
        "</mergeCells>"
    )

    last_helper_col_letter = last_col_letter
    data_validation = (
        '<dataValidations count="1">'
        '<dataValidation type="list" allowBlank="1" showInputMessage="1" '
        'showErrorMessage="1" sqref="J2">'
        f'<formula1>$O$2:${last_helper_col_letter}$2</formula1>'
        '</dataValidation>'
        '</dataValidations>'
    )

    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="B2:{last_col_letter}31"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        + cols_xml
        + sheet_data
        + merge_cells
        + data_validation
        + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" '
          'header="0.3" footer="0.3"/>'
        + '</worksheet>'
    )
    return sheet


def drop_calc_chain(all_files: dict, content_types: str, wb_rels: str) -> tuple:
    """
    Remove xl/calcChain.xml and its references so Excel force-recalcs on open.
    """
    all_files.pop("xl/calcChain.xml", None)

    # Remove from workbook.xml.rels
    wb_rels = re.sub(
        r'<Relationship[^/]+calcChain[^/]*/>',
        '',
        wb_rels,
    )
    # Remove override from [Content_Types].xml
    content_types = re.sub(
        r'<Override[^/]+calcChain[^/]*/>',
        '',
        content_types,
    )
    return content_types, wb_rels


def main():
    if not os.path.exists(WORKBOOK):
        raise SystemExit(f"{WORKBOOK} not found")

    # Read all files from the zip
    with zipfile.ZipFile(WORKBOOK, "r") as z:
        all_files = {n: z.read(n) for n in z.namelist()}

    deal_xml = all_files[DEAL_XML].decode("utf-8")

    # Extract the list of unique dates from _DealFinancial — drives both the
    # dropdown options and the "latest date per deal" default.
    date_serials = extract_unique_dates(deal_xml)
    print(f"Unique dates found in _DealFinancial: {len(date_serials)}")
    if date_serials:
        print(f"  range: {serial_to_text(date_serials[-1])} .. {serial_to_text(date_serials[0])}")

    # Rebuild the Financials sheet from scratch. This is the single source of
    # truth for the sheet's structure; it replaces the earlier "patch the
    # existing sheet" pipeline, which cannot recover if the sheet is corrupt.
    new_xml = build_sheet_xml_from_scratch(date_serials)
    print(f"Rebuilt Financials sheet from scratch "
          f"({len(FULL_LAYOUT) - 2} body rows, "
          f"{len(date_serials)} dropdown dates)")

    all_files[SHEET_XML] = new_xml.encode("utf-8")

    # Drop calcChain so Excel recomputes everything on next open
    ct_path = "[Content_Types].xml"
    rels_path = "xl/_rels/workbook.xml.rels"
    if ct_path in all_files and rels_path in all_files:
        ct = all_files[ct_path].decode("utf-8")
        rels = all_files[rels_path].decode("utf-8")
        ct, rels = drop_calc_chain(all_files, ct, rels)
        all_files[ct_path] = ct.encode("utf-8")
        all_files[rels_path] = rels.encode("utf-8")

    # Write new zip next to the original, then replace atomically
    tmp = WORKBOOK + ".new"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in all_files.items():
            z.writestr(name, data)

    shutil.move(tmp, WORKBOOK)
    print("OK:", WORKBOOK)


if __name__ == "__main__":
    main()
