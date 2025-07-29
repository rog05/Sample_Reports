#FINAL CODE

import pdfplumber
import pandas as pd
import re

pdf_path = "AXIS.pdf"             
output_excel = "AXIS1.xlsx"   

STANDALONE_HEADINGS = [
    "STANDALONE BALANCE SHEET",
    "STANDALONE STATEMENT OF PROFIT AND LOSS",
    "STATEMENT OF PROFIT & LOSS FOR THE YEAR ENDED ON 31ST MARCH",
    "Profit & Loss Statement for the Year ended 31st March",
    "STANDALONE STATEMENT OF CASH FLOWS",
    "STANDALONE CASH FLOW STATEMENT",
    "BALANCE SHEET AS AT 31ST MARCH",
    "STANDALONE STATEMENT OF PROFIT & LOSS",
    "STANDALONE STATEMENT OF FINANCIAL POSITION",
    "Cash Flow statement for the Year ended 31 March",

    "CASH FLOW STATEMENT FOR THE YEAR ENDED 31st MARCH",
    "STANDALONE STATEMENT OF FINANCIAL POSITION",
    "Standalone Profit and Loss Account",

    "STANDALONE STATEMENT OF PROFIT AND LOSS FOR THE YEAR ENDED MARCH 31"

]

CONSOLIDATED_HEADINGS = [
    "CONSOLIDATED BALANCE SHEET",
    "CONSOLIDATED STATEMENT OF PROFIT AND LOSS",
    "Consoldiated Profit & Loss Statement for the Year ended 31st March",
    "CONSOLIDATED STATEMENT OF CASH FLOWS",
    "CONSOLIDATED CASH FLOW STATEMENT",
    "CONSOLIDATED STATEMENT OF PROFIT & LOSS",
    "CONSOLIDATED STATEMENT OF FINANCIAL POSITION",
    "Consolidated Cash Flow statement for the Year ended 31 March",
    "Consolidated Profit and Loss Account"
]

STANDALONE_HEADINGS = [h.upper() for h in STANDALONE_HEADINGS]
CONSOLIDATED_HEADINGS = [h.upper() for h in CONSOLIDATED_HEADINGS]

# Notes : 2, 2.4, 9(a), 7(b), 10(a)
NOTE_RE = re.compile(r'^\d+([a-zA-Z])?(\(\w+\))?(\.\d+)?([a-zA-Z]?)?$|^\d+\.\d+(\([a-zA-Z0-9]+\))?$|^\d+\([a-zA-Z0-9]+\)$')
# Numeric value tokens (allowing commas, decimals)
VALUE_RE = re.compile(r'^-?[\d,]+(?:\.\d+)?$')

def matches_heading(text, heading_list):
    if not text:
        return None
    lines = text.splitlines()
    for line in lines[:5]:
        line_upper = line.strip().upper()
        for heading in heading_list:
            if heading in line_upper:
                return heading
    return None

def find_years_in_text(lines):
    # Find actual year titles like "31 March 2024" or "31-03-2024"
    year_regex = r'\d{1,2}\s+\w+\s+\d{4}'
    for i,line in enumerate(lines):
        found = re.findall(year_regex, line)
        found = [y.strip() for y in found]
        # Only take unique, non-blank years
        found = [y for y in found if y]
        if len(found) >= 2:
            return found[:2]
    # fallback
    return ["Year1", "Year2"]

def parse_financial_lines(lines, year_cols, cashflow_section=False):
    # Value can be (4,494), 4,494, (243), 5.67, (1,478), -.
    VALUE_RE = re.compile(r'^\(?-?[\d,]+(?:\.\d+)?\)?$|^-$')
    # Note can be 2, 2.4, 9(a), 10(a)
    NOTE_RE = re.compile(r'^\d+([a-zA-Z])?(\(\w+\))?(\.\d+)?([a-zA-Z]?)?$|^\d+\.\d+(\([a-zA-Z0-9]+\))?$|^\d+\([a-zA-Z0-9]+\)$')

    data = []
    for line in lines:
        line = re.sub(r'\s+', ' ', line.strip())
        if not line: 
            continue
        tokens = line.split(' ')
        n = len(tokens)
        particulars = ''
        note = ''
        year1 = ''
        year2 = ''
        # --- Main case: two trailing value tokens (incl. "-", (xxx), numbers)
        if n >= 2 and VALUE_RE.fullmatch(tokens[-1]) and VALUE_RE.fullmatch(tokens[-2]):
            # ...NOTE? value1 value2
            if n >= 3 and NOTE_RE.fullmatch(tokens[-3]):
                particulars = ' '.join(tokens[:-3])
                note = tokens[-3]
            else:
                particulars = ' '.join(tokens[:-2])
                note = ''
            year1 = tokens[-2]
            year2 = tokens[-1]
        # --- Special: trailing single value or note
        elif n >= 1 and VALUE_RE.fullmatch(tokens[-1]):
            particulars = ' '.join(tokens[:-1])
            note = ''
            year1 = tokens[-1]
            year2 = ''
        elif n >= 1 and NOTE_RE.fullmatch(tokens[-1]):
            particulars = ' '.join(tokens[:-1])
            note = tokens[-1]
            year1 = ''
            year2 = ''
        # --- Default: only text (e.g., section/subsection row like "ASSETS")
        else:
            particulars = line
            note = ''
            year1 = ''
            year2 = ''
        # If this is a cash flow section, always leave Notes blank
        if cashflow_section:
            note = ''
        # Always append, never skip
        data.append({
            "Particulars": particulars.strip(),
            "Notes": note,
            year_cols[0]: year1,
            year_cols[1]: year2
        })
    return pd.DataFrame(data)



extracted_data = {"Standalone": [], "Consolidated": []}

with pdfplumber.open(pdf_path) as pdf:
    for i, page in enumerate(pdf.pages):
        text = page.extract_text()
        if not text: continue
        matched_standalone = matches_heading(text, STANDALONE_HEADINGS)
        matched_consolidated = matches_heading(text, CONSOLIDATED_HEADINGS)
        if not (matched_standalone or matched_consolidated): continue

        lines = text.splitlines()
        year_cols = find_years_in_text(lines)
        # Find start of data (header idx)
        header_idx = None
        for idx, line in enumerate(lines):
            if "PARTICULARS" in line.upper() and all(y in line for y in year_cols):
                header_idx = idx
                break
        if header_idx is None:
            header_idx = 5  # fallback
        data_lines = lines[header_idx + 1:]

        # Is this a cash flow statement?
        headline = matched_standalone or matched_consolidated
        headline_lower = headline.lower()
        cashflow_section = "cash flow" in headline_lower

        df = parse_financial_lines(data_lines, year_cols, cashflow_section=cashflow_section)
        if matched_standalone:
            extracted_data["Standalone"].append((headline, df))
        if matched_consolidated:
            extracted_data["Consolidated"].append((headline, df))

# -------- Write to Excel as you need --------
with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
    for section, items in extracted_data.items():
        if not items:
            continue
        sheet_name = section[:31]
        start_row = 0
        for heading, df in items:
            worksheet = writer.book[sheet_name] if sheet_name in writer.book.sheetnames else writer.book.create_sheet(sheet_name)
            cell = worksheet.cell(row=start_row+1, column=1)
            cell.value = heading
            cell.font = cell.font.copy(bold=True)
            df.to_excel(writer, sheet_name=sheet_name, startrow=start_row + 1, index=False)
            start_row += len(df) + 3

print(f"\nExtraction complete! Data saved to '{output_excel}'")
