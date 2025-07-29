#test 2 both standalone and consolidated


import pdfplumber

pdf_path = "RIBATEXTILE.pdf"

# Headings for both statement types (add more if you discover more variants)
STANDALONE_HEADINGS = [
    "STANDALONE BALANCE SHEET",

    "STATEMENT OF PROFIT AND LOSS FOR THE YEAR ENDED MARCH 31"
    "BALANCE SHEET AS AT MARCH 31", #these two are not logical
    
    "STANDALONE STATEMENT OF PROFIT AND LOSS",
    "STATEMENT OF PROFIT & LOSS FOR THE YEAR ENDED ON 31ST MARCH",
    "Profit & Loss Statement for the Year ended 31st March",
    "STANDALONE STATEMENT OF CASH FLOWS",
    "STANDALONE CASH FLOW STATEMENT",

    "BALANCE SHEET AS AT 31ST MARCH", #not logical
    
    "STANDALONE STATEMENT OF PROFIT & LOSS",
    "STANDALONE STATEMENT OF FINANCIAL POSITION",
    "Cash Flow statement for the Year ended 31 March",
    "STANDALONE STATEMENT OF CASH FLOW",
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
    "Consolidated Profit and Loss Account",
    "CONSOLIDATED STATEMENT OF CASH FLOW"
]

   

# Uppercase for robust matching
STANDALONE_HEADINGS = [h.upper() for h in STANDALONE_HEADINGS]
CONSOLIDATED_HEADINGS = [h.upper() for h in CONSOLIDATED_HEADINGS]

def matches_heading(text, heading_list):
    if not text:
        return False
    lines = text.splitlines()
    for line in lines[:6]: 
        line_upper = line.strip().upper()
        for heading in heading_list:
            if heading in line_upper:
                return True
    return False

standalone_pages = []
consolidated_pages = []

with pdfplumber.open(pdf_path) as pdf:
    for i, page in enumerate(pdf.pages):
        text = page.extract_text()
        if not text:
            continue
        page_num = i + 1

        if matches_heading(text, STANDALONE_HEADINGS):
            standalone_pages.append(page_num)
            print(f"\n--- STANDALONE MATCH: Page {page_num} ---\n")
            print(text)

        if matches_heading(text, CONSOLIDATED_HEADINGS):
            consolidated_pages.append(page_num)
            print(f"\n--- CONSOLIDATED MATCH: Page {page_num} ---\n")
            print(text)

print()  # Just a line break for neatness
if not standalone_pages:
    print("No standalone financial statements found in the PDF.")
else:
    print(f"Standalone statements found on pages: {standalone_pages}")

if not consolidated_pages:
    print("No consolidated financial statements found in the PDF.")
else:
    print(f"Consolidated statements found on pages: {consolidated_pages}")
