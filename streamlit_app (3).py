import io
import re
from datetime import datetime
from typing import List, Dict, Optional

import pdfplumber
import pandas as pd
import streamlit as st
import zipfile


TARGET_COLUMNS = [
    "ReferralManager",
    "ReferralEmail",
    "Brand",
    "QuoteNumber",
    "QuoteDate",
    "Company",
    "FirstName",
    "LastName",
    "ContactEmail",
    "ContactPhone",
    "Address",
    "County",
    "City",
    "State",
    "ZipCode",
    "Country",
    "item_id",
    "item_desc",
    "UnitPrice",
    "TotalSales",
    "QuoteValidDate",
    "CustomerNumber",
    "manufacturer_Name",
    "PDF",
    "DemoQuote",
]


# ---------- PDF PARSING HELPERS ----------

def extract_full_text(pdf_bytes: bytes) -> str:
    """Extract text from all pages of the PDF."""
    full_text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            full_text += txt + "\n"
    return full_text


def extract_header_info(full_text: str) -> Dict[str, Optional[str]]:
    """Extract quote-level fields from the header area."""
    header: Dict[str, Optional[str]] = {}

    # Quote number + date: "Quote 120987 Date 11/24/2025"
    m_quote = re.search(r"Quote\s+(\d+)\s+Date\s+(\d{1,2}/\d{1,2}/\d{4})", full_text)
    if m_quote:
        header["QuoteNumber"] = m_quote.group(1)
        header["QuoteDate"] = m_quote.group(2)

    # Customer number
    m_cust = re.search(r"Customer\s+(\d+)", full_text)
    if m_cust:
        header["CustomerNumber"] = m_cust.group(1)

    # Contact name
    m_contact = re.search(r"Contact\s+([A-Za-z .'-]+)", full_text)
    if m_contact:
        name = m_contact.group(1).strip()
        parts = name.split()
        if len(parts) >= 2:
            header["FirstName"] = parts[0]
            header["LastName"] = " ".join(parts[1:])
        elif parts:
            header["FirstName"] = parts[0]

    # Salesperson -> ReferralManager
    m_sales = re.search(r"Salesperson\s+([A-Za-z .'-]+)", full_text)
    if m_sales:
        header["ReferralManager"] = m_sales.group(1).strip()

    # Address block between Quoted For and Quote Good Through
    if "Quoted For:" in full_text and "Quote Good Through" in full_text:
        start = full_text.index("Quoted For:")
        end = full_text.index("Quote Good Through")
        addr_block = full_text[start:end]

        # Company name
        m_company = re.search(r"Quoted For:\s*(.+?)\s+Ship To:", addr_block)
        if m_company:
            header["Company"] = m_company.group(1).strip()

        # Street address - first address in the block
        m_addr = re.search(
            r"(\d{3,6}\s+[A-Za-z0-9 .]+?)\s+\d{3,6}\s+[A-Za-z0-9 ]+",
            addr_block,
        )
        if m_addr:
            header["Address"] = m_addr.group(1).strip()

        # City, state, zip
        m_city = re.search(
            r"([A-Za-z .]+),\s*([A-Z]{2})\s+(\d{5})(?:-\d{4})?",
            addr_block,
        )
        if m_city:
            header["City"] = m_city.group(1).strip()
            header["State"] = m_city.group(2)
            header["ZipCode"] = m_city.group(3)

        if "United States of America" in addr_block:
            header["Country"] = "USA"

    # Quote valid date
    m_valid = re.search(r"Quote Good Through\s+(\d{1,2}/\d{1,2}/\d{4})", full_text)
    if m_valid:
        header["QuoteValidDate"] = m_valid.group(1)

    return header


def extract_line_items(full_text: str) -> List[Dict[str, str]]:
    """Extract all normal line items based on per-line parsing.

    We look for lines like:
    '1 COP2.750.BLACK 100 FT 33,500.00000 MFT 3,350.00'
    or
    '1 HW.MAGFOOT-170 27 EAC 3,600.00000 EAC 97,200.00'
    and then take the *next* line as the description.
    """
    items: List[Dict[str, str]] = []

    lines = [ln.strip() for ln in full_text.splitlines()]

    pattern = re.compile(
        r"^(\d+)\s+([A-Z0-9.\-]+)\s+(\d+)\s+[A-Z/]+\s+([\d,]+\.\d+)\s*[A-Z/]+\s+([\d,]+\.\d{2})$"
    )

    for idx, line in enumerate(lines):
        m = pattern.match(line)
        if not m:
            continue

        line_no = m.group(1)
        item_id = m.group(2)
        qty = m.group(3)
        unit_price = m.group(4)
        total = m.group(5)

        # Description = the very next non-empty line after this line
        description = ""
        if idx + 1 < len(lines):
            description = lines[idx + 1].strip()

        items.append(
            {
                "line_no": line_no,
                "item_id": item_id,
                "qty": qty,
                "unit_price": unit_price,
                "total": total,
                "description": description,
            }
        )

    return items


def extract_tax_item(full_text: str) -> Optional[Dict[str, str]]:
    """Extract Tax as a separate line item from the summary area, if non-zero."""
    block = full_text
    if "Product" in full_text and "Total" in full_text:
        start = full_text.index("Product")
        end = full_text.index("Total", start)
        block = full_text[start:end]

    m = re.search(r"\bTax\s+([\d,]+\.\d{2})", block)
    if not m:
        return None

    amount_str = m.group(1)
    try:
        val = float(amount_str.replace(",", ""))
    except Exception:
        return None

    if abs(val) < 0.005:
        # Tax is effectively zero -> ignore
        return None

    return {
        "line_no": "TAX",
        "item_id": "Tax",
        "qty": "1",
        "unit_price": amount_str,
        "total": amount_str,
        "description": "Tax",
    }


def normalize_date_str(date_str: Optional[str]) -> Optional[str]:
    """Return date as mm/dd/yyyy string (no special characters)."""
    if not date_str:
        return None
    try:
        dt = datetime.strptime(date_str, "%m/%d/%Y")
        return dt.strftime("%m/%d/%Y")
    except Exception:
        return date_str


def build_rows_for_pdf(
    pdf_bytes: bytes,
    filename: str,
    fallback_referral_manager: Optional[str],
    referral_email: Optional[str],
    brand: Optional[str],
) -> List[Dict]:
    """Parse one PDF and return a list of row dicts following TARGET_COLUMNS."""
    full_text = extract_full_text(pdf_bytes)
    header = extract_header_info(full_text)
    items = extract_line_items(full_text)

    # Append Tax line item if present and non-zero
    tax_item = extract_tax_item(full_text)
    if tax_item:
        items.append(tax_item)

    rows: List[Dict] = []

    quote_number_text = header.get("QuoteNumber")
    quote_date_text = normalize_date_str(header.get("QuoteDate"))
    quote_valid_text = normalize_date_str(header.get("QuoteValidDate"))

    referral_manager = header.get("ReferralManager") or fallback_referral_manager or None

    for it in items:
        try:
            unit_price_val = float(it["unit_price"].replace(",", ""))
        except Exception:
            unit_price_val = None

        try:
            total_val = float(it["total"].replace(",", ""))
        except Exception:
            total_val = None

        row = {
            "ReferralManager": referral_manager,
            "ReferralEmail": referral_email or None,
            "Brand": brand or None,
            "QuoteNumber": quote_number_text,
            "QuoteDate": quote_date_text,
            "Company": header.get("Company"),
            "FirstName": header.get("FirstName"),
            "LastName": header.get("LastName"),
            "ContactEmail": None,
            "ContactPhone": None,
            "Address": header.get("Address"),
            "County": None,
            "City": header.get("City"),
            "State": header.get("State"),
            "ZipCode": header.get("ZipCode"),
            "Country": header.get("Country"),
            "item_id": it["item_id"],
            "item_desc": it["description"],
            "UnitPrice": unit_price_val,
            "TotalSales": total_val,
            "QuoteValidDate": quote_valid_text,
            "CustomerNumber": header.get("CustomerNumber"),
            "manufacturer_Name": None,
            "PDF": filename,
            "DemoQuote": None,
        }
        rows.append(row)

    return rows


# ---------- STREAMLIT UI ----------

st.set_page_config(page_title="Cadre Quote PDF → Excel", layout="wide")

st.title("Cadre Quote PDF → Excel Extractor")

st.markdown(
    """
Upload Cadre Wire quote PDFs and download all line items in a single Excel file.

- Supports up to **100 PDFs** per run.  
- Designed for the **same layout and alignment** as Cadre quotes.  
- Each product line item becomes its own row in Excel.
"""
)

with st.sidebar:
    st.header("Optional Defaults")
    fallback_referral_manager = st.text_input(
        "Fallback Referral Manager (used only if PDF has no Salesperson)",
        value="",
    )
    referral_email = st.text_input(
        "Referral Email (optional)",
        value="",
        help="If provided, this goes into the ReferralEmail column.",
    )
    brand = st.text_input("Brand", value="Cadre Wire Group")

uploaded_files = st.file_uploader(
    "Upload up to 100 quote PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    help="Files must follow the Cadre quote layout (same format / alignment).",
)

process = st.button("Process PDFs")

if process:
    if not uploaded_files:
        st.error("Please upload at least one PDF.")
    elif len(uploaded_files) > 100:
        st.error("Please upload 100 PDFs or fewer at a time.")
    else:
        # Read all files once so we can both parse and offer a ZIP download
        file_data: List[Dict[str, bytes]] = []
        for f in uploaded_files:
            pdf_bytes = f.read()
            file_data.append({"name": f.name, "bytes": pdf_bytes})

        all_rows: List[Dict] = []
        progress = st.progress(0.0)
        status = st.empty()

        for idx, fd in enumerate(file_data, start=1):
            name = fd["name"]
            pdf_bytes = fd["bytes"]
            status.text(f"Processing {idx}/{len(file_data)}: {name}")
            try:
                rows = build_rows_for_pdf(
                    pdf_bytes=pdf_bytes,
                    filename=name,
                    fallback_referral_manager=fallback_referral_manager,
                    referral_email=referral_email,
                    brand=brand,
                )
                all_rows.extend(rows)
            except Exception as e:
                st.warning(f"Error processing {name}: {e}")
            progress.progress(idx / len(file_data))

        if not all_rows:
            st.error("No line items were found in the uploaded PDFs.")
        else:
            df = pd.DataFrame(all_rows, columns=TARGET_COLUMNS)
            st.success(
                f"Parsed {len(file_data)} PDF(s) with {len(df)} total line items."
            )

            st.subheader("Preview (first 50 rows)")
            st.dataframe(df.head(50), use_container_width=True)

            # Excel download
            excel_buf = io.BytesIO()
            with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Quotes")
            excel_buf.seek(0)

            st.download_button(
                label="Download Excel Spreadsheet",
                data=excel_buf,
                file_name="quotes_extracted.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )

            # ZIP of PDFs with same names as in the Excel 'PDF' column
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for fd in file_data:
                    zf.writestr(fd["name"], fd["bytes"])
            zip_buf.seek(0)

            st.download_button(
                label="Download Uploaded PDFs (ZIP)",
                data=zip_buf,
                file_name="quotes_pdfs.zip",
                mime="application/zip",
            )
else:
    st.info("Upload PDFs and click **Process PDFs** to start.")
