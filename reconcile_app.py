import streamlit as st
import pandas as pd
import os
import re
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from datetime import datetime
from dateutil.parser import parse as du_parse
from io import BytesIO

# --- Constants & Setup ---
st.set_page_config(page_title="Reconcile Assistant", page_icon="🏨", layout="wide")

RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

MONTH_MAP = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,"jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
MONTH_ABBR = {v: k.title() for k, v in MONTH_MAP.items()}

# --- Helper Functions ---

def normalize_key(val):
    if not val: return ""
    return re.sub(r"[^A-Z0-9]", "", str(val).strip().upper())

def infer_hotel(filename):
    n = filename.lower()
    norm = re.sub(r"[^a-z0-9]", " ", n)
    if "katathani" in norm or " kt " in f" {norm} ": return "KT"
    if "the shore" in norm or " ts " in f" {norm} ": return "TS"
    if "waters" in norm or " wat " in f" {norm} ": return "WAT"
    if "little shore" in norm or " tlkl " in f" {norm} ": return "TLKL"
    if "sands" in norm or " san " in f" {norm} ": return "SAN"
    if "leaf" in norm: return "LFS" # Simplified
    return "Unknown"

def infer_ota(filename):
    n = filename.lower()
    if "booking" in n: return "Booking.com"
    if "expedia" in n or "hotels" in n: return "Expedia"
    if "agoda" in n: return "Agoda"
    return "OTA"

def extract_period(filename):
    n = filename.lower()
    # Try finding Month-Year pattern
    m = re.search(r'\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\W?(\d{2,4})', n)
    if m:
        y = m.group(2)
        y = int(y) if len(y)==4 else int("20"+y)
        return f"{m.group(1).title()}'{str(y)[-2:]}"
    return "Unknown Period"

# --- PDF Parsing Engine (The Fix) ---

def parse_expedia_pdf_stream(pdf_file, use_ocr=False):
    """
    Reads PDF from bytes.
    Strategy 1: Text Extraction + Sanitization (Fixes the broken newline issue)
    Strategy 2: OCR (if enabled and requested)
    """
    data_rows = []
    
    # 1. Try Standard Extraction
    with pdfplumber.open(pdf_file) as pdf:
        full_text = ""
        for page in pdf.pages:
            # If user wants OCR, we skip standard text or try to mix? 
            # For simplicity, if OCR is ON, we assume standard text is garbage.
            if use_ocr:
                continue 
            
            txt = page.extract_text() or ""
            full_text += txt + "\n"

    # 2. If OCR is requested or text is empty, try OCR
    if use_ocr or (not full_text.strip() and not use_ocr):
        try:
            import pytesseract
            from pdf2image import convert_from_bytes
            
            # Convert PDF bytes to images
            images = convert_from_bytes(pdf_file.getvalue())
            full_text = ""
            for img in images:
                # Simple OCR
                full_text += pytesseract.image_to_string(img) + "\n"
        except ImportError:
            st.error("OCR libraries not installed. Please rely on text extraction or install poppler/tesseract.")
            return []
        except Exception as e:
            st.error(f"OCR Failed: {e}")
            return []

    # 3. Process the text (The "Cleaner" Logic)
    # The Problem: PDF splits "25-DEC" into "25\n\n\n-DEC". 
    # Solution: Flatten everything to a single line of tokens first.
    
    # Remove quotes and commas which are just formatting noise in this PDF
    clean_stream = full_text.replace('"', ' ').replace(',', '')
    # Replace newlines with spaces to join broken data
    clean_stream = clean_stream.replace('\n', ' ')
    # Normalize multiple spaces
    clean_stream = re.sub(r'\s+', ' ', clean_stream)

    # Regex to find Booking patterns in the clean stream
    # Pattern: (Expedia Collect) (12345678) ... (25-DEC-2025) ... (1234.56)
    # We look for the "Anchor" which is the Booking ID (8-15 digits) nearby a Date
    
    # Find all IDs
    # We iterate through potential matches in the massive string
    
    # Regex: Look for ID followed reasonably closely by a Date
    # ID: \d{8,15}
    # Date: \d{1,2}-[A-Za-z]{3}-\d{4}
    
    # We scan for the triple: Type ... ID ... Date
    matches = re.finditer(r'(Expedia Collect|Hotel Collect)\s.*?(\d{8,15})\s.*?(\d{1,2}-[A-Za-z]{3}-\d{4})', clean_stream, re.IGNORECASE)
    
    for m in matches:
        bt = m.group(1)
        rid = m.group(2)
        dt_txt = m.group(3)
        
        # Determine Amount: Look ahead from the date
        # The amount is usually a number with a decimal, e.g., 1200.00
        # We grab a chunk of text after the date
        end_pos = m.end()
        chunk = clean_stream[end_pos:end_pos+100] # Look at next 100 chars
        
        # Find all numbers with decimals
        prices = re.findall(r'(\d+\.\d{2})', chunk)
        
        amt, tax, tot = "0.00", "0.00", "0.00"
        if prices:
            # Usually the largest or the last ones are the total
            # In the file: Amount ... Tax ... Total
            if len(prices) >= 2:
                amt = prices[-2] # 2nd to last
                tot = prices[-1] # Last
            else:
                amt = prices[0]
                tot = prices[0]

        # Guest name is tricky in flattened text, skip for now or use placeholder
        guest = "Guest"

        try:
            dt_obj = du_parse(dt_txt)
        except:
            dt_obj = None

        data_rows.append({
            "Booking Type": bt,
            "Reservation ID": rid,
            "Date": dt_obj,
            "Guest": guest,
            "Amount": float(amt),
            "Total": float(tot)
        })

    return data_rows

# --- Main App Logic ---

st.title("🏨 Hotel Reconciliation Assistant")
st.markdown("Easy comparison between **Hoteliers Guru** and **OTA (Booking/Expedia)** files.")

with st.expander("ℹ️ How to use"):
    st.markdown("""
    1. Upload your **Hoteliers Guru** Excel file(s).
    2. Upload your **Commission** file(s) (Excel or PDF).
    3. If PDF reading fails, check the **'Force OCR'** box (requires setup).
    4. Click **Process**. The system will match bookings and generate a report.
    """)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Hoteliers Input")
    hot_files = st.file_uploader("Upload Hoteliers Excel", type=["xlsx", "xlsm"], accept_multiple_files=True)

with col2:
    st.subheader("2. OTA Input")
    ota_files = st.file_uploader("Upload OTA Files (Excel/PDF)", type=["xlsx", "pdf", "csv"], accept_multiple_files=True)
    use_ocr = st.checkbox("Force OCR (Use for image-only PDFs)", help="Only check this if standard processing fails. Requires Tesseract installed on server.")

if st.button("🚀 Process Reconciliation", type="primary"):
    if not hot_files or not ota_files:
        st.error("Please upload files for both sides.")
    else:
        # --- 1. Process Hoteliers ---
        hot_data = {} # Key: Hotel -> List of Rows
        
        with st.spinner("Reading Hoteliers files..."):
            for f in hot_files:
                try:
                    wb = load_workbook(f, data_only=True)
                    ws = wb.active
                    
                    # Find headers
                    headers = {cell.value.lower().strip(): i for i, cell in enumerate(ws[2]) if cell.value}
                    res_idx = -1
                    for k, v in headers.items():
                        if "reservation" in k and "id" in k: res_idx = v; break
                        if "reservation" in k: res_idx = v
                    
                    if res_idx == -1:
                        # Fallback for simple sheet
                         res_idx = 1 # Column B

                    # Iterate
                    for row in ws.iter_rows(min_row=3, values_only=True):
                        rid = normalize_key(row[res_idx]) if res_idx < len(row) else ""
                        if rid:
                            hotel_code = infer_hotel(f.name)
                            if hotel_code not in hot_data: hot_data[hotel_code] = []
                            hot_data[hotel_code].append({
                                "id": rid,
                                "raw": row,
                                "file": f.name
                            })
                except Exception as e:
                    st.error(f"Error reading {f.name}: {e}")

        # --- 2. Process OTA ---
        results = []
        
        with st.spinner("Processing OTA files & Matching..."):
            for f in ota_files:
                hotel_code = infer_hotel(f.name)
                period = extract_period(f.name)
                ota_type = infer_ota(f.name)
                
                # Extract OTA Data
                ota_bookings = [] # List of dicts {id, amount, etc}
                
                if f.name.lower().endswith(".pdf"):
                    # Use our PDF engine
                    rows = parse_expedia_pdf_stream(f, use_ocr=use_ocr)
                    for r in rows:
                        ota_bookings.append(r)
                else:
                    # Excel OTA
                    try:
                        wb = load_workbook(f, data_only=True)
                        ws = wb.active
                        # Simple header detection
                        headers = {str(cell.value).lower(): i for i, cell in enumerate(ws[1]) if cell.value}
                        
                        # Find ID col
                        id_col = -1
                        for h, i in headers.items():
                            if "reservation" in h: id_col = i; break
                        
                        # Find Amount cols (for Booking.com logic)
                        final_col = -1
                        comm_col = -1
                        for h, i in headers.items():
                            if "final" in h and "amount" in h: final_col = i
                            if "commission" in h and "amount" in h: comm_col = i

                        for row in ws.iter_rows(min_row=2, values_only=True):
                            if id_col != -1 and id_col < len(row):
                                rid = normalize_key(row[id_col])
                                if rid:
                                    f_amt = row[final_col] if final_col != -1 else 0
                                    c_amt = row[comm_col] if comm_col != -1 else 0
                                    ota_bookings.append({
                                        "Reservation ID": rid,
                                        "Final Amount": f_amt,
                                        "Commission Amount": c_amt,
                                        "raw": row
                                    })
                    except Exception as e:
                        st.error(f"Error reading Excel OTA {f.name}: {e}")

                # --- 3. Match ---
                # Get relevant Hoteliers data
                hot_list = hot_data.get(hotel_code, [])
                if not hot_list:
                    # Try finding "Unknown" or fallback
                    hot_list = hot_data.get("Unknown", [])
                
                hot_map = {item["id"]: item for item in hot_list}
                
                # Logic
                matched = []
                only_ota = []
                only_hot = list(hot_map.keys()) # Will remove found ones
                
                for ob in ota_bookings:
                    oid = normalize_key(ob.get("Reservation ID"))
                    
                    # Booking.com Special Filter: Ignore if Final & Comm are 0
                    if ota_type == "Booking.com":
                        try:
                            f_val = float(ob.get("Final Amount", 0) or 0)
                            c_val = float(ob.get("Commission Amount", 0) or 0)
                            if f_val == 0 or c_val == 0:
                                continue # Skip this booking, don't flag as missing
                        except:
                            pass

                    if oid in hot_map:
                        matched.append(oid)
                        if oid in only_hot: only_hot.remove(oid)
                    else:
                        only_ota.append(oid)
                
                results.append({
                    "file": f.name,
                    "hotel": hotel_code,
                    "ota": ota_type,
                    "period": period,
                    "matched": len(matched),
                    "missing_in_hoteliers": len(only_ota),
                    "missing_in_ota": len(only_hot),
                    "details_ota": only_ota,
                    "details_hot": only_hot
                })

        # --- 4. Display & Download ---
        st.success("Analysis Complete!")
        
        # Summary Table
        st.subheader("Summary")
        df_res = pd.DataFrame(results)
        st.dataframe(df_res[["file", "hotel", "ota", "matched", "missing_in_hoteliers", "missing_in_ota"]])
        
        # Create Excel Report
        output_buffer = BytesIO()
        wb_out = Workbook()
        wb_out.remove(wb_out.active)
        
        for res in results:
            ws_name = f"{res['hotel']}-{res['ota']}"[:30]
            ws = wb_out.create_sheet(ws_name)
            
            ws.append(["File", res['file']])
            ws.append(["Period", res['period']])
            ws.append([])
            
            ws.append(["Category", "Reservation ID", "Status"])
            
            for item in res['details_ota']:
                cell = ws.cell(row=ws.max_row+1, column=1, value="Missing in Hoteliers")
                ws.cell(row=ws.max_row, column=2, value=str(item))
                for c in range(1, 4): ws.cell(row=ws.max_row, column=c).fill = RED_FILL
            
            for item in res['details_hot']:
                cell = ws.cell(row=ws.max_row+1, column=1, value="Missing in OTA")
                ws.cell(row=ws.max_row, column=2, value=str(item))
                for c in range(1, 4): ws.cell(row=ws.max_row, column=c).fill = GREEN_FILL
                
        wb_out.save(output_buffer)
        
        st.download_button(
            label="📥 Download Full Reconciliation Report (Excel)",
            data=output_buffer.getvalue(),
            file_name="Reconciliation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )