import streamlit as st
from datetime import datetime
import re
import os
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Set the page configuration
st.set_page_config(
    page_title="PSS Maker",
    page_icon="favicon.png"
)

# Initialize session state for DOCX storage
if "docx_bytes" not in st.session_state:
    st.session_state.docx_bytes = None
    st.session_state.filename = None

# Hide Streamlit's default UI components
hide_st_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)

# Pre-filled data dictionary
pre_filled_data = {
    "001": {
        "full_name": "Mahendra Tripathi",
        "designation": "Country General Manager & Director",
        "company_name": "Lamberti India Pvt. Ltd.",
        "city_state": "Rajkot, Gujarat",
        "po_id": "LIPL202526????",
        "custom_message": "Sending you Pre-Shipment sample of Guar Gum Powder Modified."
    },
    "002": {
        "full_name": "Mahendra Tripathi",
        "designation": "Country General Manager & Director",
        "company_name": "Lamberti India Pvt. Ltd.",
        "city_state": "Rajkot, Gujarat",
        "po_id": "LIPL2025260115",
        "custom_message": "Sending you Pre-Shipment sample of FARINA GUAR 200 MESH 5000 T/C."
    }
}


def create_docx(date_str, salutation1, full_name, designation, company_name, city_state,
                salutation2, po_id, custom_line, item_details, letterhead_path=None, seal_path=None):
    """
    Build a .docx Document and return bytes.
    date_str expected in DD/MM/YYYY format (string).
    """
    doc = Document()

    # Optional: insert letterhead image (full-width-ish)
    if letterhead_path and os.path.exists(letterhead_path):
        try:
            doc.add_picture(letterhead_path, width=Inches(6.5))
        except Exception:
            # ignore image errors (size/format issues)
            pass

    doc.add_paragraph()

    # Top line: Kindly Att. and Date aligned on same row using a table
    table = doc.add_table(rows=1, cols=2)
    # set some widths (optional; docx may ignore precise widths)
    try:
        table.columns[0].width = Inches(4.0)
        table.columns[1].width = Inches(2.5)
    except Exception:
        pass

    left_cell = table.cell(0, 0)
    right_cell = table.cell(0, 1)

    p_left = left_cell.paragraphs[0]
    run = p_left.add_run("Kindly Att.")
    run.font.size = Pt(10)

    p_right = right_cell.paragraphs[0]
    p_right.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run2 = p_right.add_run(f"Date: {date_str}")
    run2.font.size = Pt(10)

    doc.add_paragraph()

    # Recipient block
    p = doc.add_paragraph()
    run = p.add_run(f"{salutation1} {full_name},")
    run.bold = True
    run.font.size = Pt(10)

    p = doc.add_paragraph(f"({designation})")
    p.runs[0].font.size = Pt(10)

    p = doc.add_paragraph(company_name + ",")
    p.runs[0].font.size = Pt(10)

    p = doc.add_paragraph(city_state)
    p.runs[0].font.size = Pt(10)

    doc.add_paragraph()

    # Greeting and body
    p = doc.add_paragraph()
    r = p.add_run(f"Dear {salutation2},")
    r.font.size = Pt(10)
    doc.add_paragraph()
    doc.add_paragraph(custom_line)
    doc.add_paragraph(f"P.O. ID: {po_id}")
    doc.add_paragraph()

    # Items
    for item_label, (code, weight) in item_details.items():
        try:
            weight_str = f"{float(weight):.2f}"
        except Exception:
            weight_str = str(weight)
        doc.add_paragraph(f"{item_label}) {code} - {weight_str} MT")

    doc.add_paragraph()
    doc.add_paragraph("Kindly acknowledge receipt of the same.")
    doc.add_paragraph()

    # Signature block with seal inserted between Yours Faithfully and Authorised Signatory
    p = doc.add_paragraph()
    run = p.add_run("Yours Faithfully,")
    run.bold = True
    run.font.size = Pt(10)

    # Small vertical gap before the seal
    doc.add_paragraph()

    # Insert seal if available (centered)
    if seal_path and os.path.exists(seal_path):
        try:
            p_seal = doc.add_paragraph()
            p_seal.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_seal = p_seal.add_run()
            # adjust width to taste; default here ~1.5 inches
            run_seal.add_picture(seal_path, width=Inches(1.5))
            # Add spacing after seal
            doc.add_paragraph()
        except Exception:
            # ignore image insertion problems
            pass

    # Continue with signature lines
    doc.add_paragraph("Authorised Signatory")
    doc.add_paragraph("Aravally Processed Agrotech Pvt Ltd")

    # Save to bytes
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()


# UI
st.title("PSS DOCX MAKER")

with st.form("docx_form"):
    # Date handling:
    date_picker = st.date_input("Pick date (calendar)", value=datetime.today())
    default_date_str = date_picker.strftime("%d/%m/%Y")
    date_text = st.text_input("Or edit date (DD/MM/YYYY)", value=default_date_str,
                              help="Enter date in DD/MM/YYYY. If invalid, the calendar selection will be used.")

    salutation1 = st.selectbox("Salutation1", ["Mr.", "Mrs."])
    user_code = st.text_input("Enter Code to auto-fill details (e.g. 001 or 002)")

    if user_code in pre_filled_data:
        data = pre_filled_data[user_code]
        full_name = st.text_input("Full Name", value=data["full_name"])
        designation = st.text_input("Designation", value=data["designation"])
        company_name = st.text_input("Company Name", value=data["company_name"])
        city_state = st.text_input("City, State", value=data["city_state"])
        po_id = st.text_input("P.O. ID", value=data["po_id"])
        custom_line = st.text_input("Pre-Shipment Sample Properties:", value=data.get("custom_message", "Sending you Pre-Shipment sample of"))
    else:
        full_name = st.text_input("Full Name")
        designation = st.text_input("Designation")
        company_name = st.text_input("Company Name")
        city_state = st.text_input("City, State")
        po_id = st.text_input("P.O. ID")
        custom_line = st.text_input("Pre-Shipment Sample Properties:", value="Sending you Pre-Shipment sample of")

    salutation2 = st.selectbox("Salutation2", ["Sir", "Ma’am"])
    total_containers = st.number_input("Total Number of Containers", min_value=1, step=1, value=1)
    current_container = st.number_input("Current Container Number", min_value=1, step=1, value=1)

    num_items = st.selectbox("Number of items", [1, 2, 3, 4, 5, 6], index=0)

    item_details = {}
    item_labels = ['A', 'B', 'C', 'D', 'E', 'F']
    for i in range(num_items):
        code = st.text_input(f"Item {item_labels[i]} Code")
        weight = st.number_input(f"Item {item_labels[i]} Weight (MT)", min_value=0.0, step=0.1, value=4.50, format="%.2f")
        item_details[item_labels[i]] = (code, weight)

    submitted = st.form_submit_button("Generate DOCX")

if submitted:
    # Try to parse the user-edited date_text in DD/MM/YYYY. If parsing fails, fall back to date_picker.
    date_str_final = None
    date_text = date_text.strip()
    date_match = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", date_text)
    if date_match:
        d, m, y = date_match.groups()
        try:
            dt = datetime(int(y), int(m), int(d))
            date_str_final = dt.strftime("%d/%m/%Y")
        except ValueError:
            date_str_final = None

    if not date_str_final:
        date_str_final = date_picker.strftime("%d/%m/%Y")

    # Provide paths for letterhead and seal (adjust filenames if different)
    letterhead_path = "letterhead.png"  # optional; will be used only if exists
    # If you saved the seal at /mnt/data/APAPL SEAL.png in this environment, try that path first,
    # otherwise fallback to a local file name.
    seal_candidates = [
        "/mnt/data/APAPL SEAL.png",
        "APAPL SEAL.png",
        "APAPL_SEAL.png",
        "seal.png"
    ]
    seal_path = None
    for candidate in seal_candidates:
        if os.path.exists(candidate):
            seal_path = candidate
            break

    # File naming
    suffix = "MOD" if user_code == "001" else "FAR" if user_code == "002" else "GEN"
    safe_po_id = re.sub(r'[\/:*?"<>|]', '', po_id or "")
    po_suffix = safe_po_id[-3:] if len(safe_po_id) >= 3 else "000"
    filename = f"PSS LIPL {suffix} {po_suffix} {int(current_container)} of {int(total_containers)}.docx"

    # Create DOCX and store in session
    docx_bytes = create_docx(date_str_final, salutation1, full_name, designation, company_name,
                             city_state, salutation2, po_id, custom_line, item_details,
                             letterhead_path if os.path.exists(letterhead_path) else None,
                             seal_path)

    st.session_state.docx_bytes = docx_bytes
    st.session_state.filename = filename

# Always show download button if DOCX is ready
if st.session_state.docx_bytes:
    st.download_button(
        "Download DOCX",
        st.session_state.docx_bytes,
        file_name=st.session_state.filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
