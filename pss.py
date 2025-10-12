import streamlit as st
from datetime import datetime
import re
import os
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# ---------- Helper: replace text in document ----------
def replace_text_in_paragraph(paragraph, mapping):
    for run in paragraph.runs:
        for key, val in mapping.items():
            if key in run.text:
                run.text = run.text.replace(key, val)

def replace_text_in_table(table, mapping):
    for row in table.rows:
        for cell in row.cells:
            replace_text_in_block(cell, mapping)

def replace_text_in_block(block, mapping):
    for paragraph in block.paragraphs:
        replace_text_in_paragraph(paragraph, mapping)
    for table in getattr(block, "tables", []):
        replace_text_in_table(table, mapping)

def apply_replacements(doc, mapping):
    replace_text_in_block(doc)
    for section in doc.sections:
        header = section.header
        footer = section.footer
        try:
            replace_text_in_block(header, mapping)
        except Exception:
            pass
        try:
            replace_text_in_block(footer, mapping)
        except Exception:
            pass

# ---------- Helper: find template path ----------
def find_template_for_code(code):
    code = (code or "").strip()
    candidates = []
    if code == "001":
        candidates = [
            "MOD PSS.docx",
            "/mnt/data/MOD PSS.docx",
            "templates/MOD PSS.docx"
        ]
    elif code == "002":
        candidates = [
            "FAR PSS.docx",
            "/mnt/data/FAR PSS.docx",
            "templates/FAR PSS.docx"
        ]
    else:
        return None

    for c in candidates:
        if os.path.exists(c):
            return c
    return None

# ---------- Create document from template and mapping ----------
def create_docx_from_template(template_path, mapping):
    doc = Document(template_path)
    apply_replacements(doc, mapping)
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()

# ---------- Fallback: create docx from scratch (if template missing) ----------
def create_docx_fallback(date_str, salutation1, full_name, designation, company_name, city_state,
                         salutation2, po_id, custom_line, letterhead_path=None, seal_path=None):
    doc = Document()
    if letterhead_path and os.path.exists(letterhead_path):
        try:
            doc.add_picture(letterhead_path, width=Inches(6.5))
        except Exception:
            pass
    doc.add_paragraph()

    table = doc.add_table(rows=1, cols=2)
    try:
        table.columns[0].width = Inches(4.0)
        table.columns[1].width = Inches(2.5)
    except Exception:
        pass
    left_cell = table.cell(0, 0)
    right_cell = table.cell(0, 1)
    p_left = left_cell.paragraphs[0]
    p_left.add_run("Kindly Att.").font.size = Pt(10)
    p_right = right_cell.paragraphs[0]
    p_right.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    p_right.add_run(f"Date: {date_str}").font.size = Pt(10)

    doc.add_paragraph()
    doc.add_paragraph(f"{salutation1} {full_name},").runs[0].font.size = Pt(10)
    doc.add_paragraph(f"({designation})").runs[0].font.size = Pt(10)
    doc.add_paragraph(company_name + ",").runs[0].font.size = Pt(10)
    doc.add_paragraph(city_state).runs[0].font.size = Pt(10)
    doc.add_paragraph()
    doc.add_paragraph(f"Dear {salutation2},").runs[0].font.size = Pt(10)
    doc.add_paragraph()
    doc.add_paragraph(custom_line)
    doc.add_paragraph(f"P.O. ID: {po_id}")
    doc.add_paragraph()

    # placeholder list (B1..B4) - left blank in fallback unless you change defaults
    for label in ["A", "B", "C", "D"]:
        doc.add_paragraph(f"{label}) ")

    doc.add_paragraph()
    doc.add_paragraph("Kindly acknowledge receipt of the same.")
    doc.add_paragraph()

    doc.add_paragraph("Yours Faithfully,").runs[0].font.size = Pt(10)
    doc.add_paragraph()

    # insert seal if available
    if seal_path and os.path.exists(seal_path):
        try:
            p_seal = doc.add_paragraph()
            p_seal.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p_seal.add_run().add_picture(seal_path, width=Inches(1.5))
            doc.add_paragraph()
        except Exception:
            pass

    doc.add_paragraph("Authorised Signatory")
    doc.add_paragraph("Aravally Processed Agrotech Pvt Ltd")

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()

# ---------- Streamlit UI ----------
st.set_page_config(page_title="PSS Maker (Template fill)", page_icon="favicon.png")

hide_st_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)

st.title("PSS Maker — Template fill")

if "docx_bytes" not in st.session_state:
    st.session_state.docx_bytes = None
    st.session_state.filename = None

# prefilled data (PO default as requested)
pre_filled_data = {
    "001": {
        "full_name": "Mahendra Tripathi",
        "designation": "Country General Manager & Director",
        "company_name": "Lamberti India Pvt. Ltd.",
        "city_state": "Rajkot, Gujarat",
        "po_id": "PO012",
        "custom_message": "Sending you Pre-Shipment sample of Guar Gum Powder Modified."
    },
    "002": {
        "full_name": "Mahendra Tripathi",
        "designation": "Country General Manager & Director",
        "company_name": "Lamberti India Pvt. Ltd.",
        "city_state": "Rajkot, Gujarat",
        "po_id": "PO012",
        "custom_message": "Sending you Pre-Shipment sample of FARINA GUAR 200 MESH 5000 T/C."
    }
}

with st.form("form"):
    # only calendar date picker (no editable date text)
    date_picker = st.date_input("Pick date (calendar)", value=datetime.today())

    salutation1 = st.selectbox("Salutation1", ["Mr.", "Mrs."])
    user_code = st.text_input("Enter Code to auto-fill details (e.g. 001 or 002)")

    if user_code in pre_filled_data:
        data = pre_filled_data[user_code]
        full_name = st.text_input("Full Name", value=data["full_name"])
        designation = st.text_input("Designation", value=data["designation"])
        company_name = st.text_input("Company Name", value=data["company_name"])
        city_state = st.text_input("City, State", value=data["city_state"])
        po_id = st.text_input("P.O. ID (will be placed into template)", value=data.get("po_id", "PO012"))
        custom_line = st.text_input("Pre-Shipment Sample Properties:", value=data.get("custom_message", "Sending you Pre-Shipment sample of"))
    else:
        full_name = st.text_input("Full Name")
        designation = st.text_input("Designation")
        company_name = st.text_input("Company Name")
        city_state = st.text_input("City, State")
        po_id = st.text_input("P.O. ID (will be placed into template)", value="PO012")
        custom_line = st.text_input("Pre-Shipment Sample Properties:", value="Sending you Pre-Shipment sample of")

    salutation2 = st.selectbox("Salutation2", ["Sir", "Ma’am"])
    total_containers = st.number_input("Total Number of Containers", min_value=1, step=1, value=1)
    current_container = st.number_input("Current Container Number", min_value=1, step=1, value=1)

    submitted = st.form_submit_button("Generate DOCX from template")

if submitted:
    # date formatted as DD/MM/YYYY (no editable input)
    date_str_final = date_picker.strftime("%d/%m/%Y")

    # prepare mapping: date + PO replacements; B1..B4 empty by default
    mapping = {}
    mapping["{{DD/MM/YYYY}}"] = date_str_final
    mapping["DD/MM/YYYY"] = date_str_final

    po_value = po_id.strip() if po_id and po_id.strip() else "PO012"
    mapping["{{PO012}}"] = po_value
    mapping["PO012"] = po_value
    mapping["{{P.O. ID}}"] = po_value
    mapping["{{P.O. ID:}}"] = po_value

    # B1..B4 -> empty strings (no item inputs)
    for i in range(1, 5):
        mapping[f"{{{{B{i}}}}}"] = ""
        mapping[f"B{i}"] = ""
        mapping[f"{{{{{ 'B'+str(i) }}}}}"] = ""  # triple-brace variant

    # find template
    template_path = find_template_for_code(user_code)

    # locate letterhead/seal if present
    letterhead_path = "letterhead.png" if os.path.exists("letterhead.png") else None
    seal_candidates = ["/mnt/data/APAPL SEAL.png", "APAPL SEAL.png", "APAPL_SEAL.png", "seal.png"]
    seal_path = None
    for cand in seal_candidates:
        if os.path.exists(cand):
            seal_path = cand
            break

    # filename
    suffix = "MOD" if user_code == "001" else "FAR" if user_code == "002" else "GEN"
    safe_po = re.sub(r'[\/:*?"<>|]', '', po_value)
    po_suffix = safe_po[-3:] if len(safe_po) >= 3 else "000"
    filename = f"PSS LIPL {suffix} {po_suffix} {int(current_container)} of {int(total_containers)}.docx"

    final_bytes = None
    if template_path:
        try:
            final_bytes = create_docx_from_template(template_path, mapping)
            st.success(f"Template used: {template_path}")
        except Exception as e:
            st.error(f"Error applying template replacements: {e}")
            final_bytes = create_docx_fallback(date_str_final, salutation1, full_name, designation,
                                               company_name, city_state, salutation2, po_value,
                                               custom_line, letterhead_path=letterhead_path,
                                               seal_path=seal_path)
    else:
        st.warning("Template for this code not found — generating fallback DOCX.")
        final_bytes = create_docx_fallback(date_str_final, salutation1, full_name, designation,
                                           company_name, city_state, salutation2, po_value,
                                           custom_line, letterhead_path=letterhead_path,
                                           seal_path=seal_path)

    st.session_state.docx_bytes = final_bytes
    st.session_state.filename = filename

if st.session_state.get("docx_bytes"):
    st.download_button(
        "Download filled DOCX",
        st.session_state.docx_bytes,
        file_name=st.session_state.filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
