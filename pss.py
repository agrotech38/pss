import streamlit as st
from datetime import datetime
import re
import os
from io import BytesIO
from docx import Document
from docx.shared import Pt
import requests

# ----------------- Utility: replacement helpers -----------------
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
        try:
            replace_text_in_block(section.header, mapping)
        except Exception:
            pass
        try:
            replace_text_in_block(section.footer, mapping)
        except Exception:
            pass

# ----------------- Template selection / fetching -----------------
def find_local_template_for_code(code):
    code = (code or "").strip()
    candidates = []
    if code == "001":
        candidates = ["MOD PSS.docx", "/mnt/data/MOD PSS.docx", "templates/MOD PSS.docx"]
    elif code == "002":
        candidates = ["FAR PSS.docx", "/mnt/data/FAR PSS.docx", "templates/FAR PSS.docx"]
    else:
        return None
    for c in candidates:
        if os.path.exists(c):
            return c
    return None

def download_template_from_url(url):
    """
    Download a raw .docx from a given URL and return bytes, or None on failure.
    """
    try:
        resp = requests.get(url, timeout=20)
        resp.raise_for_status()
        return resp.content
    except Exception:
        return None

def load_docx_from_bytes(bts):
    bio = BytesIO(bts)
    return Document(bio)

# ----------------- Create filled docx -----------------
def create_docx_from_template_bytes(template_bytes, mapping):
    doc = load_docx_from_bytes(template_bytes)
    apply_replacements(doc, mapping)
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()

def create_docx_from_template_file(path, mapping):
    doc = Document(path)
    apply_replacements(doc, mapping)
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()

# ----------------- Streamlit UI -----------------
st.set_page_config(page_title="PSS Template Filler (DOCX only)")

st.title("PSS Template Filler — DOCX only (no images)")

# session storage
if "docx_bytes" not in st.session_state:
    st.session_state.docx_bytes = None
    st.session_state.filename = None

st.markdown("Provide code (001 or 002). If templates are hosted on GitHub, optionally paste the raw template URL(s).")

with st.form("form"):
    date_picker = st.date_input("Pick date (calendar)", value=datetime.today())

    salutation1 = st.selectbox("Salutation1", ["Mr.", "Mrs."])
    user_code = st.text_input("Template Code (e.g. 001 or 002)", value="001").strip()

    # Optional: input raw URLs for templates (GitHub raw links). Leave blank to use local files.
    st.info("If your templates are stored in a GitHub repo, paste the raw file URL below (optional). Example raw URL: https://raw.githubusercontent.com/username/repo/branch/path/MOD%20PSS.docx")
    url_mod = st.text_input("MOD PSS.docx raw URL (for code 001)", value="")
    url_far = st.text_input("FAR PSS.docx raw URL (for code 002)", value="")

    # User details to place into template placeholders
    full_name = st.text_input("Full Name", value="Mahendra Tripathi")
    designation = st.text_input("Designation", value="Country General Manager & Director")
    company_name = st.text_input("Company Name", value="Lamberti India Pvt. Ltd.")
    city_state = st.text_input("City, State", value="Rajkot, Gujarat")

    po_id = st.text_input("P.O. ID (will replace PO012)", value="PO012")
    custom_line = st.text_input("Pre-Shipment Sample Properties", value="Sending you Pre-Shipment sample of Guar Gum Powder Modified.")
    salutation2 = st.selectbox("Salutation2", ["Sir", "Ma’am"])

    # B1..B4 placeholders (left empty by default)
    st.markdown("Batch placeholders (B1..B4). Leave blank if not needed.")
    b1 = st.text_input("B1", value="")
    b2 = st.text_input("B2", value="")
    b3 = st.text_input("B3", value="")
    b4 = st.text_input("B4", value="")

    total_containers = st.number_input("Total Number of Containers", min_value=1, step=1, value=1)
    current_container = st.number_input("Current Container Number", min_value=1, step=1, value=1)

    submitted = st.form_submit_button("Generate DOCX")

if submitted:
    # format date
    date_str_final = date_picker.strftime("%d/%m/%Y")

    # Prepare mapping of common placeholder variants to values
    mapping = {}
    mapping["{{DD/MM/YYYY}}"] = date_str_final
    mapping["DD/MM/YYYY"] = date_str_final

    po_value = po_id.strip() if po_id and po_id.strip() else "PO012"
    mapping["{{PO012}}"] = po_value
    mapping["PO012"] = po_value
    mapping["{{P.O. ID}}"] = po_value
    mapping["{{P.O. ID:}}"] = po_value
    mapping["{{PO_ID}}"] = po_value
    mapping["PO_ID"] = po_value

    # Basic personal details (in case template uses them)
    mapping["{{FULL_NAME}}"] = full_name
    mapping["FULL_NAME"] = full_name
    mapping["{{DESIGNATION}}"] = designation
    mapping["DESIGNATION"] = designation
    mapping["{{COMPANY_NAME}}"] = company_name
    mapping["COMPANY_NAME"] = company_name
    mapping["{{CITY_STATE}}"] = city_state
    mapping["CITY_STATE"] = city_state
    mapping["{{CUSTOM_LINE}}"] = custom_line
    mapping["CUSTOM_LINE"] = custom_line
    mapping["{{SALUTATION2}}"] = salutation2
    mapping["SALUTATION2"] = salutation2

    # B1..B4
    mapping["{{B1}}"] = b1
    mapping["B1"] = b1
    mapping["{{B2}}"] = b2
    mapping["B2"] = b2
    mapping["{{B3}}"] = b3
    mapping["B3"] = b3
    mapping["{{B4}}"] = b4
    mapping["B4"] = b4

    # find local template first
    template_path = find_local_template_for_code(user_code)

    # If not found and user provided URL, download from appropriate URL
    template_bytes = None
    if template_path:
        try:
            final_bytes = create_docx_from_template_file(template_path, mapping)
            used_template_info = f"Local template used: {template_path}"
        except Exception as e:
            st.error(f"Error applying replacements to local template: {e}")
            final_bytes = None
            used_template_info = "Local template failed"
    else:
        # Try to download from provided URLs
        url_candidate = None
        if user_code == "001" and url_mod.strip():
            url_candidate = url_mod.strip()
        elif user_code == "002" and url_far.strip():
            url_candidate = url_far.strip()

        if url_candidate:
            bts = download_template_from_url(url_candidate)
            if bts:
                try:
                    final_bytes = create_docx_from_template_bytes(bts, mapping)
                    used_template_info = f"Downloaded template used from URL: {url_candidate}"
                except Exception as e:
                    st.error(f"Error applying replacements to downloaded template: {e}")
                    final_bytes = None
                    used_template_info = "Downloaded template failed"
            else:
                final_bytes = None
                used_template_info = f"Failed to download template from URL: {url_candidate}"
        else:
            final_bytes = None
            used_template_info = "No template found locally or URL provided"

    # If we still don't have final_bytes, inform user and do not create images or fallback heavy doc.
    if not final_bytes:
        st.error(f"Could not locate or process a template. {used_template_info}. Place 'MOD PSS.docx' or 'FAR PSS.docx' next to app.py or provide a raw GitHub URL.")
    else:
        # compute output filename
        suffix = "MOD" if user_code == "001" else "FAR" if user_code == "002" else "GEN"
        safe_po = re.sub(r'[\/:*?"<>|]', '', po_value)
        po_suffix = safe_po[-3:] if len(safe_po) >= 3 else "000"
        filename = f"PSS LIPL {suffix} {po_suffix} {int(current_container)} of {int(total_containers)}.docx"

        st.session_state.docx_bytes = final_bytes
        st.session_state.filename = filename
        st.success(used_template_info)

if st.session_state.get("docx_bytes"):
    st.download_button(
        "Download filled DOCX",
        st.session_state.docx_bytes,
        file_name=st.session_state.filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
