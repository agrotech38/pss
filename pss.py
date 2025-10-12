import streamlit as st
from datetime import datetime
import re
import os
from io import BytesIO
from docx import Document

# ----------------- Replacement helpers -----------------
def replace_in_paragraph_by_text(paragraph, mapping):
    """
    Replace tokens by working with the full paragraph text.
    This handles tokens split across runs because it rewrites the whole paragraph text.
    """
    text = paragraph.text
    new_text = text
    for key, val in mapping.items():
        if key in new_text:
            new_text = new_text.replace(key, val)
    if new_text != text:
        # assign new text (this replaces runs)
        paragraph.text = new_text

def replace_text_in_table(table, mapping):
    for row in table.rows:
        for cell in row.cells:
            replace_text_in_block(cell, mapping)

def replace_text_in_block(block, mapping):
    """
    Replace tokens in a Document, Header, Footer, or _Cell block.
    """
    for paragraph in getattr(block, "paragraphs", []):
        replace_in_paragraph_by_text(paragraph, mapping)
    for table in getattr(block, "tables", []):
        replace_text_in_table(table, mapping)

def apply_replacements(doc, mapping):
    # body
    replace_text_in_block(doc, mapping)
    # headers/footers
    for section in doc.sections:
        try:
            replace_text_in_block(section.header, mapping)
        except Exception:
            pass
        try:
            replace_text_in_block(section.footer, mapping)
        except Exception:
            pass

# ----------------- Template lookup -----------------
def find_local_template_for_code(code):
    code = (code or "").strip()
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

# ----------------- Create filled docx -----------------
def create_docx_from_template_file(path, mapping):
    doc = Document(path)
    apply_replacements(doc, mapping)
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()

# ----------------- Streamlit UI -----------------
st.set_page_config(page_title="PSS Generator")
st.title("PSS Generator")

# session storage
if "docx_bytes" not in st.session_state:
    st.session_state.docx_bytes = None
    st.session_state.filename = None

with st.form("form"):
    date_picker = st.date_input("Calendar Date", value=datetime.today())

    user_code = st.text_input("Enter the Code", value="000").strip()

    po_id = st.text_input("P.O. ID", value="LIPL202526")
    b1 = st.text_input("Batch 1", value="")
    b2 = st.text_input("Batch 2", value="")
    b3 = st.text_input("Batch 3", value="")
    b4 = st.text_input("Batch 4", value="")

    total_containers = st.number_input("Total Container", min_value=1, step=1, value=1)
    current_container = st.number_input("Current Container", min_value=1, step=1, value=1)

    submitted = st.form_submit_button("Generate")

if submitted:
    # Format date as DD/MM/YYYY
    date_str_final = date_picker.strftime("%d/%m/%Y")

    # Build mapping
    mapping = {}
    mapping["{{DD/MM/YYYY}}"] = date_str_final
    mapping["DD/MM/YYYY"] = date_str_final

    po_value = po_id.strip() if po_id and po_id.strip() else "PO012"
    for key in ["{{PO012}}"]:
        mapping[key] = po_value

    mapping["{{B1}}"] = b1
    mapping["B1"] = b1
    mapping["{{B2}}"] = b2
    mapping["B2"] = b2
    mapping["{{B3}}"] = b3
    mapping["B3"] = b3
    mapping["{{B4}}"] = b4
    mapping["B4"] = b4

    template_path = find_local_template_for_code(user_code)

    if not template_path:
        st.error(
            "Template for this code not found. Place the appropriate template next to the app or in /mnt/data/.\n"
            "Use Appropriate Code"
        )
    else:
        try:
            final_bytes = create_docx_from_template_file(template_path, mapping)

            # filename
            suffix = "MOD" if user_code == "001" else "FAR" if user_code == "002" else "GEN"
            safe_po = re.sub(r'[\/:*?"<>|]', '', po_value)
            po_suffix = safe_po[-3:] if len(safe_po) >= 3 else "000"
            filename = f"PSS {suffix} LIPL {po_suffix} {int(current_container)} of {int(total_containers)}.docx"

            st.session_state.docx_bytes = final_bytes
            st.session_state.filename = filename
            st.success(f"Template {template_path} is being used")
        except Exception as e:
            st.error(f"Failed to process template: {e}")

if st.session_state.get("docx_bytes"):
    st.download_button(
        "Download",
        st.session_state.docx_bytes,
        file_name=st.session_state.filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
