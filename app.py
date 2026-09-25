#!/usr/bin/env python3
"""
Mid-Year Performance Review Compiler - Streamlit Web Application
Zero-Credentials / Privacy-First: Drag & Drop Google Form response sheets (.xlsx or .csv)
and generate beautifully styled Word (.docx) documents for each employee.
"""

import io
import os
import re
import zipfile
from collections import defaultdict

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# --- Page Configuration ---
st.set_page_config(
    page_title="Mid-Year Review Compiler",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS styling for premium look
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1B365D;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    .card-box {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1.2rem;
        margin-bottom: 1rem;
    }
    .badge-success {
        background-color: #DEF7EC;
        color: #03543F;
        padding: 0.2rem 0.6rem;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .badge-info {
        background-color: #E1EFFE;
        color: #1E429F;
        padding: 0.2rem 0.6rem;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)


def normalize_name(name_str):
    """Standardizes names for clean case-insensitive matching."""
    if not name_str:
        return ""
    cleaned = re.sub(r'[\(\[\{].*?[\)\]\}]', '', str(name_str))
    cleaned = re.sub(r'\s+', ' ', cleaned).strip().lower()
    return cleaned


def clean_question_text(q_text):
    """
    Cleans Google Form questions by:
    1. Taking ONLY the first main question sentence.
    2. Stripping leading numbers like '4.', '5.', 'Q1.', etc.
    3. Dropping all instruction sub-lines (e.g. 'Please list 3-5...', 'Focus on...').
    """
    if not q_text:
        return ""
    lines = [l.strip() for l in str(q_text).strip().split('\n') if l.strip()]
    if not lines:
        return ""
    first_line = lines[0]
    cleaned = re.sub(r'^\s*(?:Q\s*\d+[\.\:\-]*|\d+[\.\:\)\-]\s*)+', '', first_line).strip()
    return cleaned


def combine_original_peer_feedback(responses):
    """
    Combines the exact original feedback from all nominated peers into a single list
    of structured points, preserving 100% of the original wording, tone, and intent
    without any reviewer labels or alteration.
    """
    combined_points = []
    seen = set()

    for resp in responses:
        if not resp or not str(resp).strip():
            continue
        text = str(resp).strip()
        raw_lines = [l.strip() for l in text.split('\n') if l.strip()]
        for line in raw_lines:
            clean_item = re.sub(r'^\s*(?:\d+[\.\)]|[•\-\*])\s*', '', line).strip()
            clean_item = re.sub(r'[\*\#]+$', '', clean_item).strip()
            if clean_item:
                norm = clean_item.lower()
                if norm not in seen:
                    seen.add(norm)
                    combined_points.append(clean_item)

    return combined_points if combined_points else responses


def style_cell_shading(cell, color_hex="F1F5F9"):
    """Adds background shading to Word table cells."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def generate_word_doc_bytes(metadata, self_qa, peer_grouped):
    """Builds a formatted Word document in-memory and returns its byte stream."""
    doc = Document()

    # Document Margins (1 inch)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Title
    title_p = doc.add_paragraph()
    title_run = title_p.add_run("MID-YEAR PERFORMANCE REVIEW")
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(27, 54, 93)
    title_p.paragraph_format.space_after = Pt(2)

    subtitle_p = doc.add_paragraph()
    sub_run = subtitle_p.add_run("Consolidated Self-Evaluation & Peer Feedback")
    sub_run.font.size = Pt(11)
    sub_run.font.italic = True
    sub_run.font.color.rgb = RGBColor(100, 116, 139)
    subtitle_p.paragraph_format.space_after = Pt(14)

    # Employee Summary Card
    table = doc.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    fields = [
        ("Employee Name:", metadata.get("name", "N/A"), 0, 0),
        ("Department:", metadata.get("department", "N/A"), 0, 1),
        ("Email Address:", metadata.get("email", "N/A"), 1, 0),
        ("Review Date:", metadata.get("date", "N/A"), 1, 1),
    ]

    for label, val, r, c in fields:
        cell = table.cell(r, c)
        style_cell_shading(cell, "F8FAFC")
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        r_lbl = p.add_run(f"{label} ")
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(10)
        r_lbl.font.color.rgb = RGBColor(30, 41, 59)
        r_val = p.add_run(str(val))
        r_val.font.size = Pt(10)
        r_val.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # ==================== PART 1: SELF REVIEW ====================
    h1 = doc.add_heading("Part 1 – Self Review", level=1)
    h1.runs[0].font.color.rgb = RGBColor(27, 54, 93)
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(6)

    if not self_qa:
        p = doc.add_paragraph("No self-review responses submitted.")
        p.runs[0].font.italic = True
        p.runs[0].font.color.rgb = RGBColor(100, 116, 139)
    else:
        for q_idx, (question, answer) in enumerate(self_qa.items(), 1):
            q_clean = clean_question_text(question)

            q_p = doc.add_paragraph()
            q_p.paragraph_format.space_before = Pt(10)
            q_p.paragraph_format.space_after = Pt(2)
            q_run = q_p.add_run(f"Q{q_idx}. {q_clean}")
            q_run.font.bold = True
            q_run.font.size = Pt(11)
            q_run.font.color.rgb = RGBColor(30, 41, 59)

            ans_p = doc.add_paragraph()
            ans_p.paragraph_format.left_indent = Inches(0.2)
            ans_p.paragraph_format.space_after = Pt(8)
            ans_run = ans_p.add_run(str(answer) if answer else "No response provided.")
            ans_run.font.size = Pt(10.5)
            ans_run.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # ==================== PART 2: COMBINED PEER FEEDBACK SUMMARY ====================
    h2 = doc.add_heading("Part 2 – Peer Feedback Summary", level=1)
    h2.runs[0].font.color.rgb = RGBColor(27, 54, 93)
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(2)

    note_p = doc.add_paragraph("Combined summary of feedback from nominated peers.")
    note_p.runs[0].font.italic = True
    note_p.runs[0].font.size = Pt(9.5)
    note_p.runs[0].font.color.rgb = RGBColor(100, 116, 139)
    note_p.paragraph_format.space_after = Pt(6)

    if not peer_grouped:
        p = doc.add_paragraph("No peer review responses submitted for this employee.")
        p.runs[0].font.italic = True
        p.runs[0].font.color.rgb = RGBColor(100, 116, 139)
    else:
        for q_idx, (question, raw_responses) in enumerate(peer_grouped.items(), 1):
            q_clean = clean_question_text(question)

            q_p = doc.add_paragraph()
            q_p.paragraph_format.space_before = Pt(12)
            q_p.paragraph_format.space_after = Pt(3)
            q_run = q_p.add_run(f"Question {q_idx}: {q_clean}")
            q_run.font.bold = True
            q_run.font.size = Pt(11)
            q_run.font.color.rgb = RGBColor(30, 41, 59)

            combined_items = combine_original_peer_feedback(raw_responses)

            for item in combined_items:
                f_p = doc.add_paragraph(style='List Bullet')
                f_p.paragraph_format.left_indent = Inches(0.25)
                f_p.paragraph_format.space_after = Pt(3)
                feed_run = f_p.add_run(str(item))
                feed_run.font.size = Pt(10)
                feed_run.font.color.rgb = RGBColor(51, 65, 85)

    # ==================== PART 3: MANAGER EVALUATION (NEW PAGE) ====================
    doc.add_page_break()
    h3 = doc.add_heading("Part 3 – Manager's Feedback & Evaluation", level=1)
    h3.runs[0].font.color.rgb = RGBColor(27, 54, 93)
    h3.paragraph_format.space_before = Pt(12)
    h3.paragraph_format.space_after = Pt(4)

    mgr_note = doc.add_paragraph("Section for the reporting manager to provide overall summary observations, key achievements, growth areas, and guidance.")
    mgr_note.runs[0].font.italic = True
    mgr_note.runs[0].font.size = Pt(9.5)
    mgr_note.runs[0].font.color.rgb = RGBColor(100, 116, 139)
    mgr_note.paragraph_format.space_after = Pt(8)

    # Shaded container box for manager's feedback
    mgr_table = doc.add_table(rows=1, cols=1)
    mgr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    mgr_cell = mgr_table.cell(0, 0)
    style_cell_shading(mgr_cell, "F8FAFC")
    
    mgr_p = mgr_cell.paragraphs[0]
    mgr_p.paragraph_format.space_before = Pt(8)
    mgr_p.paragraph_format.space_after = Pt(8)
    
    lbl_run = mgr_p.add_run("Manager's Summary & Comments:\n")
    lbl_run.font.bold = True
    lbl_run.font.size = Pt(10.5)
    lbl_run.font.color.rgb = RGBColor(30, 41, 59)

    # Blank spacing for typing comments
    mgr_p.add_run("\n\n\n\n\n")

    # Sign-off block
    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    sign_table = doc.add_table(rows=1, cols=2)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    c1 = sign_table.cell(0, 0)
    c1_p = c1.paragraphs[0]
    r_c1 = c1_p.add_run("Manager Name: __________________________")
    r_c1.font.size = Pt(10)
    r_c1.font.color.rgb = RGBColor(71, 85, 105)

    c2 = sign_table.cell(0, 1)
    c2_p = c2.paragraphs[0]
    r_c2 = c2_p.add_run("Date: _______________")
    r_c2.font.size = Pt(10)
    r_c2.font.color.rgb = RGBColor(71, 85, 105)

    doc_stream = io.BytesIO()
    doc.save(doc_stream)
    doc_stream.seek(0)
    return doc_stream


def load_dataframe(file_obj):
    """Loads CSV or Excel into a pandas DataFrame."""
    if file_obj.name.endswith(".xlsx") or file_obj.name.endswith(".xls"):
        return pd.read_excel(file_obj)
    return pd.read_csv(file_obj)


# --- Streamlit UI ---

st.markdown('<div class="main-header">📝 Mid-Year Review Compiler</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Upload Self-Review and Peer-Review response spreadsheets to generate individual Word (.docx) review reports.</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown("##### 1. Self-Review Responses Sheet")
    self_file = st.file_uploader(
        "Upload Google Form Self-Review (.xlsx or .csv)", 
        type=["xlsx", "xls", "csv"],
        key="self_file"
    )

with col2:
    st.markdown("##### 2. Peer-Review Responses Sheet")
    peer_file = st.file_uploader(
        "Upload Google Form Peer-Review (.xlsx or .csv)", 
        type=["xlsx", "xls", "csv"],
        key="peer_file"
    )

if self_file and peer_file:
    try:
        df_self = load_dataframe(self_file)
        df_peer = load_dataframe(peer_file)

        # Detect columns in Self Review
        self_cols = [str(c).strip() for c in df_self.columns]
        emp_name_col = None
        dept_col = None
        email_col = None
        timestamp_col = None
        self_questions = []

        for c in self_cols:
            c_low = c.lower()
            if any(k in c_low for k in ['employee name', 'your name', 'name']) and not emp_name_col:
                emp_name_col = c
            elif 'department' in c_low and not dept_col:
                dept_col = c
            elif 'email' in c_low and not email_col:
                email_col = c
            elif 'timestamp' in c_low and not timestamp_col:
                timestamp_col = c
            else:
                self_questions.append(c)

        if not emp_name_col and len(self_cols) > 1:
            emp_name_col = self_cols[1]

        # Detect columns in Peer Review
        peer_cols = [str(c).strip() for c in df_peer.columns]
        target_col = None
        reviewer_name_col = None
        reviewer_dept_col = None
        peer_questions = []

        for c in peer_cols:
            c_low = c.lower()
            if any(k in c_low for k in ['colleague', 'reviewing', 'employee you are reviewing', 'nominee', 'peer name', 'candidate', 'person being reviewed', 'select the employee']) and not target_col:
                target_col = c
            elif 'reviewer' in c_low and 'name' in c_low and not reviewer_name_col:
                reviewer_name_col = c
            elif 'reviewer' in c_low and 'department' in c_low and not reviewer_dept_col:
                reviewer_dept_col = c
            elif 'timestamp' in c_low:
                continue
            else:
                peer_questions.append(c)

        if not target_col:
            for c in peer_cols:
                c_low = c.lower()
                if ('colleague' in c_low or 'nominee' in c_low or 'reviewing' in c_low) and 'reviewer' not in c_low:
                    target_col = c
                    break
        if not target_col:
            for c in peer_cols:
                c_low = c.lower()
                if 'name' in c_low and 'reviewer' not in c_low:
                    target_col = c
                    break
        if not target_col and len(peer_cols) > 3:
            target_col = peer_cols[3]

        # Parse Self Reviews
        self_by_emp = {}
        for _, row in df_self.iterrows():
            raw_name = str(row[emp_name_col]).strip() if pd.notna(row[emp_name_col]) else ""
            if not raw_name:
                continue
            norm_name = normalize_name(raw_name)
            dept = str(row[dept_col]).strip() if dept_col and pd.notna(row[dept_col]) else "N/A"
            email = str(row[email_col]).strip() if email_col and pd.notna(row[email_col]) else "N/A"
            date_val = str(row[timestamp_col]).strip().split()[0] if timestamp_col and pd.notna(row[timestamp_col]) else "N/A"

            qa = {}
            for q in self_questions:
                val = str(row[q]).strip() if pd.notna(row[q]) else ""
                qa[q] = val

            self_by_emp[norm_name] = {
                "display_name": raw_name,
                "department": dept,
                "email": email,
                "date": date_val,
                "qa": qa
            }

        # Parse Peer Reviews
        peer_by_emp = defaultdict(lambda: defaultdict(list))
        for _, row in df_peer.iterrows():
            raw_target = str(row[target_col]).strip() if pd.notna(row[target_col]) else ""
            if not raw_target:
                continue
            norm_target = normalize_name(raw_target)

            for q in peer_questions:
                if q in [target_col, reviewer_name_col, reviewer_dept_col]:
                    continue
                if pd.notna(row[q]) and str(row[q]).strip():
                    peer_by_emp[norm_target][q].append(str(row[q]).strip())

        # Collect unique employees
        all_emp_keys = sorted(set(self_by_emp.keys()) | set(peer_by_emp.keys()))

        st.divider()

        # Metrics Card
        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric("Total Candidates", len(all_emp_keys))
        with m2:
            st.metric("Self-Reviews Uploaded", len(self_by_emp))
        with m3:
            st.metric("Peer-Reviewed Candidates", len(peer_by_emp))

        st.write("")

        # Action: Compile All Button
        if st.button("🚀 Compile All Reviews into Word Documents", type="primary", use_container_width=True):
            with st.spinner("Generating Word (.docx) documents..."):
                zip_buffer = io.BytesIO()
                generated_files = []

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for emp_key in all_emp_keys:
                        self_data = self_by_emp.get(emp_key)
                        peer_grouped = peer_by_emp.get(emp_key, {})

                        if self_data:
                            display_name = self_data["display_name"]
                            metadata = {
                                "name": display_name,
                                "email": self_data.get("email", "N/A"),
                                "department": self_data.get("department", "N/A"),
                                "date": self_data.get("date", "N/A")
                            }
                            self_qa = self_data.get("qa", {})
                        else:
                            display_name = emp_key.title()
                            metadata = {
                                "name": display_name,
                                "email": "N/A",
                                "department": "N/A",
                                "date": "N/A"
                            }
                            self_qa = {}

                        doc_stream = generate_word_doc_bytes(metadata, self_qa, peer_grouped)
                        safe_name = re.sub(r'[^\w\-_\. ]', '_', display_name)
                        clean_filename = f"MidYear_Review_{safe_name}.docx"
                        
                        # Add to zip
                        zip_file.writestr(clean_filename, doc_stream.getvalue())
                        generated_files.append((display_name, clean_filename, doc_stream.getvalue()))

                zip_buffer.seek(0)

                st.session_state["zip_data"] = zip_buffer.getvalue()
                st.session_state["generated_files"] = generated_files
                st.success(f"Successfully compiled reviews for {len(all_emp_keys)} employees!")

        # Download Section if generated
        if "zip_data" in st.session_state:
            st.markdown("### 📥 Download Compiled Reviews")
            st.download_button(
                label=f"📦 Download All Reviews as ZIP ({len(all_emp_keys)} Word Documents)",
                data=st.session_state["zip_data"],
                file_name="Compiled_MidYear_Reviews.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

            st.write("")
            with st.expander("📄 View & Download Individual Documents", expanded=True):
                for disp_name, fname, file_bytes in st.session_state.get("generated_files", []):
                    c_name, c_btn = st.columns([3, 1])
                    with c_name:
                        st.markdown(f"**{disp_name}** (`{fname}`)")
                    with c_btn:
                        st.download_button(
                            label="⬇️ Download",
                            data=file_bytes,
                            file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_{fname}"
                        )

    except Exception as e:
        st.error(f"Error reading uploaded spreadsheets: {e}")
        st.info("Please ensure you are uploading the standard Google Form Excel (.xlsx) or CSV files.")

else:
    st.info("👆 Please upload both the **Self-Review** and **Peer-Review** spreadsheets to begin.")
