#!/usr/bin/env python3
"""
Mid-Year Performance Review Compilation Script
Connects directly to Google Drive & Google Sheets to compile employee self-reviews
and question-aggregated anonymous peer reviews into styled Word (.docx) documents.
Supports Master Spreadsheets with multiple employee rows and automated Google Drive folder uploads.
"""

import os
import sys
import io
import re
import argparse
from collections import defaultdict

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
except ImportError:
    print("Error: 'python-docx' is required. Please install it using: pip install python-docx")
    sys.exit(1)

# Default Google Spreadsheet IDs from sample forms
DEFAULT_PEER_SHEET_ID = "1sX-7XTFMzyz59XGv_nalEgz0N-_6Ssx8tUxVItAMf5Q"
DEFAULT_SELF_SHEET_ID = "1Xfd2Cp2VuutKWDCtnSYuIu-1pZU5yes1rQ0zi52AjIg"

# Scopes needed for Google Drive and Google Sheets APIs
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets.readonly'
]


def get_google_services(credentials_path="credentials.json", token_path="token.json"):
    """Authenticates and returns Drive and Sheets service clients."""
    creds = None
    if os.path.exists(token_path):
        try:
            creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        except Exception as e:
            print(f"Warning: Could not load {token_path}: {e}")

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Token refresh failed: {e}. Re-authenticating...")
                creds = None

        if not creds:
            if not os.path.exists(credentials_path):
                raise FileNotFoundError(
                    f"Credentials file '{credentials_path}' not found. Please place credentials.json in working directory."
                )
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_path, 'w') as token:
            token.write(creds.to_json())

    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)
    return drive_service, sheets_service


def normalize_name(name_str):
    """Cleans and standardizes names for robust case-insensitive matching."""
    if not name_str:
        return ""
    cleaned = re.sub(r'[\(\[\{].*?[\)\]\}]', '', str(name_str))
    cleaned = re.sub(r'\s+', ' ', cleaned).strip().lower()
    return cleaned


def read_sheet_data(sheets_service, spreadsheet_id):
    """Reads all rows from the first tab of the spreadsheet."""
    try:
        sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', [])
        if not sheets:
            return []
        first_sheet_title = sheets[0]['properties']['title']
        
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, 
            range=f"'{first_sheet_title}'!A1:ZZ"
        ).execute()
        return result.get('values', [])
    except Exception as e:
        print(f"    [!] Error reading spreadsheet {spreadsheet_id}: {e}")
        return []


def parse_master_self_reviews(rows):
    """
    Parses Master Self Review Sheet.
    Returns: dict[normalized_employee_name] = {
        'display_name': str,
        'department': str,
        'email': str,
        'date': str,
        'qa': dict[question, answer]
    }
    """
    if not rows or len(rows) < 2:
        return {}

    headers = [str(h).strip() for h in rows[0]]
    
    name_col_idx = None
    dept_col_idx = None
    email_col_idx = None
    timestamp_col_idx = None
    question_cols = []

    for idx, h in enumerate(headers):
        h_low = h.lower()
        if 'employee name' in h_low or h_low == 'name' or 'your name' in h_low:
            name_col_idx = idx
        elif 'department' in h_low:
            dept_col_idx = idx
        elif 'email' in h_low:
            email_col_idx = idx
        elif 'timestamp' in h_low:
            timestamp_col_idx = idx
        else:
            question_cols.append((idx, h))

    if name_col_idx is None and len(headers) > 1:
        name_col_idx = 1

    employees_self = {}

    for row in rows[1:]:
        if not row or name_col_idx >= len(row):
            continue
        
        raw_name = str(row[name_col_idx]).strip()
        if not raw_name:
            continue
        
        norm_name = normalize_name(raw_name)
        dept = str(row[dept_col_idx]).strip() if dept_col_idx is not None and dept_col_idx < len(row) else "N/A"
        email = str(row[email_col_idx]).strip() if email_col_idx is not None and email_col_idx < len(row) else "N/A"
        date_val = str(row[timestamp_col_idx]).strip().split()[0] if timestamp_col_idx is not None and timestamp_col_idx < len(row) else "N/A"

        qa = {}
        for q_idx, q_text in question_cols:
            ans = str(row[q_idx]).strip() if q_idx < len(row) else ""
            qa[q_text] = ans

        employees_self[norm_name] = {
            "display_name": raw_name,
            "department": dept,
            "email": email,
            "date": date_val,
            "qa": qa
        }

    return employees_self


def parse_master_peer_reviews(rows):
    """
    Parses Master Peer Review Sheet.
    Returns: dict[normalized_reviewee_name] = list of {
        'reviewer_name': str,
        'reviewer_dept': str,
        'relationship': str,
        'responses': dict[question, answer]
    }
    """
    if not rows or len(rows) < 2:
        return defaultdict(list)

    headers = [str(h).strip() for h in rows[0]]

    reviewer_name_col = None
    reviewer_dept_col = None
    reviewee_name_col = None
    relationship_col = None
    question_cols = []

    for idx, h in enumerate(headers):
        h_low = h.lower()
        if 'colleague' in h_low or 'reviewing' in h_low or 'employee you are reviewing' in h_low:
            reviewee_name_col = idx
        elif 'reviewer' in h_low and 'name' in h_low or 'your (reviewer) name' in h_low:
            reviewer_name_col = idx
        elif 'reviewer' in h_low and 'department' in h_low or 'your (reviewer) department' in h_low:
            reviewer_dept_col = idx
        elif 'how have you worked' in h_low or 'relationship' in h_low or 'working relationship' in h_low:
            relationship_col = idx
        elif 'timestamp' in h_low:
            continue
        else:
            if reviewer_name_col is None and 'your name' in h_low:
                reviewer_name_col = idx
            elif reviewer_dept_col is None and 'department' in h_low:
                reviewer_dept_col = idx
            else:
                question_cols.append((idx, h))

    if reviewer_name_col is None and len(headers) > 1:
        reviewer_name_col = 1
    if reviewer_dept_col is None and len(headers) > 2:
        reviewer_dept_col = 2
    if reviewee_name_col is None and len(headers) > 3:
        reviewee_name_col = 3

    peer_reviews_by_reviewee = defaultdict(list)

    for row in rows[1:]:
        if not row or reviewee_name_col >= len(row):
            continue
        
        raw_reviewee = str(row[reviewee_name_col]).strip()
        if not raw_reviewee:
            continue
        
        norm_reviewee = normalize_name(raw_reviewee)
        rev_name = str(row[reviewer_name_col]).strip() if reviewer_name_col is not None and reviewer_name_col < len(row) else "Peer Reviewer"
        rev_dept = str(row[reviewer_dept_col]).strip() if reviewer_dept_col is not None and reviewer_dept_col < len(row) else ""
        rel = str(row[relationship_col]).strip() if relationship_col is not None and relationship_col < len(row) else ""

        responses = {}
        for q_idx, q_text in question_cols:
            if q_idx == reviewer_name_col or q_idx == reviewer_dept_col or q_idx == reviewee_name_col or q_idx == relationship_col:
                continue
            ans = str(row[q_idx]).strip() if q_idx < len(row) else ""
            responses[q_text] = ans

        peer_reviews_by_reviewee[norm_reviewee].append({
            "reviewer_name": rev_name,
            "reviewer_dept": rev_dept,
            "relationship": rel,
            "responses": responses
        })

    return peer_reviews_by_reviewee


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
        # Split by explicit newlines
        raw_lines = [l.strip() for l in text.split('\n') if l.strip()]
        for line in raw_lines:
            # Strip leading bullet/numbering e.g. "1. ", "• ", "- "
            clean_item = re.sub(r'^\s*(?:\d+[\.\)]|[•\-\*])\s*', '', line).strip()
            # Strip trailing markdown asterisks if any
            clean_item = re.sub(r'[\*\#]+$', '', clean_item).strip()
            if clean_item:
                norm = clean_item.lower()
                if norm not in seen:
                    seen.add(norm)
                    combined_points.append(clean_item)

    return combined_points if combined_points else responses


def group_peer_responses_question_wise(peer_submissions):
    """
    Extracts all raw peer responses for each question into a list:
    dict[question] = [response_str_1, response_str_2, ...]
    """
    grouped = defaultdict(list)
    for sub in peer_submissions:
        for question, answer in sub["responses"].items():
            if answer and str(answer).strip():
                grouped[question].append(str(answer).strip())
    return grouped


def list_employee_folders(drive_service, parent_folder_id):
    """Fetches all employee/candidate subfolders inside the root folder (with full pagination)."""
    folders = []
    page_token = None
    query = f"'{parent_folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"

    while True:
        results = drive_service.files().list(
            q=query, 
            fields="nextPageToken, files(id, name)",
            pageSize=100,
            pageToken=page_token
        ).execute()
        
        folders.extend(results.get('files', []))
        page_token = results.get('nextPageToken')
        if not page_token:
            break

    return folders


def style_cell_shading(cell, color_hex="F1F5F9"):
    """Adds soft background shading to a table cell."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def clean_question_text(q_text):
    """
    Cleans Google Form questions by:
    1. Taking ONLY the first main question sentence/line.
    2. Stripping leading numbers like '4.', '5.', 'Q1.', etc.
    3. Dropping all instruction sub-lines (e.g. 'Please list 3-5...', 'Focus on...').
    """
    if not q_text:
        return ""
    lines = [l.strip() for l in q_text.strip().split('\n') if l.strip()]
    if not lines:
        return ""
    first_line = lines[0]
    # Remove leading numbering like '4. ', '4) ', '4: ', 'Q1. ', '1. '
    cleaned = re.sub(r'^\s*(?:Q\s*\d+[\.\:\-]*|\d+[\.\:\)\-]\s*)+', '', first_line).strip()
    return cleaned


def create_word_doc(metadata, self_qa, peer_grouped, output_filepath):
    """Builds a professionally formatted Word document with Self and Combined Peer reviews."""
    doc = Document()

    # Document Margins (1 inch)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Document Header Title
    title_p = doc.add_paragraph()
    title_run = title_p.add_run("MID-YEAR PERFORMANCE REVIEW")
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(27, 54, 93) # Deep Navy
    title_p.paragraph_format.space_after = Pt(2)

    subtitle_p = doc.add_paragraph()
    sub_run = subtitle_p.add_run("Consolidated Self-Evaluation & Peer Feedback")
    sub_run.font.size = Pt(11)
    sub_run.font.italic = True
    sub_run.font.color.rgb = RGBColor(100, 116, 139) # Slate
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
        r_val = p.add_run(val)
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
            ans_run = ans_p.add_run(answer if answer else "No response provided.")
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

            # Combine original peer points under this question
            combined_items = combine_original_peer_feedback(raw_responses)

            for item in combined_items:
                f_p = doc.add_paragraph(style='List Bullet')
                f_p.paragraph_format.left_indent = Inches(0.25)
                f_p.paragraph_format.space_after = Pt(3)
                
                feed_run = f_p.add_run(item)
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

    doc.save(output_filepath)
    print(f"  [+] Saved Word Document: {output_filepath}")


def upload_docx_to_drive(drive_service, file_path, folder_id, file_name):
    """Uploads the generated .docx directly to the target Google Drive / Shared Drive folder."""
    try:
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }
        media = MediaIoBaseUpload(
            io.FileIO(file_path, 'rb'),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            resumable=True
        )
        file = drive_service.files().create(
            body=file_metadata, 
            media_body=media, 
            fields='id',
            supportsAllDrives=True
        ).execute()
        print(f"    [+] Uploaded to Google Drive Folder: File ID {file.get('id')}")
    except Exception as e:
        print(f"    [!] Failed to upload to Google Drive: {e}")


def main():
    parser = argparse.ArgumentParser(description="Compile Mid-Year Performance Reviews from Google Sheets/Drive")
    parser.add_argument("--self-sheet-id", type=str, default=DEFAULT_SELF_SHEET_ID, help="Spreadsheet ID for Self Reviews")
    parser.add_argument("--peer-sheet-id", type=str, default=DEFAULT_PEER_SHEET_ID, help="Spreadsheet ID for Peer Reviews")
    parser.add_argument("--folder-id", type=str, default="", help="Optional: Google Drive / Shared Drive Folder ID to upload compiled docx files into")
    parser.add_argument("--output-dir", type=str, default="./compiled_reviews", help="Local directory to save generated Word documents")
    parser.add_argument("--no-upload", action="store_true", help="Do not upload .docx back to Google Drive (local only)")
    parser.add_argument("--credentials", type=str, default="credentials.json", help="Path to Google credentials.json")
    parser.add_argument("--token", type=str, default="token.json", help="Path to Google token.json")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)

    print("=" * 70)
    print(" MID-YEAR PERFORMANCE REVIEW COMPILATION PIPELINE")
    print("=" * 70)

    try:
        drive_service, sheets_service = get_google_services(
            credentials_path=args.credentials, 
            token_path=args.token
        )
    except Exception as e:
        print(f"Authentication Error: {e}")
        sys.exit(1)

    # 1. Fetch & Parse Self Review Rows
    print(f"Reading Self Review Spreadsheet: {args.self_sheet_id}...")
    self_raw_rows = read_sheet_data(sheets_service, args.self_sheet_id)
    self_reviews_by_emp = parse_master_self_reviews(self_raw_rows)
    print(f"  -> Found self-reviews for {len(self_reviews_by_emp)} employee(s): {', '.join(e['display_name'] for e in self_reviews_by_emp.values())}")

    # 2. Fetch & Parse Peer Review Rows
    print(f"\nReading Peer Review Spreadsheet: {args.peer_sheet_id}...")
    peer_raw_rows = read_sheet_data(sheets_service, args.peer_sheet_id)
    peer_reviews_by_emp = parse_master_peer_reviews(peer_raw_rows)
    print(f"  -> Found peer reviews for {len(peer_reviews_by_emp)} colleague(s): {', '.join(peer_reviews_by_emp.keys())}")

    # 3. Collect list of all unique employees across self and peer reviews
    all_employee_keys = set(self_reviews_by_emp.keys()) | set(peer_reviews_by_emp.keys())
    print(f"\nTotal unique employees to compile: {len(all_employee_keys)}")

    if args.folder_id and not args.no_upload:
        print(f"Target Google Drive Folder for uploads: {args.folder_id}")

    # 4. Compile Documents for each employee
    print("\n" + "-" * 70)
    for idx, emp_key in enumerate(sorted(all_employee_keys), 1):
        # Resolve display name and metadata
        self_data = self_reviews_by_emp.get(emp_key)
        peer_submissions = peer_reviews_by_emp.get(emp_key, [])

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
            # Name derived from peer review
            display_name = emp_key.title()
            metadata = {
                "name": display_name,
                "email": "N/A",
                "department": "N/A",
                "date": "N/A"
            }
            self_qa = {}

        peer_grouped = group_peer_responses_question_wise(peer_submissions)

        print(f"[{idx}/{len(all_employee_keys)}] Compiling Review for: {display_name}")
        print(f"    - Self Review Questions: {len(self_qa)}")
        print(f"    - Peer Reviewers Count:  {len(peer_submissions)}")

        # Create Word Document
        clean_file_name = re.sub(r'[^\w\-_\. ]', '_', display_name)
        doc_filename = f"MidYear_Review_{clean_file_name}.docx"
        local_doc_path = os.path.join(args.output_dir, doc_filename)

        create_word_doc(metadata, self_qa, peer_grouped, local_doc_path)

        # Upload directly to the target Google Drive folder
        if args.folder_id and not args.no_upload:
            upload_docx_to_drive(drive_service, local_doc_path, args.folder_id, doc_filename)

    print("\n" + "=" * 70)
    print(" Compilation finished successfully!")
    print(f" Output Word documents saved locally in: {os.path.abspath(args.output_dir)}")
    if args.folder_id and not args.no_upload:
        print(f" Output Word documents uploaded to Drive Folder: {args.folder_id}")
    print("=" * 70)


if __name__ == "__main__":
    main()
