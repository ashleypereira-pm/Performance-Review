# 📝 Performance Review Compiler

A tool to compile employee **Self-Reviews** and **Peer Feedback** into consolidated, professionally styled Microsoft Word (`.docx`) documents.

---

## 🌟 Key Features

1. **Part 1 – Self Review:**
   * Formats employee self-evaluation responses under clean question headers (`Q1.`, `Q2.`, etc.).
   * Automatically strips Google Forms question prefixes (`4.`, `5.`, `6.`) and instruction text (*"Please list 3-5..."*).

2. **Part 2 – Peer Feedback Summary:**
   * Combines all original feedback from 2–3 nominated peers into a single unified summary block under each question.
   * **100% Original Tone & Intent:** Preserves verbatim responses without AI rephrasing or alteration.
   * **Anonymous:** Excludes reviewer names and reviewer labels from the document.

3. **Part 3 – Manager's Feedback & Evaluation:**
   * Includes a designated, shaded container box at the end of each document for the reporting manager to write their summary review paragraph, performance observations, and guidance.
   * Includes Manager Name and Date sign-off placeholders.

4. **Executive Styling:**
   * Navy blue headers (`#1B365D`), muted metadata table cards, and standard 1-inch margins.

---

## 🚀 Two Ways to Run

### Method 1: Streamlit Web UI (Zero-Credentials / Privacy-First)

Ideal for HR teams. No Google credentials, API keys, or tokens required on the server.

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Launch the web application:**
   ```bash
   streamlit run app.py
   ```

3. **How HR Uses It:**
   * Open the Google Form response sheets in Google Sheets $\rightarrow$ `File` $\rightarrow$ `Download` $\rightarrow$ `Microsoft Excel (.xlsx)` or `.csv`.
   * Drag and drop the **Self-Review** and **Peer-Review** files into the Streamlit interface.
   * Click **"🚀 Compile All Reviews into Word Documents"**.
   * Download the complete `Compiled_MidYear_Reviews.zip` archive or download individual candidate documents.

---

### Method 2: Command-Line Interface (CLI with Direct Google Drive Upload)

Connects directly to Google Drive & Google Sheets API using `token.json` and `credentials.json` to process and optionally upload documents directly to a Google Drive folder.

#### 1. Compile Locally (No Upload):
```bash
python compile_reviews.py --no-upload
```
*(Generated documents are saved in `./compiled_reviews/`)*

#### 2. Compile & Upload Directly to a Google Drive Folder:
```bash
python compile_reviews.py --folder-id <YOUR_GOOGLE_DRIVE_FOLDER_ID>
```

#### Optional CLI Arguments:
* `--self-sheet-id <ID>` : Override the Self-Review Spreadsheet ID.
* `--peer-sheet-id <ID>` : Override the Peer-Review Spreadsheet ID.
* `--output-dir <PATH>` : Change local output directory (default: `./compiled_reviews`).
* `--credentials <PATH>` : Path to Google `credentials.json` (default: `credentials.json`).
* `--token <PATH>` : Path to Google `token.json` (default: `token.json`).

---

## 📁 Folder Structure

```
performance_review_compiler/
├── app.py                  # Streamlit Web Application (Drag & drop, zero credentials)
├── compile_reviews.py      # Python CLI script (Google Drive API integration)
├── requirements.txt        # Python package dependencies
└── README.md               # User & setup documentation
```

---

## 📋 Document Format Sample

```
MID-YEAR PERFORMANCE REVIEW
Consolidated Self-Evaluation & Peer Feedback

┌─────────────────────────────────────────────────────────────┐
│ Employee Name: Aksha         Department: Sales & BD         │
│ Email Address: aksha@...     Review Date: 25/09/2026        │
└─────────────────────────────────────────────────────────────┘

Part 1 – Self Review

Q1. What are your key achievements during the last 6 months?
    1. Achieved and exceeded monthly sales targets consistently.
    2. Generated new leads and converted them into customers.
    3. Improved client relationships and increased repeat business.

Q2. What did you learn, and where do you think you can improve?
    1. Learned how to handle client conversations more effectively...

---------------------------------------------------------------

Part 2 – Peer Feedback Summary
Combined summary of feedback from nominated peers.

Question 1: What does this person do particularly well?
  • She is good at maintaining positive relationships with clients and handles conversations professionally.
  • She is approachable and easy to work with, especially when coordination is needed between team members.
  • She takes feedback positively and adapts well to changes and new requirements.
  • Communicates her thoughts clearly and is able to justify her statements appropriately.
  • Good at networking and managing people.

Question 2: What could this person improve?
  • She could improve her prioritization when handling multiple leads or tasks at the same time.
  • She could be more proactive in sharing updates on ongoing client discussions.

---------------------------------------------------------------

Part 3 – Manager's Feedback & Evaluation
Section for the reporting manager to provide overall summary observations, key achievements, growth areas, and guidance.

┌─────────────────────────────────────────────────────────────┐
│ Manager's Summary & Comments:                               │
│                                                             │
│ [Manager types their performance summary paragraph here]    │
│                                                             │
└─────────────────────────────────────────────────────────────┘

Manager Name: __________________________        Date: _______________
```
