"""
build_consolidated_submission.py
Builds the final submission deliverables:
- Removes ZIP file (LMS does not accept ZIP)
- Converts Step 04 (RTM) and Step 08 (Database Design) Excel files to PDF
- Converts all Word step documents and Presentation PPTX to PDF
- Merges ALL Steps 01 to 18 into ONE single consolidated master PDF (with professional Table of Contents cover)
- Creates an ALL-IN-ONE Master Submission PDF (Items 1 to 5 merged into 1 PDF)
- Leaves final_submission pristine with only the exact required submission files (<= 20 files, exactly 8 clean files).
Project ID: P083 | Gourav Kumar (2500320140055)
Batch: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01
Mentor: Mr. Dinesh Sereno
"""
import os
import shutil
import win32com.client
from PyPDF2 import PdfMerger
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
FINAL_DIR = os.path.join(BASE_DIR, "final_submission")
DOCS_DIR = os.path.join(BASE_DIR, "docs")
TEMP_DIR = os.path.join(BASE_DIR, "temp_submission_pdfs")

os.makedirs(FINAL_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=70, bottom=70, left=90, right=90):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def style_table(table, col_widths, headers, data):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_row = table.rows[0]
    for i, h_text in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.text = h_text
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 80, 80, 100, 100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.name = "Calibri"
            r.font.size = Pt(9.5)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for r_idx, row_data in enumerate(data):
        row = table.rows[r_idx + 1]
        bg = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = str(val)
            set_cell_background(cell, bg)
            set_cell_margins(cell, 70, 70, 90, 90)
            p = cell.paragraphs[0]
            for r in p.runs:
                r.font.name = "Calibri"
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = Inches(width)

def generate_cover_page():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(12)
    title_p.paragraph_format.space_after = Pt(4)
    r1 = title_p.add_run("MCA MAJOR CAPSTONE PROJECT (PROJECT ID: P083)")
    r1.font.name = "Calibri"
    r1.font.size = Pt(18)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(2)
    sub_p.paragraph_format.space_after = Pt(12)
    r2 = sub_p.add_run("CONSOLIDATED PROJECT LIFECYCLE DELIVERABLES (STEPS 01 TO 18)")
    r2.font.name = "Calibri"
    r2.font.size = Pt(13)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(0x02, 0x84, 0xC7)

    desc_p = doc.add_paragraph(
        "This master document compiles the complete lifecycle documentation created from Step 1 through Step 18 "
        "for the MCA Major Capstone Project. It represents the end-to-end engineering rigor applied to the "
        "Customer Acquisition Cost (CAC) Optimization & Revenue Forecasting Financial Analytics System."
    )
    desc_p.paragraph_format.space_after = Pt(12)

    # Metadata table
    meta_data = [
        ["Project ID", "P083"],
        ["Candidate Name", "Gourav Kumar"],
        ["Roll Number", "2500320140055"],
        ["Academic Batch", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"],
        ["Faculty Mentor", "Mr. Dinesh Sereno (ABES Engineering College)"],
        ["Live Streamlit App", "https://monthly-budget-project-hgcyxbnvwnvyx8g4ngpnax.streamlit.app/"],
        ["GitHub Repository", "https://github.com/gkumar8272/Monthly-Budget-Project"]
    ]
    meta_table = doc.add_table(rows=len(meta_data) + 1, cols=2)
    style_table(meta_table, [2.2, 4.6], ["Attribute", "Candidate & Project Details"], meta_data)

    p_idx = doc.add_paragraph()
    p_idx.paragraph_format.space_before = Pt(16)
    p_idx.paragraph_format.space_after = Pt(6)
    r_idx = p_idx.add_run("Consolidated Steps Table of Contents")
    r_idx.font.name = "Calibri"
    r_idx.font.size = Pt(12)
    r_idx.font.bold = True
    r_idx.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    toc_data = [
        ["Step 01", "CAC Project Understanding", "Business problem definition, cost breakdown & objectives"],
        ["Step 02", "HMW Design Thinking Statements", "10 prioritized How-Might-We innovation frameworks"],
        ["Step 03", "Needs Analysis", "Persona definitions, workflow friction points & stakeholder matrix"],
        ["Step 04", "Requirements Traceability Matrix (RTM)", "End-to-end bidirectional mapping (BRD -> TRD -> UI -> Test)"],
        ["Step 05", "Business Requirements Document (BRD)", "8 Core modules, CAC formulas, validation rules & constraints"],
        ["Step 07", "Technical Requirements Document (TRD)", "Tech stack, SQLite in-memory caching, indexing & APIs"],
        ["Step 08", "Database Design Specification", "3 relational tables, 1 analytical view, schema definitions"],
        ["Step 09", "ER Architecture & Schema Diagram", "Entity relationships, keys, normalization & ER diagram"],
        ["Step 10", "Dataset Validation Report", "Zero-loss audit, distribution checks & clean data certification"],
        ["Step 11", "Agile User Stories & Backlog", "Sprint backlog, user stories with Gherkin acceptance criteria"],
        ["Step 12", "UI/UX Framework & Wireframes", "8-screen visual layout hierarchy & responsive design rules"],
        ["Step 13", "Test Cases Specification", "14 comprehensive test scenarios (DB, Formulas, Latency)"],
        ["Step 14", "Sprint Architecture & Build Plan", "Phase-by-phase implementation plan & milestone timeline"],
        ["Step 15", "Build Verification Report", "14/14 automated Pytest execution results (100% pass)"],
        ["Step 16", "Production Deployment Runbook", "Streamlit Community Cloud hosting & operations runbook"],
        ["Step 17", "Comprehensive Project Document", "Complete academic final project report (17 sections)"],
        ["Step 18", "Project Presentation Deck", "12-slide executive presentation deck"],
        ["BI Guide", "Tableau & Power BI Dashboard Guide", "Field mappings, DAX/LOD measures & BI dashboard guide"]
    ]
    toc_table = doc.add_table(rows=len(toc_data) + 1, cols=3)
    style_table(toc_table, [1.0, 2.8, 3.0], ["Step No.", "Deliverable Title", "Lifecycle Scope & Description"], toc_data)

    cover_docx = os.path.join(TEMP_DIR, "00_Consolidated_Steps_Cover.docx")
    doc.save(cover_docx)
    return cover_docx

def main():
    print("=" * 70)
    print("BUILDING FINAL CONSOLIDATED SUBMISSION (NO ZIP, ALL STEPS IN ONE PDF)")
    print("=" * 70)

    # 1. Clean up any existing ZIP files
    zip_target = os.path.join(FINAL_DIR, "06_P083_2500320140055_Complete_Project.zip")
    if os.path.exists(zip_target):
        os.remove(zip_target)
        print("  [REMOVED] 06_P083_2500320140055_Complete_Project.zip (ZIP not accepted by LMS)")

    # 2. Clean up redundant loose step PDFs from final_submission
    redundant_loose_files = [
        "09_Dataset_Validation_Report.pdf",
        "10_Test_Cases_Specification.pdf",
        "11_Build_Verification_Report.pdf",
        "12_Deployment_Runbook.pdf",
        "13_Tableau_and_PowerBI_Dashboard_Guide.pdf",
        "07_Requirements_Traceability_Matrix_RTM.xlsx",
        "08_Database_Design_Specification.xlsx",
        "14_Dataset_cac_eda_dataset.csv"
    ]
    for rf in redundant_loose_files:
        p = os.path.join(FINAL_DIR, rf)
        if os.path.exists(p):
            os.remove(p)
            print(f"  [REMOVED LOOSE FILE] {rf}")

    # Initialize COM Objects
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = 0

    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False

    ppt_app = win32com.client.Dispatch("PowerPoint.Application")

    def docx_to_pdf(docx_path, pdf_path):
        doc = word_app.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        doc.Close()
        print(f"  [DOCX -> PDF] {os.path.basename(pdf_path)} ({os.path.getsize(pdf_path):,} bytes)")

    def xlsx_to_pdf(xlsx_path, pdf_path):
        wb = excel_app.Workbooks.Open(os.path.abspath(xlsx_path), ReadOnly=True)
        wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
        wb.Close(False)
        print(f"  [XLSX -> PDF] {os.path.basename(pdf_path)} ({os.path.getsize(pdf_path):,} bytes)")

    def pptx_to_pdf(pptx_path, pdf_path):
        deck = ppt_app.Presentations.Open(os.path.abspath(pptx_path), ReadOnly=True, Untitled=False, WithWindow=False)
        deck.SaveAs(os.path.abspath(pdf_path), 32)
        deck.Close()
        print(f"  [PPTX -> PDF] {os.path.basename(pdf_path)} ({os.path.getsize(pdf_path):,} bytes)")

    try:
        # Generate Item 1: Project Document
        print("\n[Item 1] Project Document...")
        proj_doc = os.path.join(DOCS_DIR, "2500320140055_Project_Document.docx")
        p1_pdf = os.path.join(FINAL_DIR, "01_Project_Document_P083.pdf")
        p1_docx = os.path.join(FINAL_DIR, "01_Project_Document_P083.docx")
        docx_to_pdf(proj_doc, p1_pdf)
        shutil.copy2(proj_doc, p1_docx)

        # Generate Item 2: Presentation
        print("\n[Item 2] Presentation...")
        ppt_src = os.path.join(DOCS_DIR, "18_Presentation.pptx")
        if not os.path.exists(ppt_src):
            ppt_src = os.path.join(DOCS_DIR, "2500320140055_Presentation.pptx")
        p2_pdf = os.path.join(FINAL_DIR, "02_Project_Presentation_P083.pdf")
        p2_pptx = os.path.join(FINAL_DIR, "02_Project_Presentation_P083.pptx")
        pptx_to_pdf(ppt_src, p2_pdf)
        shutil.copy2(ppt_src, p2_pptx)

        # Generate Item 3: GitHub Report
        print("\n[Item 3] GitHub Report...")
        gh_docx = os.path.join(DOCS_DIR, "03_GitHub_Report_P083.docx")
        if not os.path.exists(gh_docx):
            import generate_github_report
            generate_github_report.generate_report()
            if os.path.exists(os.path.join(FINAL_DIR, "03_GitHub_Report_P083.docx")):
                shutil.move(os.path.join(FINAL_DIR, "03_GitHub_Report_P083.docx"), gh_docx)
        p3_pdf = os.path.join(FINAL_DIR, "03_GitHub_Report_P083.pdf")
        docx_to_pdf(gh_docx, p3_pdf)

        # Generate Item 4: Complete Code
        print("\n[Item 4] Complete Code Compendium...")
        code_docx = os.path.join(DOCS_DIR, "04_Complete_Code_P083.docx")
        if not os.path.exists(code_docx):
            import generate_complete_code_doc
            generate_complete_code_doc.build_code_document()
            if os.path.exists(os.path.join(FINAL_DIR, "04_Complete_Code_P083.docx")):
                shutil.move(os.path.join(FINAL_DIR, "04_Complete_Code_P083.docx"), code_docx)
        p4_pdf = os.path.join(FINAL_DIR, "04_Complete_Code_P083.pdf")
        docx_to_pdf(code_docx, p4_pdf)

        # Generate Item 5: Consolidated All-Steps Master PDF
        print("\n[Item 5] Generating Consolidated Steps PDF with Table of Contents...")
        cover_docx = generate_cover_page()
        cover_pdf = os.path.join(TEMP_DIR, "00_Consolidated_Steps_Cover.pdf")
        docx_to_pdf(cover_docx, cover_pdf)

        # Convert RTM Excel
        rtm_xlsx = os.path.join(DOCS_DIR, "04_Traceability_Matrix.xlsx")
        rtm_pdf = os.path.join(TEMP_DIR, "04_Traceability_Matrix.pdf")
        xlsx_to_pdf(rtm_xlsx, rtm_pdf)

        # Convert Database Design Excel
        db_xlsx = os.path.join(DOCS_DIR, "08_Database_Design.xlsx")
        db_pdf = os.path.join(TEMP_DIR, "08_Database_Design.pdf")
        xlsx_to_pdf(db_xlsx, db_pdf)

        # Convert all step Word docs
        step_word_list = [
            ("Step 01", os.path.join(DOCS_DIR, "01_Customer_Acquisition_Cost_Project_Understanding.docx")),
            ("Step 02", os.path.join(DOCS_DIR, "02_HMW_Statements_Updated.docx")),
            ("Step 03", os.path.join(DOCS_DIR, "03_Needs_Analysis.docx")),
            ("Step 05", os.path.join(DOCS_DIR, "05_BRD.docx")),
            ("Step 07", os.path.join(DOCS_DIR, "07_TRD.docx")),
            ("Step 09", os.path.join(DOCS_DIR, "09_ER_Diagram.docx")),
            ("Step 10", os.path.join(DOCS_DIR, "10_Validation_Report.docx")),
            ("Step 11", os.path.join(DOCS_DIR, "11_User_Stories.docx")),
            ("Step 12", os.path.join(DOCS_DIR, "12_UI_Framework.docx")),
            ("Step 13", os.path.join(DOCS_DIR, "13_Test_Cases.docx")),
            ("Step 14", os.path.join(DOCS_DIR, "14_BUILD_PLAN.docx")),
            ("Step 15", os.path.join(DOCS_DIR, "15_Build_Verification_Report.docx")),
            ("Step 16", os.path.join(DOCS_DIR, "16_Deployment_Runbook.docx")),
            ("Step 17", os.path.join(DOCS_DIR, "2500320140055_Project_Document.docx")),
            ("BI Guide", os.path.join(DOCS_DIR, "Tableau_and_PowerBI_Dashboard_Guide.docx"))
        ]

        step_pdfs = {}
        for step_label, doc_path in step_word_list:
            if not os.path.exists(doc_path):
                alt = doc_path.replace("_Updated.docx", ".docx")
                if os.path.exists(alt):
                    doc_path = alt
            tmp_pdf = os.path.join(TEMP_DIR, f"{os.path.basename(doc_path)}.pdf")
            docx_to_pdf(doc_path, tmp_pdf)
            step_pdfs[step_label] = tmp_pdf

        # Merge all steps in numerical sequence
        print("\nMerging all steps in chronological sequence into master PDF...")
        merger = PdfMerger()
        merger.append(cover_pdf)
        merger.append(step_pdfs["Step 01"])
        merger.append(step_pdfs["Step 02"])
        merger.append(step_pdfs["Step 03"])
        merger.append(rtm_pdf)                 # Step 04
        merger.append(step_pdfs["Step 05"])
        merger.append(step_pdfs["Step 07"])
        merger.append(db_pdf)                  # Step 08
        merger.append(step_pdfs["Step 09"])
        merger.append(step_pdfs["Step 10"])
        merger.append(step_pdfs["Step 11"])
        merger.append(step_pdfs["Step 12"])
        merger.append(step_pdfs["Step 13"])
        merger.append(step_pdfs["Step 14"])
        merger.append(step_pdfs["Step 15"])
        merger.append(step_pdfs["Step 16"])
        merger.append(step_pdfs["Step 17"])
        merger.append(p2_pdf)                  # Step 18 Presentation
        merger.append(step_pdfs["BI Guide"])

        p5_pdf = os.path.join(FINAL_DIR, "05_All_Project_Steps_1_to_18_Consolidated.pdf")
        merger.write(p5_pdf)
        merger.close()
        print(f"  [CONSOLIDATED STEPS PDF OK] 05_All_Project_Steps_1_to_18_Consolidated.pdf ({os.path.getsize(p5_pdf):,} bytes)")

        # Generate Ultimate Master PDF: 00_P083_ALL_IN_ONE_MASTER_SUBMISSION.pdf
        print("\nGenerating Ultimate All-In-One Master Submission PDF (Items 1 to 5 Combined)...")
        master_merger = PdfMerger()
        master_merger.append(p1_pdf)
        master_merger.append(p2_pdf)
        master_merger.append(p3_pdf)
        master_merger.append(p4_pdf)
        master_merger.append(p5_pdf)
        master_pdf = os.path.join(FINAL_DIR, "00_P083_ALL_IN_ONE_MASTER_SUBMISSION.pdf")
        master_merger.write(master_pdf)
        master_merger.close()
        print(f"  [ALL-IN-ONE MASTER PDF OK] 00_P083_ALL_IN_ONE_MASTER_SUBMISSION.pdf ({os.path.getsize(master_pdf):,} bytes)")

    finally:
        try: word_app.Quit()
        except: pass
        try: excel_app.Quit()
        except: pass
        try: ppt_app.Quit()
        except: pass

    # Clean up temp working folder
    shutil.rmtree(TEMP_DIR, ignore_errors=True)

    # Inventory Verification
    print("\n" + "=" * 70)
    print("FINAL SUBMISSION FOLDER INVENTORY CHECK:")
    print("=" * 70)
    final_files = sorted(os.listdir(FINAL_DIR))
    for i, f in enumerate(final_files, 1):
        fp = os.path.join(FINAL_DIR, f)
        size_kb = os.path.getsize(fp) / 1024
        print(f"  {i:02d}. {f:<55} ({size_kb:8.1f} KB)")

    print("=" * 70)
    print(f"TOTAL FILES IN final_submission: {len(final_files)} (Constraint: <= 20 files, NO ZIP)")
    print("=" * 70)

if __name__ == "__main__":
    main()
