"""
build_final_submission.py
Orchestrates the entire final submission process for LMS upload:
Converts DOCX/PPTX to PDF, merges all steps into a single consolidated PDF,
creates the clean ZIP archive, and copies key deliverables.
Ensures total file count is strictly under the 20-file limit.
Project ID: P083 | Candidate: Gourav Kumar (2500320140055)
"""
import os
import shutil
import zipfile
import win32com.client
from PyPDF2 import PdfMerger

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
FINAL_DIR = os.path.join(BASE_DIR, "final_submission")
DOCS_DIR = os.path.join(BASE_DIR, "docs")
TEMP_PDF_DIR = os.path.join(BASE_DIR, "temp_submission_pdfs")

os.makedirs(FINAL_DIR, exist_ok=True)
os.makedirs(TEMP_PDF_DIR, exist_ok=True)

print("Starting Final Submission Preparation...")

# Initialize Word and PowerPoint COM objects
word_app = win32com.client.Dispatch("Word.Application")
word_app.Visible = False
word_app.DisplayAlerts = 0
ppt_app = win32com.client.Dispatch("PowerPoint.Application")

def convert_docx_to_pdf(docx_path, pdf_path):
    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path)
    try:
        doc = word_app.Documents.Open(abs_docx, ReadOnly=True)
        doc.SaveAs(abs_pdf, FileFormat=17) # 17 = wdFormatPDF
        doc.Close()
        print(f"  [PDF OK] {os.path.basename(abs_pdf)} ({os.path.getsize(abs_pdf):,} bytes)")
        return True
    except Exception as e:
        print(f"  [PDF ERROR] Could not convert {abs_docx}: {e}")
        return False

def convert_pptx_to_pdf(pptx_path, pdf_path):
    abs_pptx = os.path.abspath(pptx_path)
    abs_pdf = os.path.abspath(pdf_path)
    try:
        deck = ppt_app.Presentations.Open(abs_pptx, ReadOnly=True, Untitled=False, WithWindow=False)
        deck.SaveAs(abs_pdf, 32) # 32 = ppSaveAsPDF
        deck.Close()
        print(f"  [PDF OK] {os.path.basename(abs_pdf)} ({os.path.getsize(abs_pdf):,} bytes)")
        return True
    except Exception as e:
        print(f"  [PDF ERROR] Could not convert {abs_pptx}: {e}")
        return False

try:
    # ── ITEM 1: Project Document (PDF & DOCX) ──────────────────────────────
    print("\n1. Generating Item 1: Project Document...")
    proj_doc_docx = os.path.join(DOCS_DIR, "2500320140055_Project_Document.docx")
    out_doc_pdf = os.path.join(FINAL_DIR, "01_Project_Document_P083.pdf")
    out_doc_docx = os.path.join(FINAL_DIR, "01_Project_Document_P083.docx")
    convert_docx_to_pdf(proj_doc_docx, out_doc_pdf)
    shutil.copy2(proj_doc_docx, out_doc_docx)

    # ── ITEM 2: Project Presentation (PDF & PPTX) ─────────────────────────
    print("\n2. Generating Item 2: Project Presentation...")
    ppt_source = os.path.join(DOCS_DIR, "18_Presentation.pptx")
    if not os.path.exists(ppt_source):
        ppt_source = os.path.join(DOCS_DIR, "2500320140055_Presentation.pptx")
    out_ppt_pdf = os.path.join(FINAL_DIR, "02_Project_Presentation_P083.pdf")
    out_ppt_pptx = os.path.join(FINAL_DIR, "02_Project_Presentation_P083.pptx")
    convert_pptx_to_pdf(ppt_source, out_ppt_pdf)
    shutil.copy2(ppt_source, out_ppt_pptx)

    # ── ITEM 3: GitHub Report (PDF) ───────────────────────────────────────
    print("\n3. Generating Item 3: GitHub Report...")
    gh_docx = os.path.join(FINAL_DIR, "03_GitHub_Report_P083.docx")
    if not os.path.exists(gh_docx):
        import generate_github_report
        generate_github_report.generate_report()
    out_gh_pdf = os.path.join(FINAL_DIR, "03_GitHub_Report_P083.pdf")
    convert_docx_to_pdf(gh_docx, out_gh_pdf)

    # ── ITEM 4: Complete Code (PDF) ───────────────────────────────────────
    print("\n4. Generating Item 4: Complete Code Compendium...")
    code_docx = os.path.join(FINAL_DIR, "04_Complete_Code_P083.docx")
    if not os.path.exists(code_docx):
        import generate_complete_code_doc
        generate_complete_code_doc.build_code_document()
    out_code_pdf = os.path.join(FINAL_DIR, "04_Complete_Code_P083.pdf")
    convert_docx_to_pdf(code_docx, out_code_pdf)

    # ── ITEM 5: All Project Steps 1 to 18 Consolidated PDF ────────────────
    print("\n5. Generating Item 5: Consolidated All-Steps Master PDF...")
    step_docs = [
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
        ("Tableau & PowerBI", os.path.join(DOCS_DIR, "Tableau_and_PowerBI_Dashboard_Guide.docx"))
    ]

    temp_pdfs_to_merge = []
    for step_name, doc_path in step_docs:
        if not os.path.exists(doc_path):
            alt = doc_path.replace("_Updated.docx", ".docx")
            if os.path.exists(alt):
                doc_path = alt
        if os.path.exists(doc_path):
            tmp_pdf = os.path.join(TEMP_PDF_DIR, f"{os.path.basename(doc_path)}.pdf")
            if convert_docx_to_pdf(doc_path, tmp_pdf):
                temp_pdfs_to_merge.append(tmp_pdf)

    # Also add presentation PDF at the end of consolidated doc
    if os.path.exists(out_ppt_pdf):
        temp_pdfs_to_merge.append(out_ppt_pdf)

    # Merge into a single master PDF
    print(f"Merging {len(temp_pdfs_to_merge)} individual step documents into master consolidated PDF...")
    merger = PdfMerger()
    for p in temp_pdfs_to_merge:
        merger.append(p)
    consolidated_pdf_path = os.path.join(FINAL_DIR, "05_All_Project_Steps_1_to_18_Consolidated.pdf")
    merger.write(consolidated_pdf_path)
    merger.close()
    print(f"  [CONSOLIDATED PDF OK] 05_All_Project_Steps_1_to_18_Consolidated.pdf ({os.path.getsize(consolidated_pdf_path):,} bytes)")

    # ── Convert Key Individual Step Documents for direct LMS upload ────────
    print("\nConverting key individual reports...")
    key_individual_conversions = [
        (os.path.join(DOCS_DIR, "10_Validation_Report.docx"), os.path.join(FINAL_DIR, "09_Dataset_Validation_Report.pdf")),
        (os.path.join(DOCS_DIR, "13_Test_Cases.docx"), os.path.join(FINAL_DIR, "10_Test_Cases_Specification.pdf")),
        (os.path.join(DOCS_DIR, "15_Build_Verification_Report.docx"), os.path.join(FINAL_DIR, "11_Build_Verification_Report.pdf")),
        (os.path.join(DOCS_DIR, "16_Deployment_Runbook.docx"), os.path.join(FINAL_DIR, "12_Deployment_Runbook.pdf")),
        (os.path.join(DOCS_DIR, "Tableau_and_PowerBI_Dashboard_Guide.docx"), os.path.join(FINAL_DIR, "13_Tableau_and_PowerBI_Dashboard_Guide.pdf"))
    ]
    for src, dst in key_individual_conversions:
        if os.path.exists(src):
            convert_docx_to_pdf(src, dst)

finally:
    # Safely close Word and PPT
    try:
        word_app.Quit()
    except Exception:
        pass
    try:
        ppt_app.Quit()
    except Exception:
        pass

# ── ITEM 6: Project ZIP file ──────────────────────────────────────────────
print("\n6. Creating Item 6: Clean Project ZIP File...")
zip_path = os.path.join(FINAL_DIR, "06_P083_2500320140055_Complete_Project.zip")
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(BASE_DIR):
        # Exclude node_modules, .git, .venv, final_submission, temp_submission_pdfs
        if any(skip in root for skip in ["node_modules", ".git", ".venv", "final_submission", "temp_submission_pdfs", "__pycache__", ".pytest_cache"]):
            continue
        for file in files:
            if file.endswith((".pyc", ".tmp", ".log")):
                continue
            if file.startswith("~$"):
                continue
            full_path = os.path.join(root, file)
            rel_path = os.path.relpath(full_path, BASE_DIR)
            zf.write(full_path, rel_path)

print(f"  [ZIP OK] 06_P083_2500320140055_Complete_Project.zip ({os.path.getsize(zip_path):,} bytes)")

# ── Copy Key Supporting Artifacts ─────────────────────────────────────────
print("\nCopying key supporting spreadsheets and data files...")
shutil.copy2(os.path.join(DOCS_DIR, "04_Traceability_Matrix.xlsx"), os.path.join(FINAL_DIR, "07_Requirements_Traceability_Matrix_RTM.xlsx"))
shutil.copy2(os.path.join(DOCS_DIR, "08_Database_Design.xlsx"), os.path.join(FINAL_DIR, "08_Database_Design_Specification.xlsx"))
shutil.copy2(os.path.join(BASE_DIR, "data", "cac_eda_dataset.csv"), os.path.join(FINAL_DIR, "14_Dataset_cac_eda_dataset.csv"))

# Archive working DOCX files in docs/ and remove from final_submission so final_submission has only the required deliverables
for f in ["03_GitHub_Report_P083.docx", "04_Complete_Code_P083.docx"]:
    f_path = os.path.join(FINAL_DIR, f)
    if os.path.exists(f_path):
        shutil.copy2(f_path, os.path.join(DOCS_DIR, f))
        os.remove(f_path)

# Clean up temp PDF dir
shutil.rmtree(TEMP_PDF_DIR, ignore_errors=True)

# ── Final Inventory Check ─────────────────────────────────────────────────
print("\n" + "=" * 70)
print("FINAL SUBMISSION FOLDER INVENTORY CHECK:")
print("=" * 70)
final_files = sorted(os.listdir(FINAL_DIR))
for i, f in enumerate(final_files, 1):
    fp = os.path.join(FINAL_DIR, f)
    size_kb = os.path.getsize(fp) / 1024
    print(f"  {i:02d}. {f:<55} ({size_kb:8.1f} KB)")

print("=" * 70)
print(f"TOTAL FILES IN final_submission: {len(final_files)} (LMS Limit: <= 20 files)")
if len(final_files) <= 20:
    print("SUCCESS: Within the 20-file limit! Ready for instant LMS upload.")
else:
    print("WARNING: Exceeds 20-file limit!")
