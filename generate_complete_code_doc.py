"""
generate_complete_code_doc.py
Compiles all production source code files into a single, beautifully formatted DOCX document for LMS submission.
Project ID: P083 | Candidate: Gourav Kumar (2500320140055)
"""
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
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

def set_table_borders(table, color="CBD5E1"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def style_table(table, col_widths, headers, data):
    set_table_borders(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "1E3A8A")
        set_cell_margins(hdr_cells[i], top=70, bottom=70, left=80, right=80)
        p = hdr_cells[i].paragraphs[0]
        for run in p.runs:
            run.font.bold = True
            run.font.size = Pt(8.5)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for r_idx, row_data in enumerate(data):
        row_cells = table.rows[r_idx + 1].cells
        bg_color = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = str(val)
            set_cell_background(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=50, bottom=50, left=70, right=70)
            p = row_cells[c_idx].paragraphs[0]
            for run in p.runs:
                run.font.size = Pt(8)
                run.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = Inches(w)

def add_code_block(doc, code_text):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = tbl.rows[0].cells[0]
    c.width = Inches(6.8)
    set_cell_background(c, "F8FAFC")
    set_cell_margins(c, top=80, bottom=80, left=100, right=100)
    
    # Border
    tcPr = c._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>
            <w:left w:val="single" w:sz="18" w:space="0" w:color="0284C7"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = c.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run(code_text)
    run.font.name = "Consolas"
    run.font.size = Pt(7.5)
    run.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

def build_code_document():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(9.5)
    normal_style.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    # Title
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_inst = p_inst.add_run("ABES ENGINEERING COLLEGE, GHAZIABAD\nDEPARTMENT OF COMPUTER APPLICATIONS\n")
    r_inst.font.size = Pt(12)
    r_inst.font.bold = True
    r_inst.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("COMPLETE PROJECT SOURCE CODE COMPENDIUM")
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Customer Acquisition Cost (CAC) & Marketing Unit Economics Analysis Platform")
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_div = p_div.add_run("―" * 60)
    r_div.font.color.rgb = RGBColor(0x02, 0x84, 0xC7)

    # Cover Table
    meta_data = [
        ["Project ID", "P083"],
        ["Project Title", "Customer Acquisition Cost (CAC) & Marketing Unit Economics Analysis"],
        ["Candidate Name", "Gourav Kumar"],
        ["Roll Number", "2500320140055"],
        ["Batch ID", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"],
        ["Academic Program", "Master of Computer Applications (MCA) — 2nd Year Data Analytics"],
        ["Faculty Mentor", "Mr. Dinesh Sereno, Assistant Professor"],
        ["GitHub Repository URL", "https://github.com/gkumar8272/Monthly-Budget-Project"],
        ["Live Production Streamlit URL", "https://monthly-budget-project-hgcyxbnvwnvyx8g4ngpnax.streamlit.app/"],
        ["Pytest Pass Rate", "14/14 Passed (100% Pass Rate in 2.70 seconds)"],
        ["Date of Submission", "November 21, 2026"]
    ]
    t_meta = doc.add_table(rows=len(meta_data)+1, cols=2)
    style_table(t_meta, [2.3, 4.4], ["Attribute", "Project Specification"], meta_data)

    code_files = [
        ("1. Application Driver", "app.py", "Main multi-page entrypoint, Executive Cockpit (UI-01), KPI cards and budget run-rate alerts."),
        ("2. Page: Channel CAC Trends", "pages/1_Channel_CAC.py", "Channel-level monthly spend vs. CAC trendlines and 12-month payback benchmark line (UI-02)."),
        ("3. Page: Conversion Funnel", "pages/2_Funnel_Analysis.py", "Multi-stage conversion funnel drop-off and capital leakage diagnostics (UI-03)."),
        ("4. Page: Volume vs. Efficiency", "pages/3_Compare.py", "2x2 quadrant scatter matrix segregating channels into Stars, Cash Cows, Bleeders (UI-04)."),
        ("5. Page: Statistical Hypothesis Testing", "pages/4_Statistical_Testing.py", "Shapiro-Wilk normality testing, One-Way ANOVA F-test, and Kruskal-Wallis H-test (UI-05)."),
        ("6. Page: Budget Allocation Simulator", "pages/5_Budget_Simulator.py", "Diminishing marginal ROAS logarithmic response curve budget reallocation sandbox (UI-06)."),
        ("7. Page: Database Explorer & Export", "pages/6_Data_Explorer.py", "Paginated database browser with one-click CSV and Excel export functionality (UI-07)."),
        ("8. Page: Formula Catalog & Quality", "pages/7_Formula_Catalog.py", "Auditable KaTeX mathematical definitions and data quality audit checklist (UI-08)."),
        ("9. Engine: Database Pool Manager", "src/db.py", "Thread-safe SQLite connection pool manager with read-only concurrency support."),
        ("10. Engine: Cached Data Loader", "src/data_loader.py", "In-memory query caching via Streamlit @st.cache_data for sub-15ms query execution."),
        ("11. Engine: Data Sanitization Pipeline", "src/data_cleaner.py", "Deduplication and median null-imputation pipeline."),
        ("12. Engine: CAC & Unit Economics Engine", "src/cac_calculator.py", "Vectorized zero-division guarded CAC, LTV, and payback calculation functions."),
        ("13. Engine: Statistical Inference Engine", "src/stats_engine.py", "SciPy and Statsmodels inferential testing module."),
        ("14. Engine: Marginal ROAS Simulator", "src/simulator.py", "Logarithmic response curves modeling channel budget scaling returns."),
        ("15. Engine: Tailwind CSS UI Components", "src/ui.py", "Tailwind CSS iframe injection layer for executive cards and responsive layout."),
        ("16. Test: Database Integrity Suite", "tests/test_data.py", "Automated pytest suite for connectivity, PK uniqueness, FK referential integrity."),
        ("17. Test: Analysis & Math Suite", "tests/test_analysis.py", "Automated pytest suite for mathematical formula parity, ANOVA, and zero-division guards."),
        ("18. Test: App & SLA Suite", "tests/test_app.py", "Automated pytest suite for query latency SLA (<3000ms) and multi-page imports."),
        ("19. Cloud Environment Dependencies", "requirements.txt", "Pinned production package dependencies for Streamlit Community Cloud."),
        ("20. UI Theme Configuration", ".streamlit/config.toml", "Executive corporate slate/navy theme and server configuration.")
    ]

    p_toc = doc.add_paragraph()
    p_toc.paragraph_format.space_before = Pt(16)
    p_toc.paragraph_format.space_after = Pt(4)
    r_toc = p_toc.add_run("Code Inventory & Table of Contents")
    r_toc.font.size = Pt(14)
    r_toc.font.bold = True
    r_toc.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    toc_rows = [[sec, path, role] for sec, path, role in code_files]
    t_toc = doc.add_table(rows=len(toc_rows)+1, cols=3)
    style_table(t_toc, [2.2, 2.0, 2.6], ["Component", "File Path", "Architecture Purpose"], toc_rows)

    doc.add_page_break()

    # Loop through and append each code file
    for title, rel_path, desc in code_files:
        full_path = os.path.abspath(rel_path)
        if not os.path.exists(full_path):
            continue

        p_h = doc.add_paragraph()
        p_h.paragraph_format.space_before = Pt(14)
        p_h.paragraph_format.space_after = Pt(2)
        p_h.paragraph_format.keep_with_next = True
        r_h = p_h.add_run(f"{title}: {rel_path}")
        r_h.font.size = Pt(11.5)
        r_h.font.bold = True
        r_h.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

        p_d = doc.add_paragraph()
        p_d.paragraph_format.space_before = Pt(0)
        p_d.paragraph_format.space_after = Pt(4)
        r_d = p_d.add_run(f"Purpose: {desc} | File Path: {rel_path}")
        r_d.font.size = Pt(8.5)
        r_d.font.italic = True
        r_d.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

        try:
            with open(full_path, "r", encoding="utf-8", errors="ignore") as f:
                code_content = f.read()
            add_code_block(doc, code_content)
        except Exception as e:
            p_err = doc.add_paragraph(f"[Error loading code: {e}]")
            p_err.runs[0].font.color.rgb = RGBColor(0xDC, 0x26, 0x26)

        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    out_docx = r"c:\Users\asus\Downloads\MERN-Expense-Tracker-main\final_submission\04_Complete_Code_P083.docx"
    doc.save(out_docx)
    print(f"Complete Code document saved to {out_docx}")

if __name__ == "__main__":
    build_code_document()
