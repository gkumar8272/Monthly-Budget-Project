import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def generate_rtm_excel():
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # Sheet 1: RTM Traceability Matrix
    # -------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "RTM_Traceability_Matrix"
    ws1.views.sheetView[0].showGridLines = True
    
    # Styles
    title_font = Font(name="Calibri", size=16, bold=True, color="1E3A8A")
    subtitle_font = Font(name="Calibri", size=10, italic=True, color="475569")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    placeholder_header_font = Font(name="Calibri", size=10, bold=True, italic=True, color="E2E8F0")
    data_font = Font(name="Calibri", size=10, color="1E293B")
    data_bold = Font(name="Calibri", size=10, bold=True, color="1E293B")
    status_font = Font(name="Calibri", size=10, bold=True, color="166534") # Dark green
    
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    placeholder_fill = PatternFill(start_color="475569", end_color="475569", fill_type="solid") # Slate gray
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    status_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid") # Light green
    
    must_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") # Light red
    must_font = Font(name="Calibri", size=10, bold=True, color="991B1B")
    should_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid") # Light blue
    should_font = Font(name="Calibri", size=10, bold=True, color="1E40AF")
    could_fill = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid") # Light gray
    could_font = Font(name="Calibri", size=10, bold=True, color="4B5563")

    thin_border_side = Side(border_style="thin", color="CBD5E1")
    border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    # Title Block
    ws1["A1"] = "CUSTOMER ACQUISITION COST (CAC) ANALYSIS — REQUIREMENTS TRACEABILITY MATRIX (RTM)"
    ws1["A1"].font = title_font
    ws1["A2"] = "Baseline: Step 4 | Gourav Kumar (2500320140055), MCA 2nd Yr DA, ABES | Mentor: Prof. Alpna Lodhi"
    ws1["A2"].font = subtitle_font
    ws1.row_dimensions[1].height = 25
    ws1.row_dimensions[2].height = 18
    
    headers = [
        ("OQ ID", header_fill, header_font),
        ("HMW ID", header_fill, header_font),
        ("Req ID (BR)", header_fill, header_font),
        ("Requirement Statement (What the Business Needs)", header_fill, header_font),
        ("Stakeholder (SH)", header_fill, header_font),
        ("Priority", header_fill, header_font),
        ("Status", header_fill, header_font),
        ("Feature (Step 6)", placeholder_fill, placeholder_header_font),
        ("DB Table.Column (Step 8)", placeholder_fill, placeholder_header_font),
        ("User Story (Step 11)", placeholder_fill, placeholder_header_font),
        ("UI Screen (Step 12)", placeholder_fill, placeholder_header_font),
        ("Test Case TC-F/TC-T (Step 13)", placeholder_fill, placeholder_header_font)
    ]
    
    row_idx = 4
    ws1.row_dimensions[row_idx].height = 28
    for col_idx, (h_name, h_fill, h_font) in enumerate(headers, 1):
        cell = ws1.cell(row=row_idx, column=col_idx, value=h_name)
        cell.fill = h_fill
        cell.font = h_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    rtm_data = [
        ("OQ-01", "HMW-01", "BR-01", "The business needs to capture and attribute all marketing campaign expenditures by distinct acquisition channel so that non-performing channels can be isolated from blended aggregates.", "SH-01", "Must", "Identified"),
        ("OQ-01", "HMW-01", "BR-02", "The business needs to calculate and compare channel-specific CAC against blended CAC on monthly and quarterly timeframes so that underperforming paid campaigns are identified immediately.", "SH-01", "Must", "Identified"),
        ("OQ-02", "HMW-01", "BR-01", "The business needs to capture and attribute all marketing campaign expenditures by distinct acquisition channel so that non-performing channels can be isolated from blended aggregates.", "SH-01", "Must", "Identified"),
        ("OQ-02", "HMW-01", "BR-02", "The business needs to calculate and compare channel-specific CAC against blended CAC on monthly and quarterly timeframes so that underperforming paid campaigns are identified immediately.", "SH-01", "Must", "Identified"),
        ("OQ-07", "HMW-02", "BR-03", "The business needs to monitor prospective customer progression through multi-stage conversion funnels so that top-of-funnel drop-offs and capital leakage are detected.", "SH-01", "Must", "Identified"),
        ("OQ-08", "HMW-03", "BR-08", "The business needs to simulate the impact of marketing budget scaling and channel reallocation on marginal ROAS so that capital is not over-allocated past the point of diminishing returns.", "SH-02", "Should", "Identified"),
        ("OQ-04", "HMW-04", "BR-04", "The business needs to model channel-specific capital payback periods against an enterprise benchmark of 12 months so that finance can guarantee working capital solvency.", "SH-03", "Must", "Identified"),
        ("OQ-03", "HMW-05", "BR-05", "The business needs to evaluate Customer Lifetime Value against acquisition costs (LTV:CAC ratio) across distinct customer subscription tiers so that acquisition spend is restricted on churn-prone segments.", "SH-02, SH-03", "Must", "Identified"),
        ("OQ-03", "HMW-05", "BR-12", "The business needs to dynamically adjust customer churn rate assumptions in LTV calculations so that scenario planning can account for macroeconomic changes.", "SH-03", "Could", "Identified"),
        ("OQ-09", "HMW-05", "BR-05", "The business needs to evaluate Customer Lifetime Value against acquisition costs (LTV:CAC ratio) across distinct customer subscription tiers so that acquisition spend is restricted on churn-prone segments.", "SH-02, SH-03", "Must", "Identified"),
        ("OQ-06", "HMW-06", "BR-09", "The business needs to track historical monthly and seasonal acquisition cost trends so that upcoming quarterly campaign budgets can be insulated against high-demand auction price surges.", "SH-01", "Should", "Identified"),
        ("OQ-05", "HMW-07", "BR-07", "The business needs to conduct hypothesis testing on channel cost distributions so that performance differences are proven statistically significant at alpha = 0.05.", "SH-04, SH-05", "Must", "Identified"),
        ("OQ-02", "HMW-08", "BR-10", "The business needs to evaluate channel volume-versus-conversion efficiency trade-offs so that marketing teams can select the optimal channel mix between low-cost leads and high-converting customers.", "SH-02", "Should", "Identified"),
        ("OQ-09", "HMW-09", "BR-06", "The business needs to evaluate customer segment profitability and cash payback velocity across plan tiers so that sales and marketing resources focus on high-margin customer cohorts.", "SH-03", "Must", "Identified"),
        ("OQ-01", "HMW-10", "BR-11", "The business needs to provide an auditable calculation catalog detailing mathematical definitions, metric formulas, and lineage so that internal stakeholders and evaluators can verify data integrity.", "SH-05", "Could", "Identified"),
        ("OQ-05", "HMW-10", "BR-07", "The business needs to conduct hypothesis testing on channel cost distributions so that performance differences are proven statistically significant at alpha = 0.05.", "SH-04, SH-05", "Must", "Identified")
    ]

    for row_num, item in enumerate(rtm_data, start=5):
        ws1.row_dimensions[row_num].height = 24
        fill_to_use = zebra_fill if row_num % 2 == 1 else white_fill
        
        # OQ, HMW, BR
        for col_idx in range(1, 13):
            cell = ws1.cell(row=row_num, column=col_idx)
            cell.border = border
            cell.fill = fill_to_use
            
            if col_idx == 1:
                cell.value = item[0]
                cell.font = data_bold
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 2:
                cell.value = item[1]
                cell.font = data_bold
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 3:
                cell.value = item[2]
                cell.font = data_bold
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 4:
                cell.value = item[3]
                cell.font = data_font
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            elif col_idx == 5:
                cell.value = item[4]
                cell.font = data_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 6:
                cell.value = item[5]
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if item[5] == "Must":
                    cell.fill = must_fill
                    cell.font = must_font
                elif item[5] == "Should":
                    cell.fill = should_fill
                    cell.font = should_font
                else:
                    cell.fill = could_fill
                    cell.font = could_font
            elif col_idx == 7:
                cell.value = item[6]
                cell.fill = status_fill
                cell.font = status_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = "" # Empty for later steps (Steps 6, 8, 11, 12, 13)
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # Add Auto-Filter across table headers
    last_row = 4 + len(rtm_data)
    ws1.auto_filter.ref = f"A4:L{last_row}"

    # Set Column Widths
    col_widths = {
        "A": 12,  # OQ ID
        "B": 12,  # HMW ID
        "C": 14,  # Req ID (BR)
        "D": 50,  # Requirement
        "E": 18,  # Stakeholder
        "F": 12,  # Priority
        "G": 14,  # Status
        "H": 18,  # Feature (Step 6)
        "I": 22,  # DB Table.Col (Step 8)
        "J": 18,  # User Story (Step 11)
        "U": 18,  # UI Screen (Step 12)
        "L": 22   # Test Case (Step 13)
    }
    for col_letter, width in col_widths.items():
        ws1.column_dimensions[col_letter].width = width

    # -------------------------------------------------------------
    # Sheet 2: Coverage_Analysis
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Coverage_Analysis")
    ws2.views.sheetView[0].showGridLines = True
    
    ws2["A1"] = "TRACEABILITY & COVERAGE AUDIT REPORT"
    ws2["A1"].font = title_font
    ws2["A2"] = "Verification of Bidirectional Linkage: OQ <-> HMW <-> BR"
    ws2["A2"].font = subtitle_font
    ws2.row_dimensions[1].height = 24
    ws2.row_dimensions[2].height = 18

    # Summary KPI Box
    summary_headers = ["Metric / Dimension", "Expected Count", "Actual Linked Count", "Unlinked / Orphaned", "Coverage Status"]
    ws2.row_dimensions[4].height = 24
    for j, sh in enumerate(summary_headers, 1):
        c = ws2.cell(row=4, column=j, value=sh)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border

    sum_data = [
        ("Operational Questions (OQs)", 9, 9, 0, "100.0% COVERED"),
        ("How Might We (HMW) Statements", 10, 10, 0, "100.0% COVERED"),
        ("Business Requirements (BRs)", 12, 12, 0, "100.0% COVERED"),
        ("Bidirectional Traceability Links", 16, 16, 0, "100.0% BASELINED")
    ]
    for r_idx, srow in enumerate(sum_data, 5):
        ws2.row_dimensions[r_idx].height = 20
        for c_idx, val in enumerate(srow, 1):
            c = ws2.cell(row=r_idx, column=c_idx, value=val)
            c.border = border
            c.font = data_font
            if c_idx == 1:
                c.font = data_bold
                c.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx == 5:
                c.fill = status_fill
                c.font = status_font
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="center", vertical="center")

    # Detailed Audit Tables
    ws2["A11"] = "1. Operational Questions (OQ) Traceability Breakdown"
    ws2["A11"].font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")
    
    oq_audit_headers = ["OQ ID", "Operational Question Summary", "Linked HMWs", "Linked BRs", "Audit Status"]
    ws2.row_dimensions[12].height = 22
    for j, h in enumerate(oq_audit_headers, 1):
        c = ws2.cell(row=12, column=j, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border

    oq_audit_rows = [
        ("OQ-01", "Overall blended CAC trends across channels", "HMW-01, HMW-10", "BR-01, BR-02, BR-11", "PASSED"),
        ("OQ-02", "Channel CAC vs. conversion rate trade-off", "HMW-01, HMW-08", "BR-01, BR-02, BR-10", "PASSED"),
        ("OQ-03", "LTV:CAC ratio across customer tiers", "HMW-05", "BR-05, BR-12", "PASSED"),
        ("OQ-04", "Capital payback duration in months", "HMW-04", "BR-04", "PASSED"),
        ("OQ-05", "Statistical significance of channel cost variance", "HMW-07, HMW-10", "BR-07", "PASSED"),
        ("OQ-06", "Seasonal spikes and holiday quarter CAC inflation", "HMW-06", "BR-09", "PASSED"),
        ("OQ-07", "Top-of-funnel drop-off and conversion leakage", "HMW-02", "BR-03", "PASSED"),
        ("OQ-08", "Marginal ROAS budget scaling curves", "HMW-03", "BR-08", "PASSED"),
        ("OQ-09", "Segment-level profitability and payback velocity", "HMW-05, HMW-09", "BR-05, BR-06", "PASSED")
    ]

    for r_idx, oq_row in enumerate(oq_audit_rows, 13):
        ws2.row_dimensions[r_idx].height = 19
        fill = zebra_fill if r_idx % 2 == 1 else white_fill
        for c_idx, val in enumerate(oq_row, 1):
            c = ws2.cell(row=r_idx, column=c_idx, value=val)
            c.border = border
            c.fill = fill
            c.font = data_font
            if c_idx == 1:
                c.font = data_bold
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx == 5:
                c.fill = status_fill
                c.font = status_font
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    # Column Widths for ws2
    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 46
    ws2.column_dimensions["C"].width = 22
    ws2.column_dimensions["D"].width = 24
    ws2.column_dimensions["E"].width = 18

    # Save Excel file
    target_path = os.path.abspath("docs/04_Traceability_Matrix.xlsx")
    wb.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    generate_rtm_excel()
