import os
import shutil
import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

def build_perfect_report():
    print("Building perfect report matching reference template...")
    doc = docx.Document()

    # Configure Normal Style
    normal_style = doc.styles['Normal']
    normal_font = normal_style.font
    normal_font.name = 'Times New Roman'
    normal_font.size = Pt(12)
    normal_font.color.rgb = RGBColor(0, 0, 0)

    # Configure Section: A4, 1-inch margins, double line border
    sec = doc.sections[0]
    sec.page_width = Inches(8.27)
    sec.page_height = Inches(11.69)
    sec.top_margin = Inches(1.0)
    sec.bottom_margin = Inches(1.0)
    sec.left_margin = Inches(1.0)
    sec.right_margin = Inches(1.0)

    # Double border matching reference: thinThickThinSmallGap sz=24 space=24
    pgBorders = parse_xml(
        r'<w:pgBorders %s w:offsetFrom="page">'
        r'<w:top w:val="thinThickThinSmallGap" w:sz="24" w:space="24" w:color="auto"/>'
        r'<w:left w:val="thinThickThinSmallGap" w:sz="24" w:space="24" w:color="auto"/>'
        r'<w:bottom w:val="thinThickThinSmallGap" w:sz="24" w:space="24" w:color="auto"/>'
        r'<w:right w:val="thinThickThinSmallGap" w:sz="24" w:space="24" w:color="auto"/>'
        r'</w:pgBorders>' % nsdecls('w')
    )
    sec._sectPr.append(pgBorders)

    # -------------------------------------------------------------
    # Formatting Helpers (Pure Black and White, Simple Text)
    # -------------------------------------------------------------
    def add_heading(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_subheading(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_p(text, bold_prefix=''):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.5
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = 'Times New Roman'
            r_pre.font.size = Pt(12)
            r_pre.font.bold = True
            r_pre.font.color.rgb = RGBColor(0, 0, 0)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_bullet(text, bold_prefix=''):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.5
        r_b = p.add_run('• ')
        r_b.font.name = 'Times New Roman'
        r_b.font.size = Pt(12)
        r_b.font.color.rgb = RGBColor(0, 0, 0)
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = 'Times New Roman'
            r_pre.font.size = Pt(12)
            r_pre.font.bold = True
            r_pre.font.color.rgb = RGBColor(0, 0, 0)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_table(caption, headers, rows_data, col_widths=None):
        if caption:
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_cap.paragraph_format.space_before = Pt(10)
            p_cap.paragraph_format.space_after = Pt(4)
            p_cap.paragraph_format.line_spacing = 1.15
            p_cap.paragraph_format.keep_with_next = True
            run_cap = p_cap.add_run(caption)
            run_cap.font.name = 'Times New Roman'
            run_cap.font.size = Pt(12)
            run_cap.font.bold = True
            run_cap.font.color.rgb = RGBColor(0, 0, 0)

        table = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Table Borders: 0.75 pt single black
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            r'<w:tblBorders %s>'
            r'<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'<w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'<w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'</w:tblBorders>' % nsdecls('w')
        )
        tblPr.append(borders)

        # Cell margins
        tblCellMar = parse_xml(
            r'<w:tblCellMar %s>'
            r'<w:top w:w="80" w:type="dxa"/>'
            r'<w:bottom w:w="80" w:type="dxa"/>'
            r'<w:left w:w="120" w:type="dxa"/>'
            r'<w:right w:w="120" w:type="dxa"/>'
            r'</w:tblCellMar>' % nsdecls('w')
        )
        tblPr.append(tblCellMar)

        # Header row
        trPr = table.rows[0]._tr.get_or_add_trPr()
        trPr.append(parse_xml(r'<w:tblHeader %s/>' % nsdecls('w')))

        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(h)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)

        # Data rows
        for r_idx, r_data in enumerate(rows_data):
            for c_idx, val in enumerate(r_data):
                cell = table.cell(r_idx + 1, c_idx)
                cell.text = ''
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15
                run = p.add_run(str(val))
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                run.font.bold = False
                run.font.color.rgb = RGBColor(0, 0, 0)

        if col_widths:
            for row in table.rows:
                for idx, w in enumerate(col_widths):
                    row.cells[idx].width = Inches(w)

        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(4)
        p_sp.paragraph_format.line_spacing = 1.0

    def add_screenshot(fig_title, img_path, width_in=5.2):
        if not os.path.exists(img_path):
            print(f"Warning: image {img_path} missing!")
            return

        p_label = doc.add_paragraph()
        p_label.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_label.paragraph_format.space_before = Pt(10)
        p_label.paragraph_format.space_after = Pt(4)
        p_label.paragraph_format.line_spacing = 1.15
        p_label.paragraph_format.keep_with_next = True
        r_lbl = p_label.add_run(fig_title)
        r_lbl.font.name = 'Times New Roman'
        r_lbl.font.size = Pt(12)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = RGBColor(0, 0, 0)

        # Image table frame (0.75 pt border around image)
        fig_tbl = doc.add_table(rows=1, cols=1)
        fig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = fig_tbl._tbl.tblPr
        borders = parse_xml(
            r'<w:tblBorders %s>'
            r'<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'</w:tblBorders>' % nsdecls('w')
        )
        tblPr.append(borders)

        tblCellMar = parse_xml(
            r'<w:tblCellMar %s>'
            r'<w:top w:w="30" w:type="dxa"/>'
            r'<w:bottom w:w="30" w:type="dxa"/>'
            r'<w:left w:w="30" w:type="dxa"/>'
            r'<w:right w:w="30" w:type="dxa"/>'
            r'</w:tblCellMar>' % nsdecls('w')
        )
        tblPr.append(tblCellMar)

        cell = fig_tbl.cell(0, 0)
        cell.width = Inches(width_in + 0.1)
        p_img = cell.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(2)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.paragraph_format.line_spacing = 1.0
        r_img = p_img.add_run()
        r_img.add_picture(img_path, width=Inches(width_in))

        # Blank space after figure
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(8)

    # =============================================================
    # COVER PAGE (Exact Match to CyberSafe_Project_Report.docx)
    # =============================================================
    p0 = doc.add_paragraph()
    p0.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r0 = p0.add_run("MINOR PROJECT REPORT")
    r0.font.name = "Times New Roman"
    r0.font.size = Pt(16)
    r0.font.bold = True

    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r1 = p1.add_run("CYBERSECURITY RISK ASSESSMENT FRAMEWORK FOR SMALL BUSINESSES")
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(16)
    r1.font.bold = True

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    meta_entries = [
        ("Project Domain:", " Cyber Security"),
        ("Project Name:", " SME Risk Framework"),
        ("Technology Used:", " Python, FastAPI, React, Tailwind CSS, MongoDB"),
        ("Application Environment:", " Localhost (127.0.0.1:8000, localhost:3000)"),
        ("Student Name:", " Mahi Patel, Jainam Shah"),
        ("Group:", " C"),
        ("Email ID:", " mahip290905@gmail.com   , sjainam1811@gmail.com "),
        ("Academic Year:", " 2026–2027"),
    ]

    for label, val in meta_entries:
        p_m = doc.add_paragraph()
        p_m.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_lbl = p_m.add_run(label)
        r_lbl.font.name = "Times New Roman"
        r_lbl.font.size = Pt(14)
        r_lbl.font.bold = True
        r_val = p_m.add_run(val)
        r_val.font.name = "Times New Roman"
        r_val.font.size = Pt(14)
        r_val.font.bold = False

    # 16 Blank lines matching sample docx (P11 to P26)
    for _ in range(16):
        p_empty = doc.add_paragraph()
        p_empty.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # =============================================================
    # 1. INTRODUCTION
    # =============================================================
    add_heading("1. INTRODUCTION")
    add_p("Cybersecurity has become an important part of modern digital life. Small and medium enterprises (SMEs) use digital systems for daily business, sales, accounting, and communication. However, these systems also expose small businesses to cyber threats.")
    add_p("One of the main challenges for small businesses is the lack of dedicated cybersecurity personnel and technical security tools.")
    add_p("Cyber attackers actively target small businesses using phishing, ransomware, weak passwords, and credential theft.")
    add_p("The main purpose of this project is to develop a practical cybersecurity risk assessment framework for small businesses.")
    add_p("A local web application named SME Risk Framework was developed to evaluate risks, map controls to the NIST Cybersecurity Framework 2.0, display interactive charts, and export single-page PDF audit reports.")
    add_p("The application runs locally and uses standard quantitative evaluation metrics.")

    # =============================================================
    # 2. PROBLEM STATEMENT
    # =============================================================
    add_heading("2. PROBLEM STATEMENT")
    add_p("Small businesses are attractive targets for cyber attackers because they often possess limited cybersecurity defenses.")
    add_p("Attackers exploit stolen credentials, unpatched software, and lack of staff training to steal funds and lock business data.")
    add_p("Existing cybersecurity frameworks such as NIST SP 800-53 and ISO 27001 contain hundreds of complex controls that are difficult and expensive for SMEs to implement.")
    add_p("This project aims to provide a lightweight, quantitative risk assessment platform that small businesses can use to evaluate their cybersecurity posture and prioritize practical defenses.")

    # =============================================================
    # 3. OBJECTIVES
    # =============================================================
    add_heading("3. OBJECTIVES")
    add_p("The main objectives are:")
    add_bullet("To research common cybersecurity risks facing small businesses (phishing, ransomware, insider threats).")
    add_bullet("To develop a risk assessment model aligned with the NIST Cybersecurity Framework 2.0.")
    add_bullet("To create a quantitative evaluation metric (Likelihood × Impact) to rank threats by severity.")
    add_bullet("To propose practical risk mitigation strategies including firewalls, multi-factor authentication, and backups.")
    add_bullet("To validate the risk assessment model using case studies on different small businesses.")
    add_bullet("To document findings and compile practical cybersecurity best practices for SMEs.")
    add_bullet("To develop a local web application featuring interactive charts and executive PDF report export.")

    # =============================================================
    # 4. RESEARCH ON COMMON SME CYBER RISKS
    # =============================================================
    add_heading("4. RESEARCH ON COMMON SME CYBER RISKS")
    add_p("Our research identified six common cyber threats that impact small businesses:")

    add_subheading("4.1 Phishing")
    add_p("Phishing involves fraudulent emails or messages designed to deceive staff into revealing passwords or transferring company funds.")
    add_p("Example: An email claiming that an urgent supplier payment must be rerouted to a new bank account.")

    add_subheading("4.2 Ransomware")
    add_p("Ransomware is malicious software that encrypts company databases and files until an extortion fee is paid.")
    add_p("Example: An infected attachment that encrypts retail point-of-sale data or medical records.")

    add_subheading("4.3 Weak Passwords")
    add_p("Weak passwords allow attackers to gain unauthorized access through automated credential stuffing and password guessing.")
    add_p("Example: Using simple or shared passwords across business systems without multi-factor authentication.")

    add_subheading("4.4 Malware")
    add_p("Malware includes viruses, keyloggers, and spyware that steal confidential business data or damage computer systems.")
    add_p("Example: Malicious software installed via compromised downloads that logs administrative keystrokes.")

    add_subheading("4.5 Data Breaches")
    add_p("Data breaches involve unauthorized access to or copying of sensitive customer, financial, or employee information.")
    add_p("Example: Unencrypted customer records stored on an office computer accessed by unauthorized third parties.")

    add_subheading("4.6 Insider Threats")
    add_p("Insider threats occur when authorized personnel misuse their access privileges to copy sensitive data or damage systems.")
    add_p("Example: A departing employee downloading client databases before leaving the business.")

    add_table(
        "Table 4.1: Common SME Cyber Threats Knowledge Base",
        ["Threat Name", "Category", "Likelihood (1-5)", "Impact (1-5)", "Score (1-25)", "Severity", "NIST CSF Code"],
        [
            ["Phishing", "Social Engineering", "5", "4", "20", "Critical", "PR.AT, PR.AA"],
            ["Ransomware", "Extortion Malware", "4", "5", "20", "Critical", "PR.DS, RC.RP"],
            ["Weak Passwords", "Access Control", "5", "4", "20", "Critical", "PR.AA, PR.AT"],
            ["Malware", "Malicious Software", "4", "4", "16", "High", "PR.PS, DE.CM"],
            ["Data Breaches", "Data Security", "3", "5", "15", "High", "PR.DS, ID.AM"],
            ["Insider Threats", "Internal Risk", "2", "4", "8", "Medium", "PR.AA, DE.CM"],
        ],
        [1.3, 1.4, 0.6, 0.6, 0.6, 0.8, 0.9]
    )

    # =============================================================
    # 5. NIST CYBERSECURITY FRAMEWORK 2.0
    # =============================================================
    add_heading("5. NIST CYBERSECURITY FRAMEWORK 2.0")
    add_p("The project adopts the official NIST Cybersecurity Framework (CSF) 2.0 released in 2024. The framework organizes cybersecurity into six core functions:")

    add_subheading("5.1 Govern (GV)")
    add_p("Establishes cybersecurity risk management strategy, expectations, roles, and organizational policy.")

    add_subheading("5.2 Identify (ID)")
    add_p("Identifies and manages cybersecurity risks across systems, people, assets, and operational data.")

    add_subheading("5.3 Protect (PR)")
    add_p("Applies technical and administrative safeguards to protect business systems from unauthorized access.")

    add_subheading("5.4 Detect (DE)")
    add_p("Discovers and analyzes anomalous cybersecurity events and potential unauthorized network intrusions.")

    add_subheading("5.5 Respond (RS)")
    add_p("Executes incident response containment and communication plans when an attack is discovered.")

    add_subheading("5.6 Recover (RC)")
    add_p("Restores business services, databases, and operational capabilities impaired by a cybersecurity incident.")

    add_subheading("5.7 NIST Implementation Tiers")
    add_p("The framework defines four implementation tiers:")
    add_bullet("Tier 1 (Partial): Informal, reactive, and ad-hoc cybersecurity practices.")
    add_bullet("Tier 2 (Risk Informed): Approved practices, but not applied organization-wide.")
    add_bullet("Tier 3 (Repeatable): Formally documented and updated cybersecurity policies.")
    add_bullet("Tier 4 (Adaptive): Continuous improvement and dynamic defense against new threats.")

    add_table(
        "Table 5.1: NIST CSF 2.0 Functions and Operational Scope",
        ["Function", "Code", "Operational Scope", "Key Categories"],
        [
            ["Govern", "GV", "Risk strategy, policy, oversight, and supply chain", "GV.OC, GV.RM, GV.RR, GV.PO, GV.SC"],
            ["Identify", "ID", "Asset management, risk assessment, and improvement", "ID.AM, ID.RA, ID.IM"],
            ["Protect", "PR", "Access control, awareness training, and data security", "PR.AA, PR.AT, PR.DS, PR.PS"],
            ["Detect", "DE", "Continuous monitoring and adverse event analysis", "DE.CM, DE.AE"],
            ["Respond", "RS", "Incident management, reporting, and threat containment", "RS.MA, RS.AN, RS.CO, RS.MI"],
            ["Recover", "RC", "Disaster recovery execution and backup restoration", "RC.RP, RC.CO"],
        ],
        [1.1, 0.6, 2.5, 2.0]
    )

    # =============================================================
    # 6. RISK EVALUATION METRIC
    # =============================================================
    add_heading("6. RISK EVALUATION METRIC")
    add_p("The evaluation metric uses a quantitative mathematical equation to compute threat risk:")
    add_p("Risk Score = Likelihood (1 - 5) × Impact (1 - 5)", bold_prefix="Formula: ")
    add_p("Scores range from 1 to 25 and map into four severity levels:")
    add_bullet("Low (1 - 5): Minor impact, handled by baseline controls.")
    add_bullet("Medium (6 - 11): Moderate impact, requires scheduled mitigation.")
    add_bullet("High (12 - 19): Serious risk, requires immediate attention.")
    add_bullet("Critical (20 - 25): Catastrophic risk capable of halting operations.")
    add_p("The Overall System Risk Score is the arithmetic average of all threat scores rounded to two decimal places.")

    add_table(
        "Table 6.1: 5x5 Quantitative Risk Evaluation Matrix",
        ["Likelihood \\ Impact", "1 (Minor)", "2 (Low)", "3 (Moderate)", "4 (Major)", "5 (Critical)"],
        [
            ["5 (Almost Certain)", "5 [Low]", "10 [Med]", "15 [High]", "20 [Crit]", "25 [Crit]"],
            ["4 (Likely)", "4 [Low]", "8 [Med]", "12 [High]", "16 [High]", "20 [Crit]"],
            ["3 (Possible)", "3 [Low]", "6 [Med]", "9 [Med]", "12 [High]", "15 [High]"],
            ["2 (Unlikely)", "2 [Low]", "4 [Low]", "6 [Med]", "8 [Med]", "10 [Med]"],
            ["1 (Rare)", "1 [Low]", "2 [Low]", "3 [Low]", "4 [Low]", "5 [Low]"],
        ],
        [1.5, 0.9, 0.9, 0.9, 1.0, 1.0]
    )

    # =============================================================
    # 7. PROPOSED RISK MITIGATION STRATEGIES
    # =============================================================
    add_heading("7. PROPOSED RISK MITIGATION STRATEGIES")
    add_p("The framework includes an automated mitigation recommendation engine. The following core countermeasures are proposed:")

    add_subheading("7.1 Multi-Factor Authentication (MFA)")
    add_p("Enforce MFA on email accounts, cloud storage, and remote access logins to stop credential attacks.")

    add_subheading("7.2 3-2-1 Automated Immutable Backups")
    add_p("Maintain 3 copies of data on 2 different media types, with 1 copy offline or cloud-immutable to protect against ransomware.")

    add_subheading("7.3 Employee Cybersecurity Training")
    add_p("Conduct regular employee awareness training and monthly simulated phishing tests to improve human defense.")

    add_subheading("7.4 Network Firewalls and Segmentation")
    add_p("Isolate point-of-sale payment terminals and guest Wi-Fi networks on separate VLANs using a managed firewall.")

    add_subheading("7.5 Managed Endpoint Detection and Response (EDR)")
    add_p("Deploy modern EDR software on all business computers for behavioral malware detection and isolation.")

    add_subheading("7.6 Principle of Least Privilege (RBAC)")
    add_p("Grant workers only the minimum system access necessary for their specific role.")

    add_table(
        "Table 7.1: Recommended Mitigation Safeguards for SME Risks",
        ["Target Threat", "Recommended Countermeasure", "Type", "NIST Code"],
        [
            ["Phishing", "Employee awareness training and mandatory MFA", "Administrative & Technical", "PR.AT, PR.AA"],
            ["Ransomware", "Automated 3-2-1 immutable cloud backups and EDR", "Technical", "PR.DS, RC.RP"],
            ["Weak Passwords", "Password manager and multi-factor authentication", "Technical", "PR.AA"],
            ["Malware", "Endpoint detection and response (EDR) and patching", "Technical", "PR.PS, DE.CM"],
            ["Data Breaches", "Full-disk encryption (BitLocker) and RBAC permissions", "Technical", "PR.DS, ID.AM"],
            ["Insider Threats", "Least privilege access and administrative logging", "Administrative & Technical", "PR.AA, DE.CM"],
        ],
        [1.3, 2.5, 1.4, 1.0]
    )

    # =============================================================
    # 8. VALIDATION WITH CASE STUDIES
    # =============================================================
    add_heading("8. VALIDATION WITH CASE STUDIES")
    add_p("To validate the model, three simulated small business case studies were evaluated. Residual risk after mitigation is computed using: Residual Score = Initial Score × (1 - Mitigation Efficiency).")

    add_subheading("8.1 Case Study 1: Retail Shop ('Apex Urban Retailers')")
    add_p("Profile: Retail store with 28 employees operating POS terminals and an online Shopify storefront.")
    add_p("Risks: Default POS passwords, phishing invoices, unencrypted customer records, obsolete computer backups.")
    add_p("Mitigations: MFA on POS terminals, payment VLAN segmentation, automated 3-2-1 cloud backups, staff training.")
    add_p("Outcome: Initial score 15.75 (High) reduced to residual score 5.25 (Low), achieving a 66.7% risk reduction.")

    add_subheading("8.2 Case Study 2: Healthcare Clinic ('MetroCare Family Health Center')")
    add_p("Profile: Outpatient medical clinic with 18 staff managing electronic health records and diagnostic imaging.")
    add_p("Risks: Ransomware threat, unencrypted patient database, unmonitored admin access, legacy diagnostic terminals.")
    add_p("Mitigations: Daily immutable cloud backup, role-based access control, BitLocker encryption, isolated VLAN.")
    add_p("Outcome: Initial score 14.75 (High) reduced to residual score 4.50 (Low), achieving a 69.5% risk reduction.")

    add_subheading("8.3 Case Study 3: Small IT Company ('NovaTech Cloud Solutions')")
    add_p("Profile: Software consultancy with 35 engineers developing cloud applications on AWS and GitHub.")
    add_p("Risks: Broad developer cloud access, spear-phishing targeting DevOps, shared staging passwords, hardcoded tokens.")
    add_p("Mitigations: Hardware FIDO2 MFA, automated Git secret scanning, least-privilege IAM roles, password manager.")
    add_p("Outcome: Initial score 16.00 (High) reduced to residual score 4.00 (Low), achieving a 75.0% risk reduction.")

    add_table(
        "Table 8.1: Validation Case Studies: Initial vs Residual Risk Comparison",
        ["Case Profile", "Sector & Workforce", "Initial Score", "Residual Score", "Risk Reduction"],
        [
            ["Retail Shop ('Apex Urban Retailers')", "Retail / 28 Staff", "15.75 [High]", "5.25 [Low]", "66.7%"],
            ["Healthcare Clinic ('MetroCare Health')", "Healthcare / 18 Staff", "14.75 [High]", "4.50 [Low]", "69.5%"],
            ["Small IT Company ('NovaTech Solutions')", "IT & Cloud / 35 Staff", "16.00 [High]", "4.00 [Low]", "75.0%"],
        ],
        [1.8, 1.4, 1.0, 1.0, 1.0]
    )

    # =============================================================
    # 9. FINDINGS AND SME BEST PRACTICES
    # =============================================================
    add_heading("9. FINDINGS AND SME BEST PRACTICES")
    add_p("Based on our evaluation, the following ten cybersecurity best practices are recommended for small businesses:")
    add_bullet("1. Mandatory Multi-Factor Authentication (MFA) on all email, cloud, and VPN accounts (Critical - PR.AC).")
    add_bullet("2. Enforce 3-2-1 Automated Immutable Backups with at least 1 offline copy (Critical - RC.RP).")
    add_bullet("3. Continuous Employee Cybersecurity Awareness Training and monthly phishing simulations (High - PR.AT).")
    add_bullet("4. Apply Principle of Least Privilege and Role-Based Access Control (High - PR.AC).")
    add_bullet("5. Next-Generation Firewalls and Network Segmentation for POS and guest networks (High - PR.PT).")
    add_bullet("6. Automated Software Patching within 14 days of public release (High - PR.MA).")
    add_bullet("7. Deploy Managed Endpoint Detection and Response (EDR) on all workstations (High - DE.CM).")
    add_bullet("8. Full-Disk and In-Transit Encryption via BitLocker and TLS 1.3 (Medium - PR.DS).")
    add_bullet("9. Formal Incident Response Checklist and printed emergency contact sheet (Medium - RS.RP, RC.RP).")
    add_bullet("10. Maintain an Accurate Hardware and Software Asset Inventory (Medium - ID.AM).")

    add_table(
        "Table 9.1: SME Cybersecurity Best Practices Checklist",
        ["Priority", "Best Practice Title", "Core Category", "NIST CSF Function"],
        [
            ["Critical", "1. Mandatory Multi-Factor Authentication (MFA)", "Access Control", "Protect (PR.AC)"],
            ["Critical", "2. Enforce 3-2-1 Automated Immutable Backups", "Data Recovery", "Recover (RC.RP)"],
            ["High", "3. Continuous Employee Cybersecurity Training", "Human Defense", "Protect (PR.AT)"],
            ["High", "4. Principle of Least Privilege and RBAC", "Access Management", "Protect (PR.AC)"],
            ["High", "5. Next-Gen Firewalls and Network Segmentation", "Network Security", "Protect (PR.PT)"],
            ["High", "6. Automated Patching and Vulnerability Management", "System Maintenance", "Protect (PR.MA)"],
            ["High", "7. Deploy Managed Endpoint EDR Antivirus", "Threat Detection", "Detect (DE.CM)"],
            ["Medium", "8. Full-Disk and In-Transit Encryption", "Data Protection", "Protect (PR.DS)"],
            ["Medium", "9. Formal Incident Response and Recovery Plan", "Incident Response", "Respond / Recover"],
            ["Medium", "10. Hardware and Software Asset Inventory", "Asset Management", "Identify (ID.AM)"],
        ],
        [0.8, 2.4, 1.4, 1.6]
    )

    # =============================================================
    # 10. TOOLS AND TECHNOLOGIES USED
    # =============================================================
    add_heading("10. TOOLS AND TECHNOLOGIES USED")
    add_p("The project uses standard modern open-source software tools:")
    add_bullet("Windows 10 / 11 Operating System")
    add_bullet("Python 3.13 Runtime Environment")
    add_bullet("FastAPI Web Framework (Port 8000)")
    add_bullet("MongoDB Community Server (Port 27017)")
    add_bullet("Motor Asynchronous Database Driver")
    add_bullet("React 18 Single-Page Application (Port 3000)")
    add_bullet("Tailwind CSS for user interface styling")
    add_bullet("Chart.js for interactive radar and heatmap charts")
    add_bullet("ReportLab library for PDF report generation")

    add_table(
        "Table 10.1: System Software Tools and Versions",
        ["Component", "Technology", "Version / Port", "Operational Purpose"],
        [
            ["Backend Runtime", "Python", "3.13.0", "Core programming language and scoring logic"],
            ["Web Framework", "FastAPI", "0.163.0 / Port 8000", "High-performance REST API routing"],
            ["Database Server", "MongoDB", "7.0.14 / Port 27017", "NoSQL database for users and assessments"],
            ["Database Driver", "Motor", "3.7.1", "Asynchronous non-blocking MongoDB communication"],
            ["Security & Auth", "Passlib & PyJWT", "1.7.4 / HS256", "Bcrypt hashing and JWT bearer authentication"],
            ["PDF Engine", "ReportLab", "5.1.0", "Single-page executive PDF report compilation"],
            ["Frontend UI", "React", "18.2.0 / Port 3000", "Component-based web application"],
            ["CSS Styling", "Tailwind CSS", "3.1.8", "Utility-first dark-theme user interface"],
            ["Visual Charts", "Chart.js", "4.6.1", "Rendering radar charts and threat heatmaps"],
        ],
        [1.3, 1.4, 1.4, 2.1]
    )

    # =============================================================
    # 11. SYSTEM ARCHITECTURE
    # =============================================================
    add_heading("11. SYSTEM ARCHITECTURE")
    add_p("The SME Risk Framework follows a client-server architecture:")
    add_bullet("Frontend: React application running in the user's web browser, managing input forms, radar charts, heatmaps, and history tables.")
    add_bullet("Backend: Python FastAPI server handling user authentication, input validation, risk scoring equations, and ReportLab PDF compilation.")
    add_bullet("Database: MongoDB database storing hashed user passwords and assessment evaluation records.")

    add_p("System Workflow:")
    add_bullet("1. The user registers an account and signs in with email and password.")
    add_bullet("2. The user enters system details, chooses controls, and rates threats (Likelihood and Impact 1 to 5).")
    add_bullet("3. The backend calculates threat scores, severity bands, domain coverage, and actionable findings.")
    add_bullet("4. The dashboard displays the results via a 6-axis radar chart and 5x5 threat heatmap.")
    add_bullet("5. The user can export the assessment data as JSON or download a single-page executive PDF report.")

    # =============================================================
    # 12. APPLICATION WALKTHROUGH
    # =============================================================
    add_heading("12. APPLICATION WALKTHROUGH")
    add_p("The application workflow includes the following steps:")

    add_subheading("12.1 Account Registration")
    add_p("The user visits the Create Account page (/register) and enters an email address, password, and password confirmation.")

    add_subheading("12.2 User Sign In")
    add_p("The user logs in (/login) with registered credentials to receive a signed JWT access token.")

    add_subheading("12.3 Assessment Form and Control Selection")
    add_p("The user enters system details, selects controls, and assigns each control to a NIST CSF function.")

    add_subheading("12.4 Threat Rating")
    add_p("The user enters identified threats and assigns Likelihood (1-5) and Impact (1-5) scores.")

    add_subheading("12.5 Interactive Visualizations")
    add_p("The dashboard renders a 6-axis NIST CSF radar chart and a 5x5 color-coded threat heatmap.")

    add_subheading("12.6 History Review and PDF Export")
    add_p("The user views saved assessments and downloads a single-page executive PDF report.")

    # =============================================================
    # 13. RESULTS AND ANALYSIS
    # =============================================================
    add_heading("13. RESULTS AND ANALYSIS")
    add_p("The application was tested locally and evaluated across the three case studies:")

    add_subheading("13.1 Result Interpretation")
    add_p("Retail Shop risk dropped from 15.75 (High) to 5.25 (Low), achieving a 66.7% risk reduction.")
    add_p("Healthcare Clinic risk dropped from 14.75 (High) to 4.50 (Low), achieving a 69.5% risk reduction.")
    add_p("Small IT Company risk dropped from 16.00 (High) to 4.00 (Low), achieving a 75.0% risk reduction.")
    add_p("Across all three case studies, the average risk reduction was 70.4%.")

    add_subheading("13.2 Analysis")
    add_p("The test confirms that the application's risk calculation formulas, domain mapping, and PDF exports are operational.")
    add_p("Foundational technical controls (MFA, immutable backups, and patching) provide the largest reduction in overall risk.")

    # =============================================================
    # 14. TESTING AND VERIFICATION
    # =============================================================
    add_heading("14. TESTING AND VERIFICATION")
    add_p("The risk model was tested by executing test cases directly against risk_model.py:")

    add_table(
        "Table 14.1: Verification Test Cases and Execution Results",
        ["Test ID", "Scenario", "Input Threats & Controls", "Expected Output", "Actual Output", "Result"],
        [
            ["TC-01", "Critical Ransomware Threat", "Ransomware (5,5), Backup (Recover)", "Score: 25.0 [Critical], Recover: 100%", "Score: 25.0 [Critical], Recover: 100%", "Pass"],
            ["TC-02", "Dual High Threats", "Phishing (4,4), Passwords (3,4), MFA (Protect)", "Score: 14.0 [High], Protect: 100%", "Score: 14.0 [High], Protect: 100%", "Pass"],
            ["TC-03", "Low Risk Threat", "Malware (1,2), Firewall (Protect)", "Score: 2.0 [Low], Protect: 100%", "Score: 2.0 [Low], Protect: 100%", "Pass"],
        ],
        [0.6, 1.2, 1.6, 1.3, 1.3, 0.4]
    )

    # =============================================================
    # 15. LESSONS LEARNED
    # =============================================================
    add_heading("15. LESSONS LEARNED")
    add_p("The project provided practical knowledge about small business cybersecurity:")
    add_bullet("Simple, quantitative risk models are more effective for SMEs than complex enterprise catalogs.")
    add_bullet("Multi-factor authentication is the most critical low-cost defense against credential attacks.")
    add_bullet("Immutable offline backups are essential to protect businesses against ransomware extortion.")
    add_bullet("Automated scoring provides consistent benchmarks and eliminates subjective human bias.")
    add_bullet("Clean visual interfaces help non-technical business owners understand risk findings quickly.")

    # =============================================================
    # 16. LIMITATIONS
    # =============================================================
    add_heading("16. LIMITATIONS")
    add_p("The project has the following limitations:")
    add_bullet("The application relies on user-provided ratings rather than active network port scanning.")
    add_bullet("The three case studies are simulated scenarios designed for model validation.")
    add_bullet("Real-time threat intelligence feeds are not currently integrated.")

    # =============================================================
    # 17. FUTURE SCOPE
    # =============================================================
    add_heading("17. FUTURE SCOPE")
    add_p("The project can be enhanced in the following ways:")
    add_bullet("Integrate automated port scanning tools such as Nmap.")
    add_bullet("Connect live CVE databases for real-time vulnerability scoring.")
    add_bullet("Add multi-tenant management for IT service providers (MSPs).")
    add_bullet("Add automated compliance checklists for GDPR, HIPAA, and PCI-DSS.")

    # =============================================================
    # 18. CONCLUSION
    # =============================================================
    add_heading("18. CONCLUSION")
    add_p("The project Cybersecurity Risk Assessment Framework for Small Businesses demonstrates that practical, quantitative risk evaluation can be accessible to small organizations.")
    add_p("A local web application named SME Risk Framework was successfully developed and tested.")
    add_p("The application includes risk scoring, NIST CSF 2.0 mapping, interactive radar charts, heatmaps, and executive PDF audit reports.")
    add_p("Validation across three simulated case studies confirmed an average risk reduction of 70.4% following the adoption of core safeguards.")

    # =============================================================
    # 19. REFERENCES
    # =============================================================
    add_heading("19. REFERENCES")
    add_p("The following resources are recommended for further reading and verification:")
    refs_data = [
        "National Institute of Standards and Technology (NIST) – CSF 2.0: https://doi.org/10.6028/NIST.CSWP.29",
        "NIST SP 800-30 Revision 1 – Guide for Conducting Risk Assessments: https://doi.org/10.6028/NIST.SP.800-30r1",
        "Cybersecurity and Infrastructure Security Agency (CISA) – Small Business Guidance: https://www.cisa.gov/",
        "ISO/IEC 27001:2022 – Information Security Management Systems: https://www.iso.org/standard/27001",
        "OWASP Foundation – Web Security Guidance: https://owasp.org/",
        "FastAPI Documentation: https://fastapi.tiangolo.com/",
        "React Documentation: https://react.dev/",
        "MongoDB Documentation: https://www.mongodb.com/docs/",
    ]
    for r in refs_data:
        add_bullet(r)

    # =============================================================
    # APPENDIX A – SCREENSHOTS
    # =============================================================
    add_heading("APPENDIX A – SCREENSHOTS")
    add_p("The following screenshots were captured from the running application:")

    # Screenshots in exact sequence requested:
    # Starts with Create Account page (1.png)
    add_screenshot("Figure 1 – Create Account Page", r"C:\Users\uasha\Videos\minor 1\1.png", width_in=5.4)
    add_screenshot("Figure 2 – User Sign In Page", r"C:\Users\uasha\Videos\minor 1\1.1.png", width_in=5.4)
    add_screenshot("Figure 3 – Risk Assessment Dashboard", r"C:\Users\uasha\Videos\minor 1\2.png", width_in=4.8)
    add_screenshot("Figure 4 – Common SME Cyber Risks Research", r"C:\Users\uasha\Videos\minor 1\3.png", width_in=5.2)
    add_screenshot("Figure 5 – NIST Cybersecurity Framework 2.0", r"C:\Users\uasha\Videos\minor 1\4.png", width_in=5.4)
    add_screenshot("Figure 6 – SME Validation Case Studies", r"C:\Users\uasha\Videos\minor 1\5.png", width_in=5.0)
    add_screenshot("Figure 7 – SME Cybersecurity Best Practices Checklist", r"C:\Users\uasha\Videos\minor 1\6.png", width_in=5.2)
    report_dir = os.path.dirname(os.path.abspath(__file__))
    preview_pdf_img = os.path.join(report_dir, "downloaded_test_preview.png")
    add_screenshot("Figure 8 – Single-Page Executive PDF Security Report", preview_pdf_img, width_in=4.8)

    # =============================================================
    # APPENDIX B – PROJECT FILES AND CREDITS
    # =============================================================
    add_heading("APPENDIX B – PROJECT FILES AND CREDITS")
    add_p("Project Files:")
    add_bullet("backend/app/main.py – FastAPI application factory.")
    add_bullet("backend/app/api/v1/routes_assessments.py – Assessment evaluation and CRUD routes.")
    add_bullet("backend/app/api/v1/routes_auth.py – User authentication and registration routes.")
    add_bullet("backend/app/api/v1/routes_export.py – PDF and JSON export generation routes.")
    add_bullet("backend/app/services/risk_model.py – Quantitative risk scoring and NIST CSF 2.0 service.")
    add_bullet("frontend/src/pages/Dashboard.jsx – Risk assessment form and history review interface.")
    add_bullet("frontend/src/data/threats.js – SME cyber threats research knowledge base.")
    add_bullet("run.bat and stop.bat – Single-command orchestration scripts.")

    add_p("Open Source Attribution:")
    add_p("The base application builds on the open-source repository by Abdul Matheen (@abdmath), licensed under the MIT License (https://github.com/abdmath/Security-Risk-Assessment-Dashboard).")
    add_p("Student Additions (Group C):")
    add_bullet("Research on 6 SME cyber threats with likelihood and impact ratings.")
    add_bullet("Full NIST CSF 2.0 model alignment including the Govern function.")
    add_bullet("Quantitative 5x5 evaluation metric and automated severity classification.")
    add_bullet("Automated threat-to-safeguard recommendation engine.")
    add_bullet("Three simulated SME validation case studies with residual risk scoring.")
    add_bullet("Ten prioritized SME cybersecurity best practices checklist.")
    add_bullet("Complete user interface redesign with dark theme and orange accents.")
    add_bullet("Single-page executive PDF security assessment report generator.")
    add_bullet("Registration password confirmation and persistent login state.")

    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r_end = p_end.add_run("END OF REPORT")
    r_end.font.name = "Times New Roman"
    r_end.font.size = Pt(14)
    r_end.font.bold = True

    output_path = os.path.join(report_dir, "Minor_Project_1_Report.docx")
    doc.save(output_path)
    print(f"Successfully saved {output_path}")

if __name__ == '__main__':
    build_perfect_report()
