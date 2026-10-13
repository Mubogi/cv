"""Build the CV (DOCX + PDF + Markdown) from one content model.

Written in the tone of the applicant's original CV, with honest additions.
Simple English. No "job I am looking for" section. Regenerate with:
    python3 build_cv.py
"""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, HRFlowable,
                                Table, TableStyle, Image as RLImage, Spacer)

HERE = os.path.dirname(os.path.abspath(__file__))
PHOTO = os.path.join(HERE, "photo_placeholder.jpg")

# Muted, professional palette (deep slate blue + warm grey)
NAVY = "1F3A5F"
NAVY_RGB = RGBColor(0x1F, 0x3A, 0x5F)
GREY = RGBColor(0x5A, 0x5A, 0x5A)
DARK = RGBColor(0x22, 0x22, 0x22)

# ----------------------------------------------------------------------------
# Content
# ----------------------------------------------------------------------------
NAME = "MUBOGI GASTAVAS JORDAN"

CONTACT = [
    "Kampala, Uganda",
    "+256 754 687 597",
    "+256 768 593 543",
    "mubogigastavasjordan@gmail.com",
]

PROFESSIONAL_SUMMARY = (
    "Results-driven professional with expertise in leadership, education, technical "
    "solutions, and strategic communication. Experienced in team coordination, event "
    "planning, financial management, and digital branding. Proven ability to lead "
    "initiatives, manage projects, and optimize operations across different sectors."
)

SKILLS = [
    ("Leadership & Project Management", "Initiative coordination, team mentorship, project planning."),
    ("Technical & Digital Solutions",
     "Web design, graphics design, IT training, computer support and troubleshooting, "
     "Microsoft Office, building websites and simple apps with AI-assisted tools."),
    ("Finance & Organizational Strategy",
     "Budgeting, credit assessment, workflow optimization, record keeping."),
    ("Communication & Marketing",
     "Public speaking, digital branding, campaign strategies, customer care."),
    ("Media & Creative",
     "Graphic design (Canva), video editing, content creation, sound and media for church services and events."),
]

WORK_EXPERIENCE = [
    ("ICT Volunteer", "UCSCU - Uganda Cooperative Savings and Credit Union Limited", "2026 - Present",
     "Provide day-to-day ICT support, set up and maintain computers, printers and internet, help staff with "
     "computer problems, and assist with other office duties."),
    ("Direct Sales Executive", "Airtel Uganda", "2025",
     "Marketing airtel Internet or wifi."),
    ("Code Instructor (Intern)", "Code Academy Uganda", "2025",
     "Helped students in accomplishing their code projects in the Code Camp."),
    ("Physics & Mathematics Teacher", "Destiny College International", "2025",
     "Delivered comprehensive STEM education, enhancing student performance."),
    ("Credit Officer", "Credit Watch Services International", "Previous Role, 2025",
     "Managed client financial assessments, providing structured financial guidance."),
    ("ICT Teacher", "Semuto Senior Secondary School", "Previous Role, 2024 - 2025",
     "Designed and implemented effective digital learning strategies."),
    ("Marketer", "Apex International University", "2023 - 2024",
     "Executed marketing strategies improving student engagement."),
    ("Graphics Designer", "Magix Graphics", "2022 - 2023",
     "Developed promotional materials to enhance brand visibility."),
    ("Sales Promoter", "Jajopa International Limited", "2019 - 2021",
     "Conducted awareness campaigns and product distribution."),
]

ADDITIONAL_ROLES = [
    ("Team Lead - Media & Technical Operations", "2024 - Present",
     "Oversaw translation services, sound systems, and media production. Managed technical setups for "
     "communication and public engagement."),
    ("Founder & Chairperson - Apex Christian Union", "2023 - 2025",
     "Established a Christian Union fostering spiritual growth, collaboration and mentorship among students. "
     "Organized large-scale events like outreaches to enhance spiritual and educational engagement."),
    ("GRC - IT Faculty, Apex International University", "2023 - 2024",
     "Represented student interests and facilitated IT faculty initiatives. Advocated for student needs and "
     "strengthened academic collaboration."),
]

# Honest project list, in the same plain tone
PROJECTS = [
    ("Attendance Hub - Church & Event Attendance (in use)",
     "Records church and event attendance, keeps member details and prints reports. Running on a computer, with an Android app version."),
    ("School Management System (in use)",
     "Helps a school manage students, teachers, fees, marks and reports, and works without internet."),
    ("Savings Bank SACCO App (in use)",
     "A simple savings and credit app for SACCO members, made with Flutter."),
    ("Church Media & Sound (ongoing)",
     "Running sound, media and livestreaming for church services and events, and training others to do it."),
    ("ClinicSync - Clinic & Pharmacy System (in development)",
     "A system for clinics and drug shops covering sales, stock with expiry dates, daily expenses and end-of-day cash. Still being built and tested."),
    ("UCSCU Connect (in development)",
     "A system being built to help UCSCU with its work. Still in development."),
    ("JD Hub CBS - SACCO Banking System (prototype)",
     "A banking-style system for SACCOs, built as a prototype to learn how core banking works."),
    ("MyMarket.ug - Online Market (in development)",
     "An online market where vendors can show their goods and customers can pay by mobile money. Still being finished."),
    ("BizDocs - Invoices for Small Businesses (in development)",
     "An app for making invoices, receipts and quotations for small businesses. Still being built."),
    ("Kids AI Studio (prototype)",
     "Turns one sentence into a short children's video or song using AI."),
    ("MT5 Trading Bot (prototype)",
     "A practice bot built to learn about trading, money and risk."),
    ("KJ Chat App (prototype)",
     "A chat app built to learn how messaging apps work."),
    ("JD Hub Website (in development)",
     "The website for Jordan Design Hub."),
]

EDUCATION = [
    ("Apex International University",
     "Bachelor's Degree in Information Technology - Graduated 2026"),
    ("Certificate in Theology", "In progress"),
    ("Mbale Secondary School", "Uganda Advanced Certificate of Education (2019 - 2020)"),
    ("Lugazi Homestone Boarding School", "Uganda Certificate of Education (2015 - 2018)"),
    ("Namakwekwe Primary School", "Primary Leaving Examination (2008 - 2014)"),
]

SELF_TAUGHT = [
    "Graphic design with Canva - posters, logos and branding.",
    "Video editing and content creation.",
    "Church media, sound systems and livestreaming.",
    "AI tools (OpenHands) and how to build and use AI agents.",
    "Building websites and simple apps by working on real projects.",
    "Teaching and helping others learn computer skills.",
]

LANGUAGES = [
    "English: Excellent (Speaking, Reading, Writing)",
    "Luganda: Very Good (Reading), Good (Speaking, Writing)",
    "Lumasaba: Excellent (Speaking), Very Good (Reading), Good (Writing)",
]

REFERENCES = [
    ("Nangede Jacqueline", "Director, Jajopa International Limited", "0781 934 093"),
    ("Emmanuel Buyi", "Wallace Engineering Services", "0753 190 857"),
    ("Mr. Nsereko Elisha", "Dean of Students, Apex International University", "+256 760 749 685"),
]

# ----------------------------------------------------------------------------
# DOCX helpers
# ----------------------------------------------------------------------------
def shade_para(paragraph, hexcolor):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    pPr.append(shd)

def shade_cell(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)

def heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    shade_para(p, NAVY)
    r = p.add_run("  " + text.upper())
    r.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    return p

def bullet(doc, text, bold_lead=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    if bold_lead:
        r = p.add_run(bold_lead)
        r.bold = True
    p.add_run(text)

def build_docx(path):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10.5)

    for s in doc.sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.6)
        s.right_margin = Inches(0.6)

    # Header: photo + name/contact, subtle grey rule under
    hdr = doc.add_table(rows=1, cols=2)
    hdr.autofit = False
    hdr.columns[0].width = Inches(1.4)
    hdr.columns[1].width = Inches(5.9)
    cell_photo, cell_info = hdr.cell(0, 0), hdr.cell(0, 1)
    cell_photo.paragraphs[0].text = ""
    cell_photo.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(PHOTO):
        cell_photo.paragraphs[0].add_run().add_picture(PHOTO, width=Inches(1.2))

    p = cell_info.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(NAME)
    r.bold = True
    r.font.size = Pt(20)
    r.font.color.rgb = NAVY_RGB
    p = cell_info.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(" · ".join(CONTACT))
    r.font.size = Pt(9)
    r.font.color.rgb = GREY

    rule = doc.add_paragraph()
    rule.paragraph_format.space_before = Pt(2)
    rule.paragraph_format.space_after = Pt(6)
    pPr = rule._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), NAVY)
    pbdr.append(bottom)
    pPr.append(pbdr)

    heading(doc, "Professional Summary")
    p = doc.add_paragraph(PROFESSIONAL_SUMMARY)
    p.paragraph_format.space_after = Pt(4)

    heading(doc, "Work Experience")
    for role, org, when, desc in WORK_EXPERIENCE:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(f"{role} - {org}")
        r.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = NAVY_RGB
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(when)
        r.italic = True
        r.font.size = Pt(9)
        r.font.color.rgb = GREY
        p = doc.add_paragraph(desc)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Pt(6)

    heading(doc, "Skills & Proficiencies")
    for label, val in SKILLS:
        bullet(doc, val, bold_lead=f"{label}: ")

    heading(doc, "Projects & Practical Work")
    for title, desc in PROJECTS:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"{title}: ")
        r.bold = True
        p.add_run(desc)

    heading(doc, "Additional Roles")
    for role, when, desc in ADDITIONAL_ROLES:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(f"{role} ({when})")
        r.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = NAVY_RGB
        p = doc.add_paragraph(desc)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Pt(6)

    heading(doc, "Education")
    for place, detail in EDUCATION:
        bullet(doc, detail, bold_lead=f"{place}: ")

    heading(doc, "What I Have Learned On My Own")
    for s in SELF_TAUGHT:
        bullet(doc, s)

    heading(doc, "Languages")
    for lang in LANGUAGES:
        bullet(doc, lang)

    heading(doc, "References")
    for name, title, phone in REFERENCES:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"{name} - {title} - ")
        p.add_run(phone)

    doc.save(path)
    return path

# ----------------------------------------------------------------------------
# PDF
# ----------------------------------------------------------------------------
def build_pdf(path):
    styles = getSampleStyleSheet()
    navy = HexColor("#" + NAVY)
    grey = HexColor("#5A5A5A")

    name_s = ParagraphStyle('name', parent=styles['Normal'], fontName='Helvetica-Bold',
                            fontSize=21, textColor=navy, leading=24, spaceAfter=3)
    contact_s = ParagraphStyle('contact', parent=styles['Normal'], fontName='Helvetica',
                               fontSize=9, textColor=grey, spaceAfter=0)
    h_s = ParagraphStyle('h', parent=styles['Normal'], fontName='Helvetica-Bold',
                         fontSize=10.2, textColor=HexColor("#FFFFFF"), spaceBefore=7, spaceAfter=3,
                         backColor=navy, borderPadding=(3, 3, 3, 3))
    body = ParagraphStyle('body', parent=styles['Normal'], fontName='Helvetica',
                          fontSize=9.6, leading=12.4, spaceAfter=2, alignment=TA_LEFT)
    bullet_s = ParagraphStyle('bul', parent=body, leftIndent=12, bulletIndent=2, spaceAfter=1.6)
    role_s = ParagraphStyle('role', parent=styles['Normal'], fontName='Helvetica-Bold',
                            fontSize=10.3, textColor=navy, spaceBefore=4, spaceAfter=0)
    meta_s = ParagraphStyle('meta', parent=styles['Normal'], fontName='Helvetica-Oblique',
                            fontSize=9, textColor=grey, spaceAfter=2)
    desc_s = ParagraphStyle('desc', parent=body, leftIndent=6, spaceAfter=2)

    story = []

    info = [Paragraph(NAME, name_s), Paragraph(" · ".join(CONTACT), contact_s)]
    photo_cell = RLImage(PHOTO, width=30 * mm, height=38.5 * mm) if os.path.exists(PHOTO) else ""
    header_tbl = Table([[photo_cell, info]], colWidths=[36 * mm, 145 * mm])
    header_tbl.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (0, 0), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(header_tbl)
    story.append(HRFlowable(width="100%", thickness=1.4, color=navy, spaceBefore=4, spaceAfter=2))

    def h(text):
        story.append(Paragraph(text.upper(), h_s))

    def ul(text):
        story.append(Paragraph(text, bullet_s, bulletText='\u2022'))

    h("Professional Summary")
    story.append(Paragraph(PROFESSIONAL_SUMMARY, body))

    h("Work Experience")
    for role, org, when, desc in WORK_EXPERIENCE:
        story.append(Paragraph(f"{role} - {org}", role_s))
        story.append(Paragraph(when, meta_s))
        story.append(Paragraph(desc, desc_s))

    h("Skills & Proficiencies")
    for label, val in SKILLS:
        story.append(Paragraph(f"<b>{label}:</b> {val}", bullet_s, bulletText='\u2022'))

    h("Projects & Practical Work")
    for title, desc in PROJECTS:
        story.append(Paragraph(f"<b>{title}:</b> {desc}", bullet_s, bulletText='\u2022'))

    h("Additional Roles")
    for role, when, desc in ADDITIONAL_ROLES:
        story.append(Paragraph(f"{role} ({when})", role_s))
        story.append(Paragraph(desc, desc_s))

    h("Education")
    for place, detail in EDUCATION:
        story.append(Paragraph(f"<b>{place}:</b> {detail}", bullet_s, bulletText='\u2022'))

    h("What I Have Learned On My Own")
    for s in SELF_TAUGHT:
        ul(s)

    h("Languages")
    for lang in LANGUAGES:
        ul(lang)

    h("References")
    for name, title, phone in REFERENCES:
        story.append(Paragraph(f"<b>{name}</b> - {title} - {phone}", bullet_s, bulletText='\u2022'))

    doc = SimpleDocTemplate(path, pagesize=A4,
                            leftMargin=14 * mm, rightMargin=14 * mm,
                            topMargin=10 * mm, bottomMargin=10 * mm,
                            title=f"{NAME} - CV")
    doc.build(story)
    return path

# ----------------------------------------------------------------------------
# Markdown
# ----------------------------------------------------------------------------
def build_md(path):
    L = [f"# {NAME}", "", " · ".join(CONTACT), "",
         "> Photo: replace `photo_placeholder.jpg` with your passport photo, then rebuild.", ""]
    L += ["## Professional Summary", PROFESSIONAL_SUMMARY, ""]
    L += ["## Work Experience"]
    for role, org, when, desc in WORK_EXPERIENCE:
        L += [f"### {role} - {org}", f"*{when}*", "", desc, ""]
    L += ["## Skills & Proficiencies"]
    L += [f"- **{label}:** {val}" for label, val in SKILLS]
    L += ["", "## Projects & Practical Work"]
    L += [f"- **{title}:** {desc}" for title, desc in PROJECTS]
    L += ["", "## Additional Roles"]
    for role, when, desc in ADDITIONAL_ROLES:
        L += [f"### {role} ({when})", desc, ""]
    L += ["## Education"]
    L += [f"- **{place}:** {detail}" for place, detail in EDUCATION]
    L += ["", "## What I Have Learned On My Own"]
    L += [f"- {s}" for s in SELF_TAUGHT]
    L += ["", "## Languages"]
    L += [f"- {lang}" for lang in LANGUAGES]
    L += ["", "## References"]
    L += [f"- **{name}** - {title} - {phone}" for name, title, phone in REFERENCES]
    L += [""]
    with open(path, "w") as f:
        f.write("\n".join(L))
    return path


if __name__ == "__main__":
    print(build_docx(os.path.join(HERE, "CV_Mubogi_Jordan.docx")))
    print(build_pdf(os.path.join(HERE, "CV_Mubogi_Jordan.pdf")))
    print(build_md(os.path.join(HERE, "CV_Master.md")))
