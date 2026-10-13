"""Build the CV (DOCX + PDF + Markdown) from one content model.

Honest, general-purpose CV written in simple English. It favours any role the
applicant can do, not only ICT. Regenerate all outputs with: python3 build_cv.py
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
from reportlab.lib.colors import HexColor, white
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, HRFlowable,
                                Table, TableStyle, Image as RLImage, Spacer)

HERE = os.path.dirname(os.path.abspath(__file__))
PHOTO = os.path.join(HERE, "photo_placeholder.jpg")

NAVY = "0B3D91"
LIGHT = "CFE0FF"
ACCENT = RGBColor(0x0B, 0x3D, 0x91)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x55, 0x55, 0x55)

# ----------------------------------------------------------------------------
# Content model
# ----------------------------------------------------------------------------
NAME = "MUBOGI GASTAVAS JORDAN"
TITLE = "ICT & Computer Skills  ·  Teaching  ·  Sales & Customer Care  ·  Media & Design"

CONTACT = [
    "Kampala, Uganda",
    "+256 754 687 597",
    "+256 768 593 543",
    "mubogigastavasjordan@gmail.com",
    "github.com/Mubogi",
]

SUMMARY = (
    "I am an ICT professional from Kampala with a degree in Information Technology. "
    "I have worked in teaching, sales, customer care, finance and media, and I also build "
    "websites and simple apps. I learn new things quickly, I am honest and hardworking, and "
    "I work well with people and under pressure. I am ready to do any job I am given and to "
    "give it my best."
)

QUICK_FACTS = [
    ("What I am looking for", "Any job I can do well - ICT, teaching, sales, customer care, media or office work"),
    ("Availability", "Full-time or part-time; ready to start anytime; can work under pressure"),
    ("Location", "Kampala, Uganda (ready to travel)"),
    ("Currently", "Volunteering at UCSCU; studying a Certificate in Theology"),
    ("Goal", "To keep learning and grow into a bigger role"),
]

SKILLS = [
    ("ICT & Computer Skills",
     "Using and setting up computers, Microsoft Office (Word, Excel, PowerPoint), internet and "
     "email, fixing simple computer problems, teaching computer skills to others"),
    ("Websites & Apps",
     "I can build websites and simple apps using Python, Django, Flask, JavaScript, HTML, CSS, "
     "Flutter and Node.js, with help from AI tools such as OpenHands. I understand how they work, "
     "and I test and set them up myself."),
    ("Design & Media",
     "Graphic design (Canva), making posters and logos, editing videos, content creation, running "
     "sound and media for church services and events, livestreaming"),
    ("Sales & Customer Care",
     "Direct sales, marketing, talking to customers, following up with clients, building good "
     "relationships with people"),
    ("Money & Records",
     "Budgeting, keeping records, assessing loans, managing small funds, organising work"),
    ("Leadership & People",
     "Leading small teams, planning events, public speaking, teaching and mentoring others, "
     "working with different kinds of people"),
]

EXPERIENCE = [
    {
        "role": "ICT Volunteer",
        "org": "UCSCU - Uganda Cooperative Savings and Credit Union Limited",
        "meta": "Kampala, Uganda  |  2026 - Present",
        "bullets": [
            "I help with everyday ICT work at the office.",
            "I set up and take care of computers, printers and the internet.",
            "I help staff when they have computer problems (troubleshooting).",
            "I run errands to town and help with other office duties when needed.",
            "I am learning how SACCOs work from the inside.",
        ],
    },
    {
        "role": "Founder & Product Builder",
        "org": "Jordan Design Hub (JD Hub)",
        "meta": "Kampala, Uganda  |  2025 - Present",
        "bullets": [
            "I started Jordan Design Hub, where I build websites, simple apps and design work for schools, churches, small businesses and SACCOs.",
            "I listen to what the client needs, plan the work, build it (using AI tools such as OpenHands), test it and explain it back to them.",
            "I have built several working systems. Some are now used by churches and schools.",
            "I also design posters, logos and videos for clients.",
            "I handle the money and planning for each project myself.",
        ],
    },
    {
        "role": "Team Lead - Media & Technical Operations",
        "org": "Freelance / Church & Events",
        "meta": "2024 - Present",
        "bullets": [
            "I lead a small team that runs sound, media and translation for services and events.",
            "I set up and run the equipment and help with livestreaming.",
        ],
    },
    {
        "role": "Direct Sales Executive",
        "org": "Airtel Uganda",
        "meta": "2025",
        "bullets": [
            "I sold Airtel internet and WiFi packages and followed up with customers.",
        ],
    },
    {
        "role": "Code Instructor (Intern)",
        "org": "Code Academy Uganda",
        "meta": "2025",
        "bullets": [
            "I helped students with their coding projects and built their confidence.",
        ],
    },
    {
        "role": "Physics & Mathematics Teacher",
        "org": "Destiny College International",
        "meta": "2025",
        "bullets": [
            "I taught Physics and Mathematics and helped students improve their performance.",
        ],
    },
    {
        "role": "Credit Officer",
        "org": "Credit Watch Services International",
        "meta": "2025",
        "bullets": [
            "I checked clients' finances and gave them simple financial advice.",
        ],
    },
    {
        "role": "ICT Teacher",
        "org": "Semuto Senior Secondary School",
        "meta": "2024 - 2025",
        "bullets": [
            "I taught computer skills and helped students learn practical ICT.",
        ],
    },
    {
        "role": "Marketer",
        "org": "Apex International University",
        "meta": "2023 - 2024",
        "bullets": [
            "I did marketing that increased student interest and enrolment.",
        ],
    },
    {
        "role": "Graphics Designer",
        "org": "Magix Graphics",
        "meta": "2022 - 2023",
        "bullets": [
            "I designed posters and other materials for clients.",
        ],
    },
    {
        "role": "Sales Promoter",
        "org": "Jajopa International Limited",
        "meta": "2019 - 2021",
        "bullets": [
            "I carried out awareness campaigns and distributed products.",
        ],
    },
]

# Honest project list: what is in use, what is still being built.
PROJECTS = [
    ("Attendance Hub - Church & Event Attendance (in use)", [
        "A system that records who attends church services and events, keeps member details and prints reports.",
        "It is running on a computer, and I also made an Android app version.",
    ]),
    ("School Management System (in use)", [
        "Helps a school manage students, teachers, fees, marks and reports. It works without internet.",
        "It is installed and used by a school.",
    ]),
    ("Savings Bank SACCO App (in use)", [
        "A simple savings and credit app for SACCO members, made with Flutter.",
    ]),
    ("Church Media & Sound (ongoing)", [
        "I run sound, media and livestreaming for church services and events, and train others to do it.",
    ]),
    ("ClinicSync - Clinic & Pharmacy System (in development)", [
        "A system for clinics and drug shops: sales, stock with expiry dates, daily expenses and end-of-day cash.",
        "It is still being built and tested, not yet in use.",
    ]),
    ("UCSCU Connect (in development)", [
        "A system being built to help UCSCU with its work. It is still in development and not yet in use.",
    ]),
    ("JD Hub CBS - SACCO Banking System (prototype)", [
        "A banking-style system for SACCOs, built as a prototype to learn how core banking works.",
    ]),
    ("MyMarket.ug - Online Market (in development)", [
        "An online market where vendors can show their goods and customers can pay by mobile money. Still being finished.",
    ]),
    ("BizDocs - Invoices for Small Businesses (in development)", [
        "An app for making invoices, receipts and quotations for small businesses. Still being built.",
    ]),
    ("Kids AI Studio (prototype)", [
        "A fun project that turns one sentence into a short children's video or song using AI.",
    ]),
    ("MT5 Trading Bot (prototype)", [
        "A practice bot I built to learn about trading, money and risk.",
    ]),
    ("KJ Chat App (prototype)", [
        "A chat app I built to learn how messaging apps work.",
    ]),
    ("JD Hub Website (in development)", [
        "The website for my business, Jordan Design Hub.",
    ]),
]

LEADERSHIP = [
    ("Founder & Chairperson - Apex Christian Union", "2023 - 2025",
     "I started the Christian Union at my university and led it, organising outreaches and events."),
    ("GRC - IT Faculty, Apex International University", "2023 - 2024",
     "I represented students and helped with faculty activities."),
    ("Team Lead - Media & Technical Operations", "2024 - Present",
     "I lead the media and sound team at church."),
]

EDUCATION = [
    ("Bachelor of Science in Information Technology",
     "Apex International University  |  Graduated 2026"),
    ("Certificate in Theology", "In progress  |  Currently studying"),
    ("Uganda Advanced Certificate of Education (UACE)", "Mbale Secondary School  |  2019 - 2020"),
    ("Uganda Certificate of Education (UCE)", "Lugazi Homestone Boarding School  |  2015 - 2018"),
    ("Primary Leaving Examination (PLE)", "Namakwekwe Primary School  |  2008 - 2014"),
]

SELF_TAUGHT = [
    "Graphic design with Canva - posters, logos and branding",
    "Video editing and content creation",
    "Church media, sound systems and livestreaming",
    "AI tools (OpenHands) and how to build and use AI agents",
    "Building websites and simple apps by working on real projects",
    "Teaching and helping other people learn computer skills",
]

LANGUAGES = [
    "English - excellent (speaking, reading, writing)",
    "Luganda - very good (reading); good (speaking, writing)",
    "Lumasaba - excellent (speaking); very good (reading); good (writing)",
]

REFERENCES = [
    "Nangede Jacqueline - Director, Jajopa International Limited - 0781 934 093",
    "Emmanuel Buyi - Wallace Engineering Services - 0753 190 857",
    "Mr. Nsereko Elisha - Dean of Students, Apex International University - +256 760 749 685",
]

# ----------------------------------------------------------------------------
# DOCX helpers
# ----------------------------------------------------------------------------
def shade(element, hexcolor):
    pPr = element._p.get_or_add_pPr() if hasattr(element, '_p') else element._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    if hasattr(element, '_p'):
        pPr.append(shd)
    else:
        pPr.append(shd)

def shade_cell(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)

def shade_para(paragraph, hexcolor):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    pPr.append(shd)

def heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Pt(0)
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
        s.top_margin = Inches(0.45)
        s.bottom_margin = Inches(0.45)
        s.left_margin = Inches(0.55)
        s.right_margin = Inches(0.55)

    # Header banner (navy) with photo + name
    hdr = doc.add_table(rows=1, cols=2)
    hdr.autofit = False
    hdr.columns[0].width = Inches(1.4)
    hdr.columns[1].width = Inches(5.9)
    cell_photo = hdr.cell(0, 0)
    cell_info = hdr.cell(0, 1)
    shade_cell(cell_photo, NAVY)
    shade_cell(cell_info, NAVY)

    cell_photo.paragraphs[0].text = ""
    cell_photo.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(PHOTO):
        cell_photo.paragraphs[0].add_run().add_picture(PHOTO, width=Inches(1.2))

    p = cell_info.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(NAME)
    r.bold = True
    r.font.size = Pt(19)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p = cell_info.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(TITLE)
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0xCF, 0xE0, 0xFF)
    p = cell_info.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(" | ".join(CONTACT))
    r.font.size = Pt(8.5)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    heading(doc, "About Me")
    p = doc.add_paragraph(SUMMARY)
    p.paragraph_format.space_after = Pt(4)

    heading(doc, "Quick Facts")
    for label, val in QUICK_FACTS:
        bullet(doc, val, bold_lead=f"{label}: ")

    heading(doc, "Skills")
    for label, val in SKILLS:
        bullet(doc, val, bold_lead=f"{label}: ")

    heading(doc, "Work Experience")
    for job in EXPERIENCE:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(f"{job['role']} - {job['org']}")
        r.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = ACCENT
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(job['meta'])
        r.italic = True
        r.font.size = Pt(9)
        r.font.color.rgb = GREY
        for b in job['bullets']:
            bullet(doc, b)

    heading(doc, "Projects & Practical Work")
    for title, bullets in PROJECTS:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(10.5)
        for b in bullets:
            bullet(doc, b)

    heading(doc, "Leadership & Volunteer Roles")
    for role, when, desc in LEADERSHIP:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"{role} ({when}): ")
        r.bold = True
        p.add_run(desc)

    heading(doc, "Education & Training")
    for award, detail in EDUCATION:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"{award} - ")
        r.bold = True
        p.add_run(detail)

    heading(doc, "What I Have Learned On My Own")
    for s in SELF_TAUGHT:
        bullet(doc, s)

    heading(doc, "Languages")
    for lang in LANGUAGES:
        bullet(doc, lang)

    heading(doc, "References")
    for ref in REFERENCES:
        bullet(doc, ref)

    doc.save(path)
    return path

# ----------------------------------------------------------------------------
# PDF
# ----------------------------------------------------------------------------
def build_pdf(path):
    styles = getSampleStyleSheet()
    navy = HexColor("#0B3D91")
    grey = HexColor("#555555")

    name_s = ParagraphStyle('name', parent=styles['Normal'], fontName='Helvetica-Bold',
                            fontSize=19, textColor=white, leading=22, spaceAfter=2)
    title_s = ParagraphStyle('title', parent=styles['Normal'], fontName='Helvetica',
                             fontSize=9.5, textColor=HexColor("#CFE0FF"), alignment=TA_LEFT, spaceAfter=3)
    contact_s = ParagraphStyle('contact', parent=styles['Normal'], fontName='Helvetica',
                               fontSize=8.5, textColor=white, spaceAfter=0)
    h_s = ParagraphStyle('h', parent=styles['Normal'], fontName='Helvetica-Bold',
                         fontSize=10.2, textColor=white, spaceBefore=9, spaceAfter=4,
                         backColor=navy, borderPadding=(4, 4, 4, 4), leftIndent=0)
    body = ParagraphStyle('body', parent=styles['Normal'], fontName='Helvetica',
                          fontSize=9.5, leading=12.6, spaceAfter=2, alignment=TA_LEFT)
    bullet_s = ParagraphStyle('bul', parent=body, leftIndent=12, bulletIndent=2, spaceAfter=1.8)
    role_s = ParagraphStyle('role', parent=styles['Normal'], fontName='Helvetica-Bold',
                            fontSize=10.2, textColor=navy, spaceBefore=5, spaceAfter=0)
    meta_s = ParagraphStyle('meta', parent=styles['Normal'], fontName='Helvetica-Oblique',
                            fontSize=8.8, textColor=grey, spaceAfter=1)

    story = []

    info = [
        Paragraph(NAME, name_s),
        Paragraph(TITLE, title_s),
        Paragraph(" | ".join(CONTACT), contact_s),
    ]
    photo_cell = RLImage(PHOTO, width=30 * mm, height=38.5 * mm) if os.path.exists(PHOTO) else ""
    header_tbl = Table([[photo_cell, info]], colWidths=[36 * mm, 145 * mm])
    header_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), navy),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (0, 0), 8),
        ('LEFTPADDING', (1, 0), (1, 0), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(header_tbl)
    story.append(Spacer(1, 4))

    def h(text):
        story.append(Paragraph(text.upper(), h_s))

    def ul(text):
        story.append(Paragraph(text, bullet_s, bulletText='\u2022'))

    h("About Me")
    story.append(Paragraph(SUMMARY, body))

    h("Quick Facts")
    for label, val in QUICK_FACTS:
        story.append(Paragraph(f"<b>{label}:</b> {val}", bullet_s, bulletText='\u2022'))

    h("Skills")
    for label, val in SKILLS:
        story.append(Paragraph(f"<b>{label}:</b> {val}", bullet_s, bulletText='\u2022'))

    h("Work Experience")
    for job in EXPERIENCE:
        story.append(Paragraph(f"{job['role']} - {job['org']}", role_s))
        story.append(Paragraph(job['meta'], meta_s))
        for b in job['bullets']:
            ul(b)

    h("Projects & Practical Work")
    for title, bullets in PROJECTS:
        story.append(Paragraph(title, role_s))
        for b in bullets:
            ul(b)

    h("Leadership & Volunteer Roles")
    for role, when, desc in LEADERSHIP:
        story.append(Paragraph(f"<b>{role}</b> ({when}): {desc}", bullet_s, bulletText='\u2022'))

    h("Education & Training")
    for award, detail in EDUCATION:
        story.append(Paragraph(f"<b>{award}</b> - {detail}", bullet_s, bulletText='\u2022'))

    h("What I Have Learned On My Own")
    for s in SELF_TAUGHT:
        ul(s)

    h("Languages")
    for lang in LANGUAGES:
        ul(lang)

    h("References")
    for ref in REFERENCES:
        ul(ref)

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
    L = [f"# {NAME}", f"**{TITLE}**", "", " · ".join(CONTACT), "",
         "> Photo: replace `photo_placeholder.jpg` with your passport photo, then rebuild.", ""]
    L += ["## About Me", SUMMARY, ""]
    L += ["## Quick Facts"]
    L += [f"- **{label}:** {val}" for label, val in QUICK_FACTS]
    L += ["", "## Skills"]
    L += [f"- **{label}:** {val}" for label, val in SKILLS]
    L += ["", "## Work Experience"]
    for job in EXPERIENCE:
        L += [f"### {job['role']} — {job['org']}", f"*{job['meta']}*", ""]
        L += [f"- {b}" for b in job['bullets']]
        L += [""]
    L += ["## Projects & Practical Work"]
    for title, bullets in PROJECTS:
        L += [f"### {title}"]
        L += [f"- {b}" for b in bullets]
        L += [""]
    L += ["## Leadership & Volunteer Roles"]
    L += [f"- **{role}** ({when}): {desc}" for role, when, desc in LEADERSHIP]
    L += ["", "## Education & Training"]
    L += [f"- **{award}** - {detail}" for award, detail in EDUCATION]
    L += ["", "## What I Have Learned On My Own"]
    L += [f"- {s}" for s in SELF_TAUGHT]
    L += ["", "## Languages"]
    L += [f"- {lang}" for lang in LANGUAGES]
    L += ["", "## References"]
    L += [f"- {ref}" for ref in REFERENCES]
    L += [""]
    with open(path, "w") as f:
        f.write("\n".join(L))
    return path


if __name__ == "__main__":
    print(build_docx(os.path.join(HERE, "CV_Mubogi_Jordan.docx")))
    print(build_pdf(os.path.join(HERE, "CV_Mubogi_Jordan.pdf")))
    print(build_md(os.path.join(HERE, "CV_Master.md")))
