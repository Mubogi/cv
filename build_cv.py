"""Build the master CV (DOCX + PDF + Markdown) from one content model.

Three-page, well-designed CV with a photo placeholder, aimed at being sendable
anywhere. Regenerate all outputs with:  python3 build_cv.py
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
                                Table, TableStyle, Image as RLImage)

HERE = os.path.dirname(os.path.abspath(__file__))
PHOTO = os.path.join(HERE, "photo_placeholder.jpg")

ACCENT = RGBColor(0x0B, 0x3D, 0x91)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x55, 0x55, 0x55)

# ----------------------------------------------------------------------------
# Content model
# ----------------------------------------------------------------------------
NAME = "MUBOGI GASTAVAS JORDAN"
TITLE = "ICT Professional  |  Full-Stack Developer  |  Educator, Leader & Creative"

CONTACT = [
    "Kampala, Uganda",
    "+256 754 687 597",
    "+256 768 593 543",
    "mubogigastavasjordan@gmail.com",
    "github.com/Mubogi",
]

SUMMARY = (
    "Results-driven ICT professional with a Bachelor of Science in Information Technology and "
    "wide experience across software development, teaching, sales, finance, media and leadership. "
    "I have designed, built and shipped over a dozen working software products for Ugandan "
    "schools, SACCOs, clinics, SMEs and churches - including live marketplaces, banking and "
    "clinic systems, mobile apps and a generative-AI kids' studio. I am a self-taught "
    "multi-skilled creator: graphic design, content creation, video editing, church media and ICT "
    "training. I lead teams, plan events, manage money carefully, and I am equally comfortable in "
    "front of a class, a client or a camera. I am open to any role - part-time or full-time - and "
    "I work well under pressure."
)

PROFILE = [
    ("Role I want", "Any lucrative ICT, software, teaching, sales, media or administrative role"),
    ("Availability", "Full-time or part-time; short-term contracts welcome; works well under pressure"),
    ("Location", "Kampala, Uganda (open to remote and travel)"),
    ("Currently", "ICT Volunteer at UCSCU; studying a Certificate in Theology; learning AI agents"),
    ("Next goal", "Enrolling for a Master's in IT / Software Engineering"),
]

SKILLS = [
    ("Software & Web Development",
     "Python, Django 5, Flask, SQLAlchemy, JavaScript, TypeScript, React 18, Node.js, Express, "
     "Socket.IO, GraphQL/Apollo, SQL (SQLite & PostgreSQL), REST APIs, HTML5, CSS3, Bootstrap 5, "
     "Tailwind CSS, Alpine.js, MUI, Jinja2, PWA"),
    ("Mobile & Desktop",
     "Flutter (Dart), Kotlin (native Android, Room), Capacitor, PyInstaller (Windows EXE), "
     "offline-first sync (PouchDB / IndexedDB), LAN device linking"),
    ("AI & Automation",
     "AI-assisted development with OpenHands, learning to design and build AI agents, prompt "
     "engineering, open-source generative AI (text-to-video, image-to-video), ML basics "
     "(scikit-learn), automating content and document generation (ReportLab PDFs)"),
    ("Data & Backend Systems",
     "Database design, Prisma, migrations, authentication, role-based access control, audit "
     "logging, webhooks, mobile-money payments (Flutterwave / MTN MoMo / Airtel)"),
    ("Creative & Media",
     "Graphic design (Canva, InShot), content creation, video editing, church media and sound "
     "systems, livestreaming (OBS / OpenLP), branding and promotional materials"),
    ("IT & Technical Support",
     "ICT training and teaching, computer literacy, web design, networking basics, software "
     "installation, troubleshooting, Git & GitHub, GitHub Actions, cloud deployment (Render), "
     "Linux command line"),
    ("Finance & Organisational Strategy",
     "Budgeting, credit assessment, financial management, workflow optimisation, record keeping, "
     "project planning"),
    ("Leadership & Communication",
     "Team mentorship, initiative coordination, event planning, public speaking, client relations, "
     "translation services, digital branding and campaign strategies"),
]

EXPERIENCE = [
    {
        "role": "ICT Volunteer",
        "org": "UCSCU - Uganda Cooperative Savings and Credit Union Limited",
        "meta": "Kampala, Uganda  |  2026 - Present",
        "bullets": [
            "Building UCSCU Connect, a SACCO apex platform and internal workplace hub, to move apex operations off spreadsheets and paper.",
            "Worked on the SACCO registry with dues tracking, a Central Finance Facility (CFF) inter-lending workflow, and sector KPI dashboards.",
            "Helped add PEARLS quarterly reporting (WOCCU ratios) with CSV exports for regulators such as UMRA and the Bank of Uganda.",
            "Contributed to the training catalogue, events and certification, insurance (CIC Africa) claims, stationery shop, announcements, and governance tools (meetings, e-voting), plus a full admin audit log.",
            "Helped build workplace tools: office attendance on the local network, a visitors register, a staff discussion hub, a whistleblowing channel, document space and a shared calendar.",
            "Set up role-based access so admin, staff, board, SACCO and regulator users each see only what they should.",
        ],
    },
    {
        "role": "Founder & Product Builder",
        "org": "Jordan Design Hub (JD Hub)",
        "meta": "Kampala, Uganda  |  2025 - Present",
        "bullets": [
            "Founded and run Jordan Design Hub, delivering web and mobile apps, an online academy, ICT consultancy and media services to schools, SACCOs, clinics, SMEs and churches in Uganda.",
            "Took more than a dozen products from idea to a working, deployed state - web platforms, mobile apps, desktop software and AI tools (full list under Key Projects).",
            "Handle the whole journey: listening to the client, deciding what to build, building it with AI-assisted development (OpenHands), testing, deploying, and explaining it back to the client.",
            "Set up automated processes that generate whitepaper PDFs, 48 lesson-outline PDFs, 90 marketing posters and a promo video.",
            "Design the brand, marketing materials and content that promote the products, and manage the budgets behind each project.",
        ],
    },
    {
        "role": "Team Lead - Media & Technical Operations",
        "org": "Freelance / Ministry & Events",
        "meta": "2024 - Present",
        "bullets": [
            "Oversaw translation services, sound systems and media production for services and events.",
            "Managed technical setups for communication and public engagement, including livestreaming.",
        ],
    },
    {
        "role": "Direct Sales Executive",
        "org": "Airtel Uganda",
        "meta": "2025",
        "bullets": [
            "Marketed Airtel internet and WiFi packages, growing customer uptake through direct engagement and follow-up.",
        ],
    },
    {
        "role": "Code Instructor (Intern)",
        "org": "Code Academy Uganda",
        "meta": "2025",
        "bullets": [
            "Coached students through their coding projects at the Code Camp, building their confidence and problem-solving skills.",
        ],
    },
    {
        "role": "Physics & Mathematics Teacher",
        "org": "Destiny College International",
        "meta": "2025",
        "bullets": [
            "Delivered comprehensive STEM education, improving student performance through practical, engaging lessons.",
        ],
    },
    {
        "role": "Credit Officer",
        "org": "Credit Watch Services International",
        "meta": "2025",
        "bullets": [
            "Assessed clients' finances and gave structured financial guidance, supporting sound lending decisions.",
        ],
    },
    {
        "role": "ICT Teacher",
        "org": "Semuto Senior Secondary School",
        "meta": "2024 - 2025",
        "bullets": [
            "Designed and delivered effective digital learning strategies, helping students build practical computer skills.",
        ],
    },
    {
        "role": "Marketer",
        "org": "Apex International University",
        "meta": "2023 - 2024",
        "bullets": [
            "Executed marketing strategies that improved student engagement and enrolment interest.",
        ],
    },
    {
        "role": "Graphics Designer",
        "org": "Magix Graphics",
        "meta": "2022 - 2023",
        "bullets": [
            "Developed promotional materials and designs that enhanced brand visibility.",
        ],
    },
    {
        "role": "Sales Promoter",
        "org": "Jajopa International Limited",
        "meta": "2019 - 2021",
        "bullets": [
            "Conducted awareness campaigns and product distribution, growing product reach.",
        ],
    },
]

# Financial / institutional systems
PROJECTS = [
    ("ClinicSync - Clinic, Pharmacy & Drug-Shop System (live, Android + Windows + Linux)", [
        "An offline-first hospital, clinic and drug-shop system for low-resource settings: all reads and writes hit the local database first and sync to the cloud when the network returns.",
        "Point-of-sale with fast drug search, selling by tablet / strip / box, printable 80 mm thermal receipts; daily expenses; and a one-click end-of-day balance sheet with MTN/Airtel mobile-money reconciliation.",
        "Inventory with FEFO (first-expire-first-out) batch expiry and reorder levels; tiered SaaS plans (Basic / Premium / Pro).",
        "Shipped installers for Android (APK), Windows (EXE) and Linux (AppImage / .deb), with device-bound activation licensing.",
        "Built with React/Vite, Tailwind, Express, Prisma and PouchDB.",
    ]),
    ("JD Hub CBS - Co-operative Banking System for SACCOs (prototype)", [
        "A modern core-banking system for SACCOs and credit unions, re-branded from a UCSCU CBS client.",
        "Multi-user demo hub that gives each user an isolated instance behind a reverse proxy (JupyterHub-style).",
        "Built with React 18, TypeScript, MUI, Apollo GraphQL and a mock GraphQL backend.",
    ]),
    ("MyMarket.ug - Multi-Vendor Marketplace (live)", [
        "An online marketplace where Ugandan vendors each get their own shop page; customers browse and buy without an account.",
        "Takes real mobile-money payments through Flutterwave (MTN MoMo / Airtel); product photos are compressed for fast loading on 3G.",
        "Built with Python, Flask, SQLAlchemy, Tailwind CSS and Alpine.js.",
    ]),
    ("BizDocs - Invoicing for Small Businesses (live)", [
        "An offline-first Android app that creates invoices, receipts, quotations and letters, with finger-signature locking and an audit trail.",
        "Handles URA TIN details, UGX amounts and MTN/Airtel mobile-money payment tracking; exports to A4 or thermal PDF.",
        "Built with Flutter (Dart).",
    ]),
    ("UCSCU Connect - SACCO Apex Platform (in progress)", [
        "A platform to help UCSCU run apex operations and internal workplace processes digitally (see the UCSCU role above).",
        "Built with Python, Flask and SQLAlchemy.",
    ]),
    ("Offline-First School Management System (in use)", [
        "Covers students, teachers, fees, marks and reports, works without internet, and uses 9 role-based access levels; licensed per school and tied to the computer (HWID).",
        "Includes a QR parent kiosk and a multi-tenant cloud hub for sync, licensing and backups.",
        "Built with Django as an installable desktop application.",
    ]),
    ("Attendance Hub - Church & Event Check-in (web + native Android)", [
        "A check-in system for churches and events with member registration, reports and Excel export; the native Android version needs no server or internet at all.",
        "The Android app adds LAN device linking so several phones on the same Wi-Fi share attendance in real time, plus activation keys and PDF reports.",
        "Built with Django (web) and Kotlin + Room + NanoHTTPD (Android).",
    ]),
    ("Kids AI Studio - Generative-AI Kids' Video Maker (open source)", [
        "Turns one sentence into a finished animated kids' story video or sing-along song, using open-source generative AI models on a free Kaggle GPU.",
        "Generates scenes, music, narration and captions into a finished MP4, with an image-to-video mode for personal pictures.",
        "Built in Python with ffmpeg and text-to-video / image-to-video models.",
    ]),
    ("MT5 Algorithmic Trading Bot (open source)", [
        "A discipline-first MetaTrader 5 bot with ML signal generation (RandomForest), a 0.01-lot cap, a $10 daily circuit breaker, an 80/20 profit vault and Discord remote control.",
        "Includes a vectorized backtester with walk-forward out-of-sample validation and per-strategy kill switches.",
        "Built in Python (scikit-learn, MT5 API).",
    ]),
    ("Savings Bank SACCO App (recovered & improved)", [
        "A Flutter savings-and-credit app for SACCO members, recovered and improved from an earlier build.",
        "Built with Flutter (Dart).",
    ]),
    ("JD Hub - Company Website & Platform (live)", [
        "The public site for Jordan Design Hub: systems showcase, academy catalogue, services, design gallery and WhatsApp lead capture.",
        "Includes a simple shop with mobile-money payment and a secure 48-hour download link after approval.",
        "Built with Django 5 and Bootstrap 5; deployed on Render.",
    ]),
    ("KJ - Chat App (prototype)", [
        "A messaging app with online status, big file sharing, voice and video calls, customisable themes, a private 'vault mode' and a social feed.",
        "Built with Node.js, Express and Socket.IO, with an Android build packaged for testing.",
    ]),
]

LEADERSHIP = [
    ("Founder & Chairperson - Apex Christian Union", "2023 - 2025",
     "Established the Christian Union, fostering spiritual growth, collaboration and mentorship; organised large-scale outreaches and events."),
    ("GRC - IT Faculty, Apex International University", "2023 - 2024",
     "Represented student interests and facilitated IT faculty initiatives; advocated for student needs and strengthened academic collaboration."),
    ("Team Lead - Media & Technical Operations", "2024 - Present",
     "Led translation services, sound systems and media production; managed technical setups for communication and public engagement."),
]

EDUCATION = [
    ("Bachelor of Science in Information Technology",
     "Apex International University  |  Graduated 2026 (graduation for 2024-2026 completers held together)"),
    ("Certificate in Theology", "In progress  |  Currently studying"),
    ("Uganda Advanced Certificate of Education (UACE)", "Mbale Secondary School  |  2019 - 2020"),
    ("Uganda Certificate of Education (UCE)", "Lugazi Homestone Boarding School  |  2015 - 2018"),
    ("Primary Leaving Examination (PLE)", "Namakwekwe Primary School  |  2008 - 2014"),
]

SELF_TAUGHT = [
    "Graphic design (Canva) - designing posters, branding and promotional materials",
    "Content creation and video editing (InShot and other tools)",
    "Church media and sound systems, plus livestreaming setups",
    "AI agents - how they are built and used (currently learning in depth)",
    "Generative AI for media - text-to-video and image-to-video",
    "Web, mobile and desktop development, learned largely by building real projects",
    "ICT training and teaching - passing skills on to students and colleagues",
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
# DOCX
# ----------------------------------------------------------------------------
def add_hr(doc, color='0B3D91', sz='8'):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), sz)
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pbdr.append(bottom)
    pPr.append(pbdr)

def shade(paragraph, hexcolor):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    pPr.append(shd)

def heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(11)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Pt(4)
    shade(p, 'E8EEF9')
    r = p.add_run("  " + text.upper())
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = ACCENT
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

    hdr = doc.add_table(rows=1, cols=2)
    hdr.autofit = False
    hdr.columns[0].width = Inches(1.35)
    hdr.columns[1].width = Inches(5.9)
    cell_photo = hdr.cell(0, 0)
    cell_info = hdr.cell(0, 1)
    cell_photo.paragraphs[0].text = ""
    if os.path.exists(PHOTO):
        cell_photo.paragraphs[0].add_run().add_picture(PHOTO, width=Inches(1.25))
    cell_photo.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = cell_info.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(NAME)
    r.bold = True
    r.font.size = Pt(20)
    r.font.color.rgb = ACCENT
    p = cell_info.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(TITLE)
    r.font.size = Pt(10.5)
    r.font.color.rgb = GREY
    p = cell_info.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(" | ".join(CONTACT))
    r.font.size = Pt(9)
    r.font.color.rgb = DARK

    add_hr(doc)

    heading(doc, "Professional Summary")
    p = doc.add_paragraph(SUMMARY)
    p.paragraph_format.space_after = Pt(4)

    heading(doc, "Profile")
    for label, val in PROFILE:
        bullet(doc, val, bold_lead=f"{label}: ")

    heading(doc, "Core Skills")
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
        r.font.size = Pt(9.5)
        r.font.color.rgb = GREY
        for b in job['bullets']:
            bullet(doc, b)

    heading(doc, "Key Projects")
    for title, bullets in PROJECTS:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(5)
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

    heading(doc, "Self-Taught Skills")
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

    name_s = ParagraphStyle('name', parent=styles['Title'], fontName='Helvetica-Bold',
                            fontSize=20, textColor=navy, spaceAfter=1, leading=23)
    title_s = ParagraphStyle('title', parent=styles['Normal'], fontName='Helvetica',
                             fontSize=10.5, textColor=grey, alignment=TA_LEFT, spaceAfter=2)
    contact_s = ParagraphStyle('contact', parent=styles['Normal'], fontName='Helvetica',
                               fontSize=8.8, textColor=HexColor("#1A1A1A"), spaceAfter=2)
    h_s = ParagraphStyle('h', parent=styles['Normal'], fontName='Helvetica-Bold',
                         fontSize=10.2, textColor=navy, spaceBefore=8, spaceAfter=3,
                         backColor=HexColor("#E8EEF9"), borderPadding=(3, 3, 3, 3),
                         leftIndent=3)
    body = ParagraphStyle('body', parent=styles['Normal'], fontName='Helvetica',
                          fontSize=9.4, leading=12.4, spaceAfter=2, alignment=TA_LEFT)
    bullet_s = ParagraphStyle('bul', parent=body, leftIndent=12, bulletIndent=2, spaceAfter=1.6)
    role_s = ParagraphStyle('role', parent=styles['Normal'], fontName='Helvetica-Bold',
                            fontSize=10.1, textColor=HexColor("#0B3D91"), spaceBefore=5, spaceAfter=0)
    meta_s = ParagraphStyle('meta', parent=styles['Normal'], fontName='Helvetica-Oblique',
                            fontSize=8.7, textColor=grey, spaceAfter=1)

    story = []

    info = [
        Paragraph(NAME, name_s),
        Paragraph(TITLE, title_s),
        Paragraph(" | ".join(CONTACT), contact_s),
    ]
    photo_cell = RLImage(PHOTO, width=32 * mm, height=41 * mm) if os.path.exists(PHOTO) else ""
    header_tbl = Table([[photo_cell, info]], colWidths=[36 * mm, 145 * mm])
    header_tbl.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (0, 0), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(header_tbl)
    story.append(HRFlowable(width="100%", thickness=1, color=navy, spaceBefore=3, spaceAfter=2))

    def h(text):
        story.append(Paragraph(text.upper(), h_s))

    def ul(text):
        story.append(Paragraph(text, bullet_s, bulletText='\u2022'))

    h("Professional Summary")
    story.append(Paragraph(SUMMARY, body))

    h("Profile")
    for label, val in PROFILE:
        story.append(Paragraph(f"<b>{label}:</b> {val}", bullet_s, bulletText='\u2022'))

    h("Core Skills")
    for label, val in SKILLS:
        story.append(Paragraph(f"<b>{label}:</b> {val}", bullet_s, bulletText='\u2022'))

    h("Work Experience")
    for job in EXPERIENCE:
        story.append(Paragraph(f"{job['role']} - {job['org']}", role_s))
        story.append(Paragraph(job['meta'], meta_s))
        for b in job['bullets']:
            ul(b)

    h("Key Projects")
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

    h("Self-Taught Skills")
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
         "> Photo: place your passport photo where the PDF/DOCX shows `photo_placeholder.jpg`.", ""]
    L += ["## Professional Summary", SUMMARY, ""]
    L += ["## Profile"]
    L += [f"- **{label}:** {val}" for label, val in PROFILE]
    L += ["", "## Core Skills"]
    L += [f"- **{label}:** {val}" for label, val in SKILLS]
    L += ["", "## Work Experience"]
    for job in EXPERIENCE:
        L += [f"### {job['role']} — {job['org']}", f"*{job['meta']}*", ""]
        L += [f"- {b}" for b in job['bullets']]
        L += [""]
    L += ["## Key Projects"]
    for title, bullets in PROJECTS:
        L += [f"### {title}"]
        L += [f"- {b}" for b in bullets]
        L += [""]
    L += ["## Leadership & Volunteer Roles"]
    L += [f"- **{role}** ({when}): {desc}" for role, when, desc in LEADERSHIP]
    L += ["", "## Education & Training"]
    L += [f"- **{award}** - {detail}" for award, detail in EDUCATION]
    L += ["", "## Self-Taught Skills"]
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
