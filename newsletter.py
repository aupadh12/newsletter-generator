import streamlit as st
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, ListFlowable, ListItem, KeepTogether, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib.colors import lightgrey, Color
from reportlab.pdfgen import canvas
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re

# --------- Helper functions ---------

# Custom canvas with watermark
class WatermarkCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        
    def showPage(self):
        self.draw_watermark()
        canvas.Canvas.showPage(self)
        
    def draw_watermark(self):
        self.saveState()
        self.setFillColor(lightgrey)
        self.setFont("Helvetica", 50)
        self.rotate(45)
        self.drawCentredText(400, 100, "NEWSLETTER")
        self.restoreState()

# Make URLs clickable in PDF
def process_content_pdf(content):
    custom_link_pattern = re.compile(r'\[([^\]]+)\]\(([^\)]+)\)')
    processed_content = re.sub(custom_link_pattern, r'<a href="\2" color="blue">\1</a>', content)

    url_pattern = re.compile(r'(?<!href=")(https?://\S+|mailto:\S+|\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,})(?!")')

    def replace_link(match):
        start_pos = match.start()
        before_text = processed_content[:start_pos]
        if '<a ' in before_text and '</a>' not in before_text[before_text.rfind('<a '):]:
            return match.group(0)
        link = match.group(0)
        if '@' in link and not link.startswith("mailto:"):
            link = f"mailto:{link}"
        return f'<a href="{link}" color="blue">{match.group(0)}</a>'

    return re.sub(url_pattern, replace_link, processed_content)

# Convert text to bullet list
def text_to_bullets(text):
    if not text:
        return []
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    return lines

# Add hyperlink in DOCX
def add_hyperlink(paragraph, url, text=None):
    if not text:
        text = url
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    hyperlink.set(qn('w:history'), '1')

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    c = OxmlElement('w:color')
    c.set(qn('w:val'), '0000FF')
    rPr.append(c)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)

    new_run.append(rPr)
    new_run_text = OxmlElement('w:t')
    new_run_text.text = text
    new_run.append(new_run_text)
    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)
    return paragraph

# Add text with links in DOCX
def add_text_with_links(paragraph, text):
    custom_link_pattern = re.compile(r'\[([^\]]+)\]\(([^\)]+)\)')
    last_end = 0

    for match in custom_link_pattern.finditer(text):
        paragraph.add_run(text[last_end:match.start()])
        display_text = match.group(1)
        url = match.group(2)
        add_hyperlink(paragraph, url, display_text)
        last_end = match.end()

    remaining_text = text[last_end:]
    url_pattern = re.compile(r'(https?://\S+|mailto:\S+|\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,})')
    last_end = 0
    for match in url_pattern.finditer(remaining_text):
        paragraph.add_run(remaining_text[last_end:match.start()])
        link = match.group(0)
        if '@' in link and not link.startswith("mailto:"):
            link = f"mailto:{link}"
        add_hyperlink(paragraph, link, match.group(0))
        last_end = match.end()
    paragraph.add_run(remaining_text[last_end:])

# Color mapping
COLOR_OPTIONS = {
    "Cream White": Color(0.98, 0.96, 0.92),
    "Off White": Color(0.96, 0.94, 0.90),
    "Light Blue": Color(0.90, 0.95, 1.0),
    "Light Green": Color(0.90, 0.98, 0.90),
    "Light Pink": Color(0.98, 0.90, 0.95),
    "Light Yellow": Color(0.98, 0.98, 0.85)
}

# --------- PDF Generation ---------
def create_pdf(top_image, intro_text, intro_color, sections, contact_info, bottom_image, month="", year="", issue_number=""):
    buffer = BytesIO()

    PAGE_WIDTH, PAGE_HEIGHT = A4
    LEFT_MARGIN = RIGHT_MARGIN = 0.75 * inch
    TOP_MARGIN = BOTTOM_MARGIN = 0.75 * inch
    CONTENT_WIDTH = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=LEFT_MARGIN,
        rightMargin=RIGHT_MARGIN,
        topMargin=TOP_MARGIN,
        bottomMargin=BOTTOM_MARGIN,
        canvasmaker=WatermarkCanvas
    )

    styles = getSampleStyleSheet()
    intro_style = ParagraphStyle(
        'IntroStyle',
        parent=styles['Normal'],
        backColor=intro_color,
        borderPadding=12
    )

    story = []

    # Header with month, year, and issue number
    if month or year or issue_number:
        header_parts = []
        if month and year:
            header_parts.append(f"{month} {year}")
        elif month:
            header_parts.append(month)
        elif year:
            header_parts.append(year)
        if issue_number:
            header_parts.append(f"Issue #{issue_number}")
        
        header_text = " | ".join(header_parts)
        header_style = ParagraphStyle('HeaderStyle', parent=styles['Heading1'], alignment=1, spaceAfter=0)
        story.append(Paragraph(header_text, header_style))

    # Top image (max 5 inches width)
    if top_image:
        top_image.seek(0)
        img = Image(top_image)
        max_width = min(5*inch, CONTENT_WIDTH)
        if img.drawWidth > max_width:
            img.drawWidth = max_width
            img.drawHeight = img.drawHeight * (max_width / img.imageWidth)
        story.append(img)

    # Intro section full-width
    if intro_text:
        intro_table = Table(
            [[Paragraph(process_content_pdf(intro_text), intro_style)]],
            colWidths=[CONTENT_WIDTH]
        )
        intro_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('RIGHTPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ]))
        story.append(intro_table)
        story.append(Spacer(1, 0.3*inch))

    # Sections
    for section in sections:
        title = section.get("title", "")
        content = section.get("content", "")
        section_color = section.get("color", COLOR_OPTIONS["Off White"])
        bullet_points = text_to_bullets(content) if content else []

        heading_style = ParagraphStyle('HeadingStyle', parent=styles['Heading2'], spaceAfter=6)
        bullet_style = ParagraphStyle('BulletStyle', parent=styles['Normal'], leftIndent=12, bulletIndent=0, spaceAfter=8, leading=18)

        if title or bullet_points:
            table_data = []
            if title:
                table_data.append([Paragraph(f"<b>{title}</b>", heading_style)])
            if bullet_points:
                bullet_items = [ListItem(Paragraph(process_content_pdf(bp.strip()), bullet_style), bulletText='•') for bp in bullet_points if bp.strip()]
                bullet_list = ListFlowable(bullet_items, bulletType='bullet', leftIndent=12)
                table_data.append([bullet_list])
            
            section_table = Table(table_data, colWidths=[CONTENT_WIDTH])
            section_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), section_color),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 12),
                ('RIGHTPADDING', (0, 0), (-1, -1), 12),
                ('TOPPADDING', (0, 0), (-1, -1), 12),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ]))
            story.append(KeepTogether(section_table))
            story.append(Spacer(1, 0.3*inch))

    # Contact Info full-width
    if contact_info:
        contact_table = Table(
            [[Paragraph("<b>Contact Information</b>", styles['Heading2']),
              Paragraph(process_content_pdf(contact_info), styles['Normal'])]],
            colWidths=[CONTENT_WIDTH]
        )
        contact_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('RIGHTPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ]))
        story.append(contact_table)

    # Bottom image (max 4 inches width)
    if bottom_image:
        story.append(Spacer(1, 0.6*inch))
        bottom_image.seek(0)
        img = Image(bottom_image)
        max_width = min(4*inch, CONTENT_WIDTH)
        if img.drawWidth > max_width:
            img.drawWidth = max_width
            img.drawHeight = img.drawHeight * (max_width / img.imageWidth)
        story.append(img)

    doc.build(story)
    buffer.seek(0)
    return buffer

# --------- DOCX Generation ---------
def create_docx(top_image, intro_text, intro_color, sections, contact_info, bottom_image, month="", year="", issue_number=""):
    doc = Document()

    # Header with month, year, and issue number
    if month or year or issue_number:
        header_parts = []
        if month and year:
            header_parts.append(f"{month} {year}")
        elif month:
            header_parts.append(month)
        elif year:
            header_parts.append(year)
        if issue_number:
            header_parts.append(f"Issue #{issue_number}")
        
        header_text = " | ".join(header_parts)
        header_p = doc.add_heading(header_text, level=1)
        header_p.alignment = 1  # Center alignment

    if top_image:
        doc.add_picture(top_image, width=Inches(6))

    if intro_text:
        p = doc.add_paragraph()
        add_text_with_links(p, intro_text)
        doc.add_paragraph()

    for section in sections:
        title = section.get("title", "")
        content = section.get("content", "")
        bullet_points = text_to_bullets(content) if content else []

        if title:
            doc.add_heading(title, level=2)
        if bullet_points:
            for bp in bullet_points:
                p = doc.add_paragraph(style='List Bullet')
                add_text_with_links(p, bp)
        doc.add_paragraph()

    if contact_info:
        doc.add_heading("Contact Information", level=2)
        p = doc.add_paragraph()
        add_text_with_links(p, contact_info)

    if bottom_image:
        doc.add_picture(bottom_image, width=Inches(6))

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --------- Streamlit UI ---------
st.title("📬 Newsletter Generator (Aligned PDF & DOCX with Clickable Links)")
st.write("Fill in the details to create a perfectly aligned newsletter with clickable links, bullet points, intro, contact info, and images.")

# Newsletter header information
col1, col2, col3 = st.columns(3)
with col1:
    month = st.text_input("Month", placeholder="e.g., January")
with col2:
    year = st.text_input("Year", placeholder="e.g., 2024")
with col3:
    issue_number = st.text_input("Issue Number", placeholder="e.g., 1")

top_image = st.file_uploader("Upload Top Image", type=["png", "jpg", "jpeg"])
bottom_image = st.file_uploader("Upload Bottom Image", type=["png", "jpg", "jpeg"])

col1, col2 = st.columns([3, 1])
with col1:
    intro_text = st.text_area("Newsletter Introduction", height=120)
with col2:
    intro_color_name = st.selectbox("Background Color", list(COLOR_OPTIONS.keys()), key="intro_color")
    intro_color = COLOR_OPTIONS[intro_color_name]

if 'sections' not in st.session_state:
    st.session_state.sections = [{"title": "", "content": "", "color": COLOR_OPTIONS["Off White"]}]

num_sections = st.number_input("Number of Sections", min_value=1, max_value=10, value=len(st.session_state.sections))

if len(st.session_state.sections) != num_sections:
    if num_sections > len(st.session_state.sections):
        st.session_state.sections.extend([{"title": "", "content": "", "color": COLOR_OPTIONS["Off White"]}] * (num_sections - len(st.session_state.sections)))
    else:
        st.session_state.sections = st.session_state.sections[:num_sections]

for i in range(num_sections):
    col1, col2, col3 = st.columns([3, 1, 1])
    with col1:
        st.subheader(f"Section {i+1}")
    with col2:
        if st.button("↑", key=f"up_{i}", disabled=i==0):
            st.session_state.sections[i], st.session_state.sections[i-1] = st.session_state.sections[i-1], st.session_state.sections[i]
            st.rerun()
    with col3:
        if st.button("↓", key=f"down_{i}", disabled=i==num_sections-1):
            st.session_state.sections[i], st.session_state.sections[i+1] = st.session_state.sections[i+1], st.session_state.sections[i]
            st.rerun()
    
    col1, col2 = st.columns([3, 1])
    with col1:
        st.session_state.sections[i]["title"] = st.text_input(f"Title for Section {i+1}", value=st.session_state.sections[i]["title"], key=f"title_{i}")
        st.session_state.sections[i]["content"] = st.text_area(f"Content for Section {i+1} (each line = bullet point)", value=st.session_state.sections[i]["content"], key=f"content_{i}", height=120)
    with col2:
        section_color_name = st.selectbox("Background Color", list(COLOR_OPTIONS.keys()), 
                                        index=list(COLOR_OPTIONS.values()).index(st.session_state.sections[i]["color"]), 
                                        key=f"section_color_{i}")
        st.session_state.sections[i]["color"] = COLOR_OPTIONS[section_color_name]

sections = st.session_state.sections

contact_info = st.text_area("Contact Information (supports links & emails)", height=100)

if st.button("Generate Newsletter"):
    pdf_buffer = create_pdf(top_image, intro_text, intro_color, sections, contact_info, bottom_image, month, year, issue_number)
    docx_buffer = create_docx(top_image, intro_text, intro_color, sections, contact_info, bottom_image, month, year, issue_number)

    st.download_button("📥 Download Newsletter PDF", data=pdf_buffer, file_name="newsletter.pdf", mime="application/pdf")
    st.download_button("📥 Download Newsletter DOCX", data=docx_buffer, file_name="newsletter.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
