import streamlit as st
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, ListFlowable, ListItem, KeepTogether, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib.colors import lightgrey, Color
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re

# --------- Helper functions ---------

# Custom canvas with watermark, header, footer, and banner (with bottom image)
class WatermarkCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        self.month = kwargs.pop('month', '')
        self.year = kwargs.pop('year', '')
        self.bottom_image = kwargs.pop('bottom_image', None)
        self.total_pages = 0
        canvas.Canvas.__init__(self, *args, **kwargs)
        
    def showPage(self):
        self.draw_watermark()
        self.draw_month_year_on_image()
        self.draw_bottom_image()
        self.draw_banner()
        canvas.Canvas.showPage(self)
        
    def save(self):
        self.total_pages = self._pageNumber
        canvas.Canvas.save(self)
        
    def draw_watermark(self):
        self.saveState()
        self.setFillColor(lightgrey)
        self.setFont("Helvetica", 50)
        self.rotate(45)
        self.drawCentredText(400, 100, "NEWSLETTER")
        self.restoreState()
    
    def draw_month_year_on_image(self):
        if self.month and self.year and self._pageNumber == 1:
            self.saveState()
            self.setFont("Helvetica-Bold", 14)
            self.setFillColor(Color(0, 0, 0))  # Black text for visibility
            month_year_text = f"{self.month} {self.year}"
            text_width = self.stringWidth(month_year_text, "Helvetica-Bold", 14)
            # Position on top right of content area
            x = 0.75*inch + CONTENT_WIDTH - text_width - 20
            y = A4[1] - 1.2*inch  # Below header, on top image area
            self.drawString(x, y, month_year_text)
            self.restoreState()
    
    def draw_bottom_image(self):
        if self.bottom_image is not None:
            try:
                self.saveState()
                self.bottom_image.seek(0)
                img = ImageReader(self.bottom_image)
                
                # Calculate scaling to fit page width
                max_width = A4[0] - 1.5*inch
                max_height = 1*inch
                
                img_width, img_height = img.getSize()
                width_ratio = max_width / img_width
                height_ratio = max_height / img_height
                scale_ratio = min(width_ratio, height_ratio)
                
                final_width = img_width * scale_ratio
                final_height = img_height * scale_ratio
                
                # Center horizontally, place at very bottom
                x = (A4[0] - final_width) / 2
                y = 0.1*inch
                
                self.drawImage(img, x, y, width=final_width, height=final_height, mask='auto')
                self.restoreState()
            except Exception:
                pass
    
    def draw_banner(self):
        self.saveState()
        banner_start_y = 1.2*inch
        banner_height = 1.5*inch
        
        # Dark blue banner (adjusted to not overlap image)
        self.setFillColor(Color(0.1, 0.2, 0.4))  # Dark blue
        self.rect(0, banner_start_y, A4[0], banner_height, fill=1)
        
        # AWS logo and text (positioned in banner area)
        text_y = banner_start_y + banner_height - 0.6*inch
        
        # AWS logo placeholder (white text)
        self.setFillColor(Color(1, 1, 1))  # White
        self.setFont("Helvetica-Bold", 16)
        self.drawString(0.75*inch, text_y, "AWS")
        
        # "Stay tuned for more updates" text
        self.setFont("Helvetica-Bold", 14)
        self.setFillColor(Color(0.8, 0.2, 0.6))  # Dark pink
        self.drawString(2*inch, text_y, "St")
        self.setFillColor(Color(1, 1, 1))  # White
        st_width = self.stringWidth("St", "Helvetica-Bold", 14)
        self.drawString(2*inch + st_width, text_y, "ay tuned for more updates")
        
        self.restoreState()

def process_content_pdf(content):
    custom_link_pattern = re.compile(r'\[([^\]]+)\]\(([^)]+)\)')
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

def text_to_bullets(text):
    if not text:
        return []
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    return lines

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

def add_text_with_links(paragraph, text):
    custom_link_pattern = re.compile(r'\[([^\]]+)\]\(([^)]+)\)')
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

COLOR_OPTIONS = {
    "Cream White": Color(0.98, 0.96, 0.92),
    "Off White": Color(0.96, 0.94, 0.90),
    "Light Blue": Color(0.90, 0.95, 1.0),
    "Light Green": Color(0.90, 0.98, 0.90),
    "Light Pink": Color(0.98, 0.90, 0.95),
    "Light Yellow": Color(0.98, 0.98, 0.85)
}

# --------- PDF Generation ---------
def create_pdf(top_image, intro_text, intro_color, sections, contact_info, bottom_image, month="", year=""):
    buffer = BytesIO()
    PAGE_WIDTH, PAGE_HEIGHT = A4
    LEFT_MARGIN = RIGHT_MARGIN = 0.75 * inch
    TOP_MARGIN = BOTTOM_MARGIN = 0.75 * inch
    CONTENT_WIDTH = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN

    class CustomCanvas(WatermarkCanvas):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, month=month, year=year, bottom_image=bottom_image, **kwargs)

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=LEFT_MARGIN,
        rightMargin=RIGHT_MARGIN,
        topMargin=0.75*inch,
        bottomMargin=0.75*inch,
        canvasmaker=CustomCanvas
    )

    styles = getSampleStyleSheet()
    intro_style = ParagraphStyle(
        'IntroStyle',
        parent=styles['Normal'],
        backColor=intro_color,
        borderPadding=12
    )

    story = []
    
    # Month/year above top image on right side
    if month and year:
        month_year_para = Paragraph(f"<b>{month} {year}</b>", ParagraphStyle('MonthYear', parent=styles['Normal'], fontSize=14, textColor=Color(0,0,0), alignment=2))
        story.append(month_year_para)
        story.append(Spacer(1, 0.1*inch))
    
    # Top image
    if top_image:
        top_image.seek(0)
        img = Image(top_image)
        img.drawWidth = CONTENT_WIDTH
        calculated_height = img.drawHeight * (CONTENT_WIDTH / img.imageWidth)
        img.drawHeight = min(calculated_height, 2*inch)
        image_table = Table([[img]], colWidths=[CONTENT_WIDTH])
        image_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]))
        story.append(image_table)
    
    # Introduction
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
    
    # Contact Information
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
    
    # Add bottom image at the end
    if bottom_image:
        story.append(Spacer(1, 0.5*inch))
        bottom_image.seek(0)
        img = Image(bottom_image)
        img.drawWidth = CONTENT_WIDTH
        calculated_height = img.drawHeight * (CONTENT_WIDTH / img.imageWidth)
        img.drawHeight = min(calculated_height, 1*inch)
        image_table = Table([[img]], colWidths=[CONTENT_WIDTH])
        image_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]))
        story.append(image_table)
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# --------- DOCX Generation ---------
def create_docx(top_image, intro_text, intro_color, sections, contact_info, bottom_image, month="", year=""):
    doc = Document()
    
    # Header with What's new at Amazon Web Services and month/year
    header_table = doc.add_table(rows=1, cols=2)
    header_table.autofit = False
    header_table.columns[0].width = Inches(4)
    header_table.columns[1].width = Inches(2)
    
    left_cell = header_table.cell(0, 0)
    left_cell.text = "What's new at Amazon Web Services"
    left_cell.paragraphs[0].runs[0].bold = True
    
    right_cell = header_table.cell(0, 1)
    if month and year:
        right_cell.text = f"{month} {year}"
        right_cell.paragraphs[0].runs[0].bold = True
        right_cell.paragraphs[0].alignment = 2  # Right alignment
    
    doc.add_paragraph()
    
    # Top image
    if top_image:
        top_image.seek(0)
        doc.add_picture(top_image, width=Inches(6), height=Inches(2))
    
    # Introduction
    if intro_text:
        p = doc.add_paragraph()
        add_text_with_links(p, intro_text)
        doc.add_paragraph()
    
    # Sections
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
    
    # Contact Information
    if contact_info:
        doc.add_heading("Contact Information", level=2)
        p = doc.add_paragraph()
        add_text_with_links(p, contact_info)
    
    # Banner text (before footer image)
    doc.add_paragraph()
    banner_p = doc.add_paragraph()
    banner_p.add_run("AWS ").bold = True
    banner_p.add_run("Stay tuned for more updates").bold = True
    banner_p.alignment = 1  # Center alignment
    
    # Bottom image in DOCX footer (at very bottom)
    if bottom_image:
        bottom_image.seek(0)
        section = doc.sections[-1]
        footer = section.footer
        
        # Add image to footer
        para = footer.add_paragraph()
        para.alignment = 1  # Center alignment
        run = para.add_run()
        run.add_picture(bottom_image, width=Inches(6))
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --------- Streamlit UI ---------
st.title("📬 Newsletter Generator (Aligned PDF & DOCX with Clickable Links)")
st.write("Fill in the details to create a perfectly aligned newsletter with clickable links, bullet points, intro, contact info, and images.")

# Newsletter header information
col1, col2 = st.columns(2)
with col1:
    month = st.text_input("Month", placeholder="e.g., January")
with col2:
    year = st.text_input("Year", placeholder="e.g., 2024")

# Image uploads
top_image = st.file_uploader("Upload Top Image", type=["png", "jpg", "jpeg"])
bottom_image = st.file_uploader("Upload Bottom Image", type=["png", "jpg", "jpeg"])

# Introduction section
col1, col2 = st.columns([3, 1])
with col1:
    intro_text = st.text_area("Newsletter Introduction", height=120)
with col2:
    intro_color_name = st.selectbox("Background Color", list(COLOR_OPTIONS.keys()), key="intro_color")
    intro_color = COLOR_OPTIONS[intro_color_name]

# Initialize sections
if 'sections' not in st.session_state:
    st.session_state.sections = [{"title": "", "content": "", "color": COLOR_OPTIONS["Off White"]}]

# Number of sections
num_sections = st.number_input("Number of Sections", min_value=1, max_value=10, value=len(st.session_state.sections))

# Adjust sections list
if len(st.session_state.sections) != num_sections:
    if num_sections > len(st.session_state.sections):
        st.session_state.sections.extend([{"title": "", "content": "", "color": COLOR_OPTIONS["Off White"]}] * (num_sections - len(st.session_state.sections)))
    else:
        st.session_state.sections = st.session_state.sections[:num_sections]

# Section inputs
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

# Contact information
contact_info = st.text_area("Contact Information (supports links & emails)", height=100)

# Generate button
if st.button("Generate Newsletter"):
    with st.spinner("Generating newsletter..."):
        try:
            pdf_buffer = create_pdf(top_image, intro_text, intro_color, sections, contact_info, bottom_image, month, year)
            docx_buffer = create_docx(top_image, intro_text, intro_color, sections, contact_info, bottom_image, month, year)
            
            st.success("Newsletter generated successfully!")
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button("📥 Download Newsletter PDF", data=pdf_buffer, file_name="newsletter.pdf", mime="application/pdf")
            with col2:
                st.download_button("📥 Download Newsletter DOCX", data=docx_buffer, file_name="newsletter.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        
        except Exception as e:
            st.error(f"Error generating newsletter: {str(e)}")
