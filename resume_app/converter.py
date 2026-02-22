try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
except Exception:
    Document = None
    Pt = None
    Inches = None
    RGBColor = None
    WD_ALIGN_PARAGRAPH = None
    WD_TAB_ALIGNMENT = None
#!/usr/bin/env python3
"""
generator.py
Convert resume.txt into a beautiful HTML (and optional PDF) resume.
Usage:
    python converter.py resume.txt
"""

import sys, re
import os
from pathlib import Path
from xml.sax.saxutils import escape
try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
except Exception:
    pythoncom = None
    win32com = None
PDF_EXPORT_PATHS = {
    'deshraj sharma': Path('D:/Deshraj Sir'),
    'rohit tawade': Path('D:/Resumes'),
    'mahesh sandur': Path('D:/Mahesh Sandur'),
    'mahesh pawar': Path('D:/Mahesh Pawar'),
    'paresh kumar': Path('D:/Paresh Kumar'),
    'ashish kumar': Path('D:/Ashish Kumar'),
}

try:
    from jinja2 import Template
except Exception:
    Template = None

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem
except Exception:
    A4 = None
    getSampleStyleSheet = None
    ParagraphStyle = None
    TA_CENTER = None
    TA_LEFT = None
    SimpleDocTemplate = None
    Paragraph = None
    Spacer = None
    ListFlowable = None
    ListItem = None

try:
    from fpdf import FPDF
except Exception:
    FPDF = None

TEMPLATE_FILE = Path(__file__).parent / 'template.jinja'


def _clean_filename_token(token):
    token = (token or '').replace('_', ' ')
    token = re.sub(r"\s+", " ", token)
    return token.strip()


def extract_filename_metadata(resume_path):
    stem = Path(resume_path).stem
    parts = [p for p in stem.split('-') if p.strip()]
    metadata = {
        'first': '',
        'last': '',
        'role': '',
        'company': ''
    }
    if not parts:
        return metadata
    metadata['first'] = _clean_filename_token(parts[0])
    if len(parts) >= 2:
        metadata['last'] = _clean_filename_token(parts[1])
    if len(parts) >= 3:
        if len(parts) >= 4:
            role_parts = [_clean_filename_token(p) for p in parts[2:-1]]
            metadata['role'] = " ".join([p for p in role_parts if p])
            metadata['company'] = _clean_filename_token(parts[-1])
        else:
            role_parts = [_clean_filename_token(p) for p in parts[2:]]
            metadata['role'] = " ".join([p for p in role_parts if p])
    return metadata


def _split_name(name):
    parts = [p for p in (name or '').replace('\n', ' ').split() if p]
    if not parts:
        return '', ''
    if len(parts) == 1:
        return parts[0], ''
    return parts[0], parts[-1]

def read_file(path):
    return Path(path).read_text(encoding='utf-8', errors='ignore').strip().splitlines()

def parse_resume(lines):
    data = {
        'name': '',
        'contact': [],
        'title': '',
        'summary': '',
        'experience': [],
        'skills': [],
        'skills_raw': '',
        'certifications': [],
        'education': [],
        'company': ''
    }
    i = 0
    # Find first non-empty line as name
    while i < len(lines) and not lines[i].strip():
        i += 1
    if i < len(lines):
        # If the next few lines are single characters, join them
        name_line = lines[i].strip()
        # Check if the next lines are single characters (vertical name)
        name_chars = [name_line]
        j = i + 1
        while j < len(lines) and len(lines[j].strip()) == 1:
            name_chars.append(lines[j].strip())
            j += 1
        if len(name_chars) > 1:
            data['name'] = ''.join(name_chars)
            i = j
        else:
            data['name'] = name_line
            i += 1
    # Collect contact until blank
    while i < len(lines) and lines[i].strip():
        contact_line = lines[i].strip()
        lowered = contact_line.lower()
        if lowered.startswith('company name:'):
            _, _, value = contact_line.partition(':')
            if value:
                data['company'] = value.strip()
        data['contact'].append(contact_line)
        i += 1
    # Insert a single blank line after any 'Nationality:' contact entry
    if data['contact']:
        new_contacts = []
        for idx, c in enumerate(data['contact']):
            new_contacts.append(c)
            try:
                is_nationality = isinstance(c, str) and c.strip().lower().startswith('nationality:')
            except Exception:
                is_nationality = False
            if is_nationality:
                # Only add a single blank if the next item is not already blank
                next_is_blank = (idx + 1 < len(data['contact']) and data['contact'][idx+1].strip() == '')
                if not next_is_blank:
                    new_contacts.append('')
        data['contact'] = new_contacts
    # After contacts, attempt to capture a standalone title line before sections
    KNOWN_SECTION_KEYS = {
        "professional summary",
        "summary",
        "professional experience",
        "experience",
        "technical skills",
        "skills",
        "certifications",
        "certification",
        "education",
        "qualification",
    }
    # Skip blank lines between contact info and potential title
    while i < len(lines) and not lines[i].strip():
        i += 1
    if i < len(lines):
        possible_title = lines[i].strip()
        key = possible_title.lower().rstrip(":")
        if key not in KNOWN_SECTION_KEYS and not key.startswith("certificat"):
            data['title'] = possible_title
            i += 1
            while i < len(lines) and not lines[i].strip():
                i += 1
    # Parse sections
    section, buffer = None, []
    def flush_buffer(sec, buf):
        text = "\n".join(buf).strip()
        if not text: 
            return
        if sec == "professional summary":
            data['summary'] = text
        elif sec == "professional experience":
            parts = [p.strip() for p in text.split("\n\n") if p.strip()]
            for p in parts:
                lines_p = [ln for ln in p.splitlines() if ln.strip()]
                if not lines_p:
                    continue
                header = lines_p[0]
                duration = None
                bullets = []
                idx = 1
                if len(lines_p) > 1 and re.search(r"\d{4}", lines_p[1]):
                    duration = lines_p[1]
                    idx = 2
                # All remaining lines: treat as bullets, regardless of starting symbol
                for ln in lines_p[idx:]:
                    bullet = ln.strip()
                    # Remove leading bullet symbols if present
                    while bullet.startswith('*') or bullet.startswith('•') or bullet.startswith('-'):
                        bullet = bullet[1:].strip()
                    if bullet:
                        bullets.append(bullet)
                # If header itself starts with a bullet symbol, treat it as a bullet, not as header
                header_clean = header.strip()
                if header_clean.startswith('*') or header_clean.startswith('•') or header_clean.startswith('-'):
                    # Move header to bullets, set header to empty
                    while header_clean.startswith('*') or header_clean.startswith('•') or header_clean.startswith('-'):
                        header_clean = header_clean[1:].strip()
                    if header_clean:
                        bullets = [header_clean] + bullets
                    header = ''
                data['experience'].append({
                    "header": header,
                    "duration": duration,
                    "bullets": bullets
                })
        elif sec == "technical skills":
            # Remove leading bullet symbols from each line
            lines = [l.strip() for l in text.splitlines() if l.strip()]
            cleaned_lines = []
            for l in lines:
                while l.startswith('*') or l.startswith('-') or l.startswith('•'):
                    l = l[1:].strip()
                cleaned_lines.append(l)
            data['skills_raw'] = '\n'.join(cleaned_lines)
            # Split on commas, semicolons, dashes, or newlines for legacy rendering
            raw_skills = re.split(r"[-,\n;]+", data['skills_raw'])
            cleaned = []
            for s in raw_skills:
                s = s.strip()
                if not s: 
                    continue
                # Add space between joined words like PythonSQL -> Python SQL
                s = re.sub(r'([a-z])([A-Z])', r'\1 \2', s)
                cleaned.append(s)
            data['skills'] = cleaned
        elif sec and sec.startswith("certificat"):
            # Remove leading symbols from each line
            certs = []
            for l in text.splitlines():
                l = l.strip()
                while l.startswith('*') or l.startswith('-') or l.startswith('•'):
                    l = l[1:].strip()
                if l:
                    certs.append(l)
            data['certifications'] = certs
        elif sec == "education":
            # Remove leading symbols from each line
            edu = []
            for l in text.splitlines():
                l = l.strip()
                while l.startswith('*') or l.startswith('-') or l.startswith('•'):
                    l = l[1:].strip()
                if l:
                    edu.append(l)
            data['education'] = edu

    while i < len(lines):
        ln = lines[i].strip()
        key = ln.lower().rstrip(":")
        if key in ("professional summary","summary"):
            flush_buffer(section, buffer); buffer=[]; section="professional summary"
        elif key in ("professional experience","experience"):
            flush_buffer(section, buffer); buffer=[]; section="professional experience"
        elif key in ("technical skills","skills"):
            flush_buffer(section, buffer); buffer=[]; section="technical skills"
        elif key.startswith("certificat"):
            flush_buffer(section, buffer); buffer=[]; section="certifications"
        elif key in ("education","qualification"):
            flush_buffer(section, buffer); buffer=[]; section="education"
        else:
            buffer.append(ln)
        i+=1
    flush_buffer(section, buffer)
    return data

def render_html(data, out_html):
    if Template is None:
        raise RuntimeError("jinja2 is not installed; HTML rendering is unavailable.")
    if not TEMPLATE_FILE.exists():
        raise RuntimeError(f"Template file not found: {TEMPLATE_FILE}")
    tpl = Template(TEMPLATE_FILE.read_text(encoding="utf-8"))
    html = tpl.render(data=data)
    out_html.write_text(html, encoding="utf-8")
    print(f"✅ Wrote HTML: {out_html}")
# PDF generation removed per user request


def sanitize_component(text, replace_space_with_hyphen=False):
    text = (text or '').strip()
    if not text:
        return ''
    text = re.sub(r"\s+", " ", text)
    safe_chars = []
    for ch in text:
        if ch in '<>:"/\\|?*':
            safe_chars.append(' ')
        else:
            safe_chars.append(ch)
    cleaned = ''.join(safe_chars)
    cleaned = re.sub(r"\s+", "-" if replace_space_with_hyphen else " ", cleaned).strip()
    cleaned = re.sub(r"-+", "-", cleaned)
    cleaned = cleaned.strip('-. ')
    return cleaned


def clean_bullet_text(text):
    if text is None:
        return ''
    stripped = str(text).strip()
    while stripped.startswith(('*', '-', '•')):
        stripped = stripped[1:].strip()
    return stripped


def build_contact_items(contact_lines):
    entries = []
    seen = set()
    for raw in contact_lines:
        line = str(raw).strip()
        if not line or line.lower().startswith('company name:'):
            continue
        line = clean_bullet_text(line)
        if ':' in line:
            label, value = line.split(':', 1)
            label = label.strip().title()
            value = value.strip()
            if not value:
                continue
            label_key = label.lower()
            if label_key == 'email':
                value = re.sub(r"mailto:\s*", '', value, flags=re.IGNORECASE)
                value = value.replace('<', '').replace('>', '').replace('[', '').replace(']', '')
                found_emails = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+", value)
                if found_emails:
                    unique_emails = []
                    for email in found_emails:
                        if email not in unique_emails:
                            unique_emails.append(email)
                    value = " / ".join(unique_emails)
            elif label_key == 'address':
                value = value.rstrip(', ')
                if value and not value.lower().endswith('india'):
                    if not value.endswith('.'):
                        value = f"{value}, India"
                    else:
                        value = f"{value} India"
            entry = {'label': label, 'label_key': label_key, 'value': value}
        else:
            entry = {'label': '', 'label_key': '', 'value': line}
        dedupe_key = (entry['label_key'], entry['value'].lower())
        if entry['value'] and dedupe_key not in seen:
            entries.append(entry)
            seen.add(dedupe_key)
    return entries


def _pdf_safe_text(text):
    text = str(text or "")
    replacements = {
        "\u2013": "-",
        "\u2014": "-",
        "\u2018": "'",
        "\u2019": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2022": "-",
        "\u00a0": " ",
    }
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    return text.encode("latin-1", errors="replace").decode("latin-1")


def _render_pdf_fpdf(data, pdf_path, file_metadata=None):
    if FPDF is None:
        raise RuntimeError("fpdf2 is not installed; FPDF PDF generation is unavailable.")

    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    file_metadata = file_metadata or {}

    pdf = FPDF(format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_margins(12, 12, 12)
    pdf.add_page()

    def write_line(text, size=10.5, style="", align="L", ln=True):
        pdf.set_font("Helvetica", style=style, size=size)
        pdf.cell(0, 6, _pdf_safe_text(text), ln=1 if ln else 0, align=align)

    def write_para(text, size=10.5, style=""):
        pdf.set_font("Helvetica", style=style, size=size)
        pdf.multi_cell(0, 6, _pdf_safe_text(text))

    candidate_name = (data.get("name") or "").strip() or "Candidate Name"
    pdf.set_font("Helvetica", style="B", size=20)
    pdf.cell(0, 10, _pdf_safe_text(candidate_name), ln=1, align="C")

    effective_title = (data.get("title") or "").strip()
    metadata_role = (file_metadata.get("role") or "").strip()
    if metadata_role and effective_title and metadata_role.lower() == effective_title.lower():
        metadata_role = ""
    headline_bits = [bit for bit in (effective_title, metadata_role) if bit]
    if headline_bits:
        pdf.set_font("Helvetica", style="B", size=11)
        pdf.cell(0, 7, _pdf_safe_text(" | ".join(dict.fromkeys(headline_bits))), ln=1, align="C")
    pdf.ln(1)

    contact_items = build_contact_items(data.get("contact", []))
    if contact_items:
        label_titles = {
            "phone": "Phone",
            "email": "Email",
            "address": "Address",
            "nationality": "Nationality",
        }
        for item in contact_items:
            label_key = item.get("label_key", "")
            value = (item.get("value") or "").strip()
            if not value:
                continue
            if label_key in label_titles:
                write_para(f"{label_titles[label_key]}: {value}", size=10.5, style="")
            else:
                write_para(value, size=10.5, style="")
        pdf.ln(2)

    summary_text = clean_bullet_text(data.get("summary", ""))
    if summary_text:
        write_line("PROFESSIONAL SUMMARY", size=11, style="B")
        write_para(summary_text)
        pdf.ln(1)

    experience_items = data.get("experience", [])
    if experience_items:
        write_line("PROFESSIONAL EXPERIENCE", size=11, style="B")
        for role in experience_items:
            header = clean_bullet_text(role.get("header", ""))
            duration = clean_bullet_text(role.get("duration", ""))
            if header and duration:
                write_para(f"{header} | {duration}", size=10.5, style="B")
            elif header:
                write_para(header, size=10.5, style="B")
            elif duration:
                write_para(duration, size=10.5, style="")

            bullets = [clean_bullet_text(item) for item in role.get("bullets", []) if clean_bullet_text(item)]
            for bullet in bullets:
                write_para(f"- {bullet}")
            pdf.ln(1)

    skills_lines = [clean_bullet_text(line) for line in data.get("skills_raw", "").splitlines() if clean_bullet_text(line)]
    if skills_lines:
        write_line("SKILLS", size=11, style="B")
        for entry in skills_lines:
            if ":" in entry:
                write_para(entry)
            else:
                write_para(f"- {entry}")
        pdf.ln(1)

    certifications = [clean_bullet_text(item) for item in data.get("certifications", []) if clean_bullet_text(item)]
    if certifications:
        write_line("CERTIFICATIONS", size=11, style="B")
        for cert in certifications:
            write_para(f"- {cert}")
        pdf.ln(1)

    education_lines = [clean_bullet_text(item) for item in data.get("education", []) if clean_bullet_text(item)]
    if education_lines:
        write_line("EDUCATION", size=11, style="B")
        for edu in education_lines:
            write_para(edu)

    pdf.output(str(pdf_path))
    return pdf_path


def _derive_output_stem(data, file_metadata, fallback_stem):
    split_first, split_last = _split_name(data.get('name', ''))
    component_specs = [
        (split_first or file_metadata.get('first'), True),
        (split_last or file_metadata.get('last'), True),
        (file_metadata.get('role') or (data.get('title') or ''), True),
        ((data.get('company') or file_metadata.get('company')), True)
    ]
    components = []
    for value, hyphenize in component_specs:
        if not value:
            continue
        components.append(sanitize_component(value, replace_space_with_hyphen=hyphenize))
    if components:
        return "-".join(components)
    safe_name = sanitize_component(data.get('name', ''), replace_space_with_hyphen=True)
    safe_title = sanitize_component(data.get('title', ''), replace_space_with_hyphen=True)
    if safe_name and safe_title:
        return f"{safe_name}-{safe_title}"
    if safe_name:
        return safe_name
    if safe_title:
        return safe_title
    return sanitize_component(fallback_stem, replace_space_with_hyphen=True) or fallback_stem


def convert_text_to_resume(resume_text, output_dir=None, source_hint=None):
    """Convert resume text into DOCX/PDF outputs and return resulting file path."""
    if isinstance(resume_text, (list, tuple)):
        lines = [str(line) for line in resume_text]
    else:
        lines = str(resume_text).splitlines()
    data = parse_resume(lines)
    resolved_hint = None
    if source_hint:
        try:
            resolved_hint = Path(source_hint)
        except Exception:
            resolved_hint = None
    file_metadata = extract_filename_metadata(resolved_hint or (source_hint or '')) if (source_hint or resolved_hint) else {
        'first': '',
        'last': '',
        'role': '',
        'company': ''
    }
    fallback_stem = (resolved_hint.stem if resolved_hint else str(source_hint or 'resume')) or 'resume'
    if output_dir is None:
        if resolved_hint and resolved_hint.parent.exists():
            output_dir = resolved_hint.parent
        else:
            output_dir = Path.cwd()
    else:
        output_dir = Path(output_dir)
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        raise RuntimeError(f"Cannot create output directory {output_dir}: {exc}")
    filename_stem = _derive_output_stem(data, file_metadata, fallback_stem)
    docx_path = output_dir / f"{filename_stem}.docx"
    result_path = render_docx(data, docx_path, file_metadata)
    return result_path


def render_pdf(data, pdf_path, file_metadata=None):
    if FPDF is not None:
        return _render_pdf_fpdf(data, pdf_path, file_metadata)
    if SimpleDocTemplate is None or ParagraphStyle is None:
        raise RuntimeError("reportlab is not installed; PDF generation is unavailable.")

    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    file_metadata = file_metadata or {}
    styles = getSampleStyleSheet()
    name_style = ParagraphStyle(
        name='ResumeName',
        parent=styles['Title'],
        fontName='Helvetica-Bold',
        fontSize=24,
        leading=28,
        alignment=TA_CENTER,
        spaceAfter=6,
    )
    headline_style = ParagraphStyle(
        name='ResumeHeadline',
        parent=styles['Heading3'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=14,
        alignment=TA_CENTER,
        spaceAfter=8,
    )
    heading_style = ParagraphStyle(
        name='ResumeSection',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=14,
        alignment=TA_LEFT,
        spaceBefore=10,
        spaceAfter=4,
    )
    body_style = ParagraphStyle(
        name='ResumeBody',
        parent=styles['BodyText'],
        fontName='Helvetica',
        fontSize=10.5,
        leading=14,
        alignment=TA_LEFT,
        spaceAfter=3,
    )
    meta_style = ParagraphStyle(
        name='ResumeMeta',
        parent=body_style,
        fontName='Helvetica-Bold',
        fontSize=10.5,
        leading=14,
        spaceAfter=2,
    )

    story = []
    candidate_name = (data.get('name') or '').strip() or 'Candidate Name'
    story.append(Paragraph(escape(candidate_name), name_style))

    effective_title = (data.get('title') or '').strip()
    metadata_role = (file_metadata.get('role') or '').strip()
    if metadata_role and effective_title and metadata_role.lower() == effective_title.lower():
        metadata_role = ''
    headline_bits = [bit for bit in (effective_title, metadata_role) if bit]
    if headline_bits:
        tagline = "  |  ".join(dict.fromkeys(headline_bits))
        story.append(Paragraph(escape(tagline), headline_style))

    contact_items = build_contact_items(data.get('contact', []))
    if contact_items:
        label_titles = {
            'phone': 'Phone',
            'email': 'Email',
            'address': 'Address',
            'nationality': 'Nationality',
        }
        for item in contact_items:
            label_key = item.get('label_key', '')
            display_value = escape(item.get('value', ''))
            if not display_value:
                continue
            if label_key in label_titles:
                label = escape(label_titles[label_key])
                story.append(Paragraph(f"<b>{label}:</b> {display_value}", body_style))
            else:
                story.append(Paragraph(display_value, body_style))
        story.append(Spacer(1, 6))

    summary_text = clean_bullet_text(data.get('summary', ''))
    if summary_text:
        story.append(Paragraph('PROFESSIONAL SUMMARY', heading_style))
        story.append(Paragraph(escape(summary_text), body_style))

    experience_items = data.get('experience', [])
    if experience_items:
        story.append(Paragraph('PROFESSIONAL EXPERIENCE', heading_style))
        for role in experience_items:
            header = clean_bullet_text(role.get('header', ''))
            duration = clean_bullet_text(role.get('duration', ''))
            if header or duration:
                header_bits = []
                if header:
                    header_bits.append(f"<b>{escape(header)}</b>")
                if duration:
                    header_bits.append(f"<font color='#444444'>{escape(duration)}</font>")
                story.append(Paragraph(" &nbsp; | &nbsp; ".join(header_bits), meta_style))
            bullets = [clean_bullet_text(item) for item in role.get('bullets', []) if clean_bullet_text(item)]
            if bullets:
                bullet_items = [ListItem(Paragraph(escape(item), body_style), leftIndent=10) for item in bullets]
                story.append(
                    ListFlowable(
                        bullet_items,
                        bulletType='bullet',
                        start='circle',
                        leftIndent=10,
                        bulletFontName='Helvetica',
                        bulletFontSize=8,
                    )
                )
                story.append(Spacer(1, 2))

    skills_lines = [clean_bullet_text(line) for line in data.get('skills_raw', '').splitlines() if clean_bullet_text(line)]
    if skills_lines:
        story.append(Paragraph('SKILLS', heading_style))
        for entry in skills_lines:
            if ':' in entry:
                label, value = entry.split(':', 1)
                story.append(Paragraph(f"<b>{escape(label.strip())}:</b> {escape(value.strip())}", body_style))
            else:
                story.append(Paragraph(f"• {escape(entry)}", body_style))

    certifications = [clean_bullet_text(item) for item in data.get('certifications', []) if clean_bullet_text(item)]
    if certifications:
        story.append(Paragraph('CERTIFICATIONS', heading_style))
        cert_items = [ListItem(Paragraph(escape(cert), body_style), leftIndent=10) for cert in certifications]
        story.append(ListFlowable(cert_items, bulletType='bullet', leftIndent=10, bulletFontSize=8))

    education_lines = [clean_bullet_text(item) for item in data.get('education', []) if clean_bullet_text(item)]
    if education_lines:
        story.append(Paragraph('EDUCATION', heading_style))
        for edu in education_lines:
            story.append(Paragraph(escape(edu), body_style))

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36,
    )
    doc.build(story)
    return pdf_path


def convert_text_to_pdf(resume_text, output_pdf_path=None, source_hint=None):
    """Convert resume text into a PDF output and return resulting PDF path."""
    if isinstance(resume_text, (list, tuple)):
        lines = [str(line) for line in resume_text]
    else:
        lines = str(resume_text).splitlines()
    data = parse_resume(lines)

    resolved_hint = None
    if source_hint:
        try:
            resolved_hint = Path(source_hint)
        except Exception:
            resolved_hint = None
    file_metadata = extract_filename_metadata(resolved_hint or (source_hint or '')) if (source_hint or resolved_hint) else {
        'first': '',
        'last': '',
        'role': '',
        'company': ''
    }
    fallback_stem = (resolved_hint.stem if resolved_hint else str(source_hint or 'resume')) or 'resume'
    filename_stem = _derive_output_stem(data, file_metadata, fallback_stem)

    if output_pdf_path is None:
        if resolved_hint and resolved_hint.parent.exists():
            output_dir = resolved_hint.parent
        else:
            output_dir = Path.cwd()
        pdf_path = Path(output_dir) / f"{filename_stem}.pdf"
    else:
        output_pdf_path = Path(output_pdf_path)
        if output_pdf_path.exists() and output_pdf_path.is_dir():
            pdf_path = output_pdf_path / f"{filename_stem}.pdf"
        elif output_pdf_path.suffix.lower() == '.pdf':
            pdf_path = output_pdf_path
        elif output_pdf_path.suffix:
            pdf_path = output_pdf_path.with_suffix('.pdf')
        else:
            pdf_path = output_pdf_path / f"{filename_stem}.pdf"

    try:
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        raise RuntimeError(f"Cannot create output directory {pdf_path.parent}: {exc}")

    return render_pdf(data, pdf_path, file_metadata)


def try_export_docx_to_pdf(docx_path, candidate_name=''):
    if win32com is None:
        return False, "pywin32 not installed"
    docx_path = Path(docx_path).resolve()
    if not docx_path.exists():
        return False, f"DOCX missing at {docx_path}"
    target_dir = PDF_EXPORT_PATHS.get(candidate_name.strip().lower())
    if target_dir:
        target_dir.mkdir(parents=True, exist_ok=True)
        pdf_path = target_dir / docx_path.with_suffix('.pdf').name
    else:
        pdf_path = docx_path.with_suffix('.pdf')
    word = None
    try:
        if pythoncom is not None:
            pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(docx_path), ReadOnly=True)
        try:
            doc.ExportAsFixedFormat(str(pdf_path), 17)
        finally:
            doc.Close(False)
        return True, pdf_path
    except Exception as exc:
        return False, exc
    finally:
        if word is not None:
            word.Quit()
        if pythoncom is not None:
            pythoncom.CoUninitialize()


def render_docx(data, docx_path, file_metadata=None):
    import os
    if Document is None:
        raise RuntimeError("python-docx is not installed; DOCX generation is unavailable.")

    # Remove stale DOCX so Word exports never collide with an open handle.
    try:
        if os.path.exists(docx_path):
            os.remove(docx_path)
    except PermissionError:
        print(f"❌ Permission denied: {docx_path}. Please close the file and try again.")
        return None

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    base_style = doc.styles['Normal']
    base_style.font.name = 'Calibri'
    base_style.font.size = Pt(12)
    base_style.paragraph_format.line_spacing = 1.15
    base_style.paragraph_format.space_after = Pt(4)

    try:
        bullet_style = doc.styles['List Bullet']
        bullet_style.font.name = 'Calibri'
        bullet_style.font.size = Pt(12)
    except KeyError:
        bullet_style = None

    accent = RGBColor(0, 0, 0)
    muted = RGBColor(96, 96, 96)

    def add_section_heading(title: str):
        heading_text = title.strip()
        if not heading_text:
            return None
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(heading_text.upper())
        run.font.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = accent
        return p

    def clean_bullet_text(text) -> str:
        if text is None:
            return ''
        stripped = str(text).strip()
        while stripped.startswith(('*', '-', '•')):
            stripped = stripped[1:].strip()
        return stripped

    def build_contact_items(contact_lines: list[str]) -> list[dict[str, str]]:
        entries: list[dict[str, str]] = []
        seen: set[tuple[str, str]] = set()
        for raw in contact_lines:
            line = raw.strip()
            if not line or line.lower().startswith('company name:'):
                continue
            line = clean_bullet_text(line)
            if ':' in line:
                label, value = line.split(':', 1)
                label = label.strip().title()
                value = value.strip()
                if not value:
                    continue
                label_key = label.lower()
                if label_key == 'email':
                    value = re.sub(r"mailto:\s*", '', value, flags=re.IGNORECASE)
                    value = value.replace('<', '').replace('>', '').replace('[', '').replace(']', '')
                    found_emails = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+", value)
                    if found_emails:
                        unique_emails: list[str] = []
                        for email in found_emails:
                            if email not in unique_emails:
                                unique_emails.append(email)
                        value = " / ".join(unique_emails)
                elif label_key == 'address':
                    value = value.rstrip(', ')
                    if not value.lower().endswith('india'):
                        if value and not value.endswith('.'):
                            value = f"{value}, India"
                        else:
                            value = f"{value} India"
                entry = {'label': label, 'label_key': label_key, 'value': value}
            else:
                entry = {'label': '', 'label_key': '', 'value': line}
            dedupe_key = (entry['label_key'], entry['value'].lower())
            if entry['value'] and dedupe_key not in seen:
                entries.append(entry)
                seen.add(dedupe_key)
        return entries

    candidate_name = (data.get('name') or '').strip() or 'Candidate Name'
    name_para = doc.add_paragraph()
    name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_para.paragraph_format.space_before = Pt(6)
    name_para.paragraph_format.space_after = Pt(2)
    name_run = name_para.add_run(candidate_name)
    name_run.font.bold = True
    name_run.font.size = Pt(28)
    name_run.font.color.rgb = accent

    effective_title = (data.get('title') or '').strip()
    file_metadata = file_metadata or {}
    metadata_role = (file_metadata.get('role') or '').strip()
    if metadata_role and effective_title and metadata_role.lower() == effective_title.lower():
        metadata_role = ''
    headline_bits = [bit for bit in (effective_title, metadata_role) if bit]
    if headline_bits:
        tagline = "  |  ".join(dict.fromkeys(headline_bits))
        tagline_para = doc.add_paragraph(tagline)
        tagline_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tagline_para.paragraph_format.space_after = Pt(6)
        tagline_run = tagline_para.runs[0]
        tagline_run.font.color.rgb = accent
        tagline_run.font.size = Pt(12)
        tagline_run.font.bold = True

    contact_items = build_contact_items(data.get('contact', []))
    if contact_items:
        label_titles = {
            'phone': 'Phone',
            'email': 'Email',
            'address': 'Address',
            'nationality': 'Nationality',
        }
        for item in contact_items:
            label_key = item.get('label_key', '')
            display_value = item.get('value', '')
            if not display_value:
                continue
            paragraph = doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.space_after = Pt(3)
            if label_key in label_titles:
                label_run = paragraph.add_run(f"{label_titles[label_key]}: ")
                label_run.font.bold = True
                label_run.font.color.rgb = accent
                label_run.font.size = Pt(12)
                value_run = paragraph.add_run(display_value)
                value_run.font.size = Pt(12)
                value_run.font.color.rgb = RGBColor(0, 0, 0)
            else:
                text_run = paragraph.add_run(display_value)
                text_run.font.size = Pt(12)
                text_run.font.color.rgb = RGBColor(0, 0, 0)
        spacer_para = doc.add_paragraph()
        spacer_para.paragraph_format.space_after = Pt(12)

    summary_text = clean_bullet_text(data.get('summary', ''))
    if summary_text:
        add_section_heading('Professional Summary')
        summary_para = doc.add_paragraph(summary_text)
        summary_para.paragraph_format.space_after = Pt(8)

    experience_items = data.get('experience', [])
    if experience_items:
        add_section_heading('Professional Experience')
        for role in experience_items:
            header = clean_bullet_text(role.get('header', ''))
            duration = clean_bullet_text(role.get('duration', ''))
            if header or duration:
                header_para = doc.add_paragraph()
                header_para.paragraph_format.space_before = Pt(8)
                header_para.paragraph_format.space_after = Pt(2)
                tab_stops = header_para.paragraph_format.tab_stops
                tab_stops.clear_all()
                tab_stops.add_tab_stop(Inches(6.0), WD_TAB_ALIGNMENT.RIGHT)
                if header:
                    header_run = header_para.add_run(header)
                    header_run.font.bold = True
                if duration:
                    header_para.add_run('\t')
                    duration_run = header_para.add_run(duration)
                    duration_run.font.color.rgb = RGBColor(0, 0, 0)
            bullets = role.get('bullets', [])
            for bullet in bullets:
                bullet_text = clean_bullet_text(bullet)
                if not bullet_text:
                    continue
                bullet_para = doc.add_paragraph(bullet_text, style='List Bullet')
                bullet_para.paragraph_format.space_after = Pt(1)

    skills_lines = [clean_bullet_text(line) for line in data.get('skills_raw', '').splitlines() if clean_bullet_text(line)]
    if skills_lines:
        add_section_heading('Skills')
        for entry in skills_lines:
            if ':' in entry:
                label, value = entry.split(':', 1)
                skill_para = doc.add_paragraph()
                skill_para.paragraph_format.space_after = Pt(1)
                label_run = skill_para.add_run(f"{label.strip()}: ")
                label_run.font.bold = True
                skill_para.add_run(value.strip())
            else:
                doc.add_paragraph(entry, style='List Bullet')

    certifications = [clean_bullet_text(item) for item in data.get('certifications', []) if clean_bullet_text(item)]
    if certifications:
        add_section_heading('Certifications')
        for cert in certifications:
            cert_para = doc.add_paragraph(cert, style='List Bullet')
            cert_para.paragraph_format.space_after = Pt(1)

    education_lines = [clean_bullet_text(item) for item in data.get('education', []) if clean_bullet_text(item)]
    if education_lines:
        add_section_heading('Education')
        for edu in education_lines:
            edu_para = doc.add_paragraph(edu)
            edu_para.paragraph_format.space_after = Pt(2)

    doc.save(docx_path)

    pdf_result = try_export_docx_to_pdf(docx_path, data.get('name', ''))
    generated_pdf = None
    if pdf_result[0]:
        generated_pdf = pdf_result[1]
        print(f"✅ Created PDF via Word: {generated_pdf}")
    else:
        reason = pdf_result[1]
        if isinstance(reason, str) and "pywin32" in reason.lower():
            print("ℹ Skipping PDF export (pywin32 not installed). Install pywin32 for Word-based PDF generation.")
        elif reason not in (None, ''):
            print(f"⚠ Could not create PDF via Word: {reason}")

    if not generated_pdf:
        fallback_pdf_path = Path(docx_path).with_suffix('.pdf')
        try:
            generated_pdf = render_pdf(data, fallback_pdf_path, file_metadata)
            print(f"✅ Created PDF via ReportLab: {generated_pdf}")
        except Exception as exc:
            print(f"⚠ Could not create PDF via ReportLab: {exc}")

    if generated_pdf:
        try:
            os.remove(docx_path)
        except Exception as exc:
            print(f"⚠ Could not remove intermediate DOCX: {exc}")
        if hasattr(os, "startfile"):
            try:
                os.startfile(str(generated_pdf))
            except Exception as exc:
                print(f"⚠ Could not open PDF automatically: {exc}")
        return Path(generated_pdf)

    print(f"✅ Created DOCX: {docx_path}")
    return Path(docx_path)

def main():
    if len(sys.argv)<2:
        print("Usage: python converter.py resume.txt")
        sys.exit(1)
    resume_path = Path(sys.argv[1])
    if not resume_path.exists():
        print("❌ File not found:", resume_path); sys.exit(1)
    resume_text = resume_path.read_text(encoding='utf-8', errors='ignore')
    result_path = convert_text_to_resume(resume_text, resume_path.parent, resume_path)
    if result_path:
        print(f"✅ Output: {result_path}")

if __name__ == "__main__":
    main()
