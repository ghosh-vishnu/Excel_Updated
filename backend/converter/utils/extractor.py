from docx import Document
from datetime import date
import json
import html
import re
import os
import pandas as pd
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table

# ------------------- Helpers -------------------
DASH = "–"  # en-dash for year ranges
EXCEL_CELL_LIMIT = 32767  # Excel max char limit per cell

def split_into_excel_cells(text, limit=EXCEL_CELL_LIMIT):
    if not text:
        return [""]
    return [text[i:i+limit] for i in range(0, len(text), limit)]

HEADER_LINE_RE = re.compile(
    r"""^\s*
        (?:[A-Za-z]\.)?
        (?:\d+(?:\.\d+)*)?
        [\.\)]?\s*
        (?:report\s*title|full\s*title|full\s*report\s*title|title\s*\(long[-\s]*form\))
        [\s:–-]*$
    """, re.I | re.X
)

def remove_emojis(text: str) -> str:
    """Universal emoji remover."""
    emoji_pattern = re.compile(
        "[" 
        "\U0001F600-\U0001F64F"  # emoticons
        "\U0001F300-\U0001F5FF"  # symbols & pictographs
        "\U0001F680-\U0001F6FF"  # transport & map
        "\U0001F700-\U0001F77F"  # alchemical
        "\U0001F780-\U0001F7FF"  # geometric
        "\U0001F800-\U0001F8FF"  # arrows
        "\U0001F900-\U0001F9FF"  # supplemental
        "\U0001FA00-\U0001FAFF"  # chess, symbols
        "\U00002600-\U000026FF"  # misc symbols
        "\U00002700-\U000027BF"  # dingbats
        "\U00002B00-\U00002BFF"  # arrows & symbols
        "\U0001F1E0-\U0001F1FF"  # flags
        "]+", flags=re.UNICODE
    )
    return emoji_pattern.sub(r'', text or "")
# ------------------- Normalization ------------------- 
def _norm(s: str) -> str:
    s = remove_emojis(s or "")
    return re.sub(r"\s+", " ", s.strip())

def _inline_title(text: str) -> str:
    m = re.split(r"[:\-–]", text, maxsplit=1)
    if len(m) > 1:
        right = m[1].strip()
        if right and not HEADER_LINE_RE.match(right):
            return right
    return ""

def _year_range_present(text: str) -> bool:
    return bool(re.search(r"20\d{2}\s*[\-–]\s*20\d{2}", text))

def _ensure_filename_start_and_year(title: str, filename: str) -> str:
    if not title.lower().startswith(filename.lower()):
        title = f"{filename} {title}"
    if not _year_range_present(title):
        title = f"{title} {DASH}2024–2030"
    return _norm(title)

# ------------------- Convert Paragraph to HTML -------------------
def paragraph_to_html(para):
    text = para.text.strip()
    if not text:
        return ""
    if para.style.name.lower().startswith("list"):
        return f"<li>{text}</li>"
    text = remove_emojis(text)
    if para.style.name.startswith("Heading"):
        level = para.style.name.replace("Heading", "").strip()
        level = int(level) if level.isdigit() else 2
        return f"<h{level}>{text}</h{level}>"
    return f"<p>{text}</p>"


def run_to_html(run):
    text = remove_emojis(run.text)
    if not text:
        return ""
    if run.bold and run.italic:
        return f"<b><i>{text}</i></b>"
    elif run.bold:
        return f"<b>{text}</b>"
    elif run.italic:
        return f"<i>{text}</i>"
    return text

from docx.oxml.ns import qn
from docx.text.run import Run

def runs_to_html(runs):
    parts = []
    for run in runs:
        txt = remove_emojis(run.text)
        if not txt and not run._element.xpath(".//w:br"):
            continue

        # check for manual line breaks inside run
        if run._element.xpath(".//w:br"):
            parts.append("<br>")

        if run.bold and run.italic:
            parts.append(f"<b><i>{txt}</i></b>")
        elif run.bold:
            parts.append(f"<b>{txt}</b>")
        elif run.italic:
            parts.append(f"<i>{txt}</i>")
        else:
            parts.append(txt)
    return "".join(parts).strip()



# ------------------- Extract Title -------------------
def extract_title(docx_path: str) -> str:
    doc = Document(docx_path)
    filename = os.path.splitext(os.path.basename(docx_path))[0]
    filename_low = filename.lower()
    blocks = [(p, (p.text or "").strip()) for p in doc.paragraphs if (p.text or "").strip()]

    capture = False
    for _, text in blocks:
        text = remove_emojis(text)
        if capture:
            return _ensure_filename_start_and_year(text, filename)
        if HEADER_LINE_RE.match(text):
            inline = _inline_title(text)
            if inline:
                return _ensure_filename_start_and_year(inline, filename)
            capture = True
            continue

    for table in doc.tables:
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                cell_text = (cell.text or "").strip().lower()
                if not cell_text:
                    continue
                if "report title" in cell_text or "full title" in cell_text or "full report title" in cell_text:
                    if c_idx + 1 < len(row.cells):
                        nxt = row.cells[c_idx+1].text.strip()
                        if nxt:
                            return _ensure_filename_start_and_year(nxt, filename)
                    if r_idx + 1 < len(table.rows):
                        nxt = row.rows[r_idx+1].cells[c_idx].text.strip()
                        if nxt:
                            return _ensure_filename_start_and_year(nxt, filename)

    for _, text in blocks:
        low = text.lower()
        if low.startswith("full report title") or low.startswith("full title"):
            inline = _inline_title(text)
            if inline:
                return _ensure_filename_start_and_year(inline, filename)
        if low.startswith(filename_low) and "forecast" in low:
            return _ensure_filename_start_and_year(text, filename)

    return "Title Not Available"

# ------------------- Extract Description -------------------
# def extract_description(docx_path):
#     doc = Document(docx_path)
#     html_output = []
#     capture, inside_list = False, False  

#     target_headings = [
#         "introduction and strategic context",
#         "market segmentation and forecast scope",
#         "market trends and innovation landscape",
#         "competitive intelligence and benchmarking",
#         "regional landscape and adoption outlook",
#         "end-user dynamics and use case",
#         "recent developments + opportunities & restraints",
#     ]

#     def clean_heading(text):
#         text = remove_emojis(text.strip())
#         text = re.sub(r'^[^\w]+', '', text)  
#         text = re.sub(r'(?i)section\s*\d+[:\-]?\s*', '', text)  
#         text = re.sub(r'^\d+[\.\-\)]\s*', '', text)  
#         text = re.sub(r'\s+', ' ', text)  
#         return text.lower().strip()

#     def run_to_html(run):
#         text = remove_emojis(run.text.strip())
#         if not text:
#             return ""
#         if run.bold and run.italic:
#             return f"<b><i>{text}</i></b>"
#         elif run.bold:
#             return f"<b>{text}</b>"
#         elif run.italic:
#             return f"<i>{text}</i>"
#         return text

#     for para in doc.paragraphs:
#         text = remove_emojis(para.text.strip())
#         if not text:
#             continue

#         cleaned = clean_heading(text)

#         if not capture and any(h in cleaned for h in target_headings):
#             capture = True  

#         if capture and "report summary, faqs, and seo schema" in cleaned:
#             break  

#         if capture:
#             content = runs_to_html(para.runs)

#             if any(h in cleaned for h in target_headings):
#                 html_output.append("<br>")
#                 matched = next(h for h in target_headings if h in cleaned)
#                 html_output.append(f"<h2>{matched.title()}</h2>")
#                 continue

#             if "list" in para.style.name.lower():
#                 if not inside_list:
#                     html_output.append("<ul>")
#                     inside_list = True
#                 html_output.append(f"<li>{content}</li>")
#                 continue
#             else:
#                 if inside_list:
#                     html_output.append("</ul>")
#                     inside_list = False

#             html_output.append(f"<p>{content}</p>")

#     if inside_list:
#         html_output.append("</ul>")
#     return "\n".join(html_output)

# def extract_description(docx_path):
    doc = Document(docx_path)
    html_output = []
    capture, inside_list = False, False  
    used_headings = set()

    target_headings = [
        "introduction and strategic context",
        "market segmentation and forecast scope",
        "market trends and innovation landscape",
        "competitive intelligence and benchmarking",
        "regional landscape and adoption outlook",
        "end-user dynamics and use case",
        "recent developments + opportunities & restraints",
        "recent developments",
        # "opportunities",   # ✅ added back
        "restraints",
        "report coverage table",
        "table of contents"
    ]

    def clean_heading(text):
        text = remove_emojis(text.strip())
        text = re.sub(r'^[^\w]+', '', text)
        text = re.sub(r'(?i)section\s*\d+[:\-]?\s*', '', text)
        text = re.sub(r'^\d+[\.\-\)]\s*', '', text)
        text = re.sub(r'\s+', ' ', text)
        return text.lower().strip()

    def runs_to_html(runs):
        parts = []
        for run in runs:
            txt = remove_emojis(run.text.strip())
            if not txt:
                continue
            if run.bold and run.italic:
                parts.append(f"<b><i>{txt}</i></b>")
            elif run.bold:
                parts.append(f"<b>{txt}</b>")
            elif run.italic:
                parts.append(f"<i>{txt}</i>")
            else:
                parts.append(txt)
        return " ".join(parts).strip()

    # ✅ Iterate paragraphs & tables
    for block in doc.element.body:
        if isinstance(block, CT_P):  
            para = Paragraph(block, doc)
            text = remove_emojis(para.text.strip())
            if not text:
                continue

            cleaned = clean_heading(text)

            # start capturing from first matching heading
            if not capture and any(h in cleaned for h in target_headings):
                capture = True  

            if capture and "report summary, faqs, and seo schema" in cleaned:
                break  

            if capture:
                content = runs_to_html(para.runs)
                matched_heading = next((h for h in target_headings if h in cleaned), None)

                if matched_heading and matched_heading not in used_headings:
                    if inside_list:
                        html_output.append("</ul>")
                        inside_list = False
                    html_output.append("<br>")
                    html_output.append(f"<h2>{matched_heading.title()}</h2>")
                    used_headings.add(matched_heading)
                elif "list" in para.style.name.lower() or re.match(r"^[•\-–]\s+", text):
                    # unordered list
                    if not inside_list:
                        html_output.append("<ul>")
                        inside_list = True
                    cleaned_item = re.sub(r"^[•\-–]\s*", "", content)
                    html_output.append(f"<li>{cleaned_item}</li>")
                else:
                    if inside_list:
                        html_output.append("</ul>")
                        inside_list = False
                    html_output.append(f"<p>{content}</p>")
                    
                    
                    

        elif isinstance(block, CT_Tbl):  
            table = Table(block, doc)
            table_html = ["<table border='1'>"]
            for row in table.rows:
                table_html.append("<tr>")
                for cell in row.cells:
                    cell_text = " ".join(
                        run.text for para in cell.paragraphs for run in para.runs
                    ).strip()
                    table_html.append(f"<td>{cell_text}</td>")
                table_html.append("</tr>")
            table_html.append("</table>")
            html_output.append("\n".join(table_html))

    if inside_list:
        html_output.append("</ul>")

    return "\n".join(html_output)
def is_list_item(para):
    """Check if paragraph is part of a Word list (numbered/bulleted)."""
    pPr = para._p.pPr
    if pPr is not None and pPr.numPr is not None:
        return True

def extract_description(docx_path):
    doc = Document(docx_path)
    html_output = []
    capture, inside_list = False, None
    used_headings = set()

    # ✅ Target headings jaha se data capture start hoga
    target_headings = [
        "introduction and strategic context",
        "market segmentation and forecast scope",
        "market trends and innovation landscape",
        "competitive intelligence and benchmarking",
        "regional landscape and adoption outlook",
        "end-user dynamics and use case",
        "recent developments + opportunities & restraints",
        "recent developments",
        "restraints",
        # "opportunities",
        "report coverage table",
        "table of contents"
    ]

    def clean_heading(text):
        text = remove_emojis(text.strip())
        text = re.sub(r'^[^\w]+', '', text)
        text = re.sub(r'(?i)section\s*\d+[:\-]?\s*', '', text)
        text = re.sub(r'^\d+[\.\-\)]\s*', '', text)
        text = re.sub(r'\s+', ' ', text)
        return text.lower().strip()

    def runs_to_html(runs):
        """Convert Word runs (bold/italic) to inline HTML."""
        parts = []
        for run in runs:
            txt = remove_emojis(run.text.strip())
            if not txt:
                continue
            if run.bold and run.italic:
                parts.append(f"<b><i>{txt}</i></b>")
            elif run.bold:
                parts.append(f"<b>{txt}</b>")
            elif run.italic:
                parts.append(f"<i>{txt}</i>")
            else:
                parts.append(txt)
        return " ".join(parts).strip()

    # ✅ Iterate over all blocks (paragraphs + tables)
    for block in doc.element.body:
        if isinstance(block, CT_P):
            para = Paragraph(block, doc)
            text = remove_emojis(para.text.strip())
            if not text:
                continue

            cleaned = clean_heading(text)

            # Start capture
            if not capture and any(h in cleaned for h in target_headings):
                capture = True

            # End capture
            if capture and "report summary, faqs, and seo schema" in cleaned:
                break

            if capture:
                content = runs_to_html(para.runs)
                matched_heading = next((h for h in target_headings if h in cleaned), None)

                # ✅ Heading
                if matched_heading and matched_heading not in used_headings:
                    if inside_list:
                        html_output.append(f"</{inside_list}>")
                        inside_list = None
                    html_output.append("<br>")
                    html_output.append(f"<h2>{matched_heading.title()}</h2>")
                    used_headings.add(matched_heading)

                # ✅ Proper list detection
                elif is_list_item(para):
                    if inside_list != "ol":
                        if inside_list:
                            html_output.append(f"</{inside_list}>")
                        html_output.append("<ol>")
                        inside_list = "ol"
                    html_output.append(f"<li>{content}</li>")

                # ✅ Normal paragraph
                else:
                    if inside_list:
                        html_output.append(f"</{inside_list}>")
                        inside_list = None
                    html_output.append(f"<p>{content}</p>")

        elif isinstance(block, CT_Tbl):
            # ✅ Table handling
            table = Table(block, doc)
            table_html = ["<table border='1'>"]
            for row in table.rows:
                table_html.append("<tr>")
                for cell in row.cells:
                    cell_text = " ".join(
                        run.text for para in cell.paragraphs for run in para.runs
                    ).strip()
                    table_html.append(f"<td>{cell_text}</td>")
                table_html.append("</tr>")
            table_html.append("</table>")
            html_output.append("\n".join(table_html))

    if inside_list:
        html_output.append(f"</{inside_list}>")

    return "\n".join(html_output)





# ------------------- TOC Extraction -------------------
# def extract_toc(docx_path):
#     doc = Document(docx_path)
#     html_output, inside_list, capture = [], False, False
#     end_reached = False

#     for para in doc.paragraphs:
#         text = remove_emojis(para.text.strip())
#         low = text.lower()

#         if not capture and "table of contents" in low:
#             capture = True
#             continue

#         if capture:
#             if "list of figures" in low:
#                 html_part = paragraph_to_html(para)
#                 if html_part:
#                     html_output.append(html_part)  
#                 end_reached = True
#                 continue  

#             if end_reached:
#                 style = getattr(para.style, "name","").lower()
#                 if "heading" in style or re.match(r"^\d+[\.\)]\s", text):
#                     break  

#             html_part = paragraph_to_html(para)
#             if html_part:
#                 if html_part.startswith("<li>"):
#                     if not inside_list:
#                         html_output.append("<ul>")
#                         inside_list = True
#                     html_output.append(html_part)
#                 else:
#                     if inside_list:
#                         html_output.append("</ul>")
#                         inside_list = False
#                     html_output.append(html_part)

#     if inside_list:
#         html_output.append("</ul>")
#     return "".join(html_output).strip()
# -----------------------------------------------------TOC Method 2-------------------------------
def extract_toc(docx_path):
    doc = Document(docx_path)
    html_output, inside_list, capture = [], False, False
    end_reached = False

    for para in doc.paragraphs:
        text = para.text.strip()
        low = text.lower()

        # Start condition
        if not capture and "table of contents" in low:
            capture = True
            continue

        if capture:
            # End condition = capture "List of Figures" + its items, then stop
            if "list of figures" in low:
                html_part = paragraph_to_html(para)
                if html_part:
                    html_output.append(html_part)   # add heading "List of Figures"
                end_reached = True
                continue  # don't break yet, because its children may follow

            if end_reached:
                style = getattr(para.style, "name","").lower()
                if "heading" in style or re.match(r"^\d+[\.\)]\s", text):
                    break  

            html_part = paragraph_to_html(para)
            if html_part:
                if html_part.startswith("<li>"):
                    if not inside_list:
                        html_output.append("<ul>")
                        inside_list = True
                    html_output.append(html_part)
                else:
                    if inside_list:
                        html_output.append("</ul>")
                        inside_list = False
                    html_output.append(html_part)

    if inside_list:
        html_output.append("</ul>")
    return remove_emojis("".join(html_output).strip())



# ------------------- FAQ Schema + Methodology -------------------
def _get_text(docx_path):
    doc = Document(docx_path)
    return "\n".join(p.text for p in doc.paragraphs if p.text and p.text.strip())

def _extract_json_block(text, type_name):
    pat = re.compile(r'"@type"\s*:\s*"' + re.escape(type_name) + r'"')
    m = pat.search(text)
    if not m:
        return ""
    start_idx = text.rfind("{", 0, m.start())
    if start_idx == -1:
        return ""
    depth, i, n = 0, start_idx, len(text)
    block_chars = []
    while i < n:
        ch = text[i]
        block_chars.append(ch)
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                break
        i += 1
    return "".join(block_chars).strip()

def extract_faq_schema(docx_path):
    text = _get_text(docx_path)
    return _extract_json_block(text, "FAQPage")

def extract_methodology_from_faqschema(docx_path):
    faq_schema_str = extract_faq_schema(docx_path)  
    if not faq_schema_str:
        return ""   
    try:
        faq_data = json.loads(faq_schema_str)
    except json.JSONDecodeError:
        return ""   
    faqs = []
    q_count = 0
    for item in faq_data.get("mainEntity", []):
        q_count += 1
        question = item.get("name", "").strip()
        answer = item.get("acceptedAnswer", {}).get("text", "").strip()
        if question and answer:
            faqs.append(
                f"<p><strong>Q{q_count}: {html.escape(question)}</strong><br>"
                f"A{q_count}: {html.escape(answer)}</p>"
            )
    return "\n".join(faqs)

# ------------------- Report Coverage -------------------
def extract_report_coverage_table_with_style(docx_path):
    doc = Document(docx_path)
    for table in doc.tables:
        first_row_text = " ".join([c.text.strip().lower() for c in table.rows[0].cells])
        if "report attribute" in first_row_text or "report coverage table" in first_row_text:
            html_parts = []
            html_parts.append('<h2><strong>7.1. Report Coverage Table</strong></h2>')
            html_parts.append('<table cellspacing="0" style="border-collapse:collapse; width:100%"><tbody>')
            for r_idx, row in enumerate(table.rows):
                html_parts.append("<tr>")
                for c_idx, cell in enumerate(row.cells):
                    text = remove_emojis(cell.text.strip())
                    bg = "#deeaf6" if r_idx % 2 == 1 else "#ffffff"
                    if r_idx == 0:
                        bg = "#5b9bd5"
                    td_style = (
                        f"background-color:{bg}; "
                        "border:1px solid #9cc2e5; vertical-align:top; padding:4px;"
                        "width:263px" if c_idx == 0 else
                        f"background-color:{bg}; border:1px solid #9cc2e5; vertical-align:top; padding:4px; width:303px"
                    )
                    html_parts.append(
                        f'<td style="{td_style}"><p><strong>{text}</strong></p></td>'
                        if c_idx == 0 or r_idx == 0 else f'<td style="{td_style}"><p>{text}</p></td>'
                    )
                html_parts.append("</tr>")
            html_parts.append("</tbody></table>")
            return "\n".join(html_parts)
    return ""

# ------------------- Extra Extractors -------------------
def extract_meta_description(docx_path):
    doc = Document(docx_path)
    capture = False
    for para in doc.paragraphs:
        text = para.text.strip()
        low = text.lower()
        if not capture and ("introduction" in low):
            capture = True
            continue
        if capture and text:
            return text
    return ""

def extract_seo_title(docx_path):
    doc = Document(docx_path)
    file_name = os.path.splitext(os.path.basename(docx_path))[0]
    revenue_forecast = ""
    for table in doc.tables:
        headers = [cell.text.strip().lower() for cell in table.rows[0].cells]
        if "report attribute" in headers and "details" in headers:
            attr_idx = headers.index("report attribute")
            details_idx = headers.index("details")
            for row in table.rows[1:]:
                attr = row.cells[attr_idx].text.strip().lower()
                details = row.cells[details_idx].text.strip()
                if "revenue forecast in 2030" in attr:
                    revenue_forecast = details.replace("USD", "$").strip()
                    break
    if revenue_forecast:
        return f"{file_name} Size ({revenue_forecast}) 2030"
    return file_name

def extract_breadcrumb_text(docx_path):
    file_name = os.path.splitext(os.path.basename(docx_path))[0]
    revenue_forecast = ""
    doc = Document(docx_path)
    for table in doc.tables:
        headers = [cell.text.strip().lower() for cell in table.rows[0].cells]
        if "report attribute" in headers and "details" in headers:
            attr_idx = headers.index("report attribute")
            details_idx = headers.index("details")
            for row in table.rows[1:]:
                attr = row.cells[attr_idx].text.strip().lower()
                details = row.cells[details_idx].text.strip()
                if "revenue forecast in 2030" in attr:
                    revenue_forecast = details.replace("USD", "$").strip()
                    break
    if revenue_forecast:
        return f"{file_name} Report 2030"
    return file_name

def extract_sku_code(docx_path):
    return os.path.splitext(os.path.basename(docx_path))[0].lower()

def extract_sku_url(docx_path):
    return os.path.splitext(os.path.basename(docx_path))[0].lower()

def extract_breadcrumb_schema(docx_path):
    text = _get_text(docx_path)
    return _extract_json_block(text, "BreadcrumbList")

# ------------------- Merge -------------------
def merge_description_and_coverage(docx_path):
    try:
        desc_html = extract_description(docx_path) or ""
        coverage_html = extract_report_coverage_table_with_style(docx_path) or ""
        merged_html = desc_html + "\n\n" + coverage_html if (desc_html or coverage_html) else ""
        return merged_html
    except Exception as e:
        return f"ERROR: {e}"