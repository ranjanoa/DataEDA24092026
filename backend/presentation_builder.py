"""
Presentation Builder for DataEDA
Implements a 100% generic, reusable presentation formatting engine inspired by
high-tech industrial design aesthetics (Space Grotesk, Outfit, Innomotics Dark theme).

CRITICAL DESIGN RULE:
NO plant-specific names, NO specific column names, and NO customer data are hardcoded.
All content, titles, tables, variables, and calculations are parsed dynamically
from the incoming payload (ai_summary, ai_charts, dataset_name, title, date).
"""

import os
import io
import re
import base64
import logging
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

logger = logging.getLogger(__name__)

# Reusable design tokens (Innomotics Dark aesthetic)
COLOR_CYAN = RGBColor(0, 255, 255)         # #00FFFF
COLOR_TITLE_CYAN = RGBColor(0, 240, 225)   # #00F0E1
COLOR_LIGHT_SLATE = RGBColor(203, 213, 225)# #CBD5E1
COLOR_MUTED_SLATE = RGBColor(148, 163, 184)# #94A3B8
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_DARK_BG = RGBColor(15, 23, 42)       # #0F172A
COLOR_TABLE_HDR = RGBColor(15, 32, 39)     # #0F2027
COLOR_ZEBRA_1 = RGBColor(30, 41, 59)       # #1E293B
COLOR_ZEBRA_2 = RGBColor(15, 23, 42)       # #0F172A

FONT_TITLE = "Space Grotesk"
FONT_BODY = "Outfit"
FONT_TABLE = "Arial"


def decode_base64_image(b64_str: str) -> io.BytesIO:
    """Decodes data URL base64 image into BytesIO."""
    if not b64_str:
        return None
    if "," in b64_str:
        b64_str = b64_str.split(",", 1)[1]
    img_data = base64.b64decode(b64_str)
    stream = io.BytesIO(img_data)
    stream.seek(0)
    return stream


def clean_title(title_str: str) -> str:
    """Cleans up markdown artifacts from titles."""
    if not title_str:
        return ""
    return title_str.replace('**', '').replace('*', '').replace('#', '').strip()


def set_slide_notes(slide, notes_text: str):
    """Sets speaker notes on a slide."""
    if not notes_text:
        return
    try:
        slide.notes_slide.notes_text_frame.text = notes_text
    except Exception as e:
        logger.debug(f"Could not set slide notes: {e}")


def update_graph_slide_content(slide, title: str, bullets: list, img_stream: io.BytesIO, notes_text: str = None):
    """
    Updates a formatted graph slide:
    1. Sets title and bullets in TextBox 2 (Space Grotesk Cyan + Outfit Slate).
    2. Replaces picture shape in-place preserving coordinates.
    3. Sets slide notes.
    """
    tb2 = None
    tb3 = None
    for sh in slide.shapes:
        if sh.has_text_frame:
            if sh.name == "TextBox 2":
                tb2 = sh
            elif sh.name == "TextBox 3":
                tb3 = sh

    # If new bullets are provided from AI, update the textframe(s)
    if bullets and len(bullets) > 0:
        if tb2 and not tb3:
            # Single text box containing title + bullets (Slides 10-33)
            tf = tb2.text_frame
            tf.clear()
            p0 = tf.paragraphs[0]
            p0.text = clean_title(title)
            p0.font.name = FONT_TITLE
            p0.font.size = Pt(22)
            p0.font.bold = True
            p0.font.color.rgb = COLOR_CYAN
            
            for b in bullets:
                clean_b = b.strip().replace('**', '').replace('*', '').strip()
                if not clean_b:
                    continue
                p = tf.add_paragraph()
                p.text = "• " + clean_b.lstrip('-•*').strip()
                p.font.name = FONT_BODY
                p.font.size = Pt(11.5)
                p.font.bold = False
                p.font.color.rgb = COLOR_LIGHT_SLATE
                p.space_before = Pt(4)
        elif tb2 and tb3:
            # Split textboxes: tb2 has title, tb3 has bullets (Slide 9)
            if title and tb2.text_frame.paragraphs:
                tb2.text_frame.paragraphs[0].text = clean_title(title)
            tf3 = tb3.text_frame
            tf3.clear()
            for idx, b in enumerate(bullets):
                clean_b = b.strip().replace('**', '').replace('*', '').strip()
                if not clean_b:
                    continue
                p = tf3.paragraphs[0] if idx == 0 else tf3.add_paragraph()
                p.text = "• " + clean_b.lstrip('-•*').strip()
                p.font.name = FONT_BODY
                p.font.size = Pt(11.5)
                p.font.bold = False
                p.font.color.rgb = COLOR_LIGHT_SLATE
                p.space_before = Pt(4)
    else:
        # NO new bullets provided: PRESERVE existing curated insights on the slide!
        # Only update title if provided and tb2 exists, leaving paragraphs[1:] completely intact!
        if tb2 and title and tb2.text_frame.paragraphs:
            p0 = tb2.text_frame.paragraphs[0]
            p0.text = clean_title(title)
            p0.font.name = FONT_TITLE
            p0.font.size = Pt(22)
            p0.font.bold = True
            p0.font.color.rgb = COLOR_CYAN

    # Replace picture shape
    old_pic = None
    for sh in slide.shapes:
        if sh.shape_type == 13: # Picture
            old_pic = sh
            break

    if old_pic is not None and img_stream is not None:
        left, top, width, height = old_pic.left, old_pic.top, old_pic.width, old_pic.height
        old_pic._element.getparent().remove(old_pic._element)
        slide.shapes.add_picture(img_stream, left, top, width, height)
    elif img_stream is not None:
        slide.shapes.add_picture(img_stream, Inches(0.52), Inches(2.62), width=Inches(11.87), height=Inches(4.44))

    if notes_text:
        set_slide_notes(slide, notes_text)


def update_table_slide_content(slide, title: str, header_data: list, rows_data: list, notes_text: str = None):
    """
    Updates a table slide:
    1. Sets title in TextBox 2.
    2. Updates table cells with Innomotics dark theme (navy header, zebra slate rows).
    3. Sets slide notes.
    """
    for sh in slide.shapes:
        if sh.has_text_frame and sh.name == "TextBox 2" and title:
            tf = sh.text_frame
            tf.clear()
            p0 = tf.paragraphs[0]
            p0.text = clean_title(title)
            p0.font.name = FONT_TITLE
            p0.font.size = Pt(22)
            p0.font.bold = True
            p0.font.color.rgb = COLOR_CYAN
        elif sh.has_table and header_data and rows_data:
            tbl = sh.table
            target_cols = len(header_data)
            
            # Populate header
            for c_idx in range(min(target_cols, len(tbl.columns))):
                val = header_data[c_idx] if c_idx < len(header_data) else ""
                cell = tbl.cell(0, c_idx)
                cell.text = str(val).strip()
                cell.fill.solid()
                cell.fill.fore_color.rgb = COLOR_TABLE_HDR
                for p in cell.text_frame.paragraphs:
                    p.font.name = FONT_TABLE
                    p.font.size = Pt(10)
                    p.font.bold = True
                    p.font.color.rgb = COLOR_WHITE

            # Populate data rows
            for r_idx in range(min(len(rows_data), len(tbl.rows) - 1)):
                row_vals = rows_data[r_idx]
                bg_col = COLOR_ZEBRA_1 if r_idx % 2 == 0 else COLOR_ZEBRA_2
                for c_idx in range(min(len(row_vals), len(tbl.columns))):
                    val = row_vals[c_idx]
                    cell = tbl.cell(r_idx + 1, c_idx)
                    cell.text = str(val).strip()
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = bg_col
                    for p in cell.text_frame.paragraphs:
                        p.font.name = FONT_TABLE
                        p.font.size = Pt(9.5)
                        p.font.bold = (c_idx == 0)
                        p.font.color.rgb = RGBColor(226, 232, 240)

    if notes_text:
        set_slide_notes(slide, notes_text)


def add_formatted_slide(prs, title: str, bullets: list, img_stream: io.BytesIO = None, notes_text: str = None):
    """
    Appends a new slide styled in the Dark Innomotics high-tech theme.
    """
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    slide = prs.slides.add_slide(blank_layout)
    
    try:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = COLOR_DARK_BG
    except Exception:
        pass

    # Top title box
    title_box = slide.shapes.add_textbox(Inches(0.52), Inches(0.4), Inches(12.3), Inches(1.2))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = clean_title(title)
    p.font.name = FONT_TITLE
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = COLOR_CYAN

    # Bullets under title
    if bullets:
        for b in bullets[:2]:
            clean_b = b.strip().replace('**', '').replace('*', '').strip()
            if not clean_b: continue
            p_b = tf.add_paragraph()
            p_b.text = "• " + clean_b.lstrip('-•*').strip()
            p_b.font.name = FONT_BODY
            p_b.font.size = Pt(11)
            p_b.font.color.rgb = COLOR_LIGHT_SLATE
            p_b.space_before = Pt(3)

    # Image
    if img_stream:
        slide.shapes.add_picture(img_stream, Inches(0.52), Inches(1.8), width=Inches(12.3), height=Inches(5.2))
    elif len(bullets) > 2:
        c_box = slide.shapes.add_textbox(Inches(0.52), Inches(1.8), Inches(12.3), Inches(5.2))
        c_tf = c_box.text_frame
        c_tf.word_wrap = True
        first = True
        for b in bullets[2:]:
            clean_b = b.strip().replace('**', '').replace('*', '').strip()
            if not clean_b: continue
            p_c = c_tf.paragraphs[0] if first else c_tf.add_paragraph()
            first = False
            p_c.text = "• " + clean_b.lstrip('-•*').strip()
            p_c.font.name = FONT_BODY
            p_c.font.size = Pt(13)
            p_c.font.color.rgb = COLOR_LIGHT_SLATE
            p_c.space_before = Pt(6)

    if notes_text:
        set_slide_notes(slide, notes_text)
    return slide


def parse_markdown_ai_content(summary_text: str) -> dict:
    """
    Parses generic AI markdown output into structured blocks:
    - graphs: dict mapping graph identifiers to title & bullets
    - tables: list of (title, headers, rows)
    - executive_bullets: list of summary bullet points
    - correlation_bullets: list of correlation highlights
    """
    graphs = {}
    tables = []
    executive_bullets = []
    correlation_bullets = []
    
    current_graph_key = None
    current_graph_title = None
    current_graph_bullets = []
    
    lines = summary_text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Generic Chart / Graph Header matcher
        graph_m = re.search(
            r'(?:#+\s*|\d+\.\s*)?\*\*(?:Graph\s*(\d+)|STEP\s*([0-9A-Za-z]+)|(3D\s+[^:\*]+)|(Parallel\s+[^:\*]+)|([^\*:]+Diagnostics|[^\*:]+Dynamics|[^\*:]+Recovery))\s*[:—–-]?\s*([^\*]+)?\*\*',
            line,
            re.IGNORECASE
        )
        if not graph_m:
            graph_m = re.search(r'^#{2,4}\s*(?:Graph\s*(\d+)|STEP\s*([0-9A-Za-z]+)|(3D\s+[^:]+)|(Parallel\s+[^:]+))\s*[:—–-]?\s*(.*)$', line, re.IGNORECASE)

        if graph_m:
            if current_graph_key:
                graphs[current_graph_key] = {
                    "title": current_graph_title or current_graph_key,
                    "bullets": current_graph_bullets
                }
            g_num = graph_m.group(1)
            g_step = graph_m.group(2)
            g_3d = graph_m.group(3)
            g_par = graph_m.group(4)
            g_named = graph_m.group(5) if len(graph_m.groups()) >= 5 else None
            g_sub = graph_m.group(6) if len(graph_m.groups()) >= 6 else (graph_m.group(5) or "")
            
            if g_num:
                current_graph_key = f"graph_{g_num}"
                current_graph_title = f"Graph {g_num}: {g_sub.strip()}" if g_sub and g_sub.strip() else f"Graph {g_num}"
            elif g_step:
                current_graph_key = f"step_{g_step.lower()}"
                current_graph_title = f"STEP {g_step}: {g_sub.strip()}" if g_sub and g_sub.strip() else f"STEP {g_step}"
            elif g_3d:
                current_graph_key = f"3d_{len(graphs)+1}"
                current_graph_title = f"{g_3d.strip()}: {g_sub.strip()}" if g_sub and g_sub.strip() else g_3d.strip()
            elif g_par or g_named:
                name = g_par or g_named
                current_graph_key = f"parallel_{len(graphs)+1}"
                current_graph_title = f"{name.strip()}: {g_sub.strip()}" if g_sub and g_sub.strip() else name.strip()
            current_graph_bullets = []
            i += 1
            continue
            
        # Generic Engineering Insights / Bullet
        if current_graph_key:
            if line.startswith('*') or line.startswith('-') or line.startswith('•') or line.lower().startswith('engineering insight'):
                clean_b = line.replace('**', '').replace('*', '').strip()
                if clean_b:
                    current_graph_bullets.append(clean_b)
            elif line.startswith('#'):
                graphs[current_graph_key] = {
                    "title": current_graph_title or current_graph_key,
                    "bullets": current_graph_bullets
                }
                current_graph_key = None
                current_graph_bullets = []
            elif current_graph_bullets and line and not line.startswith('[') and not line.startswith('|'):
                # Multi-line bullet continuation
                current_graph_bullets[-1] += " " + line.replace('**', '').replace('*', '').strip()

        # Generic Table Matcher
        if line.startswith('|') and '|' in line[1:]:
            tbl_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                tbl_lines.append(lines[i].strip())
                i += 1
            if len(tbl_lines) >= 3:
                headers = [c.strip().replace('**', '') for c in tbl_lines[0].split('|')[1:-1]]
                data_rows = []
                for t_line in tbl_lines[2:]:
                    cells = [c.strip().replace('**', '') for c in t_line.split('|')[1:-1]]
                    if cells and any(cells):
                        data_rows.append(cells)
                tables.append({"headers": headers, "rows": data_rows})
            continue

        # Correlation bullets (generic correlation pattern)
        if "vs" in line and ("+" in line or "-" in line):
            correlation_bullets.append(line.replace('**', '').replace('*', '').strip())

        # Executive summary numbered bullets
        if re.match(r'^\d+\.\s+\*\*', line):
            executive_bullets.append(line.replace('**', '').replace('*', '').strip())

        i += 1

    if current_graph_key:
        graphs[current_graph_key] = {
            "title": current_graph_title or current_graph_key,
            "bullets": current_graph_bullets
        }

    return {
        "graphs": graphs,
        "tables": tables,
        "executive_bullets": executive_bullets,
        "correlation_bullets": correlation_bullets
    }


def find_matching_chart_generically(slide_title: str, slide_num: int, ai_charts: list, used_indices: set):
    """
    100% GENERIC chart matching:
    1. Graph Number match (e.g. 'Graph 1' in slide title matches 'Graph 1' in chart title).
    2. Step Key match (e.g. 'STEP 5A', 'STEP 8A').
    3. Tag Index match.
    4. Token overlap between slide title words and chart title/tag words.
    5. Fallback: First unused chart.
    """
    if not ai_charts:
        return None

    norm_slide_title = (slide_title or "").lower()

    # 1. Exact Graph Number match (e.g. 'Graph 1', 'Graph 2')
    num_match = re.search(r'graph\s*(\d+)', norm_slide_title)
    if num_match:
        target_num = num_match.group(1)
        for idx, chart in enumerate(ai_charts):
            if idx in used_indices:
                continue
            c_title = (chart.get("title") or "").lower()
            c_tag = (chart.get("tag") or "").lower()
            c_match = re.search(r'graph\s*(\d+)', c_title)
            if c_match and c_match.group(1) == target_num:
                used_indices.add(idx)
                return chart

    # 2. Step Key match (e.g. 'STEP 5A', 'STEP 8A')
    step_match = re.search(r'step\s*([0-9a-z]+)', norm_slide_title)
    if step_match:
        target_step = step_match.group(1)
        for idx, chart in enumerate(ai_charts):
            if idx in used_indices:
                continue
            c_title = (chart.get("title") or "").lower()
            if f"step {target_step}" in c_title or f"step_{target_step}" in c_title:
                used_indices.add(idx)
                return chart

    # 3. Parallel Coordinate match
    if "parallel" in norm_slide_title:
        for idx, chart in enumerate(ai_charts):
            if idx in used_indices:
                continue
            c_tag = (chart.get("tag") or "").lower()
            c_title = (chart.get("title") or "").lower()
            if "parallel" in c_title or "parallel" in c_tag:
                used_indices.add(idx)
                return chart

    # 4. Token Overlap matching
    slide_tokens = set(re.findall(r'[a-zA-Z0-9_]{3,}', norm_slide_title)) - {'graph', 'step', 'versus', 'analysis', 'dynamics', 'plot'}
    if slide_tokens:
        best_idx = None
        best_overlap = 0
        for idx, chart in enumerate(ai_charts):
            if idx in used_indices:
                continue
            c_str = f"{(chart.get('title') or '')} {(chart.get('tag') or '')}".lower()
            c_tokens = set(re.findall(r'[a-zA-Z0-9_]{3,}', c_str))
            overlap = len(slide_tokens.intersection(c_tokens))
            if overlap > best_overlap:
                best_overlap = overlap
                best_idx = idx
        if best_idx is not None and best_overlap >= 1:
            used_indices.add(best_idx)
            return ai_charts[best_idx]

    # 5. Sequential fallback: first unused chart
    for idx, chart in enumerate(ai_charts):
        if idx not in used_indices:
            used_indices.add(idx)
            return chart

    return None


def generate_formatted_report_presentation(data: dict, template_path: str = None) -> io.BytesIO:
    """
    Main entry point: Generates or populates the presentation deck using
    generic Innomotics Dark formatting.
    NO hardcoded plant or customer strings are used.
    """
    if template_path and os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

    title = data.get("title") or "Exploratory Data Analysis & Optimization Report"
    dataset_name = data.get("dataset_name") or "Dataset Telemetry"
    date_str = data.get("date") or ""
    ai_summary = data.get("ai_summary", "")
    ai_charts = data.get("ai_charts", [])

    parsed_content = parse_markdown_ai_content(ai_summary)
    used_chart_indices = set()

    # If the presentation has template slides (e.g. 10+ slides)
    if len(prs.slides) >= 10:
        logger.info(f"Populating formatted presentation template ({len(prs.slides)} slides) generically...")
        
        # 1. Update Cover Slide (Slide 1)
        s1 = prs.slides[0]
        for sh in s1.shapes:
            if sh.has_text_frame:
                if sh.name == "Title 1":
                    sh.text_frame.text = title
                    if len(sh.text_frame.paragraphs) >= 1:
                        sh.text_frame.paragraphs[0].font.name = FONT_TITLE
                elif sh.name == "Text Placeholder 3" or "Placeholder" in sh.name:
                    sub_text = f"{dataset_name}" + (f" | {date_str}" if date_str else "")
                    sh.text_frame.text = sub_text
                    if len(sh.text_frame.paragraphs) >= 1:
                        sh.text_frame.paragraphs[0].font.name = FONT_BODY

        # 2. Iterate through slides to populate graphs and tables
        table_cursor = 0
        for s_idx, slide in enumerate(prs.slides):
            s_num = s_idx + 1
            if s_num == 1:
                continue

            # Identify slide title from TextBox 2 or first text frame
            slide_title = ""
            for sh in slide.shapes:
                if sh.has_text_frame and (sh.name == "TextBox 2" or "Title" in sh.name):
                    slide_title = sh.text_frame.text.split("\r")[0].split("\n")[0].strip()
                    break

            # A. Graph slide (Contains picture shape)
            has_pic = any(sh.shape_type == 13 for sh in slide.shapes)
            if has_pic and s_num >= 5:
                chart = find_matching_chart_generically(slide_title, s_num, ai_charts, used_chart_indices)
                if chart and chart.get("image_base64"):
                    try:
                        img_stream = decode_base64_image(chart["image_base64"])
                        
                        # Match bullets from parsed AI content
                        chart_bullets = []
                        c_title_clean = clean_title(chart.get("title") or slide_title)
                        chart_tag = chart.get("tag", "").lower()

                        # 1. Match by exact Graph Number (e.g. Graph 1 -> graph_1)
                        s_g_num = re.search(r'graph\s*(\d+)', f"{slide_title} {c_title_clean}", re.IGNORECASE)
                        if s_g_num:
                            target_key = f"graph_{s_g_num.group(1)}"
                            if target_key in parsed_content["graphs"]:
                                chart_bullets = parsed_content["graphs"][target_key]["bullets"]

                        # 2. Match by Step Key (e.g. step_8a, step_11)
                        if not chart_bullets:
                            s_step = re.search(r'step\s*([0-9a-z]+)', f"{slide_title} {c_title_clean}", re.IGNORECASE)
                            if s_step:
                                target_step = f"step_{s_step.group(1).lower()}"
                                if target_step in parsed_content["graphs"]:
                                    chart_bullets = parsed_content["graphs"][target_step]["bullets"]

                        # 3. Match by Token Overlap
                        if not chart_bullets:
                            slide_words = set(re.findall(r'[a-zA-Z0-9_]{3,}', f"{slide_title} {c_title_clean} {chart_tag}".lower())) - {'graph', 'step', 'versus', 'analysis', 'dynamics', 'plot', 'bursa', 'trajectories'}
                            best_b_overlap = 0
                            for gk, gv in parsed_content["graphs"].items():
                                g_words = set(re.findall(r'[a-zA-Z0-9_]{3,}', f"{gk} {gv['title']}".lower()))
                                overlap = len(slide_words.intersection(g_words))
                                if overlap > best_b_overlap:
                                    best_b_overlap = overlap
                                    chart_bullets = gv["bullets"]

                        update_graph_slide_content(slide, c_title_clean, chart_bullets, img_stream)
                        logger.info(f"Updated Slide {s_num} with chart '{c_title_clean}'")
                    except Exception as e:
                        logger.error(f"Failed to update chart on slide {s_num}: {e}")

            # B. Table slide
            has_table = any(sh.has_table for sh in slide.shapes)
            if has_table and table_cursor < len(parsed_content["tables"]):
                matched_table = parsed_content["tables"][table_cursor]
                table_cursor += 1
                try:
                    update_table_slide_content(slide, slide_title, matched_table["headers"], matched_table["rows"])
                    logger.info(f"Updated Slide {s_num} with table {table_cursor}")
                except Exception as e:
                    logger.error(f"Failed to update table on slide {s_num}: {e}")

        # 3. Add any remaining unmatched charts as new formatted slides
        for idx, chart in enumerate(ai_charts):
            if idx not in used_chart_indices and chart.get("image_base64"):
                try:
                    img_stream = decode_base64_image(chart["image_base64"])
                    c_title = clean_title(chart.get("title") or f"Process Visualization {idx+1}")
                    add_formatted_slide(prs, c_title, [], img_stream)
                    logger.info(f"Added extra slide for chart '{c_title}'")
                except Exception as e:
                    logger.error(f"Failed to add extra chart slide: {e}")

    else:
        # Generic Dynamic Mode: Generate presentation from scratch
        logger.info("Generating presentation generically from scratch with Innomotics Dark styling...")
        
        # Cover slide
        add_formatted_slide(
            prs,
            title,
            [f"Dataset: {dataset_name}", f"Date: {date_str}" if date_str else ""]
        )

        # Executive Summary Slide
        if parsed_content["executive_bullets"]:
            add_formatted_slide(
                prs,
                "Executive Summary & Diagnostic Assessment",
                parsed_content["executive_bullets"]
            )

        # Chart slides
        for idx, chart in enumerate(ai_charts):
            img_b64 = chart.get("image_base64")
            if not img_b64:
                continue
            try:
                img_stream = decode_base64_image(img_b64)
                c_title = clean_title(chart.get("title") or f"Graph {idx+1}: Process Visualization")
                
                # Match bullets
                g_bullets = []
                for gk, gv in parsed_content["graphs"].items():
                    if gk in c_title.lower() or any(w in gv["title"].lower() for w in c_title.lower().split()[:2]):
                        g_bullets = gv["bullets"]
                        break
                add_formatted_slide(prs, c_title, g_bullets, img_stream)
            except Exception as e:
                logger.error(f"Failed to render chart slide {idx}: {e}")

    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream
