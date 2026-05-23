"""Generate .pptx from slide JSON data.

Usage:
    python tools/create_ppt.py <slide_data.json> [-c config.yaml] [-o output.pptx]

Supports:
    - Custom PPT template (.pptx) for theme / layouts
    - Granular font configuration per heading level
    - 9 built-in layout types
"""

import json
import sys
import os
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import yaml


# ============================================================
# Config helpers
# ============================================================

def _default_config():
    return {
        "slide": {"width": 13.33, "height": 7.5},
        "template": {"path": "", "layout_mapping": {}},
        "theme": {
            "primary": "#0057B8", "secondary": "#DDEEFF",
            "accent": "#FF6B35", "text_dark": "#1A1A2E",
            "text_light": "#FFFFFF", "background": "#FAFAFA",
        },
        "font": {
            "cover_title": "Microsoft YaHei", "cover_subtitle": "Microsoft YaHei",
            "cover_info": "Microsoft YaHei", "heading": "Microsoft YaHei",
            "body": "Microsoft YaHei", "caption": "Microsoft YaHei",
            "page_number": "Microsoft YaHei",
            "cover_title_size": 40, "cover_subtitle_size": 20,
            "cover_info_size": 14, "heading_size": 32, "body_size": 18,
            "caption_size": 12, "page_number_size": 10,
            "section_number_size": 72, "end_title_size": 44,
        },
        "content": {"max_bullets": 6, "max_words_per_bullet": 20},
        "output": {"add_slide_numbers": True},
    }


def load_config(config_path: str) -> dict:
    cfg = _default_config()
    if config_path and os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            user = yaml.safe_load(f) or {}
        _deep_merge(cfg, user)
    return cfg


def _deep_merge(base: dict, override: dict):
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            _deep_merge(base[k], v)
        else:
            base[k] = v


# ============================================================
# Color helpers
# ============================================================

def parse_color(val) -> RGBColor:
    if isinstance(val, RGBColor):
        return val
    h = str(val).lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _lighten(hex_str: str, amount: int = 60) -> RGBColor:
    c = parse_color(hex_str)
    return RGBColor(min(c[0] + amount, 255), min(c[1] + amount, 255), min(c[2] + amount, 255))


# ============================================================
# Template helpers
# ============================================================

def _get_layout(prs, layout_type: str, mapping: dict) -> object:
    """Pick the right slide layout from the template."""
    defaults = {
        "TITLE": 0, "TOC": 1, "TEXT": 1, "IMAGE_TEXT": 2,
        "TEXT_IMAGE": 2, "RESULT": 3, "FLOWCHART": 4,
        "SECTION": 5, "END": 6,
    }
    idx = mapping.get(layout_type, defaults.get(layout_type, len(prs.slide_layouts) - 1))
    # Clamp to valid range
    idx = max(0, min(idx, len(prs.slide_layouts) - 1))
    return prs.slide_layouts[idx]


def _fill_placeholders(slide, data: dict):
    """Populate template placeholders with slide content."""
    # PHTYPE constants: TITLE=1, BODY=2, SUBTITLE=3, PICTURE=8
    PHTYPE = {1: "title", 3: "subtitle", 2: "body"}

    for ph in slide.placeholders:
        try:
            pht = ph.placeholder_format.type
        except Exception:
            continue

        if pht in (1, 3):  # TITLE / SUBTITLE
            key = PHTYPE[pht]
            text = data.get(key, data.get("title", ""))
            if text:
                ph.text = str(text)
        elif pht == 2:  # BODY
            bullets = data.get("bullets", [])
            if bullets:
                ph.text = "\n".join(bullets)


# ============================================================
# Drawing primitives
# ============================================================

def _add_textbox(slide, left, top, width, height, text, font_name, font_size,
                 bold=False, color=None, alignment=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.alignment = alignment
    if color:
        p.font.color.rgb = parse_color(color)
    return txBox


def _add_bullet_list(slide, left, top, width, height, bullets, font_name, font_size,
                     color, line_spacing=1.5):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"  {bullet}"
        p.font.name = font_name
        p.font.size = Pt(font_size)
        p.font.color.rgb = parse_color(color)
        p.space_after = Pt(font_size * 0.6)
    return txBox


def _add_placeholder_box(slide, left, top, width, height, label, cfg):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = parse_color(cfg["theme"]["secondary"])
    shape.line.fill.background()

    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = label
    p.font.name = cfg["font"]["caption"]
    p.font.size = Pt(cfg["font"]["caption_size"])
    p.font.color.rgb = parse_color(cfg["theme"]["primary"])
    p.alignment = PP_ALIGN.CENTER
    p.space_before = Pt(0)

    txBody = shape.text_frame._txBody
    bodyPr = txBody.find(qn('a:bodyPr'))
    if bodyPr is not None:
        bodyPr.set('anchor', 'ctr')
    return shape


def _add_title_bar(slide, title, cfg):
    """Styled title bar at the top of a content slide."""
    _add_textbox(
        slide, left=0.8, top=0.4, width=cfg["slide"]["width"] - 1.6, height=0.7,
        text=title,
        font_name=cfg["font"]["heading"],
        font_size=cfg["font"]["heading_size"],
        bold=True,
        color=cfg["theme"]["primary"],
    )
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.05), Inches(2.0), Inches(0.03),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = parse_color(cfg["theme"]["accent"])
    line.line.fill.background()


def _add_page_number(slide, page_num, cfg):
    _add_textbox(
        slide,
        left=cfg["slide"]["width"] - 1.0, top=cfg["slide"]["height"] - 0.5,
        width=0.8, height=0.4,
        text=str(page_num),
        font_name=cfg["font"]["page_number"],
        font_size=cfg["font"]["page_number_size"],
        color="#999999",
        alignment=PP_ALIGN.RIGHT,
    )


# ============================================================
# Layout handlers
# ============================================================

def _create_title_slide(slide, data, cfg):
    w, h = cfg["slide"]["width"], cfg["slide"]["height"]
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = parse_color(cfg["theme"]["primary"])

    _add_textbox(slide, 1.5, 2.0, w - 3.0, 1.2,
                 data.get("title", ""),
                 cfg["font"]["cover_title"], cfg["font"]["cover_title_size"],
                 bold=True, color=cfg["theme"]["text_light"],
                 alignment=PP_ALIGN.CENTER)

    subtitle = data.get("subtitle", "")
    if subtitle:
        _add_textbox(slide, 1.5, 3.3, w - 3.0, 0.6,
                     subtitle,
                     cfg["font"]["cover_subtitle"], cfg["font"]["cover_subtitle_size"],
                     color=cfg["theme"]["text_light"],
                     alignment=PP_ALIGN.CENTER)

    author = data.get("author", "")
    date = data.get("date", "")
    if author or date:
        info_text = "  |  ".join(filter(None, [author, date]))
        _add_textbox(slide, 1.5, 4.8, w - 3.0, 0.5,
                     info_text,
                     cfg["font"]["cover_info"], cfg["font"]["cover_info_size"],
                     color="#CCD8E8", alignment=PP_ALIGN.CENTER)


def _create_toc_slide(slide, data, cfg):
    _add_title_bar(slide, "目录", cfg)
    bullets = data.get("bullets", [])
    numbered = [f"{i+1}.  {b}" for i, b in enumerate(bullets)]
    _add_bullet_list(slide, 1.5, 1.6, cfg["slide"]["width"] - 3.0, 5.0,
                     numbered, cfg["font"]["body"], cfg["font"]["body_size"],
                     cfg["theme"]["text_dark"], line_spacing=1.8)


def _create_text_slide(slide, data, cfg):
    _add_title_bar(slide, data.get("title", ""), cfg)
    bullets = data.get("bullets", [])
    _add_bullet_list(slide, 0.8, 1.5, cfg["slide"]["width"] - 1.6, 5.0,
                     bullets, cfg["font"]["body"], cfg["font"]["body_size"],
                     cfg["theme"]["text_dark"])


def _create_image_text_slide(slide, data, cfg):
    w = cfg["slide"]["width"]
    _add_title_bar(slide, data.get("title", ""), cfg)

    label = data.get("image_suggestion", "")
    label_text = f"【配图】\n{label}" if label else "【配图区域】"
    _add_placeholder_box(slide, 0.8, 1.5, w * 0.45, 5.0, label_text, cfg)

    bullets = data.get("bullets", [])
    _add_bullet_list(slide, w * 0.48 + 0.8, 1.5, w * 0.45, 5.0,
                     bullets, cfg["font"]["body"], cfg["font"]["body_size"],
                     cfg["theme"]["text_dark"])


def _create_text_image_slide(slide, data, cfg):
    w = cfg["slide"]["width"]
    _add_title_bar(slide, data.get("title", ""), cfg)

    bullets = data.get("bullets", [])
    _add_bullet_list(slide, 0.8, 1.5, w * 0.45, 5.0,
                     bullets, cfg["font"]["body"], cfg["font"]["body_size"],
                     cfg["theme"]["text_dark"])

    label = data.get("image_suggestion", "")
    label_text = f"【配图】\n{label}" if label else "【配图区域】"
    _add_placeholder_box(slide, w * 0.48 + 0.8, 1.5, w * 0.45, 5.0, label_text, cfg)


def _create_result_slide(slide, data, cfg):
    w, h = cfg["slide"]["width"], cfg["slide"]["height"]
    _add_title_bar(slide, data.get("title", ""), cfg)

    chart_label = data.get("image_suggestion", "【图表区域】")
    _add_placeholder_box(slide, 0.8, 1.4, w - 1.6, h * 0.55, chart_label, cfg)

    bullets = data.get("bullets", [])
    _add_bullet_list(slide, 0.8, 1.4 + h * 0.57, w - 1.6, h * 0.35,
                     bullets, cfg["font"]["body"], cfg["font"]["body_size"],
                     cfg["theme"]["text_dark"])


def _create_flowchart_slide(slide, data, cfg):
    w, h = cfg["slide"]["width"], cfg["slide"]["height"]
    _add_title_bar(slide, data.get("title", ""), cfg)

    label = data.get("image_suggestion", "【流程图 / 架构图】")
    _add_placeholder_box(slide, 1.5, 1.4, w - 3.0, h * 0.70, label, cfg)

    bullets = data.get("bullets", [])
    if bullets:
        _add_bullet_list(slide, 1.5, 1.4 + h * 0.72, w - 3.0, 1.5,
                         bullets, cfg["font"]["body"], cfg["font"]["body_size"],
                         cfg["theme"]["text_dark"])


def _create_section_slide(slide, data, cfg):
    w, h = cfg["slide"]["width"], cfg["slide"]["height"]

    bg_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(w), Inches(h),
    )
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = parse_color(cfg["theme"]["primary"])
    bg_shape.line.fill.background()

    section_num = data.get("section_num", "")
    if section_num:
        _add_textbox(slide, 1.0, 1.5, 3.0, 1.0, str(section_num),
                     cfg["font"]["heading"], cfg["font"]["section_number_size"],
                     bold=True, color=cfg["theme"]["text_light"],
                     alignment=PP_ALIGN.LEFT)

    _add_textbox(slide, 1.0, 2.8, w - 2.0, 1.2, data.get("title", ""),
                 cfg["font"]["heading"], cfg["font"]["heading_size"],
                 bold=True, color=cfg["theme"]["text_light"],
                 alignment=PP_ALIGN.LEFT)

    subtitle = data.get("subtitle", "")
    if subtitle:
        _add_textbox(slide, 1.0, 4.0, w - 2.0, 0.6, subtitle,
                     cfg["font"]["body"], cfg["font"]["body_size"],
                     color=cfg["theme"]["text_light"],
                     alignment=PP_ALIGN.LEFT)


def _create_end_slide(slide, data, cfg):
    w, h = cfg["slide"]["width"], cfg["slide"]["height"]
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = parse_color(cfg["theme"]["primary"])

    _add_textbox(slide, 1.0, 2.5, w - 2.0, 1.2,
                 data.get("title", "谢谢"),
                 cfg["font"]["heading"], cfg["font"]["end_title_size"],
                 bold=True, color=cfg["theme"]["text_light"],
                 alignment=PP_ALIGN.CENTER)

    subtitle = data.get("subtitle", "")
    if subtitle:
        _add_textbox(slide, 1.0, 3.8, w - 2.0, 0.6, subtitle,
                     cfg["font"]["cover_subtitle"], cfg["font"]["cover_subtitle_size"],
                     color=cfg["theme"]["text_light"],
                     alignment=PP_ALIGN.CENTER)


LAYOUT_HANDLERS = {
    "TITLE":     _create_title_slide,
    "TOC":       _create_toc_slide,
    "TEXT":      _create_text_slide,
    "IMAGE_TEXT": _create_image_text_slide,
    "TEXT_IMAGE": _create_text_image_slide,
    "RESULT":    _create_result_slide,
    "FLOWCHART": _create_flowchart_slide,
    "SECTION":   _create_section_slide,
    "END":       _create_end_slide,
}


# ============================================================
# Main
# ============================================================

def create_ppt(slide_data: dict, config: dict, output_path: str):
    template_path = config.get("template", {}).get("path", "")
    use_template = bool(template_path and os.path.exists(template_path))

    if use_template:
        prs = Presentation(template_path)
        print(f"Using template: {template_path}")
    else:
        prs = Presentation()

    # Slide dimensions
    prs.slide_width = Inches(config["slide"]["width"])
    prs.slide_height = Inches(config["slide"]["height"])

    layout_mapping = config.get("template", {}).get("layout_mapping", {})
    slides = slide_data.get("slides", [])

    for i, slide_info in enumerate(slides):
        layout_type = slide_info.get("layout", "TEXT").upper()

        if use_template:
            layout = _get_layout(prs, layout_type, layout_mapping)
            slide = prs.slides.add_slide(layout)
            _fill_placeholders(slide, slide_info)
        else:
            slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

        # Apply our handler for consistent styling
        handler = LAYOUT_HANDLERS.get(layout_type, _create_text_slide)
        try:
            handler(slide, slide_info, config)
        except Exception as e:
            print(f"Warning: slide {i + 1} ({layout_type}): {e}", file=sys.stderr)
            _add_title_bar(slide, slide_info.get("title", f"Slide {i+1}"), config)

        if config["output"].get("add_slide_numbers", True):
            page_num = slide_info.get("page", i + 1)
            _add_page_number(slide, page_num, config)

    prs.save(output_path)
    return len(slides)


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Generate PPTX from slide JSON")
    parser.add_argument("input", help="Path to slide JSON file")
    parser.add_argument("-c", "--config", default=None, help="Path to config YAML")
    parser.add_argument("-o", "--output", default="output.pptx", help="Output PPTX path")

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    with open(args.input, "r", encoding="utf-8") as f:
        slide_data = json.load(f)

    config_path = args.config
    if not config_path:
        default_cfg = Path(__file__).parent.parent / "config.yaml"
        if default_cfg.exists():
            config_path = str(default_cfg)

    config = load_config(config_path)

    # Override template path from command line env or request.yaml override
    # If user specified template in request.yaml, it flows through config

    try:
        count = create_ppt(slide_data, config, args.output)
        abs_path = os.path.abspath(args.output)
        print(f"PPT generated: {abs_path} ({count} slides)")
    except Exception as e:
        print(f"Error generating PPT: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
