"""Generate .pptx from slide JSON data.

Usage:
    python tools/create_ppt.py <slide_data.json> [-c config.yaml] [-o output.pptx]

The slide JSON should follow the format defined in skill.md.
"""

import json
import sys
import os
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import yaml


def load_config(config_path: str) -> dict:
    default = {
        "slide": {"width": 13.33, "height": 7.5},
        "theme": {
            "primary": "#0057B8",
            "secondary": "#DDEEFF",
            "accent": "#FF6B35",
            "text_dark": "#1A1A2E",
            "text_light": "#FFFFFF",
            "background": "#FAFAFA",
        },
        "font": {
            "title": "Microsoft YaHei",
            "body": "Microsoft YaHei",
            "title_size": 32,
            "body_size": 18,
            "subtitle_size": 20,
        },
        "content": {
            "max_bullets": 6,
            "max_words_per_bullet": 20,
        },
        "output": {
            "add_slide_numbers": True,
        },
    }

    if config_path and os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            user = yaml.safe_load(f) or {}
        _deep_merge(default, user)

    return default


def _deep_merge(base: dict, override: dict):
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            _deep_merge(base[k], v)
        else:
            base[k] = v


def parse_color(hex_str) -> RGBColor:
    """Parse a hex color string to RGBColor. Pass-through if already RGBColor."""
    if isinstance(hex_str, RGBColor):
        return hex_str
    hex_str = str(hex_str).lstrip("#")
    return RGBColor(
        int(hex_str[0:2], 16),
        int(hex_str[2:4], 16),
        int(hex_str[4:6], 16),
    )


def _lighten(hex_str: str, amount: int = 60) -> RGBColor:
    """Lighten a hex color by mixing with white."""
    c = parse_color(hex_str)
    r = min(c[0] + amount, 255)
    g = min(c[1] + amount, 255)
    b = min(c[2] + amount, 255)
    return RGBColor(r, g, b)


def _add_textbox(slide, left, top, width, height, text, font_name, font_size,
                 bold=False, color=None, alignment=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
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


def _add_bullet_list(slide, left, top, width, height, bullets, font_name, font_size, color, line_spacing=1.5):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, bullet in enumerate(bullets):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = f"  {bullet}"
        p.font.name = font_name
        p.font.size = Pt(font_size)
        p.font.color.rgb = parse_color(color)
        p.space_after = Pt(font_size * 0.6)
        p.level = 0
        p.alignment = PP_ALIGN.LEFT

    return txBox


def _add_placeholder_box(slide, left, top, width, height, label, cfg):
    """Add a colored placeholder rectangle with a label."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(left), Inches(top),
        Inches(width), Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = parse_color(cfg["theme"]["secondary"])
    shape.line.fill.background()

    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = label
    p.font.name = cfg["font"]["body"]
    p.font.size = Pt(14)
    p.font.color.rgb = parse_color(cfg["theme"]["primary"])
    p.alignment = PP_ALIGN.CENTER
    tf.paragraphs[0].space_before = Pt(0)

    # Vertical center
    shape.text_frame.word_wrap = True
    from pptx.oxml.ns import qn
    txBody = shape.text_frame._txBody
    bodyPr = txBody.find(qn('a:bodyPr'))
    if bodyPr is not None:
        bodyPr.set('anchor', 'ctr')

    return shape


def add_slide_number(slide, page_num, cfg):
    """Add slide number in the bottom-right corner."""
    _add_textbox(
        slide,
        left=cfg["slide"]["width"] - 1.0,
        top=cfg["slide"]["height"] - 0.5,
        width=0.8,
        height=0.4,
        text=str(page_num),
        font_name=cfg["font"]["body"],
        font_size=10,
        color="#999999",
        alignment=PP_ALIGN.RIGHT,
    )


def add_title_bar(slide, title, cfg):
    """Add a styled title at the top of the slide."""
    _add_textbox(
        slide,
        left=0.8,
        top=0.4,
        width=cfg["slide"]["width"] - 1.6,
        height=0.7,
        text=title,
        font_name=cfg["font"]["title"],
        font_size=cfg["font"]["title_size"],
        bold=True,
        color=cfg["theme"]["primary"],
    )

    # Thin accent line under title
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.8), Inches(1.05),
        Inches(2.0), Inches(0.03),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = parse_color(cfg["theme"]["accent"])
    line.line.fill.background()


def create_title_slide(slide, data, cfg):
    """封面页"""
    w, h = cfg["slide"]["width"], cfg["slide"]["height"]

    # Background
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = parse_color(cfg["theme"]["primary"])

    # Title
    _add_textbox(
        slide, left=1.5, top=2.0, width=w - 3.0, height=1.2,
        text=data.get("title", ""),
        font_name=cfg["font"]["title"],
        font_size=40,
        bold=True,
        color=cfg["theme"]["text_light"],
        alignment=PP_ALIGN.CENTER,
    )

    # Subtitle
    subtitle = data.get("subtitle", "")
    if subtitle:
        _add_textbox(
            slide, left=1.5, top=3.3, width=w - 3.0, height=0.6,
            text=subtitle,
            font_name=cfg["font"]["body"],
            font_size=cfg["font"]["subtitle_size"],
            color=cfg["theme"]["text_light"],
            alignment=PP_ALIGN.CENTER,
        )

    # Author & Date
    author = data.get("author", "")
    date = data.get("date", "")
    if author or date:
        info_text = "  |  ".join(filter(None, [author, date]))
        _add_textbox(
            slide, left=1.5, top=4.8, width=w - 3.0, height=0.5,
            text=info_text,
            font_name=cfg["font"]["body"],
            font_size=14,
            color="#CCD8E8",
            alignment=PP_ALIGN.CENTER,
        )


def create_toc_slide(slide, data, cfg):
    """目录页"""
    add_title_bar(slide, "目录", cfg)

    bullets = data.get("bullets", [])
    numbered = [f"{i+1}.  {b}" for i, b in enumerate(bullets)]
    _add_bullet_list(
        slide, left=1.5, top=1.6, width=cfg["slide"]["width"] - 3.0, height=5.0,
        bullets=numbered,
        font_name=cfg["font"]["body"],
        font_size=cfg["font"]["body_size"],
        color=cfg["theme"]["text_dark"],
        line_spacing=1.8,
    )


def create_text_slide(slide, data, cfg):
    """纯文字要点页"""
    add_title_bar(slide, data.get("title", ""), cfg)

    bullets = data.get("bullets", [])
    _add_bullet_list(
        slide, left=0.8, top=1.5, width=cfg["slide"]["width"] - 1.6, height=5.0,
        bullets=bullets,
        font_name=cfg["font"]["body"],
        font_size=cfg["font"]["body_size"],
        color=cfg["theme"]["text_dark"],
    )


def create_image_text_slide(slide, data, cfg):
    """左图右文页"""
    w = cfg["slide"]["width"]
    add_title_bar(slide, data.get("title", ""), cfg)

    # Image placeholder (left 50%)
    label = data.get("image_suggestion", "")
    label_text = f"【配图】\n{label}" if label else "【配图区域】"
    _add_placeholder_box(slide, left=0.8, top=1.5, width=w * 0.45, height=5.0,
                         label=label_text, cfg=cfg)

    # Bullets (right 50%)
    bullets = data.get("bullets", [])
    _add_bullet_list(
        slide, left=w * 0.48 + 0.8, top=1.5, width=w * 0.45, height=5.0,
        bullets=bullets,
        font_name=cfg["font"]["body"],
        font_size=cfg["font"]["body_size"],
        color=cfg["theme"]["text_dark"],
    )


def create_text_image_slide(slide, data, cfg):
    """左文右图页"""
    w = cfg["slide"]["width"]
    add_title_bar(slide, data.get("title", ""), cfg)

    # Bullets (left 50%)
    bullets = data.get("bullets", [])
    _add_bullet_list(
        slide, left=0.8, top=1.5, width=w * 0.45, height=5.0,
        bullets=bullets,
        font_name=cfg["font"]["body"],
        font_size=cfg["font"]["body_size"],
        color=cfg["theme"]["text_dark"],
    )

    # Image placeholder (right 50%)
    label = data.get("image_suggestion", "")
    label_text = f"【配图】\n{label}" if label else "【配图区域】"
    _add_placeholder_box(slide, left=w * 0.48 + 0.8, top=1.5, width=w * 0.45, height=5.0,
                         label=label_text, cfg=cfg)


def create_result_slide(slide, data, cfg):
    """数据/结果页 — 大图表 + 简短结论"""
    w, h = cfg["slide"]["width"], cfg["slide"]["height"]
    add_title_bar(slide, data.get("title", ""), cfg)

    # Chart placeholder (top 65%)
    chart_label = data.get("image_suggestion", "【图表区域】")
    _add_placeholder_box(slide, left=0.8, top=1.4, width=w - 1.6, height=h * 0.55,
                         label=chart_label, cfg=cfg)

    # Conclusion bullets (bottom 35%)
    bullets = data.get("bullets", [])
    _add_bullet_list(
        slide, left=0.8, top=1.4 + h * 0.57, width=w - 1.6, height=h * 0.35,
        bullets=bullets,
        font_name=cfg["font"]["body"],
        font_size=cfg["font"]["body_size"],
        color=cfg["theme"]["text_dark"],
    )


def create_flowchart_slide(slide, data, cfg):
    """流程图/架构图页"""
    w, h = cfg["slide"]["width"], cfg["slide"]["height"]
    add_title_bar(slide, data.get("title", ""), cfg)

    label = data.get("image_suggestion", "【流程图 / 架构图】")
    _add_placeholder_box(slide, left=1.5, top=1.4, width=w - 3.0, height=h * 0.70,
                         label=label, cfg=cfg)

    # Optional bottom text
    bullets = data.get("bullets", [])
    if bullets:
        _add_bullet_list(
            slide, left=1.5, top=1.4 + h * 0.72, width=w - 3.0, height=1.5,
            bullets=bullets,
            font_name=cfg["font"]["body"],
            font_size=14,
            color=cfg["theme"]["text_dark"],
        )


def create_section_slide(slide, data, cfg):
    """章节分隔页"""
    w, h = cfg["slide"]["width"], cfg["slide"]["height"]

    # Large colored block background
    bg_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0), Inches(w), Inches(h),
    )
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = parse_color(cfg["theme"]["primary"])
    bg_shape.line.fill.background()

    # Section number (if provided)
    section_num = data.get("section_num", "")
    if section_num:
        _add_textbox(
            slide, left=1.0, top=1.5, width=3.0, height=1.0,
            text=str(section_num),
            font_name=cfg["font"]["title"],
            font_size=72,
            bold=True,
            color=parse_color(cfg["theme"]["text_light"]),
            alignment=PP_ALIGN.LEFT,
        )

    # Section title
    _add_textbox(
        slide, left=1.0, top=2.8, width=w - 2.0, height=1.2,
        text=data.get("title", ""),
        font_name=cfg["font"]["title"],
        font_size=40,
        bold=True,
        color=cfg["theme"]["text_light"],
        alignment=PP_ALIGN.LEFT,
    )

    # Subtitle
    subtitle = data.get("subtitle", "")
    if subtitle:
        _add_textbox(
            slide, left=1.0, top=4.0, width=w - 2.0, height=0.6,
            text=subtitle,
            font_name=cfg["font"]["body"],
            font_size=18,
            color=parse_color(cfg["theme"]["text_light"]),
            alignment=PP_ALIGN.LEFT,
        )


def create_end_slide(slide, data, cfg):
    """结束页"""
    w, h = cfg["slide"]["width"], cfg["slide"]["height"]

    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = parse_color(cfg["theme"]["primary"])

    _add_textbox(
        slide, left=1.0, top=2.5, width=w - 2.0, height=1.2,
        text=data.get("title", "谢谢"),
        font_name=cfg["font"]["title"],
        font_size=44,
        bold=True,
        color=cfg["theme"]["text_light"],
        alignment=PP_ALIGN.CENTER,
    )

    subtitle = data.get("subtitle", "")
    if subtitle:
        _add_textbox(
            slide, left=1.0, top=3.8, width=w - 2.0, height=0.6,
            text=subtitle,
            font_name=cfg["font"]["body"],
            font_size=18,
            color=parse_color(cfg["theme"]["text_light"]),
            alignment=PP_ALIGN.CENTER,
        )


LAYOUT_HANDLERS = {
    "TITLE": create_title_slide,
    "TOC": create_toc_slide,
    "TEXT": create_text_slide,
    "IMAGE_TEXT": create_image_text_slide,
    "TEXT_IMAGE": create_text_image_slide,
    "RESULT": create_result_slide,
    "FLOWCHART": create_flowchart_slide,
    "SECTION": create_section_slide,
    "END": create_end_slide,
}


def create_ppt(slide_data: dict, config: dict, output_path: str):
    prs = Presentation()

    # Set slide dimensions
    prs.slide_width = Inches(config["slide"]["width"])
    prs.slide_height = Inches(config["slide"]["height"])

    slides = slide_data.get("slides", [])

    for i, slide_info in enumerate(slides):
        # Use blank layout for full control
        layout = prs.slide_layouts[6]  # blank
        slide = prs.slides.add_slide(layout)

        layout_type = slide_info.get("layout", "TEXT").upper()
        handler = LAYOUT_HANDLERS.get(layout_type, create_text_slide)

        try:
            handler(slide, slide_info, config)
        except Exception as e:
            print(f"Warning: error creating slide {i + 1} ({layout_type}): {e}",
                  file=sys.stderr)
            # Fallback: add title as text
            add_title_bar(slide, slide_info.get("title", f"Slide {i+1}"), config)

        # Slide number
        if config["output"].get("add_slide_numbers", True):
            page_num = slide_info.get("page", i + 1)
            add_slide_number(slide, page_num, config)

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

    # Load slide data
    with open(args.input, "r", encoding="utf-8") as f:
        slide_data = json.load(f)

    # Load config (auto-detect if not specified)
    config_path = args.config
    if not config_path:
        default_cfg = Path(__file__).parent.parent / "config.yaml"
        if default_cfg.exists():
            config_path = str(default_cfg)

    config = load_config(config_path)

    # Generate PPT
    try:
        count = create_ppt(slide_data, config, args.output)
        abs_path = os.path.abspath(args.output)
        print(f"PPT generated: {abs_path} ({count} slides)")
    except Exception as e:
        print(f"Error generating PPT: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
