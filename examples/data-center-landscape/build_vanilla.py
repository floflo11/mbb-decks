#!/usr/bin/env python3
"""
build_vanilla.py — Generates the "vanilla" comparison deck.

This is intentionally a generic-AI-style deck so the skill version can be
contrasted against it. It uses label-style titles, bullet soup, decorative
icons, a pie chart, default Calibri throughout, a date footer, and a
"Thank you / Q&A" closer. None of these are MBB conventions.

DO NOT use this as a template. Read examples/data-center-landscape/README.md
for the comparison and the skill version for the right way to do it.
"""

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt


GENERIC_BLUE = RGBColor(0x1F, 0x77, 0xB4)   # default-ish chart blue
GENERIC_TEAL = RGBColor(0x17, 0xA2, 0xB8)
GENERIC_ORANGE = RGBColor(0xFF, 0x7F, 0x0E)
LIGHT_BLUE_BG = RGBColor(0xE8, 0xF0, 0xF8)
DARK_TEXT = RGBColor(0x33, 0x33, 0x33)
GREY = RGBColor(0x99, 0x99, 0x99)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def add_text(slide, x, y, w, h, text, *, size=14, color=DARK_TEXT, bold=False,
             italic=False, align=PP_ALIGN.LEFT, font="Calibri"):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    return box


def add_label_title(slide, text):
    """Generic 'label' style title — what AI deck tools default to."""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(0), Inches(0), SLIDE_W, Inches(1.0))
    bg.fill.solid()
    bg.fill.fore_color.rgb = GENERIC_BLUE
    bg.line.fill.background()
    add_text(slide, Inches(0.4), Inches(0.25), SLIDE_W - Inches(0.8), Inches(0.6),
             text, size=28, color=WHITE, bold=True)


def add_decorative_icon(slide, x, y, glyph, color=GENERIC_TEAL, size=Inches(0.8)):
    """Decorative circle-with-letter pseudo-icon. The kind of thing AI tools
    add to make slides feel 'visual'."""
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    circle.fill.solid()
    circle.fill.fore_color.rgb = color
    circle.line.fill.background()
    tf = circle.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = glyph
    r.font.name = "Calibri"
    r.font.size = Pt(28)
    r.font.bold = True
    r.font.color.rgb = WHITE


def add_footer(slide, page, total):
    add_text(slide, Inches(0.4), Inches(7.15), Inches(4), Inches(0.25),
             "CONFIDENTIAL — Internal Use Only", size=9, color=GREY, italic=True)
    add_text(slide, Inches(5.5), Inches(7.15), Inches(2), Inches(0.25),
             "May 2026", size=9, color=GREY, align=PP_ALIGN.CENTER)
    add_text(slide, SLIDE_W - Inches(1.5), Inches(7.15), Inches(1), Inches(0.25),
             f"Slide {page}", size=9, color=GREY, align=PP_ALIGN.RIGHT)


def cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = LIGHT_BLUE_BG
    bg.line.fill.background()
    add_text(s, Inches(1), Inches(2.5), Inches(11), Inches(1.5),
             "Data Center Industry Report", size=44, color=GENERIC_BLUE, bold=True)
    add_text(s, Inches(1), Inches(4), Inches(11), Inches(0.6),
             "An Overview of the Cloud Infrastructure Market", size=22, color=DARK_TEXT, italic=True)
    add_text(s, Inches(1), Inches(6.5), Inches(11), Inches(0.4),
             "Prepared by: Iris Meng  |  May 2026  |  CONFIDENTIAL",
             size=12, color=GREY)


def agenda(prs, page, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_label_title(s, "Agenda")
    items = [
        "1.  Introduction",
        "2.  Market Overview",
        "3.  Key Players in the Industry",
        "4.  Major Market Trends",
        "5.  Challenges and Opportunities",
        "6.  Future Outlook",
        "7.  Conclusion and Q&A",
    ]
    for i, item in enumerate(items):
        add_text(s, Inches(1.2), Inches(1.6 + 0.6 * i), Inches(11), Inches(0.5),
                 item, size=20, color=DARK_TEXT)
    add_footer(s, page, total)


def introduction(prs, page, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_label_title(s, "Introduction")
    add_decorative_icon(s, Inches(0.6), Inches(1.4), "i", GENERIC_TEAL)
    add_text(s, Inches(1.7), Inches(1.5), Inches(11), Inches(0.5),
             "About this report", size=18, color=GENERIC_BLUE, bold=True)
    add_text(s, Inches(1.7), Inches(2.0), Inches(11), Inches(4.5),
             "The data center industry is one of the most important sectors in the modern "
             "economy — driving cloud computing, artificial intelligence, and digital "
             "transformation across every industry vertical. This report provides a "
             "comprehensive overview of the data center landscape, including market sizing, "
             "competitive dynamics, key trends, and future outlook.\n\n"
             "We will explore how hyperscale providers and colocation operators compete, "
             "where new capacity is being built, and what the next several years of "
             "infrastructure investment will look like — with particular attention to the "
             "transformative impact of AI workloads.",
             size=14, color=DARK_TEXT)
    add_footer(s, page, total)


def market_overview(prs, page, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_label_title(s, "Market Overview")
    add_decorative_icon(s, Inches(0.6), Inches(1.4), "M", GENERIC_ORANGE)
    add_text(s, Inches(1.7), Inches(1.5), Inches(11), Inches(0.5),
             "Key statistics", size=18, color=GENERIC_BLUE, bold=True)
    bullets = [
        "The global data center market is large and growing rapidly",
        "Cloud infrastructure spending is increasing year over year",
        "Hyperscale providers are the largest segment of the market",
        "Colocation continues to play an important role for enterprises",
        "AI workloads are a major driver of new capacity",
        "Power and cooling are increasingly important considerations",
        "Sustainability is a growing focus for operators and customers alike",
        "The industry is geographically diverse with strong regional players",
    ]
    box = s.shapes.add_textbox(Inches(1.7), Inches(2.0), Inches(11), Inches(5))
    tf = box.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.4
        r = p.add_run()
        r.text = "•  " + b
        r.font.name = "Calibri"
        r.font.size = Pt(14)
        r.font.color.rgb = DARK_TEXT
    add_footer(s, page, total)


def key_players(prs, page, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_label_title(s, "Key Players")
    # Pie chart with 3D-ish styling (one of the AI-deck giveaways)
    cd = CategoryChartData()
    cd.categories = ["AWS", "Azure", "Google Cloud", "Alibaba", "Oracle", "Other"]
    cd.add_series("Share", [32, 23, 11, 4, 3, 27])
    cf = s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(0.5), Inches(1.3),
                            Inches(7), Inches(5.5), cd).chart
    cf.has_title = True
    cf.chart_title.text_frame.text = "Cloud Market Share"
    cf.has_legend = True
    cf.legend.position = XL_LEGEND_POSITION.RIGHT

    add_text(s, Inches(8), Inches(1.5), Inches(5), Inches(0.5),
             "Major hyperscale providers:", size=16, color=GENERIC_BLUE, bold=True)
    notes = [
        "Amazon Web Services — the market leader",
        "Microsoft Azure — strong enterprise position",
        "Google Cloud — known for AI and data analytics",
        "Alibaba Cloud — leader in Asia-Pacific",
        "Oracle Cloud — focused on enterprise applications",
    ]
    box = s.shapes.add_textbox(Inches(8), Inches(2.1), Inches(5), Inches(5))
    tf = box.text_frame
    tf.word_wrap = True
    for i, n in enumerate(notes):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.3
        r = p.add_run()
        r.text = "•  " + n
        r.font.name = "Calibri"
        r.font.size = Pt(13)
        r.font.color.rgb = DARK_TEXT
    add_footer(s, page, total)


def trends(prs, page, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_label_title(s, "Major Market Trends")
    add_decorative_icon(s, Inches(0.6), Inches(1.4), "T", GENERIC_ORANGE)
    add_text(s, Inches(1.7), Inches(1.5), Inches(11), Inches(0.5),
             "Trends shaping the future of data centers", size=18, color=GENERIC_BLUE, bold=True)
    bullets = [
        "Artificial Intelligence — driving unprecedented demand for compute capacity",
        "Sustainability — operators are racing to reduce their carbon footprint",
        "Edge Computing — pushing data centers closer to end users",
        "Liquid Cooling — becoming the new standard for high-density racks",
        "Geographic Diversification — new regions emerging beyond traditional hubs",
        "Custom Silicon — hyperscalers building their own AI accelerators",
        "Renewable Power — solar, wind, and increasingly nuclear PPAs are now table stakes",
    ]
    box = s.shapes.add_textbox(Inches(1.7), Inches(2.0), Inches(11), Inches(5))
    tf = box.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.35
        r = p.add_run()
        r.text = "•  " + b
        r.font.name = "Calibri"
        r.font.size = Pt(14)
        r.font.color.rgb = DARK_TEXT
    add_footer(s, page, total)


def challenges(prs, page, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_label_title(s, "Challenges and Opportunities")
    add_decorative_icon(s, Inches(0.6), Inches(1.4), "?", GENERIC_TEAL)
    add_text(s, Inches(1.7), Inches(1.5), Inches(11), Inches(0.5),
             "Key challenges facing the industry", size=18, color=GENERIC_BLUE, bold=True)
    bullets = [
        "Power availability is becoming a major constraint, especially in established markets",
        "Cooling at high densities is technically demanding and capital intensive",
        "Talent shortages in critical infrastructure roles",
        "Supply chain delays for transformers, switchgear, and other long-lead equipment",
        "Permitting and community opposition slowing some projects",
        "Rapidly changing AI workload profiles make long-term planning difficult",
    ]
    box = s.shapes.add_textbox(Inches(1.7), Inches(2.0), Inches(11), Inches(5))
    tf = box.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.4
        r = p.add_run()
        r.text = "•  " + b
        r.font.name = "Calibri"
        r.font.size = Pt(14)
        r.font.color.rgb = DARK_TEXT
    add_footer(s, page, total)


def future_outlook(prs, page, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_label_title(s, "Future Outlook")
    add_decorative_icon(s, Inches(0.6), Inches(1.4), "★", GENERIC_ORANGE)
    add_text(s, Inches(1.7), Inches(1.5), Inches(11), Inches(0.5),
             "What to expect in the coming years", size=18, color=GENERIC_BLUE, bold=True)
    bullets = [
        "Strong growth is expected to continue",
        "AI will remain the primary demand driver",
        "Power innovation will accelerate, including SMRs and behind-the-meter generation",
        "Geographic diversification will continue as legacy hubs hit constraints",
        "M&A activity is likely to increase as smaller operators consolidate",
        "Technology innovation will continue at a rapid pace",
    ]
    box = s.shapes.add_textbox(Inches(1.7), Inches(2.0), Inches(11), Inches(5))
    tf = box.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.4
        r = p.add_run()
        r.text = "•  " + b
        r.font.name = "Calibri"
        r.font.size = Pt(14)
        r.font.color.rgb = DARK_TEXT
    add_footer(s, page, total)


def conclusion(prs, page, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_label_title(s, "Conclusion")
    add_text(s, Inches(1), Inches(2), Inches(11), Inches(4),
             "The data center industry is at an inflection point. Driven by AI, "
             "sustainability requirements, and changing customer expectations, the "
             "next several years will bring significant change. Operators that can "
             "secure power, build at scale, and serve AI workloads will be best "
             "positioned for success. Thank you for your attention.",
             size=18, color=DARK_TEXT, italic=True)
    add_footer(s, page, total)


def main(output_path):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    paged = [agenda, introduction, market_overview, key_players,
             trends, challenges, future_outlook, conclusion]
    total = 1 + len(paged) + 1  # cover + paged + thank-you
    cover(prs)
    page = 2
    for fn in paged:
        fn(prs, page, total)
        page += 1
    thank_you_no_footer(prs)
    prs.save(output_path)
    print(f"Wrote {output_path} ({total} slides)")


def thank_you_no_footer(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = GENERIC_BLUE
    bg.line.fill.background()
    add_text(s, Inches(0.5), Inches(2.8), SLIDE_W - Inches(1), Inches(1.5),
             "Thank You!", size=72, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.5), Inches(4.5), SLIDE_W - Inches(1), Inches(0.6),
             "Questions & Discussion", size=28, color=WHITE, italic=True, align=PP_ALIGN.CENTER)


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "vanilla.pptx"
    main(out)
