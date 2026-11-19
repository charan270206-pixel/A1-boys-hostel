import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ==============================================================================
# GLOBAL CONSTANTS & PALETTE
# ==============================================================================
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

COLOR_BG_DARK        = RGBColor(11, 17, 32)      # #0B1120 Deep Space Navy
COLOR_CARD_BG        = RGBColor(22, 32, 54)      # #162036 Dark Slate Surface
COLOR_CARD_BORDER    = RGBColor(45, 60, 90)      # #2D3C5A Muted Border
COLOR_CARD_HIGHLIGHT = RGBColor(26, 42, 70)      # #1A2A46 Highlight Card Surface
COLOR_INNER_BOX      = RGBColor(15, 23, 42)      # #0F172A Deep Inset

COLOR_CYAN_ACCENT    = RGBColor(6, 182, 212)     # #06B6D4 Electric Cyan
COLOR_CYAN_LIGHT     = RGBColor(34, 211, 238)    # #22D3EE Bright Cyan
COLOR_BLUE_ACCENT    = RGBColor(56, 189, 248)    # #38BDF8 Sky Blue
COLOR_BLUE_DARK      = RGBColor(37, 99, 235)     # #2563EB Royal Blue
COLOR_GREEN_ACCENT   = RGBColor(16, 185, 129)    # #10B981 Emerald Green
COLOR_RED_ACCENT     = RGBColor(244, 63, 94)     # #F43F5E Rose Red
COLOR_AMBER_ACCENT   = RGBColor(245, 158, 11)    # #F59E0B Amber Alert
COLOR_PURPLE_ACCENT  = RGBColor(168, 85, 247)    # #A855F7 Violet Accent

COLOR_TEXT_PRIMARY   = RGBColor(248, 250, 252)   # #F8FAFC Pure White
COLOR_TEXT_SECONDARY = RGBColor(203, 213, 225)   # #CBD5E1 Light Slate
COLOR_TEXT_MUTED     = RGBColor(148, 163, 184)   # #94A3B8 Supporting Slate

FONT_HEADING = 'Segoe UI'
FONT_BODY    = 'Segoe UI'
FONT_MONO    = 'Consolas'

# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================
def create_deck():
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    return prs

def apply_background(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT)
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_BG_DARK
    bg.line.fill.background()
    return bg

def add_header(slide, badge_text, title_text, subtitle_text, slide_num):
    # Top Accent Line
    glow = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.025))
    glow.fill.solid()
    glow.fill.fore_color.rgb = COLOR_CYAN_ACCENT
    glow.line.fill.background()

    # Category Badge
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.48), Inches(4.0), Inches(0.28))
    badge.fill.solid()
    badge.fill.fore_color.rgb = COLOR_INNER_BOX
    badge.line.color.rgb = COLOR_CYAN_ACCENT
    badge.line.width = Pt(1)
    tf_b = badge.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = tf_b.margin_right = tf_b.margin_top = tf_b.margin_bottom = 0
    p_b = tf_b.paragraphs[0]
    p_b.text = f"  {badge_text.upper()}"
    p_b.font.name = FONT_HEADING
    p_b.font.size = Pt(9)
    p_b.font.bold = True
    p_b.font.color.rgb = COLOR_CYAN_LIGHT

    # Slide Number Badge
    num_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.533), Inches(0.48), Inches(1.0), Inches(0.28))
    num_box.fill.solid()
    num_box.fill.fore_color.rgb = COLOR_INNER_BOX
    num_box.line.color.rgb = COLOR_BLUE_ACCENT
    num_box.line.width = Pt(1)
    tf_n = num_box.text_frame
    tf_n.word_wrap = True
    tf_n.margin_left = tf_n.margin_right = tf_n.margin_top = tf_n.margin_bottom = 0
    p_n = tf_n.paragraphs[0]
    p_n.text = f"{slide_num:02d} / 15"
    p_n.alignment = PP_ALIGN.CENTER
    p_n.font.name = FONT_MONO
    p_n.font.size = Pt(9.5)
    p_n.font.bold = True
    p_n.font.color.rgb = COLOR_BLUE_ACCENT

    # Main Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.78), Inches(11.733), Inches(0.42))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = tf_t.margin_right = tf_t.margin_top = tf_t.margin_bottom = 0
    p_t = tf_t.paragraphs[0]
    p_t.text = title_text
    p_t.font.name = FONT_HEADING
    p_t.font.size = Pt(21)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_TEXT_PRIMARY

    # Subtitle / Context
    sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.23), Inches(11.733), Inches(0.3))
    tf_s = sub_box.text_frame
    tf_s.word_wrap = True
    tf_s.margin_left = tf_s.margin_right = tf_s.margin_top = tf_s.margin_bottom = 0
    p_s = tf_s.paragraphs[0]
    p_s.text = subtitle_text
    p_s.font.name = FONT_BODY
    p_s.font.size = Pt(11.5)
    p_s.font.color.rgb = COLOR_BLUE_ACCENT

    # Footer Metadata Line
    foot = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.015))
    foot.fill.solid()
    foot.fill.fore_color.rgb = COLOR_CARD_BORDER
    foot.line.fill.background()

    foot_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.12), Inches(11.733), Inches(0.25))
    tf_f = foot_box.text_frame
    tf_f.word_wrap = True
    tf_f.margin_left = tf_f.margin_right = tf_f.margin_top = tf_f.margin_bottom = 0
    p_f = tf_f.paragraphs[0]
    p_f.text = "Cybersecurity & Cryptography Methods  •  Computer Science Engineering Seminar  •  Academic Session 2026"
    p_f.font.name = FONT_BODY
    p_f.font.size = Pt(8.5)
    p_f.font.color.rgb = COLOR_TEXT_MUTED

def add_card(slide, left, top, width, height, bg_color=COLOR_CARD_BG, border_color=COLOR_CARD_BORDER, border_width=1):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(border_width)
    return card

def add_photo_card(slide, img_path, left, top, width, height, caption=None, border_color=COLOR_CYAN_ACCENT):
    pic = slide.shapes.add_picture(img_path, left, top, width, height)
    frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    frame.fill.background()
    frame.line.color.rgb = border_color
    frame.line.width = Pt(1.5)
    
    if caption:
        cap_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top + height, width, Inches(0.36))
        cap_box.fill.solid()
        cap_box.fill.fore_color.rgb = COLOR_INNER_BOX
        cap_box.line.color.rgb = COLOR_CARD_BORDER
        cap_box.line.width = Pt(1)
        tf = cap_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.08)
        tf.margin_right = Inches(0.08)
        tf.margin_top = Inches(0.04)
        tf.margin_bottom = Inches(0.04)
        p = tf.paragraphs[0]
        p.text = caption
        p.font.name = FONT_BODY
        p.font.size = Pt(8)
        p.font.color.rgb = COLOR_TEXT_MUTED
    return pic

def set_speaker_notes(slide, notes_text):
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = notes_text.strip()

print("Helper core ready.")
