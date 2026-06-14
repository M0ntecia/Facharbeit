"""
Build a modern, Keynote-style PowerPoint presentation for the Facharbeit
"Wie Schule wirkt: Strukturelle Unterschiede und ihr Effekt auf das Wohlbefinden
von Schuelerinnen und Schuelern" (Seminarfachkurs Glueck).

The resulting .pptx is saved to "Wie_Schule_wirkt_Praesentation.pptx" in the
workspace root and can be opened in Keynote on iPad (Keynote will import the
file and convert it to its native format).

Design goals:
  * Pastel palette (#CFEFFF, #FFD9E8, #FFF7C7, white, dark grey)
  * Generous whitespace, rounded shapes, soft shadows
  * Vector illustrations composed from primitive shapes (no clipart)
  * Per-slide transitions and gentle per-paragraph build-in animations
  * Detailed speaker notes (~1 minute each, ~15 min total)
"""

from __future__ import annotations

import copy
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn, nsmap
from lxml import etree


# ----------------------------------------------------------------------------- 
# Palette and helpers
# ----------------------------------------------------------------------------- 

BLUE      = RGBColor(0xCF, 0xEF, 0xFF)
PINK      = RGBColor(0xFF, 0xD9, 0xE8)
YELLOW    = RGBColor(0xFF, 0xF7, 0xC7)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
DARK      = RGBColor(0x2B, 0x2D, 0x42)   # dark grey / near-navy
GREY      = RGBColor(0x6B, 0x6F, 0x7B)
MUTED     = RGBColor(0x9A, 0x9D, 0xA8)
ACCENT    = RGBColor(0x4C, 0x9A, 0xFF)   # accent blue
GREEN     = RGBColor(0xCB, 0xEF, 0xD3)
GREEN_DK  = RGBColor(0x3F, 0x9B, 0x5A)
PINK_DK   = RGBColor(0xD6, 0x4F, 0x8B)
YELLOW_DK = RGBColor(0xB8, 0x8E, 0x1F)
BLUE_DK   = RGBColor(0x2C, 0x6B, 0xA8)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def color_to_xml(rgb: RGBColor) -> str:
    return "".join(f"{c:02X}" for c in rgb)


def add_background(slide, rgb: RGBColor) -> None:
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H
    )
    bg.line.fill.background()
    bg.fill.solid()
    bg.fill.fore_color.rgb = rgb
    bg.shadow.inherit = False
    # send to back is done by adding it first; we rely on insertion order


def add_rect(
    slide,
    x, y, w, h,
    fill=WHITE,
    line=None,
    radius=0.08,
    shadow=True,
    rounded=True,
):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    if rounded:
        # adjust corner radius (relative, 0..0.5 of shorter side)
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    if not shadow:
        _remove_shadow(shp)
    else:
        _soft_shadow(shp)
    return shp


def add_oval(slide, x, y, w, h, fill=WHITE, line=None, shadow=True):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    if shadow:
        _soft_shadow(shp)
    else:
        _remove_shadow(shp)
    return shp


def add_triangle(slide, x, y, w, h, fill=WHITE, line=None, shadow=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    if shadow:
        _soft_shadow(shp)
    else:
        _remove_shadow(shp)
    return shp


def _get_spPr(shape):
    return shape._element.spPr


def _soft_shadow(shape) -> None:
    """Apply a subtle soft outer shadow via spPr/effectLst."""
    spPr = _get_spPr(shape)
    for old in spPr.findall(qn("a:effectLst")):
        spPr.remove(old)
    effectLst = etree.SubElement(spPr, qn("a:effectLst"))
    outerShdw = etree.SubElement(
        effectLst,
        qn("a:outerShdw"),
        {
            "blurRad": "63500",
            "dist": "25400",
            "dir": "5400000",
            "algn": "t",
            "rotWithShape": "0",
        },
    )
    srgb = etree.SubElement(outerShdw, qn("a:srgbClr"), {"val": "000000"})
    etree.SubElement(srgb, qn("a:alpha"), {"val": "12000"})


def _remove_shadow(shape) -> None:
    spPr = _get_spPr(shape)
    for old in spPr.findall(qn("a:effectLst")):
        spPr.remove(old)
    etree.SubElement(spPr, qn("a:effectLst"))


def add_text(
    slide,
    x, y, w, h,
    text,
    size=20,
    bold=False,
    color=DARK,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    font="Helvetica Neue",
    line_spacing=1.15,
):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    if isinstance(text, str):
        lines = [text]
    else:
        lines = list(text)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = ln
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb


def add_bullets(
    slide,
    x, y, w, h,
    bullets,
    size=18,
    color=DARK,
    bullet_color=ACCENT,
    font="Helvetica Neue",
    line_spacing=1.4,
    bold=False,
):
    """A clean bullet list with a coloured dot prefix."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    for i, txt in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing
        # dot
        dot = p.add_run()
        dot.text = "•  "
        dot.font.name = font
        dot.font.size = Pt(size + 2)
        dot.font.bold = True
        dot.font.color.rgb = bullet_color
        # text
        r = p.add_run()
        r.text = txt
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb


def add_decor_blob(slide, x, y, w, h, color, alpha=20):
    """A faint translucent oval used as a soft decorative blob.

    alpha is in percent (0-100); kept low for subtle backgrounds.
    The blob is rendered without outline and without shadow.
    """
    shp = add_oval(slide, x, y, w, h, fill=color, shadow=False)
    spPr = _get_spPr(shp)
    solidFill = spPr.find(qn("a:solidFill"))
    if solidFill is not None:
        srgb = solidFill.find(qn("a:srgbClr"))
        if srgb is not None:
            etree.SubElement(srgb, qn("a:alpha"), {"val": str(alpha * 1000)})
    return shp


def set_notes(slide, text: str) -> None:
    notes_tf = slide.notes_slide.notes_text_frame
    notes_tf.clear()
    paragraphs = [p for p in text.strip().split("\n") if p.strip() != ""]
    for i, para in enumerate(paragraphs):
        p = notes_tf.paragraphs[0] if i == 0 else notes_tf.add_paragraph()
        r = p.add_run()
        r.text = para.strip()
        r.font.size = Pt(13)
        r.font.name = "Helvetica Neue"


# -----------------------------------------------------------------------------
# Transition and animation XML
# -----------------------------------------------------------------------------

P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
P15_NS = "http://schemas.microsoft.com/office/powerpoint/2012/main"
P159_NS = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

NSMAP_FULL = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "mc": MC_NS,
    "p14": P14_NS,
    "p15": P15_NS,
    "p159": P159_NS,
}


def _nsq(prefix: str, local: str) -> str:
    return f"{{{NSMAP_FULL[prefix]}}}{local}"


def set_transition(slide, kind: str) -> None:
    """Insert a <p:transition> element (with the given preset) into the slide.

    Supported `kind` values map to Keynote-style transitions:
      fade, zoom, movein, magicmove, scale, push, dissolve, pop, slide, rotate
    """
    sld = slide._element  # <p:sld>
    # remove any existing transition
    for old in sld.findall(qn("p:transition")):
        sld.remove(old)

    spd = "med"
    if kind == "fade":
        xml = f'<p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}"><p:fade/></p:transition>'
    elif kind == "zoom":
        # PowerPoint 2010+ zoom (in)
        xml = (
            f'<mc:AlternateContent xmlns:mc="{MC_NS}">'
            f'<mc:Choice xmlns:p14="{P14_NS}" Requires="p14">'
            f'<p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}" p14:dur="900">'
            f'<p14:zoom dir="in"/></p:transition></mc:Choice>'
            f'<mc:Fallback><p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}"><p:fade/></p:transition></mc:Fallback>'
            f'</mc:AlternateContent>'
        )
    elif kind == "movein":
        # cover (from left) - "move in" maps well to cover
        xml = f'<p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}"><p:cover dir="l"/></p:transition>'
    elif kind == "magicmove":
        # Morph (PowerPoint 2016+) - Keynote's Magic Move equivalent
        xml = (
            f'<mc:AlternateContent xmlns:mc="{MC_NS}">'
            f'<mc:Choice xmlns:p159="{P159_NS}" Requires="p159">'
            f'<p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}" p14:dur="1200" xmlns:p14="{P14_NS}">'
            f'<p159:morph option="byObject"/></p:transition></mc:Choice>'
            f'<mc:Fallback><p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}"><p:fade/></p:transition></mc:Fallback>'
            f'</mc:AlternateContent>'
        )
    elif kind == "scale":
        # 2010 'zoom' out variant works as Scale
        xml = (
            f'<mc:AlternateContent xmlns:mc="{MC_NS}">'
            f'<mc:Choice xmlns:p14="{P14_NS}" Requires="p14">'
            f'<p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}" p14:dur="900">'
            f'<p14:zoom dir="out"/></p:transition></mc:Choice>'
            f'<mc:Fallback><p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}"><p:fade/></p:transition></mc:Fallback>'
            f'</mc:AlternateContent>'
        )
    elif kind == "push":
        xml = f'<p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}"><p:push dir="l"/></p:transition>'
    elif kind == "dissolve":
        xml = f'<p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}"><p:dissolve/></p:transition>'
    elif kind == "pop":
        # 'split' is the closest standard transition that resembles "pop"
        xml = (
            f'<mc:AlternateContent xmlns:mc="{MC_NS}">'
            f'<mc:Choice xmlns:p14="{P14_NS}" Requires="p14">'
            f'<p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}" p14:dur="700">'
            f'<p14:flash/></p:transition></mc:Choice>'
            f'<mc:Fallback><p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}"><p:split orient="horz" dir="out"/></p:transition></mc:Fallback>'
            f'</mc:AlternateContent>'
        )
    elif kind == "slide":
        xml = f'<p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}"><p:push dir="r"/></p:transition>'
    elif kind == "rotate":
        xml = (
            f'<mc:AlternateContent xmlns:mc="{MC_NS}">'
            f'<mc:Choice xmlns:p14="{P14_NS}" Requires="p14">'
            f'<p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}" p14:dur="1100">'
            f'<p14:ferris dir="l"/></p:transition></mc:Choice>'
            f'<mc:Fallback><p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}"><p:wheel spokes="1"/></p:transition></mc:Fallback>'
            f'</mc:AlternateContent>'
        )
    else:
        xml = f'<p:transition xmlns:p="{NSMAP_FULL["p"]}" spd="{spd}"><p:fade/></p:transition>'

    el = etree.fromstring(xml)
    sld.append(el)


def add_fade_build(slide, text_shape_ids: list[int], delay_ms: int = 350) -> None:
    """Add a simple build-in: fade-in for each shape sequentially.

    `text_shape_ids` is a list of shape spIds (in render order). Each gets a
    fade-in animation auto-triggered after the previous one with `delay_ms`.
    The first fades in on slide load.
    """
    if not text_shape_ids:
        return

    sld = slide._element

    # build the cTn/par children string
    par_items = []
    for idx, sp_id in enumerate(text_shape_ids):
        delay = "0" if idx == 0 else str(delay_ms)
        par_items.append(
            f'''<p:par>
              <p:cTn id="{100 + idx*10}" presetID="10" presetClass="entr" presetSubtype="0"
                     fill="hold" grpId="0" nodeType="afterEffect">
                <p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>
                <p:childTnLst>
                  <p:set>
                    <p:cBhvr>
                      <p:cTn id="{101 + idx*10}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
                      <p:tgtEl><p:spTgt spid="{sp_id}"/></p:tgtEl>
                      <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                    </p:cBhvr>
                    <p:to><p:strVal val="visible"/></p:to>
                  </p:set>
                  <p:anim calcmode="lin" valueType="num">
                    <p:cBhvr additive="base">
                      <p:cTn id="{102 + idx*10}" dur="600" fill="hold"/>
                      <p:tgtEl><p:spTgt spid="{sp_id}"/></p:tgtEl>
                      <p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst>
                    </p:cBhvr>
                    <p:tavLst>
                      <p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav>
                      <p:tav tm="100000"><p:val><p:fltVal val="1"/></p:val></p:tav>
                    </p:tavLst>
                  </p:anim>
                </p:childTnLst>
              </p:cTn>
            </p:par>'''
        )

    # build buildLst (so PPT does not render objects until animation triggers)
    bld_items = "".join(
        f'<p:bldP spid="{sp_id}" grpId="0"/>' for sp_id in text_shape_ids
    )

    timing_xml = f'''
    <p:timing xmlns:p="{NSMAP_FULL["p"]}">
      <p:tnLst>
        <p:par>
          <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
            <p:childTnLst>
              <p:seq concurrent="1" nextAc="seek">
                <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                  <p:childTnLst>
                    {''.join(par_items)}
                  </p:childTnLst>
                </p:cTn>
                <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
                <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
              </p:seq>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
      <p:bldLst>
        {bld_items}
      </p:bldLst>
    </p:timing>
    '''

    el = etree.fromstring(timing_xml)

    # remove any existing timing
    for old in sld.findall(qn("p:timing")):
        sld.remove(old)
    sld.append(el)


def shape_id(shape) -> int:
    return int(shape._element.nvSpPr.cNvPr.get("id"))


# -----------------------------------------------------------------------------
# Illustration primitives (composed of shapes)
# -----------------------------------------------------------------------------

def draw_school(slide, cx, cy, scale=1.0):
    """A friendly, minimalist school illustration."""
    s = scale
    # roof
    add_triangle(
        slide,
        cx - Inches(2.2 * s), cy - Inches(2.0 * s),
        Inches(4.4 * s), Inches(1.4 * s),
        fill=PINK_DK,
    )
    # building body
    add_rect(
        slide,
        cx - Inches(1.9 * s), cy - Inches(0.6 * s),
        Inches(3.8 * s), Inches(2.4 * s),
        fill=WHITE, line=BLUE_DK, radius=0.05, shadow=True,
    )
    # bell tower
    add_rect(
        slide,
        cx - Inches(0.4 * s), cy - Inches(2.7 * s),
        Inches(0.8 * s), Inches(0.9 * s),
        fill=YELLOW, line=BLUE_DK, radius=0.2, shadow=False,
    )
    # door
    add_rect(
        slide,
        cx - Inches(0.45 * s), cy + Inches(0.8 * s),
        Inches(0.9 * s), Inches(1.0 * s),
        fill=BLUE, line=BLUE_DK, radius=0.12, shadow=False,
    )
    # windows
    for dx in (-Inches(1.25 * s), Inches(0.5 * s)):
        add_rect(
            slide,
            cx + dx, cy - Inches(0.2 * s),
            Inches(0.75 * s), Inches(0.65 * s),
            fill=YELLOW, line=BLUE_DK, radius=0.18, shadow=False,
        )
    # sun
    add_oval(
        slide,
        cx + Inches(2.0 * s), cy - Inches(2.5 * s),
        Inches(0.75 * s), Inches(0.75 * s),
        fill=YELLOW_DK, shadow=False,
    )


def draw_student(slide, x, y, scale=1.0, face=PINK, body=BLUE, mood="happy"):
    s = scale
    # head
    add_oval(slide, x, y, Inches(0.7 * s), Inches(0.7 * s), fill=face, shadow=False)
    # body
    add_rect(
        slide,
        x - Inches(0.05 * s), y + Inches(0.65 * s),
        Inches(0.8 * s), Inches(1.1 * s),
        fill=body, radius=0.35, shadow=False,
    )
    # face
    if mood == "happy":
        # eyes
        add_oval(slide, x + Inches(0.18 * s), y + Inches(0.25 * s), Inches(0.07 * s), Inches(0.07 * s), fill=DARK, shadow=False)
        add_oval(slide, x + Inches(0.45 * s), y + Inches(0.25 * s), Inches(0.07 * s), Inches(0.07 * s), fill=DARK, shadow=False)
        # smile arc using a small ellipse
        add_oval(slide, x + Inches(0.22 * s), y + Inches(0.40 * s), Inches(0.26 * s), Inches(0.16 * s), fill=PINK_DK, shadow=False)
    else:
        add_oval(slide, x + Inches(0.18 * s), y + Inches(0.27 * s), Inches(0.07 * s), Inches(0.07 * s), fill=DARK, shadow=False)
        add_oval(slide, x + Inches(0.45 * s), y + Inches(0.27 * s), Inches(0.07 * s), Inches(0.07 * s), fill=DARK, shadow=False)
        # frown - small rect under nose
        add_rect(
            slide,
            x + Inches(0.22 * s), y + Inches(0.5 * s),
            Inches(0.26 * s), Inches(0.04 * s),
            fill=DARK, radius=0.5, shadow=False,
        )


def draw_balance(slide, cx, cy, scale=1.0):
    """A stylised balance / scale."""
    s = scale
    # base triangle
    add_triangle(slide, cx - Inches(0.8 * s), cy + Inches(0.5 * s), Inches(1.6 * s), Inches(1.5 * s), fill=BLUE_DK, shadow=False)
    # vertical pole
    add_rect(slide, cx - Inches(0.06 * s), cy - Inches(1.8 * s), Inches(0.12 * s), Inches(2.6 * s), fill=BLUE_DK, radius=0.4, shadow=False)
    # cross beam
    add_rect(slide, cx - Inches(2.2 * s), cy - Inches(1.85 * s), Inches(4.4 * s), Inches(0.16 * s), fill=BLUE_DK, radius=0.5, shadow=False)
    # left pan
    add_oval(slide, cx - Inches(2.9 * s), cy - Inches(1.5 * s), Inches(1.6 * s), Inches(0.4 * s), fill=PINK, shadow=True)
    # right pan
    add_oval(slide, cx + Inches(1.3 * s), cy - Inches(1.5 * s), Inches(1.6 * s), Inches(0.4 * s), fill=YELLOW, shadow=True)


# -----------------------------------------------------------------------------
# Build slides
# -----------------------------------------------------------------------------

def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    blank = prs.slide_layouts[6]

    # ------------------------------------------------------------------- 1
    slide = prs.slides.add_slide(blank)
    add_background(slide, WHITE)
    add_decor_blob(slide, Inches(-3.5), Inches(-3), Inches(7), Inches(7), BLUE, alpha=22)
    add_decor_blob(slide, Inches(10.5), Inches(5.5), Inches(5), Inches(5), PINK, alpha=22)
    add_decor_blob(slide, Inches(11), Inches(-1.5), Inches(2.6), Inches(2.6), YELLOW, alpha=35)

    s_seminar = add_text(
        slide, Inches(0.8), Inches(0.9), Inches(8), Inches(0.4),
        "SEMINARFACHKURS GLÜCK", size=14, bold=True, color=BLUE_DK,
    )

    # main title
    s_title = add_text(
        slide, Inches(0.8), Inches(1.6), Inches(8.0), Inches(1.2),
        "Wie Schule wirkt",
        size=58, bold=True, color=DARK, line_spacing=1.0,
    )

    # subtitle lines
    s_sub = add_text(
        slide, Inches(0.8), Inches(3.0), Inches(7.6), Inches(2.2),
        ["Strukturelle Unterschiede und ihr",
         "Effekt auf das Wohlbefinden von",
         "Schülerinnen und Schülern"],
        size=24, color=GREY, line_spacing=1.3,
    )

    # author pill
    pill = add_rect(slide, Inches(0.8), Inches(6.0), Inches(3.6), Inches(0.7), fill=YELLOW, radius=0.45, shadow=False)
    pill_t = add_text(slide, Inches(0.8), Inches(6.0), Inches(3.6), Inches(0.7),
             "Svea Timphus  •  März 2026",
             size=16, bold=True, color=DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # illustration on right side (kept compact, away from text)
    draw_school(slide, Inches(10.7), Inches(4.0), scale=0.95)
    draw_student(slide, Inches(8.8), Inches(5.6), scale=0.85, face=PINK, body=BLUE, mood="happy")
    draw_student(slide, Inches(11.9), Inches(5.6), scale=0.85, face=YELLOW, body=PINK_DK, mood="sad")

    add_fade_build(slide, [shape_id(s_seminar), shape_id(s_title), shape_id(s_sub), shape_id(pill), shape_id(pill_t)])
    set_transition(slide, "fade")

    set_notes(slide, """
Herzlich willkommen zu meiner Präsentation im Seminarfachkurs Glück.
Mein Thema lautet: „Wie Schule wirkt – Strukturelle Unterschiede und ihr Effekt auf das Wohlbefinden von Schülerinnen und Schülern.“
In den nächsten rund 15 Minuten möchte ich euch zeigen, dass Schule weit mehr ist als nur ein Ort, an dem Wissen vermittelt wird. Sie ist auch ein Lebensraum, der unser Wohlbefinden, unsere Motivation und unser Selbstbild stark prägt.
Wir werden uns gemeinsam ansehen, wie das deutsche Schulsystem aufgebaut ist, welche Auswirkungen Leistungsdruck und Schulklima haben, wie es im internationalen Vergleich abschneidet und welche Reformideen gerade diskutiert werden.
Lasst uns mit der Frage starten, warum dieses Thema überhaupt so wichtig ist.
""")

    # ------------------------------------------------------------------- 2
    slide = prs.slides.add_slide(blank)
    add_background(slide, WHITE)
    add_decor_blob(slide, Inches(-3), Inches(4), Inches(7), Inches(6), BLUE, alpha=25)
    add_decor_blob(slide, Inches(10), Inches(-2), Inches(6), Inches(5), PINK, alpha=22)

    s_eyebrow = add_text(slide, Inches(0.8), Inches(0.7), Inches(8), Inches(0.4),
                         "01  ·  EINSTIEG", size=13, bold=True, color=BLUE_DK)
    s_h = add_text(slide, Inches(0.8), Inches(1.1), Inches(11), Inches(1.0),
                   "Warum ist das Thema wichtig?",
                   size=38, bold=True, color=DARK)

    s_bullets = add_bullets(
        slide, Inches(0.8), Inches(2.6), Inches(6.8), Inches(2.6),
        [
            "Schule prägt unseren Alltag",
            "Leistungsdruck nimmt zu",
            "Bildung beeinflusst unsere Zukunft",
            "Wohlbefinden ist Teil erfolgreicher Bildung",
        ],
        size=20, bullet_color=PINK_DK,
    )

    # leitfrage box
    qbox = add_rect(slide, Inches(0.8), Inches(5.5), Inches(11.6), Inches(1.5),
                    fill=YELLOW, radius=0.18, shadow=True)
    add_text(slide, Inches(1.1), Inches(5.65), Inches(11.0), Inches(0.4),
             "LEITFRAGE", size=12, bold=True, color=YELLOW_DK)
    qtxt = add_text(slide, Inches(1.1), Inches(6.0), Inches(11.0), Inches(0.9),
                    "„Inwiefern beeinflussen strukturelle Unterschiede im deutschen Schulsystem das Wohlbefinden von Schülerinnen und Schülern?“",
                    size=17, bold=True, color=DARK, line_spacing=1.25)

    # balance illustration on right (labels below pans, away from title)
    draw_balance(slide, Inches(10.6), Inches(3.95), scale=0.8)
    add_text(slide, Inches(7.5), Inches(2.85), Inches(2.4), Inches(0.4), "BILDUNG", size=12, bold=True, color=PINK_DK, align=PP_ALIGN.CENTER)
    add_text(slide, Inches(11.0), Inches(2.85), Inches(2.4), Inches(0.4), "WOHLBEFINDEN", size=12, bold=True, color=YELLOW_DK, align=PP_ALIGN.CENTER)

    add_fade_build(slide, [shape_id(s_eyebrow), shape_id(s_h), shape_id(s_bullets), shape_id(qbox), shape_id(qtxt)])
    set_transition(slide, "zoom")

    set_notes(slide, """
Bevor wir tief ins Thema einsteigen, kurz die Frage: Warum lohnt sich dieser Blick auf die Schule überhaupt?
Wir verbringen einen riesigen Teil unseres Tages in der Schule. Sie ist nicht nur Lernort, sondern auch sozialer Lebensraum, in dem Freundschaften entstehen, Selbstbild geformt wird und Zukunftsperspektiven gesetzt werden.
Gleichzeitig nimmt der Leistungsdruck in den letzten Jahren spürbar zu. Bildung entscheidet darüber, welche Berufe und Studienwege offenstehen – und damit auch über gesellschaftliche Teilhabe.
Studien wie die der WHO und des Robert Koch-Instituts zeigen außerdem: Wohlbefinden ist kein Beiwerk, sondern eine zentrale Voraussetzung dafür, dass Lernen überhaupt gut funktioniert.
Daraus ergibt sich die Leitfrage meiner Facharbeit, die ihr hier in der gelben Box seht: Inwiefern beeinflussen strukturelle Unterschiede im deutschen Schulsystem das Wohlbefinden von Schülerinnen und Schülern?
Um diese Frage beantworten zu können, brauchen wir zunächst ein paar Grundlagen.
""")

    # ------------------------------------------------------------------- 3
    slide = prs.slides.add_slide(blank)
    add_background(slide, WHITE)
    add_decor_blob(slide, Inches(-3), Inches(-3), Inches(6), Inches(6), YELLOW, alpha=25)
    add_decor_blob(slide, Inches(11), Inches(6), Inches(5), Inches(5), PINK, alpha=20)

    s_eyebrow = add_text(slide, Inches(0.8), Inches(0.7), Inches(8), Inches(0.4),
                         "02  ·  GRUNDLAGEN", size=13, bold=True, color=YELLOW_DK)
    s_h = add_text(slide, Inches(0.8), Inches(1.1), Inches(11), Inches(1.0),
                   "Drei Begriffe als Kompass",
                   size=38, bold=True, color=DARK)

    # central node
    central_size = Inches(1.5)
    central_x = (SLIDE_W - central_size) / 2
    central_y = Inches(3.9)
    central = add_oval(slide, central_x, central_y, central_size, central_size,
                       fill=DARK, shadow=True)
    add_text(slide, central_x, central_y, central_size, central_size,
             "Schule\n&\nWohlbefinden", size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)

    def branch(x, y, color_bg, color_accent, title, items, w=Inches(3.5), h=Inches(2.5)):
        card = add_rect(slide, x, y, w, h, fill=color_bg, radius=0.10, shadow=True)
        add_text(slide, x + Inches(0.25), y + Inches(0.2), w - Inches(0.5), Inches(0.5),
                 title, size=17, bold=True, color=color_accent)
        bullets = add_bullets(slide, x + Inches(0.25), y + Inches(0.85),
                              w - Inches(0.5), h - Inches(1.0),
                              items, size=13, bullet_color=color_accent)
        return card, bullets

    c1, b1 = branch(Inches(0.6), Inches(3.0), BLUE, BLUE_DK,
                    "Strukturmerkmale",
                    ["Schulformen", "Notensystem", "Versetzungen", "Abschlüsse"])
    c2, b2 = branch(Inches(4.9), Inches(5.2), PINK, PINK_DK,
                    "Wohlbefinden",
                    ["psychische Gesundheit", "soziale Beziehungen", "Lebenszufriedenheit"],
                    w=Inches(3.5), h=Inches(2.2))
    c3, b3 = branch(Inches(9.2), Inches(3.0), GREEN, GREEN_DK,
                    "Bildungsgerechtigkeit",
                    ["faire Chancen", "unabhängig von sozialer Herkunft"])

    # connector lines (decorative, thin)
    for x, y, w, h in [
        (Inches(4.1), Inches(4.6), Inches(2.0), Inches(0.04)),
        (Inches(7.25), Inches(4.6), Inches(1.95), Inches(0.04)),
        (Inches(6.65), Inches(5.0), Inches(0.04), Inches(0.3)),
    ]:
        add_rect(slide, x, y, w, h, fill=MUTED, radius=0.5, shadow=False)

    add_fade_build(slide, [shape_id(s_eyebrow), shape_id(s_h), shape_id(central),
                           shape_id(c1), shape_id(c2), shape_id(c3)])
    set_transition(slide, "movein")

    set_notes(slide, """
Für die ganze Arbeit nutze ich drei Schlüsselbegriffe – sozusagen als Kompass.
Der erste sind die Strukturmerkmale von Schule: also alles, was das Bildungssystem als Institution prägt. Dazu gehören die verschiedenen Schulformen, das Notensystem, Versetzungsregelungen und die Abschlüsse.
Der zweite Begriff ist Wohlbefinden. Damit meine ich – im Sinne der WHO – nicht nur die Abwesenheit von Krankheit, sondern psychische Gesundheit, soziale Beziehungen in der Klasse und die Lebenszufriedenheit insgesamt.
Der dritte Begriff ist Bildungsgerechtigkeit. Sie beschreibt, ob alle Kinder unabhängig von ihrer sozialen Herkunft faire Chancen auf Bildung haben.
Diese drei Begriffe wirken zusammen: Strukturelle Merkmale entscheiden mit darüber, wie Wohlbefinden und Bildungsgerechtigkeit erlebt werden. Schauen wir uns deshalb das deutsche Schulsystem konkret an.
""")

    # ------------------------------------------------------------------- 4
    slide = prs.slides.add_slide(blank)
    add_background(slide, WHITE)
    add_decor_blob(slide, Inches(11), Inches(-2.5), Inches(5), Inches(5), BLUE, alpha=22)
    add_decor_blob(slide, Inches(-2.5), Inches(5.5), Inches(5), Inches(5), PINK, alpha=18)

    s_eyebrow = add_text(slide, Inches(0.8), Inches(0.7), Inches(8), Inches(0.4),
                         "03  ·  ÜBERBLICK", size=13, bold=True, color=BLUE_DK)
    s_h = add_text(slide, Inches(0.8), Inches(1.1), Inches(9), Inches(1.2),
                   "Das deutsche Schulsystem", size=42, bold=True, color=DARK)
    s_sub = add_text(slide, Inches(0.8), Inches(1.95), Inches(11), Inches(0.5),
                     "Fünf Schulformen – fünf Wege",
                     size=18, color=GREY)

    schools = [
        ("Gymnasium",      "🎓", "Abitur · Studium\nhohe akademische\nAnforderungen", BLUE,   BLUE_DK),
        ("Realschule",     "📘", "Mittlerer Abschluss\nallgemein +\npraktisch",        YELLOW, YELLOW_DK),
        ("Hauptschule",    "🔧", "berufliche Bildung\nseltener geworden",              PINK,   PINK_DK),
        ("Gesamtschule",   "🤝", "längeres gemein-\nsames Lernen\nmehrere Abschlüsse", GREEN,  GREEN_DK),
        ("Berufliche Schule", "🏭", "duales System\nSchule + Betrieb",                 BLUE,   ACCENT),
    ]
    card_w = Inches(2.3)
    gap = Inches(0.2)
    total = card_w * 5 + gap * 4
    start_x = (SLIDE_W - total) / 2
    top = Inches(3.0)

    card_ids = []
    for i, (name, icon, body, bg, acc) in enumerate(schools):
        x = start_x + (card_w + gap) * i
        card = add_rect(slide, x, top, card_w, Inches(3.5), fill=bg, radius=0.10, shadow=True)
        add_text(slide, x, top + Inches(0.3), card_w, Inches(0.7),
                 icon, size=34, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + Inches(0.1), top + Inches(1.1), card_w - Inches(0.2), Inches(0.6),
                 name, size=15, bold=True, color=acc, align=PP_ALIGN.CENTER)
        # small underline bar
        add_rect(slide, x + card_w / 2 - Inches(0.25), top + Inches(1.65),
                 Inches(0.5), Inches(0.05), fill=acc, radius=0.5, shadow=False)
        add_text(slide, x + Inches(0.15), top + Inches(1.85), card_w - Inches(0.3), Inches(1.5),
                 body, size=11, color=DARK, align=PP_ALIGN.CENTER, line_spacing=1.3)
        card_ids.append(shape_id(card))

    add_fade_build(slide, [shape_id(s_eyebrow), shape_id(s_h), shape_id(s_sub)] + card_ids)
    set_transition(slide, "magicmove")

    set_notes(slide, """
In Deutschland ist das Schulsystem mehrgliedrig aufgebaut – das heißt: Schon nach der Grundschule werden Kinder auf unterschiedliche Schulformen verteilt.
Das Gymnasium ist die Schulform mit den höchsten akademischen Anforderungen. Ziel ist das Abitur und der Weg an die Universität – aber auch der Leistungsdruck ist besonders hoch.
Die Realschule führt in der Regel zum mittleren Schulabschluss. Sie verbindet allgemeine Bildung mit mehr Praxisorientierung als das Gymnasium und ist oft ein Sprungbrett in die berufliche Ausbildung.
Die Hauptschule war traditionell stark auf das Berufsleben ausgerichtet, hat aber an Bedeutung verloren und wird in vielen Bundesländern mit anderen Schulformen zusammengelegt. Hier wird in der Facharbeit auch das Problem der Stigmatisierung angesprochen.
Die Gesamtschule – ein Reformmodell seit den 1970er-Jahren – lässt Schülerinnen und Schüler länger gemeinsam lernen und bietet verschiedene Abschlüsse unter einem Dach.
Und schließlich die Beruflichen Schulen, die mit dem dualen Ausbildungssystem international als deutsches Vorzeigemodell gelten.
Damit ist klar, wie viele Wege es gibt – aber auch, wie früh die Weichen gestellt werden. Das führt uns direkt zum nächsten Punkt: dem Leistungsdruck.
""")

    # ------------------------------------------------------------------- 5
    slide = prs.slides.add_slide(blank)
    add_background(slide, WHITE)
    add_decor_blob(slide, Inches(-3), Inches(5), Inches(6), Inches(6), PINK, alpha=22)
    add_decor_blob(slide, Inches(11), Inches(-2), Inches(5), Inches(5), YELLOW, alpha=22)

    s_eyebrow = add_text(slide, Inches(0.8), Inches(0.6), Inches(8), Inches(0.4),
                         "04  ·  LEISTUNGSDRUCK", size=13, bold=True, color=PINK_DK)
    s_h = add_text(slide, Inches(0.8), Inches(1.0), Inches(11), Inches(1.2),
                   "Leistungsdruck & Wohlbefinden",
                   size=40, bold=True, color=DARK)

    # left card - Leistungsdruck
    left_card = add_rect(slide, Inches(0.8), Inches(2.6), Inches(5.4), Inches(3.2),
                         fill=BLUE, radius=0.10, shadow=True)
    add_text(slide, Inches(1.0), Inches(2.75), Inches(5.0), Inches(0.5),
             "Leistungsdruck", size=22, bold=True, color=BLUE_DK)
    left_b = add_bullets(slide, Inches(1.1), Inches(3.45), Inches(5.0), Inches(2.3),
                         ["Noten", "Prüfungen", "Konkurrenz", "Erwartungen"],
                         size=17, bullet_color=BLUE_DK)

    # right card - Folgen
    right_card = add_rect(slide, Inches(7.0), Inches(2.6), Inches(5.4), Inches(3.2),
                          fill=PINK, radius=0.10, shadow=True)
    add_text(slide, Inches(7.2), Inches(2.75), Inches(5.0), Inches(0.5),
             "Folgen", size=22, bold=True, color=PINK_DK)
    right_b = add_bullets(slide, Inches(7.3), Inches(3.45), Inches(5.0), Inches(2.3),
                          ["Stress", "Schlafprobleme", "Unsicherheit",
                           "geringeres Selbstwertgefühl"],
                          size=17, bullet_color=PINK_DK)

    # arrow between cards
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.25), Inches(3.95), Inches(0.7), Inches(0.5))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = DARK
    arrow.line.fill.background()
    _remove_shadow(arrow)

    # bottom merksatz
    merk_box = add_rect(slide, Inches(0.8), Inches(6.1), Inches(11.6), Inches(0.95),
                        fill=YELLOW, radius=0.20, shadow=True)
    merk_txt = add_text(slide, Inches(1.0), Inches(6.15), Inches(11.2), Inches(0.85),
                        "„Leistungsdruck kann Motivation fördern, aber auch das Wohlbefinden beeinträchtigen.“",
                        size=18, bold=True, color=DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_fade_build(slide, [shape_id(s_eyebrow), shape_id(s_h), shape_id(left_card),
                           shape_id(left_b), shape_id(arrow),
                           shape_id(right_card), shape_id(right_b),
                           shape_id(merk_box), shape_id(merk_txt)])
    set_transition(slide, "scale")

    set_notes(slide, """
Leistungsdruck gehört zu den wichtigsten Belastungsfaktoren, die schulisches Wohlbefinden beeinflussen.
Links seht ihr die Quellen dieses Drucks: Noten entscheiden über Versetzungen und Abschlüsse, Prüfungen schaffen ständige Bewertungssituationen, der Vergleich mit Mitschülern erzeugt Konkurrenz, und auch die Erwartungen von Eltern, Lehrkräften und der Gesellschaft spielen eine Rolle.
Auf der rechten Seite seht ihr, was daraus werden kann: Stress, Schlafprobleme, Unsicherheit und ein geringeres Selbstwertgefühl. Die KiGGS-Studie des Robert Koch-Instituts und die HBSC-Studie der WHO zeigen genau diese Zusammenhänge: schulischer Stress gehört zu den häufigsten Belastungsfaktoren im Jugendalter.
Wichtig ist aber: Leistungsdruck ist nicht per se schlecht. Er kann Motivation fördern – und wirkt für jede Person unterschiedlich. Während die einen ihn als Ansporn erleben, ist er für andere eine echte psychische Belastung.
Das ist auch der Kernsatz unten in Gelb. Festhalten lässt sich: Druck ist ambivalent – er hat zwei Seiten. Genauso wichtig ist deshalb das Schulklima, zu dem wir jetzt kommen.
""")

    # ------------------------------------------------------------------- 6
    slide = prs.slides.add_slide(blank)
    add_background(slide, WHITE)
    add_decor_blob(slide, Inches(10), Inches(-2.5), Inches(5), Inches(5), GREEN, alpha=25)
    add_decor_blob(slide, Inches(-2), Inches(5.5), Inches(4), Inches(4), BLUE, alpha=18)

    s_eyebrow = add_text(slide, Inches(0.8), Inches(0.7), Inches(8), Inches(0.4),
                         "05  ·  SCHULKLIMA", size=13, bold=True, color=GREEN_DK)
    s_h = add_text(slide, Inches(0.8), Inches(1.1), Inches(11), Inches(1.2),
                   "Schulklima trägt das Lernen",
                   size=40, bold=True, color=DARK)

    cards = [
        ("Lehrer-Schüler-\nVerhältnis", "👩‍🏫",
         ["respektvoll", "vertrauensvoll", "Unterstützung"],
         BLUE, BLUE_DK),
        ("Mitbestimmung", "🗣️",
         ["SV & Klassenrat", "Projekte mitgestalten", "Selbstwirksamkeit"],
         PINK, PINK_DK),
        ("Feedbackkultur", "💬",
         ["Noten + Worte", "formatives Lernen", "individuelle Stärken"],
         YELLOW, YELLOW_DK),
    ]
    card_w = Inches(3.8)
    gap = Inches(0.25)
    total = card_w * 3 + gap * 2
    start_x = (SLIDE_W - total) / 2
    top = Inches(2.7)

    card_ids = []
    for i, (name, icon, items, bg, acc) in enumerate(cards):
        x = start_x + (card_w + gap) * i
        card = add_rect(slide, x, top, card_w, Inches(4.1), fill=bg, radius=0.08, shadow=True)
        add_text(slide, x, top + Inches(0.35), card_w, Inches(0.8),
                 icon, size=44, align=PP_ALIGN.CENTER)
        add_text(slide, x + Inches(0.3), top + Inches(1.4), card_w - Inches(0.6), Inches(1.1),
                 name, size=18, bold=True, color=acc, align=PP_ALIGN.CENTER, line_spacing=1.1)
        add_bullets(slide, x + Inches(0.5), top + Inches(2.7), card_w - Inches(0.9), Inches(1.3),
                    items, size=13, bullet_color=acc)
        card_ids.append(shape_id(card))

    add_fade_build(slide, [shape_id(s_eyebrow), shape_id(s_h)] + card_ids)
    set_transition(slide, "push")

    set_notes(slide, """
Neben dem Druck spielt das Schulklima eine entscheidende Rolle dafür, wie wohl wir uns in der Schule fühlen. Drei Aspekte stehen dabei besonders im Vordergrund.
Erstens: das Lehrer-Schüler-Verhältnis. Lehrkräfte sind nicht nur Wissensvermittler, sondern auch Bezugspersonen. Wer sich ernst genommen und unterstützt fühlt, lernt motivierter und ist weniger gestresst. Umgekehrt können ungerechte Behandlung oder fehlende Unterstützung das Wohlbefinden deutlich verschlechtern.
Zweitens: die Mitbestimmung. Wenn Schülerinnen und Schüler über Schülervertretungen, Klassenräte oder Projekte echte Einflussmöglichkeiten haben, stärkt das ihr Gefühl von Selbstwirksamkeit – also das Gefühl, etwas bewirken zu können.
Drittens: die Feedbackkultur. Noten allein sind oft zu wenig. Qualitative Rückmeldungen, die individuelle Stärken benennen und Lernprozesse begleiten, sind besonders wertvoll. In der Facharbeit wird daher das Konzept der formativen Bewertung hervorgehoben.
All das zusammen ergibt ein Schulklima, das nicht nur das Wohlbefinden, sondern auch die Lernergebnisse verbessert. Bevor wir den Blick ins Ausland richten, möchte ich aber kurz euch einbinden.
""")

    # ------------------------------------------------------------------- 7
    slide = prs.slides.add_slide(blank)
    add_background(slide, WHITE)
    add_decor_blob(slide, Inches(-2), Inches(-2), Inches(5), Inches(5), YELLOW, alpha=22)
    add_decor_blob(slide, Inches(11), Inches(5.5), Inches(4), Inches(4), BLUE, alpha=22)

    s_eyebrow = add_text(slide, Inches(0.8), Inches(0.7), Inches(8), Inches(0.4),
                         "06  ·  INTERAKTION", size=13, bold=True, color=ACCENT)
    s_h = add_text(slide, Inches(0.8), Inches(1.1), Inches(11), Inches(1.0),
                   "Kurze Umfrage", size=42, bold=True, color=DARK)

    qbox = add_rect(slide, Inches(1.5), Inches(2.2), Inches(10.3), Inches(1.3),
                    fill=WHITE, line=MUTED, radius=0.18, shadow=True)
    qtxt = add_text(slide, Inches(1.5), Inches(2.2), Inches(10.3), Inches(1.3),
                    "„Was beeinflusst euer Wohlbefinden in der Schule am stärksten?“",
                    size=22, bold=True, color=DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    options = [
        ("📚", "Notendruck",  YELLOW, YELLOW_DK),
        ("👩‍🏫", "Lehrkräfte", BLUE,   BLUE_DK),
        ("👥", "Mitschüler",  PINK,   PINK_DK),
        ("🕒", "Zeitstress",  GREEN,  GREEN_DK),
    ]
    btn_w = Inches(2.6)
    gap = Inches(0.25)
    total = btn_w * 4 + gap * 3
    start_x = (SLIDE_W - total) / 2
    btn_y = Inches(4.0)
    btn_ids = []
    for i, (icon, label, bg, acc) in enumerate(options):
        x = start_x + (btn_w + gap) * i
        btn = add_rect(slide, x, btn_y, btn_w, Inches(1.7), fill=bg, radius=0.18, shadow=True)
        add_text(slide, x, btn_y + Inches(0.25), btn_w, Inches(0.6),
                 icon, size=30, align=PP_ALIGN.CENTER)
        add_text(slide, x, btn_y + Inches(1.0), btn_w, Inches(0.6),
                 label, size=16, bold=True, color=acc,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        btn_ids.append(shape_id(btn))

    handzeichen = add_text(slide, Inches(0.8), Inches(5.95), Inches(11.6), Inches(0.5),
                           "→ Kurzes Handzeichen, bitte!",
                           size=16, color=GREY, align=PP_ALIGN.CENTER)

    # animated reveal text at bottom (last to appear)
    reveal_box = add_rect(slide, Inches(1.0), Inches(6.55), Inches(11.3), Inches(0.8),
                          fill=DARK, radius=0.45, shadow=True)
    reveal_txt = add_text(slide, Inches(1.0), Inches(6.55), Inches(11.3), Inches(0.8),
                          "Laut Facharbeit spielen ALLE diese Faktoren eine wichtige Rolle.",
                          size=16, bold=True, color=WHITE,
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_fade_build(slide, [shape_id(s_eyebrow), shape_id(s_h), shape_id(qbox), shape_id(qtxt)]
                   + btn_ids + [shape_id(handzeichen), shape_id(reveal_box), shape_id(reveal_txt)])
    set_transition(slide, "dissolve")

    set_notes(slide, """
Jetzt seid ihr kurz dran. Schaut auf die vier Karten – Notendruck, Lehrkräfte, Mitschüler und Zeitstress.
Bitte hebt einmal kurz die Hand bei der Antwortmöglichkeit, die euer eigenes Wohlbefinden in der Schule am stärksten beeinflusst. Ich zähle einmal kurz mit.
…
Spannend! Was meine Facharbeit zeigt: Es gibt nicht die EINE Ursache. Genau dieser Mix – Noten und Prüfungen, Beziehungen zu Lehrkräften und Mitschülerinnen sowie der Zeitdruck im Alltag – wirkt zusammen.
Genau deshalb ist es so wichtig, beim Thema Wohlbefinden nicht nur einen Faktor zu betrachten, sondern die strukturellen Rahmenbedingungen insgesamt.
Werfen wir nach diesem kleinen Stimmungsbild den Blick über die Grenzen Deutschlands hinaus.
""")

    # ------------------------------------------------------------------- 8
    slide = prs.slides.add_slide(blank)
    add_background(slide, WHITE)
    add_decor_blob(slide, Inches(-3), Inches(5), Inches(6), Inches(6), BLUE, alpha=18)
    add_decor_blob(slide, Inches(10), Inches(-2), Inches(5), Inches(5), PINK, alpha=18)

    s_eyebrow = add_text(slide, Inches(0.8), Inches(0.7), Inches(8), Inches(0.4),
                         "07  ·  INTERNATIONAL", size=13, bold=True, color=BLUE_DK)
    s_h = add_text(slide, Inches(0.8), Inches(1.1), Inches(11), Inches(1.2),
                   "Andere Länder, andere Wege",
                   size=40, bold=True, color=DARK)

    countries = [
        ("🇩🇪", "Deutschland",
         ["frühe Aufteilung", "viele Noten"], BLUE, BLUE_DK),
        ("🇫🇮", "Finnland",
         ["länger gemeinsam", "spätere Differenzierung"], PINK, PINK_DK),
        ("🇬🇧 🇺🇸", "Angelsächsisch",
         ["mehr Praxis", "mehr Wahlmöglichkeiten"], YELLOW, YELLOW_DK),
    ]
    card_w = Inches(3.8)
    gap = Inches(0.25)
    total = card_w * 3 + gap * 2
    start_x = (SLIDE_W - total) / 2
    top = Inches(2.7)
    card_ids = []
    for i, (flag, name, items, bg, acc) in enumerate(countries):
        x = start_x + (card_w + gap) * i
        card = add_rect(slide, x, top, card_w, Inches(3.6), fill=bg, radius=0.08, shadow=True)
        add_text(slide, x, top + Inches(0.35), card_w, Inches(0.8),
                 flag, size=44, align=PP_ALIGN.CENTER)
        add_text(slide, x + Inches(0.3), top + Inches(1.4), card_w - Inches(0.6), Inches(0.6),
                 name, size=20, bold=True, color=acc, align=PP_ALIGN.CENTER)
        # divider
        add_rect(slide, x + card_w/2 - Inches(0.3), top + Inches(2.05),
                 Inches(0.6), Inches(0.05), fill=acc, radius=0.5, shadow=False)
        add_bullets(slide, x + Inches(0.5), top + Inches(2.25), card_w - Inches(1.0), Inches(1.3),
                    items, size=14, bullet_color=acc)
        card_ids.append(shape_id(card))

    foot_box = add_rect(slide, Inches(0.8), Inches(6.55), Inches(11.6), Inches(0.7),
                        fill=DARK, radius=0.4, shadow=False)
    foot_txt = add_text(slide, Inches(0.8), Inches(6.55), Inches(11.6), Inches(0.7),
                        "Kein System ist perfekt – jedes Land setzt eigene Schwerpunkte.",
                        size=15, bold=True, color=WHITE,
                        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_fade_build(slide, [shape_id(s_eyebrow), shape_id(s_h)] + card_ids
                   + [shape_id(foot_box), shape_id(foot_txt)])
    set_transition(slide, "pop")

    set_notes(slide, """
Im internationalen Vergleich werden die Besonderheiten unseres Systems erst richtig sichtbar.
In Deutschland erfolgt die Aufteilung in unterschiedliche Schulformen sehr früh – meist nach der vierten Klasse. Noten spielen eine zentrale Rolle und gibt es vergleichsweise viele.
Finnland geht einen anderen Weg: Dort lernen Kinder bis zur neunten Klasse gemeinsam. Leistungsdifferenzierung erfolgt also viel später, und in den ersten Schuljahren wird seltener standardisiert benotet. Stattdessen gibt es mehr qualitative Rückmeldung. Finnland gilt zudem deshalb als Vorzeigesystem, weil Lehrkräfte ein hohes gesellschaftliches Ansehen genießen.
In angelsächsischen Systemen, etwa in Großbritannien oder den USA, lernen Kinder ebenfalls länger gemeinsam in Gesamtschulen, können dafür aber mehr eigene Schwerpunkte und Wahlfächer setzen. Projektorientiertes und praktisches Arbeiten spielt eine größere Rolle – allerdings unterscheiden sich die einzelnen Schulen oft stark.
Wichtig: Keines dieser Systeme ist perfekt. Jedes muss unterschiedliche gesellschaftliche Anforderungen erfüllen. Aber genau dieser Vergleich liefert wertvolle Ideen für die Stärken- und Schwächen-Analyse, die wir uns jetzt anschauen.
""")

    # ------------------------------------------------------------------- 9
    slide = prs.slides.add_slide(blank)
    add_background(slide, WHITE)
    add_decor_blob(slide, Inches(11), Inches(5.5), Inches(4), Inches(4), YELLOW, alpha=22)
    add_decor_blob(slide, Inches(-2), Inches(-2), Inches(4.5), Inches(4.5), BLUE, alpha=20)

    s_eyebrow = add_text(slide, Inches(0.8), Inches(0.7), Inches(8), Inches(0.4),
                         "08  ·  KRITISCHE BEWERTUNG", size=13, bold=True, color=YELLOW_DK)
    s_h = add_text(slide, Inches(0.8), Inches(1.1), Inches(11), Inches(1.2),
                   "Stärken & Schwächen",
                   size=40, bold=True, color=DARK)

    # left green card
    lcard = add_rect(slide, Inches(0.9), Inches(2.7), Inches(5.7), Inches(4.4),
                     fill=GREEN, radius=0.10, shadow=True)
    add_text(slide, Inches(1.2), Inches(2.85), Inches(5.0), Inches(0.5),
             "✓  Stärken", size=24, bold=True, color=GREEN_DK)
    lb = add_bullets(slide, Inches(1.3), Inches(3.65), Inches(5.2), Inches(3.0),
                     [
                         "verschiedene Bildungswege",
                         "duales Ausbildungssystem",
                         "hohe akademische Standards",
                     ],
                     size=18, bullet_color=GREEN_DK, line_spacing=1.6)

    # right pink card
    rcard = add_rect(slide, Inches(6.8), Inches(2.7), Inches(5.7), Inches(4.4),
                     fill=PINK, radius=0.10, shadow=True)
    add_text(slide, Inches(7.1), Inches(2.85), Inches(5.0), Inches(0.5),
             "✗  Schwächen", size=24, bold=True, color=PINK_DK)
    rb = add_bullets(slide, Inches(7.2), Inches(3.65), Inches(5.2), Inches(3.0),
                     [
                         "frühe Selektion",
                         "Leistungsdruck",
                         "soziale Ungleichheiten",
                     ],
                     size=18, bullet_color=PINK_DK, line_spacing=1.6)

    add_fade_build(slide, [shape_id(s_eyebrow), shape_id(s_h),
                           shape_id(lcard), shape_id(lb),
                           shape_id(rcard), shape_id(rb)])
    set_transition(slide, "slide")

    set_notes(slide, """
Aus all dem ergibt sich ein recht klares Bild – mit zwei Seiten.
Zu den Stärken gehört zunächst die Vielfalt an Bildungswegen. Verschiedene Schulformen ermöglichen, dass Lernende auf unterschiedlichem Niveau gezielt gefördert werden. Besonders hervorzuheben ist das duale Ausbildungssystem, das international als Modell gilt und maßgeblich dazu beiträgt, dass Deutschland eine vergleichsweise niedrige Jugendarbeitslosigkeit hat. Zudem bietet besonders das Gymnasium hohe akademische Standards, die gut auf ein Studium vorbereiten.
Auf der anderen Seite stehen klare Schwächen. Die frühe Selektion nach der vierten Klasse legt Bildungswege zu einem Zeitpunkt fest, an dem Leistungen noch stark von familiärem Hintergrund abhängen. Der Leistungsdruck durch Noten, Prüfungen und Versetzungsentscheidungen kann erheblichen Stress erzeugen. Und schließlich ist der Zusammenhang zwischen sozialer Herkunft und Bildungserfolg in Deutschland besonders ausgeprägt – das zeigen OECD-Studien immer wieder.
Genau deshalb wird über Reformen diskutiert – und einige davon stelle ich euch jetzt vor.
""")

    # ------------------------------------------------------------------- 10
    slide = prs.slides.add_slide(blank)
    add_background(slide, WHITE)
    add_decor_blob(slide, Inches(-2), Inches(-2), Inches(5), Inches(5), GREEN, alpha=22)
    add_decor_blob(slide, Inches(10), Inches(5.5), Inches(4.5), Inches(4.5), PINK, alpha=20)

    s_eyebrow = add_text(slide, Inches(0.8), Inches(0.7), Inches(8), Inches(0.4),
                         "09  ·  REFORMANSÄTZE", size=13, bold=True, color=GREEN_DK)
    s_h = add_text(slide, Inches(0.8), Inches(1.1), Inches(11), Inches(1.2),
                   "Schritte in die Zukunft",
                   size=40, bold=True, color=DARK)

    circles = [
        ("Lebens-\nkompetenzen", "Finanzen, Medien, Alltagskompetenzen stärker im Lehrplan",
         BLUE, BLUE_DK),
        ("Spätere\nDifferenzierung", "Länger gemeinsam lernen – weniger Druck früh, fundiertere Entscheidungen",
         PINK, PINK_DK),
        ("Alternative\nBewertung", "Portfolios, Feedback und Projekte ergänzen klassische Noten",
         YELLOW, YELLOW_DK),
    ]
    circle_d = Inches(2.5)
    gap = Inches(1.0)
    total = circle_d * 3 + gap * 2
    start_x = (SLIDE_W - total) / 2
    top = Inches(2.9)

    ids = []
    text_ids = []
    for i, (label, body, bg, acc) in enumerate(circles):
        x = start_x + (circle_d + gap) * i
        c = add_oval(slide, x, top, circle_d, circle_d, fill=bg, shadow=True)
        add_text(slide, x, top, circle_d, circle_d, label,
                 size=18, bold=True, color=acc,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
        t = add_text(slide, x - Inches(0.3), top + circle_d + Inches(0.2),
                     circle_d + Inches(0.6), Inches(1.7),
                     body, size=13, color=DARK,
                     align=PP_ALIGN.CENTER, line_spacing=1.35)
        ids.append(shape_id(c))
        text_ids.append(shape_id(t))

    # arrow towards future
    arrow_shape = slide.shapes.add_shape(
        MSO_SHAPE.PENTAGON,
        Inches(11.0), Inches(3.85), Inches(1.2), Inches(0.7),
    )
    arrow_shape.fill.solid()
    arrow_shape.fill.fore_color.rgb = DARK
    arrow_shape.line.fill.background()
    _remove_shadow(arrow_shape)
    add_text(slide, Inches(11.0), Inches(3.85), Inches(1.2), Inches(0.7),
             "Zukunft", size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_fade_build(slide, [shape_id(s_eyebrow), shape_id(s_h)] + ids + text_ids + [shape_id(arrow_shape)])
    set_transition(slide, "rotate")

    set_notes(slide, """
Welche Ideen werden also diskutiert, um Bildung und Wohlbefinden besser zu verbinden? Drei zentrale Reformansätze stehen in der Facharbeit im Mittelpunkt.
Erstens: Mehr Lebenskompetenzen im Curriculum. Konkret heißt das, Themen wie Finanzbildung, Medienkompetenz und praktische Alltagsfähigkeiten stärker zu integrieren. Auch die OECD betont die wachsende Bedeutung von Problemlösung, kritischem Denken und Teamarbeit.
Zweitens: Eine spätere Leistungsdifferenzierung – also nicht schon nach der vierten Klasse die Weichen zu stellen. Skandinavische Systeme zeigen, dass längeres gemeinsames Lernen soziale Ungleichheiten reduzieren und Bildungsentscheidungen fundierter machen kann.
Drittens: Alternative Bewertungsformen. Neben Noten könnten Portfolios, projektbasierte Bewertungen und ausführliche Lernrückmeldungen den Lernprozess stärker in den Mittelpunkt rücken – statt nur die Prüfungsergebnisse.
Diese drei Pfeile zeigen gemeinsam in Richtung einer Schule der Zukunft: eine Schule, die mehr ist als Leistungsbewertung. Damit komme ich zu meinem Fazit.
""")

    # ------------------------------------------------------------------- 11
    slide = prs.slides.add_slide(blank)
    add_background(slide, WHITE)
    add_decor_blob(slide, Inches(-3), Inches(4), Inches(7), Inches(6), BLUE, alpha=22)
    add_decor_blob(slide, Inches(10), Inches(-2), Inches(6), Inches(5), PINK, alpha=18)
    add_decor_blob(slide, Inches(11.5), Inches(5.5), Inches(4), Inches(4), YELLOW, alpha=30)

    s_eyebrow = add_text(slide, Inches(0.8), Inches(0.7), Inches(8), Inches(0.4),
                         "10  ·  FAZIT", size=13, bold=True, color=ACCENT)
    s_h = add_text(slide, Inches(0.8), Inches(1.1), Inches(11.5), Inches(1.6),
                   "Schule beeinflusst mehr als nur Noten.",
                   size=40, bold=True, color=DARK, line_spacing=1.05)

    items = [
        ("✓", "Strukturen wirken auf das Wohlbefinden.",            BLUE,   BLUE_DK),
        ("✓", "Leistungsdruck und Schulklima sind zentral.",         PINK,   PINK_DK),
        ("✓", "Reformen können Bildung und Glück verbinden.",        YELLOW, YELLOW_DK),
    ]
    ids = []
    for i, (mark, txt, bg, acc) in enumerate(items):
        y = Inches(3.1 + i * 0.95)
        card = add_rect(slide, Inches(0.8), y, Inches(8.6), Inches(0.78),
                        fill=bg, radius=0.45, shadow=True)
        add_text(slide, Inches(1.0), y, Inches(0.8), Inches(0.78),
                 mark, size=22, bold=True, color=acc,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(1.8), y, Inches(7.0), Inches(0.78),
                 txt, size=17, color=DARK,
                 align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        ids.append(shape_id(card))

    thank_box = add_rect(slide, Inches(0.8), Inches(6.4), Inches(8.6), Inches(0.85),
                         fill=DARK, radius=0.4, shadow=True)
    thank_txt = add_text(slide, Inches(0.8), Inches(6.4), Inches(8.6), Inches(0.85),
                         "Danke für eure Aufmerksamkeit ✨",
                         size=22, bold=True, color=WHITE,
                         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # right-side pastel illustration (compact, kept right of the cards)
    draw_school(slide, Inches(11.4), Inches(4.6), scale=0.65)
    draw_student(slide, Inches(10.3), Inches(5.55), scale=0.55, face=PINK, body=BLUE, mood="happy")
    draw_student(slide, Inches(12.3), Inches(5.55), scale=0.55, face=YELLOW, body=PINK_DK, mood="happy")

    add_fade_build(slide, [shape_id(s_eyebrow), shape_id(s_h)] + ids + [shape_id(thank_box), shape_id(thank_txt)])
    set_transition(slide, "fade")

    set_notes(slide, """
Damit zum Schluss meine drei Kernaussagen.
Erstens: Schulische Strukturen wirken sich messbar auf das Wohlbefinden der Schülerinnen und Schüler aus. Die Wahl der Schulform, das Bewertungssystem und die Versetzungsregelungen sind keine reine Organisationsfrage – sie beeinflussen, wie wohl wir uns fühlen.
Zweitens: Leistungsdruck und Schulklima sind zentrale Faktoren. Druck kann motivieren, aber auch belasten – und ein gutes Schulklima mit unterstützenden Lehrkräften, Mitbestimmung und qualitativem Feedback kann viel auffangen.
Drittens: Reformen wie spätere Differenzierung, mehr Lebenskompetenzen und alternative Bewertungsformen können Bildung und Glück enger miteinander verbinden.
Damit komme ich zurück zur Leitfrage: Strukturen im deutschen Schulsystem beeinflussen das Wohlbefinden ganz entscheidend – sowohl positiv durch differenzierte Wege und hohe Standards, als auch belastend durch frühe Selektion und Leistungsdruck.
Ich danke euch herzlich für eure Aufmerksamkeit und freue mich jetzt auf eure Fragen und Anmerkungen!
""")

    # Save
    out = Path(__file__).parent / "Wie_Schule_wirkt_Praesentation.pptx"
    prs.save(out)
    print(f"Saved: {out}  (slides={len(prs.slides)})")


if __name__ == "__main__":
    build()
