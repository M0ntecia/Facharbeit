"""
Generator für die Keynote-Style-Präsentation:
"Wie Schule wirkt – Strukturelle Unterschiede und ihr Effekt auf das Wohlbefinden
von Schülerinnen und Schülern" (Seminarfachkurs Glück)

Erzeugt eine .pptx-Datei, die in Apple Keynote auf dem iPad importiert werden kann.
Inhalte basieren ausschließlich auf der vorliegenden Facharbeit.
"""
from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn, nsmap
from lxml import etree

# ---------------------------------------------------------------------------
# Designsystem
# ---------------------------------------------------------------------------
BLUE       = RGBColor(0xCF, 0xEF, 0xFF)   # #CFEFFF Hellblau
PINK       = RGBColor(0xFF, 0xD9, 0xE8)   # #FFD9E8 Pastellrosa
YELLOW     = RGBColor(0xFF, 0xF7, 0xC7)   # #FFF7C7 Hellgelb
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
DARK       = RGBColor(0x2C, 0x2C, 0x2E)   # Schrift dunkelgrau
DARK_SOFT  = RGBColor(0x4A, 0x4A, 0x4F)   # Sekundärschrift
ACCENT     = RGBColor(0x4F, 0x8F, 0xBF)   # Tiefes Akzentblau (sparsam)
GREEN_SOFT = RGBColor(0xCD, 0xEF, 0xD9)   # für Stärken-Karte
PINK_SOFT  = RGBColor(0xFF, 0xC8, 0xD9)   # für Schwächen-Karte
SHADOW_GRY = RGBColor(0xE6, 0xE6, 0xEC)
LINE_GREY  = RGBColor(0xD8, 0xD8, 0xDE)

# Folienformat 16:9
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

FONT_HEAD = "Helvetica Neue"  # Keynote-typisch; Fallback auf System
FONT_BODY = "Helvetica Neue"


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def set_solid_fill(shape, color: RGBColor):
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = color


def set_no_line(shape):
    shape.line.fill.background()


def set_line(shape, color: RGBColor, width_pt: float = 0.75):
    line = shape.line
    line.color.rgb = color
    line.width = Pt(width_pt)


def add_background(slide, color: RGBColor = WHITE):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    set_solid_fill(bg, color)
    set_no_line(bg)
    bg.shadow.inherit = False
    return bg


def add_soft_blob(slide, x, y, w, h, color: RGBColor, shape=MSO_SHAPE.OVAL, alpha=None):
    blob = slide.shapes.add_shape(shape, x, y, w, h)
    set_solid_fill(blob, color)
    set_no_line(blob)
    if alpha is not None:
        # Transparenz auf Füllung setzen (0..100000)
        sp = blob.fill.fore_color._xFill
        srgb = sp.find(qn("a:srgbClr"))
        if srgb is not None:
            alpha_el = etree.SubElement(srgb, qn("a:alpha"))
            alpha_el.set("val", str(int(alpha * 1000)))
    return blob


def add_text(
    slide,
    x, y, w, h,
    text: str,
    *,
    size: int = 18,
    bold: bool = False,
    color: RGBColor = DARK,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    font: str = FONT_BODY,
    line_spacing: float | None = None,
):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = anchor

    lines = text.split("\n") if isinstance(text, str) else list(text)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = ln
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
    return tb


def add_rich_textbox(slide, x, y, w, h, *, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    tf.vertical_anchor = anchor
    return tf


def add_paragraph(tf, text, *, size=16, bold=False, color=DARK, align=PP_ALIGN.LEFT,
                  font=FONT_BODY, space_before=0, space_after=4, first=False, line_spacing=1.15):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return p


def add_card(slide, x, y, w, h, fill: RGBColor, *, radius_factor=0.08, line_color: RGBColor | None = None):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    set_solid_fill(card, fill)
    if line_color is None:
        set_no_line(card)
    else:
        set_line(card, line_color, 0.75)
    # Eckenradius über Adjustment (0..0.5)
    try:
        card.adjustments[0] = radius_factor
    except Exception:
        pass
    # leichter Schatten
    apply_soft_shadow(card)
    return card


def apply_soft_shadow(shape, *, blur=14, dist=4, alpha=22):
    """Fügt einen dezenten Schlagschatten via XML hinzu."""
    spPr = shape.fill._xPr  # CT_ShapeProperties
    # remove existing effectLst
    for e in spPr.findall(qn("a:effectLst")):
        spPr.remove(e)
    effectLst = etree.SubElement(spPr, qn("a:effectLst"))
    outerShdw = etree.SubElement(
        effectLst, qn("a:outerShdw"),
        attrib={
            "blurRad": str(blur * 12700),
            "dist": str(dist * 12700),
            "dir": "5400000",
            "rotWithShape": "0",
        },
    )
    srgb = etree.SubElement(outerShdw, qn("a:srgbClr"), attrib={"val": "000000"})
    etree.SubElement(srgb, qn("a:alpha"), attrib={"val": str(alpha * 1000)})


def add_circle_icon(slide, cx, cy, r, fill: RGBColor, glyph: str = "", glyph_color=DARK, glyph_size=24):
    x = cx - r
    y = cy - r
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, r * 2, r * 2)
    set_solid_fill(circ, fill)
    set_no_line(circ)
    apply_soft_shadow(circ, blur=10, dist=2, alpha=18)
    if glyph:
        tb = add_text(slide, x, y, r * 2, r * 2, glyph,
                      size=glyph_size, bold=True, color=glyph_color,
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return circ


def add_line(slide, x1, y1, x2, y2, color=LINE_GREY, width_pt=1.25, dashed=False):
    line = slide.shapes.add_connector(1, x1, y1, x2, y2)
    line.line.color.rgb = color
    line.line.width = Pt(width_pt)
    if dashed:
        ln = line.line._get_or_add_ln()
        prstDash = etree.SubElement(ln, qn("a:prstDash"))
        prstDash.set("val", "dash")
    return line


# ---------------------------------------------------------------------------
# Übergänge (slide transitions) – schreibt p:transition direkt in slide-XML
# ---------------------------------------------------------------------------

P_NS  = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS  = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
P15_NS = "http://schemas.microsoft.com/office/powerpoint/2012/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _add_transition_xml(slide, xml_str: str):
    """Fügt das übergebene <p:transition>-Element zur Slide hinzu."""
    sld = slide._element
    # eventuelle alte transition entfernen
    for t in sld.findall(qn("p:transition")):
        sld.remove(t)
    el = etree.fromstring(xml_str)
    sld.append(el)


def set_transition(slide, name: str):
    """Setzt einen Übergang. 'name' ist ein logischer Bezeichner."""
    nsdecl = (
        f' xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}"'
        f' xmlns:p14="{P14_NS}" xmlns:mc="{MC_NS}"'
    )
    spd = ' spd="med"'
    name_l = name.lower()

    if name_l == "fade":
        xml = f'<p:transition{nsdecl}{spd}><p:fade/></p:transition>'

    elif name_l == "zoom":
        # PowerPoint 2010+ "zoom" via mc:AlternateContent (p14:zoom)
        xml = (
            f'<p:transition{nsdecl}{spd}>'
            f'  <mc:AlternateContent>'
            f'    <mc:Choice xmlns:p14="{P14_NS}" Requires="p14">'
            f'      <p14:zoom/>'
            f'    </mc:Choice>'
            f'    <mc:Fallback><p:fade/></mc:Fallback>'
            f'  </mc:AlternateContent>'
            f'</p:transition>'
        )

    elif name_l in ("move in", "movein", "push"):
        xml = f'<p:transition{nsdecl}{spd}><p:push dir="l"/></p:transition>'

    elif name_l in ("magic move", "magicmove", "morph"):
        # PowerPoint 2016 morph
        xml = (
            f'<p:transition{nsdecl}{spd}>'
            f'  <mc:AlternateContent>'
            f'    <mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
            f'      <p159:morph option="byObject"/>'
            f'    </mc:Choice>'
            f'    <mc:Fallback><p:fade/></mc:Fallback>'
            f'  </mc:AlternateContent>'
            f'</p:transition>'
        )

    elif name_l == "scale":
        # In PPTX kein direktes Pendant – wir nutzen p14:zoom (in=Out) als Annäherung
        xml = (
            f'<p:transition{nsdecl}{spd}>'
            f'  <mc:AlternateContent>'
            f'    <mc:Choice xmlns:p14="{P14_NS}" Requires="p14">'
            f'      <p14:zoom/>'
            f'    </mc:Choice>'
            f'    <mc:Fallback><p:fade/></mc:Fallback>'
            f'  </mc:AlternateContent>'
            f'</p:transition>'
        )

    elif name_l == "dissolve":
        xml = f'<p:transition{nsdecl}{spd}><p:dissolve/></p:transition>'

    elif name_l == "pop":
        # kein PPTX-Pendant – Annäherung mit Zoom/Fade
        xml = (
            f'<p:transition{nsdecl}{spd}>'
            f'  <mc:AlternateContent>'
            f'    <mc:Choice xmlns:p14="{P14_NS}" Requires="p14">'
            f'      <p14:zoom/>'
            f'    </mc:Choice>'
            f'    <mc:Fallback><p:fade/></mc:Fallback>'
            f'  </mc:AlternateContent>'
            f'</p:transition>'
        )

    elif name_l == "slide":
        xml = f'<p:transition{nsdecl}{spd}><p:pull dir="l"/></p:transition>'

    elif name_l == "rotate":
        # Annäherung – PPTX hat keinen 3D-Rotate; verwenden wir „newsflash"
        xml = (
            f'<p:transition{nsdecl}{spd}>'
            f'  <mc:AlternateContent>'
            f'    <mc:Choice xmlns:p14="{P14_NS}" Requires="p14">'
            f'      <p14:newsflash/>'
            f'    </mc:Choice>'
            f'    <mc:Fallback><p:fade/></mc:Fallback>'
            f'  </mc:AlternateContent>'
            f'</p:transition>'
        )

    else:
        xml = f'<p:transition{nsdecl}{spd}><p:fade/></p:transition>'

    _add_transition_xml(slide, xml)


# ---------------------------------------------------------------------------
# Build-In-Animationen für Texte (Fade-In nacheinander)
# ---------------------------------------------------------------------------

def add_sequential_fade_in(slide, shape_ids: list[int]):
    """Fügt eine simple sequentielle Fade-In-Animation für die angegebenen
    shape_ids (spIds) hinzu. Jede Form erscheint per Klick nacheinander."""
    if not shape_ids:
        return
    nsdecl = (
        f' xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}"'
    )
    # Aufbau einer click-basierten Sequenz (wie Standard "On click" in PowerPoint).
    children_xml = []
    for i, spid in enumerate(shape_ids):
        children_xml.append(f'''
<p:par>
  <p:cTn id="{10 + i*10}" fill="hold">
    <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
    <p:childTnLst>
      <p:par>
        <p:cTn id="{11 + i*10}" fill="hold">
          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
          <p:childTnLst>
            <p:par>
              <p:cTn id="{12 + i*10}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">
                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                <p:childTnLst>
                  <p:set>
                    <p:cBhvr>
                      <p:cTn id="{13 + i*10}" dur="1" fill="hold">
                        <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      </p:cTn>
                      <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                      <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                    </p:cBhvr>
                    <p:to><p:strVal val="visible"/></p:to>
                  </p:set>
                  <p:anim calcmode="lin" valueType="num">
                    <p:cBhvr additive="base">
                      <p:cTn id="{14 + i*10}" dur="500" fill="hold"/>
                      <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                      <p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst>
                    </p:cBhvr>
                    <p:tavLst>
                      <p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav>
                      <p:tav tm="100000"><p:val><p:fltVal val="1"/></p:val></p:tav>
                    </p:tavLst>
                  </p:anim>
                </p:childTnLst>
              </p:cTn>
            </p:par>
          </p:childTnLst>
        </p:cTn>
      </p:par>
    </p:childTnLst>
  </p:cTn>
</p:par>''')

    timing_xml = f'''<p:timing{nsdecl}>
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>
                {''.join(children_xml)}
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
    {''.join(f'<p:bldP spid="{sid}" grpId="0"/>' for sid in shape_ids)}
  </p:bldLst>
</p:timing>'''

    sld = slide._element
    for t in sld.findall(qn("p:timing")):
        sld.remove(t)
    sld.append(etree.fromstring(timing_xml))


def shape_id(shape):
    return int(shape._element.nvSpPr.cNvPr.get("id"))


# ---------------------------------------------------------------------------
# Sprechnotizen
# ---------------------------------------------------------------------------

def set_notes(slide, notes_text: str):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.clear()
    paragraphs = notes_text.strip().split("\n\n")
    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = para.strip()
        run.font.size = Pt(13)
        run.font.name = FONT_BODY
        run.font.color.rgb = DARK


# ---------------------------------------------------------------------------
# Wiederkehrende Folienelemente (Header / Footer / Seitenzahl)
# ---------------------------------------------------------------------------

def add_page_meta(slide, page_num: int, total: int = 11, accent: RGBColor = BLUE):
    # kleiner Akzentstreifen unten links
    bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(0.55), Inches(7.05),
                                 Inches(0.45), Inches(0.12))
    set_solid_fill(bar, accent)
    set_no_line(bar)
    try:
        bar.adjustments[0] = 0.5
    except Exception:
        pass
    add_text(slide, Inches(1.05), Inches(6.95), Inches(2.0), Inches(0.3),
             f"{page_num:02d} / {total:02d}",
             size=11, color=DARK_SOFT)
    add_text(slide, Inches(10.7), Inches(6.95), Inches(2.4), Inches(0.3),
             "Seminarfachkurs Glück",
             size=11, color=DARK_SOFT, align=PP_ALIGN.RIGHT)


# ===========================================================================
# Folie 1 – Titelfolie
# ===========================================================================

def slide_01_titel(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    add_background(s, WHITE)

    # Großer dekorativer Pastell-Blob hinten links + rechts
    add_soft_blob(s, Inches(-2.2), Inches(-1.8), Inches(6.5), Inches(6.5), BLUE)
    add_soft_blob(s, Inches(9.2), Inches(4.0), Inches(5.5), Inches(5.5), PINK)
    add_soft_blob(s, Inches(7.0), Inches(-1.4), Inches(3.3), Inches(3.3), YELLOW)

    # Titel-Bereich (links)
    title_tf = add_rich_textbox(s, Inches(0.9), Inches(2.0), Inches(7.6), Inches(3.6))
    add_paragraph(title_tf, "Wie Schule wirkt",
                  size=64, bold=True, color=DARK, font=FONT_HEAD,
                  first=True, line_spacing=1.0, space_after=8)
    add_paragraph(title_tf,
                  "Strukturelle Unterschiede und ihr Effekt",
                  size=28, bold=False, color=DARK_SOFT, line_spacing=1.15, space_after=2)
    add_paragraph(title_tf,
                  "auf das Wohlbefinden von Schüler:innen",
                  size=28, bold=False, color=DARK_SOFT, line_spacing=1.15, space_after=18)

    # Pastell-Pille mit Untertitel
    pill = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(0.9), Inches(5.5), Inches(3.6), Inches(0.55))
    set_solid_fill(pill, YELLOW)
    set_no_line(pill)
    try:
        pill.adjustments[0] = 0.5
    except Exception:
        pass
    add_text(s, Inches(0.9), Inches(5.5), Inches(3.6), Inches(0.55),
             "Seminarfachkurs Glück",
             size=15, bold=True, color=DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Kleine Meta-Zeile
    add_text(s, Inches(0.9), Inches(6.15), Inches(8), Inches(0.4),
             "Svea Timphus  ·  Gymnasium Damme  ·  2026",
             size=13, color=DARK_SOFT)

    # ---- Illustration: Schule + Schüler:innen (rechts) -------------------
    # Schulgebäude
    base_x, base_y = Inches(8.7), Inches(2.05)

    # Dach (Dreieck)
    roof = s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                              Inches(8.55), Inches(1.55),
                              Inches(3.4), Inches(0.95))
    set_solid_fill(roof, PINK)
    set_no_line(roof)
    apply_soft_shadow(roof, blur=12, dist=3, alpha=18)

    # Schulkörper
    body = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              base_x, base_y, Inches(3.1), Inches(2.5))
    set_solid_fill(body, BLUE)
    set_no_line(body)
    try:
        body.adjustments[0] = 0.06
    except Exception:
        pass
    apply_soft_shadow(body, blur=14, dist=4, alpha=20)

    # Glocke / Uhr
    clock = s.shapes.add_shape(MSO_SHAPE.OVAL,
                               Inches(10.0), Inches(1.85),
                               Inches(0.45), Inches(0.45))
    set_solid_fill(clock, WHITE)
    set_no_line(clock)
    add_text(s, Inches(10.0), Inches(1.85), Inches(0.45), Inches(0.45),
             "★", size=14, bold=True, color=ACCENT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Türe
    door = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(9.95), Inches(3.6), Inches(0.55), Inches(0.95))
    set_solid_fill(door, YELLOW)
    set_no_line(door)
    try:
        door.adjustments[0] = 0.18
    except Exception:
        pass

    # Fenster (4 Stück)
    for i, (fx, fy) in enumerate([(8.95, 2.55), (10.95, 2.55),
                                  (8.95, 3.55), (10.95, 3.55)]):
        win = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(fx), Inches(fy),
                                 Inches(0.55), Inches(0.55))
        set_solid_fill(win, WHITE)
        set_no_line(win)
        try:
            win.adjustments[0] = 0.18
        except Exception:
            pass

    # Schüler:innen Köpfe als kleine Avatare unter der Schule
    avatars = [
        (Inches(8.65), Inches(4.85), YELLOW, "☺"),
        (Inches(9.45), Inches(4.85), PINK,   "☺"),
        (Inches(10.25), Inches(4.85), BLUE,  "☹"),
        (Inches(11.05), Inches(4.85), YELLOW, "☺"),
    ]
    for ax, ay, col, gly in avatars:
        a = s.shapes.add_shape(MSO_SHAPE.OVAL, ax, ay, Inches(0.7), Inches(0.7))
        set_solid_fill(a, col)
        set_no_line(a)
        apply_soft_shadow(a, blur=8, dist=2, alpha=18)
        add_text(s, ax, ay, Inches(0.7), Inches(0.7), gly,
                 size=22, bold=True, color=DARK,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Boden / Linie
    add_line(s, Inches(8.4), Inches(5.7), Inches(11.95), Inches(5.7),
             color=LINE_GREY, width_pt=2)

    add_page_meta(s, 1, accent=BLUE)
    set_transition(s, "fade")

    set_notes(s, """
Herzlich willkommen zu meiner Präsentation im Seminarfachkurs Glück. Mein Thema lautet: „Wie Schule wirkt – strukturelle Unterschiede und ihr Effekt auf das Wohlbefinden von Schülerinnen und Schülern."

Schule prägt unseren Alltag wie kaum eine andere Institution. Wir verbringen einen großen Teil unserer Jugend in ihr – und dabei geht es längst nicht nur um Noten und Abschlüsse. Schule ist ein Lebensraum. Sie entscheidet mit, ob wir uns wohlfühlen, ob wir mit Freude lernen, oder ob wir unter Druck geraten.

In den nächsten 15 Minuten möchte ich euch zeigen, wie stark die Strukturen unseres Schulsystems das Wohlbefinden von Schülerinnen und Schülern beeinflussen – und welche Alternativen andere Länder uns vorleben.

Lasst uns also gemeinsam einen Blick darauf werfen, warum dieses Thema gerade im Seminarfachkurs Glück so passend ist.
""")


# ===========================================================================
# Folie 2 – Warum ist das Thema wichtig?
# ===========================================================================

def slide_02_relevanz(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(s, WHITE)
    add_soft_blob(s, Inches(-1.5), Inches(-1.5), Inches(4.5), Inches(4.5), YELLOW)
    add_soft_blob(s, Inches(11.5), Inches(5.0), Inches(3.5), Inches(3.5), BLUE)

    # kleine Section-Marke
    add_text(s, Inches(0.9), Inches(0.55), Inches(6), Inches(0.35),
             "01 · EINSTIEG", size=11, bold=True, color=ACCENT)

    # Headline
    add_text(s, Inches(0.9), Inches(0.95), Inches(11.5), Inches(1.2),
             "Warum ist das Thema wichtig?",
             size=44, bold=True, color=DARK, font=FONT_HEAD)

    # Bullet-Liste links
    tf = add_rich_textbox(s, Inches(0.9), Inches(2.4), Inches(6.5), Inches(3.6))
    bullets = [
        ("●", "Schule prägt unseren Alltag"),
        ("●", "Leistungsdruck nimmt zu"),
        ("●", "Bildung beeinflusst die Zukunft"),
        ("●", "Wohlbefinden ist Teil erfolgreicher Bildung"),
    ]
    for i, (b, t) in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(14)
        p.line_spacing = 1.2
        r1 = p.add_run(); r1.text = b + "  "
        r1.font.size = Pt(20); r1.font.bold = True; r1.font.color.rgb = ACCENT
        r2 = p.add_run(); r2.text = t
        r2.font.size = Pt(20); r2.font.color.rgb = DARK; r2.font.name = FONT_BODY

    # Visual: Waage Bildung ⇄ Wohlbefinden (rechts)
    cx = Inches(10.4)
    base_y = Inches(2.6)
    # Ständer
    pole = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                              cx - Inches(0.04), base_y, Inches(0.08), Inches(2.4))
    set_solid_fill(pole, DARK_SOFT); set_no_line(pole)
    foot = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              cx - Inches(0.6), base_y + Inches(2.35),
                              Inches(1.2), Inches(0.18))
    set_solid_fill(foot, DARK_SOFT); set_no_line(foot)
    # Querbalken
    bar = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                             cx - Inches(2.0), base_y - Inches(0.06),
                             Inches(4.0), Inches(0.12))
    set_solid_fill(bar, DARK_SOFT); set_no_line(bar)
    # Schalen
    left_pan = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  cx - Inches(2.7), base_y + Inches(0.05),
                                  Inches(1.6), Inches(0.55))
    set_solid_fill(left_pan, BLUE); set_no_line(left_pan)
    apply_soft_shadow(left_pan, blur=8, dist=2, alpha=18)
    right_pan = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   cx + Inches(1.1), base_y + Inches(0.05),
                                   Inches(1.6), Inches(0.55))
    set_solid_fill(right_pan, PINK); set_no_line(right_pan)
    apply_soft_shadow(right_pan, blur=8, dist=2, alpha=18)
    # Beschriftung
    add_text(s, cx - Inches(2.85), base_y + Inches(0.8), Inches(1.9), Inches(0.3),
             "Bildung", size=14, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    add_text(s, cx + Inches(0.95), base_y + Inches(0.8), Inches(1.9), Inches(0.3),
             "Wohlbefinden", size=14, bold=True, color=DARK, align=PP_ALIGN.CENTER)

    # Hervorgehobene Leitfrage (unten)
    qbox = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(0.9), Inches(5.85),
                              Inches(11.55), Inches(1.05))
    set_solid_fill(qbox, BLUE); set_no_line(qbox)
    try:
        qbox.adjustments[0] = 0.18
    except Exception:
        pass
    apply_soft_shadow(qbox, blur=12, dist=3, alpha=20)
    add_text(s, Inches(1.2), Inches(5.85), Inches(11.0), Inches(1.05),
             "\u201eInwiefern beeinflussen strukturelle Unterschiede im deutschen Schulsystem das Wohlbefinden von Sch\u00fclerinnen und Sch\u00fclern?\u201c",
             size=18, bold=True, color=DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    add_page_meta(s, 2, accent=YELLOW)
    set_transition(s, "zoom")

    set_notes(s, """
Vielleicht denkt ihr: Schule ist halt Schule – warum sollten wir uns gerade im Kurs Glück damit beschäftigen? Genau das möchte ich auf dieser Folie erklären.

Erstens: Schule prägt unseren Alltag. Wir verbringen Jahre in ihr und sie ist neben der Familie die wichtigste Sozialisationsinstanz unserer Jugend. Zweitens nimmt der Leistungsdruck spürbar zu. Studien wie die KiGGS-Studie des Robert Koch-Instituts und die HBSC-Studie der WHO zeigen, dass Jugendliche schulischen Stress als zentralen Belastungsfaktor erleben. Drittens entscheidet Bildung mit über unsere Zukunft – beruflich, finanziell und sozial. Und viertens: Wohlbefinden ist kein Luxus, sondern Voraussetzung dafür, dass Bildung überhaupt gelingt.

Genau hier setzt meine Leitfrage an: „Inwiefern beeinflussen strukturelle Unterschiede im deutschen Schulsystem das Wohlbefinden von Schülerinnen und Schülern?" Diese Frage zieht sich wie ein roter Faden durch meine gesamte Arbeit – und damit auch durch diese Präsentation.

Bevor wir aber in die Analyse einsteigen, brauchen wir ein paar gemeinsame Grundbegriffe.
""")


# ===========================================================================
# Folie 3 – Grundlagen / Mindmap
# ===========================================================================

def slide_03_grundlagen(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(s, WHITE)

    add_text(s, Inches(0.9), Inches(0.55), Inches(6), Inches(0.35),
             "02 · GRUNDLAGEN", size=11, bold=True, color=ACCENT)
    add_text(s, Inches(0.9), Inches(0.95), Inches(11.5), Inches(1.0),
             "Grundlagen",
             size=44, bold=True, color=DARK, font=FONT_HEAD)
    add_text(s, Inches(0.9), Inches(1.85), Inches(11.5), Inches(0.4),
             "Drei Begriffe, die meine Arbeit zusammenhalten.",
             size=16, color=DARK_SOFT)

    # zentrale Mindmap-Mitte
    cx = Inches(6.66); cy = Inches(4.6)
    center = s.shapes.add_shape(MSO_SHAPE.OVAL,
                                cx - Inches(0.85), cy - Inches(0.85),
                                Inches(1.7), Inches(1.7))
    set_solid_fill(center, DARK); set_no_line(center)
    apply_soft_shadow(center, blur=14, dist=4, alpha=24)
    add_text(s, cx - Inches(0.85), cy - Inches(0.85), Inches(1.7), Inches(1.7),
             "Schule\n& Glück",
             size=15, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Drei Knoten
    nodes = [
        # (x_center, y_center, fill, title, bullets)
        (Inches(2.3),  Inches(3.7), BLUE,
         "Strukturmerkmale",
         ["Schulformen", "Notensystem", "Versetzungen", "Abschlüsse"]),
        (Inches(11.0), Inches(3.7), PINK,
         "Wohlbefinden",
         ["psychische Gesundheit", "soziale Beziehungen", "Lebenszufriedenheit"]),
        (Inches(6.66), Inches(6.45), YELLOW,
         "Bildungsgerechtigkeit",
         ["faire Chancen unabhängig", "von sozialer Herkunft"]),
    ]

    for nx, ny, col, title, items in nodes:
        # Verbindungslinie zur Mitte
        add_line(s, nx, ny, cx, cy, color=LINE_GREY, width_pt=1.5, dashed=True)

        # Karte
        card_w = Inches(3.4); card_h = Inches(1.85)
        card = add_card(s, nx - card_w/2, ny - card_h/2, card_w, card_h, col,
                        radius_factor=0.18)

        # Inhalt
        tf = add_rich_textbox(s, nx - card_w/2 + Inches(0.25),
                              ny - card_h/2 + Inches(0.18),
                              card_w - Inches(0.5), card_h - Inches(0.36))
        add_paragraph(tf, title, size=18, bold=True, color=DARK,
                      first=True, space_after=4, line_spacing=1.1)
        for it in items:
            add_paragraph(tf, "• " + it, size=12, color=DARK_SOFT,
                          space_after=1, line_spacing=1.15)

    add_page_meta(s, 3, accent=PINK)
    set_transition(s, "move in")

    set_notes(s, """
Damit wir alle vom selben sprechen, klären wir drei zentrale Begriffe.

Erstens: Strukturmerkmale von Schule. Damit meine ich die institutionellen Rahmenbedingungen – also Schulformen wie Gymnasium, Realschule oder Hauptschule, das Notensystem, Versetzungsregelungen und Abschlüsse. Diese Strukturen wirken sich nicht nur organisatorisch aus, sondern auch sozial.

Zweitens: Wohlbefinden. Die Weltgesundheitsorganisation versteht darunter mehr als nur das Fehlen von Krankheit. Wohlbefinden umfasst psychische Gesundheit, soziale Beziehungen und subjektive Lebenszufriedenheit. Übertragen auf Schule heißt das: Es geht um emotionale Sicherheit, Zugehörigkeit und das Gefühl, etwas bewirken zu können.

Drittens: Bildungsgerechtigkeit. Sie meint faire Bildungschancen unabhängig von sozialer Herkunft. In Deutschland hängt der Bildungserfolg laut OECD vergleichsweise stark vom Elternhaus ab – ein Punkt, der uns gleich noch begegnen wird.

Diese drei Begriffe – Struktur, Wohlbefinden und Gerechtigkeit – bilden das Dreieck, in dem sich meine Arbeit bewegt. Schauen wir uns jetzt an, wie das deutsche System konkret aufgebaut ist.
""")


# ===========================================================================
# Folie 4 – Das deutsche Schulsystem
# ===========================================================================

def slide_04_schulsystem(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(s, WHITE)
    add_soft_blob(s, Inches(11.6), Inches(-1.5), Inches(3.8), Inches(3.8), YELLOW)

    add_text(s, Inches(0.9), Inches(0.55), Inches(6), Inches(0.35),
             "03 · ÜBERSICHT", size=11, bold=True, color=ACCENT)
    add_text(s, Inches(0.9), Inches(0.95), Inches(11.5), Inches(1.0),
             "Das deutsche Schulsystem",
             size=44, bold=True, color=DARK, font=FONT_HEAD)
    add_text(s, Inches(0.9), Inches(1.85), Inches(11.5), Inches(0.4),
             "Mehrgliedrigkeit – fünf Wege nach der Grundschule.",
             size=16, color=DARK_SOFT)

    # Spalten / Karten
    schools = [
        ("Gymnasium", BLUE, "🎓",
         "Allg. Hochschulreife · Abitur · stärker theoretisch · hohe Anforderungen."),
        ("Realschule", YELLOW, "📘",
         "Mittlerer Schulabschluss · Mischung aus Theorie & Praxis · Brücke zu Beruf oder weiterführender Schule."),
        ("Hauptschule", PINK, "🔧",
         "Praktisch-beruflich · Zugang zu Ausbildung · in vielen Bundesländern integriert oder ersetzt."),
        ("Gesamtschule", GREEN_SOFT, "🤝",
         "Längeres gemeinsames Lernen · mehrere Abschlüsse möglich · Fokus auf individuelle Förderung."),
        ("Berufliche Schule", BLUE, "🛠️",
         "Duales System · Verbindung von Schule & Betrieb · international anerkanntes Modell."),
    ]

    n = len(schools)
    total_w = Inches(12.0)
    gap = Inches(0.18)
    card_w = Emu(int((total_w - gap * (n - 1)) / n))
    start_x = Inches(0.66)
    y = Inches(2.6)
    h = Inches(3.9)

    for i, (name, col, icon, desc) in enumerate(schools):
        x = start_x + (card_w + gap) * i
        # Hintergrundkarte
        card = add_card(s, x, y, card_w, h, WHITE, radius_factor=0.08,
                        line_color=LINE_GREY)
        # Farbkopf
        head = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  x, y, card_w, Inches(1.05))
        set_solid_fill(head, col); set_no_line(head)
        try:
            head.adjustments[0] = 0.1
        except Exception:
            pass
        # weißes Rechteck unten verdeckt unteren Rundungsbogen des Heads
        cover = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   x, y + Inches(0.65), card_w, Inches(0.4))
        set_solid_fill(cover, col); set_no_line(cover)

        # Icon-Kreis
        ic = s.shapes.add_shape(MSO_SHAPE.OVAL,
                                x + card_w/2 - Inches(0.4),
                                y + Inches(0.2),
                                Inches(0.8), Inches(0.8))
        set_solid_fill(ic, WHITE); set_no_line(ic)
        apply_soft_shadow(ic, blur=8, dist=2, alpha=18)
        add_text(s, x + card_w/2 - Inches(0.4),
                 y + Inches(0.2), Inches(0.8), Inches(0.8),
                 icon, size=24, bold=True, color=DARK,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Titel
        add_text(s, x + Inches(0.1), y + Inches(1.2),
                 card_w - Inches(0.2), Inches(0.5),
                 name, size=18, bold=True, color=DARK,
                 align=PP_ALIGN.CENTER)
        # Beschreibung
        add_text(s, x + Inches(0.18), y + Inches(1.75),
                 card_w - Inches(0.36), h - Inches(1.85),
                 desc, size=12, color=DARK_SOFT,
                 align=PP_ALIGN.CENTER)

    # Hinweiszeile
    add_text(s, Inches(0.9), Inches(6.6), Inches(11.5), Inches(0.4),
             "Bildungspolitik ist Ländersache – die Strukturen variieren regional.",
             size=13, color=DARK_SOFT, align=PP_ALIGN.CENTER)

    add_page_meta(s, 4, accent=BLUE)
    set_transition(s, "magic move")

    set_notes(s, """
Schauen wir uns jetzt das deutsche Schulsystem im Überblick an. Charakteristisch ist seine Mehrgliedrigkeit – also die Aufteilung nach der Grundschule auf verschiedene Schulformen.

Das Gymnasium ist die Schulform mit dem höchsten akademischen Anspruch. Es führt zur allgemeinen Hochschulreife und bereitet auf ein Studium vor. Der Unterricht ist theoretisch ausgerichtet, die Anforderungen sind hoch.

Die Realschule liegt in der Mitte. Sie verbindet Allgemeinbildung mit Praxisbezug, führt zum Mittleren Schulabschluss und ist eine wichtige Brücke zwischen schulischer und beruflicher Bildung.

Die Hauptschule war traditionell stark berufsorientiert. Heute spielt sie eine kleinere Rolle, weil sie in vielen Bundesländern mit anderen Schulformen zusammengeführt wurde – auch, weil sie mit gesellschaftlicher Stigmatisierung zu kämpfen hatte.

Die Gesamtschule bietet längeres gemeinsames Lernen und mehrere Abschlüsse innerhalb einer Schule. Sie versucht, soziale Ungleichheit zu reduzieren.

Schließlich die beruflichen Schulen – sie sind zentraler Bestandteil des dualen Ausbildungssystems, das international als Erfolgsmodell gilt.

Wichtig ist: Bildung ist Ländersache. Das heißt, die Strukturen unterscheiden sich von Bundesland zu Bundesland. Dieses System schafft viele Wege – aber es erzeugt auch Druck. Wie sich das auswirkt, sehen wir auf der nächsten Folie.
""")


# ===========================================================================
# Folie 5 – Leistungsdruck und Wohlbefinden
# ===========================================================================

def slide_05_leistungsdruck(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(s, WHITE)
    add_soft_blob(s, Inches(-1.5), Inches(5.0), Inches(4.0), Inches(4.0), PINK)
    add_soft_blob(s, Inches(11.5), Inches(-1.0), Inches(3.5), Inches(3.5), BLUE)

    add_text(s, Inches(0.9), Inches(0.55), Inches(8), Inches(0.35),
             "04 · ANALYSE", size=11, bold=True, color=ACCENT)
    add_text(s, Inches(0.9), Inches(0.95), Inches(11.5), Inches(1.0),
             "Leistungsdruck und Wohlbefinden",
             size=40, bold=True, color=DARK, font=FONT_HEAD)
    add_text(s, Inches(0.9), Inches(1.85), Inches(11.5), Inches(0.4),
             "Was Schule fordert – und was sie auslösen kann.",
             size=16, color=DARK_SOFT)

    # Linke Karte – Leistungsdruck
    left = add_card(s, Inches(0.9), Inches(2.5), Inches(5.5), Inches(3.4),
                    BLUE, radius_factor=0.1)
    add_text(s, Inches(1.15), Inches(2.65), Inches(5), Inches(0.55),
             "Leistungsdruck", size=22, bold=True, color=DARK)
    tf1 = add_rich_textbox(s, Inches(1.15), Inches(3.2),
                           Inches(5), Inches(2.6))
    items_l = ["Noten", "Prüfungen", "Konkurrenz", "Erwartungen"]
    for i, it in enumerate(items_l):
        p = tf1.paragraphs[0] if i == 0 else tf1.add_paragraph()
        p.space_after = Pt(8)
        p.line_spacing = 1.2
        r1 = p.add_run(); r1.text = "›  "
        r1.font.size = Pt(20); r1.font.bold = True; r1.font.color.rgb = ACCENT
        r2 = p.add_run(); r2.text = it
        r2.font.size = Pt(20); r2.font.color.rgb = DARK; r2.font.name = FONT_BODY

    # Mittlerer Pfeil / Verbindung
    arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                               Inches(6.55), Inches(3.95),
                               Inches(0.3), Inches(0.5))
    set_solid_fill(arrow, DARK_SOFT); set_no_line(arrow)

    # Rechte Karte – Folgen
    right = add_card(s, Inches(6.95), Inches(2.5), Inches(5.5), Inches(3.4),
                     PINK, radius_factor=0.1)
    add_text(s, Inches(7.2), Inches(2.65), Inches(5), Inches(0.55),
             "Mögliche Folgen", size=22, bold=True, color=DARK)
    tf2 = add_rich_textbox(s, Inches(7.2), Inches(3.2),
                           Inches(5), Inches(2.6))
    items_r = ["Stress", "Schlafprobleme", "Unsicherheit", "geringeres Selbstwertgefühl"]
    for i, it in enumerate(items_r):
        p = tf2.paragraphs[0] if i == 0 else tf2.add_paragraph()
        p.space_after = Pt(8)
        p.line_spacing = 1.2
        r1 = p.add_run(); r1.text = "›  "
        r1.font.size = Pt(20); r1.font.bold = True; r1.font.color.rgb = ACCENT
        r2 = p.add_run(); r2.text = it
        r2.font.size = Pt(20); r2.font.color.rgb = DARK; r2.font.name = FONT_BODY

    # Merksatz unten
    note = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(0.9), Inches(6.05),
                              Inches(11.55), Inches(0.85))
    set_solid_fill(note, YELLOW); set_no_line(note)
    try:
        note.adjustments[0] = 0.4
    except Exception:
        pass
    apply_soft_shadow(note, blur=12, dist=3, alpha=18)
    add_text(s, Inches(1.2), Inches(6.05), Inches(11.0), Inches(0.85),
             "\u201eLeistungsdruck kann Motivation f\u00f6rdern \u2013 aber auch das Wohlbefinden beeintr\u00e4chtigen.\u201c",
             size=17, bold=True, color=DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    add_page_meta(s, 5, accent=PINK)
    set_transition(s, "scale")

    set_notes(s, """
Leistungsdruck ist eines der wichtigsten Themen meiner Arbeit. Die Folie ist deshalb bewusst zweigeteilt: Links seht ihr die Quellen, rechts die möglichen Folgen.

Auf der linken Seite stehen die typischen Auslöser von Druck: Noten, Prüfungen, Konkurrenz untereinander – und Erwartungen. Erwartungen kommen dabei aus mehreren Richtungen: von Eltern, Lehrkräften und der Gesellschaft. Schon die PISA-Studie zeigt, dass viele Schülerinnen und Schüler sich allein wegen ihrer Noten massiv unter Druck fühlen.

Auf der rechten Seite seht ihr, was daraus entstehen kann. Die KiGGS-Studie und die HBSC-Studie der WHO weisen nach, dass schulischer Leistungsdruck mit Stress, Schlafproblemen, Konzentrationsschwierigkeiten, einem geringeren Selbstwertgefühl und teils sogar mit psychosomatischen Beschwerden einhergeht.

Wichtig ist aber – und das steht im Merksatz unten: Leistungsdruck ist nicht per se negativ. Eine gewisse Anforderung kann auch motivieren. Entscheidend ist, ob das Maß stimmt und ob das Umfeld unterstützt.

Damit kommen wir zu genau diesem Umfeld – dem Schulklima.
""")


# ===========================================================================
# Folie 6 – Schulklima
# ===========================================================================

def slide_06_schulklima(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(s, WHITE)
    add_soft_blob(s, Inches(11.0), Inches(5.5), Inches(3.8), Inches(3.8), YELLOW)

    add_text(s, Inches(0.9), Inches(0.55), Inches(8), Inches(0.35),
             "05 · BEZIEHUNGEN", size=11, bold=True, color=ACCENT)
    add_text(s, Inches(0.9), Inches(0.95), Inches(11.5), Inches(1.0),
             "Schulklima",
             size=44, bold=True, color=DARK, font=FONT_HEAD)
    add_text(s, Inches(0.9), Inches(1.85), Inches(11.5), Inches(0.4),
             "Atmosphäre, Beziehungen und das Gefühl von Zugehörigkeit.",
             size=16, color=DARK_SOFT)

    # Drei große Karten
    cards = [
        ("Lehrer-Schüler-\nVerhältnis", BLUE, "👩‍🏫",
         ["Vertrauen & Respekt",
          "Unterstützung statt Kontrolle",
          "Reduziert Stress, fördert Lernfreude"]),
        ("Mitbestimmung", PINK, "🗳️",
         ["Schülervertretungen, Klassenrat",
          "Beteiligung an Schulprojekten",
          "Stärkt Selbstwirksamkeit"]),
        ("Feedbackkultur", YELLOW, "💬",
         ["Mehr als nur Noten",
          "Formative Rückmeldungen",
          "Begleitet Lernprozesse individuell"]),
    ]

    n = len(cards)
    total_w = Inches(11.5)
    gap = Inches(0.3)
    card_w = Emu(int((total_w - gap * (n - 1)) / n))
    start_x = Inches(0.9)
    y = Inches(2.7)
    h = Inches(3.9)

    for i, (title, col, icon, items) in enumerate(cards):
        x = start_x + (card_w + gap) * i
        card = add_card(s, x, y, card_w, h, WHITE, line_color=LINE_GREY,
                        radius_factor=0.06)
        # Pastellbalken oben
        bar = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 x, y, card_w, Inches(0.55))
        set_solid_fill(bar, col); set_no_line(bar)
        try:
            bar.adjustments[0] = 0.25
        except Exception:
            pass
        cover = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   x, y + Inches(0.3), card_w, Inches(0.3))
        set_solid_fill(cover, col); set_no_line(cover)

        # Icon
        add_circle_icon(s, x + card_w/2, y + Inches(1.25), Inches(0.55),
                        col, glyph=icon, glyph_size=26, glyph_color=DARK)

        # Titel
        add_text(s, x + Inches(0.2), y + Inches(2.0),
                 card_w - Inches(0.4), Inches(0.95),
                 title, size=20, bold=True, color=DARK,
                 align=PP_ALIGN.CENTER)
        # Bullets
        tf = add_rich_textbox(s, x + Inches(0.3), y + Inches(2.95),
                              card_w - Inches(0.6), h - Inches(3.05))
        for j, it in enumerate(items):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.space_after = Pt(6); p.line_spacing = 1.15
            r1 = p.add_run(); r1.text = "•  "
            r1.font.size = Pt(13); r1.font.bold = True; r1.font.color.rgb = ACCENT
            r2 = p.add_run(); r2.text = it
            r2.font.size = Pt(13); r2.font.color.rgb = DARK_SOFT; r2.font.name = FONT_BODY

    add_page_meta(s, 6, accent=BLUE)
    set_transition(s, "push")

    set_notes(s, """
Wenn wir über Wohlbefinden in der Schule sprechen, müssen wir über das Schulklima sprechen. Damit ist die soziale Atmosphäre gemeint – also Beziehungen, Vertrauen, Sicherheit und Zugehörigkeit. Drei Punkte sind dabei besonders wichtig.

Erstens: das Lehrer-Schüler-Verhältnis. Lehrkräfte sind nicht nur Wissensvermittler, sondern auch Bezugspersonen. Wenn Schülerinnen und Schüler sich respektiert und unterstützt fühlen, sinkt das Stresserleben und die Motivation steigt deutlich.

Zweitens: Mitbestimmung. Klassenrat, Schülervertretungen oder Beteiligung an Projekten geben uns das Gefühl, gestalten zu können. In der Forschung nennt man das Selbstwirksamkeit – und sie ist einer der stärksten Schutzfaktoren gegen schulischen Stress.

Drittens: Feedbackkultur. Hier geht es darum, Lernen zu begleiten – nicht nur am Ende mit einer Note zu beurteilen. Formative Rückmeldungen wie Lernfeedback oder Portfolios helfen, Stärken zu erkennen und Schwächen gezielt zu verbessern.

Zusammen ergibt das ein Bild: Schulklima entscheidet mit darüber, ob Schule ein Lebensraum ist, in dem man sich entfalten kann. Wie wir das selbst erleben, schauen wir uns auf der nächsten Folie kurz interaktiv an.
""")


# ===========================================================================
# Folie 7 – Publikum einbinden / Umfrage
# ===========================================================================

def slide_07_umfrage(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(s, WHITE)
    add_soft_blob(s, Inches(-1.8), Inches(-1.8), Inches(4.5), Inches(4.5), BLUE)
    add_soft_blob(s, Inches(11.0), Inches(5.0), Inches(4), Inches(4), PINK)

    add_text(s, Inches(0.9), Inches(0.55), Inches(8), Inches(0.35),
             "06 · INTERAKTION", size=11, bold=True, color=ACCENT)
    add_text(s, Inches(0.9), Inches(0.95), Inches(11.5), Inches(1.0),
             "Kurze Umfrage",
             size=44, bold=True, color=DARK, font=FONT_HEAD)

    # Große Frage in der Mitte
    qbox = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(0.9), Inches(2.05),
                              Inches(11.55), Inches(1.4))
    set_solid_fill(qbox, YELLOW); set_no_line(qbox)
    try:
        qbox.adjustments[0] = 0.18
    except Exception:
        pass
    apply_soft_shadow(qbox, blur=14, dist=4, alpha=22)
    add_text(s, Inches(1.2), Inches(2.05), Inches(11.0), Inches(1.4),
             "\u201eWas beeinflusst euer Wohlbefinden in der Schule am st\u00e4rksten?\u201c",
             size=22, bold=True, color=DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    # Vier Antwort-Buttons
    buttons = [
        ("📚", "Notendruck", BLUE),
        ("👩‍🏫", "Lehrkräfte", PINK),
        ("👥", "Mitschüler:innen", YELLOW),
        ("🕒", "Zeitstress", BLUE),
    ]
    n = len(buttons)
    total_w = Inches(11.5)
    gap = Inches(0.25)
    bw = Emu(int((total_w - gap * (n - 1)) / n))
    bx = Inches(0.9); by = Inches(3.85); bh = Inches(1.4)
    for i, (icon, label, col) in enumerate(buttons):
        x = bx + (bw + gap) * i
        btn = add_card(s, x, by, bw, bh, col, radius_factor=0.25)
        add_text(s, x, by + Inches(0.2), bw, Inches(0.55),
                 icon, size=30, bold=True, color=DARK,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x, by + Inches(0.8), bw, Inches(0.5),
                 label, size=17, bold=True, color=DARK,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_text(s, Inches(0.9), Inches(5.45), Inches(11.55), Inches(0.4),
             "Kurzes Handzeichen.",
             size=16, color=DARK_SOFT, align=PP_ALIGN.CENTER)

    # Auflösung als Pastell-Pille
    reveal = add_card(s, Inches(2.5), Inches(6.0),
                      Inches(8.35), Inches(0.85), PINK, radius_factor=0.4)
    add_text(s, Inches(2.5), Inches(6.0), Inches(8.35), Inches(0.85),
             "Alle Faktoren spielen laut Facharbeit eine wichtige Rolle.",
             size=18, bold=True, color=DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_page_meta(s, 7, accent=YELLOW)
    set_transition(s, "dissolve")

    set_notes(s, """
Bevor ich mit der nächsten inhaltlichen Folie weitermache, möchte ich euch kurz einbinden. Schaut bitte auf die Frage in der Mitte: „Was beeinflusst euer Wohlbefinden in der Schule am stärksten?"

Ich gebe euch vier Möglichkeiten: Notendruck, Lehrkräfte, Mitschülerinnen und Mitschüler oder Zeitstress. Hebt einfach kurz die Hand bei dem Punkt, der euch am ehesten beeinflusst.

(Pause für die Umfrage. Antworten kurz kommentieren.)

Danke. Wenig überraschend: Es ist meistens nicht nur einer dieser Punkte. Genau das zeigt auch meine Facharbeit: Notendruck, Lehrkräfte, soziale Beziehungen und Zeitdruck wirken zusammen. Sie sind wie Schichten, die sich gegenseitig verstärken oder abfedern können.

Mit diesem Bild im Kopf schauen wir nun, wie andere Länder mit denselben Herausforderungen umgehen.
""")


# ===========================================================================
# Folie 8 – Internationaler Vergleich
# ===========================================================================

def slide_08_international(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(s, WHITE)

    add_text(s, Inches(0.9), Inches(0.55), Inches(8), Inches(0.35),
             "07 · BLICK ÜBER DEN TELLERRAND", size=11, bold=True, color=ACCENT)
    add_text(s, Inches(0.9), Inches(0.95), Inches(11.5), Inches(1.0),
             "Internationaler Vergleich",
             size=44, bold=True, color=DARK, font=FONT_HEAD)
    add_text(s, Inches(0.9), Inches(1.85), Inches(11.5), Inches(0.4),
             "Drei Bildungssysteme – drei Philosophien.",
             size=16, color=DARK_SOFT)

    countries = [
        ("Deutschland", "🇩🇪", BLUE,
         ["frühe Aufteilung nach Klasse 4",
          "viele Noten & Prüfungen",
          "starke Differenzierung"]),
        ("Finnland", "🇫🇮", YELLOW,
         ["längeres gemeinsames Lernen",
          "spätere Leistungsdifferenzierung",
          "geringer Notenfokus"]),
        ("Angelsächsisch", "🇬🇧", PINK,
         ["umfassende Gesamtschulen",
          "mehr Wahlmöglichkeiten",
          "stärkere Praxis- & Projektorientierung"]),
    ]

    n = len(countries)
    total_w = Inches(11.5)
    gap = Inches(0.3)
    cw = Emu(int((total_w - gap * (n - 1)) / n))
    start_x = Inches(0.9)
    y = Inches(2.6)
    h = Inches(3.9)

    for i, (name, flag, col, items) in enumerate(countries):
        x = start_x + (cw + gap) * i
        add_card(s, x, y, cw, h, WHITE, line_color=LINE_GREY, radius_factor=0.07)

        # Header in Pastellfarbe
        head = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  x, y, cw, Inches(1.05))
        set_solid_fill(head, col); set_no_line(head)
        try:
            head.adjustments[0] = 0.12
        except Exception:
            pass
        cover = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   x, y + Inches(0.55), cw, Inches(0.5))
        set_solid_fill(cover, col); set_no_line(cover)

        add_text(s, x, y + Inches(0.18), cw, Inches(0.5),
                 flag, size=30, bold=True, color=DARK,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x, y + Inches(0.7), cw, Inches(0.4),
                 name, size=18, bold=True, color=DARK,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Bullets
        tf = add_rich_textbox(s, x + Inches(0.3), y + Inches(1.3),
                              cw - Inches(0.6), h - Inches(1.4))
        for j, it in enumerate(items):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.space_after = Pt(8); p.line_spacing = 1.2
            r1 = p.add_run(); r1.text = "›  "
            r1.font.size = Pt(15); r1.font.bold = True; r1.font.color.rgb = ACCENT
            r2 = p.add_run(); r2.text = it
            r2.font.size = Pt(15); r2.font.color.rgb = DARK_SOFT; r2.font.name = FONT_BODY

    # Hinweis unten
    bottom = add_card(s, Inches(2.0), Inches(6.7), Inches(9.3), Inches(0.55),
                      YELLOW, radius_factor=0.5)
    add_text(s, Inches(2.0), Inches(6.7), Inches(9.3), Inches(0.55),
             "Kein System ist perfekt – aber jedes setzt eigene Schwerpunkte.",
             size=15, bold=True, color=DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_page_meta(s, 8, accent=PINK)
    set_transition(s, "pop")

    set_notes(s, """
Deutschland ist nur ein Modell unter vielen. Internationale Vergleichsstudien wie PISA helfen uns, einzuordnen, was wir hier eigentlich tun.

In Deutschland erfolgt die Aufteilung der Schülerinnen und Schüler typischerweise schon nach Klasse 4. Wir setzen außerdem stark auf Noten, Prüfungen und Differenzierung zwischen den Schulformen.

Finnland geht einen ganz anderen Weg. Dort lernen Kinder bis zur 9. Klasse gemeinsam. Erst danach erfolgt eine Differenzierung. Auch beim Notensystem ist Finnland zurückhaltender und setzt stärker auf qualitative Rückmeldungen. Das Resultat: relativ geringe Leistungsunterschiede zwischen sozialen Gruppen – und international anerkannt gute Leistungen.

In angelsächsischen Systemen, also im Vereinigten Königreich oder den USA, dominieren umfassende Gesamtschulen. Schüler:innen können oft früh eigene Schwerpunkte setzen, Kurse wählen und arbeiten häufiger projektorientiert. Auch hier gibt es aber Schwächen, etwa starke regionale Unterschiede.

Was lernen wir daraus? Kein System ist perfekt – aber wenn man Wohlbefinden ernst nimmt, sind alternative Strukturen möglich. Wie wir die deutsche Variante einordnen können, sehen wir auf der nächsten Folie.
""")


# ===========================================================================
# Folie 9 – Stärken und Schwächen
# ===========================================================================

def slide_09_swot(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(s, WHITE)
    add_soft_blob(s, Inches(11.0), Inches(-1.5), Inches(3.5), Inches(3.5), BLUE)
    add_soft_blob(s, Inches(-1.5), Inches(5.0), Inches(3.5), Inches(3.5), YELLOW)

    add_text(s, Inches(0.9), Inches(0.55), Inches(8), Inches(0.35),
             "08 · BEWERTUNG", size=11, bold=True, color=ACCENT)
    add_text(s, Inches(0.9), Inches(0.95), Inches(11.5), Inches(1.0),
             "Stärken und Schwächen",
             size=44, bold=True, color=DARK, font=FONT_HEAD)
    add_text(s, Inches(0.9), Inches(1.85), Inches(11.5), Inches(0.4),
             "Das deutsche Bildungssystem zeigt zwei Gesichter.",
             size=16, color=DARK_SOFT)

    # Linke grüne Karte – Stärken
    left = add_card(s, Inches(0.9), Inches(2.5), Inches(5.7), Inches(4.2),
                    GREEN_SOFT, radius_factor=0.08)

    # Plus-Icon-Kreis oben
    add_circle_icon(s, Inches(1.55), Inches(2.95), Inches(0.42),
                    WHITE, glyph="+", glyph_size=28, glyph_color=DARK)
    add_text(s, Inches(2.2), Inches(2.7), Inches(4), Inches(0.55),
             "Stärken", size=24, bold=True, color=DARK,
             anchor=MSO_ANCHOR.MIDDLE)

    tf_l = add_rich_textbox(s, Inches(1.3), Inches(3.55),
                            Inches(5.0), Inches(2.9))
    strengths = [
        "Verschiedene Bildungswege",
        "Duales Ausbildungssystem",
        "Hohe akademische Standards",
    ]
    for i, st in enumerate(strengths):
        p = tf_l.paragraphs[0] if i == 0 else tf_l.add_paragraph()
        p.space_after = Pt(12); p.line_spacing = 1.25
        r1 = p.add_run(); r1.text = "✓  "
        r1.font.size = Pt(20); r1.font.bold = True; r1.font.color.rgb = DARK
        r2 = p.add_run(); r2.text = st
        r2.font.size = Pt(20); r2.font.color.rgb = DARK; r2.font.name = FONT_BODY

    # Rechte rosa Karte – Schwächen
    right = add_card(s, Inches(6.75), Inches(2.5), Inches(5.7), Inches(4.2),
                     PINK_SOFT, radius_factor=0.08)

    add_circle_icon(s, Inches(7.4), Inches(2.95), Inches(0.42),
                    WHITE, glyph="–", glyph_size=30, glyph_color=DARK)
    add_text(s, Inches(8.05), Inches(2.7), Inches(4), Inches(0.55),
             "Schwächen", size=24, bold=True, color=DARK,
             anchor=MSO_ANCHOR.MIDDLE)

    tf_r = add_rich_textbox(s, Inches(7.15), Inches(3.55),
                            Inches(5.0), Inches(2.9))
    weaknesses = [
        "Frühe Selektion",
        "Hoher Leistungsdruck",
        "Soziale Ungleichheiten",
    ]
    for i, w in enumerate(weaknesses):
        p = tf_r.paragraphs[0] if i == 0 else tf_r.add_paragraph()
        p.space_after = Pt(12); p.line_spacing = 1.25
        r1 = p.add_run(); r1.text = "•  "
        r1.font.size = Pt(20); r1.font.bold = True; r1.font.color.rgb = DARK
        r2 = p.add_run(); r2.text = w
        r2.font.size = Pt(20); r2.font.color.rgb = DARK; r2.font.name = FONT_BODY

    # Schlusszeile
    add_text(s, Inches(0.9), Inches(6.85), Inches(11.55), Inches(0.4),
             "Beides gehört zur ehrlichen Bewertung – Reformen setzen genau hier an.",
             size=14, color=DARK_SOFT, align=PP_ALIGN.CENTER)

    add_page_meta(s, 9, accent=BLUE)
    set_transition(s, "slide")

    set_notes(s, """
Wenn wir das deutsche System bewerten, müssen wir fair sein – und beide Seiten anschauen.

Auf der Habenseite stehen drei klare Stärken. Erstens: Wir haben verschiedene Bildungswege, die individuelle Förderung ermöglichen. Zweitens: Das duale Ausbildungssystem gilt international als Vorbild. Es verbindet Schule und Betrieb und sorgt unter anderem für eine relativ niedrige Jugendarbeitslosigkeit. Drittens: Die akademischen Standards, gerade am Gymnasium, sind hoch und international anerkannt.

Auf der anderen Seite gibt es aber auch deutliche Schwächen. Erstens: die frühe Selektion nach Klasse 4. In diesem Alter sind Leistungsunterschiede stark vom Elternhaus geprägt – die Aufteilung wirkt damit oft sozial selektiv. Zweitens: der hohe Leistungsdruck, der psychisch belastet, wie wir bereits gesehen haben. Drittens: soziale Ungleichheiten. In Deutschland hängt der Bildungserfolg laut OECD vergleichsweise stark vom sozioökonomischen Hintergrund der Familie ab.

Diese ehrliche Bestandsaufnahme ist wichtig, denn sie ist die Grundlage für die Reformansätze, die wir uns jetzt ansehen.
""")


# ===========================================================================
# Folie 10 – Reformansätze
# ===========================================================================

def slide_10_reformen(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(s, WHITE)
    add_soft_blob(s, Inches(11.5), Inches(-1.0), Inches(3.5), Inches(3.5), PINK)
    add_soft_blob(s, Inches(-2.0), Inches(5.0), Inches(4), Inches(4), BLUE)

    add_text(s, Inches(0.9), Inches(0.55), Inches(8), Inches(0.35),
             "09 · AUSBLICK", size=11, bold=True, color=ACCENT)
    add_text(s, Inches(0.9), Inches(0.95), Inches(11.5), Inches(1.0),
             "Reformansätze",
             size=44, bold=True, color=DARK, font=FONT_HEAD)
    add_text(s, Inches(0.9), Inches(1.85), Inches(11.5), Inches(0.4),
             "Drei Hebel für ein gerechteres und gesünderes Schulsystem.",
             size=16, color=DARK_SOFT)

    # Pfeil in Richtung Zukunft (dezent rechts)
    arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                               Inches(8.7), Inches(6.45),
                               Inches(2.8), Inches(0.45))
    set_solid_fill(arrow, BLUE); set_no_line(arrow)
    apply_soft_shadow(arrow, blur=10, dist=2, alpha=18)
    add_text(s, Inches(8.7), Inches(6.45), Inches(2.8), Inches(0.45),
             "Zukunft", size=14, bold=True, color=DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Drei große Kreise mit Erklärung darunter
    circles = [
        (Inches(2.7), BLUE, "📚",
         "Mehr Lebens-\nkompetenzen",
         "Finanzbildung, Medienkompetenz und Alltagsfähigkeiten stärker im Curriculum."),
        (Inches(6.66), YELLOW, "⏳",
         "Spätere Leistungs-\ndifferenzierung",
         "Längeres gemeinsames Lernen reduziert sozial bedingte Ungleichheiten."),
        (Inches(10.6), PINK, "📝",
         "Alternative\nBewertungsformen",
         "Lernfeedback, Portfolio- und Projektarbeit ergänzen klassische Noten."),
    ]

    cy = Inches(3.5)
    r = Inches(0.9)
    for cx, col, icon, title, desc in circles:
        circle = s.shapes.add_shape(MSO_SHAPE.OVAL,
                                    cx - r, cy - r, r * 2, r * 2)
        set_solid_fill(circle, col); set_no_line(circle)
        apply_soft_shadow(circle, blur=14, dist=4, alpha=22)
        add_text(s, cx - r, cy - r, r * 2, r * 2,
                 icon, size=42, bold=True, color=DARK,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Titel unter Kreis
        add_text(s, cx - Inches(1.7), cy + r + Inches(0.18),
                 Inches(3.4), Inches(0.85),
                 title, size=18, bold=True, color=DARK,
                 align=PP_ALIGN.CENTER)
        # Beschreibung
        add_text(s, cx - Inches(1.85), cy + r + Inches(1.0),
                 Inches(3.7), Inches(1.4),
                 desc, size=13, color=DARK_SOFT,
                 align=PP_ALIGN.CENTER)

    add_page_meta(s, 10, accent=YELLOW)
    set_transition(s, "rotate")

    set_notes(s, """
Wo könnte Reform ansetzen? In der Bildungsforschung werden vor allem drei Hebel diskutiert. Sie passen direkt zu den Schwächen, die wir gerade gesehen haben.

Der erste Hebel: mehr Lebenskompetenzen im Curriculum. Viele Jugendliche kritisieren, dass Schule zu wenig auf den Alltag vorbereitet – Stichworte sind Steuern, Mietverträge oder Versicherungen. Dazu kommt das Thema Finanzbildung, das international zunehmend ernst genommen wird. Auch die OECD betont Kompetenzen wie Problemlösen, kritisches Denken und Teamarbeit.

Der zweite Hebel: spätere Leistungsdifferenzierung. Wenn Kinder länger gemeinsam lernen, hängt der Bildungsweg weniger stark vom Elternhaus ab. Skandinavische Modelle zeigen, dass das funktionieren kann.

Der dritte Hebel: alternative Bewertungsformen. Statt sich allein auf Noten zu verlassen, könnten Lernfeedback, Portfolio-Arbeit und projektbasierte Formate viel stärker in den Vordergrund rücken. Lernen wird so begleitet, nicht nur abgeprüft.

Diese drei Hebel zeigen: Es gibt realistische Wege, Bildung und Wohlbefinden besser zu verbinden. Damit kommen wir zu meinem Fazit.
""")


# ===========================================================================
# Folie 11 – Fazit
# ===========================================================================

def slide_11_fazit(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(s, WHITE)
    add_soft_blob(s, Inches(-2.0), Inches(-2.0), Inches(5.5), Inches(5.5), BLUE)
    add_soft_blob(s, Inches(11.0), Inches(4.7), Inches(4.5), Inches(4.5), PINK)
    add_soft_blob(s, Inches(9.5), Inches(-1.5), Inches(3.0), Inches(3.0), YELLOW)

    add_text(s, Inches(0.9), Inches(0.55), Inches(6), Inches(0.35),
             "10 · FAZIT", size=11, bold=True, color=ACCENT)

    # Große Headline
    add_text(s, Inches(0.9), Inches(1.0), Inches(11.5), Inches(1.7),
             "Schule beeinflusst mehr\nals nur Noten.",
             size=46, bold=True, color=DARK, font=FONT_HEAD,
             line_spacing=1.05)

    # Drei Kernaussagen als Karten
    cards = [
        (BLUE, "Strukturen wirken auf das Wohlbefinden."),
        (YELLOW, "Leistungsdruck und Schulklima spielen eine zentrale Rolle."),
        (PINK, "Reformen können Bildung und Glück verbinden."),
    ]
    y0 = Inches(3.55)
    for i, (col, txt) in enumerate(cards):
        y = y0 + Inches(0.78) * i
        card = add_card(s, Inches(0.9), y, Inches(8.6), Inches(0.7),
                        col, radius_factor=0.4)
        # Häkchen-Kreis
        add_circle_icon(s, Inches(1.25), y + Inches(0.35), Inches(0.22),
                        WHITE, glyph="✓", glyph_size=14, glyph_color=DARK)
        add_text(s, Inches(1.7), y, Inches(7.7), Inches(0.7),
                 txt, size=17, bold=True, color=DARK,
                 anchor=MSO_ANCHOR.MIDDLE)

    # Pastell-Illustration rechts: ruhiger Kreis mit Herz/Gleichgewicht
    cx = Inches(11.05); cy = Inches(4.2); rr = Inches(1.05)
    halo = s.shapes.add_shape(MSO_SHAPE.OVAL,
                              cx - rr - Inches(0.25), cy - rr - Inches(0.25),
                              (rr + Inches(0.25)) * 2, (rr + Inches(0.25)) * 2)
    set_solid_fill(halo, YELLOW); set_no_line(halo)
    main = s.shapes.add_shape(MSO_SHAPE.OVAL,
                              cx - rr, cy - rr, rr * 2, rr * 2)
    set_solid_fill(main, PINK); set_no_line(main)
    apply_soft_shadow(main, blur=14, dist=4, alpha=22)
    add_text(s, cx - rr, cy - rr, rr * 2, rr * 2,
             "♡", size=64, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Dank-Pille unten
    thanks = add_card(s, Inches(0.9), Inches(6.65),
                      Inches(11.55), Inches(0.65),
                      DARK, radius_factor=0.5)
    add_text(s, Inches(0.9), Inches(6.65), Inches(11.55), Inches(0.65),
             "Danke für eure Aufmerksamkeit!",
             size=20, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    set_transition(s, "fade")

    set_notes(s, """
Damit komme ich zum Fazit meiner Arbeit. Mein Leitsatz lautet: Schule beeinflusst mehr als nur Noten.

Erste Kernaussage: Strukturen wirken auf das Wohlbefinden. Die Mehrgliedrigkeit, das Notensystem, Versetzungsregelungen und die soziale Selektion sind nicht neutral – sie prägen, wie sich Schülerinnen und Schüler in der Schule erleben.

Zweite Kernaussage: Leistungsdruck und Schulklima spielen eine zentrale Rolle. Druck kann motivieren, aber er kann auch krank machen. Entscheidend sind die Beziehungen, die Mitbestimmung und die Feedbackkultur, die ihn entweder abfedern oder verstärken.

Dritte Kernaussage: Reformen können Bildung und Glück stärker miteinander verbinden. Internationale Beispiele zeigen, dass das möglich ist – durch mehr Lebenskompetenzen, spätere Differenzierung und alternative Bewertungsformen.

Schule ist also nicht einfach nur ein Ort, an dem wir Wissen erwerben. Sie ist ein Lebensraum, der unser Wohlbefinden prägt. Und genau deshalb passt das Thema so gut in unseren Seminarfachkurs Glück.

Ich danke euch herzlich für eure Aufmerksamkeit – und freue mich jetzt auf eure Fragen und eure Diskussion.
""")


# ===========================================================================
# Build & Save
# ===========================================================================

def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slide_01_titel(prs)
    slide_02_relevanz(prs)
    slide_03_grundlagen(prs)
    slide_04_schulsystem(prs)
    slide_05_leistungsdruck(prs)
    slide_06_schulklima(prs)
    slide_07_umfrage(prs)
    slide_08_international(prs)
    slide_09_swot(prs)
    slide_10_reformen(prs)
    slide_11_fazit(prs)

    out = "/workspace/praesentation/Wie_Schule_wirkt.pptx"
    prs.save(out)
    print(f"OK -> {out}")


if __name__ == "__main__":
    build()
