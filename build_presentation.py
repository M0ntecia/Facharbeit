# -*- coding: utf-8 -*-
"""
Build a modern, Apple-Keynote-style presentation from the Facharbeit
"Wie Schule wirkt" (Seminarfachkurs Glueck, Svea Timphus).

Output: Glueck_Schule_Wohlbefinden.pptx  (16:9, opens natively in Keynote / iPad)

Contents are derived EXCLUSIVELY from the Facharbeit.
Each slide gets its own transition + sequential build-in text animations,
plus ~1 minute presenter notes.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "Glueck_Schule_Wohlbefinden.pptx")

# ----------------------------------------------------------------------------
# Theme
# ----------------------------------------------------------------------------
BLUE      = RGBColor(0xCF, 0xEF, 0xFF)
PINK      = RGBColor(0xFF, 0xD9, 0xE8)
YELLOW    = RGBColor(0xFF, 0xF7, 0xC7)
MINT      = RGBColor(0xD7, 0xF2, 0xDD)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
DARK      = RGBColor(0x2E, 0x2E, 0x38)
GRAY      = RGBColor(0x6C, 0x6C, 0x78)

BLUE_D    = RGBColor(0x4F, 0xA3, 0xD9)
PINK_D    = RGBColor(0xE8, 0x6F, 0xA1)
YELLOW_D  = RGBColor(0xD9, 0xB8, 0x3A)
MINT_D    = RGBColor(0x55, 0xB0, 0x77)

FONT      = "Helvetica Neue"
FONT_HEAD = "Helvetica Neue"

EMU_IN = 914400
SW = 13.333   # slide width  (inches)
SH = 7.5      # slide height (inches)

NS_P  = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A  = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"

# ----------------------------------------------------------------------------
# Presentation scaffold
# ----------------------------------------------------------------------------
prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]


def slide():
    return prs.slides.add_slide(BLANK)


# ----------------------------------------------------------------------------
# Low-level helpers
# ----------------------------------------------------------------------------
def _spPr(shape):
    return shape._element.spPr


def no_line(shape):
    shape.line.fill.background()


def solid(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def fill_alpha(shape, alpha_pct):
    """Add transparency to a shape's solid fill. alpha_pct 0..100 (opacity)."""
    sp = _spPr(shape)
    sf = sp.find(qn("a:solidFill"))
    srgb = sf.find(qn("a:srgbClr"))
    srgb.append(parse_xml(
        '<a:alpha xmlns:a="%s" val="%d"/>' % (NS_A, int(alpha_pct * 1000))))


def shadow(shape, blur=15, dist=4, alpha=22, direction=5400000):
    """Soft, subtle drop shadow. blur/dist in points, alpha = opacity %."""
    sp = _spPr(shape)
    old = sp.find(qn("a:effectLst"))
    if old is not None:
        sp.remove(old)
    xml = (
        '<a:effectLst xmlns:a="%s">'
        '<a:outerShdw blurRad="%d" dist="%d" dir="%d" rotWithShape="0">'
        '<a:srgbClr val="2E2E38"><a:alpha val="%d"/></a:srgbClr>'
        '</a:outerShdw></a:effectLst>'
        % (NS_A, int(blur * 12700), int(dist * 12700), direction, int(alpha * 1000))
    )
    sp.append(parse_xml(xml))


def round_radius(shape, frac):
    """Set corner radius of a rounded rectangle (0..0.5 of short side)."""
    try:
        shape.adjustments[0] = frac
    except Exception:
        pass


def rrect(s, x, y, w, h, color, radius=0.12, line=None, line_w=1.0,
          shadow_on=True, alpha=None):
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                            Inches(x), Inches(y), Inches(w), Inches(h))
    round_radius(sp, radius)
    solid(sp, color)
    if alpha is not None:
        fill_alpha(sp, alpha)
    if line is None:
        no_line(sp)
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    if shadow_on:
        shadow(sp)
    sp.shadow.inherit = False
    return sp


def oval(s, x, y, w, h, color, alpha=None, shadow_on=False, line=None, line_w=1.0):
    sp = s.shapes.add_shape(MSO_SHAPE.OVAL,
                            Inches(x), Inches(y), Inches(w), Inches(h))
    solid(sp, color)
    if alpha is not None:
        fill_alpha(sp, alpha)
    if line is None:
        no_line(sp)
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    if shadow_on:
        shadow(sp)
    sp.shadow.inherit = False
    return sp


def blob(s, x, y, d, color, alpha=35):
    """Decorative soft background circle."""
    return oval(s, x, y, d, d, color, alpha=alpha, shadow_on=False)


def line_shape(s, x, y, w, h, color, weight=1.5):
    cn = s.shapes.add_connector(2, Inches(x), Inches(y), Inches(x + w), Inches(y + h))
    cn.line.color.rgb = color
    cn.line.width = Pt(weight)
    return cn


def picture(s, path, x, y, w=None, h=None):
    kw = {}
    if w is not None:
        kw["width"] = Inches(w)
    if h is not None:
        kw["height"] = Inches(h)
    pic = s.shapes.add_picture(path, Inches(x), Inches(y), **kw)
    return pic


def textbox(s, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", PP_ALIGN.LEFT)
        if "space_after" in p:
            para.space_after = Pt(p["space_after"])
        if "space_before" in p:
            para.space_before = Pt(p["space_before"])
        para.line_spacing = p.get("line", 1.05)
        for run in p["runs"]:
            r = para.add_run()
            r.text = run["t"]
            f = r.font
            f.size = Pt(run.get("size", 18))
            f.bold = run.get("bold", False)
            f.italic = run.get("italic", False)
            f.name = run.get("font", FONT)
            f.color.rgb = run.get("color", DARK)
    return tb


def run(t, size=18, bold=False, color=DARK, italic=False, font=FONT):
    return {"t": t, "size": size, "bold": bold, "color": color,
            "italic": italic, "font": font}


def para(runs, align=PP_ALIGN.LEFT, space_after=0, space_before=0, line=1.05):
    return {"runs": runs, "align": align, "space_after": space_after,
            "space_before": space_before, "line": line}


def bullet(text, accent=BLUE_D, size=18, color=DARK, space_after=9, bold=False):
    return para([run("\u25CF  ", size=size - 4, color=accent, bold=True),
                 run(text, size=size, color=color, bold=bold)],
                space_after=space_after, line=1.06)


def check(text, accent=MINT_D, size=20, space_after=14):
    return para([run("\u2713  ", size=size, color=accent, bold=True),
                 run(text, size=size, color=DARK)],
                space_after=space_after, line=1.08)


# ----------------------------------------------------------------------------
# Eyebrow / pill label
# ----------------------------------------------------------------------------
def pill(s, x, y, text, fill=BLUE, txt=DARK, w=None, size=13):
    if w is None:
        w = 0.32 + 0.092 * len(text)
    p = rrect(s, x, y, w, 0.42, fill, radius=0.5, shadow_on=False)
    tf = p.text_frame
    tf.word_wrap = False
    tf.margin_top = 0
    tf.margin_bottom = 0
    para0 = tf.paragraphs[0]
    para0.alignment = PP_ALIGN.CENTER
    r = para0.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.name = FONT
    r.font.color.rgb = txt
    return p


# ----------------------------------------------------------------------------
# Transition injection
# ----------------------------------------------------------------------------
def _transition_el(name, attrs="", spd="med"):
    if name == "morph":
        return parse_xml(
            '<mc:AlternateContent xmlns:mc="%s">'
            '<mc:Choice xmlns:p159="%s" Requires="p159">'
            '<p:transition xmlns:p="%s" xmlns:p14="%s" spd="slow" p14:dur="900">'
            '<p159:morph option="byObject"/></p:transition>'
            '</mc:Choice>'
            '<mc:Fallback>'
            '<p:transition xmlns:p="%s"><p:fade/></p:transition>'
            '</mc:Fallback></mc:AlternateContent>'
            % (NS_MC, NS_P159, NS_P, NS_P14, NS_P))
    return parse_xml(
        '<p:transition xmlns:p="%s" xmlns:p14="%s" spd="%s" p14:dur="700">'
        '<p:%s%s/></p:transition>'
        % (NS_P, NS_P14, spd, name, (" " + attrs) if attrs else ""))


def set_transition(sld, name, attrs=""):
    el = sld._element
    clr = el.find(qn("p:clrMapOvr"))
    trans = _transition_el(name, attrs)
    clr.addnext(trans)
    return trans


# ----------------------------------------------------------------------------
# Build-in (entrance) animation injection -- sequential fade
# ----------------------------------------------------------------------------
class IdGen:
    def __init__(self, start=3):
        self.n = start

    def __call__(self):
        v = self.n
        self.n += 1
        return v


def _effect(spid, mode, para_idx, delay, ids, dur=480):
    a, b, c, d, e = ids(), ids(), ids(), ids(), ids()
    if mode == "para":
        tgt = ('<p:spTgt spid="%d"><p:txEl><p:pRg st="%d" end="%d"/>'
               '</p:txEl></p:spTgt>' % (spid, para_idx, para_idx))
    else:
        tgt = '<p:spTgt spid="%d"/>' % spid
    return (
        '<p:par><p:cTn id="%d" fill="hold">'
        '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
        '<p:par><p:cTn id="%d" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        '<p:par><p:cTn id="%d" presetID="10" presetClass="entr" presetSubtype="0" '
        'fill="hold" grpId="0" nodeType="afterEffect">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '<p:tgtEl>%s</p:tgtEl>'
        '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
        '</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
        '<p:animEffect transition="in" filter="fade"><p:cBhvr>'
        '<p:cTn id="%d" dur="%d"/><p:tgtEl>%s</p:tgtEl>'
        '</p:cBhvr></p:animEffect>'
        '</p:childTnLst></p:cTn></p:par>'
        '</p:childTnLst></p:cTn></p:par>'
        '</p:childTnLst></p:cTn></p:par>'
        % (a, delay, b, c, d, tgt, e, dur, tgt)
    )


def set_build(sld, specs, first_delay=300, step_delay=180):
    """specs: ordered list of dicts:
        {'spid':int, 'mode':'shape'} or {'spid':int,'mode':'para','n':int}
    """
    ids = IdGen(3)
    effects = []
    bldps = []
    seen_bld = set()
    first = True
    for sp in specs:
        spid = sp["spid"]
        mode = sp["mode"]
        if mode == "para":
            for i in range(sp["n"]):
                delay = first_delay if first else step_delay
                effects.append(_effect(spid, "para", i, delay, ids))
                first = False
            if spid not in seen_bld:
                bldps.append('<p:bldP spid="%d" grpId="0" build="p"/>' % spid)
                seen_bld.add(spid)
        else:
            delay = first_delay if first else step_delay
            effects.append(_effect(spid, "shape", 0, delay, ids))
            first = False

    bld = ("<p:bldLst>" + "".join(bldps) + "</p:bldLst>") if bldps else ""
    timing = parse_xml(
        '<p:timing xmlns:p="%s" xmlns:a="%s">'
        '<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" '
        'nodeType="tmRoot"><p:childTnLst>'
        '<p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        '%s'
        '</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/>'
        '</p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/>'
        '</p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
        '%s</p:timing>'
        % (NS_P, NS_A, "".join(effects), bld))
    sld._element.append(timing)


def notes(sld, text):
    sld.notes_slide.notes_text_frame.text = text


# ----------------------------------------------------------------------------
# Decorative background
# ----------------------------------------------------------------------------
def background(s, color=WHITE):
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0,
                            prs.slide_width, prs.slide_height)
    solid(bg, color)
    no_line(bg)
    bg.shadow.inherit = False
    return bg


def corner_blobs(s, scheme=(BLUE, PINK, YELLOW)):
    blob(s, -1.4, -1.6, 3.6, scheme[0], alpha=40)
    blob(s, SW - 2.2, SH - 2.4, 3.8, scheme[1], alpha=34)


def sid(shape):
    return shape.shape_id


# ============================================================================
#  SLIDE 1 -- Titel  (Fade)
# ============================================================================
def slide1():
    s = slide()
    background(s)
    blob(s, -1.6, -1.8, 4.2, BLUE, alpha=45)
    blob(s, 10.2, -1.4, 3.2, YELLOW, alpha=42)
    blob(s, 9.8, 5.2, 3.6, PINK, alpha=40)

    eb = pill(s, 0.95, 1.18, "Seminarfachkurs Gl\u00fcck", fill=PINK, size=14)

    title = textbox(s, 0.9, 1.95, 6.9, 2.7, [
        para([run("Wie Schule", size=58, bold=True, color=DARK, font=FONT_HEAD)],
             line=0.98, space_after=2),
        para([run("wirkt", size=58, bold=True, color=BLUE_D, font=FONT_HEAD)],
             line=0.98),
    ])
    sub = textbox(s, 0.95, 4.35, 6.6, 1.7, [
        para([run("Strukturelle Unterschiede und ihr Effekt auf das "
                  "Wohlbefinden von Sch\u00fclerinnen und Sch\u00fclern",
                  size=21, color=GRAY)], line=1.2),
    ])
    author = textbox(s, 0.95, 6.35, 6.6, 0.6, [
        para([run("Facharbeit von Svea Timphus", size=15, bold=True, color=DARK),
              run("   \u00b7   Gymnasium Damme", size=15, color=GRAY)]),
    ])

    pic = picture(s, os.path.join(ASSETS, "cover_school.png"), 7.35, 2.25, w=5.7)

    set_transition(s, "fade")
    set_build(s, [
        {"spid": sid(eb), "mode": "shape"},
        {"spid": sid(title), "mode": "para", "n": 2},
        {"spid": sid(sub), "mode": "para", "n": 1},
        {"spid": sid(author), "mode": "para", "n": 1},
        {"spid": sid(pic), "mode": "shape"},
    ])
    notes(s,
        "Herzlich willkommen. Mein Thema lautet \u201eWie Schule wirkt\u201c \u2013 es "
        "geht um die strukturellen Unterschiede im deutschen Schulsystem und ihren "
        "Effekt auf das Wohlbefinden von Sch\u00fclerinnen und Sch\u00fclern. "
        "Schule ist weit mehr als ein Ort der Wissensvermittlung: Sie ist ein "
        "pr\u00e4gender Lebensraum, in dem wir einen Gro\u00dfteil unseres Alltags "
        "verbringen. Genau deshalb passt dieses Thema so gut in unseren "
        "Seminarfachkurs Gl\u00fcck. Auf dem Bild seht ihr schon den Kern: dieselbe "
        "Schule, aber sehr unterschiedliche Gef\u00fchle \u2013 ein Sch\u00fcler "
        "geht gut gelaunt hin, eine Sch\u00fclerin f\u00fchlt sich gestresst. "
        "Woran das liegen kann, schauen wir uns jetzt gemeinsam an.")


# ============================================================================
#  SLIDE 2 -- Warum wichtig  (Zoom)
# ============================================================================
def slide2():
    s = slide()
    background(s)
    corner_blobs(s, (BLUE, YELLOW, PINK))

    eb = pill(s, 0.95, 0.7, "Relevanz", fill=YELLOW, size=14)
    head = textbox(s, 0.9, 1.18, 7.0, 1.0, [
        para([run("Warum ist das Thema wichtig?", size=38, bold=True,
                  color=DARK, font=FONT_HEAD)])])

    bl = textbox(s, 1.0, 2.5, 6.4, 2.6, [
        bullet("Schule pr\u00e4gt den Alltag junger Menschen", BLUE_D, size=21),
        bullet("Leistungsdruck nimmt zu", PINK_D, size=21),
        bullet("Bildung beeinflusst Zukunft und Teilhabe", YELLOW_D, size=21),
        bullet("Wohlbefinden ist Teil erfolgreicher Bildung", MINT_D, size=21,
               space_after=0),
    ])

    box = rrect(s, 0.95, 5.35, 11.4, 1.5, BLUE, radius=0.18)
    qtf = box.text_frame
    qtf.word_wrap = True
    qtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    qtf.margin_left = Inches(0.35)
    qtf.margin_right = Inches(0.35)
    qp = qtf.paragraphs[0]
    qp.alignment = PP_ALIGN.CENTER
    r1 = qp.add_run(); r1.text = "Leitfrage:  "
    r1.font.size = Pt(19); r1.font.bold = True; r1.font.name = FONT
    r1.font.color.rgb = BLUE_D
    r2 = qp.add_run()
    r2.text = ("Inwiefern beeinflussen strukturelle Unterschiede im deutschen "
               "Schulsystem das Wohlbefinden von Sch\u00fclerinnen und "
               "Sch\u00fclern?")
    r2.font.size = Pt(19); r2.font.bold = True; r2.font.name = FONT
    r2.font.color.rgb = DARK

    pic = picture(s, os.path.join(ASSETS, "balance_scale.png"), 7.7, 1.95, w=5.2)

    set_transition(s, "zoom")
    set_build(s, [
        {"spid": sid(eb), "mode": "shape"},
        {"spid": sid(head), "mode": "para", "n": 1},
        {"spid": sid(pic), "mode": "shape"},
        {"spid": sid(bl), "mode": "para", "n": 4},
        {"spid": sid(box), "mode": "shape"},
    ])
    notes(s,
        "Warum lohnt sich der Blick auf dieses Thema? Schule pr\u00e4gt unseren "
        "Alltag wie kaum eine andere Institution \u2013 wir verbringen einen "
        "Gro\u00dfteil unserer Zeit dort. Gleichzeitig nimmt der Leistungsdruck "
        "sp\u00fcrbar zu, und Bildung entscheidet ma\u00dfgeblich \u00fcber "
        "Zukunftschancen und gesellschaftliche Teilhabe. Wichtig ist: Wohlbefinden "
        "ist kein netter Zusatz, sondern ein Bestandteil erfolgreicher Bildung. "
        "Die Waage im Bild bringt es auf den Punkt \u2013 Bildung und Wohlbefinden "
        "geh\u00f6ren ins Gleichgewicht. Daraus ergibt sich meine Leitfrage, die "
        "die ganze Arbeit tr\u00e4gt: Inwiefern beeinflussen strukturelle "
        "Unterschiede im deutschen Schulsystem das Wohlbefinden der "
        "Sch\u00fclerinnen und Sch\u00fcler? Um das zu beantworten, kl\u00e4ren "
        "wir zun\u00e4chst die wichtigsten Grundbegriffe.")


# ============================================================================
#  SLIDE 3 -- Grundlagen  (Move In / cover)
# ============================================================================
def slide3():
    s = slide()
    background(s)
    blob(s, -1.5, 4.4, 3.4, YELLOW, alpha=36)
    blob(s, 11.0, -1.4, 3.2, BLUE, alpha=36)

    eb = pill(s, 0.95, 0.62, "Begriffe", fill=BLUE, size=14)
    head = textbox(s, 0.9, 1.08, 8.0, 1.0, [
        para([run("Grundlagen", size=40, bold=True, color=DARK, font=FONT_HEAD)])])

    cards = []
    data = [
        ("Strukturmerkmale", BLUE, BLUE_D,
         ["Schulformen", "Notensystem", "Versetzungen", "Abschl\u00fcsse"]),
        ("Wohlbefinden", PINK, PINK_D,
         ["psychische Gesundheit", "soziale Beziehungen", "Lebenszufriedenheit"]),
        ("Bildungs-\ngerechtigkeit", YELLOW, YELLOW_D,
         ["faire Chancen \u2013 unabh\u00e4ngig von sozialer Herkunft"]),
    ]
    x = 0.95
    w = 3.78
    gap = 0.18
    for title, fill, accent, items in data:
        card = rrect(s, x, 2.35, w, 3.95, WHITE, radius=0.07,
                     line=fill, line_w=1.5)
        oval(s, x + 0.32, 2.7, 0.62, 0.62, fill, shadow_on=False)
        head_t = textbox(s, x + 0.32, 3.55, w - 0.6, 1.0, [
            para([run(title, size=21, bold=True, color=DARK)], line=0.98)])
        body = textbox(s, x + 0.34, 4.75, w - 0.62, 1.4,
                       [bullet(i, accent, size=15.5, space_after=7) for i in items])
        cards.append((card, head_t, body))
        x += w + gap

    set_transition(s, "cover", 'dir="u"')
    specs = [{"spid": sid(eb), "mode": "shape"},
             {"spid": sid(head), "mode": "para", "n": 1}]
    for card, head_t, body in cards:
        specs.append({"spid": sid(head_t), "mode": "para", "n": 1})
        n = len(body.text_frame.paragraphs)
        specs.append({"spid": sid(body), "mode": "para", "n": n})
    set_build(s, specs, step_delay=140)
    notes(s,
        "Bevor wir analysieren, drei zentrale Begriffe. Erstens die "
        "Strukturmerkmale von Schule: Damit sind die organisatorischen "
        "Rahmenbedingungen gemeint \u2013 also die verschiedenen Schulformen, das "
        "Notensystem, Versetzungsregeln und Abschl\u00fcsse. Zweitens das "
        "Wohlbefinden: Es umfasst die psychische Gesundheit, soziale Beziehungen "
        "und die allgemeine Lebenszufriedenheit. Die WHO definiert Gesundheit "
        "n\u00e4mlich nicht nur als Abwesenheit von Krankheit, sondern als "
        "k\u00f6rperliches, geistiges und soziales Wohlbefinden. Und drittens die "
        "Bildungsgerechtigkeit: also faire Bildungschancen, unabh\u00e4ngig von "
        "der sozialen Herkunft. Diese drei Begriffe sind das Werkzeug, mit dem "
        "wir jetzt das deutsche Schulsystem genauer betrachten.")


# ============================================================================
#  SLIDE 4 -- Deutsches Schulsystem  (Magic Move / morph)
# ============================================================================
def slide4():
    s = slide()
    background(s)
    blob(s, -1.6, -1.6, 3.4, BLUE, alpha=34)
    blob(s, 11.2, 5.0, 3.4, PINK, alpha=34)

    eb = pill(s, 0.95, 0.6, "Mehrgliedriges System", fill=YELLOW, size=14)
    head = textbox(s, 0.9, 1.06, 11.0, 1.0, [
        para([run("Das deutsche Schulsystem", size=38, bold=True, color=DARK,
                  font=FONT_HEAD)])])

    rows = [
        ("Gymnasium", BLUE, BLUE_D,
         "H\u00f6chste akademische Anforderungen \u2013 Ziel ist das Abitur."),
        ("Realschule", PINK, PINK_D,
         "Mittlere Schulform mit mehr Praxisbezug \u2013 Mittlerer Abschluss."),
        ("Hauptschule", YELLOW, YELLOW_D,
         "Praktisch orientiert \u2013 f\u00fchrt meist in die Berufsausbildung."),
        ("Gesamtschule", MINT, MINT_D,
         "L\u00e4ngeres gemeinsames Lernen \u2013 mehrere Abschl\u00fcsse "
         "m\u00f6glich."),
        ("Berufliche Schule", BLUE, BLUE_D,
         "Duales System \u2013 verbindet Betrieb und Schule eng miteinander."),
    ]
    shapes = []
    y = 2.35
    h = 0.82
    gap = 0.16
    for name, fill, accent, desc in rows:
        card = rrect(s, 0.95, y, 11.4, h, WHITE, radius=0.22, line=fill, line_w=1.4)
        tag = rrect(s, 1.2, y + 0.14, 0.18, h - 0.28, accent, radius=0.5,
                    shadow_on=False)
        t = textbox(s, 1.6, y, 10.5, h, [
            para([run(name + "   ", size=18, bold=True, color=DARK),
                  run(desc, size=15.5, color=GRAY)],
                 line=1.0)], anchor=MSO_ANCHOR.MIDDLE)
        shapes.append((card, t))
        y += h + gap

    set_transition(s, "morph")
    specs = [{"spid": sid(eb), "mode": "shape"},
             {"spid": sid(head), "mode": "para", "n": 1}]
    for card, t in shapes:
        specs.append({"spid": sid(card), "mode": "shape"})
        specs.append({"spid": sid(t), "mode": "para", "n": 1})
    set_build(s, specs, step_delay=120)
    notes(s,
        "Das deutsche System ist mehrgliedrig \u2013 nach der Grundschule werden "
        "die Kinder auf verschiedene Schulformen aufgeteilt. Das Gymnasium hat die "
        "h\u00f6chsten akademischen Anforderungen und f\u00fchrt zum Abitur. Die "
        "Realschule liegt in der Mitte, verbindet Allgemeinbildung mit mehr "
        "Praxis und endet mit dem Mittleren Abschluss. Die Hauptschule ist stark "
        "praktisch ausgerichtet und f\u00fchrt meist direkt in eine Ausbildung. "
        "Die Gesamtschule geht einen anderen Weg: l\u00e4ngeres gemeinsames "
        "Lernen und mehrere Abschl\u00fcsse unter einem Dach. Und schlie\u00dflich "
        "die beruflichen Schulen mit dem dualen System, das Betrieb und Schule "
        "verbindet \u2013 international gilt das als Vorzeigemodell. Wichtig ist: "
        "Diese fr\u00fche Aufteilung legt Bildungswege sehr fr\u00fch fest. Und "
        "genau das bringt uns zum Leistungsdruck.")


# ============================================================================
#  SLIDE 5 -- Leistungsdruck & Wohlbefinden  (Scale / circle)
# ============================================================================
def slide5():
    s = slide()
    background(s)
    blob(s, 10.8, -1.5, 3.4, YELLOW, alpha=34)

    eb = pill(s, 0.95, 0.58, "Kapitel 4", fill=PINK, size=14)
    head = textbox(s, 0.9, 1.04, 11.6, 1.0, [
        para([run("Leistungsdruck und Wohlbefinden", size=34, bold=True,
                  color=DARK, font=FONT_HEAD)])])

    # left column
    left = rrect(s, 0.95, 2.25, 3.35, 3.55, BLUE, radius=0.08)
    lh = textbox(s, 1.25, 2.5, 2.8, 0.6, [
        para([run("Leistungsdruck", size=20, bold=True, color=DARK)])])
    lb = textbox(s, 1.25, 3.25, 2.85, 2.4, [
        bullet("Noten", BLUE_D, size=18),
        bullet("Pr\u00fcfungen", BLUE_D, size=18),
        bullet("Konkurrenz", BLUE_D, size=18),
        bullet("Erwartungen", BLUE_D, size=18, space_after=0),
    ])

    # right column
    right = rrect(s, 4.5, 2.25, 3.35, 3.55, PINK, radius=0.08)
    rh = textbox(s, 4.8, 2.5, 2.85, 0.6, [
        para([run("Folgen", size=20, bold=True, color=DARK)])])
    rb = textbox(s, 4.8, 3.25, 2.95, 2.5, [
        bullet("Stress", PINK_D, size=18),
        bullet("Schlafprobleme", PINK_D, size=18),
        bullet("Unsicherheit", PINK_D, size=18),
        bullet("geringeres Selbstwertgef\u00fchl", PINK_D, size=18, space_after=0),
    ])

    pic = picture(s, os.path.join(ASSETS, "stressed_student.png"), 8.05, 1.95, w=5.0)

    note = rrect(s, 0.95, 6.1, 11.4, 0.95, YELLOW, radius=0.3)
    ntf = note.text_frame
    ntf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ntf.word_wrap = True
    ntf.margin_left = Inches(0.3)
    ntf.margin_right = Inches(0.3)
    np = ntf.paragraphs[0]
    np.alignment = PP_ALIGN.CENTER
    rr = np.add_run()
    rr.text = ("\u201eLeistungsdruck kann Motivation f\u00f6rdern, aber auch das "
               "Wohlbefinden beeintr\u00e4chtigen.\u201c")
    rr.font.size = Pt(18); rr.font.bold = True; rr.font.italic = True
    rr.font.name = FONT; rr.font.color.rgb = DARK

    set_transition(s, "circle")
    set_build(s, [
        {"spid": sid(eb), "mode": "shape"},
        {"spid": sid(head), "mode": "para", "n": 1},
        {"spid": sid(pic), "mode": "shape"},
        {"spid": sid(left), "mode": "shape"},
        {"spid": sid(lh), "mode": "para", "n": 1},
        {"spid": sid(lb), "mode": "para", "n": 4},
        {"spid": sid(right), "mode": "shape"},
        {"spid": sid(rh), "mode": "para", "n": 1},
        {"spid": sid(rb), "mode": "para", "n": 4},
        {"spid": sid(note), "mode": "shape"},
    ], step_delay=120)
    notes(s,
        "Jetzt zum Kern: dem Leistungsdruck. Links seht ihr, wodurch er entsteht "
        "\u2013 durch Noten, Pr\u00fcfungen, st\u00e4ndige Konkurrenz und hohe "
        "Erwartungen von Eltern, Lehrkr\u00e4ften und Gesellschaft. Rechts die "
        "m\u00f6glichen Folgen: Stress, Schlafprobleme, Unsicherheit und ein "
        "geringeres Selbstwertgef\u00fchl. Studien der WHO zeigen, dass "
        "schulischer Stress zu den h\u00e4ufigsten Belastungsfaktoren im "
        "Jugendalter geh\u00f6rt, und das Robert-Koch-Institut best\u00e4tigt das "
        "f\u00fcr Deutschland. Entscheidend ist aber der Merksatz unten: "
        "Leistungsdruck ist nicht nur negativ \u2013 er kann auch motivieren. Es "
        "kommt auf das Ma\u00df an. Ob aus Druck Motivation oder Belastung wird, "
        "h\u00e4ngt stark vom Umfeld ab \u2013 und damit vom Schulklima.")


# ============================================================================
#  SLIDE 6 -- Schulklima  (Push)
# ============================================================================
def slide6():
    s = slide()
    background(s)
    blob(s, -1.6, 4.6, 3.4, BLUE, alpha=34)
    blob(s, 11.2, -1.4, 3.2, MINT, alpha=40)

    eb = pill(s, 0.95, 0.62, "Beziehungsebene", fill=MINT, size=14)
    head = textbox(s, 0.9, 1.08, 8.0, 1.0, [
        para([run("Schulklima", size=40, bold=True, color=DARK, font=FONT_HEAD)])])
    pic = picture(s, os.path.join(ASSETS, "teacher_student.png"), 9.25, 0.35, w=3.35)

    cards = []
    data = [
        ("Lehrer-Sch\u00fcler-\nVerh\u00e4ltnis", BLUE, BLUE_D,
         ["Bezugspersonen, nicht nur Wissensvermittler",
          "Respekt reduziert Stress",
          "st\u00e4rkt die Lernmotivation"]),
        ("Mitbestimmung", PINK, PINK_D,
         ["Sch\u00fclervertretung & Klassenrat",
          "Beteiligung an Projekten",
          "st\u00e4rkt Selbstwirksamkeit"]),
        ("Feedbackkultur", YELLOW, YELLOW_D,
         ["Noten + qualitative R\u00fcckmeldung",
          "formatives Feedback begleitet",
          "f\u00f6rdert realistische Selbsteinsch\u00e4tzung"]),
    ]
    x = 0.95
    w = 3.78
    gap = 0.18
    for title, fill, accent, items in data:
        card = rrect(s, x, 2.65, w, 3.85, WHITE, radius=0.07, line=fill, line_w=1.5)
        bar = rrect(s, x, 2.65, w, 0.16, fill, radius=0.5, shadow_on=False)
        head_t = textbox(s, x + 0.3, 2.95, w - 0.55, 1.0, [
            para([run(title, size=19, bold=True, color=DARK)], line=0.98)])
        body = textbox(s, x + 0.32, 4.15, w - 0.6, 2.2,
                       [bullet(i, accent, size=14.5, space_after=8) for i in items])
        cards.append((card, head_t, body))
        x += w + gap

    set_transition(s, "push", 'dir="l"')
    specs = [{"spid": sid(eb), "mode": "shape"},
             {"spid": sid(head), "mode": "para", "n": 1},
             {"spid": sid(pic), "mode": "shape"}]
    for card, head_t, body in cards:
        specs.append({"spid": sid(head_t), "mode": "para", "n": 1})
        specs.append({"spid": sid(body), "mode": "para", "n": 3})
    set_build(s, specs, step_delay=130)
    notes(s,
        "Neben dem Druck entscheidet vor allem das Schulklima \u00fcber unser "
        "Wohlbefinden \u2013 also die Qualit\u00e4t der sozialen Beziehungen in "
        "der Schule. Drei Faktoren sind besonders wichtig. Erstens das "
        "Lehrer-Sch\u00fcler-Verh\u00e4ltnis: Lehrkr\u00e4fte sind Bezugspersonen, "
        "nicht nur Wissensvermittler. Ein respektvolles Verh\u00e4ltnis senkt "
        "Stress und steigert die Motivation. Zweitens die Mitbestimmung \u2013 "
        "etwa \u00fcber Sch\u00fclervertretung, Klassenrat oder Projekte. Wer "
        "mitentscheiden darf, erlebt Selbstwirksamkeit und f\u00fchlt sich "
        "ernst genommen. Drittens die Feedbackkultur: Neben Noten helfen "
        "qualitative R\u00fcckmeldungen, die eigenen St\u00e4rken realistisch "
        "einzusch\u00e4tzen. Ein gutes Klima ist also ein echter Schutzfaktor. "
        "Doch wie erlebt ihr das eigentlich selbst? Das frage ich euch jetzt.")


# ============================================================================
#  SLIDE 7 -- Publikum einbinden  (Dissolve)
# ============================================================================
def slide7():
    s = slide()
    background(s)
    blob(s, -1.5, -1.6, 3.6, PINK, alpha=38)
    blob(s, 11.0, 5.0, 3.4, BLUE, alpha=38)

    eb = pill(s, 5.45, 0.7, "Kurze Umfrage", fill=YELLOW, size=15)
    q = textbox(s, 1.4, 1.45, 10.5, 1.7, [
        para([run("Was beeinflusst euer Wohlbefinden in der Schule am "
                  "st\u00e4rksten?", size=32, bold=True, color=DARK,
                  font=FONT_HEAD)], align=PP_ALIGN.CENTER, line=1.06)])

    btns = []
    data = [
        ("\U0001F4DA  Notendruck", BLUE),
        ("\U0001F469\u200D\U0001F3EB  Lehrkr\u00e4fte", PINK),
        ("\U0001F465  Mitsch\u00fcler", YELLOW),
        ("\U0001F552  Zeitstress", MINT),
    ]
    x = 1.05
    w = 2.7
    gap = 0.24
    for text, fill in data:
        b = rrect(s, x, 3.45, w, 1.25, fill, radius=0.22)
        tf = b.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        rr = p0.add_run(); rr.text = text
        rr.font.size = Pt(18); rr.font.bold = True; rr.font.name = FONT
        rr.font.color.rgb = DARK
        btns.append(b)
        x += w + gap

    hand = textbox(s, 1.4, 4.95, 10.5, 0.6, [
        para([run("Kurzes Handzeichen \u2013 was trifft auf euch zu?",
                  size=18, color=GRAY)], align=PP_ALIGN.CENTER)])

    reveal = rrect(s, 2.4, 5.75, 8.5, 1.05, DARK, radius=0.3)
    rtf = reveal.text_frame
    rtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    rtf.word_wrap = True
    rtf.margin_left = Inches(0.3); rtf.margin_right = Inches(0.3)
    rp = rtf.paragraphs[0]
    rp.alignment = PP_ALIGN.CENTER
    rr = rp.add_run()
    rr.text = ("Alle Faktoren spielen laut Facharbeit eine wichtige Rolle.")
    rr.font.size = Pt(19); rr.font.bold = True; rr.font.name = FONT
    rr.font.color.rgb = WHITE

    set_transition(s, "dissolve")
    set_build(s, [
        {"spid": sid(eb), "mode": "shape"},
        {"spid": sid(q), "mode": "para", "n": 1},
        {"spid": sid(btns[0]), "mode": "shape"},
        {"spid": sid(btns[1]), "mode": "shape"},
        {"spid": sid(btns[2]), "mode": "shape"},
        {"spid": sid(btns[3]), "mode": "shape"},
        {"spid": sid(hand), "mode": "para", "n": 1},
        {"spid": sid(reveal), "mode": "shape"},
    ], step_delay=160)
    notes(s,
        "An dieser Stelle m\u00f6chte ich euch kurz einbinden. Meine Frage: Was "
        "beeinflusst euer Wohlbefinden in der Schule am st\u00e4rksten? Ihr habt "
        "vier M\u00f6glichkeiten \u2013 der Notendruck, die Lehrkr\u00e4fte, die "
        "Mitsch\u00fclerinnen und Mitsch\u00fcler oder der Zeitstress. \u00dcberlegt "
        "kurz, und dann ein kurzes Handzeichen: Wer w\u00fcrde sich am ehesten "
        "f\u00fcr Notendruck entscheiden? Wer f\u00fcr die Lehrkr\u00e4fte? Wer "
        "f\u00fcr die Mitsch\u00fcler? Und wer f\u00fcr Zeitstress? "
        "(Kurz Reaktionen abwarten.) Spannend \u2013 und genau das passt zum "
        "Ergebnis meiner Arbeit: Es gibt nicht den einen Faktor. Alle diese "
        "Aspekte spielen laut der Facharbeit zusammen eine wichtige Rolle. Wie "
        "andere L\u00e4nder damit umgehen, zeigt der n\u00e4chste Blick \u00fcber "
        "die Grenze.")


# ============================================================================
#  SLIDE 8 -- Internationaler Vergleich  (Pop / newsflash)
# ============================================================================
def slide8():
    s = slide()
    background(s)
    blob(s, -1.5, -1.5, 3.2, BLUE, alpha=34)
    blob(s, 11.2, 5.0, 3.4, YELLOW, alpha=36)

    eb = pill(s, 0.95, 0.6, "Blick \u00fcber die Grenze", fill=BLUE, size=14)
    head = textbox(s, 0.9, 1.06, 8.5, 1.0, [
        para([run("Internationaler Vergleich", size=38, bold=True, color=DARK,
                  font=FONT_HEAD)])])
    pic = picture(s, os.path.join(ASSETS, "international.png"), 9.55, 0.45, w=3.25)

    cards = []
    data = [
        ("Deutschland", BLUE, BLUE_D,
         ["fr\u00fche Aufteilung", "viele Noten & Pr\u00fcfungen"]),
        ("Finnland", PINK, PINK_D,
         ["l\u00e4ngeres gemeinsames Lernen",
          "sp\u00e4tere Leistungsdifferenzierung"]),
        ("Angels\u00e4chsische\nSysteme", YELLOW, YELLOW_D,
         ["mehr Praxis & Projekte", "mehr Wahlm\u00f6glichkeiten"]),
    ]
    x = 0.95
    w = 3.78
    gap = 0.18
    for title, fill, accent, items in data:
        card = rrect(s, x, 2.65, w, 3.5, WHITE, radius=0.08, line=fill, line_w=1.5)
        oval(s, x + 0.32, 2.98, 0.55, 0.55, fill, shadow_on=False)
        head_t = textbox(s, x + 0.3, 3.65, w - 0.55, 1.0, [
            para([run(title, size=19, bold=True, color=DARK)], line=0.98)])
        body = textbox(s, x + 0.32, 4.9, w - 0.6, 1.2,
                       [bullet(i, accent, size=15.5, space_after=8) for i in items])
        cards.append((card, head_t, body))
        x += w + gap

    foot = rrect(s, 3.4, 6.5, 6.5, 0.7, MINT, radius=0.4)
    ftf = foot.text_frame; ftf.vertical_anchor = MSO_ANCHOR.MIDDLE
    fp = ftf.paragraphs[0]; fp.alignment = PP_ALIGN.CENTER
    fr = fp.add_run(); fr.text = "Kein System ist perfekt."
    fr.font.size = Pt(18); fr.font.bold = True; fr.font.name = FONT
    fr.font.color.rgb = DARK

    set_transition(s, "newsflash")
    specs = [{"spid": sid(eb), "mode": "shape"},
             {"spid": sid(head), "mode": "para", "n": 1},
             {"spid": sid(pic), "mode": "shape"}]
    for card, head_t, body in cards:
        specs.append({"spid": sid(head_t), "mode": "para", "n": 1})
        specs.append({"spid": sid(body), "mode": "para", "n": 2})
    specs.append({"spid": sid(foot), "mode": "shape"})
    set_build(s, specs, step_delay=130)
    notes(s,
        "Um Deutschland einzuordnen, hilft der internationale Vergleich. In "
        "Deutschland teilen wir die Sch\u00fclerinnen und Sch\u00fcler fr\u00fch "
        "auf, und Noten sowie Pr\u00fcfungen spielen eine sehr gro\u00dfe Rolle. "
        "Finnland macht es anders: Dort lernen die Kinder bis zur neunten Klasse "
        "gemeinsam, die Leistungsdifferenzierung kommt deutlich sp\u00e4ter, und "
        "es wird seltener standardisiert getestet. Finnland erzielt damit \u00fcber "
        "Jahre sehr gute Ergebnisse bei geringeren sozialen Unterschieden. Die "
        "angels\u00e4chsischen Systeme wiederum setzen st\u00e4rker auf Praxis, "
        "Projekte und Wahlm\u00f6glichkeiten. Wichtig bleibt aber: Kein System ist "
        "perfekt \u2013 jedes muss unterschiedliche gesellschaftliche Aufgaben "
        "erf\u00fcllen. Mit diesem Blick bewerten wir nun das deutsche System.")


# ============================================================================
#  SLIDE 9 -- Staerken & Schwaechen  (Slide / pull)
# ============================================================================
def slide9():
    s = slide()
    background(s)
    blob(s, 5.6, -1.7, 3.4, YELLOW, alpha=32)

    eb = pill(s, 0.95, 0.62, "Kritische Bewertung", fill=PINK, size=14)
    head = textbox(s, 0.9, 1.08, 10.0, 1.0, [
        para([run("St\u00e4rken und Schw\u00e4chen", size=40, bold=True,
                  color=DARK, font=FONT_HEAD)])])

    # Staerken (green)
    left = rrect(s, 0.95, 2.45, 5.55, 4.2, MINT, radius=0.06)
    lh = textbox(s, 1.35, 2.8, 4.8, 0.7, [
        para([run("\u2713  St\u00e4rken", size=24, bold=True, color=MINT_D)])])
    lb = textbox(s, 1.4, 3.75, 4.85, 2.7, [
        bullet("verschiedene Bildungswege", MINT_D, size=19, space_after=14),
        bullet("duales Ausbildungssystem", MINT_D, size=19, space_after=14),
        bullet("hohe akademische Standards", MINT_D, size=19, space_after=0),
    ])

    # Schwaechen (pink)
    right = rrect(s, 6.85, 2.45, 5.55, 4.2, PINK, radius=0.06)
    rh = textbox(s, 7.25, 2.8, 4.8, 0.7, [
        para([run("\u2715  Schw\u00e4chen", size=24, bold=True, color=PINK_D)])])
    rb = textbox(s, 7.3, 3.75, 4.85, 2.7, [
        bullet("fr\u00fche Selektion", PINK_D, size=19, space_after=14),
        bullet("Leistungsdruck", PINK_D, size=19, space_after=14),
        bullet("soziale Ungleichheiten", PINK_D, size=19, space_after=0),
    ])

    set_transition(s, "pull", 'dir="r"')
    set_build(s, [
        {"spid": sid(eb), "mode": "shape"},
        {"spid": sid(head), "mode": "para", "n": 1},
        {"spid": sid(left), "mode": "shape"},
        {"spid": sid(lh), "mode": "para", "n": 1},
        {"spid": sid(lb), "mode": "para", "n": 3},
        {"spid": sid(right), "mode": "shape"},
        {"spid": sid(rh), "mode": "para", "n": 1},
        {"spid": sid(rb), "mode": "para", "n": 3},
    ], step_delay=130)
    notes(s,
        "Fassen wir Deutschland kritisch zusammen \u2013 mit St\u00e4rken und "
        "Schw\u00e4chen. Zu den St\u00e4rken: Das System bietet verschiedene "
        "Bildungswege, sodass unterschiedliche Begabungen gef\u00f6rdert werden "
        "k\u00f6nnen. Das duale Ausbildungssystem gilt international als "
        "Erfolgsmodell und sorgt f\u00fcr eine niedrige Jugendarbeitslosigkeit. "
        "Und besonders die Gymnasien haben hohe akademische Standards. "
        "Dem gegen\u00fcber stehen klare Schw\u00e4chen: Die fr\u00fche Selektion "
        "schon nach der vierten Klasse \u2013 zu einem Zeitpunkt, an dem soziale "
        "Herkunft stark mitwirkt. Dazu der hohe Leistungsdruck und in der Folge "
        "soziale Ungleichheiten, die das System eher verst\u00e4rken als "
        "ausgleichen kann. Genau hier setzen Reform\u00fcberlegungen an.")


# ============================================================================
#  SLIDE 10 -- Reformansaetze  (Rotate / wheel)
# ============================================================================
def slide10():
    s = slide()
    background(s)
    blob(s, -1.5, 4.6, 3.4, BLUE, alpha=34)
    blob(s, 11.2, -1.4, 3.2, PINK, alpha=34)

    eb = pill(s, 0.95, 0.6, "Blick nach vorn", fill=YELLOW, size=14)
    head = textbox(s, 0.9, 1.06, 8.5, 1.0, [
        para([run("Reformans\u00e4tze", size=40, bold=True, color=DARK,
                  font=FONT_HEAD)])])
    pic = picture(s, os.path.join(ASSETS, "future_arrow.png"), 9.65, 0.45, w=3.25)

    cards = []
    data = [
        ("Mehr\nLebens-\nkompetenzen", BLUE, BLUE_D,
         "Finanzbildung, Medien- und Alltagskompetenzen im Lehrplan."),
        ("Sp\u00e4tere\nDifferen-\nzierung", PINK, PINK_D,
         "L\u00e4nger gemeinsam lernen \u2013 mehr Zeit zur Entwicklung."),
        ("Alternative\nBewertung", YELLOW, YELLOW_D,
         "Lernfeedback, Portfolio- und projektbasierte Bewertung."),
    ]
    x = 1.25
    cw = 3.5
    gap = 0.55
    for title, fill, accent, desc in data:
        circ = oval(s, x + 0.45, 2.75, 2.6, 2.6, fill, shadow_on=True)
        ctf = circ.text_frame; ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
        ctf.word_wrap = True
        cp = ctf.paragraphs[0]; cp.alignment = PP_ALIGN.CENTER
        cr = cp.add_run(); cr.text = title
        cr.font.size = Pt(15); cr.font.bold = True; cr.font.name = FONT
        cr.font.color.rgb = DARK
        cp.line_spacing = 1.0
        desc_t = textbox(s, x, 5.65, cw, 1.4, [
            para([run(desc, size=15, color=GRAY)], align=PP_ALIGN.CENTER, line=1.1)])
        cards.append((circ, desc_t))
        x += cw + gap

    set_transition(s, "wheel", 'spokes="2"')
    specs = [{"spid": sid(eb), "mode": "shape"},
             {"spid": sid(head), "mode": "para", "n": 1},
             {"spid": sid(pic), "mode": "shape"}]
    for circ, desc_t in cards:
        specs.append({"spid": sid(circ), "mode": "shape"})
        specs.append({"spid": sid(desc_t), "mode": "para", "n": 1})
    set_build(s, specs, step_delay=150)
    notes(s,
        "Aus den Schw\u00e4chen ergeben sich drei vielversprechende "
        "Reformans\u00e4tze. Erstens: mehr Lebenskompetenzen im Lehrplan \u2013 "
        "also Finanzbildung, Medienkompetenz und praktische Alltagsf\u00e4higkeiten, "
        "damit Schule besser aufs echte Leben vorbereitet. Zweitens: eine "
        "sp\u00e4tere Leistungsdifferenzierung nach skandinavischem Vorbild. Wenn "
        "Kinder l\u00e4nger gemeinsam lernen, haben sie mehr Zeit, ihre "
        "F\u00e4higkeiten zu entwickeln und fundierter zu entscheiden. Und "
        "drittens: alternative Bewertungsformen \u2013 ausf\u00fchrliches "
        "Lernfeedback, Portfolioarbeit oder projektbasierte Bewertung, die den "
        "Lernprozess in den Mittelpunkt stellen statt nur das Pr\u00fcfungsergebnis. "
        "Diese Ans\u00e4tze zielen darauf, Bildungsqualit\u00e4t und Wohlbefinden "
        "zu verbinden. Damit komme ich zum Fazit.")


# ============================================================================
#  SLIDE 11 -- Fazit  (Fade)
# ============================================================================
def slide11():
    s = slide()
    background(s)
    blob(s, -1.6, -1.7, 4.0, BLUE, alpha=42)
    blob(s, 10.6, 5.0, 3.6, PINK, alpha=42)
    blob(s, 10.8, -1.5, 3.0, YELLOW, alpha=40)

    eb = pill(s, 0.95, 0.95, "Fazit", fill=MINT, size=14)
    head = textbox(s, 0.9, 1.5, 7.3, 2.0, [
        para([run("Schule beeinflusst mehr als nur Noten.", size=40, bold=True,
                  color=DARK, font=FONT_HEAD)], line=1.04)])

    body = textbox(s, 0.95, 3.7, 7.0, 2.5, [
        check("Strukturen wirken auf das Wohlbefinden."),
        check("Leistungsdruck und Schulklima spielen eine zentrale Rolle."),
        check("Reformen k\u00f6nnen Bildung und Gl\u00fcck verbinden.",
              space_after=0),
    ])

    thanks = rrect(s, 0.95, 6.25, 6.7, 0.85, BLUE, radius=0.4)
    ttf = thanks.text_frame; ttf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ttf.margin_left = Inches(0.3)
    tp = ttf.paragraphs[0]; tp.alignment = PP_ALIGN.CENTER
    tr = tp.add_run(); tr.text = "Danke f\u00fcr eure Aufmerksamkeit!"
    tr.font.size = Pt(22); tr.font.bold = True; tr.font.name = FONT
    tr.font.color.rgb = DARK

    pic = picture(s, os.path.join(ASSETS, "closing_calm.png"), 7.95, 2.5, w=5.1)

    set_transition(s, "fade")
    set_build(s, [
        {"spid": sid(eb), "mode": "shape"},
        {"spid": sid(head), "mode": "para", "n": 1},
        {"spid": sid(pic), "mode": "shape"},
        {"spid": sid(body), "mode": "para", "n": 3},
        {"spid": sid(thanks), "mode": "shape"},
    ], step_delay=200)
    notes(s,
        "Zur\u00fcck zur Leitfrage: Inwiefern beeinflussen strukturelle "
        "Unterschiede das Wohlbefinden? Mein Fazit lautet \u2013 Schule "
        "beeinflusst weit mehr als nur Noten. Erstens: Strukturen wie die "
        "fr\u00fche Aufteilung und das Notensystem wirken nachweislich auf das "
        "Wohlbefinden. Zweitens: Leistungsdruck und Schulklima spielen dabei eine "
        "zentrale Rolle \u2013 sie k\u00f6nnen belasten, aber bei gutem Umfeld "
        "auch sch\u00fctzen und motivieren. Und drittens: Durchdachte Reformen "
        "k\u00f6nnten Bildung und Gl\u00fcck st\u00e4rker miteinander verbinden. "
        "Wohlbefinden ist eben nicht nur Privatsache, sondern h\u00e4ngt stark "
        "von den Rahmenbedingungen ab, die wir gestalten k\u00f6nnen. Vielen Dank "
        "f\u00fcr eure Aufmerksamkeit \u2013 gerne beantworte ich jetzt eure "
        "Fragen.")


# ----------------------------------------------------------------------------
for fn in (slide1, slide2, slide3, slide4, slide5, slide6,
           slide7, slide8, slide9, slide10, slide11):
    fn()

prs.save(OUT)
print("Saved:", OUT)
print("Slides:", len(prs.slides._sldIdLst))

# ----------------------------------------------------------------------------
# Export presenter notes to a readable Markdown file
# ----------------------------------------------------------------------------
TITLES = [
    "Titel \u2013 Wie Schule wirkt",
    "Warum ist das Thema wichtig?",
    "Grundlagen",
    "Das deutsche Schulsystem",
    "Leistungsdruck und Wohlbefinden",
    "Schulklima",
    "Publikum einbinden \u2013 Kurze Umfrage",
    "Internationaler Vergleich",
    "St\u00e4rken und Schw\u00e4chen",
    "Reformans\u00e4tze",
    "Fazit",
]
TRANSITIONS = ["Fade", "Zoom", "Move In (Cover)", "Magic Move (Morph)",
               "Scale (Circle)", "Push", "Dissolve", "Pop (Newsflash)",
               "Slide (Pull)", "Rotate (Wheel)", "Fade"]

notes_path = os.path.join(os.path.dirname(OUT), "Sprechnotizen.md")
with open(notes_path, "w", encoding="utf-8") as fh:
    fh.write("# Sprechnotizen \u2013 \u201eWie Schule wirkt\u201c\n\n")
    fh.write("Seminarfachkurs Gl\u00fcck \u00b7 Facharbeit von Svea Timphus \u00b7 "
             "Vortragsdauer ca. 15 Minuten\n\n")
    fh.write("> Pro Folie ca. 1 Minute. Die Notizen sind frei vortragbar "
             "formuliert und enthalten \u00dcberg\u00e4nge zur n\u00e4chsten "
             "Folie. Sie sind zus\u00e4tzlich direkt in der Keynote-/PowerPoint-"
             "Datei als Pr\u00e4sentatornotizen hinterlegt.\n\n---\n\n")
    for i, sld in enumerate(prs.slides):
        txt = sld.notes_slide.notes_text_frame.text if sld.has_notes_slide else ""
        fh.write("## Folie %d \u2013 %s\n\n" % (i + 1, TITLES[i]))
        fh.write("*\u00dcbergang: %s*\n\n" % TRANSITIONS[i])
        fh.write(txt.strip() + "\n\n---\n\n")
print("Saved:", notes_path)
