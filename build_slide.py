from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

SW, SH = 1672.0, 941.0
prs = Presentation()
prs.slide_width = Emu(12192000); prs.slide_height = Emu(6858000)
slide = prs.slides.add_slide(prs.slide_layouts[6])
EX = 12192000 / SW; EY = 6858000 / SH
def X(v): return Emu(int(round(v * EX)))
def Y(v): return Emu(int(round(v * EY)))
def rgb(h): return RGBColor.from_string(h)
FONT = "Arial"

def rect(x, y, w, h, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE, radius=None, shadow=False):
    s = slide.shapes.add_shape(shape, X(x), Y(y), X(w), Y(h))
    if fill: s.fill.solid(); s.fill.fore_color.rgb = rgb(fill)
    else: s.fill.background()
    if line: s.line.color.rgb = rgb(line); s.line.width = Pt(lw)
    else: s.line.fill.background()
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = min(0.5, radius / min(w, h))
    spPr = s._element.spPr
    for e in spPr.findall(qn('a:effectLst')): spPr.remove(e)
    eff = etree.SubElement(spPr, qn('a:effectLst'))
    if shadow:
        sh = etree.SubElement(eff, qn('a:outerShdw'), blurRad="63500", dist="19050", dir="5400000", algn="t", rotWithShape="0")
        c = etree.SubElement(sh, qn('a:srgbClr'), val="000000"); etree.SubElement(c, qn('a:alpha'), val="14000")
    # kill style-based text/effects
    return s

def text(x, y, w, h, runs, size=12, color="000000", bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=None, name=None):
    """runs: str or list of lines; each line str or list of (text, dict) runs"""
    tb = slide.shapes.add_textbox(X(x), Y(y), X(w), Y(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    lines = runs if isinstance(runs, list) else [runs]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if spacing: p.line_spacing = (Pt(spacing) if spacing > 3 else spacing)
        parts = ln if isinstance(ln, list) else [(ln, {})]
        for t, o in parts:
            r = p.add_run(); r.text = t
            f = r.font; f.name = o.get('font', FONT); f.size = Pt(o.get('size', size))
            f.bold = o.get('bold', bold); f.color.rgb = rgb(o.get('color', color))
            rPr = r._r.get_or_add_rPr()
            for tag in ('a:ea', 'a:cs'):
                e = etree.SubElement(rPr, qn(tag)); e.set('typeface', o.get('font', FONT))
    return tb

def pic(path, x, y, w, h):
    return slide.shapes.add_picture(path, X(x), Y(y), X(w), Y(h))

def line(x1, y1, x2, y2, color, wpt, dash=None, head=None, tail=None):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, X(x1), Y(y1), X(x2), Y(y2))
    c.line.color.rgb = rgb(color); c.line.width = Pt(wpt)
    ln = c.line._get_or_add_ln()
    if dash:
        d = etree.SubElement(ln, qn('a:prstDash')); d.set('val', dash)
    if head:
        e = etree.SubElement(ln, qn('a:headEnd')); e.set('type', head); e.set('w', 'med'); e.set('len', 'med')
    if tail:
        e = etree.SubElement(ln, qn('a:tailEnd')); e.set('type', tail); e.set('w', 'med'); e.set('len', 'med')
    return c

# ---------- background ----------
bg = slide.background.fill; bg.solid(); bg.fore_color.rgb = rgb("FFFFFF")

# ---------- header ----------
pill = rect(42, 20, 173, 65, fill="131D2D", shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=32.5)
tf = pill.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "DarcARK"; r.font.name = FONT; r.font.size = Pt(14.1); r.font.bold = False; r.font.color.rgb = rgb("FFFFFF")

text(300, 8, 1072, 80, "INNOVATION & UNIQUENESS", size=31.2, color="0B1B4D", bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(300, 72, 1072, 44, "3D LiDAR → Adaptive 2.5D Conversion Engine", size=17.2, color="333B4A", bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# gradient underline
gl = rect(585, 125, 503, 6, fill="0A74D6", shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=3)
spPr = gl._element.spPr
sf = spPr.find(qn('a:solidFill')); idx = list(spPr).index(sf); spPr.remove(sf)
grad = etree.Element(qn('a:gradFill'), rotWithShape="1")
gs = etree.SubElement(grad, qn('a:gsLst'))
for pos, col in ((0, "0A74D6"), (50000, "19B6A6"), (100000, "64D94F")):
    g = etree.SubElement(gs, qn('a:gs'), pos=str(pos)); etree.SubElement(g, qn('a:srgbClr'), val=col)
etree.SubElement(grad, qn('a:lin'), ang="0", scaled="0")
spPr.insert(idx, grad)

# SIH logo + text
pic('logo.png', 1372, 4, 112, 122)
text(1490, 22, 170, 30, "SMART INDIA", size=12.7, color="0B2A5B", bold=True)
text(1490, 50, 170, 30, "HACKATHON", size=12.7, color="0B2A5B", bold=True)
text(1490, 80, 170, 30, "2026", size=12.7, color="0B2A5B", bold=True)

# ---------- panels ----------
def panel(x, y, w, h, fill, border):
    return rect(x, y, w, h, fill=fill, line=border, lw=1.25, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=22, shadow=False)

def badge(cx, cy, d, fill, n):
    b = rect(cx - d/2, cy - d/2, d, d, fill=fill, shape=MSO_SHAPE.OVAL)
    tf = b.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = n; r.font.name = FONT; r.font.size = Pt(19); r.font.bold = True; r.font.color.rgb = rgb("FFFFFF")

TITLE = "0B1B4D"; BODY = "1B2433"

# Panel 1
panel(26, 150, 604, 357, "F5FAFE", "7FBDEC")
badge(70, 182, 54, "0A74D6", "1")
pic('p1g.png', 58, 248, 534, 194)
text(117, 160, 500, 44, "Distance-Adaptive Foveation", size=16.7, color=TITLE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
text(117, 203, 540, 80, ["Cell resolution changes with distance —", "fine cells near the sensor and progressively", "coarser cells farther away."], size=12.5, color=BODY, spacing=14.1)
for lx, lw_, t1, t2 in ((150, 76, "Fine", "(0 – 10 m)"), (270, 100, "Medium", "(10 – 40 m)"), (420, 108, "Coarse", "(40 – 120 m)")):
    text(lx - 20, 442, lw_ + 40, 24, t1, size=9.2, color=BODY, bold=False, align=PP_ALIGN.CENTER)
    text(lx - 20, 464, lw_ + 40, 24, t2, size=9.2, color=BODY, align=PP_ALIGN.CENTER)

# Panel 2
panel(1041, 150, 605, 357, "F3F9F3", "A3D7AC")
badge(1085, 182, 54, "17963F", "2")
text(1127, 158, 500, 44, "Macro N-Block Alignment", size=16.7, color=TITLE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
text(1127, 203, 515, 100, ["The space is divided into macro blocks, where each", "block is assigned a resolution level. Cell sizes remain", "spatially aligned, avoiding overlap and duplicate", "boundaries."], size=12.5, color=BODY, spacing=14.1)
pic('p2g.png', 1060, 357, 562, 135)
# arrows + dashed verticals above graphic
for xa, xb, col in ((1127, 1267, "1E7FD6"), (1267, 1410, "1FA34A"), (1410, 1598, "F0951C")):
    line(xa, 355, xb, 355, col, 1.75, head='triangle', tail='triangle')
for xd in (1127, 1267, 1410, 1598):
    line(xd, 340, xd, 358, "1B2433", 1.0, dash='dash')
text(1140, 314, 110, 40, ["Fine", "Block"], size=9.2, color=BODY, align=PP_ALIGN.CENTER, spacing=0.95)
text(1290, 311, 110, 40, ["Medium", "Block"], size=9.2, color=BODY, align=PP_ALIGN.CENTER, spacing=0.95)
text(1452, 306, 110, 40, ["Coarse", "Block"], size=9.2, color=BODY, align=PP_ALIGN.CENTER, spacing=0.95)

# Panel 3
panel(26, 527, 604, 350, "FEFCF6", "FAD98A")
badge(70, 562, 54, "F0A81C", "3")
text(122, 540, 500, 44, "Height-Preserving 2.5D Cells", size=16.7, color=TITLE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
text(122, 583, 505, 60, ["Each 2.5D cell stores both spatial and vertical", "information for accurate environment representation."], size=12.5, color=BODY, spacing=14.1)
pic('cube.png', 104, 660, 178, 190)
text(50, 738, 60, 44, ["Cell", "(x, y)"], size=10, color=BODY, align=PP_ALIGN.CENTER, spacing=0.95)
rect(294, 648, 314, 207, fill="E9F3FC", line="BBD6EE", lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=14)
items = ["Min Z (lowest height)", "Max Z (highest height)", "Mean Z (average height)", "Ground height",
         "Relative height (to ground)", "Intensity (reflectivity)", "Occupancy (object presence)"]
for i, t in enumerate(items):
    cy = 673 + i * 25.8
    rect(316 - 5.5, cy - 5.5, 11, 11, fill="0A74D6", shape=MSO_SHAPE.OVAL)
    text(345, cy - 13, 260, 26, t, size=10.3, color=BODY, anchor=MSO_ANCHOR.MIDDLE)

# Center labels
pic('center.png', 636, 180, 400, 456)
text(736, 638, 200, 28, "3D LiDAR", size=12, color="0B1B4D", bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
pic('arrow.png', 820, 666, 32, 46)
pic('grid.png', 660, 696, 354, 104)
text(636, 800, 400, 30, "Adaptive 2.5D Representation", size=12.4, color="0B1B4D", bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Panel 4
panel(1041, 527, 605, 350, "F7F6FD", "B3A6E8")
badge(1085, 561, 54, "6D48CC", "4")
text(1136, 538, 500, 44, "3D vs 2.5D Validation", size=16.7, color=TITLE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
text(1136, 577, 505, 80, ["The same scene is evaluated in both 3D and 2.5D", "representations for terrain analysis, object detection", "and adaptive spatial representation."], size=12.5, color=BODY, spacing=13.2)
pic('th1.png', 1074, 648, 192, 86)
pic('th2.png', 1398, 648, 224, 84)
g1 = rect(1090, 738, 171, 39, fill="DCDFE3", shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=8)
g2 = rect(1418, 738, 196, 39, fill="D6D9DC", shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=8)
text(1090, 738, 171, 39, ["3D Point Cloud", "(Reference)"], size=8.3, color=BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, spacing=0.9)
text(1418, 738, 196, 39, ["Adaptive 2.5D", "Representation"], size=8.3, color=BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, spacing=0.9)
line(1302, 696, 1374, 696, "1B2433", 1.5, head='triangle', tail='triangle')
text(1282, 717, 112, 44, ["Evaluation", "& Comparison"], size=8.7, color=BODY, align=PP_ALIGN.CENTER, spacing=0.95)
rect(1060, 790, 566, 73, fill="E7F3FD", line="C4DCF0", lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=14)
pic('ic1.png', 1078, 806, 44, 44)
pic('ic2.png', 1244, 802, 64, 48)
pic('ic3.png', 1444, 802, 50, 48)
text(1138, 806, 120, 46, [[("Memory", {})], [("Usage ↓", {})]], size=10.3, color=BODY, spacing=0.95)
text(1316, 806, 130, 46, [[("Processing", {})], [("Latency ↓", {})]], size=10.3, color=BODY, spacing=0.95)
text(1508, 806, 120, 46, [[("Detection /", {})], [("Accuracy ↔", {})]], size=10.3, color=BODY, spacing=0.95)

# Footer
rect(0, 892, 1672, 49, fill="0166BB")
text(736, 905, 300, 28, [[("@SIH", {}), ("   Idea Submission", {})]], size=10.9, color="FFFFFF", anchor=MSO_ANCHOR.MIDDLE)
text(1580, 903, 40, 28, "4", size=10.9, color="FFFFFF", align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

prs.save('/mnt/user-data/outputs/DarcARK_Innovation_Uniqueness.pptx')
