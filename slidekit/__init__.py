"""Thin helpers over python-pptx for technical slides. Use what helps and ignore the rest.

Everything stays native and follows the template's theme: colors are theme slots, never hex,
and fonts are inherited, so swapping the template restyles the deck and PowerPoint can edit it all.

Positions and sizes are in inches.

Colors are "ink" (text), "paper" (background) and "a1" to "a6" (the theme's accents). A suffix
changes them: "a3-" is darker, "a3+" is a pale tint, "a3~" is muted (half strength, for dimming).

Text markup: **bold**, *italic*, _{sub}, ^{sup}, [a3:colored]. Newlines start a new paragraph.
"""
import copy
import functools
import io
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path

import pypandoc
from lxml import etree
from pptx import Presentation
from pptx.enum.chart import XL_TICK_MARK
from pptx.enum.dml import MSO_THEME_COLOR as TC
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, PP_PLACEHOLDER as PH
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

COLORS = {"ink": TC.TEXT_1, "paper": TC.BACKGROUND_1, "a1": TC.ACCENT_1, "a2": TC.ACCENT_2,
          "a3": TC.ACCENT_3, "a4": TC.ACCENT_4, "a5": TC.ACCENT_5, "a6": TC.ACCENT_6}
ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
ANCHOR = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}
TOKEN = re.compile(r"\*\*(?P<b>.+?)\*\*|\*(?P<i>.+?)\*|_\{(?P<sub>.+?)\}|\^\{(?P<sup>.+?)\}"
                   r"|\[(?P<c>\w+[-+~]?):(?P<ct>.+?)\]")
OPACITY = {"+": 18, "~": 45}  # percent, for pale tints and muted colors
TEMPLATE = str(Path(__file__).parent.parent / "assets" / "clean.pptx")


def I(inches):
    return Emu(int(round(inches * 914400)))


def paint(color_format, name):
    """Set a theme color. The pale tint is the color at low opacity, which works on any background."""
    color_format.theme_color = COLORS[name.rstrip("+-~")]
    if name[-1] in OPACITY:
        clr = color_format._xFill.find(qn("a:schemeClr"))
        etree.SubElement(clr, qn("a:alpha"), val=str(OPACITY[name[-1]] * 1000))
    elif name.endswith("-"):
        color_format.brightness = -0.3


def theme_colors(prs):
    """The theme's colors as hex, so that matplotlib figures can match any template."""
    root = etree.fromstring(prs.slide_master.part.part_related_by(RT.THEME).blob)
    scheme = root.find(f".//{{{A}}}clrScheme")
    slots = {"ink": "dk1", "paper": "lt1", "a1": "accent1", "a2": "accent2", "a3": "accent3",
             "a4": "accent4", "a5": "accent5", "a6": "accent6"}
    out = {}
    for name, slot in slots.items():
        el = scheme.find(f"{{{A}}}{slot}")[0]
        out[name] = "#" + (el.get("val") if el.tag.endswith("srgbClr") else el.get("lastClr"))
    return out


def open_deck(path):
    """Open a .pptx or .potx. python-pptx refuses .potx files, so relabel them in memory."""
    path = str(path)
    if not path.lower().endswith(".potx"):
        return Presentation(path)
    buf = io.BytesIO()
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(b"presentationml.template.main+xml", b"presentationml.presentation.main+xml")
            dst.writestr(item, data)
    buf.seek(0)
    return Presentation(buf)


def _run(p, s, st):
    r = p.add_run()
    r.text = s
    if st.get("size"):
        r.font.size = Pt(st["size"])
    if st.get("b"):
        r.font.bold = True
    if st.get("i"):
        r.font.italic = True
    if st.get("font"):
        r.font.name = st["font"]
    if st.get("color"):
        paint(r.font.color, st["color"])
    if st.get("baseline"):
        r.font._element.set("baseline", st["baseline"])


def runs(p, s, **st):
    """Append markup to a paragraph as real, separately styled text runs."""
    pos = 0
    for m in TOKEN.finditer(s):
        if m.start() > pos:
            _run(p, s[pos:m.start()], st)
        if m.group("b") is not None:
            runs(p, m.group("b"), **{**st, "b": True})
        elif m.group("i") is not None:
            runs(p, m.group("i"), **{**st, "i": True})
        elif m.group("sub") is not None:
            runs(p, m.group("sub"), **{**st, "baseline": "-25000"})
        elif m.group("sup") is not None:
            runs(p, m.group("sup"), **{**st, "baseline": "30000"})
        else:
            runs(p, m.group("ct"), **{**st, "color": m.group("c")})
        pos = m.end()
    if pos < len(s):
        _run(p, s[pos:], st)


def _clone_chart(part):
    """PowerPoint refuses a file where two slides share one chart, so each copy needs its own."""
    from pptx.parts.chart import ChartPart
    pkg = part.package
    clone = ChartPart.load(pkg.next_partname(ChartPart.partname_template), part.content_type, pkg, part.blob)
    external = clone._element.find(qn("c:externalData"))
    if external is not None:
        clone._element.remove(external)
    workbook = part.chart_workbook.xlsx_part
    if workbook is not None:
        clone.chart_workbook.update_from_xlsx_blob(workbook.blob)
    return clone


def duplicate_slide(prs, slide):
    """Copy a slide to the end of the deck, for step-by-step builds. Charts get their own copy."""
    new = prs.slides.add_slide(slide.slide_layout)
    tree = new.shapes._spTree
    for el in list(tree)[2:]:
        tree.remove(el)
    rids = {}
    for rel in slide.part.rels.values():
        if rel.reltype.endswith(("/slideLayout", "/notesSlide")):
            continue
        if rel.is_external:
            rids[rel.rId] = new.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        elif rel.reltype.endswith("/chart"):
            rids[rel.rId] = new.part.rels.get_or_add(rel.reltype, _clone_chart(rel.target_part))
        else:
            rids[rel.rId] = new.part.rels.get_or_add(rel.reltype, rel.target_part)
    r_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    for el in list(slide.shapes._spTree)[2:]:
        copied = copy.deepcopy(el)
        for node in copied.iter():
            for key, value in node.attrib.items():
                if key.startswith(r_ns) and value in rids:
                    node.set(key, rids[value])
        tree.append(copied)
    return new


def delete_slide(prs, slide):
    ids = prs.slides._sldIdLst
    for sld_id in list(ids):
        if prs.part.related_part(sld_id.rId) is slide.part:
            prs.part.drop_rel(sld_id.rId)
            ids.remove(sld_id)


def move_slide(prs, slide, index):
    """Move a slide to a 0-based position."""
    ids = prs.slides._sldIdLst
    el = next(i for i in ids if prs.part.related_part(i.rId) is slide.part)
    ids.remove(el)
    ids.insert(index, el)


def add_slide(prs, title=None, layout="Title and Content", footer=None):
    """Add a slide from a named layout and carry over its footer and slide number.

    python-pptx leaves footers and slide numbers off new slides, which breaks most templates.
    """
    lay = next((l for l in prs.slide_layouts if l.name == layout), None)
    if lay is None:
        raise ValueError(f"No layout {layout!r}. This template has: {[l.name for l in prs.slide_layouts]}")
    s = prs.slides.add_slide(lay)
    if title and s.shapes.title is not None:
        runs(s.shapes.title.text_frame.paragraphs[0], title)
    for ph in lay.placeholders:
        t = ph.placeholder_format.type
        if t == PH.SLIDE_NUMBER or (t == PH.FOOTER and footer):
            el = copy.deepcopy(ph._element)
            el.find(qn("p:nvSpPr")).find(qn("p:cNvPr")).set("id", str(s.shapes._next_shape_id))
            s.shapes._spTree.append(el)
            if t == PH.FOOTER:
                body = el.find(qn("p:txBody"))
                for para in body.findall(qn("a:p")):
                    body.remove(para)
                para = etree.SubElement(body, qn("a:p"))
                etree.SubElement(etree.SubElement(para, qn("a:r")), qn("a:t")).text = footer
    return s


def text(parent, markup, x, y, w, h, size=18, color="ink", align="l", anchor="t", font=None, raw=False, name=None):
    """A text box with no padding, so it lines up exactly with shapes. parent is a slide or a group.

    raw=True skips the markup, for text full of asterisks or underscores.
    """
    box_ = parent.shapes.add_textbox(I(x), I(y), I(w), I(h))
    if name:
        box_.name = name
    tf = box_.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = ANCHOR[anchor]
    for i, line in enumerate(markup.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN[align]
        if raw:
            _run(p, line, {"size": size, "color": color, "font": font})
        else:
            runs(p, line, size=size, color=color, font=font)
    return box_


def code(parent, source, x, y, w, h, size=18, name="Code"):
    """Plain monospaced code, exactly as written: no box, no syntax colors."""
    return text(parent, source, x, y, w, h, size=size, font="Consolas", raw=True, name=name)


def bullets(slide, items, x=None, y=None, w=None, h=None, anchor=None):
    """Fill the layout's body placeholder, so real bullets and levels come from the template.

    Each item is (level, markup) or (level, markup, points_of_space_before).
    """
    body = next(p for p in slide.placeholders if p.placeholder_format.type in (PH.OBJECT, PH.BODY))
    if x is not None:
        body.left, body.top, body.width, body.height = I(x), I(y), I(w), I(h)
    tf = body.text_frame
    if anchor:
        tf.vertical_anchor = ANCHOR[anchor]
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.level = item[0]
        runs(p, item[1])
        if len(item) > 2:
            p.space_before = Pt(item[2])
    return body


def note(slide, markup, size=20, gap=0.12, lines=1):
    """A small line just above the footer, for a caveat or citation. Same place on every slide.

    Write caveats in *italics* and citations upright. Raise gap if the template has a band there.
    """
    prs = slide.part.package.presentation_part.presentation
    foot = [p.top for p in slide.slide_layout.placeholders if p.placeholder_format.type == PH.FOOTER and p.top > prs.slide_height * 0.6]
    bottom = (min(foot) if foot else prs.slide_height - I(0.45)) / 914400 - gap
    body = next((p for p in slide.slide_layout.slide_master.placeholders if p.placeholder_format.type == PH.BODY), None)
    x, w = (body.left / 914400, body.width / 914400) if body is not None else (0.75, prs.slide_width / 914400 - 1.5)
    h = lines * size * 1.25 / 72
    return text(slide, markup, x, bottom - h, w, h, size=size, anchor="b", name="Note")


def _shape_text(shape, markup, size, color="ink"):
    tf = shape.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    runs(p, markup, size=size, color=color)


def _dress(s, fill, line, width, name):
    s.shadow.inherit = False
    if fill:
        s.fill.solid()
        paint(s.fill.fore_color, fill)
    else:
        s.fill.background()
    if line:
        paint(s.line.color, line)
        s.line.width = Pt(width)
    else:
        s.line.fill.background()
    if name:
        s.name = name


def box(parent, x, y, w, h, color="ink", label=None, size=14, radius=0.05, line=1.25, name=None):
    """A rounded box: pale fill, darker outline in the same color."""
    s = parent.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, I(x), I(y), I(w), I(h))
    s.adjustments[0] = radius / min(w, h)
    _dress(s, color + "+", color, line, name)
    if label:
        _shape_text(s, label, size)
    return s


def rect(parent, x, y, w, h, fill="ink+", line=None, width=1.0, label=None, size=14, shape=MSO_SHAPE.RECTANGLE, name=None):
    """A plain rectangle or other preset shape, for marks whose size carries data."""
    s = parent.shapes.add_shape(shape, I(x), I(y), I(w), I(h))
    _dress(s, fill, line, width, name)
    if label:
        _shape_text(s, label, size)
    return s


def outline(parent, x, y, w, h, color="ink", width=2.25, radius=0.05, name=None):
    """An unfilled rounded box, for marking a region of a figure or chart."""
    s = parent.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, I(x), I(y), I(w), I(h))
    s.adjustments[0] = radius / min(w, h)
    _dress(s, None, color, width, name)
    return s


def circle(parent, cx, cy, r, fill="paper", line="ink", width=1.0, label=None, size=14,
           shape=MSO_SHAPE.OVAL, name=None):
    """A circle centered on (cx, cy). shape=MSO_SHAPE.FLOWCHART_OR draws a circled plus."""
    s = parent.shapes.add_shape(shape, I(cx - r), I(cy - r), I(2 * r), I(2 * r))
    _dress(s, fill, line, width, name)
    if label:
        _shape_text(s, label, size)
    return s


def _head(line, end="tailEnd"):
    etree.SubElement(line._get_or_add_ln(), qn(f"a:{end}"), type="triangle", w="med", len="med")


def _site(shape, side):
    """Where a connector attaches to a shape, and that point's index. Round shapes have eight sites."""
    round_ = shape.auto_shape_type in (MSO_SHAPE.OVAL, MSO_SHAPE.FLOWCHART_OR,
                                       MSO_SHAPE.FLOWCHART_SUMMING_JUNCTION)
    x, y, w, h = shape.left, shape.top, shape.width, shape.height
    return {"top": (x + w // 2, y, 0), "bottom": (x + w // 2, y + h, 4 if round_ else 2),
            "left": (x, y + h // 2, 2 if round_ else 1), "right": (x + w, y + h // 2, 6 if round_ else 3)}[side]


def arrow(parent, a, b, color="ink", width=1.5, head=True, name=None):
    """A straight connector. a and b are (x, y), or (shape, "top"|"bottom"|"left"|"right").

    Ends given as a shape stay attached when someone moves that shape in PowerPoint.
    """
    c = parent.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 0, 0)
    if isinstance(a[0], (int, float)):
        c.begin_x, c.begin_y = I(a[0]), I(a[1])
    else:
        x, y, idx = _site(*a)
        c._connect_begin_to(a[0], idx)
        c.begin_x, c.begin_y = x, y
    if isinstance(b[0], (int, float)):
        c.end_x, c.end_y = I(b[0]), I(b[1])
    else:
        x, y, idx = _site(*b)
        c._connect_end_to(b[0], idx)
        c.end_x, c.end_y = x, y
    paint(c.line.color, color)
    c.line.width = Pt(width)
    if head:
        _head(c.line)
    if name:
        c.name = name
    return c


def path(parent, steps, color="ink", width=1.5, fill=None, head=False, name=None):
    """A custom line, curve or filled shape, adjustable with Edit Points in PowerPoint.

    Steps are ("M", x, y), ("L", x, y), ("C", x1, y1, x2, y2, x, y) and optionally ("Z",).
    """
    pts = [st[1:] for st in steps if st[0] != "Z"]
    xs = [v for p in pts for v in p[0::2]]
    ys = [v for p in pts for v in p[1::2]]
    x0, y0 = min(xs), min(ys)
    w, h = max(xs) - x0, max(ys) - y0
    s = parent.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x0), I(y0), I(w), I(h))
    pt = lambda x, y: f'<a:pt x="{I(x - x0)}" y="{I(y - y0)}"/>'
    parts = []
    for st in steps:
        if st[0] == "M":
            parts.append(f"<a:moveTo>{pt(st[1], st[2])}</a:moveTo>")
        elif st[0] == "L":
            parts.append(f"<a:lnTo>{pt(st[1], st[2])}</a:lnTo>")
        elif st[0] == "C":
            parts.append(f"<a:cubicBezTo>{pt(st[1], st[2])}{pt(st[3], st[4])}{pt(st[5], st[6])}</a:cubicBezTo>")
        else:
            parts.append("<a:close/>")
    no_fill = "" if fill else ' fill="none"'
    geom = etree.fromstring(
        f'<a:custGeom xmlns:a="{A}"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        f'<a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="{I(w)}" h="{I(h)}"{no_fill}>'
        f'{"".join(parts)}</a:path></a:pathLst></a:custGeom>'
    )
    sp_pr = s._element.spPr
    sp_pr.replace(sp_pr.find(qn("a:prstGeom")), geom)
    _dress(s, fill, color, width, name)
    if head:
        _head(s.line)
    return s


@functools.lru_cache(maxsize=None)
def _omml(latex):
    """LaTeX to Office math, using pandoc's converter. Pandoc ships as a wheel on every major OS."""
    with tempfile.TemporaryDirectory() as d:
        src, out = Path(d) / "m.tex", Path(d) / "m.docx"
        src.write_text("\\[" + latex + "\\]", encoding="utf-8")
        subprocess.run([pypandoc.get_pandoc_path(), "-f", "latex", "-t", "docx", str(src), "-o", str(out)],
                       check=True, capture_output=True)
        with zipfile.ZipFile(out) as z:
            doc = etree.fromstring(z.read("word/document.xml"))
    math = doc.find(f".//{{{M}}}oMath")
    if math is None:
        raise ValueError(f"pandoc could not parse this LaTeX: {latex}")
    return etree.tostring(math)


def _stretch_braces(math):
    """Pandoc writes braces as a limit, which does not stretch. Make them real group characters."""
    for tag, pos, just, chars in ((f"{{{M}}}limLow", "bot", "top", "⏟"), (f"{{{M}}}limUpp", "top", "bot", "⏞")):
        for lim_el in list(math.iter(tag)):
            lim, base = lim_el.find(f"{{{M}}}lim"), lim_el.find(f"{{{M}}}e")
            if "".join(lim.itertext()).strip() != chars:
                continue
            group = etree.Element(f"{{{M}}}groupChr")
            pr = etree.SubElement(group, f"{{{M}}}groupChrPr")
            etree.SubElement(pr, f"{{{M}}}chr").set(f"{{{M}}}val", chars)
            etree.SubElement(pr, f"{{{M}}}pos").set(f"{{{M}}}val", pos)
            etree.SubElement(pr, f"{{{M}}}vertJc").set(f"{{{M}}}val", just)
            group.append(base)
            lim_el.getparent().replace(lim_el, group)


def eq(slide, latex, x, y, w, h, size=26, align="c", anchor="m", display=True, fallback=None, name="Equation"):
    """LaTeX to a native PowerPoint equation you can click into and edit. Returns nothing usable.

    display=False sets it inline, with a compact sum and fraction. fallback is what apps without
    native equations show; it defaults to the LaTeX source.
    """
    math = etree.fromstring(_omml(latex))
    _stretch_braces(math)
    for bad in list(math.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")):
        bad.getparent().remove(bad)
    for r in math.iter(f"{{{M}}}r"):
        pr = r.find(f"{{{M}}}rPr")
        is_text = pr is not None and pr.find(f"{{{M}}}nor") is not None
        # PowerPoint shrinks math in limits and scripts by itself, but not plain text, so size that here.
        small = any(a.tag in (f"{{{M}}}lim", f"{{{M}}}sub", f"{{{M}}}sup") for a in r.iterancestors())
        rpr = etree.Element(f"{{{A}}}rPr", lang="en-US", sz=str(int(size * (72 if is_text and small else 100))))
        etree.SubElement(rpr, f"{{{A}}}latin", typeface="+mn-lt" if is_text else "Cambria Math")
        r.find(f"{{{M}}}t").addprevious(rpr)
    box_ = text(slide, fallback or latex, x, y, w, h, size=size, align=align, anchor=anchor, name=name)
    sp = box_._element
    plain = copy.deepcopy(sp)
    para = etree.fromstring(f'<a:p xmlns:a="{A}" xmlns:a14="{A14}" xmlns:m="{M}"><a:pPr algn="'
                            f'{"ctr" if align == "c" else align}"/></a:p>')
    if display:
        holder = etree.SubElement(etree.SubElement(para, f"{{{A14}}}m"), f"{{{M}}}oMathPara")
        jc = etree.SubElement(etree.SubElement(holder, f"{{{M}}}oMathParaPr"), f"{{{M}}}jc")
        jc.set(f"{{{M}}}val", {"c": "center", "l": "left", "r": "right"}[align])
    else:
        # PowerPoint renders an equation alone in a paragraph as display math; a leading space keeps it inline.
        etree.SubElement(etree.SubElement(para, f"{{{A}}}r"), f"{{{A}}}t").text = " "
        holder = etree.SubElement(para, f"{{{A14}}}m")
    holder.append(math)
    body = sp.find(qn("p:txBody"))
    body.remove(body.find(qn("a:p")))
    body.append(para)
    wrap = etree.Element(f"{{{MC}}}AlternateContent", nsmap={"mc": MC})
    sp.addprevious(wrap)
    choice = etree.SubElement(wrap, f"{{{MC}}}Choice", Requires="a14", nsmap={"a14": A14})
    choice.append(sp)
    etree.SubElement(wrap, f"{{{MC}}}Fallback").append(plain)


def equations(slide):
    """Equations sit in a wrapper that slide.shapes cannot see. Returns [(wrapper, x, y, w, h)]."""
    found = []
    for wrap in slide._element.iter(f"{{{MC}}}AlternateContent"):
        choice = wrap.find(f"{{{MC}}}Choice")
        sp = choice.find(qn("p:sp")) if choice is not None else None
        if sp is None or not list(sp.iter(f"{{{A14}}}m")):
            continue
        xfrm = sp.find(qn("p:spPr")).find(qn("a:xfrm"))
        off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
        found.append((wrap, *(int(v) / 914400 for v in (off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy")))))
    return found


def replace_eq(slide, index, latex, **kwargs):
    """Swap one equation for new LaTeX, in the same spot."""
    wrap, x, y, w, h = equations(slide)[index]
    wrap.getparent().remove(wrap)
    eq(slide, latex, x, y, w, h, **kwargs)


def table(slide, rows, x, y, widths, row_h=0.42, size=20, align=None, name="Table"):
    """A native table with only horizontal rules: above and below the header, and at the bottom.

    The first row is the header. widths are column widths; align is a string like "lrr".
    """
    frame = slide.shapes.add_table(len(rows), len(rows[0]), I(x), I(y), I(sum(widths)), I(row_h * len(rows)))
    frame.name = name
    tbl = frame.table
    pr = tbl._tbl.tblPr
    for attr in ("firstRow", "bandRow"):
        pr.set(attr, "0")
    style = pr.find(qn("a:tableStyleId"))
    if style is not None:
        style.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # PowerPoint's "No Style, No Grid"
    for j, w in enumerate(widths):
        tbl.columns[j].width = I(w)
    for i, row in enumerate(rows):
        tbl.rows[i].height = I(row_h)
        for j, value in enumerate(row):
            cell = tbl.cell(i, j)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = I(0.08)
            cell.margin_top = cell.margin_bottom = I(0.02)
            p = cell.text_frame.paragraphs[0]
            p.alignment = ALIGN[(align or "l" * len(row))[j]]
            runs(p, str(value), size=size, color="ink", b=i == 0)
            tc_pr = cell._tc.get_or_add_tcPr()
            edges = [("a:lnT", i == 0, 1.5), ("a:lnB", i == 0 or i == len(rows) - 1, 1.5 if i else 0.75)]
            for k, (tag, on, weight) in enumerate(edges):
                ln = etree.Element(qn(tag), w=str(int(Pt(weight))) if on else "0")
                if on:
                    fill = etree.SubElement(ln, qn("a:solidFill"))
                    etree.SubElement(fill, qn("a:schemeClr"), val="tx1")
                else:
                    etree.SubElement(ln, qn("a:noFill"))
                tc_pr.insert(k, ln)
            etree.SubElement(tc_pr, qn("a:noFill"))
    return frame


def style_chart(chart, font=16, legend=False, title=None):
    """Remove the default chart look: quiet axes, faint gridlines, theme colors, no legend."""
    chart.has_legend = legend
    chart.font.size = Pt(font)
    paint(chart.font.color, "ink")
    if title:
        chart.has_title = True
        chart.chart_title.include_in_layout = False
        p = chart.chart_title.text_frame.paragraphs[0]
        runs(p, title, size=font, color="ink")
        for r in p.runs:
            r.font.bold = False
    else:
        chart.has_title = False
    for axis in (chart.value_axis, chart.category_axis):
        paint(axis.format.line.color, "ink")
        axis.major_tick_mark = XL_TICK_MARK.OUTSIDE
        axis.minor_tick_mark = XL_TICK_MARK.NONE
        axis.has_minor_gridlines = False
    chart.value_axis.has_major_gridlines = True
    paint(chart.value_axis.major_gridlines.format.line.color, "ink+")
    chart.value_axis.major_gridlines.format.line.width = Pt(0.75)
    chart.category_axis.has_major_gridlines = False
    # PowerPoint hides some category labels to save room unless told otherwise.
    no_multi = chart.category_axis._element.find(qn("c:noMultiLvlLbl"))
    if no_multi is not None:
        no_multi.addprevious(etree.Element(f"{{{C}}}tickLblSkip", val="1"))


def axis_title(axis, markup, size=16):
    axis.has_title = True
    p = axis.axis_title.text_frame.paragraphs[0]
    runs(p, markup, size=size, color="ink")
    for r in p.runs:
        r.font.bold = False


def fix_plot(frame, x, y, w, h):
    """Pin a chart's inner plot area to known slide coordinates, so overlays can be placed from data.

    Leave at least 0.2 inches between the plot and the chart's frame, or PowerPoint shifts the plot.
    """
    fx, fy, fw, fh = (v / 914400 for v in (frame.left, frame.top, frame.width, frame.height))
    area = frame.chart._chartSpace.chart.plotArea
    old = area.find(qn("c:layout"))
    if old is not None:
        area.remove(old)
    area.insert(0, etree.fromstring(
        f'<c:layout xmlns:c="{C}"><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/>'
        f'<c:yMode val="edge"/><c:x val="{(x - fx) / fw:.5f}"/><c:y val="{(y - fy) / fh:.5f}"/>'
        f'<c:w val="{w / fw:.5f}"/><c:h val="{h / fh:.5f}"/></c:manualLayout></c:layout>'
    ))
