#!/usr/bin/env python3
"""Convert notes_final.html into an editable DOCX with the same content."""
import re
from html.parser import HTMLParser

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

NAVY = RGBColor(0x12, 0x3A, 0x63)
TEAL = RGBColor(0x0E, 0x6F, 0x66)
GREY = RGBColor(0x5B, 0x6B, 0x7B)
RED = RGBColor(0xA4, 0x40, 0x2F)
BODY = RGBColor(0x1B, 0x22, 0x2B)

VOID = {"br", "img", "hr", "meta", "link", "input"}


class Node:
    def __init__(self, tag, attrs=None, parent=None):
        self.tag = tag
        self.attrs = dict(attrs or {})
        self.children = []
        self.parent = parent

    @property
    def cls(self):
        return self.attrs.get("class", "")

    def text(self):
        out = []
        for c in self.children:
            out.append(c if isinstance(c, str) else c.text())
        return "".join(out)

    def find_all(self, tag):
        res = []
        for c in self.children:
            if isinstance(c, Node):
                if c.tag == tag:
                    res.append(c)
                res.extend(c.find_all(tag))
        return res

    def iter(self):
        for c in self.children:
            if isinstance(c, Node):
                yield c
                yield from c.iter()


class Builder(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node("root")
        self.stack = [self.root]

    def handle_starttag(self, tag, attrs):
        node = Node(tag, attrs, self.stack[-1])
        self.stack[-1].children.append(node)
        if tag not in VOID:
            self.stack.append(node)

    def handle_startendtag(self, tag, attrs):
        self.stack[-1].children.append(Node(tag, attrs, self.stack[-1]))

    def handle_endtag(self, tag):
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                del self.stack[i:]
                return

    def handle_data(self, data):
        if data.strip() or " " in data:
            self.stack[-1].children.append(data)


def shade(el, hexcolor):
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), hexcolor)
    el.append(sh)


def para_shading(p, hexcolor):
    shade(p._p.get_or_add_pPr(), hexcolor)


def para_border(p, side="left", color="123A63", size=18):
    pPr = p._p.get_or_add_pPr()
    bdr = pPr.find(qn("w:pBdr"))
    if bdr is None:
        bdr = OxmlElement("w:pBdr")
        pPr.append(bdr)
    el = OxmlElement(f"w:{side}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(size))
    el.set(qn("w:space"), "4")
    el.set(qn("w:color"), color)
    bdr.append(el)


def table_header_repeat(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


class Doc:
    def __init__(self):
        self.d = Document()
        sec = self.d.sections[0]
        sec.page_width = Cm(21)
        sec.page_height = Cm(29.7)
        sec.left_margin = sec.right_margin = Cm(1.7)
        sec.top_margin = Cm(1.6)
        sec.bottom_margin = Cm(1.6)

        st = self.d.styles["Normal"]
        st.font.name = "Calibri"
        st.font.size = Pt(9.5)
        st.font.color.rgb = BODY
        st.paragraph_format.space_after = Pt(3)
        st.paragraph_format.line_spacing = 1.05
        rpr = st.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.append(rf)
        rf.set(qn("w:ascii"), "Calibri")
        rf.set(qn("w:hAnsi"), "Calibri")

        for name, size, color, before, after in [
            ("Heading 1", 15, NAVY, 14, 5),
            ("Heading 2", 11.5, NAVY, 9, 3),
            ("Heading 3", 10, TEAL, 7, 2),
        ]:
            h = self.d.styles[name]
            h.font.name = "Calibri"
            h.font.size = Pt(size)
            h.font.bold = True
            h.font.color.rgb = color
            h.paragraph_format.space_before = Pt(before)
            h.paragraph_format.space_after = Pt(after)
            h.paragraph_format.keep_with_next = True

        self.fig_no = 0
        self.footer()

    def footer(self):
        sec = self.d.sections[0]
        p = sec.footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("DBMS Unit III  ·  One-Day Exam Notes  ·  page ")
        r.font.size = Pt(7.5)
        r.font.color.rgb = GREY
        fld = OxmlElement("w:fldSimple")
        fld.set(qn("w:instr"), "PAGE")
        p._p.append(fld)

    # ---------- inline ----------
    def runs(self, p, node, fmt=None):
        fmt = dict(fmt or {})
        for c in node.children:
            if isinstance(c, str):
                txt = re.sub(r"\s+", " ", c)
                if not txt.strip() and p.runs and p.runs[-1].text.endswith(" "):
                    continue
                self.add_run(p, txt, fmt)
                continue
            t = c.tag
            if t == "br":
                self.add_run(p, "\n", fmt)
            elif t in ("b", "strong"):
                self.runs(p, c, {**fmt, "bold": True})
            elif t in ("i", "em", "cite"):
                self.runs(p, c, {**fmt, "italic": True})
            elif t == "sub":
                self.runs(p, c, {**fmt, "sub": True})
            elif t == "sup":
                self.runs(p, c, {**fmt, "sup": True})
            elif t in ("code", "kbd"):
                self.runs(p, c, {**fmt, "mono": True})
            elif t == "span" and ("mono" in c.cls or "code" in c.cls):
                self.runs(p, c, {**fmt, "mono": True})
            elif t == "span" and ("tag" in c.cls):
                self.runs(p, c, {**fmt, "bold": True, "small": True, "caps": True,
                                 "color": TEAL})
            elif t == "small" or (t == "span" and "small" in c.cls):
                self.runs(p, c, {**fmt, "small": True})
            elif t in ("svg", "script", "style"):
                continue
            else:
                self.runs(p, c, fmt)

    def add_run(self, p, txt, fmt):
        r = p.add_run(txt)
        if fmt.get("bold"):
            r.bold = True
        if fmt.get("italic"):
            r.italic = True
        if fmt.get("sub"):
            r.font.subscript = True
        if fmt.get("sup"):
            r.font.superscript = True
        if fmt.get("mono"):
            r.font.name = "Consolas"
            r.font.size = Pt(8)
        if fmt.get("small"):
            r.font.size = Pt(8)
        if fmt.get("color"):
            r.font.color.rgb = fmt["color"]
        if fmt.get("caps"):
            r.font.all_caps = True
            r.font.size = Pt(7)

    # ---------- blocks ----------
    def para(self, node, style=None, shade_hex=None, border=None):
        p = self.d.add_paragraph(style=style)
        self.runs(p, node)
        if shade_hex:
            para_shading(p, shade_hex)
        if border:
            para_border(p, color=border)
        return p

    def heading(self, node, style):
        p = self.d.add_paragraph(style=style)
        main = []
        src = None
        tag = None
        for ch in node.children:
            if isinstance(ch, str):
                main.append(ch)
            elif ch.tag == "span" and "sub" in ch.cls:
                src = ch.text().strip()
            elif ch.tag == "span" and "tag" in ch.cls:
                tag = ch.text().strip()
            elif ch.tag == "span" and "mk" in ch.cls:
                continue
            else:
                main.append(ch.text())
        txt = re.sub(r"\s+", " ", "".join(main)).strip()
        r = p.add_run(txt)
        if src:
            r2 = p.add_run("   " + src)
            r2.bold = False
            r2.italic = True
            r2.font.size = Pt(8)
            r2.font.color.rgb = GREY
        if tag:
            r3 = p.add_run("   " + tag)
            r3.bold = True
            r3.font.all_caps = True
            r3.font.size = Pt(7)
            r3.font.color.rgb = TEAL
        if style == "Heading 1":
            para_border(p, side="bottom", color="123A63", size=8)
        return p

    def pre(self, node):
        p = self.d.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.2)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        raw = node.text()
        lines = [ln.rstrip() for ln in raw.strip("\n").split("\n")]
        for i, ln in enumerate(lines):
            r = p.add_run(ln)
            r.font.name = "Consolas"
            r.font.size = Pt(8)
            r.font.color.rgb = RGBColor(0x1B, 0x22, 0x2B)
            if i < len(lines) - 1:
                r.add_break()
        para_shading(p, "F4F6F9")
        para_border(p, side="left", color="9FB3C8", size=12)

    def figure(self, node):
        self.fig_no += 1
        p = self.d.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        p.add_run().add_picture(f"figs/fig{self.fig_no}.png", width=Cm(13.4))
        cap = node.find_all("figcaption")
        if cap:
            c = self.d.add_paragraph()
            c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            c.paragraph_format.space_after = Pt(6)
            r = c.add_run(cap[0].text().strip())
            r.italic = True
            r.font.size = Pt(8)
            r.font.color.rgb = GREY

    def list_block(self, node, level=0):
        ordered = node.tag == "ol"
        n = 0
        for item in [c for c in node.children if isinstance(c, Node) and c.tag == "li"]:
            n += 1
            p = self.d.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5 + 0.45 * level)
            p.paragraph_format.first_line_indent = Cm(-0.4)
            p.paragraph_format.space_after = Pt(2)
            bullet = f"{n}. " if ordered else "\u2022  "
            if "check" in node.cls:
                bullet = "\u2610  "
            r = p.add_run(bullet)
            r.font.color.rgb = TEAL if not ordered else NAVY
            r.bold = ordered
            sub = [c for c in item.children if isinstance(c, Node) and c.tag in ("ul", "ol")]
            flat = Node("li")
            flat.children = [c for c in item.children
                             if not (isinstance(c, Node) and c.tag in ("ul", "ol"))]
            self.runs(p, flat)

    def table(self, node):
        rows = []
        for tr in node.find_all("tr"):
            if tr.parent is not None and tr.parent.tag != "tbody" and tr.parent.tag != "table" \
                    and tr.parent.tag != "thead":
                continue
            cells = []
            for cell in [c for c in tr.children if isinstance(c, Node) and c.tag in ("td", "th")]:
                cells.append((cell.tag, cell))
            rows.append(cells)
        rows = [r for r in rows if r]
        if not rows:
            return
        ncols = max(len(r) for r in rows)
        t = self.d.add_table(rows=len(rows), cols=ncols)
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = True
        for ri, row in enumerate(rows):
            for ci in range(ncols):
                cell = t.cell(ri, ci)
                cell.paragraphs[0].paragraph_format.space_after = Pt(1)
                cell.paragraphs[0].paragraph_format.space_before = Pt(1)
                if ci >= len(row):
                    continue
                tag, src = row[ci]
                self.runs(cell.paragraphs[0], src, {"small": True})
                for r in cell.paragraphs[0].runs:
                    r.font.size = Pt(8)
                    if tag == "th":
                        r.bold = True
                        r.font.color.rgb = NAVY
                if ri == 0:
                    shade(cell._tc.get_or_add_tcPr(), "E8EEF6")
        if rows and rows[0] and rows[0][0][0] == "th":
            table_header_repeat(t.rows[0])
        sp = self.d.add_paragraph()
        sp.paragraph_format.space_after = Pt(2)
        sp.paragraph_format.space_before = Pt(0)

    def block(self, c):
        t = c.tag
        if t in ("h1", "h2"):
            self.heading(c, "Heading 1")
        elif t == "h3":
            self.heading(c, "Heading 2")
        elif t in ("h4", "h5", "h6"):
            self.heading(c, "Heading 3")
        elif t == "p":
            cls = c.cls
            if "subtitle" in cls:
                p = self.d.add_paragraph()
                self.runs(p, c)
                for r in p.runs:
                    r.font.size = Pt(10)
                    r.italic = True
                    r.font.color.rgb = GREY
            elif "small" in cls or "fine" in cls or "note" in cls or "legend" in cls:
                p = self.para(c)
                for r in p.runs:
                    r.font.size = Pt(8)
            else:
                self.para(c)
        elif t == "pre":
            self.pre(c)
        elif t in ("ul", "ol"):
            self.list_block(c)
        elif t == "table":
            self.table(c)
        elif t == "figure":
            self.figure(c)
        elif t == "section":
            self.d.add_page_break()
            self.container(c)
        elif t in ("div", "header", "main", "body", "root", "span", "li", "td",
                   "figcaption", "article", "aside", "footer"):
            if c.text().strip() or c.find_all("table") or c.find_all("figure") \
                    or c.find_all("svg"):
                self.container(c)
        elif t in ("svg", "script", "style"):
            return
        else:
            self.container(c)

    def container(self, node):
        for c in node.children:
            if isinstance(c, str):
                if c.strip():
                    self.d.add_paragraph(re.sub(r"\s+", " ", c).strip())
                continue
            self.block(c)

    def cover(self, node, toc_pages=None):
        for c in node.children:
            if not isinstance(c, Node):
                continue
            t = c.tag
            cls = c.cls
            if "kicker" in cls:
                p = self.d.add_paragraph()
                r = p.add_run(c.text().strip().upper())
                r.font.size = Pt(8)
                r.bold = True
                r.font.color.rgb = TEAL
                r.font.all_caps = True
            elif t == "h1":
                p = self.d.add_paragraph()
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(2)
                main = []
                sub = []
                for ch in c.children:
                    if isinstance(ch, str):
                        main.append(ch)
                    elif ch.tag == "span":
                        sub.append(ch.text())
                    else:
                        main.append(ch.text())
                r = p.add_run(re.sub(r"\s+", " ", "".join(main)).strip())
                r.font.size = Pt(30)
                r.bold = True
                r.font.color.rgb = NAVY
                if sub:
                    p2 = self.d.add_paragraph()
                    p2.paragraph_format.space_after = Pt(10)
                    r2 = p2.add_run(re.sub(r"\s+", " ", "".join(sub)).strip())
                    r2.font.size = Pt(14)
                    r2.italic = True
                    r2.font.color.rgb = TEAL
            elif "lede" in cls:
                p = self.d.add_paragraph()
                self.runs(p, c)
                for r in p.runs:
                    r.font.size = Pt(10)
                    r.font.color.rgb = GREY
                p.paragraph_format.space_after = Pt(10)
            elif "tocfull" in cls:
                rows = [r for r in c.children if isinstance(r, Node)]
                tb = self.d.add_table(rows=len(rows), cols=3)
                tb.autofit = True
                for i, row in enumerate(rows):
                    spans = row.find_all("span")
                    if len(spans) >= 2:
                        label, page = spans[0].text().strip(), spans[1].text().strip()
                    else:
                        label, page = row.text().strip(), ""
                    a = tb.cell(i, 0).paragraphs[0]
                    self.add_run(a, label, {"bold": False})
                    b = tb.cell(i, 1).paragraphs[0]
                    self.add_run(b, ". . . . . . . . . . . . . . . . . . . . . . . . . . . . . .", {"small": True, "color": GREY})
                    cc = tb.cell(i, 2).paragraphs[0]
                    cc.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    self.add_run(cc, page, {"bold": True, "color": NAVY})
                for r in tb.rows:  # no visible borders
                    for cell in r.cells:
                        cell.paragraphs[0].paragraph_format.space_after = Pt(2)
                self.d.add_paragraph().paragraph_format.space_after = Pt(2)
            elif "unitmap" in cls:
                items = [r for r in c.children if isinstance(r, Node)]
                tb = self.d.add_table(rows=2, cols=len(items))
                tb.style = "Table Grid"
                for i, it in enumerate(items):
                    hs = it.find_all("h4")
                    ps = it.find_all("p")
                    cell = tb.cell(0, i).paragraphs[0]
                    self.add_run(cell, hs[0].text().strip() if hs else "", {"bold": True, "color": NAVY, "small": True})
                    cell2 = tb.cell(1, i).paragraphs[0]
                    self.add_run(cell2, ps[0].text().strip() if ps else "", {"small": True})
                    shade(tb.cell(0, i)._tc.get_or_add_tcPr(), "E8EEF6")
                self.d.add_paragraph().paragraph_format.space_after = Pt(2)
            elif "rules" in cls:
                for i, row in enumerate([r for r in c.children if isinstance(r, Node)], 1):
                    p = self.d.add_paragraph()
                    p.paragraph_format.left_indent = Cm(0.5)
                    p.paragraph_format.first_line_indent = Cm(-0.4)
                    r = p.add_run(f"{i}.  ")
                    r.bold = True
                    r.font.color.rgb = NAVY
                    body = Node("x")
                    body.children = [ch for ch in row.children
                                     if not (isinstance(ch, Node) and ch.tag == "span")]
                    self.runs(p, body)
            elif "band" in cls:
                p = self.d.add_paragraph()
                self.runs(p, c)
                for r in p.runs:
                    r.font.size = Pt(8.5)
                    r.font.color.rgb = GREY
                para_shading(p, "F4F6F9")
                para_border(p, side="left", color="9FB3C8", size=12)
            elif t == "div" and "rule" == cls:
                p = self.d.add_paragraph()
                para_border(p, side="bottom", color="123A63", size=12)
                p.paragraph_format.space_after = Pt(6)
            else:
                self.block(c)


def main():
    html = open("notes_final.html").read()
    body_match = re.search(r"<body[^>]*>(.*)</body>", html, re.S)
    src = body_match.group(1) if body_match else html
    src = re.sub(r"<span class=\"mk\">.*?</span>", "", src, flags=re.S)
    b = Builder()
    b.feed(src)

    doc = Doc()
    root = b.root
    body = None
    for n in root.iter():
        if n.tag == "body":
            body = n
            break
    body = body or root

    # cover section (first section element) rendered first
    sections = [n for n in body.iter() if n.tag == "section"]
    cover_children = [c for c in body.children if isinstance(c, Node) and c.tag == "section"]
    first = cover_children[0] if cover_children else None
    if first is not None:
        doc.d.add_page_break()
        doc.cover(first)

    for s in cover_children[1:]:
        doc.container(s)

    doc.d.save("/workspace/project/DBMS_One_Day_Exam_Notes.docx")
    print("saved docx; figures:", doc.fig_no)


if __name__ == "__main__":
    main()
