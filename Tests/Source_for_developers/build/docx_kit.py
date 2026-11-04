"""Мини-набор для сборки DOCX-документов ELKA (python-docx)"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

BLUE = RGBColor(0x1F, 0x4E, 0x78)

def _shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)

class Doc:
    def __init__(self, title, subtitle, version="8.2"):
        self.d = Document()
        sec = self.d.sections[0]
        sec.page_width = Cm(21); sec.page_height = Cm(29.7)
        sec.left_margin = sec.right_margin = Cm(2); sec.top_margin = Cm(1.8); sec.bottom_margin = Cm(1.8)
        st = self.d.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(10.5)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
        st.paragraph_format.space_after = Pt(4)
        for lvl, size in ((1, 16), (2, 13), (3, 11.5)):
            hs = self.d.styles[f"Heading {lvl}"]; hs.font.name = "Arial"; hs.font.size = Pt(size); hs.font.color.rgb = BLUE; hs.font.bold = True
            hs.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial"); hs.element.rPr.rFonts.set(qn("w:ascii"), "Arial"); hs.element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
            hs.paragraph_format.space_before = Pt(14 if lvl == 1 else 10); hs.paragraph_format.space_after = Pt(6)
            hs.paragraph_format.keep_with_next = True
        # footer with page numbers
        p = sec.footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"ELKA v{version} — {title}   •   стр. "); r.font.size = Pt(8); r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        self._field(p, "PAGE")
        # title block
        t = self.d.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = t.add_run("ELKA v" + version); r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
        t = self.d.add_paragraph(); r = t.add_run(title); r.font.size = Pt(24); r.font.bold = True; r.font.color.rgb = BLUE
        t = self.d.add_paragraph(); r = t.add_run(subtitle); r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
        self.d.add_paragraph()

    def _field(self, p, code):
        r = p.add_run(); f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin"); r._r.append(f1)
        r2 = p.add_run(); it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = code; r2._r.append(it)
        r3 = p.add_run(); f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end"); r3._r.append(f2)
        for x in (r, r2, r3): x.font.size = Pt(8)

    def h1(self, t): self.d.add_heading(t, 1)
    def h2(self, t): self.d.add_heading(t, 2)
    def h3(self, t): self.d.add_heading(t, 3)
    def p(self, t, bold=False, italic=False, size=None, color=None):
        para = self.d.add_paragraph()
        self._rich(para, t, bold, italic, size, color)
        return para
    def _rich(self, para, t, bold=False, italic=False, size=None, color=None):
        # **жирный** внутри текста
        parts = t.split("**")
        for i, part in enumerate(parts):
            if not part: continue
            r = para.add_run(part); r.bold = bold or (i % 2 == 1); r.italic = italic
            if size: r.font.size = Pt(size)
            if color: r.font.color.rgb = color
    def bullets(self, items):
        for it in items:
            para = self.d.add_paragraph(style="List Bullet"); self._rich(para, it)
    def steps(self, items):
        for i, it in enumerate(items, 1):
            para = self.d.add_paragraph()
            para.paragraph_format.left_indent = Cm(0.9); para.paragraph_format.first_line_indent = Cm(-0.9)
            r = para.add_run(f"{i}.\t"); r.bold = True; r.font.color.rgb = BLUE
            para.paragraph_format.tab_stops.add_tab_stop(Cm(0.9))
            self._rich(para, it)
    def note(self, t, kind="info"):
        colors = {"info": ("DEEAF6", "1F4E78", "Подсказка"), "warn": ("FFF2CC", "7F6000", "Внимание"), "bad": ("F8CBAD", "C00000", "Важно"), "ok": ("E2EFDA", "375623", "Результат")}
        fill, fg, label = colors[kind]
        tb = self.d.add_table(rows=1, cols=1); tb.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = tb.cell(0, 0); _shade(c, fill)
        para = c.paragraphs[0]; r = para.add_run(label + ": "); r.bold = True; r.font.color.rgb = RGBColor.from_string(fg)
        self._rich(para, t)
        self.d.add_paragraph().paragraph_format.space_after = Pt(2)
    def table(self, header, rows, widths=None, size=9):
        tb = self.d.add_table(rows=1, cols=len(header)); tb.style = "Table Grid"; tb.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(header):
            c = tb.rows[0].cells[i]; _shade(c, "1F4E78")
            c.paragraphs[0].text = ""; r = c.paragraphs[0].add_run(str(h)); r.bold = True; r.font.size = Pt(size); r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        # repeat header
        trPr = tb.rows[0]._tr.get_or_add_trPr(); th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trPr.append(th)
        for k, row in enumerate(rows):
            cells = tb.add_row().cells
            for i, v in enumerate(row):
                cells[i].paragraphs[0].text = ""
                self._rich(cells[i].paragraphs[0], "" if v is None else str(v), size=size)
                if k % 2 == 1: _shade(cells[i], "F2F2F2")
        if widths:
            tb.autofit = False
            grid = tb._tbl.tblGrid
            for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
                if i < len(widths): gc.set(qn("w:w"), str(int(widths[i] * 567)))
            tblPr = tb._tbl.tblPr
            lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
            for row in tb.rows:
                for i, w in enumerate(widths): row.cells[i].width = Cm(w)
        self.d.add_paragraph().paragraph_format.space_after = Pt(2)
        return tb
    def img(self, path, width_cm=17, caption=None):
        if not os.path.exists(path): self.p(f"[нет изображения {os.path.basename(path)}]", italic=True); return
        self.d.add_picture(path, width=Cm(width_cm))
        self.d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if caption:
            c = self.d.add_paragraph(); c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = c.add_run(caption); r.italic = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    def code(self, t):
        para = self.d.add_paragraph(); r = para.add_run(t); r.font.name = "Consolas"; r.font.size = Pt(9)
        para.paragraph_format.left_indent = Cm(0.8)
    def pagebreak(self): self.d.add_page_break()
    def save(self, path): self.d.save(path)
