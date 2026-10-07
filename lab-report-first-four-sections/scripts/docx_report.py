# -*- coding: utf-8 -*-
"""Assemble the first four sections of a Chinese lab report into the template.

The template is never modified in place: the script opens a copy of the template
.docx, writes the content of the requested section rows, and saves a new file.
Everything else in the template (cover, later sections, grade table, raw-data
page) is left exactly as it was.

Usage:
    python docx_report.py --template template.docx --spec content.py \
        --assets report_assets --out "outputs/实验报告_前四项_完成版.docx"

The spec is a small Python module so Chinese text and long paragraphs need no
escaping. Recognised module-level names:

    EXPERIMENT_NAME = "光栅光谱仪的使用"   # optional, fills the cover 实验名称 field
    TABLE_INDEX     = 1                    # 0-based index into doc.tables
    SECTION_ROWS    = [0, 1, 2, 3]          # table rows to fill
    SECTIONS        = [SEC1, SEC2, SEC3, SEC4]
    CONFIG          = {"max_w_cm": 11.5, "max_h_cm": 6.9}   # optional overrides

Each section is a list of blocks:

    ("h", "（一）光谱的产生")                      bold sub-heading
    ("p", "正文…")                                  body paragraph, 2-char indent
    ("p", "① …", 0, 200)                            no indent, 200 twips left
    ("eq", omml_children, 4)                        numbered native equation
    ("img", "ppt_p05_img5.png", "图 2-2  …（来源：PPT 第 5 页）")
    ("tbl", ["谱线", "标准波长 (nm)"], [["汞谱线", "404.7、…"]], "表 4-1  …")
"""
import argparse
import importlib.util
import os
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
# omml.py ships next to this file; import it whether the script is executed
# directly or loaded as a module.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

DEFAULTS = {
    "body_pt": 10.5,        # 五号
    "cap_pt": 9.0,          # 图表名小一号
    "cn_font": "宋体",
    "en_font": "Times New Roman",
    "max_w_cm": 11.5,       # keep figures inside the printable column
    "max_h_cm": 6.9,
    "embed_max_px": 1600,   # downscale big images so the .docx stays portable
    "cell_text_twips": 9744,
}


class Report:
    def __init__(self, template, assets_dir, tmp_dir, cfg=None):
        self.doc = Document(template)
        self.doc_part_path = template
        self.assets = assets_dir
        self.tmp = tmp_dir
        self.cfg = dict(DEFAULTS)
        if cfg:
            self.cfg.update(cfg)

    # ---------------------------------------------------------------- runs --
    def set_font(self, run, size=None, bold=False):
        cfg = self.cfg
        run.font.size = Pt(size if size is not None else cfg["body_pt"])
        run.font.bold = bold
        run.font.color.rgb = RGBColor(0, 0, 0)
        rpr = run._element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.insert(0, rf)
        rf.set(qn("w:ascii"), cfg["en_font"])
        rf.set(qn("w:hAnsi"), cfg["en_font"])
        rf.set(qn("w:cs"), cfg["en_font"])
        rf.set(qn("w:eastAsia"), cfg["cn_font"])

    def add_text(self, par, text, size=None, bold=False):
        self.set_font(par.add_run(text), size=size, bold=bold)
        return par

    def set_indent(self, par, first_line_chars=0, left_twips=0):
        """python-docx cannot express firstLineChars, so patch w:ind directly."""
        ppr = par._p.get_or_add_pPr()
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            ppr.append(ind)
        if first_line_chars:
            ind.set(qn("w:firstLineChars"), str(first_line_chars))
            ind.set(qn("w:firstLine"),
                    str(int(first_line_chars * self.cfg["body_pt"] * 20 / 100)))
        else:
            ind.set(qn("w:firstLineChars"), "0")
            ind.set(qn("w:firstLine"), "0")
        if left_twips:
            ind.set(qn("w:left"), str(left_twips))

    # ----------------------------------------------------------- paragraphs --
    def body_par(self, cell, text, indent=200, left=0, bold=False, align=None):
        par = cell.add_paragraph()
        pf = par.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(2)
        pf.line_spacing = 1.28
        if align is not None:
            par.alignment = align
        self.set_indent(par, indent, left)
        self.add_text(par, text, bold=bold)
        return par

    def heading_par(self, cell, text):
        return self.body_par(cell, text, indent=0, bold=True,
                             align=WD_ALIGN_PARAGRAPH.LEFT)

    # ------------------------------------------------------------ equations --
    def equation_par(self, cell, children, number=None):
        """Native Word equation, centred, with a right-aligned number.

        The equation is an inline m:oMath inside an ordinary paragraph that uses
        a centre tab stop and a right tab stop - that is how Word itself lays out
        numbered display equations.
        """
        from omml import omml

        width = self.cfg["cell_text_twips"]
        par = cell.add_paragraph()
        pf = par.paragraph_format
        pf.space_before = Pt(3)
        pf.space_after = Pt(4)
        pf.line_spacing = 1.0
        ppr = par._p.get_or_add_pPr()
        ind = OxmlElement("w:ind")
        ind.set(qn("w:firstLineChars"), "0")
        ind.set(qn("w:firstLine"), "0")
        ppr.append(ind)
        tabs = OxmlElement("w:tabs")
        for val, pos in (("center", width // 2), ("right", width)):
            tb = OxmlElement("w:tab")
            tb.set(qn("w:val"), val)
            tb.set(qn("w:pos"), str(pos))
            tabs.append(tb)
        ppr.append(tabs)
        lead = OxmlElement("w:r")
        lead.append(OxmlElement("w:tab"))
        par._p.append(lead)
        par._p.append(omml(children))
        if number is not None:
            tail = OxmlElement("w:r")
            tail.append(OxmlElement("w:tab"))
            t = OxmlElement("w:t")
            t.text = "（%d）" % number
            tail.append(t)
            rpr = OxmlElement("w:rPr")
            rf = OxmlElement("w:rFonts")
            rf.set(qn("w:ascii"), self.cfg["en_font"])
            rf.set(qn("w:hAnsi"), self.cfg["en_font"])
            rf.set(qn("w:eastAsia"), self.cfg["cn_font"])
            rpr.append(rf)
            sz = OxmlElement("w:sz")
            sz.set(qn("w:val"), str(int(self.cfg["body_pt"] * 2)))
            rpr.append(sz)
            tail.insert(0, rpr)
            par._p.append(tail)
        return par

    # ------------------------------------------------------- atomic blocks --
    @staticmethod
    def _no_borders(tbl, width_cm=None):
        tblpr = tbl._tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement("w:" + edge)
            el.set(qn("w:val"), "none")
            el.set(qn("w:sz"), "0")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
            borders.append(el)
        tblpr.append(borders)
        if width_cm is not None:
            tw = OxmlElement("w:tblW")
            tw.set(qn("w:w"), str(int(width_cm * 567)))
            tw.set(qn("w:type"), "dxa")
            tblpr.append(tw)
        tbli = OxmlElement("w:tblInd")
        tbli.set(qn("w:w"), "0")
        tbli.set(qn("w:type"), "dxa")
        tblpr.append(tbli)

    @staticmethod
    def _no_split(row):
        trpr = row._tr.get_or_add_trPr()
        trpr.append(OxmlElement("w:cantSplit"))

    def wrapper(self, cell, width_cm):
        """A borderless 1x1 table whose row cannot break across pages.

        Needed because keepNext/keepLines are ignored for paragraphs inside a
        table cell: without this, a figure ends up at the bottom of one page and
        its caption at the top of the next.
        """
        tbl = cell.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        self._no_borders(tbl, width_cm)
        self._no_split(tbl.rows[0])
        inner = tbl.rows[0].cells[0]
        inner.width = Cm(width_cm)
        return inner

    def picture_par(self, cell, filename, caption):
        path = os.path.join(self.assets, filename)
        cfg = self.cfg
        with Image.open(path) as im:
            w_px, h_px = im.size
            w_cm = cfg["max_w_cm"]
            h_cm = w_cm * h_px / w_px
            if h_cm > cfg["max_h_cm"]:
                h_cm = cfg["max_h_cm"]
                w_cm = h_cm * w_px / h_px
            if max(w_px, h_px) > cfg["embed_max_px"]:
                os.makedirs(self.tmp, exist_ok=True)
                scale = cfg["embed_max_px"] / float(max(w_px, h_px))
                small = im.convert("RGB").resize(
                    (max(1, int(w_px * scale)), max(1, int(h_px * scale))),
                    Image.LANCZOS)
                embed_path = os.path.join(self.tmp, filename)
                small.save(embed_path, "PNG", optimize=True)
            else:
                embed_path = path
        holder = self.wrapper(cell, cfg["max_w_cm"] + 0.6)
        par = holder.paragraphs[0]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_before = Pt(4)
        par.paragraph_format.space_after = Pt(2)
        self.set_indent(par, 0, 0)
        par.add_run().add_picture(embed_path, width=Cm(w_cm))
        cap = holder.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_before = Pt(0)
        cap.paragraph_format.space_after = Pt(4)
        self.set_indent(cap, 0, 0)
        self.add_text(cap, caption, size=cfg["cap_pt"])
        return par

    def table_par(self, cell, header, rows, caption, widths_cm=None):
        cfg = self.cfg
        n = len(header)
        if widths_cm is None:
            total = cfg["max_w_cm"]
            widths_cm = [total / n] * n
        holder = self.wrapper(cell, sum(widths_cm) + 0.4)
        for p in holder._tc.findall(qn("w:p")):
            holder._tc.remove(p)
        tbl = holder.add_table(rows=1 + len(rows), cols=n)
        tbl.style = "Table Grid"
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        for j, txt in enumerate(header):
            c = tbl.rows[0].cells[j]
            par = c.paragraphs[0]
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self.set_indent(par, 0, 0)
            self.add_text(par, txt, bold=True)
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:fill"), "D9D9D9")
            c._tc.get_or_add_tcPr().append(shd)
        for i, row in enumerate(rows, start=1):
            for j, txt in enumerate(row):
                c = tbl.rows[i].cells[j]
                par = c.paragraphs[0]
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                self.set_indent(par, 0, 0)
                self.add_text(par, txt)
        for j, w in enumerate(widths_cm):
            tbl.columns[j].width = Cm(w)
        for row in tbl.rows:
            self._no_split(row)
        cap = holder.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_before = Pt(2)
        cap.paragraph_format.space_after = Pt(6)
        self.set_indent(cap, 0, 0)
        self.add_text(cap, caption, size=cfg["cap_pt"])

    # -------------------------------------------------------------- filling --
    @staticmethod
    def clear_extras(cell):
        """Drop the template's empty placeholder paragraphs after the heading."""
        ps = cell._tc.findall(qn("w:p"))
        for p in ps[1:]:
            cell._tc.remove(p)

    def fill_cell(self, cell, blocks):
        self.clear_extras(cell)
        for blk in blocks:
            kind = blk[0]
            if kind == "h":
                self.heading_par(cell, blk[1])
            elif kind == "p":
                text = blk[1]
                indent = blk[2] if len(blk) > 2 else 200
                left = blk[3] if len(blk) > 3 else 0
                self.body_par(cell, text, indent=indent, left=left)
            elif kind == "eq":
                self.equation_par(cell, blk[1], blk[2] if len(blk) > 2 else None)
            elif kind == "img":
                self.picture_par(cell, blk[1], blk[2])
            elif kind == "tbl":
                self.table_par(cell, blk[1], blk[2], blk[3],
                               blk[4] if len(blk) > 4 else None)
            else:
                raise ValueError("unknown block type: %r" % (kind,))
        # a table cell must end with a paragraph
        cell.add_paragraph()

    def fill_cover(self, label, value):
        for par in self.doc.paragraphs:
            if par.text.strip().startswith(label):
                for run in par.runs:
                    if run.font.underline and run.text.strip():
                        run.text = value
                        return True
                for run in par.runs:
                    if run.font.underline:
                        run.text = value
                        return True
        return False

    def write(self, out_path, sections, rows, table_index=1):
        if os.path.abspath(out_path) == os.path.abspath(self.doc_part_path):
            raise SystemExit("refusing to overwrite the template")
        table = self.doc.tables[table_index]
        for idx, blocks in zip(rows, sections):
            self.fill_cell(table.rows[idx].cells[0], blocks)
        self.doc.save(out_path)
        return out_path


def load_spec(path):
    spec_dir = os.path.dirname(os.path.abspath(path))
    sys.path.insert(0, spec_dir)
    name = os.path.splitext(os.path.basename(path))[0]
    mod_spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(mod_spec)
    mod_spec.loader.exec_module(mod)
    return mod


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--spec", required=True)
    ap.add_argument("--assets", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--tmp-dir", default=None,
                    help="scratch dir for downscaled embed copies")
    args = ap.parse_args()

    spec = load_spec(args.spec)
    tmp = args.tmp_dir or os.path.join(os.path.dirname(os.path.abspath(args.out)),
                                       "_embed")
    rep = Report(args.template, args.assets, tmp,
                 getattr(spec, "CONFIG", None))

    name = getattr(spec, "EXPERIMENT_NAME", None)
    if name:
        print("cover 实验名称 filled:", rep.fill_cover("实验名称：", name))

    rows = getattr(spec, "SECTION_ROWS", [0, 1, 2, 3])
    table_index = getattr(spec, "TABLE_INDEX", 1)
    rep.write(args.out, spec.SECTIONS, rows, table_index)
    print("saved", args.out, os.path.getsize(args.out))


if __name__ == "__main__":
    main()
