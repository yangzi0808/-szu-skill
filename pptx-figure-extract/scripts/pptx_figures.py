# -*- coding: utf-8 -*-
"""Extract reusable figures from a PowerPoint deck.

Three subcommands cover the whole job:

    dump    slide text / tables / notes, a shape->media map with geometry, and
            the media inventory (use this to decide which pictures are worth
            keeping and whether PowerPoint crops them)
    media   copy ppt/media/* out of the package
    build   produce the final figure set from a JSON spec, either by copying a
            media file or by cropping a region out of a rendered slide, and write
            a manifest

Only .pptx can be read directly. For a legacy .ppt, use the
`office-com-windows` skill: it reads .ppt through PowerPoint COM and can export
slides as PNG, which `build` can then crop from.

Examples:
    python pptx_figures.py dump deck.pptx work/ppt_dump
    python pptx_figures.py media deck.pptx work/media
    python pptx_figures.py build deck.pptx --slides-png work/slides_hi \
        --spec figures.json --out outputs/report_assets

Spec format (JSON list, `out` plus exactly one source):
    [
      {"out": "ppt_p02_img1.png", "media": "image3.png",
       "slide": 2, "desc": "氢原子能级跃迁图"},
      {"out": "ppt_p12_img14.png", "slide": 12, "shape": "对象 5",
       "desc": "开机提示对话框"},
      {"out": "ppt_p02_img2.png", "slide": 2,
       "box": [539433, 4941570, 3429000, 1643063], "desc": "发射光谱照片"}
    ]
`box` is [left, top, width, height] in EMU. Use `shape` when a shape's geometry
already matches what you want; use `box` to union several shapes or add padding.
"""
import argparse
import json
import os
import re
import sys
import zipfile

from PIL import Image
from pptx import Presentation

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def shape_media(shape):
    """Relationship targets of every picture fill inside a shape."""
    try:
        part = shape.part
        rids = []
        for blip in shape._element.findall(".//" + A_NS + "blip"):
            rid = blip.get(R_NS + "embed") or blip.get(R_NS + "link")
            if rid and rid in part.rels:
                rids.append(part.rels[rid].target_ref.split("/")[-1])
        return rids
    except Exception:
        return []


def walk(shapes, slide_no, depth, lines):
    for sh in shapes:
        entry = {
            "slide": slide_no,
            "name": sh.name,
            "type": str(sh.shape_type),
            "left": getattr(sh, "left", None),
            "top": getattr(sh, "top", None),
            "width": getattr(sh, "width", None),
            "height": getattr(sh, "height", None),
            "media": shape_media(sh),
        }
        lines.append("  " * depth + json.dumps(entry, ensure_ascii=False))
        if sh.__class__.__name__ == "GroupShape":
            walk(sh.shapes, slide_no, depth + 1, lines)


def cmd_dump(args):
    prs = Presentation(args.pptx)
    os.makedirs(args.out_dir, exist_ok=True)
    text_lines = ["slide_size_emu=%s x %s" % (prs.slide_width, prs.slide_height),
                  "num_slides=%d" % len(prs.slides)]
    shape_lines = []
    for i, slide in enumerate(prs.slides, 1):
        text_lines.append("")
        text_lines.append("=" * 60)
        text_lines.append("SLIDE %d" % i)
        try:
            if slide.has_notes_slide:
                nt = slide.notes_slide.notes_text_frame.text
                if nt.strip():
                    text_lines.append("NOTES: " + nt.replace("\n", " | "))
        except Exception:
            pass
        for sh in slide.shapes:
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    t = "".join(r.text for r in p.runs)
                    if t.strip():
                        text_lines.append("  [%s] %s" % (sh.name, t))
            if getattr(sh, "has_table", False) and sh.has_table:
                text_lines.append("  TABLE %s:" % sh.name)
                for row in sh.table.rows:
                    text_lines.append("    " + " | ".join(
                        c.text.replace("\n", "/") for c in row.cells))
        walk(slide.shapes, i, 0, shape_lines)

    with open(os.path.join(args.out_dir, "slides.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(text_lines))
    with open(os.path.join(args.out_dir, "shape_map.jsonl"), "w",
              encoding="utf-8") as f:
        f.write("\n".join(shape_lines))

    z = zipfile.ZipFile(args.pptx)
    media = sorted(n for n in z.namelist()
                   if n.startswith("ppt/media/") and not n.endswith("/"))
    with open(os.path.join(args.out_dir, "media_list.txt"), "w",
              encoding="utf-8") as f:
        for n in media:
            f.write("%s\t%d\n" % (n, z.getinfo(n).file_size))

    by_slide = {}
    for n in z.namelist():
        m = re.match(r"ppt/slides/_rels/slide(\d+)\.xml\.rels$", n)
        if not m:
            continue
        idx = int(m.group(1))
        for t in re.findall(r'Target="([^"]+)"',
                            z.read(n).decode("utf-8", "ignore")):
            if "media/" in t:
                by_slide.setdefault(idx, set()).add(t.split("media/")[-1])
    with open(os.path.join(args.out_dir, "media_by_slide.txt"), "w",
              encoding="utf-8") as f:
        for k in sorted(by_slide):
            f.write("slide%02d: %s\n" % (k, ", ".join(sorted(by_slide[k]))))

    print("dumped %d slides -> %s" % (len(prs.slides), args.out_dir))
    print("check shape_map.jsonl against media_list.txt: a picture whose shape "
          "ratio differs from its media ratio is being cropped by PowerPoint, "
          "so crop it from the slide render instead of copying the media file.")


def cmd_media(args):
    os.makedirs(args.out_dir, exist_ok=True)
    z = zipfile.ZipFile(args.pptx)
    n = 0
    for name in z.namelist():
        if name.startswith("ppt/media/") and not name.endswith("/"):
            with open(os.path.join(args.out_dir, os.path.basename(name)), "wb") as f:
                f.write(z.read(name))
            n += 1
    print("extracted %d media files -> %s" % (n, args.out_dir))


def cmd_build(args):
    prs = Presentation(args.pptx)
    slide_w, slide_h = prs.slide_width, prs.slide_height
    with open(args.spec, encoding="utf-8") as f:
        spec = json.load(f)
    os.makedirs(args.out, exist_ok=True)
    media_dir = args.media_dir or os.path.join(os.path.dirname(args.spec), "media")

    index = {}
    for i, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            index[(i, sh.name)] = sh
            if sh.__class__.__name__ == "GroupShape":
                for sub in sh.shapes:
                    index.setdefault((i, sub.name), sub)

    manifest = ["文件名\t尺寸\t来源\t内容\t用途"]
    for item in spec:
        name = item["out"]
        dst = os.path.join(args.out, name)
        page = item.get("slide")
        if "media" in item:
            im = Image.open(os.path.join(media_dir, item["media"])).convert("RGB")
            im.save(dst, "PNG")
            origin = "PPT第%s页" % page if page else "PPT原件"
        else:
            box = item.get("box")
            if box is None:
                sh = index[(item["slide"], item["shape"])]
                box = [sh.left, sh.top, sh.width, sh.height]
            render = Image.open(os.path.join(
                args.slides_png, "slide-%02d.png" % item["slide"])).convert("RGB")
            sx = render.width / float(slide_w)
            sy = render.height / float(slide_h)
            px = (int(round(box[0] * sx)), int(round(box[1] * sy)),
                  int(round((box[0] + box[2]) * sx)),
                  int(round((box[1] + box[3]) * sy)))
            im = render.crop(px)
            im.save(dst, "PNG")
            origin = "PPT第%s页" % item["slide"]
        manifest.append("%s\t%dx%d\t%s\t%s\t%s" % (
            name, im.width, im.height, origin,
            item.get("desc", ""), item.get("use", "已采用")))
        print(name, im.size, item.get("desc", ""))

    with open(os.path.join(args.out, "图片清单.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(manifest))
    print("built %d figures -> %s" % (len(spec), args.out))


def main():
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("dump", help="text/tables/notes + shape map + media inventory")
    p.add_argument("pptx")
    p.add_argument("out_dir")
    p.set_defaults(func=cmd_dump)

    p = sub.add_parser("media", help="extract ppt/media/*")
    p.add_argument("pptx")
    p.add_argument("out_dir")
    p.set_defaults(func=cmd_media)

    p = sub.add_parser("build", help="produce the final figure set from a JSON spec")
    p.add_argument("pptx")
    p.add_argument("--spec", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--slides-png", help="directory of slide-NN.png renders")
    p.add_argument("--media-dir", help="directory holding the extracted media")
    p.set_defaults(func=cmd_build)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
