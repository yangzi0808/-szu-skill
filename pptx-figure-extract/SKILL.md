---
name: pptx-figure-extract
description: Pull the useful figures out of a PowerPoint deck - extract embedded media, crop annotated or cropped regions out of rendered slides, name them consistently and write a manifest. Use when a task needs figures/diagrams/screenshots taken from a .pptx (or .ppt) for reuse in another document, and when deciding which pictures are content rather than decoration.
metadata:
  short-description: 从实验 PPT 提取图片、切图与清单
---

# Extracting figures from a PowerPoint deck

Slides contain a mix of real content figures and decoration, and the picture you
see on a slide is often a *crop* of the stored image with arrows, circles or
labels drawn on top. Copying `ppt/media/*` blindly therefore produces either
decoration or a wrong crop.

## Workflow

1. **Dump first.**
   `python scripts/pptx_figures.py dump deck.pptx work/ppt_dump`
   gives `slides.txt` (text/tables/notes per slide), `shape_map.jsonl` (every
   shape with its EMU geometry and the media file behind it) and
   `media_by_slide.txt`.
2. **Render the slides** so you can see what each picture actually shows. Use the
   `office-com-windows` skill's `pptx_dump.ps1` with `PPT_PNG_DIR` set; export at
   2560x1920 so crops stay sharp.
3. **Decide per picture** (read `references/figure-selection.md`):
   - shape aspect ratio == media aspect ratio -> the media file is used whole,
     copy it with `media`;
   - ratios differ, or the slide adds arrows/circles/labels -> crop that shape's
     region out of the slide render with `build`;
   - logos, backgrounds, footers, watermarks, decorative bullets -> drop.
4. **Build the figures and the manifest** from a JSON spec:

   ```bash
   python scripts/pptx_figures.py build deck.pptx \
       --slides-png work/slides_hi --media-dir work/media \
       --spec figures.json --out outputs/report_assets
   ```

   The manifest is written as `图片清单.txt` with columns
   文件名 / 尺寸 / 来源 / 内容 / 用途, so a reviewer can see which PPT page each
   figure came from and which ones were deliberately left out.

## Naming

`ppt_p{page:02d}_img{k}.{ext}` - page is the PowerPoint page number, k is the
index of that figure within the page. Sortable, traceable, and easy to cite in a
caption as "来源：PPT 第 N 页".

## Legacy `.ppt`

`python-pptx` cannot read the binary `.ppt` format. Use the
`office-com-windows` skill to read it through PowerPoint COM (slide text, tables,
notes and slide PNGs all work). If you need the embedded media itself, ask
PowerPoint to save a `.pptx` copy to a path **outside the user profile** (for
example `C:\CodexTmp`) and then run the dump/media commands on that copy.

## Resolution and cropping

- Crop from a render at least as large as the intended print size; a 2560x1920
  slide render keeps a half-slide figure around 300 dpi at 10 cm wide.
- When several shapes belong to one figure (a picture plus its label, or the two
  dialogs of one step), pass a `box` that unions them instead of using `shape`.
- Do not upscale media files whose own pixels are smaller than the target; crop
  from the render instead, or ask the user whether the low-resolution source is
  acceptable.
