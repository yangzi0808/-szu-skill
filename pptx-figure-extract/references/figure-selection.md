# Which pictures from a lab PPT belong in the report

## Keep

| Kind | Why |
| --- | --- |
| 原理图 / 示意图 (diffraction geometry, optical path, energy levels) | the report's 实验原理 leans on them |
| 装置图 / 实物照片 with labels | shows what 实验仪器 actually is |
| 关键操作照片 (lamp aligned to a slit, high-voltage knob at minimum, light spot on the entrance slit) | the safety- or success-critical step |
| 标准值表、参数表 | needed to perform the measurement, worth a small table |

## Drop

| Kind | Why |
| --- | --- |
| slide backgrounds, title bars, decorative rectangles, company logos | not content |
| images carrying a third-party watermark (Baidu Baike, a vendor logo burned into a screenshot) | do not reproduce another party's watermark in a student report; say in the hand-off that the figure was skipped for this reason |
| software menu screenshots of every step | one or two representative shots is enough; the prose already describes the menus |
| duplicated photos of the same operation | pick the clearest |
| very low-resolution media (for example a 200x153 jpeg) when it would be printed several centimetres wide | notify the user rather than shipping a blurry figure |

## Judging a picture that PowerPoint has cropped

A picture shape has its own width/height in EMU. Compare its aspect ratio with
the media file's pixel aspect ratio:

* equal (within ~2%) -> PowerPoint shows the whole file; the media file is the
  best source and is higher resolution than the render;
* different -> PowerPoint is showing part of the file (a `srcRect` crop), so crop
  the same region out of a slide render to reproduce what the audience sees.

The same rule applies to group shapes and to `box` unions of several shapes.

## Captioning

Every figure needs a caption that lets the reader trace it back:

```
图 {章节号}-{序号}  {图标题}（来源：PPT 第 {页} 页）
```

Use the deck's own caption when the slide has one (for example "图 16-2 光栅的衍射");
otherwise write a short factual caption from the surrounding slide text. Never
invent a caption that says more than the slide does.
