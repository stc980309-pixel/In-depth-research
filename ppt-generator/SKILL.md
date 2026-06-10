---
name: ppt-generator
description: "Generate well-designed PowerPoint presentations (PPTX) using MS Office PowerPoint COM on Windows. Creates professional slides with academic/business styling, 16:9 layout, Chinese font support. Use when users ask to create PPT, PowerPoint, 演示文稿, 幻灯片, or presentation about any topic. Handles cover pages, table of contents, grids, timelines, stats, and more."
metadata:
  version: "1.0.0"
  last_updated: "2026-05-30"
  status: active
  platform: windows
compatibility: "Requires Windows + MS Office PowerPoint + pywin32"
allowed-tools:
  - Bash
  - Write
  - Read
  - Edit
---

# PPT Generator — Create PowerPoint Presentations via COM

Generate professional .pptx files using MS Office PowerPoint COM automation on Windows. Do NOT use `cli-anything-wps` — it only works with WPS Office, not MS Office.

## Quick Start

When a user asks to create a PPT about a topic:

1. **Understand requirements** — topic, audience, slide count preference, style
2. **Design slide structure** — 9 slides is the sweet spot for introductions
3. **Choose color scheme** — use the academic preset unless user specifies otherwise
4. **Write Python script** — adapt from the patterns below
5. **Execute** — run with `D:/py/Python310/python.exe`

## Required Setup

- Python: `D:/py/Python310/python.exe`
- Package: `pywin32` (already installed)
- MS Office PowerPoint must be installed

## Color Scheme (Academic Default)

```
PRIMARY_BLUE  = (26, 60, 139)    # Deep blue for backgrounds, banners
ACCENT_ORANGE = (230, 119, 51)   # Orange for highlights, accents
ACCENT_GREEN  = (24, 128, 80)    # Green for variety in multi-col layouts
ACCENT_CYAN   = (0, 168, 232)    # Cyan for variety
WHITE         = (255, 255, 255)
DARK          = (34, 34, 34)
LIGHT_BG      = (245, 248, 252)  # Light gray for card backgrounds
GRAY          = (128, 128, 128)
```

## Slide Structure (9-Slide Default for Introductions)

```
1. cover       — Title + subtitle + institution info
2. toc         — Table of contents (numbered items)
3. overview    — Summary text + key stat cards on right
4. timeline    — History/chronology with dot+line markers
5. three_col   — Three-column comparison (disciplines/features/categories)
6. grid_cards  — 2x4 or 4x2 card grid (people/schools/products)
7. quadrant    — 2x2 quadrant with colored headers
8. stats       — 8 stat cards (number + label + description)
9. closing     — Thank you + contact info + motto
```

## Critical Gotchas

1. **New presentations may have 0 slides**: Always check `pres.Slides.Count` before accessing `pres.Slides(1)`. If 0, use `pres.Slides.Add(1, 2)` for the first slide.

2. **One-line try/except**: Avoid `try: ... except: pass` on one line. Use multi-line:
   ```python
   try:
       s.Delete()
   except Exception:
       pass
   ```

3. **Set PageSetup before adding slides**: Set `pres.PageSetup.SlideWidth` and `SlideHeight` right after creating the presentation. 16:9 = 960x540.

4. **Delete default shapes**: Each new slide comes with placeholder shapes. Delete them before adding custom content.

5. **COM must be initialized**: Call `pythoncom.CoInitialize()` at the start and `pythoncom.CoUninitialize()` at the end.

6. **RGB format**: PowerPoint COM uses `(r << 16) | (g << 8) | b` integer format, not tuples.

7. **Font for Chinese**: Use `"Microsoft YaHei"` (微软雅黑) for Chinese text. It's pre-installed on Windows.

## Script Template

See `references/pptx-com-api.md` for the complete COM API reference. Every script follows this skeleton:

```python
"""Generate PPT: [topic]"""
import win32com.client
import pythoncom
import os

pythoncom.CoInitialize()

def rgb(r, g, b):
    return (r << 16) | (g << 8) | b

pp = win32com.client.Dispatch("PowerPoint.Application")
pp.Visible = True
pres = pp.Presentations.Add()
pres.PageSetup.SlideWidth = 960
pres.PageSetup.SlideHeight = 540

# --- Helper functions ---
def add_text_box(slide, left, top, width, height, text, font_size=18,
                 bold=False, color=(34,34,34), alignment=1, font_name="Microsoft YaHei"):
    shape = slide.Shapes.AddTextbox(1, left, top, width, height)
    tf = shape.TextFrame
    tf.WordWrap = -1
    tf.TextRange.Text = text
    tf.TextRange.Font.Size = font_size
    tf.TextRange.Font.Bold = bold
    tf.TextRange.Font.Color = rgb(*color)
    tf.TextRange.Font.Name = font_name
    tf.TextRange.ParagraphFormat.Alignment = alignment
    return shape

def add_rect(slide, left, top, width, height, fill_color, line_visible=False):
    shape = slide.Shapes.AddShape(1, left, top, width, height)
    shape.Fill.ForeColor.RGB = rgb(*fill_color)
    shape.Fill.Visible = -1
    shape.Line.Visible = 0 if not line_visible else -1
    return shape

# --- Build slides ---
# [Each slide: handle first slide specially, delete default shapes, add content]

# --- Save ---
output_path = r"C:\Users\Admin\Desktop\[name].pptx"
if os.path.exists(output_path):
    os.remove(output_path)
pres.SaveAs(output_path)
print(f"Saved: {output_path}")
pythoncom.CoUninitialize()
```

## Slide Layout Reference

**Cover**: Blue background (960x540), top+bottom accent lines (15px), title 60pt white centered, subtitle 28pt accent centered, detail line 16pt.

**TOC**: Left blue decorative bar (18px wide), numbering blocks (50x42 rounded rects), item labels 26pt.

**Overview**: Top blue banner (960x90) with 38pt white title, body text 18pt, stat cards on right side (138x150 light gray).

**Timeline**: Top blue banner, circles (16px) connected by vertical lines (3px), date 18pt bold blue, event 16pt.

**Three columns**: Dark full-width background, three cards (290px each) with colored headers, white content areas.

**Grid cards**: 4 columns x 2 rows, cards (215x185 light gray) with top accent strip, circle avatar area, title 18pt, subtitle 13pt.

**Quadrant**: 2x2 grid, each quadrant (435x195) with colored header (22pt white) and white content area.

**Stats**: 4x2 grid, cards (215x185) with large number (36pt blue), label (18pt orange), desc (14pt gray).

**Closing**: Blue background, thank you 50pt white, divider lines, contact info 18pt, motto 28pt accent.

## Execution

Always run the script with the known Python path and capture output:

```
"D:/py/Python310/python.exe" "C:/Users/Admin/Desktop/create_ppt.py"
```

If the script succeeds, tell the user the file path and slide count. If it fails, read the error, fix the script, and re-run. Common fixes:
- `pywintypes.com_error` with "Integer out of range" → slide count is 0, use `Slides.Add(1, 2)`
- `SyntaxError` → check for one-line try/except or encoding issues
- `AttributeError` → check COM method names (no typos)
