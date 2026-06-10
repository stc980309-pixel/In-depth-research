# PowerPoint COM API Quick Reference

## Setup

```python
import win32com.client
import pythoncom
pythoncom.CoInitialize()

pp = win32com.client.Dispatch("PowerPoint.Application")
pp.Visible = True  # True for debugging, False for background
pres = pp.Presentations.Add()
pres.PageSetup.SlideWidth = 960   # 16:9
pres.PageSetup.SlideHeight = 540
```

## Slide Operations

```python
# Add slide (layout 2 = ppLayoutText, has title + content placeholders)
slide = pres.Slides.Add(index, 2)

# Delete slide
pres.Slides(index).Delete()

# Delete all shapes on a slide
for s in list(slide.Shapes):
    try:
        s.Delete()
    except Exception:
        pass

# Get slide count
count = pres.Slides.Count
```

## Color Format

```python
def rgb(r, g, b):
    return (r << 16) | (g << 8) | b

# Common colors:
PRIMARY_BLUE = (26, 60, 139)
ACCENT_ORANGE = (230, 119, 51)
WHITE = (255, 255, 255)
DARK = (34, 34, 34)
LIGHT_BG = (245, 248, 252)
GRAY = (128, 128, 128)
```

## Adding Shapes

```python
# Text box: AddTextbox(orientation=1, left, top, width, height)
shape = slide.Shapes.AddTextbox(1, left, top, width, height)
tf = shape.TextFrame
tf.WordWrap = -1  # msoTrue
tf.TextRange.Text = "content"
tf.TextRange.Font.Size = 18
tf.TextRange.Font.Bold = True / False
tf.TextRange.Font.Color = rgb(r, g, b)
tf.TextRange.Font.Name = "Microsoft YaHei"
tf.TextRange.ParagraphFormat.Alignment = 1  # 1=left, 2=center, 3=right

# Rectangle: AddShape(type=1, left, top, width, height)
shape = slide.Shapes.AddShape(1, left, top, width, height)
shape.Fill.ForeColor.RGB = rgb(r, g, b)
shape.Fill.Visible = -1  # msoTrue
shape.Line.Visible = 0   # no border

# Circle: AddShape(type=9, left, top, width, height)
dot = slide.Shapes.AddShape(9, x, y, w, h)
dot.Fill.ForeColor.RGB = rgb(r, g, b)
dot.Line.Visible = 0
```

## Save

```python
pres.SaveAs(output_path)
print(f"Saved: {output_path}")
pythoncom.CoUninitialize()
```

## Common Layout Patterns

### Cover slide
- Full-width colored background (960x540 rect)
- Decorative lines (thin rects) at top/bottom with accent color
- Large title (60pt) centered, white
- Subtitle (28pt) centered, accent color
- Small detail line at bottom

### Section header
- Full-width colored banner at top (960x90)
- Title (38pt) centered in banner, white text

### Grid cards
- Loop: col = i % cols, row = i // cols
- x = margin + col * card_width, y = start_y + row * card_height
- Card background + top accent strip + content text

### Timeline
- Left circle markers connected by vertical line segments
- Date text (18pt bold) + event text (16pt) beside each marker

### Stats grid
- Large number (36pt bold) + label (18pt) + description (14pt) per card
```
