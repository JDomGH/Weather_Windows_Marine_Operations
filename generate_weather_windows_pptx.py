from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

OUT = "Weather_Windows_Marine_Operations.pptx"
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

NAVY = RGBColor(16, 48, 77)
BLUE = RGBColor(39, 128, 185)
TEAL = RGBColor(34, 153, 153)
GREEN = RGBColor(39, 174, 96)
RED = RGBColor(192, 57, 43)
LIGHT = RGBColor(235, 243, 248)
DARK = RGBColor(35, 45, 55)
WHITE = RGBColor(255, 255, 255)


def text_box(slide, text, x, y, w, h, size=20, color=DARK, bold=False,
             align=PP_ALIGN.LEFT, font="Aptos"):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return shape


def title(slide, heading, subtitle=None):
    text_box(slide, heading, 0.55, 0.25, 12.2, 0.55, 28, NAVY, True)
    if subtitle:
        text_box(slide, subtitle, 0.58, 0.88, 12, 0.35, 13, BLUE)


def bullet_box(slide, items, x, y, w, h, size=18):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.font.name = "Aptos"
        p.font.size = Pt(size)
        p.font.color.rgb = DARK
        p.space_after = Pt(8)
        p._p.get_or_add_pPr().insert(0, p._p._new_buChar())
    return shape


def rounded_box(slide, text, x, y, w, h, fill, color=WHITE, size=17, bold=True):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid(); shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = fill
    tf = shape.text_frame; tf.clear(); tf.vertical_anchor = 1
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text; r.font.name = "Aptos"; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
    return shape

# Slide 1
s = prs.slides.add_slide(prs.slide_layouts[6])
title(s, "Weather Windows for Marine Operations", "When is the sea state safe enough to work?")
bullet_box(s, [
    "A weather window is a period when forecast conditions remain within the limits required for a marine operation.",
    "It supports safe, efficient decisions before mobilising crew, vessels, cranes, or equipment.",
    "Relevant conditions may include waves, wind, swell, current, visibility, and vessel motion."
], 0.7, 1.45, 5.3, 2.2, 17)
text_box(s, "Decision process", 7.0, 1.35, 4.8, 0.4, 20, NAVY, True, PP_ALIGN.CENTER)
labels = [("Forecast", 7.3, BLUE), ("Define\nlimits", 9.1, TEAL), ("Compare\nconditions", 10.9, BLUE)]
for label, x, c in labels: rounded_box(s, label, x, 2.05, 1.45, 0.85, c, size=15)
for x in (8.78, 10.58):
    line = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(2.47), Inches(x+0.28), Inches(2.47))
    line.line.color.rgb = NAVY; line.line.width = Pt(2)
rounded_box(s, "SAFE WINDOW?", 8.25, 3.55, 2.65, 0.65, NAVY, size=18)
rounded_box(s, "PROCEED", 7.1, 5.0, 2.0, 0.7, GREEN, size=16)
rounded_box(s, "DELAY / RESCHEDULE", 10.0, 5.0, 2.2, 0.7, RED, size=14)
for x1, x2 in ((9.0, 8.1), (10.15, 11.0)):
    line = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(4.2), Inches(x2), Inches(5.0))
    line.line.color.rgb = NAVY; line.line.width = Pt(2)
text_box(s, "The objective is to keep people, vessels, equipment, and the schedule within acceptable risk limits.", 0.8, 5.6, 11.8, 0.7, 19, NAVY, True, PP_ALIGN.CENTER)

# Slide 2
s = prs.slides.add_slide(prs.slide_layouts[6])
title(s, "Weather Criteria — DNV N001 Basis", "Criteria must be defined for the specific vessel, activity, and phase of work")
bullet_box(s, [
    "DNV N001 does not prescribe one universal weather threshold for every operation.",
    "The operation-specific limits should cover the environmental conditions and vessel responses that govern safety.",
    "A valid window exists only when all relevant criteria remain within their allowable values for the planned activity."
], 0.65, 1.25, 6.2, 2.1, 16)
text_box(s, "Typical criteria", 7.2, 1.25, 5.2, 0.4, 21, NAVY, True, PP_ALIGN.CENTER)
criteria = ["Significant wave height (Hs)", "Wind speed and gusts", "Swell / wave direction", "Current speed and direction", "Visibility", "Vessel motion: roll, pitch, heave", "Lifting, transfer, and access limits"]
for i, item in enumerate(criteria):
    y = 1.78 + i * 0.49
    rounded_box(s, item, 7.25, y, 5.0, 0.35, LIGHT, NAVY, 13, False)
rounded_box(s, "GREEN: within limits", 0.9, 4.45, 2.6, 0.62, GREEN, size=15)
rounded_box(s, "RED: limit exceeded", 4.0, 4.45, 2.6, 0.62, RED, size=15)
text_box(s, "Controlling criterion = the first limit reached", 7.1, 4.45, 5.0, 0.62, 18, NAVY, True, PP_ALIGN.CENTER)
text_box(s, "Example acceptance criteria", 0.75, 5.55, 3.2, 0.35, 17, NAVY, True)
examples = "Hs ≤ 1.5 m    |    Wind ≤ 15 m/s    |    Current ≤ 1.2 m/s    |    Visibility ≥ 2 nm    |    Roll ≤ 5°"
text_box(s, examples, 0.75, 5.95, 11.8, 0.55, 17, DARK, False, PP_ALIGN.CENTER)
text_box(s, "Always verify the project-specific design basis, operating manual, and approved marine procedure.", 0.75, 6.75, 11.8, 0.3, 11, BLUE, False, PP_ALIGN.CENTER)

# Slide 3
s = prs.slides.add_slide(prs.slide_layouts[6])
title(s, "Simple Case: Offshore Lift Operation", "Find the period when every criterion is satisfied")
text_box(s, "Scenario: A support vessel will lift a small module near an offshore platform.", 0.65, 1.15, 12, 0.38, 17, DARK)
text_box(s, "Limits: Hs ≤ 1.5 m  |  Wind ≤ 15 m/s  |  Visibility ≥ 2 nm  |  Current ≤ 1.2 m/s  |  Roll ≤ 5°", 0.65, 1.58, 12, 0.38, 15, NAVY, True)
# Timeline
x0, y0, total_w = 1.15, 3.0, 11.0
text_box(s, "Time", 0.35, y0-0.1, 0.6, 0.3, 15, NAVY, True)
segments = [("08:00–10:00", "NOT FEASIBLE\nHs 2.0 m; wind 18 m/s", RED, 2.2), ("10:00–12:30", "FEASIBLE\nHs 1.2 m; wind 12 m/s", GREEN, 2.35), ("12:30–14:00", "FEASIBLE\nHs 1.4 m; wind 14 m/s", GREEN, 2.0), ("14:00–16:00", "NOT FEASIBLE\nHs 1.8 m; wind 17 m/s", RED, 2.45)]
x = x0
for period, desc, color, w in segments:
    rounded_box(s, period, x, y0, w, 0.5, NAVY, size=13)
    rounded_box(s, desc, x, y0+0.63, w, 1.05, color, size=13)
    x += w
text_box(s, "WEATHER WINDOW", 4.15, 5.15, 4.95, 0.4, 19, GREEN, True, PP_ALIGN.CENTER)
text_box(s, "10:00–14:00", 4.15, 5.55, 4.95, 0.55, 25, NAVY, True, PP_ALIGN.CENTER)
text_box(s, "Decision: schedule the lift only inside this window; postpone if any criterion is forecast to exceed its limit.", 0.8, 6.35, 11.8, 0.55, 17, NAVY, True, PP_ALIGN.CENTER)
text_box(s, "Teaching point: a weather window is a validated period in time—not simply a general forecast of ‘good weather’. ", 0.8, 6.92, 11.8, 0.28, 11, BLUE, False, PP_ALIGN.CENTER)

prs.save(OUT)
print(f"Created {OUT}")
