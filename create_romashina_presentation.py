from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


OUT = Path("/workspace/Ромашина О.В. презентация.pptx")
ASSETS = Path("/home/ubuntu/.cursor/projects/workspace/assets")
MISSION_PHOTO = ASSETS / "b8cadcd5-b1a8-4bc2-b789-82acb5443140.jpg"
VALUES_PHOTO = ASSETS / "cb43bb22-7768-4187-817d-8a77cfb71e1d.jpg"
WORK_PHOTO = ASSETS / "d32f3dd0-611d-4581-8b50-384b9844b818.jpg"
DESK_SIDE_PHOTO = ASSETS / "e34c7407-6386-4676-be20-f8da82ad0683.jpg"
DESK_FRONT_PHOTO = ASSETS / "67a66b21-a2f8-4dde-b70b-aab2b1f59db5.jpg"
PORTRAIT_PHOTO = ASSETS / "994691e1-b4e7-4356-8398-0b28dbbc9f29.jpg"

prs = Presentation()
prs.slide_width = Inches(13.333333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

NAVY = RGBColor(20, 42, 114)
BLUE = RGBColor(28, 92, 166)
CYAN = RGBColor(0, 166, 227)
MAGENTA = RGBColor(215, 24, 106)
CORAL = RGBColor(239, 102, 109)
INK = RGBColor(30, 41, 59)
MUTED = RGBColor(93, 108, 130)
LIGHT = RGBColor(244, 247, 251)
PALE_BLUE = RGBColor(233, 242, 251)
PALE_PINK = RGBColor(252, 234, 242)
WHITE = RGBColor(255, 255, 255)
LINE = RGBColor(217, 225, 236)


def rect(slide, x, y, w, h, fill, radius=False, line=None, transparency=0):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.fill.transparency = transparency
    shape.line.color.rgb = fill if line is None else line
    if radius:
        shape.adjustments[0] = 0.12
    return shape


def textbox(
    slide, text, x, y, w, h, size=20, color=INK, bold=False,
    font="Aptos", align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
    margin=0, italic=False, line_spacing=1.0,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box


def rich_text(slide, lines, x, y, w, h, margin=0, valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    for i, item in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item.get("text", "")
        p.font.name = item.get("font", "Aptos")
        p.font.size = Pt(item.get("size", 18))
        p.font.bold = item.get("bold", False)
        p.font.italic = item.get("italic", False)
        p.font.color.rgb = item.get("color", INK)
        p.alignment = item.get("align", PP_ALIGN.LEFT)
        p.space_after = Pt(item.get("after", 6))
        p.space_before = Pt(item.get("before", 0))
        p.line_spacing = item.get("spacing", 1.0)
    return box


def crop_picture(slide, path, x, y, w, h):
    pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))
    with Image.open(path) as im:
        image_ratio = im.width / im.height
    box_ratio = w / h
    if image_ratio > box_ratio:
        visible = box_ratio / image_ratio
        crop = (1 - visible) / 2
        pic.crop_left = crop
        pic.crop_right = crop
    else:
        visible = image_ratio / box_ratio
        crop = (1 - visible) / 2
        pic.crop_top = crop
        pic.crop_bottom = crop
    return pic


def brand_bar(slide):
    rect(slide, 0, 7.36, 8.7, 0.14, NAVY)
    rect(slide, 8.7, 7.36, 1.6, 0.14, MAGENTA)
    rect(slide, 10.3, 7.36, 3.033, 0.14, CYAN)


def title(slide, kicker, heading, number):
    textbox(slide, kicker.upper(), 0.72, 0.42, 8.8, 0.3, 11, MAGENTA, True)
    textbox(slide, heading, 0.72, 0.78, 11.4, 0.55, 28, NAVY, True)
    textbox(slide, f"{number:02d}", 12.05, 0.48, 0.55, 0.3, 11, MUTED, True, align=PP_ALIGN.RIGHT)
    rect(slide, 0.72, 1.42, 11.88, 0.02, LINE)


def footer(slide, number, dark=False):
    color = WHITE if dark else MUTED
    textbox(slide, "РОМАШИНА ОЛЬГА ВИКТОРОВНА", 0.72, 7.12, 4.2, 0.18, 8, color, True)
    textbox(slide, str(number), 12.0, 7.12, 0.6, 0.18, 8, color, True, align=PP_ALIGN.RIGHT)


def add_bullet_card(slide, x, y, w, h, num, heading, body, accent):
    card = rect(slide, x, y, w, h, WHITE, radius=True, line=LINE)
    card.line.width = Pt(1)
    rect(slide, x, y, 0.08, h, accent)
    rect(slide, x + 0.25, y + 0.23, 0.46, 0.46, accent, radius=True)
    textbox(slide, str(num), x + 0.25, y + 0.25, 0.46, 0.35, 13, WHITE, True, align=PP_ALIGN.CENTER)
    textbox(slide, heading, x + 0.84, y + 0.20, w - 1.04, 0.34, 15, NAVY, True)
    textbox(slide, body, x + 0.84, y + 0.58, w - 1.04, h - 0.70, 11.5, MUTED)


# Slide 1 — title
s = prs.slides.add_slide(blank)
rect(s, 0, 0, 5.7, 7.5, NAVY)
rect(s, 5.52, 0, 0.18, 7.5, MAGENTA)
rect(s, 5.7, 0, 7.633, 7.5, NAVY)
crop_picture(s, PORTRAIT_PHOTO, 6.98, 0, 5.65, 7.5)
rect(s, 6.98, 0, 5.65, 7.5, NAVY, transparency=88)
textbox(s, "СОЦИАЛЬНЫЙ ФОНД РОССИИ", 0.74, 0.62, 4.25, 0.32, 11, CYAN, True)
textbox(s, "ОДИН ДЕНЬ\nРАБОТЫ В СФР", 0.74, 1.38, 4.45, 1.58, 30, WHITE, True)
rect(s, 0.74, 3.22, 1.1, 0.07, MAGENTA)
rich_text(s, [
    {"text": "Ромашина Ольга Викторовна", "size": 21, "bold": True, "color": WHITE, "after": 10},
    {"text": "Ведущий специалист-эксперт", "size": 14, "bold": True, "color": WHITE, "after": 5},
    {"text": "отдела установления мер социальной поддержки семьям с детьми", "size": 12.5, "color": RGBColor(220, 229, 245), "spacing": 1.05, "after": 5},
    {"text": "Управление социального обеспечения семей с детьми", "size": 12.5, "color": RGBColor(220, 229, 245)},
], 0.74, 3.58, 4.3, 2.2)
textbox(s, "ЗАБОТА • ПОЛЕЗНОСТЬ • НАДЁЖНОСТЬ", 0.74, 6.78, 4.4, 0.25, 9, WHITE, True)
brand_bar(s)


# Slide 2 — professional path
s = prs.slides.add_slide(blank)
s.background.fill.solid()
s.background.fill.fore_color.rgb = LIGHT
title(s, "Профессиональный путь", "Опыт, знания, ответственность", 2)
rect(s, 0.72, 1.86, 3.20, 4.72, NAVY, radius=True)
textbox(s, "19", 1.04, 2.16, 2.55, 1.03, 50, WHITE, True, align=PP_ALIGN.CENTER)
textbox(s, "лет профессионального\nопыта", 1.04, 3.18, 2.55, 0.7, 18, WHITE, True, align=PP_ALIGN.CENTER)
rect(s, 1.24, 4.10, 2.15, 0.03, RGBColor(91, 130, 195))
textbox(s, "3,5 года", 1.04, 4.47, 2.55, 0.54, 27, CYAN, True, align=PP_ALIGN.CENTER)
textbox(s, "в Управлении социального\nобеспечения семей с детьми", 1.02, 5.14, 2.60, 0.75, 12.5, WHITE, align=PP_ALIGN.CENTER)

items = [
    ("ОБРАЗОВАНИЕ", "Московский государственный университет экономики, статистики и информатики", MAGENTA),
    ("ДОЛЖНОСТЬ", "Ведущий специалист-эксперт отдела установления мер социальной поддержки семьям с детьми", CYAN),
    ("ПРОФЕССИОНАЛЬНЫЙ ФОКУС", "Точная оценка права на поддержку, внимание к обстоятельствам семьи и понятная помощь гражданам", BLUE),
]
for i, (head, body, accent) in enumerate(items):
    y = 1.86 + i * 1.57
    card = rect(s, 4.28, y, 8.32, 1.28, WHITE, radius=True, line=LINE)
    card.line.width = Pt(1)
    rect(s, 4.28, y, 0.10, 1.28, accent)
    textbox(s, head, 4.70, y + 0.21, 2.45, 0.25, 11, accent, True)
    textbox(s, body, 4.70, y + 0.54, 7.42, 0.52, 14, INK, i == 1)
footer(s, 2)
brand_bar(s)


# Slide 3 — responsibilities
s = prs.slides.add_slide(blank)
s.background.fill.solid()
s.background.fill.fore_color.rgb = LIGHT
title(s, "Моя роль и обязанности", "Каждое решение — в интересах семьи", 3)
cards = [
    ("Заявления на единое пособие", "Рассматриваю заявления и контролирую полноту сведений."),
    ("Проверка документов", "Проверяю документы и корректность представленных данных."),
    ("Комплексная оценка", "Провожу комплексную оценку нуждаемости семьи."),
    ("Сложные ситуации", "Помогаю разобраться в нестандартных жизненных обстоятельствах."),
    ("Консультирование", "Разъясняю порядок назначения детских пособий, в том числе гражданам, проходящим военную службу."),
    ("Взаимодействие", "Сотрудничаю с другими отделами Управления и организациями."),
]
for i, (head, body) in enumerate(cards):
    col, row = i % 2, i // 2
    add_bullet_card(s, 0.72 + col * 6.0, 1.75 + row * 1.68, 5.62, 1.38, i + 1, head, body, MAGENTA if i in (0, 3, 4) else BLUE)
footer(s, 3)
brand_bar(s)


# Slide 4 — mission
s = prs.slides.add_slide(blank)
s.background.fill.solid()
s.background.fill.fore_color.rgb = LIGHT
title(s, "Миссия СФР", "Забота о благополучии каждого человека", 4)
crop_picture(s, MISSION_PHOTO, 0.72, 1.73, 4.65, 4.98)
frame = rect(s, 0.72, 1.73, 4.65, 4.98, WHITE, radius=True, line=LINE, transparency=100)
frame.line.width = Pt(1.5)
textbox(s, "МИССИЯ", 5.86, 1.87, 2.0, 0.3, 11, MAGENTA, True)
textbox(s, "Мы заботимся\nо благополучии\nкаждого человека", 5.86, 2.25, 6.20, 1.48, 28, NAVY, True)
rect(s, 5.86, 3.98, 1.10, 0.06, MAGENTA)
textbox(
    s,
    "Реализуя социальную политику государства, мы обеспечиваем простоту и удобство получения гражданами мер социальной поддержки на основе современных технологий.",
    5.86, 4.38, 6.20, 1.25, 16, INK, False, line_spacing=1.12,
)
rect(s, 5.86, 5.90, 6.20, 0.70, PALE_BLUE, radius=True)
textbox(s, "Моя работа помогает воплощать эту миссию каждый день.", 6.17, 6.08, 5.62, 0.3, 14, BLUE, True)
footer(s, 4)
brand_bar(s)


# Slide 5 — values
s = prs.slides.add_slide(blank)
s.background.fill.solid()
s.background.fill.fore_color.rgb = LIGHT
title(s, "Ценности фонда", "Ориентиры в ежедневной работе", 5)
values = [
    ("ЗАБОТА", "внимание к человеку"),
    ("ПОЛЕЗНОСТЬ", "реальная помощь"),
    ("ДОСТУПНОСТЬ И ОТКРЫТОСТЬ", "понятные решения"),
    ("НАДЁЖНОСТЬ", "ответственность за результат"),
    ("САМОСОВЕРШЕНСТВОВАНИЕ", "новые знания и навыки"),
    ("ЕДИНСТВО", "работа одной командой"),
]
for i, (head, body) in enumerate(values):
    col, row = i % 2, i // 2
    x, y = 0.72 + col * 2.92, 1.79 + row * 1.44
    fill = PALE_PINK if i in (0, 3, 4) else PALE_BLUE
    accent = MAGENTA if i in (0, 3, 4) else BLUE
    rect(s, x, y, 2.66, 1.14, fill, radius=True, line=fill)
    textbox(s, head, x + 0.18, y + 0.18, 2.30, 0.30, 10.5, accent, True, align=PP_ALIGN.CENTER)
    textbox(s, body, x + 0.18, y + 0.60, 2.30, 0.28, 10.5, INK, align=PP_ALIGN.CENTER)
crop_picture(s, VALUES_PHOTO, 6.76, 1.73, 5.84, 4.98)
frame = rect(s, 6.76, 1.73, 5.84, 4.98, WHITE, radius=True, line=LINE, transparency=100)
frame.line.width = Pt(1.5)
footer(s, 5)
brand_bar(s)


# Slide 6 — workday in photos
s = prs.slides.add_slide(blank)
s.background.fill.solid()
s.background.fill.fore_color.rgb = LIGHT
title(s, "Один день работы в СФР", "Рабочий день в кадре", 6)
photo_data = [
    (DESK_SIDE_PHOTO, "РАБОТА С ЗАЯВЛЕНИЯМИ", "Внимание к деталям"),
    (DESK_FRONT_PHOTO, "АНАЛИЗ ДОКУМЕНТОВ", "Точность каждого решения"),
    (WORK_PHOTO, "КОНСУЛЬТИРОВАНИЕ", "Помощь в сложной ситуации"),
]
for i, (photo, head, caption) in enumerate(photo_data):
    x = 0.72 + i * 4.02
    crop_picture(s, photo, x, 1.72, 3.67, 4.42)
    rect(s, x, 5.15, 3.67, 0.99, NAVY, transparency=10)
    textbox(s, head, x + 0.18, 5.33, 3.31, 0.24, 10.5, WHITE, True, align=PP_ALIGN.CENTER)
    textbox(s, caption, x + 0.18, 5.70, 3.31, 0.23, 10.5, RGBColor(220, 231, 246), align=PP_ALIGN.CENTER)
textbox(s, "За каждым заявлением — конкретная семья.", 2.35, 6.50, 8.63, 0.36, 17, NAVY, True, align=PP_ALIGN.CENTER, italic=True)
footer(s, 6)
brand_bar(s)


# Slide 7 — conclusion
s = prs.slides.add_slide(blank)
rect(s, 0, 0, 13.333, 7.5, NAVY)
rect(s, 0, 0, 0.18, 7.5, MAGENTA)
rect(s, 10.76, 0, 2.573, 7.5, BLUE)
crop_picture(s, WORK_PHOTO, 9.06, 0.82, 3.56, 5.42)
rect(s, 9.06, 0.82, 3.56, 5.42, NAVY, transparency=82)
textbox(s, "ОДИН ДЕНЬ РАБОТЫ В СФР", 0.86, 0.90, 5.80, 0.3, 11, CYAN, True)
textbox(s, "СПАСИБО\nЗА ВНИМАНИЕ!", 0.86, 1.52, 7.55, 1.24, 34, WHITE, True)
rect(s, 0.86, 3.08, 1.12, 0.06, MAGENTA)
textbox(
    s,
    "Опыт помогает принимать взвешенные решения, а ценности Фонда — сохранять внимание к каждому человеку.",
    0.86, 3.56, 7.08, 1.05, 18, RGBColor(222, 231, 246), False, line_spacing=1.14,
)
rect(s, 0.86, 5.15, 6.98, 0.86, WHITE, radius=True, transparency=91)
textbox(s, "Забота о людях — основа нашей работы.", 1.15, 5.39, 6.38, 0.34, 17, WHITE, True)
textbox(s, "РОМАШИНА ОЛЬГА ВИКТОРОВНА", 0.86, 6.62, 5.1, 0.26, 10, WHITE, True)
textbox(s, "07", 12.00, 6.98, 0.60, 0.2, 9, WHITE, True, align=PP_ALIGN.RIGHT)
brand_bar(s)


prs.core_properties.title = "Один день работы в СФР — Ромашина Ольга Викторовна"
prs.core_properties.subject = "Профессиональная презентация сотрудника СФР"
prs.core_properties.author = "Ромашина Ольга Викторовна"
prs.core_properties.keywords = "СФР, социальная поддержка, семьи с детьми"
prs.save(OUT)
print(OUT)
