"""Editable academic layouts with a shared hierarchy, panels and evidence captions."""
import re
from math import ceil
from pathlib import Path
from unicodedata import east_asian_width

from PIL import Image
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt


BLUE = '003F88'
RED = 'C00000'
INK = '202C3C'
MUTED = '5A6E83'
PALE = 'F7FAFD'
LINE = 'A8C3DF'
FONT = 'Microsoft YaHei'


def box(slide, x, y, width, height, text, size=20, color=INK, bold=False,
        align=PP_ALIGN.LEFT, highlights=(), name=None):
    """Make editable text; reject an obviously overfull authored box."""
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(width), Inches(height))
    if name:
        shape.name = name
    frame = shape.text_frame
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = Inches(.02)
    frame.margin_top = frame.margin_bottom = Inches(.025)
    frame.vertical_anchor = MSO_ANCHOR.TOP
    paragraphs = str(text).split('\n')
    patterns = [re.escape(value) for value in highlights if isinstance(value, str) and value]
    pattern = '(' + '|'.join(sorted(patterns, key=len, reverse=True)) + ')' if patterns else None
    estimate = 0
    for index, value in enumerate(paragraphs):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.alignment = align
        paragraph.font.name = FONT
        paragraph.font.size = Pt(size)
        paragraph.font.bold = bold
        paragraph.font.color.rgb = RGBColor.from_string(color)
        paragraph.space_after = Pt(4)
        paragraph.line_spacing = 1.12
        for piece in re.split(pattern, value) if pattern else [value]:
            run = paragraph.add_run()
            run.text = piece
            run.font.name = FONT
            run.font.size = Pt(size)
            run.font.bold = bold or piece in highlights
            run.font.color.rgb = RGBColor.from_string(RED if piece in highlights else color)
        units = sum(1 if east_asian_width(char) in ('W', 'F') else .5 for char in value)
        estimate += max(1, ceil(units * size / max((width - .04) * 72, 1))) * size * 1.12 + 4
    if estimate > (height - .05) * 72 * 1.12:
        raise ValueError(f'Academic layout text is too dense: {str(text)[:70]}; shorten it or split the slide')
    return shape


def rectangle(slide, x, y, width, height, fill=None, stroke=None, rounded=False, name=None):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(width), Inches(height))
    if name:
        shape.name = name
    if rounded:
        shape.adjustments[0] = .07
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        shape.fill.background()
    if stroke:
        shape.line.color.rgb = RGBColor.from_string(stroke)
        shape.line.width = Pt(.8)
    else:
        shape.line.fill.background()
    shape._element.spPr.append(OxmlElement('a:effectLst'))
    return shape


def picture(slide, plan, value, area, name='journal-club-evidence-image'):
    path = Path(value)
    if not path.is_absolute():
        path = Path(plan.get('_base', '.')) / path
    x, y, width, height = area
    with Image.open(path) as source:
        iw, ih = source.size
    scale = min(width / iw, height / ih)
    pw, ph = iw * scale, ih * scale
    result = slide.shapes.add_picture(str(path), Inches(x + (width - pw) / 2),
                                     Inches(y + (height - ph) / 2), width=Inches(pw), height=Inches(ph))
    result.name = name
    return result


def footer(slide, entry, number, total):
    rectangle(slide, 0, 7.16, 13.333, .34, BLUE)
    box(slide, .3, 7.19, 11.65, .24, entry['source'], 9.5, 'FFFFFF',
        name='journal-club-source')
    box(slide, 12.05, 7.18, .94, .26, f'{number}/{total}', 10.5, 'FFFFFF',
        True, PP_ALIGN.RIGHT, name='journal-club-page-number')


def heading(slide, entry, design):
    section = entry.get('section', '文献阅读')
    box(slide, .55, .27, 9.15, .59, section, 28, BLUE, True, name='journal-club-section')
    box(slide, 10.0, .43, 2.76, .32, design.get('label', '组会文献汇报'), 12, BLUE,
        align=PP_ALIGN.RIGHT)
    rectangle(slide, 0, 1.01, 13.333, .032, BLUE)
    rectangle(slide, .55, 1.27, .20, .20, stroke=INK)
    box(slide, .88, 1.14, 11.93, .62, entry['title'], 24, INK, True,
        highlights=entry.get('emphasis', []), name='journal-club-page-title')


def conclusion(slide, entry):
    takeaway = entry.get('takeaway')
    if takeaway:
        box(slide, .62, 6.65, 12.1, .42, takeaway, 17, BLUE, True, PP_ALIGN.CENTER,
            entry.get('emphasis', []), name='journal-club-takeaway')


def text_lines(slide, content, area, size=20):
    values = content.get('bullets', [])
    if not isinstance(values, list) or any(not isinstance(value, str) for value in values):
        raise ValueError('Academic bullets must be a list of strings')
    if not values:
        return
    x, y, width, height = area
    gap = .10
    row_height = (height - (len(values) - 1) * gap) / len(values)
    for index, value in enumerate(values):
        rectangle(slide, x, y + index * (row_height + gap) + .13, .065, .065, BLUE)
        box(slide, x + .18, y + index * (row_height + gap), width - .18, row_height,
            value, size, highlights=content.get('emphasis', []))


def content_images(slide, content, area, plan):
    images = content.get('images')
    if images is None:
        images = [{'path': content['image'], 'caption': content.get('caption', '')}]
    if not isinstance(images, list) or not 1 <= len(images) <= 4:
        raise ValueError('Academic images must contain one to four objects')
    columns = content.get('image_columns', min(2, len(images)))
    if columns not in (1, 2):
        raise ValueError('image_columns must be 1 or 2')
    rows = ceil(len(images) / columns)
    x, y, width, height = area
    gap = .22
    cw = (width - (columns - 1) * gap) / columns
    ch = (height - (rows - 1) * gap) / rows
    for index, image in enumerate(images):
        if not isinstance(image, dict) or not image.get('path'):
            raise ValueError('Every image needs a path')
        ix, iy = x + (index % columns) * (cw + gap), y + (index // columns) * (ch + gap)
        caption = image.get('caption')
        caption_height = .48 if caption else 0
        picture(slide, plan, image['path'], (ix, iy, cw, ch - caption_height),
                f'journal-club-image-{1 + sum(shape.shape_type == 13 for shape in slide.shapes)}')
        if caption:
            box(slide, ix, iy + ch - caption_height + .02, cw, caption_height - .02,
                caption, 12, BLUE, True, PP_ALIGN.CENTER)


def native_flow(slide, data, area):
    steps = data.get('steps') if isinstance(data, dict) else None
    if not isinstance(steps, list) or not 2 <= len(steps) <= 6:
        raise ValueError('flow needs two to six steps')
    x, y, width, height = area
    columns = len(steps) if len(steps) <= 4 else 3
    rows = ceil(len(steps) / columns)
    gap = .27
    cw = (width - gap * (columns - 1)) / columns
    caption = data.get('caption')
    usable = height - (.65 if caption else .1)
    ch = min(3.25 if width > 9 else 2.75, (usable - .3 * (rows - 1)) / rows)
    for index, value in enumerate(steps):
        step = {'title': value} if isinstance(value, str) else value
        if not isinstance(step, dict) or not step.get('title'):
            raise ValueError('Every flow step needs a title')
        cx, cy = x + (index % columns) * (cw + gap), y + (index // columns) * (ch + .3)
        rectangle(slide, cx, cy, cw, ch, PALE, LINE, True)
        if ch < 2.4:
            box(slide, cx + .12, cy + .12, .44, .45, f'{index + 1:02d}', 18, BLUE, True)
            box(slide, cx + .65, cy + .12, cw - .79, .64, step['title'],
                18, BLUE, True, name=f'journal-club-flow-step-{index + 1}')
            if step.get('detail'):
                box(slide, cx + .15, cy + .85, cw - .3, ch - .96, step['detail'], 16)
        else:
            box(slide, cx + .15, cy + .15, cw - .3, .46, f'{index + 1:02d}', 22, BLUE, True)
            box(slide, cx + .15, cy + .74, cw - .3, .73, step['title'],
                18 if cw < 3.1 else 20, BLUE, True, name=f'journal-club-flow-step-{index + 1}')
            if step.get('detail'):
                box(slide, cx + .15, cy + 1.55, cw - .3, ch - 1.70, step['detail'],
                    17 if cw < 3.1 else 18)
        if index % columns < columns - 1 and index < len(steps) - 1:
            arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(cx + cw + .04),
                                          Inches(cy + ch * .48), Inches(gap - .08), Inches(.25))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = RGBColor.from_string('8CAAC9')
            arrow.line.fill.background()
    if caption:
        used = rows * ch + (rows - 1) * .3
        box(slide, x + .03, y + used + .20, width - .06, .43, caption, 14, BLUE,
            True, PP_ALIGN.CENTER)


def native_table(slide, data, area):
    headers, rows = data.get('headers'), data.get('rows')
    if not isinstance(headers, list) or not headers or not isinstance(rows, list) or not rows:
        raise ValueError('table needs nonempty headers and rows')
    if any(not isinstance(row, list) or len(row) != len(headers) for row in rows):
        raise ValueError('Table rows must match headers')
    x, y, width, height = area
    size = data.get('font_pt', 18)
    caption = data.get('caption')
    caption_height = .43 if caption else 0
    table_height = min(height - caption_height, .58 * (len(rows) + 1))
    cell_width = (width / len(headers) - .18) * 72
    cell_height = (table_height / (len(rows) + 1) - .10) * 72
    for values in [headers, *rows]:
        for value in values:
            units = sum(1 if east_asian_width(char) in ('W', 'F') else .5 for char in str(value))
            if cell_width <= 0 or max(1, ceil(units * size / cell_width)) * size * 1.12 > cell_height * 1.12:
                raise ValueError('Academic table cells are too dense; split rows or use a wider, inspected layout')
    shape = slide.shapes.add_table(len(rows) + 1, len(headers), Inches(x), Inches(y),
                                   Inches(width), Inches(table_height))
    shape.name = 'journal-club-table'
    for row_index, values in enumerate([headers, *rows]):
        for column_index, value in enumerate(values):
            cell = shape.table.cell(row_index, column_index)
            cell.text = str(value)
            cell.margin_left = cell.margin_right = Inches(.09)
            cell.margin_top = cell.margin_bottom = Inches(.05)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(
                BLUE if row_index == 0 else ('EEF3F8' if row_index % 2 else 'FFFFFF'))
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.name = FONT
                paragraph.font.size = Pt(size)
                paragraph.font.bold = row_index == 0
                paragraph.font.color.rgb = RGBColor.from_string('FFFFFF' if row_index == 0 else INK)
    if caption:
        box(slide, x + .03, y + table_height + .08, width - .06, caption_height - .05,
            caption, 12, MUTED)


def native_equation(slide, data, area):
    if not isinstance(data, dict) or not data.get('text'):
        raise ValueError('equation needs editable text')
    x, y, width, height = area
    variables = data.get('variables', [])
    if width >= 9 and variables:
        left_width = width * .53
        rectangle(slide, x, y + .35, left_width, height - .75, PALE, LINE, True)
        box(slide, x + .20, y + .54, left_width - .4, .45, '模型表达', 19, BLUE, True)
        box(slide, x + .20, y + 1.30, left_width - .4, height - 2.05,
            data['text'], data.get('font_pt', 26), name='journal-club-equation')
        rx, rw = x + left_width + .4, width - left_width - .4
        rh = (height - .2) / len(variables)
        for index, variable in enumerate(variables):
            symbol = variable['symbol'] if isinstance(variable, dict) else str(variable)
            meaning = variable['meaning'] if isinstance(variable, dict) else ''
            vy = y + .10 + index * rh
            box(slide, rx, vy, rw, .40, symbol, 19, BLUE, True)
            if meaning:
                box(slide, rx, vy + .43, rw, rh - .47, meaning, 17)
        return
    box(slide, x + .05, y + .2, width - .1, height * .36, data['text'],
        data.get('font_pt', 26), name='journal-club-equation')
    lines = [f'{value["symbol"]}：{value["meaning"]}' if isinstance(value, dict) else str(value)
             for value in variables]
    if lines:
        box(slide, x + .05, y + height * .45, width - .1, height * .53, '\n'.join(lines), 18)


def summary_rows(slide, entry, area):
    values = entry['summary_rows']
    if not isinstance(values, list) or not 2 <= len(values) <= 4:
        raise ValueError('summary_rows needs two to four titled text objects')
    x, y, width, height = area
    row_height = height / len(values)
    rectangle(slide, x + 3.70, y + .12, .014, height - .24, LINE)
    for index, value in enumerate(values):
        if not isinstance(value, dict) or not value.get('title') or not isinstance(value.get('text'), str):
            raise ValueError('Every summary row needs title and text')
        ry = y + index * row_height
        box(slide, x + .10, ry + .18, .78, .69, f'{index + 1:02d}', 28, BLUE, True)
        box(slide, x + 1.05, ry + .23, 2.43, row_height - .40, value['title'],
            20, BLUE, True)
        box(slide, x + 4.05, ry + .22, width - 4.20, row_height - .4,
            value['text'], 19, highlights=value.get('emphasis', entry.get('emphasis', [])))
        if index < len(values) - 1:
            rectangle(slide, x + .08, ry + row_height, width - .16, .014, LINE)


def content(slide, data, area, plan):
    x, y, width, height = area
    types = [key for key in ('image', 'images', 'table', 'flow', 'equation', 'columns') if key in data]
    if len(types) > 1:
        raise ValueError('Choose one academic content type per region')
    if 'columns' in data:
        columns = data['columns']
        if not isinstance(columns, list) or not 1 <= len(columns) <= 2:
            raise ValueError('columns needs one or two lists')
        cw = (width - .35 * (len(columns) - 1)) / len(columns)
        titles = data.get('column_titles', [])
        for index, values in enumerate(columns):
            cx = x + index * (cw + .35)
            if titles:
                box(slide, cx, y, cw, .46, titles[index], 19, BLUE, True)
            if len(columns) == 2 and index == 0:
                rectangle(slide, cx + cw + .165, y, .012, height, LINE)
            text_lines(slide, {'bullets': values, 'emphasis': data.get('emphasis', [])},
                       (cx, y + (.60 if titles else 0), cw, height - (.60 if titles else 0)), 20)
        return
    bullets = data.get('bullets', [])
    if types and bullets:
        lines = sum(max(1, ceil(sum(1 if east_asian_width(char) in ('W', 'F') else .5
                                   for char in value) * 18 / max((width - .08) * 72, 1))) for value in bullets)
        intro = (lines * 18 * 1.12 + len(bullets) * 4) / 72 + .09
        if intro > height * .38:
            raise ValueError('Too much introduction above evidence; move explanation to notes or a separate panel')
        box(slide, x + .02, y, width - .04, intro, '\n'.join(bullets),
            18, highlights=data.get('emphasis', []))
        y += intro + .12
        height -= intro + .12
    if types and types[0] in ('image', 'images'):
        content_images(slide, data, (x, y, width, height), plan)
    elif 'flow' in data:
        native_flow(slide, data['flow'], (x, y, width, height))
    elif 'table' in data:
        native_table(slide, data['table'], (x, y, width, height))
    elif 'equation' in data:
        native_equation(slide, data['equation'], (x, y, width, height))
    else:
        text_lines(slide, data, (x + .07, y + .06, width - .14, height - .12),
                   18 if width < 4.5 else 20)


def panels(slide, entry, area, plan):
    values = entry['panels']
    if not isinstance(values, list) or not 1 <= len(values) <= 3:
        raise ValueError('Academic panels needs one to three objects')
    x, y, width, height = area
    gap = .26
    cw = (width - gap * (len(values) - 1)) / len(values)
    for index, value in enumerate(values):
        if not isinstance(value, dict) or not value.get('title'):
            raise ValueError('Every academic panel needs a title')
        cx = x + index * (cw + gap)
        rectangle(slide, cx, y, cw, height, PALE, LINE, True, f'journal-club-panel-{index + 1}')
        rectangle(slide, cx + .15, y + .13, cw - .3, .43, BLUE, rounded=True)
        box(slide, cx + .22, y + .15, cw - .44, .38, value['title'],
            17 if len(values) == 3 else 19, 'FFFFFF', True, PP_ALIGN.CENTER)
        content(slide, value, (cx + .17, y + .72, cw - .34, height - .87), plan)
    if entry.get('connect_panels'):
        for index in range(len(values) - 1):
            arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                          Inches(x + (index + 1) * cw + index * gap - .04),
                                          Inches(y + height * .48), Inches(gap + .08), Inches(.28))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = RGBColor.from_string('B8CDE3')
            arrow.line.fill.background()


def cover(slide, entry, design):
    rectangle(slide, 0, 0, 13.333, 1.67, BLUE)
    box(slide, .65, .58, 12.0, .45, design.get('label', '组会文献阅读汇报'),
        22, 'FFFFFF', True)
    rectangle(slide, .65, 1.25, .82, .048, 'FFFFFF')
    box(slide, .85, 2.32, 11.63, 1.48, entry['title'], 32, BLUE, True,
        PP_ALIGN.CENTER, name='journal-club-page-title')
    if entry.get('subtitle'):
        box(slide, 1.2, 3.9, 10.93, .66, entry['subtitle'], 18, MUTED,
            align=PP_ALIGN.CENTER)
    metadata = entry.get('metadata', [])
    if metadata:
        box(slide, 1.15, 4.78, 11.03, 1.52, '\n'.join(metadata), 16, BLUE,
            align=PP_ALIGN.CENTER)


def agenda(slide, entry):
    values = entry.get('agenda')
    if not isinstance(values, list) or not 2 <= len(values) <= 4:
        raise ValueError('agenda needs two to four section titles')
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(.86), Inches(2.23), Inches(3.5), Inches(3.5))
    circle.fill.solid()
    circle.fill.fore_color.rgb = RGBColor.from_string(BLUE)
    circle.line.fill.background()
    box(slide, 1.07, 3.37, 3.08, .78, entry['title'], 32, 'FFFFFF', True, PP_ALIGN.CENTER,
        name='journal-club-page-title')
    for index, value in enumerate(values):
        y = 2.05 + index * 1.0
        box(slide, 5.10, y, .7, .53, f'{index + 1:02d}', 24, BLUE, True)
        box(slide, 6.00, y, 6.50, .68, value, 24, BLUE, True)
        rectangle(slide, 6.0, y + .76, 6.43, .012, LINE)


def render(slide, entry, plan, number, total):
    design = plan['design']
    if design.get('preset') != 'academic-blue':
        raise ValueError('Unknown design preset; use academic-blue or omit design for the legacy helper')
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor.from_string('FFFFFF')
    page_type = entry.get('page_type', 'content')
    if page_type == 'cover':
        cover(slide, entry, design)
    elif page_type == 'agenda':
        box(slide, .55, .34, 11.9, .54, design.get('label', '组会文献汇报'), 24, BLUE, True)
        agenda(slide, entry)
    elif page_type in ('content', 'closing'):
        heading(slide, entry, design)
        area = (.55, 1.90, 12.23, 4.58 if entry.get('takeaway') else 4.98)
        if 'summary_rows' in entry:
            if any(key in entry for key in ('panels', 'bullets', 'columns', 'image', 'images', 'table', 'flow', 'equation')):
                raise ValueError('summary_rows cannot be combined with other slide-level content')
            summary_rows(slide, entry, area)
        elif 'panels' in entry:
            if any(key in entry for key in ('bullets', 'columns', 'image', 'images', 'table', 'flow', 'equation')):
                raise ValueError('panels cannot be combined with slide-level content; put that content inside a panel')
            panels(slide, entry, area, plan)
        else:
            content(slide, entry, area, plan)
        conclusion(slide, entry)
    else:
        raise ValueError('page_type must be cover, agenda, content or closing')
    footer(slide, entry, number, total)
