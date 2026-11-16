#!/usr/bin/env python3
"""Build editable slides from a manually authored, source-grounded JSON plan."""
import argparse
import importlib.util
import json
import re
from math import ceil, isfinite
from pathlib import Path
from unicodedata import east_asian_width
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt


def update_total_slide_numbers(prs, count):
    """Update literal /N only when it follows a PowerPoint slide-number field."""
    for master in prs.slide_masters:
        for layer in [master, *master.slide_layouts]:
            for shape in layer.shapes:
                if not shape.has_text_frame:
                    continue
                for paragraph in shape.text_frame.paragraphs:
                    elements = list(paragraph._p)
                    for index, element in enumerate(elements):
                        if not element.tag.endswith('}fld') or element.get('type') != 'slidenum':
                            continue
                        for sibling in elements[index + 1:]:
                            if not sibling.tag.endswith('}r'):
                                continue
                            text_node = next((n for n in sibling if n.tag.endswith('}t')), None)
                            if text_node is not None:
                                if re.fullmatch(r'\s*/\s*\d+\s*', text_node.text or ''):
                                    text_node.text = f'/{count}'
                                break


def region(value, name):
    if not isinstance(value, list) or len(value) != 4 or not all(
            isinstance(x, (int, float)) and not isinstance(x, bool) and isfinite(x) for x in value):
        raise ValueError(f'{name} must contain four relative coordinates')
    x, y, width, height = value
    if x < 0 or y < 0 or width <= 0 or height <= 0 or x + width > 1 or y + height > 1:
        raise ValueError(f'{name} must be inside the slide')
    return value


def add_text(slide, area, value, size=20, bold=False, color=None, font_name='Microsoft YaHei'):
    shape = slide.shapes.add_textbox(*(int(value) for value in area))
    shape.text_frame.word_wrap = True
    shape.text_frame.text = value
    for paragraph in shape.text_frame.paragraphs:
        paragraph.font.name = font_name
        paragraph.font.size = Pt(size)
        paragraph.font.bold = bold
        if color:
            paragraph.font.color.rgb = RGBColor.from_string(color.lstrip('#'))
    return shape


def image_path(plan, value):
    path = Path(value)
    return path if path.is_absolute() else Path(plan.get('_base', '.')) / path


def add_image(slide, path, area):
    from PIL import Image
    with Image.open(path) as source:
        iw, ih = source.size
    x, y, available_width, available_height = area
    scale = min(available_width / iw, available_height / ih)
    width, height = iw * scale, ih * scale
    return slide.shapes.add_picture(str(path), int(x + (available_width - width) / 2),
                                    int(y + (available_height - height) / 2),
                                    width=int(width), height=int(height))


def render_rich_content(slide, entry, area, plan, font_name='Microsoft YaHei'):
    """Render one explicit research content type as native objects or source images."""
    x, y, width, height = area
    bullets = entry.get('bullets', [])
    if bullets:
        intro_height = height * .18
        add_text(slide, (x, y, width, intro_height), '\n'.join(bullets), 17, font_name=font_name)
        y += intro_height + height * .03
        height *= .79
    if 'images' in entry:
        images = entry['images']
        if not isinstance(images, list) or not 1 <= len(images) <= 4:
            raise ValueError('images must contain one to four source image objects')
        columns = entry.get('image_columns', min(2, len(images)))
        if columns not in (1, 2):
            raise ValueError('image_columns must be 1 or 2')
        rows = (len(images) + columns - 1) // columns
        gap = width * .025
        cell_width = (width - (columns - 1) * gap) / columns
        cell_height = (height - (rows - 1) * gap) / rows
        for index, item in enumerate(images):
            if not isinstance(item, dict) or not item.get('path'):
                raise ValueError('Each image needs a path; captions and source_ids are optional')
            ix = x + (index % columns) * (cell_width + gap)
            iy = y + (index // columns) * (cell_height + gap)
            caption = item.get('caption', '')
            caption_height = min(height * .11, cell_height * .2) if caption else 0
            picture = add_image(slide, image_path(plan, item['path']),
                                (ix, iy, cell_width, cell_height - caption_height))
            picture.name = f'journal-club-image-{index + 1}'
            if caption:
                add_text(slide, (ix, iy + cell_height - caption_height, cell_width, caption_height),
                         caption, 12, font_name=font_name)
    elif 'table' in entry:
        data = entry['table']
        if not isinstance(data, dict):
            raise ValueError('table must be an object with headers and rows')
        headers, rows = data.get('headers'), data.get('rows')
        if not isinstance(headers, list) or not headers or not isinstance(rows, list) or not rows:
            raise ValueError('table needs nonempty headers and rows')
        if any(not isinstance(row, list) or len(row) != len(headers) for row in rows):
            raise ValueError('Every table row must have the same number of cells as headers')
        caption = data.get('caption', '')
        caption_height = height * .1 if caption else 0
        shape = slide.shapes.add_table(len(rows) + 1, len(headers), int(x), int(y),
                                       int(width), int(height - caption_height))
        shape.name = 'journal-club-table'
        font_size = data.get('font_pt', 16)
        for row_index, values in enumerate([headers, *rows]):
            for column_index, value in enumerate(values):
                cell = shape.table.cell(row_index, column_index)
                cell.text = str(value)
                cell.margin_left = cell.margin_right = Inches(.08)
                cell.margin_top = cell.margin_bottom = Inches(.05)
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.name = font_name
                    paragraph.font.size = Pt(font_size)
                    paragraph.font.bold = row_index == 0
                    paragraph.font.color.rgb = RGBColor.from_string(data.get('text_color', '243447').lstrip('#'))
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string('E5EDF5' if row_index == 0 else 'FFFFFF')
        if caption:
            add_text(slide, (x, y + height - caption_height, width, caption_height), caption, 12,
                     font_name=font_name)
    elif 'flow' in entry:
        data = entry['flow']
        steps = data.get('steps') if isinstance(data, dict) else None
        if not isinstance(steps, list) or not 2 <= len(steps) <= 6:
            raise ValueError('flow needs two to six steps')
        columns = len(steps) if len(steps) <= 4 else 3
        rows = (len(steps) + columns - 1) // columns
        gap = width * .04
        cell_width = (width - (columns - 1) * gap) / columns
        cell_height = min(height * .65, (height - (rows - 1) * gap) / rows)
        for index, step in enumerate(steps):
            if isinstance(step, str):
                step = {'title': step}
            if not isinstance(step, dict) or not step.get('title'):
                raise ValueError('Each flow step needs a title')
            ix = x + (index % columns) * (cell_width + gap)
            iy = y + (index // columns) * (cell_height + gap)
            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, int(ix), int(iy),
                                           int(cell_width), int(cell_height))
            shape.name = f'journal-club-flow-step-{index + 1}'
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string('F0F4F8')
            shape.line.color.rgb = RGBColor.from_string('A4B4C5')
            shape.text_frame.word_wrap = True
            shape.text = f'{index + 1}. {step["title"]}'
            if step.get('detail'):
                shape.text_frame.add_paragraph().text = step['detail']
            for paragraph in shape.text_frame.paragraphs:
                paragraph.font.name = font_name
                paragraph.font.size = Pt(data.get('font_pt', 17))
                paragraph.font.color.rgb = RGBColor.from_string('243447')
            if index % columns < columns - 1 and index < len(steps) - 1:
                arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, int(ix + cell_width + gap * .15),
                                               int(iy + cell_height * .43), int(gap * .7), int(cell_height * .14))
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = RGBColor.from_string('8094A8')
                arrow.line.fill.background()
    elif 'equation' in entry:
        data = entry['equation']
        if not isinstance(data, dict) or not isinstance(data.get('text'), str) or not data['text'].strip():
            raise ValueError('equation needs editable plain text; use a source image for complex mathematics')
        shape = add_text(slide, (x, y, width, height * .34), data['text'], data.get('font_pt', 26),
                         font_name=font_name)
        shape.name = 'journal-club-equation'
        variables = data.get('variables', [])
        if not isinstance(variables, list):
            raise ValueError('equation variables must be a list')
        lines = [f'{value["symbol"]}：{value["meaning"]}' if isinstance(value, dict) else str(value)
                 for value in variables]
        if lines:
            add_text(slide, (x, y + height * .4, width, height * .6), '\n'.join(lines), 20,
                     font_name=font_name)


def speaker_notes(entry):
    notes = entry['notes']
    for field, heading in [('reading_questions', '阅读疑问'), ('takeaways', '可借鉴的方法与验证条件')]:
        values = entry.get(field, [])
        if not isinstance(values, list):
            raise ValueError(f'{field} must be a list')
        if values:
            notes += '\n\n' + heading + '\n' + '\n'.join(str(value) for value in values)
    reading = entry.get('figure_reading')
    if reading is not None:
        if not isinstance(reading, dict):
            raise ValueError('figure_reading must be an object')
        labels = {'comparison': '比较对象', 'conditions': '试验或计算条件', 'observation': '主要观察',
                  'claim': '支持的结论', 'scope': '适用范围'}
        notes += '\n\n读图提示\n' + '\n'.join(f'{label}：{reading[key]}' for key, label in labels.items() if reading.get(key))
    duration = entry.get('duration_seconds')
    if duration is not None:
        if isinstance(duration, bool) or not isinstance(duration, (int, float)) or not isfinite(duration) or duration <= 0:
            raise ValueError('duration_seconds must be a positive finite number')
        notes += f'\n\n建议时长：{duration:g} 秒'
    return notes


def build(plan, output, template=None):
    if template and Path(template).resolve() == Path(output).resolve():
        raise ValueError('Output must not overwrite the template')
    if Path(output).exists():
        raise FileExistsError(f'Refusing to overwrite existing output: {output}')
    academic = None
    if 'design' in plan:
        if not isinstance(plan['design'], dict) or plan['design'].get('preset') != 'academic-blue':
            raise ValueError('design must be an object with preset academic-blue')
        if template:
            raise ValueError('Use inspected template layouts with --template; do not draw an independent design over its artwork')
        spec = importlib.util.spec_from_file_location('journal_club_academic_design',
                                                     Path(__file__).with_name('academic_design.py'))
        academic = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(academic)
    prs = Presentation(str(template)) if template else Presentation()
    if template:
        update_total_slide_numbers(prs, len(plan['slides']))
    # Remove sample slides, retaining their master/theme/layout package.
    for slide_id in list(prs.slides._sldIdLst):
        prs.part.drop_rel(slide_id.rId)
        prs.slides._sldIdLst.remove(slide_id)
    if not template:
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    for entry in plan['slides']:
        if not str(entry.get('notes') or '').strip() or not str(entry.get('source') or '').strip():
            raise ValueError('Every slide needs substantive notes and a source label')
        idx = entry.get('layout', min(1, len(prs.slide_layouts) - 1))
        slide = prs.slides.add_slide(prs.slide_layouts[idx])
        w, h = prs.slide_width, prs.slide_height
        if academic:
            for shape in list(slide.placeholders):
                shape._element.getparent().remove(shape._element)
            academic.render(slide, entry, plan, len(prs.slides), len(plan['slides']))
            slide._element.cSld.set('name', str(entry.get('id', f'slide-{len(prs.slides):03d}')))
            slide.notes_slide.notes_text_frame.text = speaker_notes(entry)
            continue
        def box(x, y, width, height, value, size, bold=False, color=None, font_name='Microsoft YaHei'):
            return add_text(slide, (x*w, y*h, width*w, height*h), value, size, bold, color, font_name)
        image = entry.get('image')
        image_area = (.53*w, .22*h, .40*w, .62*h)
        regions = entry.get('regions')
        columns = entry.get('columns')
        if columns is not None:
            if not isinstance(columns, list) or len(columns) not in (1, 2) or not all(isinstance(column, list) for column in columns):
                raise ValueError('columns must be a list of one or two lists')
            if image:
                raise ValueError('columns and image cannot be combined')
        rich_types = [field for field in ('images', 'table', 'flow', 'equation') if field in entry]
        if rich_types:
            if len(rich_types) != 1 or image or columns is not None:
                raise ValueError('Choose one of images, table, flow or equation; do not combine it with image or columns')
            if template and regions is not None:
                box(*region(regions.get('title'), 'title region'), entry['title'], regions.get('title_font_pt', 27),
                    True, regions.get('title_color'), regions.get('title_font_name', 'Microsoft YaHei'))
                rx, ry, rw, rh = region(regions.get('body'), 'body region')
                area = (rx*w, ry*h, rw*w, rh*h)
                font_name = regions.get('body_font_name', 'Microsoft YaHei')
            elif template:
                title = slide.shapes.title
                bodies = [shape for shape in slide.placeholders if shape != title and shape.has_text_frame and
                          str(shape.placeholder_format.type).startswith(('BODY', 'OBJECT', 'SUBTITLE'))]
                if title is None or not bodies:
                    raise ValueError(f'layout {idx} lacks suitable title/body placeholders; provide inspected regions')
                title.text = entry['title']
                target = bodies[0]
                area = (target.left, target.top, target.width, target.height)
                for body in bodies:
                    body._element.getparent().remove(body._element)
                font_name = 'Microsoft YaHei'
            else:
                for shape in list(slide.placeholders):
                    shape._element.getparent().remove(shape._element)
                box(.05, .04, .9, .12, entry['title'], 27, True)
                area, font_name = (.06*w, .19*h, .88*w, .67*h), 'Microsoft YaHei'
            render_rich_content(slide, entry, area, plan, font_name)
        elif template and regions is not None:
            # Explicitly inspected custom layouts can have design elements but no placeholders.
            title_region = region(regions.get('title'), 'title region')
            body_region = region(regions.get('body'), 'body region')
            columns = entry.get('columns')
            if columns is not None and image:
                raise ValueError('columns and image cannot be combined')
            if columns is not None and (not isinstance(columns, list) or len(columns) != 2
                                        or not all(isinstance(column, list) for column in columns)):
                raise ValueError('regions columns must be a list of two lists')
            box(*title_region, entry['title'], regions.get('title_font_pt', 27), True,
                regions.get('title_color'), regions.get('title_font_name', 'Microsoft YaHei'))
            if columns:
                box(*body_region, '\n'.join(columns[0]), regions.get('body_font_pt', 20),
                    font_name=regions.get('body_font_name', 'Microsoft YaHei'))
                box(*region(regions.get('body2'), 'body2 region'), '\n'.join(columns[1]), regions.get('body_font_pt', 20),
                    font_name=regions.get('body_font_name', 'Microsoft YaHei'))
            else:
                box(*body_region, '\n'.join(entry.get('bullets', [])), regions.get('body_font_pt', 20),
                    font_name=regions.get('body_font_name', 'Microsoft YaHei'))
            if image:
                ix, iy, iw, ih = region(regions.get('image'), 'image region')
                image_area = (ix*w, iy*h, iw*w, ih*h)
        elif template:
            title = slide.shapes.title
            bodies = [shape for shape in slide.placeholders
                      if shape != title and shape.has_text_frame and
                      str(shape.placeholder_format.type).startswith(('BODY', 'OBJECT', 'SUBTITLE'))]
            columns = entry.get('columns')
            if columns is not None and (not isinstance(columns, list) or len(columns) not in (1, 2)
                                        or not all(isinstance(column, list) for column in columns)):
                raise ValueError('columns must be a list of one or two lists')
            if columns is not None and image:
                raise ValueError('columns and image cannot be combined')
            needed = len(columns) if columns is not None else (2 if image else 1)
            if title is None or len(bodies) < needed:
                raise ValueError(f'layout {idx} lacks suitable title/body placeholders; choose another layout or edit template manually')
            title.text = entry['title']
            if columns is not None:
                for shape, column in zip(bodies, columns):
                    shape.text = '\n'.join(column)
            else:
                bodies[0].text = '\n'.join(entry.get('bullets', []))
            if image:
                target = bodies[1]
                image_area = (target.left, target.top, target.width, target.height)
                target._element.getparent().remove(target._element)
            for unused in bodies[needed:]:
                unused._element.getparent().remove(unused._element)
        else:
            for shape in list(slide.placeholders):
                shape._element.getparent().remove(shape._element)
            box(.05, .04, .9, .12, entry['title'], 27, True)
            bullets = entry.get('bullets', [])
            if columns is not None:
                if len(columns) == 1:
                    box(.06, .2, .88, .65, '\n'.join(columns[0]), 20)
                else:
                    box(.06, .2, .42, .65, '\n'.join(columns[0]), 20)
                    box(.52, .2, .42, .65, '\n'.join(columns[1]), 20)
            elif image and entry.get('image_focus'):
                if bullets:
                    line_width_pt = .88*w / 12700 - 14.4
                    lines = sum(max(1, ceil(sum(1 if east_asian_width(char) in ('W', 'F') else .5
                                               for char in text) * 18 / line_width_pt)) for text in bullets)
                    text_height = max(.08, (lines * 18 * 1.2 + 14.4) * 12700 / h)
                    if text_height > .28:
                        raise ValueError('image_focus has too much text; move detailed explanation to notes or another slide')
                    box(.06, .17, .88, text_height, '\n'.join(bullets), 18)
                    image_top = .17 + text_height + .03
                    image_area = (.06*w, image_top*h, .88*w, (.89-image_top)*h)
                else:
                    image_area = (.06*w, .17*h, .88*w, .72*h)
            else:
                box(.06, .2, .44 if image else .86, .65, '\n'.join(bullets), 20)
        if image:
            add_image(slide, image_path(plan, image), image_area)
        source_shape = box(.05, .91, .9, .06, entry['source'], 10)
        source_shape.name = 'journal-club-source'
        slide._element.cSld.set('name', str(entry.get('id', f'slide-{len(prs.slides):03d}')))
        slide.notes_slide.notes_text_frame.text = speaker_notes(entry)
    prs.save(str(output))


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('plan', type=Path, help='Authored JSON slide plan; see references/slide-plan.md for research content types')
    p.add_argument('output', type=Path)
    p.add_argument('--template', type=Path)
    args = p.parse_args()
    plan = json.loads(args.plan.read_text(encoding='utf-8'))
    plan['_base'] = str(args.plan.parent)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    build(plan, args.output, args.template)


if __name__ == '__main__':
    main()
