#!/usr/bin/env python3
"""Build editable slides from a manually authored, source-grounded JSON plan."""
import argparse
import json
import re
from pathlib import Path
from pptx import Presentation
from pptx.dml.color import RGBColor
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
    if not isinstance(value, list) or len(value) != 4 or not all(isinstance(x, (int, float)) for x in value):
        raise ValueError(f'{name} must contain four relative coordinates')
    x, y, width, height = value
    if x < 0 or y < 0 or width <= 0 or height <= 0 or x + width > 1 or y + height > 1:
        raise ValueError(f'{name} must be inside the slide')
    return value


def build(plan, output, template=None):
    if template and Path(template).resolve() == Path(output).resolve():
        raise ValueError('Output must not overwrite the template')
    if Path(output).exists():
        raise FileExistsError(f'Refusing to overwrite existing output: {output}')
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
        def box(x, y, width, height, value, size, bold=False, color=None, font_name='Microsoft YaHei'):
            shape = slide.shapes.add_textbox(int(x*w), int(y*h), int(width*w), int(height*h))
            tf = shape.text_frame
            tf.word_wrap = True
            tf.text = value
            for p in tf.paragraphs:
                p.font.name = font_name
                p.font.size = Pt(size)
                p.font.bold = bold
                if color:
                    p.font.color.rgb = RGBColor.from_string(color.lstrip('#'))
            return shape
        image = entry.get('image')
        image_area = (.53*w, .22*h, .40*w, .62*h)
        regions = entry.get('regions')
        if template and regions is not None:
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
            if image and entry.get('image_focus'):
                if bullets:
                    box(.06, .17, .88, .08, '\n'.join(bullets), 18)
                    image_area = (.06*w, .27*h, .88*w, .62*h)
                else:
                    image_area = (.06*w, .17*h, .88*w, .72*h)
            else:
                box(.06, .2, .44 if image else .86, .65, '\n'.join(bullets), 20)
        if image:
            image_path = Path(image)
            if not image_path.is_absolute():
                image_path = Path(plan['_base']) / image_path
            from PIL import Image
            with Image.open(image_path) as source:
                iw, ih = source.size
            x, y, available_width, available_height = image_area
            scale = min(available_width/iw, available_height/ih)
            width, height = iw*scale, ih*scale
            slide.shapes.add_picture(str(image_path), int(x + (available_width-width)/2),
                                     int(y + (available_height-height)/2), width=int(width), height=int(height))
        box(.05, .91, .9, .06, entry['source'], 10)
        slide.notes_slide.notes_text_frame.text = entry['notes']
    prs.save(str(output))


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('plan', type=Path, help='JSON with slides: title, bullets, source, notes, optional image/layout')
    p.add_argument('output', type=Path)
    p.add_argument('--template', type=Path)
    args = p.parse_args()
    plan = json.loads(args.plan.read_text(encoding='utf-8'))
    plan['_base'] = str(args.plan.parent)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    build(plan, args.output, args.template)


if __name__ == '__main__':
    main()
