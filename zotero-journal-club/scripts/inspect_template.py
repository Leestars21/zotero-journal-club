#!/usr/bin/env python3
"""List the dimensions, masters, layouts and editable placeholders in a PPTX."""
import argparse
import json
from collections import Counter
from pathlib import Path
from pptx import Presentation


def all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == 6:
            yield from all_shapes(shape.shapes)


def design_summary(prs):
    fonts, sizes, colors = Counter(), Counter(), Counter()
    contents = []
    for number, slide in enumerate(prs.slides, 1):
        shapes = list(all_shapes(slide.shapes))
        contents.append({'slide': number,
                         'text_objects': sum(shape.has_text_frame and bool(shape.text.strip()) for shape in shapes),
                         'pictures': sum(shape.shape_type == 13 for shape in shapes),
                         'tables': sum(shape.has_table for shape in shapes)})
        for shape in shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    font = run.font
                    weight = len(run.text)
                    name = font.name or paragraph.font.name
                    size = font.size or paragraph.font.size
                    if name:
                        fonts[name] += weight
                    if size:
                        sizes[round(size.pt, 2)] += weight
                    if font.color.type and str(font.color.type).startswith('RGB'):
                        colors[str(font.color.rgb)] += weight
    return {'fonts_by_character_count': fonts.most_common(),
            'font_sizes_pt_by_character_count': sizes.most_common(),
            'explicit_text_colors_by_character_count': colors.most_common(),
            'slide_content_counts': contents,
            'limitation': 'Explicit styles only; theme inheritance and visual layout still require rendering.'}


def inspect(path, include_design=False):
    prs = Presentation(str(path))
    layouts = []
    for i, layout in enumerate(prs.slide_layouts):
        placeholders = []
        for shape in layout.placeholders:
            fmt = shape.placeholder_format
            placeholders.append({'idx': fmt.idx, 'type': str(fmt.type), 'name': shape.name,
                                 'left': shape.left, 'top': shape.top, 'width': shape.width, 'height': shape.height})
        layouts.append({'index': i, 'name': layout.name, 'placeholders': placeholders})
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        slides.append({'number': i, 'layout': slide.slide_layout.name,
                       'text': [s.text[:160] for s in slide.shapes if s.has_text_frame and s.text.strip()]})
    result = {'width_emu': prs.slide_width, 'height_emu': prs.slide_height,
              'master_count': len(prs.slide_masters), 'layouts': layouts, 'sample_slides': slides}
    if include_design:
        result['design_summary'] = design_summary(prs)
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('template', type=Path)
    parser.add_argument('--design', action='store_true', help='Include explicit fonts, sizes, colors and content counts')
    args = parser.parse_args()
    print(json.dumps(inspect(args.template, args.design), ensure_ascii=False, indent=2))
