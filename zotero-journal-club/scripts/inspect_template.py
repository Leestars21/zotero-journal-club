#!/usr/bin/env python3
"""List the dimensions, masters, layouts and editable placeholders in a PPTX."""
import argparse
import json
from pathlib import Path
from pptx import Presentation


def inspect(path):
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
    return {'width_emu': prs.slide_width, 'height_emu': prs.slide_height,
            'master_count': len(prs.slide_masters), 'layouts': layouts, 'sample_slides': slides}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('template', type=Path)
    args = parser.parse_args()
    print(json.dumps(inspect(args.template), ensure_ascii=False, indent=2))
