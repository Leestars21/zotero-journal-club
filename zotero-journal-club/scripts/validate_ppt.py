#!/usr/bin/env python3
"""Check PPTX structure; visual and scientific verification remain manual."""
import argparse
import json
from itertools import combinations
from pathlib import Path
from zipfile import ZipFile
from pptx import Presentation


def inspect(path):
    with ZipFile(path) as package:
        bad = package.testzip()
    prs = Presentation(str(path))
    issues = []
    warnings = []
    text_objects = 0
    pictures = 0
    notes_count = 0
    for i, slide in enumerate(prs.slides, 1):
        if not slide.has_notes_slide or not slide.notes_slide.notes_text_frame.text.strip():
            issues.append(f'slide {i}: missing speaker notes')
        else:
            notes_count += 1
        text_shapes = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip():
                text_objects += 1
                text_shapes.append(shape)
            if shape.shape_type == 13:
                pictures += 1
                if min(shape.width, shape.height) < 4 * 914400:
                    warnings.append(f'slide {i}: image {shape.name!r} has a small displayed dimension; review labels at presentation size')
            if shape.left < 0 or shape.top < 0 or shape.left + shape.width > prs.slide_width or shape.top + shape.height > prs.slide_height:
                issues.append(f'slide {i}: object extends beyond slide')
        for first, second in combinations(text_shapes, 2):
            width = max(0, min(first.left + first.width, second.left + second.width) - max(first.left, second.left))
            height = max(0, min(first.top + first.height, second.top + second.height) - max(first.top, second.top))
            smaller = min(first.width * first.height, second.width * second.height)
            if smaller and width * height / smaller >= .15:
                warnings.append(f'slide {i}: text boxes {first.name!r} and {second.name!r} overlap; review visually')
    if bad:
        issues.append(f'corrupt zip member: {bad}')
    if not prs.slides:
        issues.append('no slides')
    if not text_objects:
        issues.append('no editable text')
    return {'slide_count': len(prs.slides), 'editable_text_objects': text_objects,
            'slides_with_notes': notes_count, 'picture_objects': pictures,
            'issues': issues, 'warnings': warnings,
            'not_checked': ['scientific accuracy', 'all overlaps', 'text clipping', 'visual quality', 'source correctness']}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pptx', type=Path)
    args = parser.parse_args()
    report = inspect(args.pptx)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    raise SystemExit(bool(report['issues']))
