#!/usr/bin/env python3
"""Crop explicitly located original PDF figures locally and retain their provenance."""
import argparse
import hashlib
import json
import re
from math import isfinite
from pathlib import Path
import fitz


def extract(pdf, regions, output_dir, dpi=220):
    pdf, output_dir = Path(pdf), Path(output_dir)
    if not isinstance(dpi, (int, float)) or isinstance(dpi, bool) or not isfinite(dpi) or not 72 <= dpi <= 600:
        raise ValueError('dpi must be between 72 and 600')
    if not isinstance(regions, list) or not regions:
        raise ValueError('Supply a nonempty list of explicitly inspected PDF regions')
    with fitz.open(pdf) as doc:
        if doc.needs_pass:
            raise ValueError('PDF requires a password')
        jobs, ids = [], set()
        for entry in regions:
            if not isinstance(entry, dict):
                raise ValueError('Each region must be a JSON object')
            identifier, page = entry.get('id'), entry.get('pdf_page')
            if not isinstance(identifier, str) or not re.fullmatch(r'[A-Za-z0-9_-]+', identifier) or identifier in ids:
                raise ValueError('Every region needs a unique id containing letters, numbers, hyphens or underscores')
            ids.add(identifier)
            if not isinstance(page, int) or isinstance(page, bool) or not 1 <= page <= len(doc):
                raise ValueError(f'{identifier}: pdf_page must be a one-based page within this PDF')
            bbox = entry.get('bbox')
            if not isinstance(bbox, list) or len(bbox) != 4 or not all(
                    isinstance(value, (int, float)) and not isinstance(value, bool) and isfinite(value) for value in bbox):
                raise ValueError(f'{identifier}: bbox needs four finite PDF coordinates in points')
            clip = fitz.Rect(bbox)
            if clip.is_empty or clip.is_infinite or not doc[page - 1].rect.contains(clip):
                raise ValueError(f'{identifier}: crop must remain within its observed PDF page')
            target = output_dir / f'{identifier}.png'
            jobs.append((entry, clip, target))
        manifest_path = output_dir / 'figure_manifest.json'
        for target in [manifest_path, *(job[2] for job in jobs)]:
            if target.exists():
                raise FileExistsError(f'Refusing to overwrite existing figure output: {target}')
        output_dir.mkdir(parents=True, exist_ok=True)
        sources = []
        for entry, clip, target in jobs:
            pixmap = doc[entry['pdf_page'] - 1].get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72),
                                                        clip=clip, alpha=False)
            pixmap.save(target)
            sources.append({'id': entry['id'], 'kind': entry.get('kind', 'figure'), 'origin': 'paper',
                            'path': str(target.resolve()), 'label': entry.get('label', entry['id']),
                            'caption': entry.get('caption', ''),
                            'locator': {'pdf_page': entry['pdf_page'], 'figure': entry.get('figure'),
                                        'panel': entry.get('panel'), 'table': entry.get('table'),
                                        'equation': entry.get('equation')},
                            'bbox_pdf_points': list(clip), 'dpi': dpi,
                            'pixels': [pixmap.width, pixmap.height], 'verification': 'unchecked'})
        result = {'pdf_source': {'path': str(pdf.resolve()), 'page_count': len(doc),
                                 'sha256': hashlib.sha256(pdf.read_bytes()).hexdigest()},
                  'sources': sources,
                  'warning': 'Coordinates were supplied by the caller; visually verify labels, captions and crop completeness against the PDF.'}
        manifest_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf', type=Path)
    parser.add_argument('regions_json', type=Path, help='JSON list with id, pdf_page and bbox in PDF points')
    parser.add_argument('output_dir', type=Path)
    parser.add_argument('--dpi', type=int, default=220)
    args = parser.parse_args()
    regions = json.loads(args.regions_json.read_text(encoding='utf-8'))
    print(json.dumps(extract(args.pdf, regions, args.output_dir, args.dpi), ensure_ascii=False, indent=2))
