#!/usr/bin/env python3
"""Inspect one local PDF without modifying it or using a MinerU cache."""
import argparse
import json
from pathlib import Path
import fitz


def inspect(pdf, text_dir=None):
    doc = fitz.open(pdf)
    if doc.needs_pass:
        raise ValueError('PDF requires a password')
    if text_dir is not None:
        text_dir = Path(text_dir)
        text_dir.mkdir(parents=True, exist_ok=True)
        existing = [text_dir / f'page-{index + 1:04d}.txt' for index in range(len(doc))
                    if (text_dir / f'page-{index + 1:04d}.txt').exists()]
        if existing:
            raise FileExistsError(f'Refusing to overwrite existing text: {existing[0]}')
    pages = []
    for index, page in enumerate(doc):
        text = page.get_text('text')
        target = None
        if text_dir is not None:
            target = text_dir / f'page-{index + 1:04d}.txt'
            target.write_text(text, encoding='utf-8')
        pages.append({'pdf_page': index + 1, 'page_index': index, 'text_chars': len(text.strip()),
                      'text_path': str(target) if target else None, 'image_count': len(page.get_images())})
    return {'pdf_path': str(Path(pdf).resolve()), 'paper_metadata': doc.metadata,
            'page_count': len(doc), 'pages': pages,
            'warning': 'Extracted text may contain OCR or reading-order errors; inspect original pages and figures.'}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf', type=Path)
    parser.add_argument('--text-dir', type=Path, help='Optional directory for per-page text files')
    args = parser.parse_args()
    print(json.dumps(inspect(args.pdf, args.text_dir), ensure_ascii=False, indent=2))
