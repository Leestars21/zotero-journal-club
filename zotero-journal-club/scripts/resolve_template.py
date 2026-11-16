#!/usr/bin/env python3
"""Resolve a PowerPoint template: explicit path, config default, or none."""
import argparse
import json
import os
from pathlib import Path


def resolve(explicit=None, config_path=None):
    if explicit:
        selected, origin = Path(explicit).expanduser(), 'explicit'
    else:
        config = Path(config_path).expanduser() if config_path else Path.home() / '.codex' / 'zotero-journal-club.json'
        if not config.is_file():
            return {'template': None, 'origin': 'none', 'config': str(config)}
        values = json.loads(config.read_text(encoding='utf-8'))
        configured = values.get('default_template')
        if not configured:
            return {'template': None, 'origin': 'none', 'config': str(config)}
        selected = Path(configured).expanduser()
        if not selected.is_absolute():
            selected = config.parent / selected
        origin = 'config'
    if selected.suffix.lower() != '.pptx' or not selected.is_file():
        raise ValueError(f'{origin} template is not an existing PPTX: {selected}')
    return {'template': str(selected.resolve()), 'origin': origin}


def resolve_reference(explicit=None, config_path=None):
    """A visual benchmark is separate from the template used to build the deck."""
    config = Path(config_path).expanduser() if config_path else Path.home() / '.codex' / 'zotero-journal-club.json'
    if explicit:
        selected, origin = Path(explicit).expanduser(), 'explicit'
    else:
        values = json.loads(config.read_text(encoding='utf-8')) if config.is_file() else {}
        configured = values.get('quality_reference')
        if not configured:
            return {'quality_reference': None, 'reference_origin': 'none'}
        selected, origin = Path(configured).expanduser(), 'config'
        if not selected.is_absolute():
            selected = config.parent / selected
    if selected.suffix.lower() != '.pptx' or not selected.is_file():
        raise ValueError(f'{origin} quality reference is not an existing PPTX: {selected}')
    return {'quality_reference': str(selected.resolve()), 'reference_origin': origin}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--template')
    parser.add_argument('--quality-reference', help='PPTX used as a visual quality benchmark, independently of --template')
    parser.add_argument('--config', default=os.environ.get('ZOTERO_JOURNAL_CLUB_CONFIG'))
    args = parser.parse_args()
    result = resolve(args.template, args.config)
    result.update(resolve_reference(args.quality_reference, args.config))
    print(json.dumps(result, ensure_ascii=False, indent=2))
