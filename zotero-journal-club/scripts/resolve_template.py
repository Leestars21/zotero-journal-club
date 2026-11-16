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


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--template')
    parser.add_argument('--config', default=os.environ.get('ZOTERO_JOURNAL_CLUB_CONFIG'))
    args = parser.parse_args()
    print(json.dumps(resolve(args.template, args.config), ensure_ascii=False, indent=2))
