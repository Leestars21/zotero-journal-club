#!/usr/bin/env python3
"""Read-only discovery of local capabilities; MCP availability is checked by the agent."""
import argparse
import configparser
import importlib.util
import json
import os
import re
import shutil
from pathlib import Path


def profile_data_dirs(profile_root):
    ini = profile_root / 'profiles.ini'
    if not ini.is_file():
        return []
    config = configparser.ConfigParser(interpolation=None)
    config.read(ini, encoding='utf-8-sig')
    result = []
    for section in config.sections():
        if not section.startswith('Profile'):
            continue
        raw = config.get(section, 'Path', fallback='')
        if not raw:
            continue
        directory = (profile_root / raw) if config.get(section, 'IsRelative', fallback='1') == '1' else Path(raw)
        prefs = directory / 'prefs.js'
        if not prefs.is_file():
            continue
        content = prefs.read_text(encoding='utf-8', errors='replace')
        match = re.search(r'user_pref\("extensions\.zotero\.dataDir",\s*("(?:\\.|[^"\\])*")\s*\)', content)
        if match:
            result.append(Path(json.loads(match.group(1))).expanduser())
    return result


def inspect(data_dir=None, mineru_root=None, profile_root=None, app_dir=None):
    home = Path.home()
    if profile_root is None and os.environ.get('APPDATA'):
        profile_root = Path(os.environ['APPDATA']) / 'Zotero' / 'Zotero'
    candidates = []
    if data_dir:
        candidates.append(('explicit', Path(data_dir).expanduser()))
    elif os.environ.get('ZOTERO_DATA_DIR'):
        candidates.append(('environment', Path(os.environ['ZOTERO_DATA_DIR']).expanduser()))
    else:
        if profile_root:
            candidates.extend(('profile', p) for p in profile_data_dirs(Path(profile_root)))
        candidates.append(('default_candidate', home / 'Zotero'))
    dirs = [{'origin': origin, 'path': str(p), 'exists': p.is_dir(),
             'database_present': (p / 'zotero.sqlite').is_file(),
             'storage_present': (p / 'storage').is_dir()} for origin, p in candidates]
    roots = [Path(mineru_root).expanduser()] if mineru_root else [p / 'llm-for-zotero-mineru' for _, p in candidates]
    modules = {name: importlib.util.find_spec(name) is not None for name in ('pptx', 'PIL', 'fitz')}
    powerpoint = shutil.which('POWERPNT.EXE') or shutil.which('powerpoint')
    if not powerpoint and os.name == 'nt':
        base = Path(os.environ.get('ProgramFiles', 'C:/Program Files'))
        candidate = base / 'Microsoft Office' / 'root' / 'Office16' / 'POWERPNT.EXE'
        if candidate.is_file():
            powerpoint = str(candidate)
    zotero_dir = Path(app_dir).expanduser() if app_dir else None
    if zotero_dir is None:
        executable = shutil.which('zotero') or shutil.which('zotero.exe')
        if executable:
            zotero_dir = Path(executable).parent
        elif os.name == 'nt':
            zotero_dir = Path(os.environ.get('ProgramFiles', 'C:/Program Files')) / 'Zotero'
    zotero_version = None
    if zotero_dir:
        for ini in (zotero_dir / 'app' / 'application.ini', zotero_dir / 'application.ini'):
            if ini.is_file():
                metadata = configparser.ConfigParser(interpolation=None)
                metadata.read(ini, encoding='utf-8-sig')
                zotero_version = metadata.get('App', 'Version', fallback=None)
                break
    return {'zotero_data_candidates': dirs,
            'mineru_roots': [{'path': str(p), 'exists': p.is_dir()} for p in roots],
            'python_modules': modules, 'powerpoint': powerpoint, 'zotero_version': zotero_version,
            'mcp_status': 'discover from session tools',
            'note': 'Candidates are not proof of the active Zotero library; no database was opened.'}


if __name__ == '__main__':
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--data-dir')
    p.add_argument('--mineru-root')
    p.add_argument('--profile-root', type=Path)
    p.add_argument('--app-dir', type=Path)
    args = p.parse_args()
    print(json.dumps(inspect(args.data_dir, args.mineru_root, args.profile_root, args.app_dir), ensure_ascii=False, indent=2))
