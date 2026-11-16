#!/usr/bin/env python3
"""Read-only helpers for resolving a Zotero search and inspecting a MinerU cache.

The agent supplies JSON returned by the Zotero MCP; this script does not connect
to Zotero or write to its database. Paths in the result are not copied.
"""
import argparse
import json
import re
from pathlib import Path
from urllib.parse import unquote, urlsplit


def norm_title(value):
    return ' '.join(re.findall(r'[^\W_]+', (value or '').casefold(), re.UNICODE))


def norm_doi(value):
    value = (value or '').strip().casefold()
    return re.sub(r'^(?:https?://(?:dx\.)?doi\.org/|doi:\s*)', '', value).rstrip(' .')


def unwrap(data):
    if 'result' in data:
        data = data['result']
    return data


def resolve_candidates(results, *, title=None, doi=None, key=None, author=None, year=None, limited=False):
    """Exact identity match; ambiguous and incomplete searches never auto-select."""
    if not any((title, doi, key)):
        raise ValueError('supply a title, DOI or item key')
    found = []
    for item in results:
        fields = (item.get('metadata') or {}).get('fields', {})
        title_value = item.get('title') or fields.get('title', '')
        if title and norm_title(title) != norm_title(title_value):
            continue
        if doi and norm_doi(doi) != norm_doi(fields.get('DOI', item.get('DOI', ''))):
            continue
        if key and key.casefold() != str(item.get('itemKey', item.get('key', ''))).casefold():
            continue
        if author and author.casefold() not in str(item.get('firstCreator', '')).casefold():
            creators = (item.get('metadata') or {}).get('creators', [])
            if not any(author.casefold() in str(c.get('lastName', '')).casefold() for c in creators):
                continue
        if year and str(year) not in str(item.get('year') or fields.get('date', '')):
            continue
        found.append(item)
    status = 'incomplete_search' if limited else ('not_found' if not found else 'matched' if len(found) == 1 else 'ambiguous')
    return {'status': status, 'candidates': found, 'selected': found[0] if status == 'matched' else None}


def inspect_mineru(cache, *, attachment_id=None, attachment_key=None, parent_key=None):
    cache = Path(cache)
    source_path = cache / '_llm_source.json'
    identity = json.loads(source_path.read_text(encoding='utf-8')) if source_path.is_file() else None
    checks = [('attachmentId', attachment_id), ('attachmentKey', attachment_key), ('parentItemKey', parent_key)]
    conflicts = [name for name, expected in checks if expected is not None and identity and
                 name in identity and str(identity[name]).casefold() != str(expected).casefold()]
    evidence = [name for name, expected in checks if expected is not None and identity and
                name in identity and str(identity[name]).casefold() == str(expected).casefold()]
    status = 'conflict' if conflicts else 'matched' if evidence else 'unverified'
    markdown = sorted(cache.glob('*.md')) if cache.is_dir() else []
    images = []
    for md in markdown:
        content = md.read_text(encoding='utf-8')
        for ref in re.findall(r'!\[[^\]]*\]\(([^)]+)\)', content):
            ref = unquote(ref.strip().strip('<>'))
            parsed = urlsplit(ref)
            if parsed.scheme or parsed.netloc:
                images.append({'markdown': str(md), 'reference': ref, 'status': 'remote_not_loaded'})
                continue
            path = (md.parent / parsed.path).resolve()
            images.append({'markdown': str(md), 'reference': ref, 'path': str(path),
                           'status': 'present' if path.is_file() else 'missing'})
    structured = [str(p) for p in sorted(cache.glob('*.json')) if p.name != '_llm_source.json'] if cache.is_dir() else []
    content_list = cache / 'content_list.json'
    if content_list.is_file():
        data = json.loads(content_list.read_text(encoding='utf-8'))
        if isinstance(data, list):
            for block in data:
                if not isinstance(block, dict) or not isinstance(block.get('img_path'), str):
                    continue
                ref = block['img_path']
                path = (cache / ref).resolve()
                page = block.get('page_idx')
                images.append({'structured_source': str(content_list), 'reference': ref, 'path': str(path),
                               'status': 'present' if path.is_file() else 'missing',
                               'page_index': page, 'pdf_page': page + 1 if isinstance(page, int) and page >= 0 else None,
                               'caption': block.get('image_caption', block.get('table_caption', []))})
    return {'identity_status': status, 'matched_fields': evidence, 'conflicts': conflicts,
            'markdown': [str(p) for p in markdown], 'structured_files': structured,
            'images': images, 'requires_pdf_verification': True}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    resolve = sub.add_parser('resolve')
    resolve.add_argument('search_json', type=Path)
    for flag in ('title', 'doi', 'key', 'author', 'year'):
        resolve.add_argument('--' + flag)
    cache = sub.add_parser('mineru')
    cache.add_argument('cache', type=Path)
    for flag in ('attachment-id', 'attachment-key', 'parent-key'):
        cache.add_argument('--' + flag)
    args = parser.parse_args()
    if args.command == 'resolve':
        data = unwrap(json.loads(args.search_json.read_text(encoding='utf-8')))
        result = resolve_candidates(data['results'], title=args.title, doi=args.doi, key=args.key,
                                    author=args.author, year=args.year, limited=data.get('limited', False))
    else:
        result = inspect_mineru(args.cache, attachment_id=args.attachment_id,
                                attachment_key=args.attachment_key, parent_key=args.parent_key)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
