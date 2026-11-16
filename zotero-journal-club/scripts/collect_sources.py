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
    """Accept decoded results and the text/structured envelopes returned by MCP."""
    for _ in range(6):
        if not isinstance(data, dict):
            raise ValueError('Zotero search output must be a JSON object')
        if data.get('isError') or data.get('ok') is False:
            details = data.get('result', data.get('content', 'unknown error'))
            raise ValueError(f'Zotero search failed: {details}')
        if 'results' in data:
            if not isinstance(data['results'], list):
                raise ValueError('Zotero search results must be a list')
            return data
        if isinstance(data.get('structuredContent'), dict):
            data = data['structuredContent']
        elif isinstance(data.get('result'), dict):
            data = data['result']
        elif isinstance(data.get('content'), list):
            candidates = []
            for block in data['content']:
                if not isinstance(block, dict) or block.get('type') != 'text':
                    continue
                try:
                    value = json.loads(block.get('text', ''))
                except (TypeError, json.JSONDecodeError):
                    continue
                if isinstance(value, dict) and any(key in value for key in ('results', 'result', 'ok')):
                    candidates.append(value)
            if len(candidates) != 1:
                raise ValueError('Expected exactly one JSON search result in MCP text content')
            data = candidates[0]
        else:
            raise ValueError('No Zotero search results found in the supplied JSON')
    raise ValueError('Zotero search output has too many nested envelopes')


def resolve_search(data, **identity):
    data = unwrap(data)
    count = data.get('returnedCount', len(data['results']))
    total = data.get('totalCount')
    offset = data.get('offset', 0)
    incomplete = bool(data.get('limited') or data.get('nextOffset') is not None or offset)
    if isinstance(total, int) and isinstance(count, int):
        incomplete = incomplete or count < total
    return resolve_candidates(data['results'], limited=incomplete, **identity)


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
    if identity is not None and not isinstance(identity, dict):
        raise ValueError('MinerU source identity must be a JSON object')
    checks = [('attachmentId', attachment_id), ('attachmentKey', attachment_key), ('parentItemKey', parent_key)]
    conflicts = [name for name, expected in checks if expected is not None and identity and
                 name in identity and str(identity[name]).casefold() != str(expected).casefold()]
    evidence = [name for name, expected in checks if expected is not None and identity and
                name in identity and str(identity[name]).casefold() == str(expected).casefold()]
    attachment_verified = any(name in evidence for name in ('attachmentId', 'attachmentKey'))
    status = ('conflict' if conflicts else 'matched' if attachment_verified
              else 'parent_only' if 'parentItemKey' in evidence else 'unverified')
    missing = [name for name, expected in checks if expected is not None and
               (identity is None or name not in identity)]
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
            'attachment_verified': attachment_verified and not conflicts, 'missing_identity_fields': missing,
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
        data = json.loads(args.search_json.read_text(encoding='utf-8'))
        result = resolve_search(data, title=args.title, doi=args.doi, key=args.key,
                                author=args.author, year=args.year)
    else:
        result = inspect_mineru(args.cache, attachment_id=args.attachment_id,
                                attachment_key=args.attachment_key, parent_key=args.parent_key)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
