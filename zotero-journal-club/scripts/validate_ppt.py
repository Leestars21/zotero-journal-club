#!/usr/bin/env python3
"""Check PPTX structure; visual and scientific verification remain manual."""
import argparse
import hashlib
import json
import re
from collections import Counter
from itertools import combinations
from math import ceil, isfinite
from pathlib import Path
from unicodedata import east_asian_width
from zipfile import ZipFile
from pptx import Presentation


def all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == 6:
            yield from all_shapes(shape.shapes)


def texts(shape):
    if shape.has_text_frame and shape.text.strip():
        return [shape.text]
    if shape.has_table:
        return [cell.text for row in shape.table.rows for cell in row.cells if cell.text.strip()]
    return []


def possible_text_overflow(shape):
    if not shape.has_text_frame or not shape.text.strip():
        return False
    frame = shape.text_frame
    width = shape.width - frame.margin_left - frame.margin_right
    height = shape.height - frame.margin_top - frame.margin_bottom
    if width <= 0 or height <= 0:
        return True
    estimate = 0
    for paragraph in frame.paragraphs:
        sizes = [run.font.size.pt for run in paragraph.runs if run.font.size is not None]
        size = max(sizes, default=paragraph.font.size.pt if paragraph.font.size is not None else 20)
        units = sum(1 if east_asian_width(char) in ('W', 'F') else .5 for char in paragraph.text)
        lines = max(1, ceil(units * size * 12700 / width))
        estimate += lines * size * 12700 * 1.15
    return estimate > height * 1.15


def asset_path(data, value):
    path = Path(value)
    return path if path.is_absolute() else Path(data.get('_base', '.')) / path


def source_ids(value, context, issues):
    ids = value.get('source_ids', [])
    if not isinstance(ids, list) or any(not isinstance(identifier, str) or not identifier.strip() for identifier in ids):
        issues.append(f'{context}: source_ids must be a list of nonempty identifiers')
        return []
    return ids


def content_blocks(entry, context, issues):
    blocks = [entry]
    panels = entry.get('panels', [])
    if not isinstance(panels, list) or any(not isinstance(panel, dict) for panel in panels):
        issues.append(f'{context}: panels must be a list of content objects')
    else:
        blocks.extend(panels)
    rows = entry.get('summary_rows', [])
    if not isinstance(rows, list) or any(not isinstance(row, dict) for row in rows):
        issues.append(f'{context}: summary_rows must be a list of objects')
    else:
        blocks.extend(rows)
    return blocks


def check_plan(prs, plan, evidence, issues, warnings):
    entries = plan.get('slides')
    if not isinstance(entries, list):
        issues.append('plan: slides must be a list')
        return
    if len(entries) != len(prs.slides):
        issues.append('plan: slide count differs from the PPTX')
    slide_ids, references, durations = [], {}, []
    for index, (slide, entry) in enumerate(zip(prs.slides, entries), 1):
        if not isinstance(entry, dict):
            issues.append(f'plan slide {index}: expected an object')
            continue
        identifier = str(entry.get('id', f'slide-{index:03d}'))
        slide_ids.append(identifier)
        stored_id = slide._element.cSld.get('name')
        if entry.get('id') is not None and stored_id and stored_id != identifier:
            issues.append(f'slide {index}: slide identifier differs from the plan')
        references[identifier] = set(source_ids(entry, f'slide {index}', issues))
        blocks = content_blocks(entry, f'slide {index}', issues)
        shapes = list(all_shapes(slide.shapes))
        visible = '\n'.join(text for shape in shapes for text in texts(shape))
        expected = [entry.get('title', ''), entry.get('source', ''), entry.get('subtitle', ''),
                    entry.get('takeaway', '')]
        expected.extend(entry.get('metadata', []))
        expected.extend(entry.get('agenda', []))
        images = []
        for block in blocks:
            references[identifier].update(source_ids(block, f'slide {index} region', issues))
            if block is not entry:
                expected.append(block.get('title', ''))
                expected.append(block.get('text', ''))
            if block.get('columns') is not None:
                expected.extend(value for column in block['columns'] for value in column)
                expected.extend(block.get('column_titles', []))
            else:
                expected.extend(block.get('bullets', []))
            images.extend(block.get('images', []))
            if block.get('image'):
                images.append({'path': block['image'], 'caption': block.get('caption', ''),
                               'source_ids': block.get('source_ids', [])})
        actual_images = Counter(hashlib.sha256(shape.image.blob).hexdigest() for shape in shapes if shape.shape_type == 13)
        expected_images = Counter()
        for image in images:
            ids = source_ids(image, f'slide {index} image', issues) or list(references[identifier])
            references[identifier].update(ids)
            image_file = asset_path(plan, image['path'])
            if not image_file.is_file():
                issues.append(f'slide {index}: missing image file {image_file}')
            else:
                expected_images[hashlib.sha256(image_file.read_bytes()).hexdigest()] += 1
            if image.get('caption'):
                expected.append(image['caption'])
        if expected_images - actual_images:
            issues.append(f'slide {index}: a planned source image is absent or has been replaced in the PPTX')
        rich_blocks = [(field, region.get(field)) for region in blocks for field in ('table', 'flow', 'equation')]
        for field, block in rich_blocks:
            if not isinstance(block, dict):
                continue
            references[identifier].update(source_ids(block, f'slide {index} {field}', issues))
            if field == 'table':
                expected_table = [[str(value) for value in row] for row in [block['headers'], *block['rows']]]
                actual_tables = [[[cell.text for cell in row.cells] for row in shape.table.rows]
                                 for shape in shapes if shape.has_table]
                if expected_table not in actual_tables:
                    issues.append(f'slide {index}: planned table data is missing or changed')
                if block.get('caption'):
                    expected.append(block['caption'])
            elif field == 'flow':
                for step in block['steps']:
                    expected.append(step if isinstance(step, str) else step['title'])
                    if isinstance(step, dict) and step.get('detail'):
                        expected.append(step['detail'])
                if block.get('caption'):
                    expected.append(block['caption'])
            else:
                expected.append(block['text'])
                for variable in block.get('variables', []):
                    expected.extend([variable['symbol'], variable['meaning']] if isinstance(variable, dict) else [str(variable)])
        for value in expected:
            if value and str(value) not in visible:
                issues.append(f'slide {index}: planned text is missing from the PPTX: {str(value)[:80]}')
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ''
        if entry.get('notes') and entry['notes'] not in notes:
            issues.append(f'slide {index}: authored speaker notes were not retained')
        for field in ('reading_questions', 'takeaways'):
            for value in entry.get(field, []):
                if str(value) not in notes:
                    issues.append(f'slide {index}: {field} content is missing from speaker notes')
        for value in (entry.get('figure_reading') or {}).values():
            if str(value) not in notes:
                issues.append(f'slide {index}: figure reading guidance is missing from speaker notes')
        duration = entry.get('duration_seconds')
        if duration is not None:
            if isinstance(duration, bool) or not isinstance(duration, (int, float)) or not isfinite(duration) or duration <= 0:
                issues.append(f'slide {index}: duration_seconds must be a positive finite number')
            else:
                durations.append(duration)
                if f'建议时长：{duration:g} 秒' not in notes:
                    issues.append(f'slide {index}: planned speaking duration is missing from speaker notes')
    if len(slide_ids) != len(set(slide_ids)):
        issues.append('plan: duplicate slide identifiers')
    budget = plan.get('duration_minutes')
    if budget is not None and (isinstance(budget, bool) or not isinstance(budget, (int, float)) or not isfinite(budget) or budget <= 0):
        issues.append('plan: duration_minutes must be a positive finite number')
        budget = None
    if budget is not None and len(durations) < len(entries):
        warnings.append('speaking time is incomplete: some slides have no duration_seconds')
    if isinstance(budget, (int, float)) and not isinstance(budget, bool) and budget > 0 and sum(durations) > budget * 60 * 1.1:
        warnings.append(f'planned speaking time {sum(durations):g}s exceeds the {budget:g}-minute budget')
    if evidence is None:
        if any(references.values()):
            issues.append('plan declares source_ids but no evidence ledger was supplied')
        return
    records = evidence.get('sources', [])
    if not isinstance(records, list):
        issues.append('evidence: sources must be a list')
        return
    sources = {}
    for source in records:
        if not isinstance(source, dict) or not isinstance(source.get('id'), str) or not source['id']:
            issues.append('evidence: every source needs a nonempty id')
            continue
        if source['id'] in sources:
            issues.append(f'evidence: duplicate source id {source["id"]}')
        sources[source['id']] = source
    used = set().union(*references.values()) if references else set()
    for identifier, ids in references.items():
        if not ids:
            issues.append(f'{identifier}: no source_ids linked to the supplied evidence ledger')
        for source_id in ids - sources.keys():
            issues.append(f'{identifier}: unknown source id {source_id}')
    count = (evidence.get('pdf_source') or {}).get('page_count')
    if count is not None and (not isinstance(count, int) or isinstance(count, bool) or count < 1):
        issues.append('evidence: pdf_source.page_count must be a positive integer')
        count = None
    for source_id in used & sources.keys():
        source = sources[source_id]
        origin, locator = source.get('origin'), source.get('locator')
        if origin not in ('paper', 'user_note', 'presenter_inference'):
            issues.append(f'{source_id}: missing or invalid source origin')
        if origin == 'paper' and (not isinstance(locator, dict) or not any(value is not None for value in locator.values())):
            issues.append(f'{source_id}: paper source has no locator')
        page = locator.get('pdf_page') if isinstance(locator, dict) else None
        if page is not None and (not isinstance(page, int) or isinstance(page, bool) or page < 1 or
                                 (isinstance(count, int) and page > count)):
            issues.append(f'{source_id}: invalid one-based PDF page')
        if source.get('path') and not asset_path(evidence, source['path']).is_file():
            issues.append(f'{source_id}: source asset is missing')
        if origin == 'paper' and source.get('verification') in (None, 'unchecked', 'unverified'):
            warnings.append(f'{source_id}: paper evidence still requires verification against the original')
    claims = evidence.get('claims', [])
    if not isinstance(claims, list):
        issues.append('evidence: claims must be a list')
        claims = []
    claim_ids = set()
    for claim in claims:
        if not isinstance(claim, dict):
            issues.append('evidence: each claim must be an object')
            continue
        claim_id = claim.get('id')
        if not isinstance(claim_id, str) or not claim_id or claim_id in claim_ids:
            issues.append('evidence: claims need unique nonempty identifiers')
        else:
            claim_ids.add(claim_id)
        if claim.get('origin') not in ('paper', 'user_note', 'presenter_inference'):
            issues.append(f'claim {claim_id}: missing or invalid claim origin')
        ids = source_ids(claim, f'claim {claim.get("id", "unknown")}', issues)
        if any(identifier not in sources for identifier in ids):
            issues.append(f'claim {claim.get("id", "unknown")}: unknown evidence source')
        if claim.get('origin') == 'paper' and not any(sources.get(identifier, {}).get('origin') == 'paper' for identifier in ids):
            issues.append(f'claim {claim.get("id", "unknown")}: a paper finding cannot be supported only by notes or inference')
        destinations = claim.get('destination_slide', [])
        if not isinstance(destinations, list):
            destinations = [destinations]
        for destination in destinations:
            if not isinstance(destination, (str, int)) or isinstance(destination, bool):
                issues.append(f'claim {claim_id}: invalid destination slide')
                continue
            if isinstance(destination, int) and not isinstance(destination, bool) and 1 <= destination <= len(slide_ids):
                destination = slide_ids[destination - 1]
            if destination not in references:
                issues.append(f'claim {claim.get("id", "unknown")}: unknown destination slide {destination}')
            elif not set(ids).issubset(references[destination]):
                issues.append(f'claim {claim.get("id", "unknown")}: evidence not linked on destination slide {destination}')
    for index, entry in enumerate(entries[:len(prs.slides)], 1):
        if not isinstance(entry, dict):
            continue
        images = []
        for block in content_blocks(entry, f'slide {index}', issues):
            images.extend(block.get('images', []))
            if block.get('image'):
                images.append({'path': block['image'], 'source_ids': block.get('source_ids', [])})
        for image in images:
            ids = source_ids(image, f'slide {index} image', issues) or source_ids(entry, f'slide {index}', issues)
            declared = [sources[identifier] for identifier in ids if identifier in sources and sources[identifier].get('path')]
            target = asset_path(plan, image['path'])
            if declared and target.is_file():
                digest = hashlib.sha256(target.read_bytes()).digest()
                if not any(asset_path(evidence, source['path']).is_file() and
                           hashlib.sha256(asset_path(evidence, source['path']).read_bytes()).digest() == digest for source in declared):
                    issues.append(f'slide {index}: image does not match its declared source asset')


def inspect(path, plan=None, evidence=None):
    with ZipFile(path) as package:
        bad = package.testzip()
    prs = Presentation(str(path))
    issues = []
    warnings = []
    text_objects = 0
    pictures = 0
    notes_count = 0
    tables = 0
    source_count = 0
    for i, slide in enumerate(prs.slides, 1):
        if not slide.has_notes_slide or not slide.notes_slide.notes_text_frame.text.strip():
            issues.append(f'slide {i}: missing speaker notes')
        else:
            notes_count += 1
        text_shapes = []
        found_source = False
        for shape in all_shapes(slide.shapes):
            content = texts(shape)
            text_objects += len(content)
            if shape.has_text_frame and shape.text.strip():
                text_shapes.append(shape)
                if shape.name == 'journal-club-source' or (shape.top >= prs.slide_height * .82 and
                        re.search(r'(?:来源|原文|汇报者|合成|\bsynthetic\b|\bsource\b|\bPDF\b|doi\.org/|10\.\d{4,9}/|\b(?:19|20)\d{2}\b)', shape.text, re.I)):
                    found_source = True
                if possible_text_overflow(shape):
                    warnings.append(f'slide {i}: text box {shape.name!r} may overflow; review visually')
            if shape.has_table:
                tables += 1
            if shape.shape_type == 13:
                pictures += 1
                if min(shape.width, shape.height) < 4 * 914400:
                    warnings.append(f'slide {i}: image {shape.name!r} has a small displayed dimension; review labels at presentation size')
                width, height = shape.image.size
                effective_dpi = min(width / (shape.width / 914400), height / (shape.height / 914400)) if min(shape.width, shape.height) else 0
                if effective_dpi < 120:
                    warnings.append(f'slide {i}: image {shape.name!r} has low effective resolution ({effective_dpi:.0f} dpi)')
        if found_source:
            source_count += 1
        else:
            issues.append(f'slide {i}: missing source label; use a named journal-club-source box or an explicit citation footer')
        for shape in slide.shapes:
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
    if plan is not None:
        check_plan(prs, plan, evidence, issues, warnings)
    elif evidence is not None:
        issues.append('an evidence ledger requires a slide plan to check destination bindings')
    return {'slide_count': len(prs.slides), 'editable_text_objects': text_objects,
            'slides_with_notes': notes_count, 'picture_objects': pictures, 'native_tables': tables,
            'slides_with_source_labels': source_count, 'plan_checked': plan is not None,
            'evidence_checked': plan is not None and evidence is not None,
            'issues': issues, 'warnings': warnings,
            'not_checked': ['scientific accuracy', 'all overlaps', 'text clipping', 'visual quality', 'source correctness']}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pptx', type=Path)
    parser.add_argument('--plan', type=Path, help='Authored slide plan to compare with actual PPTX content')
    parser.add_argument('--evidence', type=Path, help='Source ledger with sources and claims; requires --plan')
    args = parser.parse_args()
    def read_json(path):
        if path is None:
            return None
        value = json.loads(path.read_text(encoding='utf-8'))
        value['_base'] = str(path.parent)
        return value
    report = inspect(args.pptx, read_json(args.plan), read_json(args.evidence))
    print(json.dumps(report, ensure_ascii=False, indent=2))
    raise SystemExit(bool(report['issues']))
