import importlib.util
import json
import tempfile
import unittest
from pathlib import Path

path = Path(__file__).resolve().parents[1] / 'zotero-journal-club' / 'scripts' / 'collect_sources.py'
spec = importlib.util.spec_from_file_location('sources', path)
sources = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sources)


class SourceTests(unittest.TestCase):
    def setUp(self):
        self.items = [
            {'itemId': 1, 'itemKey': 'AAA', 'title': 'Shared title', 'firstCreator': 'Smith', 'year': '2020',
             'metadata': {'fields': {'DOI': '10.1234/A'}}},
            {'itemId': 2, 'itemKey': 'BBB', 'title': 'Shared title', 'firstCreator': 'Zhao', 'year': '2021',
             'metadata': {'fields': {'DOI': '10.1234/B'}}},
        ]

    def test_ambiguity_not_found_doi_and_paginated_results(self):
        self.assertEqual(sources.resolve_candidates(self.items, title='Shared title')['status'], 'ambiguous')
        self.assertIsNone(sources.resolve_candidates(self.items, title='missing')['selected'])
        result = sources.resolve_candidates(self.items, doi='https://doi.org/10.1234/B')
        self.assertEqual(result['selected']['itemId'], 2)
        self.assertIsNone(sources.resolve_candidates(self.items[:1], title='Shared title', limited=True)['selected'])
        self.assertEqual(sources.resolve_candidates(self.items, title='Shared title', author='Smith', year='2020')['selected']['itemId'], 1)

    def test_cache_identity_and_relative_images(self):
        with tempfile.TemporaryDirectory() as root:
            cache = Path(root)
            (cache / 'images').mkdir()
            (cache / 'images' / 'fig.png').write_bytes(b'image')
            (cache / 'full.md').write_text('![a](images/fig.png) ![b](images/missing.png)', encoding='utf-8')
            (cache / '_llm_source.json').write_text(json.dumps({'attachmentId': 9, 'parentItemKey': 'ABC'}), encoding='utf-8')
            (cache / 'content_list.json').write_text(json.dumps([{'type': 'image', 'img_path': 'images/fig.png', 'page_idx': 2}]), encoding='utf-8')
            good = sources.inspect_mineru(cache, attachment_id=9, parent_key='ABC')
            self.assertEqual(good['identity_status'], 'matched')
            self.assertEqual([x['status'] for x in good['images']], ['present', 'missing', 'present'])
            self.assertEqual(good['images'][2]['pdf_page'], 3)
            self.assertEqual(sources.inspect_mineru(cache, attachment_id=10)['identity_status'], 'conflict')
            self.assertEqual(sources.inspect_mineru(cache)['identity_status'], 'unverified')

    def test_missing_cache_is_explicit(self):
        with tempfile.TemporaryDirectory() as root:
            result = sources.inspect_mineru(Path(root) / 'absent', attachment_id=9)
            self.assertEqual(result['markdown'], [])
            self.assertTrue(result['requires_pdf_verification'])


if __name__ == '__main__':
    unittest.main()
