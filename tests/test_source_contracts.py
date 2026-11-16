import importlib.util
import json
import tempfile
import unittest
from pathlib import Path

path = Path(__file__).resolve().parents[1] / 'zotero-journal-club' / 'scripts' / 'collect_sources.py'
spec = importlib.util.spec_from_file_location('source_contracts', path)
sources = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sources)


class SourceContractTests(unittest.TestCase):
    def setUp(self):
        self.result = {'results': [{'itemKey': 'ABC', 'title': 'Exact paper'}],
                       'totalCount': 1, 'returnedCount': 1, 'limited': False}

    def test_raw_mcp_and_structured_results_select_the_same_paper(self):
        envelopes = [self.result, {'ok': True, 'result': self.result},
                     {'structuredContent': {'ok': True, 'result': self.result}},
                     {'content': [{'type': 'text', 'text': json.dumps({'ok': True, 'result': self.result})}]}]
        for envelope in envelopes:
            with self.subTest(envelope=envelope):
                self.assertEqual(sources.resolve_search(envelope, key='ABC')['selected']['itemKey'], 'ABC')

    def test_failed_or_ambiguous_envelopes_do_not_masquerade_as_no_match(self):
        for envelope in ({'ok': False, 'result': {'error': 'offline'}},
                         {'content': [{'type': 'text', 'text': json.dumps(self.result)}] * 2},
                         {'results': 'invalid'}):
            with self.subTest(envelope=envelope), self.assertRaises(ValueError):
                sources.resolve_search(envelope, key='ABC')

    def test_pagination_metadata_prevents_automatic_selection(self):
        for extra in ({'totalCount': 2}, {'nextOffset': 1}, {'offset': 1}, {'limited': True}):
            result = sources.resolve_search({**self.result, **extra}, key='ABC')
            self.assertIsNone(result['selected'])
            self.assertEqual(result['status'], 'incomplete_search')

    def test_parent_identity_does_not_verify_a_sibling_pdf(self):
        with tempfile.TemporaryDirectory() as folder:
            cache = Path(folder)
            (cache / '_llm_source.json').write_text(json.dumps({'parentItemKey': 'ABC'}), encoding='utf-8')
            result = sources.inspect_mineru(cache, attachment_id=9, parent_key='ABC')
            self.assertEqual(result['identity_status'], 'parent_only')
            self.assertFalse(result['attachment_verified'])
            self.assertIn('attachmentId', result['missing_identity_fields'])

    def test_attachment_conflict_wins_over_parent_match(self):
        with tempfile.TemporaryDirectory() as folder:
            cache = Path(folder)
            (cache / '_llm_source.json').write_text(json.dumps({'attachmentId': 10, 'parentItemKey': 'ABC'}), encoding='utf-8')
            result = sources.inspect_mineru(cache, attachment_id=9, parent_key='ABC')
            self.assertEqual(result['identity_status'], 'conflict')
            self.assertFalse(result['attachment_verified'])


if __name__ == '__main__':
    unittest.main()
