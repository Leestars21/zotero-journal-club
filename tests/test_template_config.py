import importlib.util
import json
import tempfile
import unittest
from pathlib import Path

path = Path(__file__).resolve().parents[1] / 'zotero-journal-club' / 'scripts' / 'resolve_template.py'
spec = importlib.util.spec_from_file_location('template_config', path)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class TemplateConfigTests(unittest.TestCase):
    def test_explicit_overrides_config_and_config_can_use_relative_path(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            first, second = root / 'default.pptx', root / 'specified.pptx'
            first.write_bytes(b'test')
            second.write_bytes(b'test')
            config = root / 'config.json'
            config.write_text(json.dumps({'default_template': first.name}), encoding='utf-8')
            self.assertEqual(module.resolve(config_path=config)['template'], str(first.resolve()))
            self.assertEqual(module.resolve(explicit=second, config_path=config)['template'], str(second.resolve()))

    def test_missing_default_is_error_and_absent_config_is_none(self):
        with tempfile.TemporaryDirectory() as folder:
            config = Path(folder) / 'config.json'
            self.assertIsNone(module.resolve(config_path=config)['template'])
            config.write_text('{"default_template":"missing.pptx"}', encoding='utf-8')
            with self.assertRaises(ValueError):
                module.resolve(config_path=config)

    def test_quality_reference_does_not_become_the_selected_template(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            reference, explicit = root / 'reference.pptx', root / 'new.pptx'
            reference.write_bytes(b'reference')
            explicit.write_bytes(b'explicit')
            config = root / 'config.json'
            config.write_text(json.dumps({'quality_reference': reference.name}), encoding='utf-8')
            self.assertIsNone(module.resolve(config_path=config)['template'])
            self.assertEqual(module.resolve_reference(config_path=config)['quality_reference'], str(reference.resolve()))
            self.assertEqual(module.resolve_reference(explicit=explicit, config_path=config)['quality_reference'], str(explicit.resolve()))

    def test_missing_quality_reference_is_explicit(self):
        with tempfile.TemporaryDirectory() as folder:
            config = Path(folder) / 'config.json'
            self.assertIsNone(module.resolve_reference(config_path=config)['quality_reference'])
            config.write_text('{"quality_reference":"missing.pptx"}', encoding='utf-8')
            with self.assertRaises(ValueError):
                module.resolve_reference(config_path=config)

if __name__ == '__main__':
    unittest.main()
