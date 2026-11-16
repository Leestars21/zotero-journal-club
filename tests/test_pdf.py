import importlib.util
import tempfile
import unittest
from pathlib import Path
import fitz

path = Path(__file__).resolve().parents[1] / 'zotero-journal-club' / 'scripts' / 'inspect_pdf.py'
spec = importlib.util.spec_from_file_location('inspect_pdf', path)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class PdfOnlyTests(unittest.TestCase):
    def test_pdf_only_extraction_preserves_page_numbers_and_source(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / 'paper.pdf'
            doc = fitz.open()
            for text in ('First page', 'Second page'):
                page = doc.new_page()
                page.insert_text((72, 72), text)
            doc.save(source)
            doc.close()
            before = source.read_bytes()
            result = module.inspect(source, root / 'text')
            self.assertEqual(result['page_count'], 2)
            self.assertEqual([p['pdf_page'] for p in result['pages']], [1, 2])
            self.assertEqual([p['page_index'] for p in result['pages']], [0, 1])
            self.assertIn('Second page', (root / 'text' / 'page-0002.txt').read_text(encoding='utf-8'))
            self.assertEqual(source.read_bytes(), before)
            with self.assertRaises(FileExistsError):
                module.inspect(source, root / 'text')


if __name__ == '__main__':
    unittest.main()
