import importlib.util
import tempfile
import unittest
from pathlib import Path
import fitz
from PIL import Image

path = Path(__file__).resolve().parents[1] / 'zotero-journal-club' / 'scripts' / 'extract_figures.py'
spec = importlib.util.spec_from_file_location('figure_extraction', path)
figures = importlib.util.module_from_spec(spec)
spec.loader.exec_module(figures)


class FigureExtractionTests(unittest.TestCase):
    def make_pdf(self, root):
        source = root / 'source.pdf'
        with fitz.open() as doc:
            for color in [(1, 0, 0), (0, 0, 1)]:
                page = doc.new_page(width=200, height=200)
                page.draw_rect(fitz.Rect(20, 20, 100, 100), color=color, fill=color)
            doc.save(source)
        return source

    def test_crop_uses_the_requested_page_and_keeps_source_and_locator(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = self.make_pdf(root)
            before = source.read_bytes()
            result = figures.extract(source, [{'id': 'F2', 'pdf_page': 2, 'bbox': [25, 25, 95, 95],
                                               'figure': 'Figure 2', 'panel': 'b'}], root / 'assets', dpi=144)
            entry = result['sources'][0]
            with Image.open(entry['path']) as image:
                self.assertEqual(image.size, (140, 140))
                self.assertEqual(image.getpixel((70, 70)), (0, 0, 255))
            self.assertEqual(entry['locator']['pdf_page'], 2)
            self.assertEqual(entry['locator']['panel'], 'b')
            self.assertEqual(entry['verification'], 'unchecked')
            self.assertEqual(source.read_bytes(), before)
            self.assertTrue((root / 'assets' / 'figure_manifest.json').is_file())

    def test_invalid_regions_and_existing_outputs_are_rejected_before_writing(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = self.make_pdf(root)
            valid = {'id': 'F1', 'pdf_page': 1, 'bbox': [20, 20, 100, 100]}
            for change in ({'pdf_page': 0}, {'pdf_page': 3}, {'bbox': [0, 0, 201, 100]},
                           {'bbox': [10, 10, 5, 5]}, {'id': '../outside'}, {'bbox': [0, 0, float('nan'), 5]}):
                with self.subTest(change=change), self.assertRaises(ValueError):
                    figures.extract(source, [{**valid, **change}], root / 'not-created')
                self.assertFalse((root / 'not-created').exists())
            figures.extract(source, [valid], root / 'assets')
            before = (root / 'assets' / 'F1.png').read_bytes()
            with self.assertRaises(FileExistsError):
                figures.extract(source, [valid], root / 'assets')
            self.assertEqual((root / 'assets' / 'F1.png').read_bytes(), before)


if __name__ == '__main__':
    unittest.main()
