import importlib.util
import tempfile
import unittest
from pathlib import Path
from pptx import Presentation
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches
from PIL import Image

ROOT = Path(__file__).resolve().parents[1] / 'zotero-journal-club' / 'scripts'
def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / (name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

build_ppt = load('build_ppt')
validate_ppt = load('validate_ppt')

class PresentationTests(unittest.TestCase):
    def test_default_and_template_keep_notes_and_editable_text(self):
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            template = folder / 'template.pptx'
            prs = Presentation()
            prs.slides.add_slide(prs.slide_layouts[0])
            prs.slides[0].shapes.title.text = 'SAMPLE SHOULD DISAPPEAR'
            prs.save(template)
            plan = {'slides': [{'title': '合成测试：公式', 'bullets': ['y = ax + b', 'a 表示斜率'],
                                'source': '合成资料，第 1 页', 'notes': '解释斜率与截距，衔接下一页。60 秒。'}]}
            for source in (None, template):
                output = folder / ('template-output.pptx' if source else 'default.pptx')
                build_ppt.build(plan, output, source)
                report = validate_ppt.inspect(output)
                self.assertEqual(report['slide_count'], 1)
                self.assertGreater(report['editable_text_objects'], 0)
                self.assertEqual(report['issues'], [])
                result = Presentation(output)
                self.assertNotIn('SAMPLE SHOULD DISAPPEAR', ' '.join(s.text for s in result.slides[0].shapes if s.has_text_frame))
                self.assertIn('斜率', result.slides[0].notes_slide.notes_text_frame.text)
                if source:
                    self.assertEqual(len(result.slide_masters), len(prs.slide_masters))
                    self.assertTrue(result.slides[0].shapes.title.is_placeholder)

    def test_layout_without_placeholders_fails_instead_of_drawing_over_template(self):
        with tempfile.TemporaryDirectory() as folder:
            template = Path(folder) / 'template.pptx'
            Presentation().save(template)
            plan = {'slides': [{'title': 'x', 'source': 'source', 'notes': 'notes', 'layout': 6}]}
            with self.assertRaisesRegex(ValueError, 'lacks suitable'):
                build_ppt.build(plan, Path(folder) / 'out.pptx', template)

    def test_two_columns_populate_both_placeholders(self):
        with tempfile.TemporaryDirectory() as folder:
            template = Path(folder) / 'template.pptx'
            Presentation().save(template)
            plan = {'slides': [{'title': 'title', 'columns': [['left'], ['right']],
                                'source': 'source', 'notes': 'notes', 'layout': 3}]}
            out = Path(folder) / 'result.pptx'
            build_ppt.build(plan, out, template)
            text = [s.text for s in Presentation(out).slides[0].shapes if s.has_text_frame]
            self.assertIn('left', text)
            self.assertIn('right', text)

    def test_explicit_custom_layout_regions_are_editable_and_reject_bad_geometry(self):
        with tempfile.TemporaryDirectory() as folder:
            template = Path(folder) / 'template.pptx'
            Presentation().save(template)
            regions = {'title': [.05, .02, .8, .1], 'body': [.05, .2, .4, .6],
                       'body2': [.52, .2, .4, .6], 'body_font_name': 'Arial'}
            entry = {'title': 'title', 'columns': [['left'], ['right']],
                     'source': 'source', 'notes': 'notes', 'layout': 6, 'regions': regions}
            out = Path(folder) / 'result.pptx'
            build_ppt.build({'slides': [entry]}, out, template)
            texts = [s.text for s in Presentation(out).slides[0].shapes if s.has_text_frame]
            self.assertIn('left', texts)
            self.assertIn('right', texts)
            left_shape = next(s for s in Presentation(out).slides[0].shapes if s.has_text_frame and s.text == 'left')
            self.assertEqual(left_shape.text_frame.paragraphs[0].font.name, 'Arial')
            regions['title'] = [.9, .02, .8, .1]
            with self.assertRaisesRegex(ValueError, 'inside the slide'):
                build_ppt.build({'slides': [entry]}, Path(folder) / 'bad.pptx', template)

    def test_missing_notes_or_source_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaises(ValueError):
                build_ppt.build({'slides': [{'title': 'x', 'source': '', 'notes': ''}]}, Path(folder) / 'x.pptx')

    def test_template_is_never_overwritten(self):
        with tempfile.TemporaryDirectory() as folder:
            template = Path(folder) / 'template.pptx'
            Presentation().save(template)
            before = template.read_bytes()
            with self.assertRaises(ValueError):
                build_ppt.build({'slides': []}, template, template)
            self.assertEqual(before, template.read_bytes())

    def test_existing_output_is_never_overwritten(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / 'output.pptx'
            output.write_bytes(b'existing user file')
            with self.assertRaises(FileExistsError):
                build_ppt.build({'slides': []}, output)
            self.assertEqual(output.read_bytes(), b'existing user file')

    def test_overlap_warning_does_not_claim_visual_validation(self):
        with tempfile.TemporaryDirectory() as folder:
            p = Presentation()
            slide = p.slides.add_slide(p.slide_layouts[6])
            for text in ('first', 'second'):
                shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
                shape.text = text
            slide.notes_slide.notes_text_frame.text = 'explanation'
            file = Path(folder) / 'overlap.pptx'
            p.save(file)
            report = validate_ppt.inspect(file)
            self.assertTrue(any('overlap' in warning for warning in report['warnings']))
            self.assertIn('text clipping', report['not_checked'])

    def test_image_focus_allocates_more_vertical_space_and_small_image_warns(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            image = root / 'portrait.png'
            Image.new('RGB', (100, 300), 'white').save(image)
            common = {'title': 'title', 'source': 'source', 'notes': 'notes', 'image': str(image)}
            normal, focus = root / 'normal.pptx', root / 'focus.pptx'
            build_ppt.build({'slides': [common]}, normal)
            build_ppt.build({'slides': [{**common, 'image_focus': True}]}, focus)
            image_shape = lambda file: next(s for s in Presentation(file).slides[0].shapes if s.shape_type == 13)
            self.assertGreater(image_shape(focus).height, image_shape(normal).height)
            report = validate_ppt.inspect(normal)
            self.assertTrue(any('small displayed dimension' in warning for warning in report['warnings']))

    def test_updates_only_total_next_to_slide_number_field(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(.4))
        paragraph = shape.text_frame.paragraphs[0]
        field = OxmlElement('a:fld')
        field.set('type', 'slidenum')
        field.set('id', '{42FEAD30-4C25-4D9D-BD9F-816FF90208F6}')
        field_text = OxmlElement('a:t')
        field_text.text = '1'
        field.append(field_text)
        paragraph._p.append(field)
        paragraph.add_run().text = '/20'
        other = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(2), Inches(.4))
        other.text = 'ratio /20'
        prs.slide_masters[0].shapes._spTree.append(shape._element)
        prs.slide_masters[0].shapes._spTree.append(other._element)
        build_ppt.update_total_slide_numbers(prs, 7)
        self.assertIn('/7', shape.text)
        self.assertEqual(other.text, 'ratio /20')

if __name__ == '__main__':
    unittest.main()
