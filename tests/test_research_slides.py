import importlib.util
import tempfile
import unittest
from pathlib import Path
from PIL import Image
from pptx import Presentation
from pptx.util import Inches

ROOT = Path(__file__).resolve().parents[1] / 'zotero-journal-club' / 'scripts'


def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / (name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


builder = load('build_ppt')
validator = load('validate_ppt')


class ResearchSlideTests(unittest.TestCase):
    def entry(self, **content):
        return {'title': 'Research comparison', 'source': 'Synthetic source, PDF p. 1',
                'notes': 'Explain the comparison and its limitations.', **content}

    def test_default_columns_keep_all_content_and_validator_detects_later_loss(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / 'columns.pptx'
            plan = {'slides': [self.entry(columns=[['left evidence'], ['right evidence']])]}
            builder.build(plan, output)
            self.assertEqual(validator.inspect(output, plan)['issues'], [])
            deck = Presentation(output)
            right = next(shape for shape in deck.slides[0].shapes if shape.has_text_frame and shape.text == 'right evidence')
            right.text = ''
            deck.save(output)
            self.assertTrue(any('right evidence' in issue for issue in validator.inspect(output, plan)['issues']))

    def test_native_table_flow_and_equation_remain_editable_with_and_without_template(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            template = root / 'template.pptx'
            Presentation().save(template)
            content = [self.entry(table={'headers': ['Model', 'Error'], 'rows': [['A', .12], ['B', .09]]}),
                       self.entry(flow={'steps': [{'title': 'Data', 'detail': 'Measured response'}, 'Fit', 'Validate']}),
                       self.entry(equation={'text': 'y = ax + b', 'variables': [{'symbol': 'a', 'meaning': 'slope'}]})]
            for use_template in (False, True):
                entries = [dict(entry) for entry in content]
                if use_template:
                    for entry in entries:
                        entry.update(layout=6, regions={'title': [.05, .02, .9, .1], 'body': [.05, .18, .9, .66]})
                output = root / f'native-{use_template}.pptx'
                plan = {'slides': entries}
                builder.build(plan, output, template if use_template else None)
                deck = Presentation(output)
                table = next(shape.table for shape in deck.slides[0].shapes if shape.has_table)
                self.assertEqual(table.cell(2, 1).text, '0.09')
                flow = [shape for shape in deck.slides[1].shapes if shape.name.startswith('journal-club-flow-step')]
                self.assertEqual(len(flow), 3)
                self.assertIn('Measured response', flow[0].text)
                text = '\n'.join(shape.text for shape in deck.slides[2].shapes if shape.has_text_frame)
                self.assertIn('y = ax + b', text)
                self.assertIn('a：slope', text)
                self.assertEqual(validator.inspect(output, plan)['issues'], [])

    def test_multiple_source_images_and_individual_captions_are_retained(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            for name, color in [('left', 'red'), ('right', 'blue')]:
                Image.new('RGB', (800, 600), color).save(root / f'{name}.png')
            plan = {'_base': str(root), 'slides': [self.entry(images=[{'path': 'left.png', 'caption': 'Figure 1a'},
                                                                         {'path': 'right.png', 'caption': 'Figure 1b'}])]}
            output = root / 'figures.pptx'
            builder.build(plan, output)
            deck = Presentation(output)
            self.assertEqual(len([shape for shape in deck.slides[0].shapes if shape.shape_type == 13]), 2)
            self.assertEqual(validator.inspect(output, plan)['issues'], [])

    def test_reading_questions_figure_explanation_and_timing_enter_speaker_notes(self):
        with tempfile.TemporaryDirectory() as folder:
            plan = {'duration_minutes': 1, 'slides': [self.entry(
                duration_seconds=90, reading_questions=['Is the test geometry unseen?'],
                figure_reading={'comparison': 'Two baselines', 'scope': 'Observed geometry only'},
                takeaways=['Use a grouped holdout for a new geometry claim'])]}
            output = Path(folder) / 'notes.pptx'
            builder.build(plan, output)
            notes = Presentation(output).slides[0].notes_slide.notes_text_frame.text
            for value in ('Is the test geometry unseen?', 'Two baselines', 'Observed geometry only', '90 秒'):
                self.assertIn(value, notes)
            self.assertTrue(any('speaking time' in warning for warning in validator.inspect(output, plan)['warnings']))

    def test_unsupported_content_combinations_fail_instead_of_dropping_content(self):
        with tempfile.TemporaryDirectory() as folder:
            for content in ({'columns': [['x']], 'image': 'missing.png'},
                            {'images': [], 'equation': {'text': 'x'}},
                            {'table': {'headers': ['A', 'B'], 'rows': [['x']]}}):
                with self.subTest(content=content), self.assertRaises(ValueError):
                    builder.build({'slides': [self.entry(**content)]}, Path(folder) / 'not-written.pptx')
                self.assertFalse((Path(folder) / 'not-written.pptx').exists())

    def test_missing_citation_is_reported_even_when_structure_and_notes_are_valid(self):
        with tempfile.TemporaryDirectory() as folder:
            deck = Presentation()
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1)).text = 'A scientific finding'
            slide.notes_slide.notes_text_frame.text = 'Explain this finding.'
            output = Path(folder) / 'no-source.pptx'
            deck.save(output)
            report = validator.inspect(output)
            self.assertEqual(report['slides_with_source_labels'], 0)
            self.assertTrue(any('missing source label' in issue for issue in report['issues']))

    def test_short_box_with_long_text_warns_without_claiming_complete_clipping_detection(self):
        with tempfile.TemporaryDirectory() as folder:
            plan = {'slides': [self.entry(bullets=['A very long paragraph ' * 300])]}
            output = Path(folder) / 'overflow.pptx'
            builder.build(plan, output)
            report = validator.inspect(output)
            self.assertTrue(any('may overflow' in warning for warning in report['warnings']))
            self.assertIn('text clipping', report['not_checked'])

    def test_image_focus_keeps_multiple_lines_above_the_figure_and_rejects_excessive_text(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            image = root / 'figure.png'
            Image.new('RGB', (1000, 600), 'white').save(image)
            entry = self.entry(image=str(image), image_focus=True, bullets=['First line', 'Second line', 'Third line'])
            output = root / 'focused.pptx'
            builder.build({'slides': [entry]}, output)
            deck = Presentation(output)
            body = next(shape for shape in deck.slides[0].shapes if shape.has_text_frame and 'First line' in shape.text)
            picture = next(shape for shape in deck.slides[0].shapes if shape.shape_type == 13)
            self.assertLess(body.top + body.height, picture.top)
            self.assertFalse(validator.possible_text_overflow(body))
            self.assertIn('Third line', body.text)
            dense = self.entry(image=str(image), image_focus=True, bullets=['Very long detailed content ' * 100])
            with self.assertRaises(ValueError):
                builder.build({'slides': [dense]}, root / 'dense.pptx')
            self.assertFalse((root / 'dense.pptx').exists())

    def test_table_text_has_readable_contrast_with_inherited_template_styles(self):
        def luminance(rgb):
            values = [value/255 for value in rgb]
            values = [value/12.92 if value <= .04045 else ((value+.055)/1.055)**2.4 for value in values]
            return sum(value*weight for value,weight in zip(values, [.2126,.7152,.0722]))
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / 'table.pptx'
            builder.build({'slides': [self.entry(table={'headers':['Model','Error'], 'rows':[['A','.12']]})]}, output)
            table = next(shape.table for shape in Presentation(output).slides[0].shapes if shape.has_table)
            for row in table.rows:
                for cell in row.cells:
                    foreground = luminance(cell.text_frame.paragraphs[0].font.color.rgb)
                    background = luminance(cell.fill.fore_color.rgb)
                    self.assertGreaterEqual((max(foreground,background)+.05)/(min(foreground,background)+.05), 4.5)


class EvidenceTests(unittest.TestCase):
    def make_case(self, root):
        image = root / 'figure.png'
        Image.new('RGB', (800, 600), 'red').save(image)
        entry = {'id': 'result', 'title': 'Result', 'source': 'Synthetic source, Figure 1',
                 'notes': 'Explain the data.', 'source_ids': ['F1'],
                 'images': [{'path': str(image), 'source_ids': ['F1'], 'caption': 'Figure 1'}]}
        plan = {'slides': [entry]}
        evidence = {'pdf_source': {'page_count': 2}, 'sources': [
            {'id': 'F1', 'kind': 'figure', 'origin': 'paper', 'path': str(image),
             'locator': {'pdf_page': 1, 'figure': 'Figure 1'}, 'verification': 'checked_against_pdf'}],
            'claims': [{'id': 'C1', 'text': 'A result', 'origin': 'paper', 'source_ids': ['F1'], 'destination_slide': 'result'}]}
        output = root / 'result.pptx'
        builder.build(plan, output)
        return output, plan, evidence

    def test_source_ledger_binds_actual_images_and_destination_slides(self):
        with tempfile.TemporaryDirectory() as folder:
            output, plan, evidence = self.make_case(Path(folder))
            report = validator.inspect(output, plan, evidence)
            self.assertTrue(report['evidence_checked'])
            self.assertEqual(report['issues'], [])
            evidence['claims'][0]['destination_slide'] = 'absent'
            self.assertTrue(any('unknown destination slide' in issue for issue in validator.inspect(output, plan, evidence)['issues']))

    def test_mislabeled_image_is_rejected_even_when_both_files_exist(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            output, plan, evidence = self.make_case(root)
            other = root / 'other.png'
            Image.new('RGB', (800, 600), 'blue').save(other)
            evidence['sources'][0]['path'] = str(other)
            self.assertTrue(any('does not match its declared source' in issue for issue in validator.inspect(output, plan, evidence)['issues']))

    def test_unknown_missing_and_out_of_range_evidence_is_reported(self):
        with tempfile.TemporaryDirectory() as folder:
            output, plan, evidence = self.make_case(Path(folder))
            evidence['sources'][0]['locator']['pdf_page'] = 3
            evidence['sources'][0]['path'] = str(Path(folder) / 'missing.png')
            plan['slides'][0]['source_ids'].append('MISSING')
            issues = validator.inspect(output, plan, evidence)['issues']
            for signal in ('unknown source id', 'source asset is missing', 'invalid one-based PDF page'):
                self.assertTrue(any(signal in issue for issue in issues))
            self.assertTrue(any('no evidence ledger' in issue for issue in validator.inspect(output, plan)['issues']))

    def test_paper_finding_cannot_be_backed_only_by_a_zotero_note(self):
        with tempfile.TemporaryDirectory() as folder:
            output, plan, evidence = self.make_case(Path(folder))
            evidence['sources'][0]['origin'] = 'user_note'
            issues = validator.inspect(output, plan, evidence)['issues']
            self.assertTrue(any('cannot be supported only by notes' in issue for issue in issues))


if __name__ == '__main__':
    unittest.main()
