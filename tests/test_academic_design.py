import importlib.util
import tempfile
import unittest
from pathlib import Path

from PIL import Image
from pptx import Presentation


SCRIPTS = Path(__file__).resolve().parents[1] / 'zotero-journal-club' / 'scripts'


def load(name):
    spec = importlib.util.spec_from_file_location(name, SCRIPTS / f'{name}.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


builder, validator, inspector = load('build_ppt'), load('validate_ppt'), load('inspect_template')


def entry(**values):
    return {'title': 'Research evidence', 'source': 'Synthetic example, no experimental evidence',
            'notes': 'Explain the synthetic example and its scope.', **values}


class AcademicDesignTests(unittest.TestCase):
    def test_cover_agenda_and_content_preserve_editable_hierarchy_and_page_numbers(self):
        with tempfile.TemporaryDirectory() as folder:
            plan = {'design': {'preset': 'academic-blue'}, 'slides': [
                entry(page_type='cover', title='Synthetic research presentation',
                      subtitle='A visual layout example', metadata=['Synthetic author, 2026']),
                entry(page_type='agenda', title='Outline', agenda=['Method', 'Evidence', 'Discussion']),
                entry(section='2 Evidence', bullets=['Compare identical conditions'],
                      takeaway='Only the evaluated conditions support the conclusion', emphasis=['evaluated conditions'])]}
            output = Path(folder) / 'design.pptx'
            builder.build(plan, output)
            deck = Presentation(output)
            report = validator.inspect(output, plan)
            self.assertEqual(report['issues'], [])
            self.assertEqual(report['slides_with_source_labels'], 3)
            for index, slide in enumerate(deck.slides, 1):
                number = next(shape.text for shape in slide.shapes if shape.name == 'journal-club-page-number')
                self.assertEqual(number, f'{index}/3')
                self.assertTrue(slide.notes_slide.notes_text_frame.text.strip())
            takeaway = next(shape for shape in deck.slides[2].shapes if shape.name == 'journal-club-takeaway')
            self.assertTrue(any(str(run.font.color.rgb) == 'C00000' for run in takeaway.text_frame.paragraphs[0].runs))

    def test_nested_panel_images_are_bound_and_lost_captions_are_detected(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            images = []
            for name, color in [('A', 'red'), ('B', 'blue')]:
                path = root / f'{name}.png'
                Image.new('RGB', (1000, 800), color).save(path)
                images.append(path)
            plan = {'design': {'preset': 'academic-blue'}, 'slides': [entry(
                id='comparison', source_ids=['A', 'B'], takeaway='Compare the two identified panels',
                panels=[{'title': 'Condition A', 'images': [{'path': str(images[0]), 'caption': 'Panel A', 'source_ids': ['A']}]},
                        {'title': 'Condition B', 'images': [{'path': str(images[1]), 'caption': 'Panel B', 'source_ids': ['B']}]}])]}
            evidence = {'pdf_source': {'page_count': 1}, 'sources': [
                {'id': name, 'origin': 'paper', 'kind': 'figure', 'path': str(path),
                 'locator': {'pdf_page': 1, 'panel': name}, 'verification': 'checked_against_pdf'}
                for name, path in zip(['A', 'B'], images)],
                'claims': [{'id': 'C', 'origin': 'paper', 'source_ids': ['A', 'B'], 'destination_slide': 'comparison'}]}
            output = root / 'panels.pptx'
            builder.build(plan, output)
            self.assertEqual(validator.inspect(output, plan, evidence)['issues'], [])
            deck = Presentation(output)
            next(shape for shape in deck.slides[0].shapes if shape.has_text_frame and shape.text == 'Panel B').text = ''
            deck.save(output)
            self.assertTrue(any('Panel B' in issue for issue in validator.inspect(output, plan, evidence)['issues']))
            evidence['sources'][1]['path'] = str(images[0])
            self.assertTrue(any('does not match its declared source' in issue
                                for issue in validator.inspect(output, plan, evidence)['issues']))

    def test_table_and_flow_remain_native_and_header_has_contrast(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / 'native.pptx'
            plan = {'design': {'preset': 'academic-blue'}, 'slides': [
                entry(table={'headers': ['Model', 'Synthetic metric'], 'rows': [['A', '0.12'], ['B', '0.09']]}),
                entry(panels=[{'title': 'Model equation', 'equation': {'text': 'y = ax + b',
                      'variables': [{'symbol': 'a', 'meaning': 'Slope'}, {'symbol': 'b', 'meaning': 'Intercept'}]}},
                      {'title': 'Validation method', 'flow': {'steps': ['Data', 'Fit', 'Validate']}}]) ]}
            builder.build(plan, output)
            deck = Presentation(output)
            table = next(shape.table for shape in deck.slides[0].shapes if shape.has_table)
            self.assertEqual(table.cell(2, 1).text, '0.09')
            self.assertEqual(str(table.cell(0, 0).fill.fore_color.rgb), '003F88')
            self.assertEqual(str(table.cell(0, 0).text_frame.paragraphs[0].font.color.rgb), 'FFFFFF')
            steps = [shape for shape in deck.slides[1].shapes if shape.name.startswith('journal-club-flow-step')]
            self.assertEqual(len(steps), 3)
            self.assertEqual(validator.inspect(output, plan)['issues'], [])

    def test_crowded_content_and_conflicting_templates_fail_without_writing(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            template = root / 'template.pptx'
            Presentation().save(template)
            original = template.read_bytes()
            crowded = {'design': {'preset': 'academic-blue'}, 'slides': [entry(bullets=['Very dense content ' * 500])]}
            with self.assertRaises(ValueError):
                builder.build(crowded, root / 'crowded.pptx')
            self.assertFalse((root / 'crowded.pptx').exists())
            dense_table = {'design': {'preset': 'academic-blue'}, 'slides': [entry(
                table={'headers': ['A'], 'rows': [['Dense table']] * 40})]}
            with self.assertRaises(ValueError):
                builder.build(dense_table, root / 'dense-table.pptx')
            self.assertFalse((root / 'dense-table.pptx').exists())
            with self.assertRaises(ValueError):
                builder.build(crowded, root / 'conflict.pptx', template)
            self.assertEqual(template.read_bytes(), original)
            self.assertFalse((root / 'conflict.pptx').exists())

    def test_duration_removed_after_generation_is_reported(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / 'notes.pptx'
            plan = {'design': {'preset': 'academic-blue'}, 'slides': [entry(duration_seconds=60)]}
            builder.build(plan, output)
            deck = Presentation(output)
            frame = deck.slides[0].notes_slide.notes_text_frame
            frame.text = frame.text.replace('建议时长：60 秒', '')
            deck.save(output)
            self.assertTrue(any('duration is missing' in issue for issue in validator.inspect(output, plan)['issues']))

    def test_design_inspection_reports_native_content_and_explicit_styles(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / 'inspect.pptx'
            plan = {'design': {'preset': 'academic-blue'}, 'slides': [entry(
                table={'headers': ['Model', 'Metric'], 'rows': [['A', '0.12']]}, takeaway='Synthetic data only')]}
            builder.build(plan, output)
            result = inspector.inspect(output, include_design=True)['design_summary']
            self.assertEqual(result['slide_content_counts'][0]['tables'], 1)
            self.assertIn('Microsoft YaHei', [name for name, count in result['fonts_by_character_count']])
            self.assertIn('003F88', [color for color, count in result['explicit_text_colors_by_character_count']])

    def test_summary_rows_retain_content_and_validator_detects_a_lost_row(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / 'rows.pptx'
            plan = {'design': {'preset': 'academic-blue'}, 'slides': [entry(summary_rows=[
                {'title': 'Model comparison', 'text': 'Keep the dataset and budget identical'},
                {'title': 'Validation scope', 'text': 'Hold out the target inputs'}]) ]}
            builder.build(plan, output)
            self.assertEqual(validator.inspect(output, plan)['issues'], [])
            deck = Presentation(output)
            next(shape for shape in deck.slides[0].shapes if shape.has_text_frame and
                 shape.text == 'Hold out the target inputs').text = ''
            deck.save(output)
            self.assertTrue(any('Hold out the target inputs' in issue for issue in validator.inspect(output, plan)['issues']))

    def test_long_flow_keeps_six_native_steps_and_their_details(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / 'flow.pptx'
            plan = {'design': {'preset': 'academic-blue'}, 'slides': [entry(flow={'steps': [
                {'title': f'Step {index}', 'detail': f'Observed detail {index}'} for index in range(1, 7)]})]}
            builder.build(plan, output)
            self.assertEqual(validator.inspect(output, plan)['issues'], [])
            steps = [shape for shape in Presentation(output).slides[0].shapes
                     if shape.name.startswith('journal-club-flow-step')]
            self.assertEqual(len(steps), 6)


if __name__ == '__main__':
    unittest.main()
