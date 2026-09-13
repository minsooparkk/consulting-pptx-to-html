"""Semantic procedure validation and CLI preservation checks (stdlib only)."""
import copy
import hashlib
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import unittest
import uuid

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
from procedural import ObjectInventory, build_procedural_reveal, positive_interval
from pptx_html import Converter


def obj(identifier, content=''):
    return f'<div class="ppt-object" id="{identifier}" data-object-id="{identifier}">{content}</div>'


# Entirely synthetic objects; no customer deck, course text, or font binaries.
HTML = ('<section class="ppt-slide" data-slide-index="1">'
        + obj('s001_slide_10', 'Other slide') + '</section>'
        + '<section class="ppt-slide" data-slide-index="2">'
        + obj('s002_master_2', 'Footer') + obj('s002_slide_10', 'Heading')
        + obj('s002_slide_101', 'First') + obj('s002_slide_102', '<svg><path d="M0,0 L1,1"/></svg>')
        + obj('s002_slide_103', 'Second') + obj('s002_slide_104', 'Conclusion')
        + obj('s002_slide_200', obj('s002_slide_201', 'Nested')) + '</section>')


def plan():
    return {'slides': [{'slide': 2, 'groups': [
        {'label': 'First step', 'ids': ['s002_slide_101']},
        {'label': 'Second step', 'ids': ['s002_slide_102', 's002_slide_103']}],
        'conclusionIds': ['s002_slide_104']}]}


class ProcedureValidationTests(unittest.TestCase):
    def test_order_static_defaults_and_conclusion_timing(self):
        raw = plan()
        original = copy.deepcopy(raw)
        spec, audit = build_procedural_reveal(raw, HTML)
        self.assertEqual(raw, original)
        self.assertEqual(spec['slides'][0]['groups'], raw['slides'][0]['groups'])
        self.assertEqual(spec['slides'][0]['staticIds'],
                         ['s002_master_2', 's002_slide_10', 's002_slide_200', 's002_slide_201'])
        self.assertEqual(spec['interval_ms'], 300)
        self.assertEqual(spec['scheduler'], 'cumulative-timers')
        self.assertTrue(spec['completion_clears_hidden_state'])
        self.assertEqual(audit['maximum_total_ms'], 1170)
        self.assertEqual(audit['maximum_cleanup_ms'], 1250)
        self.assertEqual(audit['slide_timings'][0]['total_groups'], 3)
        self.assertFalse(audit['browser_verified'])
        raw['slides'][0].pop('conclusionIds')
        _, without_conclusion = build_procedural_reveal(raw, HTML, 600)
        self.assertEqual(without_conclusion['maximum_total_ms'], 1170)
        self.assertEqual(without_conclusion['slide_timings'][0]['conclusion_groups'], 0)

    def test_parent_container_covers_children_but_child_can_animate_alone(self):
        raw = {'slides': [{'slide': 2, 'groups': [{'label': 'Container', 'ids': ['s002_slide_200']}]}]}
        spec, _ = build_procedural_reveal(raw, HTML)
        self.assertNotIn('s002_slide_201', spec['slides'][0]['staticIds'])
        raw['slides'][0]['groups'][0]['ids'] = ['s002_slide_201']
        raw['slides'][0]['staticIds'] = ['s002_slide_200']
        spec, _ = build_procedural_reveal(raw, HTML)
        self.assertIn('s002_slide_200', spec['slides'][0]['staticIds'])

    def test_invalid_ownership_and_duplicate_assignment_rejected(self):
        cases = []
        for ids in [[], ['missing'], ['s001_slide_10'], ['s002_slide_101', 's002_slide_101'], [123]]:
            raw = plan(); raw['slides'][0]['groups'][0]['ids'] = ids; cases.append(raw)
        raw = plan(); raw['slides'][0]['groups'][1]['ids'] = ['s002_slide_101']; cases.append(raw)
        raw = plan(); raw['slides'][0]['conclusionIds'] = ['s002_slide_101']; cases.append(raw)
        raw = plan(); raw['slides'][0]['staticIds'] = ['s002_slide_101']; cases.append(raw)
        raw = plan(); raw['slides'][0]['staticIds'] = ['s001_slide_10']; cases.append(raw)
        raw = plan(); raw['slides'][0]['staticIds'] = ['s002_slide_10', 's002_slide_10']; cases.append(raw)
        for raw in cases:
            with self.subTest(raw=raw), self.assertRaises(ValueError):
                build_procedural_reveal(raw, HTML)

    def test_ancestor_descendant_double_hiding_and_static_conflict_rejected(self):
        for groups, static in [
            ([{'label': 'One', 'ids': ['s002_slide_200', 's002_slide_201']}], []),
            ([{'label': 'One', 'ids': ['s002_slide_200']}, {'label': 'Two', 'ids': ['s002_slide_201']}], []),
            ([{'label': 'One', 'ids': ['s002_slide_200']}], ['s002_slide_201']),
        ]:
            with self.subTest(groups=groups), self.assertRaises(ValueError):
                build_procedural_reveal({'slides': [{'slide': 2, 'groups': groups, 'staticIds': static}]}, HTML)

    def test_slide_and_group_schema_rejected(self):
        cases = [None, [], {}, {'slides': []}, {'slides': [None]}]
        for slide in [0, -1, True, 2.0, '2', 99]:
            raw = plan(); raw['slides'][0]['slide'] = slide; cases.append(raw)
        raw = plan(); raw['slides'].append(copy.deepcopy(raw['slides'][0])); cases.append(raw)
        for groups in [None, [], ['wrong'], [{'label': '', 'ids': ['s002_slide_101']}],
                       [{'label': ' ', 'ids': ['s002_slide_101']}], [{'ids': ['s002_slide_101']}]]:
            raw = plan(); raw['slides'][0]['groups'] = groups; cases.append(raw)
        for raw in cases:
            with self.subTest(raw=raw), self.assertRaises(ValueError):
                build_procedural_reveal(raw, HTML)

    def test_positive_integer_interval_and_duplicate_html_id(self):
        self.assertEqual(positive_interval('300'), 300)
        self.assertEqual(positive_interval(1), 1)
        for value in [0, -1, '0', '-1', '1.5', 1.5, True, None, 'nan', float('inf')]:
            with self.subTest(value=value), self.assertRaises(ValueError):
                positive_interval(value)
        duplicate = HTML.replace('First', obj('s002_slide_101', 'Duplicate'))
        with self.assertRaisesRegex(ValueError, 'duplicate object ID'):
            ObjectInventory(duplicate)

    def test_scheduled_cleanup_cannot_overflow_browser_timer_limit(self):
        # Three groups: the largest actual timer includes the final cleanup.
        _, audit = build_procedural_reveal(plan(), HTML, 1_073_741_498)
        self.assertEqual(audit['maximum_cleanup_ms'], 2_147_483_646)
        with self.assertRaisesRegex(ValueError, 'timer delay limit'):
            build_procedural_reveal(plan(), HTML, 1_073_741_499)


class ProcedureCliTests(unittest.TestCase):
    def setUp(self):
        # Match existing tests' Windows ACL-safe temporary workspace convention.
        self.folder = ROOT / 'tests' / ('.tmp-procedure-' + uuid.uuid4().hex)
        self.folder.mkdir()
        converter = Converter(ROOT / 'examples/sample.pptx')
        try:
            rendered, _, _ = converter.convert()
        finally:
            converter.close()
        inventory = ObjectInventory(rendered)
        for slide, ids in inventory.slides.items():
            leaf_ids = [identity for identity in ids if '_slide_' in identity
                        and not any(identity in other['ancestors'] for other in inventory.objects.values())]
            if len(leaf_ids) >= 4:
                self.manifest = {'slides': [{'slide': slide, 'groups': [
                    {'label': f'Step {index + 1}', 'ids': [identity]} for index, identity in enumerate(leaf_ids[:3])],
                    'conclusionIds': [leaf_ids[3]]}]}
                break
        else:
            self.fail('Synthetic sample must provide four independent slide objects')
        self.mapping = self.folder / 'procedure.json'
        self.mapping.write_text(json.dumps(self.manifest), encoding='utf-8')

    def tearDown(self):
        shutil.rmtree(self.folder)

    def run_cli(self, target, *extra):
        return subprocess.run([sys.executable, str(ROOT / 'scripts/convert.py'),
                               str(ROOT / 'examples/sample.pptx'), '-o', str(target), *extra],
                              capture_output=True, text=True, encoding='utf-8')

    def test_cli_preserves_slide_bytes_and_reports_actual_longest_sequence(self):
        baseline, target = self.folder / 'baseline.html', self.folder / 'procedure.html'
        for file, options in [(baseline, []), (target, ['--procedure-manifest', str(self.mapping)])]:
            result = self.run_cli(file, *options)
            self.assertEqual(result.returncode, 0, result.stderr)
        before, after = baseline.read_text(encoding='utf-8'), target.read_text(encoding='utf-8')
        self.assertEqual(re.findall(r'<section class="ppt-slide".*?</section>', before, re.S),
                         re.findall(r'<section class="ppt-slide".*?</section>', after, re.S))
        audit = json.loads(target.with_suffix('.audit.json').read_text(encoding='utf-8'))
        baseline_audit = json.loads(baseline.with_suffix('.audit.json').read_text(encoding='utf-8'))
        self.assertEqual(audit['html_text_sha256'], baseline_audit['html_text_sha256'])
        self.assertEqual(audit['html_sha256'], hashlib.sha256(target.read_bytes()).hexdigest())
        self.assertEqual(audit['motion']['max_delay_ms'], 500)
        self.assertEqual(audit['motion']['normal_maximum_total_ms'], 920)
        self.assertEqual(audit['motion']['maximum_total_ms'], 1470)
        self.assertEqual(audit['procedural_reveal']['maximum_cleanup_ms'], 1550)
        self.assertFalse(audit['procedural_reveal']['browser_verified'])

    def test_invalid_mapping_and_none_conflict_create_no_output(self):
        invalid = copy.deepcopy(self.manifest)
        invalid['slides'][0]['groups'][0]['ids'] = ['unknown-object']
        bad_mapping = self.folder / 'invalid.json'
        bad_mapping.write_text(json.dumps(invalid), encoding='utf-8')
        cases = [
            ['--procedure-manifest', str(bad_mapping)],
            ['--procedure-manifest', str(self.mapping), '--motion', 'none'],
            ['--procedure-manifest', str(self.mapping), '--procedure-interval-ms', '0'],
            ['--procedure-manifest', str(self.mapping), '--procedure-interval-ms', '1.5'],
        ]
        for index, options in enumerate(cases):
            target = self.folder / f'invalid-{index}.html'
            with self.subTest(options=options):
                result = self.run_cli(target, *options)
                self.assertNotEqual(result.returncode, 0)
                self.assertFalse(target.exists())
                self.assertFalse(target.with_suffix('.audit.json').exists())

    def test_overwrite_cannot_replace_input_mapping_with_html_or_report(self):
        html_mapping = self.folder / 'mapping.html'
        html_mapping.write_text(json.dumps(self.manifest), encoding='utf-8')
        original_html_mapping = html_mapping.read_bytes()
        result = self.run_cli(html_mapping, '--procedure-manifest', str(html_mapping), '--overwrite')
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(html_mapping.read_bytes(), original_html_mapping)
        self.assertFalse(html_mapping.with_suffix('.audit.json').exists())
        original_mapping = self.mapping.read_bytes()
        target = self.folder / 'output.html'
        result = self.run_cli(target, '--procedure-manifest', str(self.mapping),
                              '--report', str(self.mapping), '--overwrite')
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(self.mapping.read_bytes(), original_mapping)
        self.assertFalse(target.exists())


if __name__ == '__main__':
    unittest.main()
