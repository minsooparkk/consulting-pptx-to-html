"""Validate explicit reveal groups against actual converted slide objects.

The mapping changes presentation behavior only. The rendered slide HTML is never
rewritten here, and every unmapped object remains visible by default.
"""
from __future__ import annotations

from html.parser import HTMLParser

MAX_TIMER_DELAY_MS = 2_147_483_647


class ObjectInventory(HTMLParser):
    """Collect object ownership and ancestry without a browser or dependencies."""

    VOID = {'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input',
            'link', 'meta', 'param', 'source', 'track', 'wbr'}

    def __init__(self, source):
        super().__init__(convert_charrefs=True)
        self.stack = []
        self.slides = {}
        self.objects = {}
        self.feed(source)
        self.close()

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        classes = attrs.get('class', '').split()
        slide = self.stack[-1]['slide'] if self.stack else None
        if tag == 'section' and 'ppt-slide' in classes:
            value = attrs.get('data-slide-index', '')
            if not value.isdigit() or int(value) < 1:
                raise ValueError('Generated slide has no valid data-slide-index')
            slide = int(value)
            if slide in self.slides:
                raise ValueError(f'Generated HTML contains duplicate slide {slide}')
            self.slides[slide] = []
        object_id = None
        if 'ppt-object' in classes:
            object_id = attrs.get('data-object-id')
            if slide is None or not object_id or attrs.get('id') != object_id:
                raise ValueError('Generated ppt-object has invalid slide ownership or ID')
            if object_id in self.objects:
                raise ValueError(f'Generated HTML contains duplicate object ID: {object_id}')
            ancestors = tuple(item['object_id'] for item in self.stack if item['object_id'])
            self.objects[object_id] = {'slide': slide, 'ancestors': ancestors}
            self.slides[slide].append(object_id)
        if tag not in self.VOID:
            self.stack.append({'tag': tag, 'slide': slide, 'object_id': object_id})

    def handle_endtag(self, tag):
        for index in range(len(self.stack) - 1, -1, -1):
            if self.stack[index]['tag'] == tag:
                del self.stack[index:]
                break

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in self.VOID:
            self.handle_endtag(tag)


def positive_interval(value):
    """argparse-compatible positive integer parser (not a float or boolean)."""
    if isinstance(value, bool):
        raise ValueError('Procedure interval must be a positive integer')
    if isinstance(value, int):
        interval = value
    elif isinstance(value, str) and value.isascii() and value.isdigit():
        interval = int(value)
    else:
        raise ValueError('Procedure interval must be a positive integer')
    if interval <= 0:
        raise ValueError('Procedure interval must be a positive integer')
    return interval


def build_procedural_reveal(data, rendered_slides, interval_ms=300):
    """Return normalized player configuration and timing audit, or raise early."""
    interval_ms = positive_interval(interval_ms)
    if not isinstance(data, dict) or not isinstance(data.get('slides'), list) or not data['slides']:
        raise ValueError('Procedure manifest must contain a non-empty slides array')
    inventory = ObjectInventory(rendered_slides)
    normalized, timings, seen_slides = [], [], set()

    for item in data['slides']:
        if not isinstance(item, dict):
            raise ValueError('Each procedure slide must be an object')
        slide = item.get('slide')
        if isinstance(slide, bool) or not isinstance(slide, int) or slide < 1:
            raise ValueError('Procedure slide must be a positive integer')
        if slide not in inventory.slides:
            raise ValueError(f'Procedure slide does not exist: {slide}')
        if slide in seen_slides:
            raise ValueError(f'Duplicate procedure slide: {slide}')
        seen_slides.add(slide)
        if not isinstance(item.get('groups'), list) or not item['groups']:
            raise ValueError(f'Procedure slide {slide} must have non-empty groups')
        assigned = set()

        def read_ids(values, field, allow_empty=False):
            if not isinstance(values, list) or (not allow_empty and not values):
                raise ValueError(f'Slide {slide} {field} must be a {"" if allow_empty else "non-empty "}ID array')
            for object_id in values:
                if not isinstance(object_id, str) or not object_id:
                    raise ValueError(f'Slide {slide} {field} contains an invalid object ID')
                if object_id not in inventory.objects:
                    raise ValueError(f'Unknown procedure object ID: {object_id}')
                if inventory.objects[object_id]['slide'] != slide:
                    raise ValueError(f'Procedure object belongs to another slide: {object_id}')
                if object_id in assigned:
                    raise ValueError(f'Procedure object assigned more than once: {object_id}')
                assigned.add(object_id)
            return list(values)

        groups = []
        for index, group in enumerate(item['groups'], 1):
            if not isinstance(group, dict) or not isinstance(group.get('label'), str) or not group['label'].strip():
                raise ValueError(f'Slide {slide} group {index} needs a non-empty text label')
            groups.append({'label': group['label'], 'ids': read_ids(group.get('ids'), f'group {index}')})
        conclusion = read_ids(item.get('conclusionIds', []), 'conclusionIds', allow_empty=True)
        dynamic = set(assigned)
        explicit_static = read_ids(item.get('staticIds', []), 'staticIds', allow_empty=True)
        for object_id in dynamic:
            overlap = dynamic.intersection(inventory.objects[object_id]['ancestors'])
            if overlap:
                raise ValueError(f'Procedure objects would hide both ancestor and descendant: {sorted(overlap)[0]}, {object_id}')
        for object_id in explicit_static:
            overlap = dynamic.intersection(inventory.objects[object_id]['ancestors'])
            if overlap:
                raise ValueError(f'Static object would be hidden by a procedure ancestor: {object_id}')
        # A selected group container already owns its descendants' visibility.
        # They must not also be animated or described as independently static.
        covered = dynamic | {object_id for object_id in inventory.slides[slide]
                             if dynamic.intersection(inventory.objects[object_id]['ancestors'])}
        static = [object_id for object_id in inventory.slides[slide] if object_id not in covered]
        normalized.append({'slide': slide, 'groups': groups,
                           'conclusionIds': conclusion, 'staticIds': static})
        total_groups = len(groups) + bool(conclusion)
        total_ms = 150 + (total_groups - 1) * interval_ms + 420
        if total_ms + 80 > MAX_TIMER_DELAY_MS:
            raise ValueError(f'Slide {slide} procedure exceeds the JavaScript timer delay limit ({MAX_TIMER_DELAY_MS} ms)')
        timings.append({'slide': slide, 'step_groups': len(groups),
                        'conclusion_groups': int(bool(conclusion)), 'total_groups': total_groups,
                        'maximum_total_ms': total_ms, 'maximum_cleanup_ms': total_ms + 80})

    config = {'mode': 'auto', 'interval_ms': interval_ms, 'initial_delay_ms': 150,
              'duration_ms': 420, 'scheduler': 'cumulative-timers',
              'completion_clears_hidden_state': True, 'slides': normalized}
    audit = {**config, 'maximum_total_ms': max(item['maximum_total_ms'] for item in timings),
             'maximum_cleanup_ms': max(item['maximum_cleanup_ms'] for item in timings),
             'slide_timings': timings, 'browser_verified': False}
    return config, audit
