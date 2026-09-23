"""Conservative geometry hints from a compact PowerPoint slide inspection."""
import math


def check_layout(slide):
    width, height = slide['size']
    shapes = slide.get('shapes', [])
    outside, overlaps, contained, overflow, unknown = [], [], [], [], []
    boxes = []
    for shape in shapes:
        box = shape.get('box')
        if not box or len(box) != 4 or not all(isinstance(n, (int, float)) and math.isfinite(n) for n in box):
            unknown.append({'id': shape.get('id'), 'reason': 'missing geometry'})
            continue
        x, y, w, h = box
        rotation = shape.get('rotation', 0) or 0
        if abs(rotation % 180) > 0.01:
            unknown.append({'id': shape['id'], 'reason': 'rotated shape; rectangle checks omitted'})
            continue
        edges = [max(0, -x), max(0, -y), max(0, x + w - width), max(0, y + h - height)]
        if max(edges) > 0.5:
            outside.append({'id': shape['id'], 'excess_ltrb': [round(v, 2) for v in edges]})
        boxes.append((shape['id'], x, y, x + w, y + h))
        if shape.get('overflow'):
            overflow.append({'id': shape['id'], 'estimate': shape['overflow']})
    for i, a in enumerate(boxes):
        for b in boxes[i + 1:]:
            dx, dy = min(a[3], b[3]) - max(a[1], b[1]), min(a[4], b[4]) - max(a[2], b[2])
            if min(dx, dy) <= 0.5:
                continue
            entry = {'ids': [a[0], b[0]], 'intersection': [round(dx, 2), round(dy, 2)]}
            a_contains = a[1] <= b[1] and a[2] <= b[2] and a[3] >= b[3] and a[4] >= b[4]
            b_contains = b[1] <= a[1] and b[2] <= a[2] and b[3] >= a[3] and b[4] >= a[4]
            (contained if a_contains or b_contains else overlaps).append(entry)
    return {'slide_id': slide['slide_id'], 'revision': slide['revision'],
            'outside': outside, 'overlap_candidates': overlaps[:40],
            'overlap_count': len(overlaps), 'containment_count': len(contained),
            'text_overflow_hints': overflow, 'unknown': unknown,
            'limits': 'Top-level unrotated rectangles only. Overlap and text bounds are hints, not visual defects. Containment usually includes backgrounds. Groups, connectors, effects, masks, and inherited layout/master content need targeted inspection or rendering.'}
