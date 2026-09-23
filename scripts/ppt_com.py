"""Live PowerPoint commands. Construct and call only on the owning STA thread."""
from __future__ import annotations
import hashlib
import json
import math
import ntpath
import time
import uuid
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client


class BridgeError(Exception):
    def __init__(self, code, message, **details):
        super().__init__(message)
        self.code, self.details = code, details


BUSY = {-2147418111, -2147417846}


def canonical_path(value):
    if not isinstance(value, str) or not ntpath.isabs(value):
        return None
    return ntpath.normcase(ntpath.normpath(value))


def image_path(value):
    if not isinstance(value, str):
        raise BridgeError('invalid_request', 'Image path must be an absolute local file path')
    path = Path(value)
    if not path.is_absolute() or not path.is_file():
        raise BridgeError('invalid_request', 'Image path must be an existing absolute local file')
    return str(path)


def _com_same(left, right):
    if left is right:
        return True
    try:
        return left._oleobj_ == right._oleobj_
    except AttributeError:
        return False


def optional(getter):
    try:
        return getter()
    except pywintypes.com_error as e:
        if e.hresult in BUSY:
            raise
        return None
    except AttributeError:
        return None


def items(collection):
    return [collection.Item(i) for i in range(1, collection.Count + 1)]


def number(value, positive=False):
    if isinstance(value, bool) or not isinstance(value, (float, int)) or not math.isfinite(value):
        raise BridgeError('invalid_request', 'Geometry and font sizes must be finite numbers')
    if positive and value <= 0:
        raise BridgeError('invalid_request', 'Widths, heights and font sizes must be positive')
    return value


def rgb(value):
    if not isinstance(value, str):
        raise BridgeError('invalid_request', 'Colors must be #RRGGBB strings')
    h = value.lstrip('#')
    if len(h) != 6:
        raise BridgeError('invalid_request', 'Colors must be #RRGGBB strings')
    try:
        r, g, b = [int(h[i:i+2], 16) for i in (0, 2, 4)]
    except ValueError:
        raise BridgeError('invalid_request', 'Invalid color')
    return r + (g << 8) + (b << 16)


class Bridge:
    def __init__(self):
        self.app = None
        self.decks = {}
        self.scopes = {}

    def _app(self):
        if self.app is None:
            try:
                self.app = win32com.client.GetActiveObject('PowerPoint.Application')
            except pywintypes.com_error:
                raise BridgeError('powerpoint_not_running', 'Open PowerPoint before using this skill')
        try:
            self.app.Presentations.Count
        except pywintypes.com_error:
            self.app = None
            self.decks.clear()
            raise BridgeError('connection_lost', 'PowerPoint connection ended; request deck again')
        return self.app

    def _inventory(self):
        current = items(self._app().Presentations)
        live = {}
        for p in current:
            found = next((k for k, old in self.decks.items() if _com_same(old, p)), None)
            live[found or uuid.uuid4().hex[:12]] = p
        self.decks = live
        return live

    def _full_name(self, p):
        return str(optional(lambda: p.FullName) or p.Name)

    def _presentation_key(self, p):
        path = canonical_path(self._full_name(p))
        if path is not None:
            return ('path', path)
        try:
            return ('com', str(p._oleobj_))
        except AttributeError:
            return ('object', id(p))

    def _deck(self, request):
        decks = self._inventory()
        reference = request.get('deck')
        if not isinstance(reference, str) or not reference:
            raise BridgeError('invalid_request', 'deck must be a current handle or absolute presentation path')
        path_reference = canonical_path(reference)
        for handle, presentation in decks.items():
            if reference == handle or (
                path_reference is not None and path_reference == canonical_path(self._full_name(presentation))
            ):
                return handle, presentation
        raise BridgeError('unknown_deck', 'Use a current deck handle or canonical presentation path')

    def _slide(self, p, request):
        sid = request.get('slide_id')
        if isinstance(sid, bool) or not isinstance(sid, int):
            raise BridgeError('invalid_request', 'slide_id must be a stable integer SlideID')
        for s in items(p.Slides):
            if s.SlideID == sid:
                return s
        raise BridgeError('unknown_slide', 'SlideID no longer exists in the specified deck')

    def _shape_map(self, slide):
        found = {}
        def visit(shapes):
            for shape in items(shapes):
                sid = shape.Id
                if sid in found:
                    raise BridgeError('ambiguous_shape', 'Duplicate shape ID', id=sid)
                found[sid] = shape
                if shape.Type == 6:
                    visit(shape.GroupItems)
        visit(slide.Shapes)
        return found

    def _record(self, shape):
        r = {'id': shape.Id, 'name': shape.Name, 'type': shape.Type,
             'box': [round(float(v), 4) for v in (shape.Left, shape.Top, shape.Width, shape.Height)],
             'rotation': round(float(shape.Rotation), 4), 'z': shape.ZOrderPosition}
        r['fill'] = optional(lambda: [shape.Fill.Type, shape.Fill.Visible, shape.Fill.ForeColor.RGB, shape.Fill.Transparency])
        if shape.HasTextFrame:
            tf = shape.TextFrame2
            tr = tf.TextRange
            r['text'] = tr.Text
            r['font'] = optional(lambda: [tr.Font.Name, tr.Font.Size, tr.Font.Bold, tr.Font.Fill.ForeColor.RGB])
            r['text_frame'] = {'margins': [tf.MarginLeft, tf.MarginTop, tf.MarginRight, tf.MarginBottom],
                               'autosize': tf.AutoSize, 'wrap': tf.WordWrap,
                               'anchor': tf.VerticalAnchor,
                               'alignment': optional(lambda: tr.ParagraphFormat.Alignment)}
            bounds = optional(lambda: [tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight])
            if bounds is not None and r['text']:
                r['text_bounds'] = [round(float(v), 3) for v in bounds]
                margins = r['text_frame']['margins']
                usable = [max(0, r['box'][2]-margins[0]-margins[2]), max(0, r['box'][3]-margins[1]-margins[3])]
                excess = [round(max(0, bounds[2]-usable[0]), 2), round(max(0, bounds[3]-usable[1]), 2)]
                r['overflow'] = {'excess_wh': excess, 'measured_hint': True} if max(excess) > 0.5 else False
        if shape.Type == 6:
            r['children'] = [self._record(c) for c in items(shape.GroupItems)]
        return r

    def _snapshot(self, p, slide, ids=None):
        shapes = items(slide.Shapes)
        ancestors = {}
        if ids is not None:
            if not isinstance(ids,list) or not ids or any(isinstance(i,bool) or not isinstance(i,int) for i in ids):
                raise BridgeError('invalid_request','ids must be a nonempty list of shape IDs')
            wanted = set(ids)
            selected = {}
            def find(collection, parents=()):
                for s in collection:
                    sid = s.Id
                    if sid in wanted:
                        selected[sid] = s
                        for parent in parents:
                            ancestors[parent.Id] = {'id':parent.Id,
                                'box':[parent.Left,parent.Top,parent.Width,parent.Height],
                                'rotation':parent.Rotation}
                    if len(selected)==len(wanted):
                        return
                    if s.Type==6:
                        find(items(s.GroupItems),parents+(s,))
            find(shapes)
            if set(selected)!=wanted:
                raise BridgeError('unknown_shape','Requested shape no longer exists')
            shapes = [selected[i] for i in ids]
        raw = {'slide_id': slide.SlideID, 'size': [p.PageSetup.SlideWidth, p.PageSetup.SlideHeight],
               'shapes': [self._record(s) for s in shapes]}
        if ids is not None:
            raw['revision_scope'] = ids
            if ancestors:
                raw['ancestors'] = list(ancestors.values())
        encoded = json.dumps(raw, sort_keys=True, ensure_ascii=False, separators=(',', ':'))
        raw['revision'] = hashlib.sha256(encoded.encode('utf-8')).hexdigest()[:20]
        return raw

    def _compact(self, raw, request):
        depth = request.get('depth', 0)
        text_limit = request.get('text_limit', 120)
        if not isinstance(depth, int) or not 0 <= depth <= 12 or not isinstance(text_limit, int) or not 0 <= text_limit <= 20000:
            raise BridgeError('invalid_request', 'depth must be 0..12 and text_limit 0..20000')
        def shrink(record, remaining):
            r = {k: v for k, v in record.items() if k not in ('children', 'text_frame', 'text_bounds', 'fill')}
            if 'text' in r and len(r['text']) > text_limit:
                r['text_length'] = len(r['text'])
                r['text'] = r['text'][:text_limit]
                r['text_truncated'] = True
            if request.get('detail'):
                r.update({k: record[k] for k in ('text_frame', 'text_bounds', 'fill') if k in record})
            if 'children' in record:
                r['child_count'] = len(record['children'])
                if remaining:
                    r['children'] = [shrink(c, remaining-1) for c in record['children']]
            return r
        selected = raw['shapes']
        if 'ids' in request:
            ids = request['ids']
            if not isinstance(ids, list):
                raise BridgeError('invalid_request', 'ids must be a list')
            flat = {}
            def collect(rows):
                for r in rows:
                    flat[r['id']] = r
                    collect(r.get('children', []))
            collect(selected)
            if any(i not in flat for i in ids):
                raise BridgeError('unknown_shape', 'A requested shape no longer exists')
            selected = [flat[i] for i in ids]
        return {**{k: v for k, v in raw.items() if k != 'shapes'}, 'shapes': [shrink(s, depth) for s in selected]}

    def _guard(self, p, slide, request):
        scope = self.scopes.get((self._presentation_key(p), slide.SlideID, request.get('expected_revision')))
        if scope and request.get('op')=='native':
            raise BridgeError('full_revision_required','Native write scripts require a full slide inspection')
        before = self._snapshot(p, slide, scope)
        if request.get('expected_revision') != before['revision']:
            if scope:
                current = self._compact(before, request)
                raise BridgeError('stale_revision', 'Inspect again before editing; the live revision differs',
                                  revision=before['revision'], current=current)
            raise BridgeError('stale_revision', 'Inspect again before editing; the live revision differs',
                              revision=before['revision'])
        return before

    def _validate_ops(self, operations, shape_map, top_level_ids):
        if not isinstance(operations, list) or not 1 <= len(operations) <= 200:
            raise BridgeError('invalid_request', 'operations must contain 1..200 edits')
        deleted = set()
        set_fields = {'x', 'y', 'w', 'h', 'text', 'font_size', 'font_name', 'bold', 'fill', 'color', 'name'}
        for op in operations:
            if not isinstance(op, dict):
                raise BridgeError('invalid_request', 'Each operation must be an object')
            kind = op.get('op')
            allowed = {'set': {'id'} | set_fields, 'replace_text': {'id', 'old', 'new'},
                       'add_text': set_fields, 'add_rect': set_fields, 'align': {'ids', 'mode'},
                       'distribute': {'ids', 'direction'}, 'duplicate': {'id'}, 'delete': {'id'},
                       'add_image': {'path', 'x', 'y', 'w', 'h', 'name'},
                       'replace_image': {'id', 'path'}}
            if kind not in allowed or set(op)-({'op'} | allowed[kind]):
                raise BridgeError('invalid_request', 'Unknown operation or field', operation=kind)
            refs = op.get('ids', [op.get('id')]) if kind not in ('add_text', 'add_rect', 'add_image') else []
            if not isinstance(refs, list) or any(isinstance(i,bool) or not isinstance(i,int) or i not in shape_map or i in deleted for i in refs):
                raise BridgeError('unknown_shape', 'Batch refers to a missing or already deleted shape')
            if kind in ('align', 'distribute') and (len(refs) < 2 or len(set(refs)) != len(refs)):
                raise BridgeError('invalid_request', 'Alignment requires distinct shape IDs')
            if kind == 'align' and op.get('mode') not in ('left', 'right', 'top', 'bottom', 'center', 'middle'):
                raise BridgeError('invalid_request', 'Unknown alignment mode')
            if kind == 'distribute' and op.get('direction') not in ('horizontal', 'vertical'):
                raise BridgeError('invalid_request', 'Unknown distribution direction')
            if kind in ('add_text', 'add_rect') and not {'x', 'y', 'w', 'h'} <= set(op):
                raise BridgeError('invalid_request', 'New shapes require x,y,w,h')
            if kind == 'add_image' and not {'path', 'x', 'y', 'w', 'h'} <= set(op):
                raise BridgeError('invalid_request', 'add_image requires path,x,y,w,h')
            if kind == 'replace_image' and shape_map[op['id']].Type not in (11, 13):
                raise BridgeError('invalid_request', 'replace_image target must be a picture')
            if kind == 'replace_image' and op['id'] not in top_level_ids:
                raise BridgeError('invalid_request', 'replace_image does not support pictures nested in groups')
            if kind in ('add_image', 'replace_image'):
                op['path'] = image_path(op['path'])
            for key in ('x', 'y', 'w', 'h', 'font_size'):
                if key in op:
                    number(op[key], key in ('w', 'h', 'font_size'))
            for key in ('text', 'font_name', 'name', 'old', 'new'):
                if key in op and not isinstance(op[key], str):
                    raise BridgeError('invalid_request', key+' must be a string')
            if kind == 'replace_text' and (not op.get('old') or 'new' not in op):
                raise BridgeError('invalid_request', 'replace_text requires nonempty old and new')
            if 'bold' in op and not isinstance(op['bold'], bool):
                raise BridgeError('invalid_request', 'bold must be boolean')
            for key in ('fill', 'color'):
                if key in op:
                    rgb(op[key])
            if kind in ('set', 'replace_text') and any(k in op for k in ('text','font_size','font_name','bold','color','old')):
                if not shape_map[op['id']].HasTextFrame:
                    raise BridgeError('invalid_request', 'Target has no text frame')
            if kind in ('delete', 'replace_image'):
                deleted.add(op['id'])

    def _set(self, shape, op):
        if 'w' in op and 'h' in op:
            lock = shape.LockAspectRatio
            try:
                shape.LockAspectRatio = 0
                shape.Width, shape.Height = op['w'], op['h']
            finally:
                shape.LockAspectRatio = lock
        else:
            for key, attr in (('w', 'Width'), ('h', 'Height')):
                if key in op:
                    setattr(shape, attr, op[key])
        for key, attr in (('x', 'Left'), ('y', 'Top'), ('name', 'Name')):
            if key in op:
                setattr(shape, attr, op[key])
        if 'text' in op:
            shape.TextFrame.TextRange.Text = op['text']
        if any(k in op for k in ('font_size', 'font_name', 'bold', 'color')):
            font = shape.TextFrame.TextRange.Font
            for key, attr in (('font_size','Size'), ('font_name','Name')):
                if key in op:
                    setattr(font, attr, op[key])
            if 'bold' in op:
                font.Bold = -1 if op['bold'] else 0
            if 'color' in op:
                font.Color.RGB = rgb(op['color'])
        if 'fill' in op:
            shape.Fill.Visible = -1
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = rgb(op['fill'])

    def _add_picture(self, slide, path, x, y, w, h):
        return slide.Shapes.AddPicture(path, 0, -1, x, y, w, h)

    def _replace_picture(self, slide, old, op):
        new = None
        try:
            new = self._add_picture(slide, op['path'], old.Left, old.Top, old.Width, old.Height)
            new.Rotation = old.Rotation
            new.Name = old.Name
            target_z = int(old.ZOrderPosition)
            for _ in range(max(0, int(slide.Shapes.Count) + 1)):
                if int(new.ZOrderPosition) <= target_z:
                    break
                new.ZOrder(3)
            old.Delete()
            return new
        except Exception:
            if new is not None:
                try:
                    new.Delete()
                except Exception:
                    pass
            raise

    def _batch(self, p, slide, request):
        before = self._guard(p, slide, request)
        shape_map = self._shape_map(slide)
        top_level_ids = {shape.Id for shape in items(slide.Shapes)}
        operations = request.get('operations')
        self._validate_ops(operations, shape_map, top_level_ids)
        if 'revision_scope' in before:
            covered = set()
            def cover(rows):
                for row in rows:
                    covered.add(row['id'])
                    cover(row.get('children',[]))
            cover(before['shapes'])
            for op in operations:
                refs = op.get('ids',[op['id']] if 'id' in op else [])
                if not set(refs)<=covered:
                    raise BridgeError('scope_mismatch','Inspect every existing shape the batch will edit')
        self.app.StartNewUndoEntry()
        applied, changed, deleted = [], set(), []
        for index, op in enumerate(operations):
            try:
                kind = op['op']
                shape = shape_map.get(op.get('id'))
                if kind == 'set':
                    self._set(shape, op)
                    changed.add(shape.Id)
                elif kind in ('add_text', 'add_rect'):
                    shape = slide.Shapes.AddTextbox(1,op['x'],op['y'],op['w'],op['h']) if kind == 'add_text' else slide.Shapes.AddShape(1,op['x'],op['y'],op['w'],op['h'])
                    changed.add(shape.Id)
                    if shape.HasTextFrame:
                        shape.TextFrame2.AutoSize = 0
                        shape.TextFrame2.WordWrap = -1
                    self._set(shape, op)
                elif kind == 'add_image':
                    shape = self._add_picture(slide, op['path'], op['x'], op['y'], op['w'], op['h'])
                    if 'name' in op:
                        shape.Name = op['name']
                    changed.add(shape.Id)
                elif kind == 'replace_image':
                    old_id = shape.Id
                    shape = self._replace_picture(slide, shape, op)
                    deleted.append(old_id)
                    changed.add(shape.Id)
                elif kind == 'replace_text':
                    found = shape.TextFrame.TextRange.Replace(op['old'], op['new'], 0, -1, 0)
                    if found is None:
                        raise BridgeError('text_not_found', 'Exact text was not found')
                    changed.add(shape.Id)
                elif kind == 'duplicate':
                    changed.add(shape.Duplicate().Item(1).Id)
                elif kind == 'delete':
                    deleted.append(shape.Id)
                    shape.Delete()
                else:
                    group = [shape_map[i] for i in op['ids']]
                    if kind == 'align':
                        mode = op['mode']
                        horizontal = mode in ('left','right','center')
                        pos, size = ('Left','Width') if horizontal else ('Top','Height')
                        lo = min(getattr(s,pos) for s in group)
                        hi = max(getattr(s,pos)+getattr(s,size) for s in group)
                        for s in group:
                            value = lo if mode in ('left','top') else hi-getattr(s,size) if mode in ('right','bottom') else (lo+hi-getattr(s,size))/2
                            setattr(s,pos,value)
                    else:
                        pos,size = ('Left','Width') if op['direction']=='horizontal' else ('Top','Height')
                        group.sort(key=lambda s:getattr(s,pos))
                        start = getattr(group[0],pos)
                        end = getattr(group[-1],pos)+getattr(group[-1],size)
                        gap = (end-start-sum(getattr(s,size) for s in group))/(len(group)-1)
                        cursor = start
                        for s in group:
                            setattr(s,pos,cursor)
                            cursor += getattr(s,size)+gap
                    changed.update(op['ids'])
                applied.append({'index': index, 'op': kind})
            except Exception as exc:
                revision = optional(lambda: self._snapshot(p,slide)['revision'])
                raise BridgeError('partial_batch', str(exc), applied=applied, failed_index=index,
                                  failing_operation_may_be_partial=True, revision=revision)
        # Return just changed objects. The new scoped revision can guard the
        # next edit of those objects without reading every other shape.
        remaining = list(changed-set(deleted))
        after = self._snapshot(p,slide,remaining or None)
        self._remember(p, after)
        response = self._compact(after, {'ids': list(changed-set(deleted))})
        response.update(applied_count=len(applied), deleted_ids=deleted)
        return response

    def _remember(self, p, snapshot):
        scope = snapshot.get('revision_scope')
        if scope is not None:
            self.scopes[(self._presentation_key(p), snapshot['slide_id'], snapshot['revision'])] = scope
            if len(self.scopes) > 256:
                self.scopes.pop(next(iter(self.scopes)))

    def _selected_ids(self, selection):
        if selection is None:
            return []
        return optional(lambda: [shape.Id for shape in items(selection.ShapeRange)]) or []

    def _context(self, request):
        decks = self._inventory()
        active = optional(lambda: self.app.ActivePresentation)
        handle = next((key for key, p in decks.items() if active is not None and _com_same(active, p)), None)
        window = optional(lambda: self.app.ActiveWindow)
        slide = optional(lambda: window.View.Slide) if window else None
        if handle is None or slide is None:
            raise BridgeError('no_active_window', 'PowerPoint has no active presentation window')
        presentation = decks[handle]
        selection = optional(lambda: window.Selection)
        selected_ids = self._selected_ids(selection)
        snapshot = self._snapshot(presentation, slide, selected_ids or None)
        self._remember(presentation, snapshot)
        return {
            'deck': handle,
            'path': self._full_name(presentation),
            'slide_id': slide.SlideID,
            'slide_index': slide.SlideIndex,
            'selected_ids': selected_ids,
            'selection_type': optional(lambda: selection.Type) if selection else None,
            'inspection': self._compact(snapshot, request),
        }

    def _outline(self, p, request, handle):
        text_limit = request.get('text_limit', 100)
        if (isinstance(text_limit, bool) or not isinstance(text_limit, int)
                or not 0 <= text_limit <= 20000):
            raise BridgeError('invalid_request', 'text_limit must be an integer from 0 to 20000')

        def walk(collection):
            for shape in items(collection):
                yield shape
                if shape.Type == 6:
                    yield from walk(shape.GroupItems)

        rows = []
        for slide in items(p.Slides):
            total_chars = 0
            meaningful = []
            placeholders = []
            order = 0
            for shape in walk(slide.Shapes):
                text = optional(lambda shape=shape: shape.TextFrame2.TextRange.Text) if optional(
                    lambda shape=shape: bool(shape.HasTextFrame)) else None
                if text:
                    total_chars += len(text)
                    value = ' '.join(text.replace('\r', ' ').replace('\n', ' ').split())
                    if value:
                        top = optional(lambda shape=shape: float(shape.Top))
                        left = optional(lambda shape=shape: float(shape.Left))
                        position = (top if top is not None else float('inf'),
                                    left if left is not None else float('inf'), order)
                        meaningful.append((position, value))
                        placeholder_type = optional(lambda shape=shape: int(shape.PlaceholderFormat.Type)) if shape.Type == 14 else None
                        if placeholder_type in (1, 3):
                            placeholders.append((position, value))
                order += 1
            title = (min(placeholders, key=lambda row: row[0])[1] if placeholders else
                     min(meaningful, key=lambda row: row[0])[1] if meaningful else None)
            rows.append({
                'slide_id': slide.SlideID,
                'index': slide.SlideIndex,
                'shape_count': slide.Shapes.Count,
                'title': title[:text_limit] if title is not None else None,
                'text_char_count': total_chars,
            })
        return {'deck': handle, 'path': self._full_name(p), 'slides': rows}

    def _slide_ids(self, value, available, field):
        if not isinstance(value, list):
            raise BridgeError('invalid_request', field + ' must be a list of slide IDs')
        result = []
        for raw in value:
            if isinstance(raw, bool) or not isinstance(raw, int) or raw not in available or raw in result:
                raise BridgeError('unknown_slide', field + ' contains an invalid or duplicate SlideID')
            result.append(raw)
        return result

    def _native_deck(self, p, request):
        from ppt_native import handle

        mode = request.get('mode')
        if mode not in ('read', 'write'):
            raise BridgeError('invalid_mode', "native_deck mode must be explicitly 'read' or 'write'")
        slides = {int(slide.SlideID): slide for slide in items(p.Slides)}
        expected = request.get('expected_revisions')
        listed = []
        if mode == 'write':
            if not isinstance(expected, dict) or not expected:
                raise BridgeError('expected_revisions_required',
                                  'native_deck writes require expected_revisions for every touched slide')
            normalized = {}
            for raw_id, revision in expected.items():
                try:
                    slide_id = int(raw_id)
                except (TypeError, ValueError):
                    raise BridgeError('invalid_request', 'expected_revisions keys must be SlideIDs')
                if isinstance(raw_id, bool) or slide_id in normalized or slide_id not in slides:
                    raise BridgeError('unknown_slide', 'expected_revisions contains an invalid SlideID')
                if not isinstance(revision, str) or not revision:
                    raise BridgeError('invalid_request', 'expected_revisions values must be full-slide revisions')
                normalized[slide_id] = revision
            listed = list(normalized)
            for slide_id in listed:
                current = self._snapshot(p, slides[slide_id])['revision']
                if normalized[slide_id] != current:
                    raise BridgeError('stale_revision', 'Inspect the slide again before native_deck write',
                                      slide_id=slide_id, revision=current)
            self.app.StartNewUndoEntry()
        else:
            raw_ids = request.get('slide_ids')
            listed = self._slide_ids(list(slides) if raw_ids is None else raw_ids, slides, 'slide_ids')

        shape_maps = {}
        listed_set = set(listed)

        def resolve(slide_id, shape_id):
            if isinstance(slide_id, bool) or not isinstance(slide_id, int) or slide_id not in slides:
                raise BridgeError('unknown_slide', 'Shape resolver received an unknown SlideID')
            if slide_id not in listed_set:
                raise BridgeError('scope_mismatch', 'Shape resolver received an undeclared SlideID')
            if slide_id not in shape_maps:
                shape_maps[slide_id] = self._shape_map(slides[slide_id])
            if shape_id not in shape_maps[slide_id]:
                raise BridgeError('unknown_shape', 'Shape ID not found', slide_id=slide_id, id=shape_id)
            return shape_maps[slide_id][shape_id]

        try:
            declared_slides = {slide_id: slides[slide_id] for slide_id in listed}
            result = handle(request, self.app, p, None, None, slides=declared_slides,
                            deck_shape_resolver=resolve)
            result['revisions'] = {str(slide_id): self._snapshot(p, slides[slide_id])['revision']
                                   for slide_id in listed}
            return result
        except Exception as exc:
            raise BridgeError('native_partial', str(exc), outcome_unknown=True,
                              declared_mode=mode, cause=getattr(exc, 'details', None))

    def _review(self, p, request):
        from ppt_review import ReviewError, review_slides

        all_slides = {int(slide.SlideID): slide for slide in items(p.Slides)}
        raw_ids = request.get('slide_ids')
        slide_ids = list(all_slides) if raw_ids is None else self._slide_ids(raw_ids, all_slides, 'slide_ids')
        try:
            return review_slides(p, [all_slides[slide_id] for slide_id in slide_ids], request)
        except ReviewError as exc:
            raise BridgeError(exc.code, str(exc), **(exc.details or {}))

    def _handle(self, request):
        op = request.get('op')
        if op == 'help':
            return {'operations':['deck','selection','context','inspect','outline','batch','check','review','render','snapshot','save','native_read','native','native_deck','navigate','render_shape','fit_text','xml','metrics'],
                    'units':'points', 'ids':'deck handle + SlideID + shape ID', 'batch':'expected_revision required'}
        if op in ('deck','selection'):
            decks = self._inventory()
            active = optional(lambda: self.app.ActivePresentation)
            active_id = next((k for k,p in decks.items() if active is not None and _com_same(active, p)), None)
            if op == 'deck':
                active_slide = optional(lambda: self.app.ActiveWindow.View.Slide) if active_id else None
                active_slide_id = active_slide.SlideID if active_slide else None
                rows = []
                for handle, p in decks.items():
                    row = {'deck': handle, 'name': p.Name, 'path': self._full_name(p),
                           'slide_count': p.Slides.Count,
                           'size': [p.PageSetup.SlideWidth, p.PageSetup.SlideHeight],
                           'active_slide_id': active_slide_id if handle == active_id else None}
                    if request.get('detail'):
                        row['slides'] = [{'slide_id': s.SlideID, 'index': s.SlideIndex,
                                          'shape_count': s.Shapes.Count} for s in items(p.Slides)]
                    rows.append(row)
                return {'active': active_id, 'decks': rows}
            window = optional(lambda:self.app.ActiveWindow)
            slide = optional(lambda:window.View.Slide) if window else None
            selection = optional(lambda:window.Selection) if window else None
            ids = self._selected_ids(selection)
            return {'deck':active_id,'slide_id':slide.SlideID if slide else None,'ids':ids,
                    'selection_type':selection.Type if selection else None}
        if op == 'context':
            return self._context(request)
        p_handle, p = self._deck(request)
        if op == 'outline':
            return self._outline(p, request, p_handle)
        if op == 'native_deck':
            return self._native_deck(p, request)
        if op == 'review':
            return self._review(p, request)
        if op == 'save':
            p.Save()
            return {'saved':True,'deck':p_handle,'path':self._full_name(p)}
        if op == 'snapshot':
            path = Path(request.get('path',''))
            suffix = '.pptm' if p.HasVBProject else '.pptx'
            if not path.is_absolute() or path.exists() or path.suffix.lower()!=suffix:
                raise BridgeError('invalid_path','Use a fresh absolute '+suffix+' path')
            p.SaveCopyAs(str(path),25 if p.HasVBProject else 24)
            return {'path':str(path),'deck':p_handle}
        slide = self._slide(p,request)
        if op == 'inspect':
            snapshot = self._snapshot(p,slide,request.get('ids'))
            self._remember(p,snapshot)
            return self._compact(snapshot,request)
        if op == 'batch':
            return self._batch(p,slide,request)
        if op == 'check':
            from ppt_layout import check_layout
            return check_layout(self._snapshot(p,slide))
        if op == 'render':
            path = Path(request.get('path',''))
            width = request.get('width',1280)
            if not path.is_absolute() or path.exists() or path.suffix.lower()!='.png':
                raise BridgeError('invalid_path','Use a fresh absolute .png path')
            if isinstance(width,bool) or not isinstance(width,int) or not 64<=width<=8192:
                raise BridgeError('invalid_request','width must be 64..8192 pixels')
            height = round(width*p.PageSetup.SlideHeight/p.PageSetup.SlideWidth)
            slide.Export(str(path),'PNG',width,height)
            return {'path':str(path),'width':width,'height':height,'slide_id':slide.SlideID,
                    'image_hash':hashlib.sha256(path.read_bytes()).hexdigest()[:20]}
        if op == 'xml':
            from ppt_xml_live import snapshot_xml
            return snapshot_xml(p,request)
        if op in ('native_read','native','navigate','render_shape','fit_text'):
            from ppt_native import handle
            write = op=='fit_text' or op=='native' and request.get('mode')=='write'
            if write:
                before = self._guard(p,slide,request)
                if op=='fit_text' and before.get('revision_scope') and request.get('id') not in before['revision_scope']:
                    raise BridgeError('scope_mismatch','Inspect the shape before fitting its text')
                self.app.StartNewUndoEntry()
            shape_map = self._shape_map(slide)
            def resolve(sid):
                if sid not in shape_map:
                    raise BridgeError('unknown_shape','Shape ID not found',id=sid)
                return shape_map[sid]
            try:
                result = handle(request,self.app,p,slide,resolve)
                if write:
                    snapshot = self._snapshot(p,slide,[request['id']] if op=='fit_text' else None)
                    self._remember(p,snapshot)
                    result['revision'] = snapshot['revision']
                    if 'revision_scope' in snapshot:
                        result['revision_scope'] = snapshot['revision_scope']
                return result
            except Exception as exc:
                if write or op == 'native':
                    raise BridgeError('native_partial',str(exc),outcome_unknown=True,
                                      declared_mode=request.get('mode'),
                                      cause=getattr(exc,'details',None))
                raise
        raise BridgeError('unknown_operation','Unknown operation; use help')

    def handle(self, request):
        if not isinstance(request,dict):
            raise BridgeError('invalid_request','Request must be an object')
        # Only repeat read operations rejected as busy. Never replay mutation bodies.
        deadline = time.monotonic()+1.0
        while True:
            try:
                return self._handle(request)
            except pywintypes.com_error as exc:
                if exc.hresult in BUSY and request.get('op') in ('deck','selection','context','inspect','outline','check','native_read') and time.monotonic()<deadline:
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.05)
                    continue
                raise BridgeError('powerpoint_busy' if exc.hresult in BUSY else 'com_error',str(exc),hresult=exc.hresult)
