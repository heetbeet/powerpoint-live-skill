"""Take a fresh PowerPoint snapshot and read one XML fragment in one request."""
from argparse import Namespace
from pathlib import Path
import time

from ppt_package import inspect, positive_id, bounded_max_chars


def snapshot_xml(presentation, request):
    path = Path(request.get('path', ''))
    if not path.is_absolute():
        raise ValueError('xml requires a fresh absolute snapshot path')
    if path.exists():
        raise ValueError('Refusing to overwrite an existing snapshot')
    # Preserve VBA when present; PPTX is the normal non-macro snapshot format.
    macros = bool(presentation.HasVBProject)
    suffix = '.pptm' if macros else '.pptx'
    if path.suffix.lower() != suffix:
        raise ValueError('Snapshot path must end in ' + suffix)
    slide_id = positive_id(str(request['slide_id']))
    shape_id = positive_id(str(request['id'])) if 'id' in request else None
    max_chars = bounded_max_chars(str(request.get('max_chars', 6000)))
    output = request.get('output')
    if output is not None and (not Path(output).is_absolute() or Path(output).exists()):
        raise ValueError('XML output requires a fresh absolute path')
    started = time.perf_counter()
    presentation.SaveCopyAs(str(path), 25 if macros else 24)
    saved_ms = round((time.perf_counter()-started)*1000, 1)
    result = inspect(Namespace(snapshot=str(path), slide_id=slide_id, shape_id=shape_id,
                               max_chars=max_chars, output=output))
    result['snapshot_ms'] = saved_ms
    result['fresh_snapshot'] = True
    return result
