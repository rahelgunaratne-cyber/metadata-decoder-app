"""
Vendored Metadata Sheet Decoder engine.

These modules (scan_metadata, apply_corrections, apply_isrc_corrections,
apply_missing_corrections, apply_format_corrections, threaded_comments) are
COPIED verbatim from the original desktop tool in `Metadata Sheet Decoder/`,
with one small adaptation: `scan_metadata.analyze()` now accepts a
`project_dir` argument so the web backend can point LEAVE-record lookups at a
directory it controls (materialized from Firestore) instead of the module's
own folder.

The original modules use absolute imports like `from scan_metadata import ...`.
To keep those working without rewriting the vendored files, we add this
package directory to `sys.path` so each module is importable as a top-level
module (e.g. `import scan_metadata`).
"""
from __future__ import annotations

import os
import sys

_ENGINE_DIR = os.path.dirname(os.path.abspath(__file__))
if _ENGINE_DIR not in sys.path:
    sys.path.insert(0, _ENGINE_DIR)
