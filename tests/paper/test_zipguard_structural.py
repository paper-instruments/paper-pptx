"""Characterization of the structural ZIP refusals in `pptx._zipguard`.

These pin behavior that must survive the removal of the six numeric resource ceilings.
Every case here targets a refusal living inside a function that removal rewrites:
`_preflight_zip_stream`, `_scan_central_directory`, `_validate_metadata`,
`_inflate_member`, and `_parse_content_types`. The twelve *policy* raise sites (the
ceilings themselves) are deliberately NOT pinned -- they are being deleted on purpose.

Message fragments are copied verbatim from `_zipguard.py`, so these tests fail on message
drift as well as behavior drift.

Coverage map -- refusals covered elsewhere, deliberately not duplicated here:

* duplicate member name -- `test_package_io_hardening.py`
  `::test_normal_open_refuses_duplicate_members`
* noncanonical member name -- `test_package_io_hardening.py`
  `::test_normal_open_refuses_noncanonical_member_names`

Refusals that are unreachable through `Presentation()` and therefore have no case below:

* `_inflate_member` "has truncated stored data" and "has truncated compressed data" both
  require `stream.read()` to return b"" inside a member's data region. `_validate_local_header`
  already refuses any member whose `data_start + compress_size` runs past its boundary, and the
  last member's boundary is `start_dir`, which `_read_all_members` refuses when it exceeds the
  archive size. So the read position is always strictly inside the file and always yields bytes.
* `_scan_central_directory` "ZIP central-directory member record is truncated" needs a short
  read inside `[central_offset, central_end)`, and `central_end` is the EOCD offset, which is
  inside the file by construction. "ends inside a member record" (pinned below) is the
  reachable truncation refusal.
"""

from __future__ import annotations

import re
import struct
import zipfile
import zlib
from pathlib import Path

import pytest

from pptx import Presentation, _zipguard
from pptx.errors import PackageLimitError

# -- ZIP record layouts, mirrored from `_zipguard` so a change there cannot silently
# -- rewrite the archives these tests build.
_LOCAL_HEADER = struct.Struct("<4s5H3L2H")
_CENTRAL_HEADER = struct.Struct("<4s6H3L5H2L")
_END_RECORD = struct.Struct("<4s4H2LH")
_FLAG_UTF8 = 0x0800

_BLOB = b"<paper-pptx characterization blob/>" * 8
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _deflate(blob: bytes) -> bytes:
    compressor = zlib.compressobj(9, zlib.DEFLATED, -zlib.MAX_WBITS)
    return compressor.compress(blob) + compressor.flush()


def _member(name, blob=_BLOB, *, method=zipfile.ZIP_DEFLATED, data=None, **overrides):
    """An honest local+central record pair for `blob`, with `overrides` replacing fields.

    `overrides` are applied to BOTH headers, because `_validate_local_header` refuses any
    local/central disagreement and would otherwise mask the refusal under test.
    """
    if data is None:
        data = _deflate(blob) if method == zipfile.ZIP_DEFLATED else blob
    spec = {
        "name": name,
        "data": data,
        "method": method,
        "flags": _FLAG_UTF8,
        "crc": zlib.crc32(blob) & 0xFFFFFFFF,
        "file_size": len(blob),
        "compress_size": len(data),
        "disk_start": 0,  # -- central-directory "disk number start" field
        "offset": None,  # -- central-directory "relative offset of local header" field
    }
    spec.update(overrides)
    return spec


def _zip_bytes(
    *members,
    disk_number=0,
    central_disk=0,
    disk_entries=None,
    total_entries=None,
    central_size_delta=0,
    central_gap=0,
):
    """Assemble a ZIP whose metadata is honest except where a keyword overrides it.

    The keywords name end-of-central-directory fields; `central_gap` inserts filler bytes
    after the last central-directory record but still inside the declared central region.
    """
    body = bytearray()
    central = bytearray()
    for spec in members:
        offset = len(body) if spec["offset"] is None else spec["offset"]
        raw_name = spec["name"].encode("utf-8")
        shared = (
            spec["flags"],
            spec["method"],
            0,  # -- mod time
            0,  # -- mod date
            spec["crc"],
            spec["compress_size"],
            spec["file_size"],
        )
        body += _LOCAL_HEADER.pack(b"PK\x03\x04", 20, *shared, len(raw_name), 0)
        body += raw_name + spec["data"]
        central += _CENTRAL_HEADER.pack(
            b"PK\x01\x02",
            20,  # -- version made by
            20,  # -- version needed
            *shared,
            len(raw_name),
            0,  # -- extra length
            0,  # -- comment length
            spec["disk_start"],
            0,  # -- internal attributes
            0,  # -- external attributes
            offset,
        )
        central += raw_name
    central += b"\x00" * central_gap

    central_offset = len(body)
    body += central
    body += _END_RECORD.pack(
        b"PK\x05\x06",
        disk_number,
        central_disk,
        len(members) if disk_entries is None else disk_entries,
        len(members) if total_entries is None else total_entries,
        len(central) + central_size_delta,
        central_offset,
        0,  # -- comment length
    )
    return bytes(body)


def _content_types(inner: bytes = b"", prologue: bytes = b"") -> bytes:
    return b'%s<Types xmlns="%s">%s</Types>' % (prologue, _CT_NS.encode(), inner)


_OVERRIDE = b'<Override PartName="/a.xml" ContentType="application/xml"/>'

_CASES = [
    # -- `_inflate_member` (the highest-risk function) ---------------------------------
    pytest.param(
        lambda: _zip_bytes(_member("a.xml", crc=0xDEADBEEF)),
        "fails its CRC check",
        id="inflate-crc-mismatch",
    ),
    pytest.param(
        lambda: _zip_bytes(_member("a.xml", file_size=len(_BLOB) + 1)),
        "does not match declared size",
        id="inflate-expanded-size-mismatch",
    ),
    pytest.param(
        # -- deflate stream cut short of its end-of-stream marker
        lambda: _zip_bytes(_member("a.xml", data=_deflate(_BLOB)[:-2])),
        "compressed data ends before deflate EOF",
        id="inflate-deflate-truncated",
    ),
    pytest.param(
        # -- complete deflate stream followed by bytes the decompressor never consumes
        lambda: _zip_bytes(_member("a.xml", data=_deflate(_BLOB) + b"\x00\x00")),
        "has trailing compressed data",
        id="inflate-trailing-compressed-data",
    ),
    # -- `_validate_metadata` ----------------------------------------------------------
    pytest.param(
        lambda: _zip_bytes(_member("a.xml"), _member("A.xml")),
        "case-ambiguous member name",
        id="metadata-case-ambiguous-name",
    ),
    pytest.param(
        # -- ZipInfo truncates its `filename` at the NUL while `orig_filename` keeps it
        lambda: _zip_bytes(_member("a.xml\x00stowaway")),
        "contains a noncanonical NUL suffix",
        id="metadata-nul-suffix-name",
    ),
    pytest.param(
        # -- an empty member name. `_is_directory_entry` runs before `_validate_metadata`,
        # -- and on Python 3.9/3.10 `ZipInfo.is_dir` reads `filename[-1]`, so an unguarded
        # -- call raised IndexError there instead of this typed refusal.
        lambda: _zip_bytes(_member("")),
        "empty member name",
        id="metadata-empty-member-name",
    ),
    pytest.param(
        lambda: _zip_bytes(_member("a.xml", flags=_FLAG_UTF8 | 0x0001)),
        "is encrypted",
        id="metadata-encrypted-member",
    ),
    pytest.param(
        # -- method 12 is bzip2: a real ZIP method the OPC ZIP mapping does not permit
        lambda: _zip_bytes(_member("a.xml", method=12)),
        "uses unsupported compression method",
        id="metadata-unsupported-compression",
    ),
    pytest.param(
        lambda: _zip_bytes(_member("a.xml"), _member("b.xml", offset=0)),
        "has an invalid or shared local-header offset",
        id="metadata-shared-header-offset",
    ),
    pytest.param(
        lambda: _zip_bytes(_member("a.xml", method=zipfile.ZIP_STORED, file_size=len(_BLOB) + 1)),
        "has inconsistent size metadata",
        id="metadata-stored-size-disagreement",
    ),
    # -- `_preflight_zip_stream` -------------------------------------------------------
    pytest.param(
        lambda: _zip_bytes(_member("a.xml"), disk_number=1),
        "multi-disk ZIP packages are not supported",
        id="preflight-multi-disk-end-record",
    ),
    pytest.param(
        lambda: _zip_bytes(_member("a.xml"), disk_entries=2),
        "ZIP central-directory counts disagree across disks",
        id="preflight-entry-counts-disagree",
    ),
    pytest.param(
        lambda: _zip_bytes(_member("a.xml"), central_size_delta=1),
        "do not identify one unambiguous region",
        id="preflight-central-region-ambiguous",
    ),
    pytest.param(
        lambda: _zip_bytes(_member("a.xml"), disk_entries=0, total_entries=0),
        "empty ZIP package has a non-empty central directory",
        id="preflight-empty-with-central-directory",
    ),
    pytest.param(
        lambda: _zip_bytes(_member("a.xml"), disk_entries=3, total_entries=3),
        "ZIP central directory is too small for its member count",
        id="preflight-central-directory-too-small",
    ),
    # -- `_scan_central_directory` -----------------------------------------------------
    pytest.param(
        lambda: _zip_bytes(_member("a.xml"), central_gap=10),
        "ZIP central directory ends inside a member record",
        id="scan-region-ends-inside-record",
    ),
    pytest.param(
        lambda: _zip_bytes(_member("a.xml", disk_start=1)),
        "multi-disk ZIP member records are not supported",
        id="scan-member-record-nonzero-disk",
    ),
    # -- `_parse_content_types`, reached only through the `_read_all_members` call site --
    pytest.param(
        lambda: _zip_bytes(_member("[Content_Types].xml", _content_types(_OVERRIDE + _OVERRIDE))),
        "[Content_Types].xml contains an ambiguous Override declaration",
        id="content-types-duplicate-override",
    ),
    pytest.param(
        lambda: _zip_bytes(
            _member("[Content_Types].xml", _content_types(prologue=b"<!DOCTYPE Types>"))
        ),
        "[Content_Types].xml contains a prohibited DTD",
        id="content-types-prohibited-dtd",
    ),
]


def test_the_unmutated_archive_builder_output_is_accepted(tmp_path):
    """Guard the guard: each refusal above must come from its mutation, not from `_zip_bytes`."""
    target = tmp_path / "honest.pptx"
    target.write_bytes(
        _zip_bytes(
            # -- an honest archive declares a content type for every part it carries
            _member("[Content_Types].xml", _content_types(_OVERRIDE)),
            _member("a.xml"),
        )
    )

    _zipguard.preflight_zip(str(target))
    with zipfile.ZipFile(target) as archive:
        reader = _zipguard.GuardedZipReader(archive)

    assert reader.order == ("[Content_Types].xml", "a.xml")
    assert reader.read("a.xml") == _BLOB


@pytest.mark.parametrize(("build", "fragment"), _CASES)
def test_normal_open_refuses_structurally_ambiguous_archive(tmp_path, build, fragment):
    target = tmp_path / "adversarial.pptx"
    target.write_bytes(build())

    with pytest.raises(PackageLimitError, match=re.escape(fragment)):
        Presentation(target)


def test_zipguard_structural_refusal_population_is_pinned():
    """`_zipguard.py` has 83 `raise PackageLimitError` sites, every one structural.

    The twelve policy sites that enforced the six numeric resource ceilings are gone. Two
    sites were added, each for a shape PowerPoint refuses: a package carrying bytes in
    front of its first member, and a package whose members do not all resolve to a
    declared content type. One was removed: the directory-entry refusal, since a ZIP
    directory record is not an OPC part and PowerPoint opens packages containing them.
    Any movement from 83 means a structural or ambiguity refusal was added or dropped, so
    a mismatch here is a prompt to check which.
    """
    source = Path(_zipguard.__file__).read_text(encoding="utf-8")

    assert source.count("raise PackageLimitError") == 83
