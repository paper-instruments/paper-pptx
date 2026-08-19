"""End-of-central-directory handling, pinned against the real PowerPoint application.

PowerPoint requires the ZIP to span its file exactly: nothing before the first local
header, and the footer's declared comment length accounting for every byte after the
footer. A package that breaks either half opens in stdlib `zipfile`, in upstream
python-pptx, and in LibreOffice, and is refused by PowerPoint with a repair prompt.

That makes these refusals the fork's core case rather than over-strictness: a deck that
opens in python-pptx but not in PowerPoint is the silent-corruption class named in
`PLAN-paper-pptx.md`. The tests below pin each half, and record where paper-pptx and
PowerPoint still disagree.

Measured by opening each shape in PowerPoint (macOS) and recording the result:

    shape                                   PowerPoint  upstream  paper-pptx
    unmutated deck                          opens       opens     opens
    archive comment, declared correctly     opens       opens     opens
    trailing bytes, declared in the footer  opens       opens     opens
    stray footer signature inside a member  opens       opens     opens
    bytes appended, undeclared              REFUSES     opens     refuses
    one newline appended                    REFUSES     opens     refuses
    comment length declared, no comment     REFUSES     opens     refuses
    two packages concatenated               REFUSES     opens     refuses
    prefix data before the archive          REFUSES     opens     refuses
    ZIP directory entries ("ppt/")          opens       opens     refuses  <- still too strict

The prefix-data row was the one shape where this package was more permissive than the
application it protects; it is refused as of this file's companion change. The remaining
disagreement is `directory_entries`, which is a member-name question rather than a footer
one and is handled elsewhere.
"""

from __future__ import annotations

import struct
import zipfile

import pytest

from pptx import Presentation
from pptx.errors import PackageLimitError

from . import corpus

#: end-of-central-directory record: signature, disk, cd-disk, entries here, entries total,
#: central-directory size, central-directory offset, comment length
_END_RECORD = struct.Struct("<4s4H2LH")
_CENTRAL_HEADER = struct.Struct("<4s6H3L5H2L")
_END_SIGNATURE = b"PK\x05\x06"

_FIXTURE = "self_generated/minimal_clean.pptx"


def _source_bytes() -> bytes:
    return corpus.fixture_path(_FIXTURE).read_bytes()


def _end_record(raw: bytes):
    offset = raw.rfind(_END_SIGNATURE)
    assert offset >= 0, "fixture carries no end-of-central-directory record"
    return offset, _END_RECORD.unpack_from(raw, offset)


# -- the footer must account for every byte after it --------------------------------------


def test_undeclared_bytes_after_the_footer_refuse(tmp_path):
    """A signing or watermarking tool appending in place. PowerPoint prompts to repair."""
    target = tmp_path / "signed.pptx"
    target.write_bytes(_source_bytes() + b"--SIGNATURE-BLOCK-APPENDED-BY-SOME-TOOL--")

    with pytest.raises(
        PackageLimitError, match="no end-of-central-directory record accounting for the end"
    ):
        Presentation(target)


def test_a_single_appended_newline_refuses(tmp_path):
    """`echo "" >> deck.pptx`. One byte is enough for PowerPoint to reject the deck."""
    target = tmp_path / "newline.pptx"
    target.write_bytes(_source_bytes() + b"\n")

    with pytest.raises(
        PackageLimitError, match="no end-of-central-directory record accounting for the end"
    ):
        Presentation(target)


def test_a_declared_comment_length_with_no_comment_refuses(tmp_path):
    """Observed on intact SEC EDGAR filings: one mangled byte in the footer.

    Every member still inflates with a clean CRC, so the package reads fine in stdlib
    `zipfile`. PowerPoint refuses it anyway, which is what makes refusing it correct here.
    """
    raw = bytearray(_source_bytes())
    offset, _ = _end_record(bytes(raw))
    struct.pack_into("<H", raw, offset + 20, 2560)  # -- the only byte that differs

    target = tmp_path / "bogus_comment_length.pptx"
    target.write_bytes(bytes(raw))

    with pytest.raises(
        PackageLimitError, match="no end-of-central-directory record accounting for the end"
    ):
        Presentation(target)


def test_trailing_bytes_declared_as_a_comment_open(tmp_path):
    """The same trailing bytes, declared in the footer, are legitimate and must open.

    This is the boundary: the objection is to *undeclared* trailing data, not to trailing
    data. PowerPoint opens this shape.
    """
    raw = _source_bytes()
    offset, fields = _end_record(raw)
    trailer = b"--SIGNATURE-BLOCK-APPENDED-BY-SOME-TOOL--"
    declared = list(fields)
    declared[-1] = len(trailer)

    target = tmp_path / "declared.pptx"
    target.write_bytes(raw[:offset] + _END_RECORD.pack(*declared) + trailer)

    assert len(Presentation(target).slides) == len(
        Presentation(corpus.fixture_path(_FIXTURE)).slides
    )


def test_a_comment_written_through_the_zip_api_opens(tmp_path):
    """The spec-correct way to attach a comment; PowerPoint opens it."""
    target = tmp_path / "commented.pptx"
    target.write_bytes(_source_bytes())
    with zipfile.ZipFile(target, "a") as archive:
        archive.comment = b"produced by an internal build pipeline"

    Presentation(target)


def test_an_archive_with_no_footer_refuses(tmp_path):
    raw = _source_bytes()
    offset, _ = _end_record(raw)

    target = tmp_path / "footerless.pptx"
    target.write_bytes(raw[:offset] + b"\x00" * 40)

    with pytest.raises(Exception) as excinfo:
        Presentation(target)
    assert type(excinfo.value).__name__ in {"PackageNotFoundError", "PackageLimitError"}


def test_a_stray_footer_signature_inside_a_member_is_ignored(tmp_path):
    """Four bytes of binary media may match the signature. PowerPoint is untroubled."""
    stray = _END_SIGNATURE + struct.pack("<4H2LH", 0, 0, 5, 5, 1234, 5678, 0)
    source = corpus.fixture_path(_FIXTURE)
    target = tmp_path / "stray.pptx"
    with zipfile.ZipFile(source) as incoming, zipfile.ZipFile(target, "w") as outgoing:
        for info in incoming.infolist():
            data = incoming.read(info.filename)
            if info.filename.endswith(".jpeg"):
                # -- stored, so the signature bytes appear literally in the archive
                outgoing.writestr(info.filename, data + stray, zipfile.ZIP_STORED)
            else:
                outgoing.writestr(info.filename, data)

    Presentation(target)


# -- the archive must start at byte zero ---------------------------------------------------


def test_two_concatenated_packages_refuse(tmp_path):
    """`cat a.pptx b.pptx` leaves two footers and two documents in one file.

    Upstream opens it, silently reads whichever document the last footer names, and drops
    the other on save. PowerPoint refuses it. The refusal here comes from the
    central-directory bounds check rather than the footer search.
    """
    raw = _source_bytes()
    target = tmp_path / "concatenated.pptx"
    target.write_bytes(raw + raw)

    with pytest.raises(PackageLimitError, match="unambiguous region"):
        Presentation(target)


def test_prefix_data_before_the_archive_refuses(tmp_path):
    """A stub before the first local header, every offset rebased: the self-extracting shape.

    Three readers, three different documents: paper-pptx used to open this and show the
    correct two slides, LibreOffice renders four pages of binary garbage rather than the
    deck, and PowerPoint prompts to repair. A permissive reader here edits whichever
    reading it happened to take, which is the hazard the package exists to prevent.
    """
    raw = _source_bytes()
    offset, fields = _end_record(raw)
    prefix = b"#!/bin/sh\n# self-extracting stub\n" + b"P" * 100
    central_size, central_offset = fields[5], fields[6]

    body = bytearray(prefix + raw)
    cursor, end = central_offset + len(prefix), central_offset + central_size + len(prefix)
    while cursor < end:
        record = _CENTRAL_HEADER.unpack_from(body, cursor)
        struct.pack_into("<L", body, cursor + 42, record[16] + len(prefix))
        cursor += _CENTRAL_HEADER.size + record[10] + record[11] + record[12]
    shifted = list(fields)
    shifted[6] = central_offset + len(prefix)
    _END_RECORD.pack_into(body, offset + len(prefix), *shifted)

    target = tmp_path / "prefixed.pptx"
    target.write_bytes(bytes(body))

    with pytest.raises(PackageLimitError, match="does not begin with a member record"):
        Presentation(target)


def test_an_empty_archive_is_not_mistaken_for_prefixed(tmp_path):
    """A zero-member archive is 22 bytes of end record and has no first member to lead with.

    Six such files turned up in a 52,941-file scan, all of them `empty.zip` test data. They
    are not packages, so the refusal below is about the missing part rather than the shape
    of the archive, and it must not be the prefix refusal.
    """
    target = tmp_path / "empty.pptx"
    with zipfile.ZipFile(target, "w"):
        pass

    with pytest.raises(Exception) as excinfo:
        Presentation(target)
    assert "does not begin with a member record" not in str(excinfo.value)
