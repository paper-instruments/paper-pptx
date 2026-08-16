"""What `save()` accepts as a destination, and what it guarantees about each kind.

Upstream required only `write()` of a stream destination. Staging the package before
touching the destination is a fork addition worth keeping -- it means a serialization
failure cannot emit a partial package -- but the rollback that was built on top of it
demanded `read`/`seek`/`tell`/`truncate` as *entry* requirements, which refused the
ordinary ways of handing `save()` a stream:

    with open(path, "wb") as f: prs.save(f)     # not readable
    prs.save(sys.stdout.buffer)                 # not seekable
    prs.save(socket.makefile("wb"))             # neither

Every stream test in this suite used `io.BytesIO`, the one object that is readable,
writable, seekable and truncatable at once, which is why none of it surfaced. The cases
below use destinations that are not.

The path branch keeps its atomicity, and the last test here is what proves that is still
worth the deviation it costs.
"""

from __future__ import annotations

import io
import os
import socket
import threading
import zipfile

import pytest

from pptx import Presentation

from . import corpus


def _deck():
    prs = Presentation(str(corpus.fixture_path("self_generated/minimal_clean.pptx")))
    return prs


def _is_readable_package(path) -> bool:
    with zipfile.ZipFile(path) as archive:
        return archive.testzip() is None and "[Content_Types].xml" in archive.namelist()


# -- stream destinations -------------------------------------------------------------------


def test_save_to_a_write_only_binary_file(tmp_path):
    """`open(path, "wb")` is the destination the docstring advertises and is not readable."""
    target = tmp_path / "wb.pptx"
    with open(target, "wb") as handle:
        _deck().save(handle)

    assert _is_readable_package(target)
    assert len(Presentation(str(target)).slides) == len(_deck().slides)


def test_save_to_an_object_exposing_only_write_and_flush():
    """A pipe, an HTTP response body, an upload handle: `write` is all upstream needed."""

    class WriteOnly:
        def __init__(self):
            self.chunks = []

        def write(self, data):
            self.chunks.append(bytes(data))
            return len(data)

        def flush(self):
            pass

    sink = WriteOnly()
    _deck().save(sink)

    payload = b"".join(sink.chunks)
    assert payload[:2] == b"PK"
    assert _is_readable_package(io.BytesIO(payload))


def test_save_to_an_unseekable_socket():
    """The shape of streaming a generated deck into a network response."""
    left, right = socket.socketpair()
    received = bytearray()

    def drain():
        while True:
            chunk = right.recv(65536)
            if not chunk:
                break
            received.extend(chunk)

    # -- a socketpair buffers only a few KB, so an undrained write of a whole deck blocks
    reader = threading.Thread(target=drain, daemon=True)
    reader.start()
    try:
        with left.makefile("wb") as sink:
            _deck().save(sink)
    finally:
        left.close()
        reader.join(timeout=30)
        right.close()

    assert _is_readable_package(io.BytesIO(bytes(received)))


def test_save_honors_the_streams_current_position():
    """Embedding a package inside a larger container stream must not destroy the prefix."""
    buffer = io.BytesIO()
    buffer.write(b"CONTAINER-HEADER")

    _deck().save(buffer)

    payload = buffer.getvalue()
    assert payload.startswith(b"CONTAINER-HEADER")
    assert _is_readable_package(io.BytesIO(payload[len(b"CONTAINER-HEADER") :]))


def test_save_at_a_nonzero_position_leaves_the_containers_own_tail_alone():
    """The suffix half of honoring the position: bytes past the package are the caller's.

    Honoring the cursor is only half an embedding. Truncating at the end of the package
    discards whatever the container had after the insertion point, which upstream's
    `zipfile.ZipFile(stream, "w")` never did -- silently, and however much of it there is.
    The tail here is deliberately longer than the package, so plain overwriting cannot
    account for its loss the way it could if the tail were short.
    """
    staged = io.BytesIO()
    _deck().save(staged)
    package_size = len(staged.getvalue())

    header, tail = b"CONTAINER-HEADER", b"TAILBYTE" * 40_000
    assert len(tail) > package_size, "tail must outlast the package or this proves nothing"

    buffer = io.BytesIO()
    buffer.write(header)
    buffer.write(tail)
    buffer.seek(len(header))

    _deck().save(buffer)

    payload = buffer.getvalue()
    assert len(payload) == len(header) + len(tail), "the container's tail was truncated"
    assert payload.startswith(header)
    assert payload[len(header) + package_size :] == tail[package_size:]
    assert _is_readable_package(io.BytesIO(payload[len(header) : len(header) + package_size]))


def test_a_shorter_package_leaves_no_tail_of_the_previous_contents():
    """Stale bytes past the end of a package read as a corrupt file, not as a failure.

    Writing a smaller document over a larger one has to drop what the old one left behind.
    A caller who reopens the destination would otherwise get a package with trailing
    garbage, which is worse than an error: it looks like it worked.
    """
    buffer = io.BytesIO(b"X" * 400_000)
    buffer.seek(0)

    _deck().save(buffer)

    payload = buffer.getvalue()
    assert b"XXXX" not in payload[-1024:], "previous contents survived past the package"
    assert _is_readable_package(io.BytesIO(payload))


def test_rollback_restores_a_destination_that_supports_it():
    """The guarantee this change is careful to keep, pinned at the destination contract."""

    class FailsPartWayThroughAWrite(io.BytesIO):
        def __init__(self, initial):
            super().__init__(initial)
            self._fail_next = True

        def write(self, data):
            if self._fail_next:
                self._fail_next = False
                # -- land some bytes first, so a missing rollback is visible
                super().write(data[:512])
                raise OSError("forced destination write failure")
            return super().write(data)

    original = b"AN EARLIER DOCUMENT" * 4096
    destination = FailsPartWayThroughAWrite(original)
    destination.seek(11)

    with pytest.raises(OSError, match="forced destination write failure"):
        _deck().save(destination)

    assert destination.getvalue() == original
    assert destination.tell() == 11


def test_a_serialization_failure_writes_nothing_to_the_stream(monkeypatch):
    """Staging's guarantee: a partial package never reaches a sink that cannot take it back."""

    class WriteOnly:
        def __init__(self):
            self.total = 0

        def write(self, data):
            self.total += len(data)
            return len(data)

        def flush(self):
            pass

    prs = _deck()
    sink = WriteOnly()

    from pptx.opc import package as package_module

    def explode(*args, **kwargs):
        raise RuntimeError("serialization blew up")

    monkeypatch.setattr(package_module.PackageWriter, "write", explode)

    with pytest.raises(RuntimeError, match="serialization blew up"):
        prs.save(sink)

    assert sink.total == 0


# -- path destinations ---------------------------------------------------------------------


def test_save_through_a_symlink_writes_the_file_the_link_names(tmp_path):
    """`os.replace` on an unresolved path destroys the link and never writes the real deck."""
    real_dir, link_dir = tmp_path / "real", tmp_path / "link"
    real_dir.mkdir()
    link_dir.mkdir()
    real = real_dir / "deck.pptx"
    _deck().save(str(real))
    before = len(Presentation(str(real)).slides)

    link = link_dir / "deck.pptx"
    os.symlink(str(real), str(link))

    prs = Presentation(str(link))
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.save(str(link))

    assert link.is_symlink(), "the symlink was replaced by a regular file"
    assert len(Presentation(str(real)).slides) == before + 1, "the real deck was not written"


def test_patch_save_through_a_symlink_writes_the_file_the_link_names(tmp_path):
    from pptx.package import patch_save

    real_dir, link_dir = tmp_path / "real", tmp_path / "link"
    real_dir.mkdir()
    link_dir.mkdir()
    real = real_dir / "deck.pptx"
    _deck().save(str(real))

    link = link_dir / "deck.pptx"
    os.symlink(str(real), str(link))

    prs = Presentation(str(real))
    prs.slides.add_slide(prs.slide_layouts[6])
    patch_save(str(real), prs, str(link))

    assert link.is_symlink(), "the symlink was replaced by a regular file"
    assert len(Presentation(str(real)).slides) == 2, "the real deck was not written"


def test_a_failed_path_save_leaves_the_existing_deck_byte_identical(tmp_path, monkeypatch):
    """The load-bearing test for the atomicity deviation.

    Upstream serialized directly into the destination, so a failure part-way through left a
    truncated file where the deck used to be. This is the guarantee that justifies departing
    from "no changes to save()", so it is pinned rather than assumed.
    """
    target = tmp_path / "deck.pptx"
    _deck().save(str(target))
    before = target.read_bytes()

    from pptx.opc import package as package_module

    original_write = package_module.PackageWriter.write

    def fail_partway(pkg_file, pkg_rels, parts):
        # -- write a plausible prefix first, so a non-atomic implementation would be
        # -- caught leaving exactly that behind
        if isinstance(pkg_file, str):
            with open(pkg_file, "wb") as handle:
                handle.write(b"PK\x03\x04partial-package-bytes")
        raise RuntimeError("blew up mid-serialization")

    monkeypatch.setattr(package_module.PackageWriter, "write", fail_partway)

    with pytest.raises(RuntimeError, match="blew up mid-serialization"):
        _deck().save(str(target))

    assert target.read_bytes() == before, "the existing deck was modified by a failed save"
    assert _is_readable_package(target)

    monkeypatch.setattr(package_module.PackageWriter, "write", original_write)
    leftovers = [p.name for p in tmp_path.iterdir() if p.name.endswith(".partial")]
    assert leftovers == [], "a temp file survived the failure: %r" % leftovers
