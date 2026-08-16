"""Overall .pptx package."""

from __future__ import annotations

from typing import IO, Iterator

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import OpcPackage
from pptx.opc.packuri import PackURI
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.image import Image, ImagePart
from pptx.parts.media import MediaPart
from pptx.util import lazyproperty


class Package(OpcPackage):
    """An overall .pptx package."""

    @lazyproperty
    def core_properties(self) -> CorePropertiesPart:
        """Instance of |CoreProperties| holding read/write Dublin Core doc properties.

        Creates a default core properties part if one is not present (not common).
        """
        try:
            return self.part_related_by(RT.CORE_PROPERTIES)
        except KeyError:
            core_props = CorePropertiesPart.default(self)
            self.relate_to(core_props, RT.CORE_PROPERTIES)
            return core_props

    def get_or_add_image_part(self, image_file: str | IO[bytes]):
        """
        Return an |ImagePart| object containing the image in *image_file*. If
        the image part already exists in this package, it is reused,
        otherwise a new one is created.
        """
        return self._image_parts.get_or_add_image_part(image_file)

    def get_or_add_media_part(self, media):
        """Return a |MediaPart| object containing the media in *media*.

        If a media part for this media bytestream ("file") is already present
        in this package, it is reused, otherwise a new one is created.
        """
        return self._media_parts.get_or_add_media_part(media)

    def next_image_partname(self, ext: str) -> PackURI:
        """Return a |PackURI| instance representing the next available image partname.

        Partname uses the next available sequence number. *ext* is used as the extention on the
        returned partname.
        """

        def first_available_image_idx():
            image_idxs = sorted(
                [
                    part.partname.idx
                    for part in self.iter_parts()
                    if (
                        part.partname.startswith("/ppt/media/image")
                        and part.partname.idx is not None
                    )
                ]
            )
            for i, image_idx in enumerate(image_idxs):
                idx = i + 1
                if idx < image_idx:
                    return idx
            return len(image_idxs) + 1

        idx = first_available_image_idx()
        return PackURI("/ppt/media/image%d.%s" % (idx, ext))

    def next_media_partname(self, ext):
        """Return |PackURI| instance for next available media partname.

        Partname is first available, starting at sequence number 1. Empty
        sequence numbers are reused. *ext* is used as the extension on the
        returned partname.
        """

        def first_available_media_idx():
            media_idxs = sorted(
                [
                    part.partname.idx
                    for part in self.iter_parts()
                    if part.partname.startswith("/ppt/media/media")
                ]
            )
            for i, media_idx in enumerate(media_idxs):
                idx = i + 1
                if idx < media_idx:
                    return idx
            return len(media_idxs) + 1

        idx = first_available_media_idx()
        return PackURI("/ppt/media/media%d.%s" % (idx, ext))

    @property
    def presentation_part(self):
        """
        Reference to the |Presentation| instance contained in this package.
        """
        return self.main_document_part

    @lazyproperty
    def _image_parts(self):
        """
        |_ImageParts| object providing access to the image parts in this
        package.
        """
        return _ImageParts(self)

    @lazyproperty
    def _media_parts(self):
        """Return |_MediaParts| object for this package.

        The media parts object provides access to all the media parts in this
        package.
        """
        return _MediaParts(self)


class _ImageParts(object):
    """Provides access to the image parts in a package."""

    def __init__(self, package):
        super(_ImageParts, self).__init__()
        self._package = package

    def __iter__(self) -> Iterator[ImagePart]:
        """Generate a reference to each |ImagePart| object in the package."""
        image_parts = []
        for rel in self._package.iter_rels():
            if rel.is_external:
                continue
            if rel.reltype != RT.IMAGE:
                continue
            image_part = rel.target_part
            if image_part in image_parts:
                continue
            image_parts.append(image_part)
            yield image_part

    def get_or_add_image_part(self, image_file: str | IO[bytes]) -> ImagePart:
        """Return |ImagePart| object containing the image in `image_file`.

        `image_file` can be either a path to an image file or a file-like object
        containing an image. If an image part containing this same image already exists,
        that instance is returned, otherwise a new image part is created.
        """
        image = Image.from_file(image_file)
        image_part = self._find_by_sha1(image.sha1)
        return image_part if image_part else ImagePart.new(self._package, image)

    def _find_by_sha1(self, sha1: str) -> ImagePart | None:
        """
        Return an |ImagePart| object belonging to this package or |None| if
        no matching image part is found. The image part is identified by the
        SHA1 hash digest of the image binary it contains.
        """
        for image_part in self:
            # ---skip unknown/unsupported image types, like SVG---
            if not hasattr(image_part, "sha1"):
                continue
            if image_part.sha1 == sha1:
                return image_part
        return None


class _MediaParts(object):
    """Provides access to the media parts in a package.

    Supports iteration and :meth:`get()` using the media object SHA1 hash as
    its key.
    """

    def __init__(self, package):
        super(_MediaParts, self).__init__()
        self._package = package

    def __iter__(self):
        """Generate a reference to each |MediaPart| object in the package."""
        # A media part can appear in more than one relationship (and commonly
        # does in the case of video). Use media_parts to keep track of those
        # that have been "yielded"; they can be skipped if they occur again.
        media_parts = []
        for rel in self._package.iter_rels():
            if rel.is_external:
                continue
            if rel.reltype not in (RT.MEDIA, RT.VIDEO):
                continue
            media_part = rel.target_part
            if media_part in media_parts:
                continue
            media_parts.append(media_part)
            yield media_part

    def get_or_add_media_part(self, media):
        """Return a |MediaPart| object containing the media in *media*.

        If this package already contains a media part for the same
        bytestream, that instance is returned, otherwise a new media part is
        created.
        """
        media_part = self._find_by_sha1(media.sha1)
        if media_part is None:
            media_part = MediaPart.new(self._package, media)
        return media_part

    def _find_by_sha1(self, sha1):
        """Return |MediaPart| object having *sha1* hash or None if not found.

        All media parts belonging to this package are considered. A media
        part is identified by the SHA1 hash digest of its bytestream
        ("file").
        """
        for media_part in self:
            if media_part.sha1 == sha1:
                return media_part
        return None


# ===========================================================================================
# paper-pptx package kernel — additive module-level utilities.
#
# `pptx.package` already held the opc `Package` class, so the kernel extends this module
# additively rather than shadowing anything. Everything below is new API: semantic XML
# comparison, part-level package diffing, and compare-based narrow save.
# ===========================================================================================

import io as _io  # noqa: E402
import os as _os  # noqa: E402
import stat as _stat  # noqa: E402
import tempfile as _tempfile  # noqa: E402
import zipfile as _zipfile  # noqa: E402
from dataclasses import dataclass as _dataclass  # noqa: E402
from typing import Sequence, Tuple, Union  # noqa: E402

from pptx.errors import UnsupportedStructureError  # noqa: E402

#: Pinned deterministic zip entry timestamp (the zip epoch).
_ZIP_EPOCH = (1980, 1, 1, 0, 0, 0)
_CONTENT_TYPES = "[Content_Types].xml"


def xml_equivalent(a: Union[bytes, str], b: Union[bytes, str]) -> bool:
    """Return True when `a` and `b` are semantically equivalent XML documents.

    Comparison is Canonical XML (C14N 2.0) with prefixes rewritten, so differences in
    attribute order, namespace-prefix spelling, XML declaration, self-closing-tag style, and
    pretty-print indentation are equivalent — while element order, attribute values, and all
    potentially meaningful text compare exactly.

    Whitespace handling is deliberately asymmetric: a whitespace-only text
    node is ignored ONLY where its parent element has element children (structural
    indentation; OOXML defines no mixed content, so such whitespace can never render). Text
    of element-childless elements — `a:t` and friends — is never normalized in any way: two
    documents differing only by a trailing space inside a text node are NOT equivalent.

    Raises |ValueError| when either argument is not well-formed XML.
    """
    return _c14n_bytes(a) == _c14n_bytes(b)


def _c14n_bytes(data: Union[bytes, str]) -> bytes:
    from xml.etree import ElementTree as _ElementTree

    out = _io.StringIO()
    try:
        _ElementTree.canonicalize(
            _drop_structural_whitespace(data),
            out=out,
            strip_text=False,
            rewrite_prefixes=True,
        )
    except _ElementTree.ParseError as e:
        raise ValueError("not well-formed XML: %s" % e)
    return out.getvalue().encode("utf-8")


def _drop_structural_whitespace(data: Union[bytes, str]) -> str:
    """Return `data` re-serialized with structural-indentation whitespace removed.

    Drops (a) whitespace-only `.text` of elements that have element children and (b)
    whitespace-only `.tail` of any element (a tail always sits inside a parent that has
    element children — this element). The text of element-childless elements is untouchable
    here by construction, so preserved-space content like `a:t` can never be altered.

    Raises |ValueError| on malformed XML.
    """
    from lxml import etree as _etree

    if isinstance(data, str):
        data = data.encode("utf-8")
    try:
        root = _etree.fromstring(data)
    except _etree.XMLSyntaxError as e:
        raise ValueError("not well-formed XML: %s" % e)
    for element in root.iter():
        if len(element) and element.text is not None and not element.text.strip():
            element.text = None
        if element.tail is not None and not element.tail.strip():
            element.tail = None
    return _etree.tostring(root, encoding="unicode")


def _is_xml_member(name: str) -> bool:
    return name.endswith(".xml") or name.endswith(".rels")


def _members_semantically_equal(name: str, a: bytes, b: bytes) -> bool:
    """True when zip member `name` carries semantically identical content in `a` and `b`.

    XML members compare via `xml_equivalent`; `[Content_Types].xml` additionally compares
    with its children sorted first, because OPC defines no significance for their order and
    every writer orders them differently. Binary members compare byte-for-byte (`a == b` is
    checked by callers first).
    """
    if not _is_xml_member(name):
        return a == b
    if name == _CONTENT_TYPES:
        return _c14n_bytes(_sorted_content_types(a)) == _c14n_bytes(_sorted_content_types(b))
    return xml_equivalent(a, b)


def _sorted_content_types(data: bytes) -> bytes:
    """Return `data` with the `[Content_Types].xml` children in a canonical sort order."""
    from lxml import etree as _etree

    root = _etree.fromstring(data)
    children = sorted(
        root,
        key=lambda e: (
            e.tag,
            e.get("Extension") or "",
            e.get("PartName") or "",
            e.get("ContentType") or "",
        ),
    )
    for child in children:
        root.append(child)  # -- re-appending moves each to the end, in sorted order
    return _etree.tostring(root)


@_dataclass(frozen=True)
class PartDelta:
    """One differing package member."""

    partname: str  #: partname-style, e.g. "/ppt/slides/slide1.xml"
    kind: str  #: "xml" | "binary"
    change: str  #: "added" | "removed" | "changed"
    detail: str  #: human-readable note on the difference

    def to_dict(self) -> dict:
        return {
            "partname": self.partname,
            "kind": self.kind,
            "change": self.change,
            "detail": self.detail,
        }


@_dataclass(frozen=True)
class PackageDiff:
    """Part-by-part semantic diff between two packages. Schema "paper-package-diff" v1.

    ``deltas`` holds one :class:`PartDelta` per package member that was added, removed,
    or semantically changed; it is empty when the two packages are equivalent.
    """

    deltas: Tuple[PartDelta, ...]

    @property
    def is_empty(self) -> bool:
        return not self.deltas

    def to_dict(self) -> dict:
        return {
            "schema": "paper-package-diff",
            "version": 1,
            "deltas": [delta.to_dict() for delta in self.deltas],
        }

    def __repr__(self) -> str:
        return "PackageDiff(%d deltas)" % len(self.deltas)


def diff_package(path_a: str, path_b: str) -> PackageDiff:
    """Return the |PackageDiff| between the packages at `path_a` and `path_b`.

    XML members are compared semantically (`xml_equivalent`; `[Content_Types].xml`
    order-insensitively), binary members by bytes. Members appearing in only one package
    report as "added" (only in `path_b`) or "removed" (only in `path_a`). Deltas are sorted
    by partname.
    """
    return _diff_maps(_read_zip_map(path_a), _read_zip_map(path_b), str(path_a), str(path_b))


def _diff_maps(map_a: dict, map_b: dict, label_a: str, label_b: str) -> PackageDiff:
    """Return the |PackageDiff| between two in-memory member maps."""
    deltas = []
    for name in sorted(set(map_a) | set(map_b)):
        partname = "/" + name
        kind = "xml" if _is_xml_member(name) else "binary"
        if name not in map_b:
            deltas.append(PartDelta(partname, kind, "removed", "only in %s" % label_a))
        elif name not in map_a:
            deltas.append(PartDelta(partname, kind, "added", "only in %s" % label_b))
        elif map_a[name] == map_b[name]:
            continue
        elif _members_semantically_equal(name, map_a[name], map_b[name]):
            continue  # -- byte-different but semantically identical: not a delta
        elif kind == "xml":
            deltas.append(PartDelta(partname, kind, "changed", "semantic XML change"))
        else:
            deltas.append(
                PartDelta(
                    partname,
                    kind,
                    "changed",
                    "binary change (%d -> %d bytes)" % (len(map_a[name]), len(map_b[name])),
                )
            )
    return PackageDiff(tuple(deltas))


def patch_save(original_path: str, document, out_path: str) -> PackageDiff:
    """Save `document` to `out_path`, restoring original bytes for unchanged XML parts.

    Compare-based narrow save: `document` (a |Presentation|) is serialized
    normally, then every XML member that is semantically identical to its counterpart in
    `original_path` is written with the ORIGINAL bytes, so unrelated parts never churn.
    Returns the residual |PackageDiff| between `original_path` and `out_path`.

    Writes are deterministic — entry order is `[Content_Types].xml`, `_rels/.rels`, then all
    remaining members sorted; every entry timestamp is fixed to 1980-01-01 — and atomic: the
    package is built in a temp file in `out_path`'s directory and moved into place with
    `os.replace`, so a mid-write failure leaves any existing `out_path` untouched. When
    nothing changed at all, `out_path` is written as an exact byte copy of `original_path`
    (the no-op round trip is byte-identical).

    Not interchangeable with |Presentation.save|, which is also atomic on a path:
    atomicity is how the bytes land, narrowness is which bytes get written. `save()`
    re-serializes every part, so even an unchanged part gets new bytes; `patch_save`
    restores the original bytes for every part that is semantically identical.

    Symlinked destinations are resolved, so the file a link names is the file replaced.

    Raises |UnsupportedStructureError| when `original_path` is not a readable zip package
    (before anything is written) and |ValueError| when `document` cannot save itself.
    """
    if not hasattr(document, "save"):
        raise ValueError(
            "document must be a Presentation (or provide .save(stream)), got %r"
            % type(document).__name__
        )
    original_map = _read_zip_map(original_path)

    buffer = _io.BytesIO()
    # -- narrow save runs through the ordinary stream-save path, so `patch_save` has a live
    # -- dependency on it: a regression there breaks this too
    document.save(buffer)
    candidate_map = _read_zip_map_from_bytes(buffer.getvalue(), "in-memory save output")

    out_map = {}
    for name, data in candidate_map.items():
        original = original_map.get(name)
        byte_differs = original is not None and data != original
        if byte_differs and _members_semantically_equal(name, original, data):
            data = original
        out_map[name] = data

    # -- the residual diff is computed BEFORE writing: out_path may equal original_path
    # -- (in-place narrow save), in which case a post-write diff would always be empty
    residual = _diff_maps(original_map, out_map, str(original_path), str(out_path))

    unchanged = set(out_map) == set(original_map) and all(
        out_map[name] == original_map[name] for name in out_map
    )
    if unchanged:
        _atomic_write_bytes(_read_file_bytes(original_path), out_path)
        return PackageDiff(())

    _atomic_write_zip(out_map, out_path)
    return residual


def _member_write_order(names: "Sequence[str]") -> "Sequence[str]":
    head = [n for n in (_CONTENT_TYPES, "_rels/.rels") if n in names]
    return head + sorted(n for n in names if n not in head)


def _atomic_write_zip(member_map: dict, out_path: str) -> None:
    def write(handle):
        with _zipfile.ZipFile(handle, "w") as zipf:
            for name in _member_write_order(list(member_map)):
                info = _zipfile.ZipInfo(name, date_time=_ZIP_EPOCH)
                info.compress_type = _zipfile.ZIP_DEFLATED
                zipf.writestr(info, member_map[name], compresslevel=6)

    _atomic_write(write, out_path)


def _atomic_write_bytes(data: bytes, out_path: str) -> None:
    _atomic_write(lambda handle: handle.write(data), out_path)


def _atomic_write(write, out_path: str) -> None:
    """Run `write(file_handle)` against a temp file, then move it into place atomically."""
    # -- realpath, not abspath: abspath normalizes but does not resolve symlinks, so the
    # -- replace below would land on the link and leave the file it names untouched
    destination = _os.path.realpath(str(out_path))
    out_dir = _os.path.dirname(destination)
    existing_mode = (
        _stat.S_IMODE(_os.stat(destination).st_mode) if _os.path.exists(destination) else None
    )
    fd, temp_path = _tempfile.mkstemp(suffix=".pptx.partial", dir=out_dir)
    try:
        with _os.fdopen(fd, "wb") as handle:
            write(handle)
        if existing_mode is not None:
            _os.chmod(temp_path, existing_mode)
        else:
            # -- mkstemp creates 0600; a NEW destination must get the mode a plain
            # -- open() would have given it, or downstream readers lose access
            active_umask = _os.umask(0)
            _os.umask(active_umask)
            _os.chmod(temp_path, 0o666 & ~active_umask)
        _os.replace(temp_path, destination)
    except BaseException:
        if _os.path.exists(temp_path):
            _os.unlink(temp_path)
        raise


def _read_file_bytes(path: str) -> bytes:
    with open(str(path), "rb") as handle:
        return handle.read()


def _read_zip_map(path: str) -> dict:
    try:
        data = _read_file_bytes(path)
    except OSError as e:
        raise UnsupportedStructureError("cannot read package %s: %s" % (path, e))
    return _read_zip_map_from_bytes(data, str(path))


def _read_zip_map_from_bytes(data: bytes, label: str) -> dict:
    try:
        with _zipfile.ZipFile(_io.BytesIO(data)) as zipf:
            names = zipf.namelist()
            if len(names) != len(set(names)):
                raise UnsupportedStructureError(
                    "package %s contains duplicate member names" % label
                )
            return {name: zipf.read(name) for name in names}
    except _zipfile.BadZipFile:
        raise UnsupportedStructureError("%s is not a zip package" % label)
