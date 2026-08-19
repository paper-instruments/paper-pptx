"""Unambiguous ZIP reads for OPC packages.

An OPC package must have exactly one interpretation. Archives whose structure admits
more than one are refused rather than resolved by guesswork: ambiguous or multi-disk
end-of-central-directory records, member counts and offsets that disagree, duplicate or
case-equivalent member names, local headers that contradict the central directory, and
data that overlaps or trails a member's declared extent.

The archive must also span its file exactly, beginning at the first byte and ending with
an end record that accounts for every byte after it. Both halves are PowerPoint's rule,
established by opening each shape in the application: it refuses a package carrying
undeclared trailing bytes, one whose end record declares an absent comment, and one with
bytes in front of its first member. No Python reader reproduces this -- stdlib ``zipfile``,
upstream python-pptx and LibreOffice accept all three -- so these refusals look like
over-strictness until measured. Declared trailing data is legitimate and is accepted.

Every member must also resolve to a content type, through an ``Override`` naming the part
or a ``Default`` matching its extension. A member with neither has no type at all, and
PowerPoint refuses the package whatever the part is for -- measured on a thumbnail it never
renders, a slide it must load, an image it draws, and a part nothing references. The rule
therefore keys on physical membership rather than on what the loader would otherwise read.
A part nothing references is *not* refused: PowerPoint opens that package and drops the
part on its next save, which is what ``save()`` does too.

Only the two compression methods permitted by the OPC ZIP mapping (stored and
deflated) are accepted. Every member is inflated from its raw compressed bytes
rather than through ``ZipFile.read()``. This allows actual output length, CRC, and
deflate end-of-stream to be checked even when central-directory sizes lie.
"""

from __future__ import annotations

import os
import stat
import struct
import unicodedata
import zlib
from contextlib import suppress
from typing import BinaryIO, Dict, List, Optional, Tuple, Union, cast
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile, ZipInfo

from lxml import etree

from pptx.errors import PackageLimitError

_READ_CHUNK_BYTES = 64 * 1024
_LOCAL_HEADER = struct.Struct("<4s5H3L2H")
_LOCAL_HEADER_SIGNATURE = b"PK\x03\x04"
_CENTRAL_HEADER = struct.Struct("<4s6H3L5H2L")
_CENTRAL_HEADER_SIGNATURE = b"PK\x01\x02"
_END_RECORD = struct.Struct("<4s4H2LH")
_END_RECORD_SIGNATURE = b"PK\x05\x06"
_ZIP64_END_RECORD = struct.Struct("<4sQ2H2L4Q")
_ZIP64_END_RECORD_SIGNATURE = b"PK\x06\x06"
_ZIP64_LOCATOR = struct.Struct("<4sLQL")
_ZIP64_LOCATOR_SIGNATURE = b"PK\x06\x07"
_MAX_END_COMMENT_BYTES = 65_535

_CONTENT_TYPES_NAME = "[Content_Types].xml"
_CONTENT_TYPES_NAMESPACE = (
    "http://schemas.openxmlformats.org/package/2006/content-types"
)
_CONTENT_TYPES_TAG = f"{{{_CONTENT_TYPES_NAMESPACE}}}Types"
_DEFAULT_TAG = f"{{{_CONTENT_TYPES_NAMESPACE}}}Default"
_OVERRIDE_TAG = f"{{{_CONTENT_TYPES_NAMESPACE}}}Override"
_CONTENT_TYPES_PARSER = etree.XMLParser(
    load_dtd=False,
    no_network=True,
    remove_blank_text=False,
    resolve_entities=False,
)

_FLAG_ENCRYPTED = 0x0001
_FLAG_DATA_DESCRIPTOR = 0x0008
_FLAG_PATCHED_DATA = 0x0020
_FLAG_STRONG_ENCRYPTION = 0x0040
_FLAG_UTF8_NAME = 0x0800

_SUPPORTED_COMPRESSION = (ZIP_STORED, ZIP_DEFLATED)
_HEX_DIGITS = frozenset("0123456789ABCDEF")
_URI_UNRESERVED = frozenset(b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~")


def compressed_size(source: object) -> Optional[int]:
    """Return physical byte size for a path or seekable stream, if available."""
    if isinstance(source, (str, os.PathLike)):
        path = cast(Union[str, "os.PathLike[str]"], source)
        return os.path.getsize(path)

    fileno = getattr(source, "fileno", None)
    if callable(fileno):
        try:
            descriptor = fileno()
        except (AttributeError, OSError, TypeError, ValueError):
            pass
        else:
            if isinstance(descriptor, int) and not isinstance(descriptor, bool):
                try:
                    return os.fstat(descriptor).st_size
                except (OSError, ValueError):
                    pass

    tell = getattr(source, "tell", None)
    seek = getattr(source, "seek", None)
    if not callable(tell) or not callable(seek):
        return None
    try:
        position = tell()
    except (AttributeError, OSError, TypeError, ValueError):
        return None
    try:
        seek(0, os.SEEK_END)
        size = tell()
    except (AttributeError, OSError, TypeError, ValueError):
        return None
    finally:
        with suppress(AttributeError, OSError, TypeError, ValueError):
            seek(position, os.SEEK_SET)
    return size if isinstance(size, int) and not isinstance(size, bool) else None


def preflight_zip(source: object) -> None:
    """Validate central-directory metadata before ``ZipFile`` parses it.

    An archive must declare one unambiguous central-directory region: a single
    end-of-central-directory record, no multi-disk structure, and member counts,
    offsets, and sizes that agree with each other and with the records actually
    present. This preflight reads only fixed-size records and skips variable
    fields without retaining them.
    """
    if isinstance(source, (str, os.PathLike)):
        path = cast(Union[str, "os.PathLike[str]"], source)
        with open(path, "rb") as stream:
            _preflight_zip_stream(stream)
        return

    read = getattr(source, "read", None)
    seek = getattr(source, "seek", None)
    tell = getattr(source, "tell", None)
    if not callable(read) or not callable(seek) or not callable(tell):
        # ``ZipFile`` will reject objects that do not implement its stream
        # protocol before it can parse or allocate central-directory entries.
        return
    _preflight_zip_stream(cast(BinaryIO, source))


def _preflight_zip_stream(stream: BinaryIO) -> None:
    """Validate the archive footer and central directory before `zipfile` parses them.

    Runs first, so an ambiguous archive refuses before any part is read. Requires one unambiguous
    central-directory region: a single end record accounting for the end of the file, no multi-disk
    structure, counts and offsets that agree, and a first byte that starts a member record. Bytes
    appended after the end record and bytes prepended before the archive both refuse.
    """
    try:
        original_position = stream.tell()
        stream.seek(0, os.SEEK_END)
        archive_size_value = cast(object, stream.tell())
    except (AttributeError, OSError, TypeError, ValueError) as exc:
        raise PackageLimitError("ZIP package input must be a seekable binary stream") from exc

    try:
        if not isinstance(archive_size_value, int) or isinstance(archive_size_value, bool):
            # Test doubles and invalid stream implementations cannot reach
            # ZipFile's central-directory parser with this result type.
            return
        archive_size = archive_size_value

        end_offset, end_fields = _find_end_record(stream, archive_size)
        (
            _signature,
            disk_number,
            central_disk,
            disk_entries,
            total_entries,
            central_size,
            central_offset,
            _comment_length,
        ) = end_fields
        if disk_number != 0 or central_disk != 0:
            raise PackageLimitError("multi-disk ZIP packages are not supported")

        zip64_required = (
            disk_entries == 0xFFFF
            or total_entries == 0xFFFF
            or central_size == 0xFFFFFFFF
            or central_offset == 0xFFFFFFFF
        )
        central_end = end_offset
        if zip64_required:
            (
                disk_entries,
                total_entries,
                central_size,
                central_offset,
                central_end,
            ) = _read_zip64_end_record(stream, end_offset, end_fields)

        if disk_entries != total_entries:
            raise PackageLimitError("ZIP central-directory counts disagree across disks")
        if central_offset < 0 or central_size < 0:
            raise PackageLimitError("ZIP central-directory metadata is negative")
        if central_offset + central_size != central_end:
            raise PackageLimitError(
                "ZIP central-directory offset and size do not identify one unambiguous region"
            )
        if total_entries == 0 and central_size != 0:
            raise PackageLimitError("empty ZIP package has a non-empty central directory")
        if total_entries and central_size < total_entries * _CENTRAL_HEADER.size:
            raise PackageLimitError("ZIP central directory is too small for its member count")

        # -- The archive must also START where the file does. Bytes in front of the first
        # -- member leave readers disagreeing about where the archive begins: PowerPoint
        # -- refuses such a package, LibreOffice renders something other than the deck, and
        # -- a permissive reader edits whichever reading it happened to take. An archive
        # -- declaring no members has no first member and legitimately begins with its own
        # -- end record.
        if total_entries:
            stream.seek(0, os.SEEK_SET)
            if stream.read(len(_LOCAL_HEADER_SIGNATURE)) != _LOCAL_HEADER_SIGNATURE:
                raise PackageLimitError(
                    "ZIP package does not begin with a member record, so bytes precede the "
                    "archive and readers disagree about where it starts; remove the leading "
                    "bytes or re-save the package from PowerPoint"
                )

        _scan_central_directory(
            stream,
            central_offset,
            central_size,
            total_entries,
        )
    finally:
        with suppress(AttributeError, OSError, TypeError, ValueError):
            stream.seek(original_position, os.SEEK_SET)


def _find_end_record(stream: BinaryIO, archive_size: int) -> Tuple[int, Tuple[int, ...]]:
    """Locate the end-of-central-directory record that accounts for the end of the file.

    Refuses when no record does, or when more than one does. Either way the member list has more
    than one reading, and PowerPoint refuses the package in that state too.
    """
    tail_size = min(archive_size, _END_RECORD.size + _MAX_END_COMMENT_BYTES)
    stream.seek(archive_size - tail_size, os.SEEK_SET)
    tail = stream.read(tail_size)
    if len(tail) != tail_size:
        raise PackageLimitError("ZIP package ends before its declared physical size")

    candidates: List[Tuple[int, Tuple[int, ...]]] = []
    cursor = 0
    while True:
        index = tail.find(_END_RECORD_SIGNATURE, cursor)
        if index < 0:
            break
        cursor = index + 1
        if index + _END_RECORD.size > len(tail):
            continue
        fields = _END_RECORD.unpack_from(tail, index)
        comment_length = fields[-1]
        if index + _END_RECORD.size + comment_length == len(tail):
            candidates.append((archive_size - tail_size + index, fields))

    if not candidates:
        raise PackageLimitError(
            "ZIP package has no end-of-central-directory record accounting for the end of "
            "the file: either bytes were appended after the archive, or the record declares "
            "an archive comment that is not there; PowerPoint refuses a package in this "
            "state, so remove the trailing bytes or re-save the package"
        )
    if len(candidates) != 1:
        raise PackageLimitError(
            "ZIP package has more than one end-of-central-directory record accounting for "
            "the end of the file, so its member list has no single reading; re-save the "
            "package from PowerPoint to rewrite one unambiguous record"
        )
    return candidates[0]


def _read_zip64_end_record(
    stream: BinaryIO,
    end_offset: int,
    legacy_fields: Tuple[int, ...],
) -> Tuple[int, int, int, int, int]:
    """Read the ZIP64 locator and end record, returning counts, central size and offset, and the
    record's own offset.

    Refuses a missing, truncated, or misshaped locator or record, ZIP64 multi-disk structure, and
    any ZIP64 field that disagrees with a non-sentinel legacy value.
    """
    locator_offset = end_offset - _ZIP64_LOCATOR.size
    if locator_offset < 0:
        raise PackageLimitError("ZIP64 package is missing its locator")
    stream.seek(locator_offset, os.SEEK_SET)
    locator_data = stream.read(_ZIP64_LOCATOR.size)
    if len(locator_data) != _ZIP64_LOCATOR.size:
        raise PackageLimitError("ZIP64 locator is truncated")
    signature, record_disk, record_offset, disk_count = _ZIP64_LOCATOR.unpack(locator_data)
    if signature != _ZIP64_LOCATOR_SIGNATURE:
        raise PackageLimitError("ZIP64 package is missing its locator")
    if record_disk != 0 or disk_count != 1:
        raise PackageLimitError("multi-disk ZIP64 packages are not supported")
    if record_offset < 0 or record_offset + _ZIP64_END_RECORD.size != locator_offset:
        raise PackageLimitError("ZIP64 end record has an ambiguous offset or size")

    stream.seek(record_offset, os.SEEK_SET)
    record_data = stream.read(_ZIP64_END_RECORD.size)
    if len(record_data) != _ZIP64_END_RECORD.size:
        raise PackageLimitError("ZIP64 end record is truncated")
    (
        signature,
        record_size,
        _made_by,
        _extract_version,
        disk_number,
        central_disk,
        disk_entries,
        total_entries,
        central_size,
        central_offset,
    ) = _ZIP64_END_RECORD.unpack(record_data)
    if signature != _ZIP64_END_RECORD_SIGNATURE or record_size != 44:
        raise PackageLimitError("ZIP64 end record has an unsupported or ambiguous shape")
    if disk_number != 0 or central_disk != 0 or disk_entries != total_entries:
        raise PackageLimitError("multi-disk ZIP64 packages are not supported")

    legacy_disk_entries = legacy_fields[3]
    legacy_total_entries = legacy_fields[4]
    legacy_central_size = legacy_fields[5]
    legacy_central_offset = legacy_fields[6]
    _validate_zip64_legacy_value(legacy_disk_entries, disk_entries, 0xFFFF, "member count")
    _validate_zip64_legacy_value(legacy_total_entries, total_entries, 0xFFFF, "member count")
    _validate_zip64_legacy_value(
        legacy_central_size,
        central_size,
        0xFFFFFFFF,
        "central-directory size",
    )
    _validate_zip64_legacy_value(
        legacy_central_offset,
        central_offset,
        0xFFFFFFFF,
        "central-directory offset",
    )
    return disk_entries, total_entries, central_size, central_offset, record_offset


def _validate_zip64_legacy_value(
    legacy: int,
    actual: int,
    sentinel: int,
    label: str,
) -> None:
    """Refuse when a ZIP64 field and its legacy counterpart disagree about the same quantity."""
    if legacy != sentinel and legacy != actual:
        raise PackageLimitError(f"ZIP64 {label} disagrees with the legacy end record")


def _scan_central_directory(
    stream: BinaryIO,
    central_offset: int,
    central_size: int,
    expected_count: int,
) -> None:
    """Walk every central-directory record, refusing a directory the region cannot hold one way.

    Refuses a truncated or unsigned record, a record naming a nonzero disk, records that overrun the
    declared region, and a member count that disagrees with the end record.
    """
    cursor = central_offset
    central_end = central_offset + central_size
    actual_count = 0
    while cursor < central_end:
        if cursor + _CENTRAL_HEADER.size > central_end:
            raise PackageLimitError("ZIP central directory ends inside a member record")
        stream.seek(cursor, os.SEEK_SET)
        header = stream.read(_CENTRAL_HEADER.size)
        if len(header) != _CENTRAL_HEADER.size:
            raise PackageLimitError("ZIP central-directory member record is truncated")
        fields = _CENTRAL_HEADER.unpack(header)
        if fields[0] != _CENTRAL_HEADER_SIGNATURE:
            raise PackageLimitError("ZIP central directory contains an unsupported record")
        name_length, extra_length, comment_length = fields[10:13]
        if fields[13] != 0:
            raise PackageLimitError("multi-disk ZIP member records are not supported")
        record_size = _CENTRAL_HEADER.size + name_length + extra_length + comment_length
        if cursor + record_size > central_end:
            raise PackageLimitError("ZIP central-directory member record exceeds its region")
        cursor += record_size
        actual_count += 1
        if actual_count > expected_count:
            raise PackageLimitError("ZIP central-directory member count exceeds its end record")

    if cursor != central_end or actual_count != expected_count:
        raise PackageLimitError("ZIP central-directory member count disagrees with its end record")


class GuardedZipReader:
    """Validate and stream every member of an already-open ``ZipFile``.

    Construction fully validates the archive and caches every member's bytes. This makes an ordinary
    ``Presentation()`` open share the same validation, including for members not reachable from OPC
    relationships.
    """

    def __init__(self, zip_file: ZipFile):
        """Wrap `zip_file`, splitting its records into all members and the subset that can be OPC
        parts.
        """
        self._zip_file = zip_file
        # -- every central-directory record, directory entries included: the physical-layout
        # -- checks derive each member's boundary from the next record's offset, so they need
        # -- the complete picture or a skipped record reads as an unexplained gap.
        self._infos = tuple(zip_file.infolist())
        # -- the subset that can be an OPC part. A directory record carries no content and no
        # -- part name can denote one, so it is ignored rather than treated as invalid input.
        self._part_infos = tuple(i for i in self._infos if not _is_directory_entry(i))
        self._validate_metadata()
        self._parts = self._read_all_members()

    @property
    def order(self) -> Tuple[str, ...]:
        """Member names in central-directory order."""
        return tuple(info.filename for info in self._part_infos)

    def read(self, name: str) -> bytes:
        """Return validated bytes for member `name`."""
        return self._parts[name]

    def read_all(self) -> Tuple[Dict[str, bytes], List[str]]:
        """Return a copy of validated parts and their archive order."""
        return dict(self._parts), list(self.order)

    def _validate_metadata(self) -> None:
        """Refuse member records the archive cannot read one way, or cannot support.

        Covers noncanonical part names (through `_validate_member_name`), duplicate and case-
        colliding names, names whose stored bytes differ from the canonical form, invalid or shared
        local-header offsets, encryption, patched-data encoding, compression outside the two the OPC
        ZIP mapping allows, non-regular filesystem entry types, negative sizes, and stored members
        whose two sizes disagree.
        """
        seen_names: set[str] = set()
        seen_equivalent_names: set[str] = set()
        seen_offsets: set[int] = set()
        for info in self._infos:
            name = info.orig_filename
            if name != info.filename:
                raise PackageLimitError(
                    f"ZIP member name {name!r} contains a noncanonical NUL suffix"
                )
            # -- a directory record is not an OPC part, so the part-name rules below do not
            # -- apply to it. It is still a physical record: the duplicate-name, offset,
            # -- encryption and compression checks that follow cover it like any other.
            directory_entry = _is_directory_entry(info)
            if not directory_entry:
                _validate_member_name(name)

            if name in seen_names:
                raise PackageLimitError(f"ZIP contains duplicate member name {name!r}")
            equivalent_name = name.casefold()
            if equivalent_name in seen_equivalent_names:
                raise PackageLimitError(f"ZIP contains case-ambiguous member name {name!r}")
            seen_names.add(name)
            seen_equivalent_names.add(equivalent_name)

            if info.header_offset < 0 or info.header_offset in seen_offsets:
                raise PackageLimitError(
                    f"ZIP member {name!r} has an invalid or shared local-header offset"
                )
            seen_offsets.add(info.header_offset)

            if info.flag_bits & (_FLAG_ENCRYPTED | _FLAG_STRONG_ENCRYPTION):
                raise PackageLimitError(f"ZIP member {name!r} is encrypted")
            if info.flag_bits & _FLAG_PATCHED_DATA:
                raise PackageLimitError(
                    f"ZIP member {name!r} uses unsupported patched-data encoding"
                )
            if info.compress_type not in _SUPPORTED_COMPRESSION:
                raise PackageLimitError(
                    f"ZIP member {name!r} uses unsupported compression method {info.compress_type}"
                )
            if directory_entry:
                continue

            unix_mode = (info.external_attr >> 16) & 0xFFFF
            file_type = stat.S_IFMT(unix_mode)
            if file_type not in (0, stat.S_IFREG):
                raise PackageLimitError(
                    f"ZIP member {name!r} uses unsupported filesystem entry type"
                )

            if info.file_size < 0 or info.compress_size < 0:
                raise PackageLimitError(f"ZIP member {name!r} has a negative size")
            if info.compress_type == ZIP_STORED and info.file_size != info.compress_size:
                raise PackageLimitError(
                    f"stored ZIP member {name!r} has inconsistent size metadata"
                )

    def _read_all_members(self) -> Dict[str, bytes]:
        """Read every part, validating physical layout as it goes.

        Each member's boundary comes from the next record's offset, so a gap or an overlap refuses
        instead of returning whatever bytes sit there. `[Content_Types].xml` is read first so the
        member list can be checked against it; an archive carrying no such member skips that check
        and fails in the OPC layer instead.
        """
        if not self._infos:
            return {}
        stream = self._zip_file.fp
        if stream is None:
            raise PackageLimitError("ZIP package stream is closed")

        start_dir = self._zip_file.start_dir
        archive_size = compressed_size(stream)
        if start_dir < 0:
            raise PackageLimitError("ZIP central-directory offset is invalid")
        if archive_size is not None and start_dir > archive_size:
            raise PackageLimitError("ZIP central directory lies beyond the package")

        sorted_infos = sorted(self._infos, key=lambda info: info.header_offset)
        boundaries = {
            info.header_offset: (
                sorted_infos[index + 1].header_offset
                if index + 1 < len(sorted_infos)
                else start_dir
            )
            for index, info in enumerate(sorted_infos)
        }

        try:
            original_position = stream.tell()
        except (AttributeError, OSError, TypeError, ValueError):
            original_position = None

        parts: Dict[str, bytes] = {}
        try:
            content_types_info = next(
                (info for info in self._part_infos if info.filename == _CONTENT_TYPES_NAME),
                None,
            )
            if content_types_info is not None:
                data_start = self._validate_local_header(
                    content_types_info,
                    boundaries[content_types_info.header_offset],
                )
                content_types = self._inflate_member(content_types_info, data_start)
                parts[content_types_info.filename] = content_types
                self._validate_content_types(content_types)

            for info in self._part_infos:
                if info is content_types_info:
                    continue
                data_start = self._validate_local_header(info, boundaries[info.header_offset])
                parts[info.filename] = self._inflate_member(info, data_start)

            # -- a directory record yields no part, but its local header is still validated
            # -- so the region it occupies cannot conceal undeclared bytes.
            for info in self._infos:
                if not _is_directory_entry(info):
                    continue
                self._validate_local_header(info, boundaries[info.header_offset])
        finally:
            if original_position is not None:
                with suppress(AttributeError, OSError, TypeError, ValueError):
                    stream.seek(original_position, os.SEEK_SET)
        return parts

    def _validate_content_types(self, content_types: bytes) -> None:
        """Refuse members that `[Content_Types].xml` gives no type, by Override or by Default
        extension.

        PowerPoint refuses such a package whatever the part is for, so accepting it here would hand
        back a deck that will not open.
        """
        defaults, overrides = _parse_content_types(content_types)

        # -- OPC gives every part a content type, by an Override naming the part or a
        # -- Default matching its extension. A member with neither has no type at all, and
        # -- PowerPoint refuses such a package whatever the part is for -- a slide, an image
        # -- it draws, or a thumbnail it never reads. Keys are normalized exactly as
        # -- `_parse_content_types` stored them.
        #
        # -- Iterate the part records, not every archive record: a directory record has no
        # -- extension and no Override, so it resolves to no content type and would be
        # -- refused here, even though PowerPoint opens packages that contain them.
        undeclared = sorted(
            info.filename
            for info in self._part_infos
            if info.filename != _CONTENT_TYPES_NAME
            and ("/" + info.filename).casefold() not in overrides
            and _member_extension(info.filename) not in defaults
        )
        if undeclared:
            raise PackageLimitError(
                "ZIP members have no content type, so their parts cannot be interpreted: "
                "%s. [Content_Types].xml declares no Default for their extension and no "
                "Override for their name; add the missing declaration or re-save the "
                "package from PowerPoint" % ", ".join(repr(name) for name in undeclared)
            )

    def _validate_local_header(self, info: ZipInfo, boundary: int) -> int:
        """Check one member's local header against its central-directory record and its boundary."""
        stream = self._zip_file.fp
        if stream is None:
            raise PackageLimitError("ZIP package stream is closed")
        if info.header_offset + _LOCAL_HEADER.size > boundary:
            raise PackageLimitError(f"ZIP member {info.filename!r} has an overlapping header")

        stream.seek(info.header_offset, os.SEEK_SET)
        header = stream.read(_LOCAL_HEADER.size)
        if len(header) != _LOCAL_HEADER.size:
            raise PackageLimitError(f"ZIP member {info.filename!r} has a truncated header")
        (
            signature,
            _extract_version,
            flags,
            compression,
            _mod_time,
            _mod_date,
            crc,
            compressed,
            expanded,
            name_length,
            extra_length,
        ) = _LOCAL_HEADER.unpack(header)
        if signature != _LOCAL_HEADER_SIGNATURE:
            raise PackageLimitError(f"ZIP member {info.filename!r} has an invalid header")
        if flags != info.flag_bits or compression != info.compress_type:
            raise PackageLimitError(
                f"ZIP member {info.filename!r} has inconsistent local-header metadata"
            )

        raw_name = stream.read(name_length)
        if len(raw_name) != name_length:
            raise PackageLimitError(f"ZIP member {info.filename!r} has a truncated name")
        try:
            local_name = raw_name.decode(
                "utf-8"
                if flags & _FLAG_UTF8_NAME
                else (getattr(self._zip_file, "metadata_encoding", None) or "cp437")
            )
        except UnicodeDecodeError as exc:
            raise PackageLimitError(
                f"ZIP member {info.filename!r} has an invalid encoded name"
            ) from exc
        if local_name != info.orig_filename:
            raise PackageLimitError(
                f"ZIP member {info.filename!r} has conflicting local and central names"
            )

        if flags & _FLAG_DATA_DESCRIPTOR:
            valid_crc = crc in (0, info.CRC)
            valid_compressed = compressed in (0, info.compress_size, 0xFFFFFFFF)
            valid_expanded = expanded in (0, info.file_size, 0xFFFFFFFF)
        else:
            valid_crc = crc == info.CRC
            valid_compressed = compressed in (info.compress_size, 0xFFFFFFFF)
            valid_expanded = expanded in (info.file_size, 0xFFFFFFFF)
        if not (valid_crc and valid_compressed and valid_expanded):
            raise PackageLimitError(
                f"ZIP member {info.filename!r} has inconsistent local size or CRC metadata"
            )

        data_start = info.header_offset + _LOCAL_HEADER.size + name_length + extra_length
        data_end = data_start + info.compress_size
        if data_start > boundary or data_end > boundary:
            raise PackageLimitError(f"ZIP member {info.filename!r} overlaps another ZIP record")
        self._validate_data_descriptor(info, data_end, boundary)
        return data_start

    def _validate_data_descriptor(self, info: ZipInfo, data_end: int, boundary: int) -> None:
        """Check a member's trailing data descriptor, refusing bytes between members that nothing
        declares.
        """
        stream = self._zip_file.fp
        if stream is None:
            raise PackageLimitError("ZIP package stream is closed")
        descriptor_size = boundary - data_end
        if not info.flag_bits & _FLAG_DATA_DESCRIPTOR:
            if descriptor_size:
                raise PackageLimitError(
                    f"ZIP member {info.filename!r} has undeclared trailing data"
                )
            return
        if descriptor_size not in (12, 16, 20, 24):
            raise PackageLimitError(f"ZIP member {info.filename!r} has an invalid data descriptor")

        stream.seek(data_end, os.SEEK_SET)
        descriptor = stream.read(descriptor_size)
        if len(descriptor) != descriptor_size:
            raise PackageLimitError(f"ZIP member {info.filename!r} has a truncated data descriptor")
        has_signature = descriptor_size in (16, 24)
        if has_signature:
            if descriptor[:4] != b"PK\x07\x08":
                raise PackageLimitError(
                    f"ZIP member {info.filename!r} has an invalid data descriptor"
                )
            descriptor = descriptor[4:]
        if len(descriptor) == 12:
            crc, compressed, expanded = struct.unpack("<III", descriptor)
        else:
            crc, compressed, expanded = struct.unpack("<IQQ", descriptor)
        if (crc, compressed, expanded) != (
            info.CRC,
            info.compress_size,
            info.file_size,
        ):
            raise PackageLimitError(
                f"ZIP member {info.filename!r} has inconsistent data-descriptor metadata"
            )

    def _inflate_member(self, info: ZipInfo, data_start: int) -> bytes:
        """Inflate one member from its raw bytes, checking CRC and declared size.

        Reads the compressed bytes directly rather than through `ZipFile.read`, so a size header
        that lies cannot pass unnoticed.
        """
        stream = self._zip_file.fp
        if stream is None:
            raise PackageLimitError("ZIP package stream is closed")
        stream.seek(data_start, os.SEEK_SET)

        chunks: List[bytes] = []
        actual_size = 0
        crc = 0

        def consume(data: bytes) -> None:
            nonlocal actual_size, crc
            if not data:
                return
            actual_size += len(data)
            crc = zlib.crc32(data, crc)
            chunks.append(data)

        remaining = info.compress_size
        if info.compress_type == ZIP_STORED:
            while remaining:
                raw = stream.read(min(_READ_CHUNK_BYTES, remaining))
                if not raw:
                    raise PackageLimitError(
                        f"ZIP member {info.filename!r} has truncated stored data"
                    )
                remaining -= len(raw)
                consume(raw)
        else:
            decompressor = zlib.decompressobj(-zlib.MAX_WBITS)
            try:
                while remaining:
                    raw = stream.read(min(_READ_CHUNK_BYTES, remaining))
                    if not raw:
                        raise PackageLimitError(
                            f"ZIP member {info.filename!r} has truncated compressed data"
                        )
                    remaining -= len(raw)
                    pending = raw
                    while pending:
                        if decompressor.eof:
                            raise PackageLimitError(
                                f"ZIP member {info.filename!r} has trailing compressed data"
                            )
                        output = decompressor.decompress(pending, _READ_CHUNK_BYTES)
                        pending = decompressor.unconsumed_tail
                        consume(output)
                        if decompressor.unused_data:
                            raise PackageLimitError(
                                f"ZIP member {info.filename!r} has trailing compressed data"
                            )

                while not decompressor.eof:
                    output = decompressor.decompress(b"", _READ_CHUNK_BYTES)
                    consume(output)
                    if not output:
                        break
            except zlib.error as exc:
                raise PackageLimitError(
                    f"ZIP member {info.filename!r} has invalid deflate data"
                ) from exc
            if not decompressor.eof:
                raise PackageLimitError(
                    f"ZIP member {info.filename!r} compressed data ends before deflate EOF"
                )

        if actual_size != info.file_size:
            raise PackageLimitError(
                f"ZIP member {info.filename!r} actual expanded size {actual_size} bytes "
                f"does not match declared size {info.file_size} bytes"
            )
        if crc & 0xFFFFFFFF != info.CRC:
            raise PackageLimitError(f"ZIP member {info.filename!r} fails its CRC check")
        return b"".join(chunks)


def _member_extension(name: str) -> str:
    """Return `name`'s extension keyed as ``Default`` declarations are stored.

    Empty when the final segment carries no period, which no ``Default`` can match.
    """
    leaf = name.rsplit("/", 1)[-1]
    return leaf.rsplit(".", 1)[-1].lower() if "." in leaf else ""


def _parse_content_types(data: bytes) -> Tuple[Dict[str, str], Dict[str, str]]:
    """Parse `[Content_Types].xml` into its Default and Override maps.

    Refuses a malformed document, a DTD, an unexpected root element, an invalid or ambiguous Default
    or Override declaration, and any other child element.
    """
    try:
        root = etree.fromstring(data, _CONTENT_TYPES_PARSER)
    except etree.XMLSyntaxError as exc:
        raise PackageLimitError("[Content_Types].xml is malformed") from exc
    docinfo = root.getroottree().docinfo
    if docinfo.doctype or docinfo.internalDTD is not None:
        raise PackageLimitError("[Content_Types].xml contains a prohibited DTD")
    if root.tag != _CONTENT_TYPES_TAG:
        raise PackageLimitError("[Content_Types].xml has an unexpected root element")

    defaults: Dict[str, str] = {}
    overrides: Dict[str, str] = {}
    for element in root:
        if not isinstance(element.tag, str):
            continue
        if element.tag == _DEFAULT_TAG:
            extension = element.get("Extension")
            content_type = element.get("ContentType")
            if (
                not extension
                or "." in extension
                or "/" in extension
                or extension != extension.strip()
                or not content_type
                or content_type != content_type.strip()
            ):
                raise PackageLimitError(
                    "[Content_Types].xml contains an invalid Default declaration"
                )
            key = extension.lower()
            if key in defaults:
                raise PackageLimitError(
                    "[Content_Types].xml contains an ambiguous Default declaration"
                )
            defaults[key] = content_type
            continue
        if element.tag == _OVERRIDE_TAG:
            part_name = element.get("PartName")
            content_type = element.get("ContentType")
            if (
                not part_name
                or not part_name.startswith("/")
                or part_name == "/"
                or part_name != part_name.strip()
                or not content_type
                or content_type != content_type.strip()
            ):
                raise PackageLimitError(
                    "[Content_Types].xml contains an invalid Override declaration"
                )
            key = part_name.casefold()
            if key in overrides:
                raise PackageLimitError(
                    "[Content_Types].xml contains an ambiguous Override declaration"
                )
            overrides[key] = content_type
            continue
        raise PackageLimitError(
            "[Content_Types].xml contains an unsupported declaration"
        )
    return defaults, overrides


def _is_directory_entry(info: ZipInfo) -> bool:
    """True for a ZIP directory record, which is never an OPC part.

    A directory record carries no content and no part name can denote one, so a reader
    ignores it rather than treating the package as invalid. Producers that emit them by
    default include the ``zip`` CLI, ``shutil.make_archive`` and Java's
    ``ZipOutputStream`` -- every "unzip, edit, rezip" pipeline passes through one.

    The trailing slash is the whole test. A record carrying only the MS-DOS directory
    attribute, with no slash, is deliberately NOT treated as a directory: such a name
    denotes a file, and a package holding both a file ``ppt`` and members under ``ppt/``
    contradicts itself. PowerPoint refuses that file, so admitting it would help nobody.

    The empty-name guard keeps this callable before ``_validate_metadata`` runs: on Python
    3.9 and 3.10 ``ZipInfo.is_dir`` reads ``filename[-1]`` and raises ``IndexError`` on an
    empty name, which would escape as an untyped crash before the empty-name refusal fires.
    An empty name is never a directory record, so returning False here defers it to
    ``_validate_member_name``, which reports it as the ``PackageLimitError`` it is. On 3.11+
    ``is_dir`` already returns False for an empty name, so this only aligns 3.9/3.10.
    """
    return bool(info.filename) and info.is_dir()


def _validate_member_name(name: str) -> None:
    """Refuse a member name that cannot denote one unambiguous OPC part.

    Covers empty names, names that are not NFC-normalized, leading or trailing slashes, backslashes,
    URI query or fragment characters, control characters, `.` and `..` path segments, drive-
    qualified paths, and percent escapes that are malformed or unsafe.
    """
    if not name:
        raise PackageLimitError("ZIP contains an empty member name")
    if name != unicodedata.normalize("NFC", name):
        raise PackageLimitError(f"ZIP member name {name!r} is not Unicode-normalized")
    if name.startswith("/") or name.endswith("/") or "\\" in name:
        raise PackageLimitError(f"ZIP member name {name!r} is noncanonical")
    if "?" in name or "#" in name:
        raise PackageLimitError(f"ZIP member name {name!r} contains a URI query or fragment")
    if any(ord(char) < 0x20 or ord(char) == 0x7F for char in name):
        raise PackageLimitError(f"ZIP member name {name!r} contains a control character")

    segments = name.split("/")
    if any(segment in ("", ".", "..") for segment in segments):
        raise PackageLimitError(f"ZIP member name {name!r} has a noncanonical path segment")
    first_segment = segments[0]
    if len(first_segment) >= 2 and first_segment[0].isalpha() and first_segment[1] == ":":
        raise PackageLimitError(f"ZIP member name {name!r} has a drive-qualified path")

    index = 0
    while index < len(name):
        if name[index] != "%":
            index += 1
            continue
        if index + 2 >= len(name) or any(
            char not in _HEX_DIGITS for char in name[index + 1 : index + 3]
        ):
            raise PackageLimitError(f"ZIP member name {name!r} has a noncanonical percent escape")
        value = int(name[index + 1 : index + 3], 16)
        if value in _URI_UNRESERVED or value in (0x00, 0x2F, 0x5C, 0x7F) or value < 0x20:
            raise PackageLimitError(f"ZIP member name {name!r} has an unsafe percent escape")
        index += 3
