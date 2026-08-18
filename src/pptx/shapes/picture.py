"""Shapes based on the `p:pic` element, including Picture and Movie."""

from __future__ import annotations

import io
from typing import IO, TYPE_CHECKING

from pptx.dml.line import LineFormat
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, PP_MEDIA_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.shapes.base import BaseShape
from pptx.shared import ParentedElementProxy
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.shapes.picture import CT_Picture
    from pptx.oxml.shapes.shared import CT_LineProperties
    from pptx.types import ProvidesPart


def _canonical_image_ext(ext: str) -> str:
    """Return lowercase canonical form of image extension `ext` (jpg == jpeg)."""
    lowered = ext.lower()
    return "jpg" if lowered == "jpeg" else lowered


_R_NS_PREFIX = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _part_xml_references_rId(root, rId: str) -> bool:
    """True when any r-namespace attribute anywhere in `root` still carries `rId`."""
    for element in root.iter():
        for attr_name, attr_value in element.attrib.items():
            if attr_value == rId and attr_name.startswith(_R_NS_PREFIX):
                return True
    return False


class _BasePicture(BaseShape):
    """Base class for shapes based on a `p:pic` element."""

    def __init__(self, pic: CT_Picture, parent: ProvidesPart):
        super(_BasePicture, self).__init__(pic, parent)
        self._pic = pic

    @property
    def crop_bottom(self) -> float:
        """`float` representing relative portion cropped from shape bottom.

        Read/write. 1.0 represents 100%. For example, 25% is represented by 0.25. Negative values
        are valid as are values greater than 1.0.
        """
        return self._pic.srcRect_b

    @crop_bottom.setter
    def crop_bottom(self, value: float):
        self._pic.srcRect_b = value

    @property
    def crop_left(self) -> float:
        """`float` representing relative portion cropped from left of shape.

        Read/write. 1.0 represents 100%. A negative value extends the side beyond the image
        boundary.
        """
        return self._pic.srcRect_l

    @crop_left.setter
    def crop_left(self, value: float):
        self._pic.srcRect_l = value

    @property
    def crop_right(self) -> float:
        """`float` representing relative portion cropped from right of shape.

        Read/write. 1.0 represents 100%.
        """
        return self._pic.srcRect_r

    @crop_right.setter
    def crop_right(self, value: float):
        self._pic.srcRect_r = value

    @property
    def crop_top(self) -> float:
        """`float` representing relative portion cropped from shape top.

        Read/write. 1.0 represents 100%.
        """
        return self._pic.srcRect_t

    @crop_top.setter
    def crop_top(self, value: float):
        self._pic.srcRect_t = value

    def get_or_add_ln(self):
        """Return the `a:ln` element for this `p:pic`-based image.

        The `a:ln` element contains the line format properties XML.
        """
        return self._pic.get_or_add_ln()

    @lazyproperty
    def line(self) -> LineFormat:
        """Provides access to properties of the picture outline, such as its color and width."""
        return LineFormat(self)

    @property
    def ln(self) -> CT_LineProperties | None:
        """The `a:ln` element for this `p:pic`.

        Contains the line format properties such as line color and width. `None` if no `a:ln`
        element is present.
        """
        return self._pic.ln


class Movie(_BasePicture):
    """A movie shape, one that places a video on a slide.

    Like `Picture`, a movie shape is based on the `p:pic` element. A movie is composed of a video
    and a *poster frame*, the placeholder image that represents the video before it is played.
    """

    @lazyproperty
    def media_format(self) -> _MediaFormat:
        """The `_MediaFormat` object for this movie.

        The `_MediaFormat` object provides access to formatting properties for the movie.
        """
        return _MediaFormat(self._pic, self)

    @property
    def media_type(self) -> PP_MEDIA_TYPE:
        """Member of `PpMediaType` describing this shape.

        The return value is unconditionally `PP_MEDIA_TYPE.MOVIE` in this case.
        """
        return PP_MEDIA_TYPE.MOVIE

    @property
    def poster_frame(self):
        """Return `Image` object containing poster frame for this movie.

        Returns `None` if this movie has no poster frame (uncommon).
        """
        slide_part, rId = self.part, self._pic.blip_rId
        if rId is None:
            return None
        return slide_part.get_image(rId)

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        """Return member of `MsoShapeType` describing this shape.

        The return value is unconditionally `MSO_SHAPE_TYPE.MEDIA` in this
        case.
        """
        return MSO_SHAPE_TYPE.MEDIA


class Picture(_BasePicture):
    """A picture shape, one that places an image on a slide.

    Based on the `p:pic` element.
    """

    def replace_image(
        self, image_file: str | IO[bytes], *, allow_format_change: bool = False
    ) -> None:
        """Replace the image behind this picture, preserving its geometry exactly.

        paper-pptx addition. Position, size, rotation, masking geometry, and crop
        (`a:srcRect`) are not touched — only the `a:blip/@r:embed` target changes. By
        default the new image's canonical format must match the existing image part's
        extension (jpg == jpeg); a mismatch refuses with `UnsupportedStructureError`.
        Passing `allow_format_change=True` permits a cross-format swap: the new
        image gets its own correctly-typed part and `[Content_Types].xml` follows
        automatically at save (it is regenerated from live parts).

        A picture with no embedded image relationship (e.g. linked-only) refuses. The new
        image part is deduplicated package-wide by content hash; the old image part simply
        becomes unreferenced when this picture held its last reference (an unreachable part
        is never serialized).
        """
        from pptx._ownership import require_shape_attached
        from pptx.errors import UnsupportedStructureError
        from pptx.parts.image import Image

        # -- validation pass, complete before any mutation --
        require_shape_attached(self, argument="picture")
        if not isinstance(allow_format_change, bool):
            raise ValueError(
                "allow_format_change must be a bool, got %r" % (allow_format_change,)
            )
        old_rId = self._pic.blip_rId
        if old_rId is None:
            raise UnsupportedStructureError(
                "picture %r has no embedded image relationship (r:embed); a linked-only"
                " image cannot be replaced" % self.name
            )
        try:
            old_part = self.part.related_part(old_rId)
        except KeyError:
            raise UnsupportedStructureError(
                "picture %r references image relationship %s which does not exist"
                % (self.name, old_rId)
            )
        try:
            # -- image parsing is lazy: .ext is what forces PIL to sniff the bytes --
            new_image = Image.from_file(image_file)
            new_ext = _canonical_image_ext(new_image.ext)
        except OSError as e:
            raise ValueError("image_file is not a recognizable image: %s" % e)
        old_ext = _canonical_image_ext(old_part.partname.ext)
        if old_ext != new_ext and not allow_format_change:
            raise UnsupportedStructureError(
                "replacement image format %r does not match existing image part format %r;"
                " pass allow_format_change=True to swap across formats" % (new_ext, old_ext)
            )

        # -- mutation --
        from pptx._transaction import PackageTransaction

        with PackageTransaction(self.part.package, self):
            image_part = self.part.package.get_or_add_image_part(io.BytesIO(new_image.blob))
            new_rId = self.part.relate_to(image_part, RT.IMAGE)
            if new_rId == old_rId:
                return  # -- identical image bytes: already in place
            self._pic.blipFill.blip.rEmbed = new_rId
            # -- another shape on this slide may still reference old_rId (pictures added from
            # -- identical bytes share one relationship); XmlPart.drop_rel only counts @r:id
            # -- references, so guard with a scan over ALL r-namespace attributes.
            if not _part_xml_references_rId(self.part._element, old_rId):
                self.part.drop_rel(old_rId)

    @property
    def auto_shape_type(self) -> MSO_SHAPE | None:
        """Member of MSO_SHAPE indicating masking shape.

        A picture can be masked by any of the so-called "auto-shapes" available in PowerPoint,
        such as an ellipse or triangle. When a picture is masked by a shape, the shape assumes the
        same dimensions as the picture and the portion of the picture outside the shape boundaries
        does not appear. Note the default value for a newly-inserted picture is
        `MSO_AUTO_SHAPE_TYPE.RECTANGLE`, which performs no cropping because the extents of the
        rectangle exactly correspond to the extents of the picture.

        The available shapes correspond to the members of `MsoAutoShapeType`.

        The return value can also be `None`, indicating the picture either has no geometry (not
        expected) or has custom geometry, like a freeform shape. A picture with no geometry will
        have no visible representation on the slide, although it can be selected. This is because
        without geometry, there is no "inside-the-shape" for it to appear in.
        """
        prstGeom = self._pic.spPr.prstGeom
        if prstGeom is None:  # ---generally means cropped with freeform---
            return None
        return prstGeom.prst

    @auto_shape_type.setter
    def auto_shape_type(self, member: MSO_SHAPE):
        MSO_SHAPE.validate(member)
        spPr = self._pic.spPr
        prstGeom = spPr.prstGeom
        if prstGeom is None:
            spPr._remove_custGeom()  # pyright: ignore[reportPrivateUsage]
            prstGeom = spPr._add_prstGeom()  # pyright: ignore[reportPrivateUsage]
        prstGeom.prst = member

    @property
    def image(self):
        """The `Image` object for this picture.

        Provides access to the properties and bytes of the image in this picture shape.
        """
        slide_part, rId = self.part, self._pic.blip_rId
        if rId is None:
            raise ValueError("no embedded image")
        return slide_part.get_image(rId)

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        """Unconditionally `MSO_SHAPE_TYPE.PICTURE` in this case."""
        return MSO_SHAPE_TYPE.PICTURE


class _MediaFormat(ParentedElementProxy):
    """Provides access to formatting properties for a Media object.

    Media format properties are things like start point, volume, and
    compression type.
    """
