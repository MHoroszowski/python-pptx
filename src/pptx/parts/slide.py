"""Slide and related objects."""

from __future__ import annotations

import copy
import re
from typing import IO, TYPE_CHECKING, cast

from pptx.enum.shapes import PROG_ID
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.slide import CT_NotesMaster, CT_NotesSlide, CT_Slide
from pptx.oxml.theme import CT_OfficeStyleSheet
from pptx.parts.chart import ChartPart
from pptx.parts.embeddedpackage import EmbeddedPackagePart
from pptx.slide import NotesMaster, NotesSlide, Slide, SlideLayout, SlideMaster
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.media import Video
    from pptx.parts.image import Image, ImagePart


class BaseSlidePart(XmlPart):
    """Base class for slide parts.

    This includes slide, slide-layout, and slide-master parts, but also notes-slide,
    notes-master, and handout-master parts.
    """

    _element: CT_Slide

    def get_image(self, rId: str) -> Image:
        """Return an |Image| object containing the image related to this slide by *rId*.

        Raises |KeyError| if no image is related by that id, which would generally indicate a
        corrupted .pptx file.
        """
        return cast("ImagePart", self.related_part(rId)).image

    def get_or_add_image_part(self, image_file: str | IO[bytes]):
        """Return `(image_part, rId)` pair corresponding to `image_file`.

        The returned |ImagePart| object contains the image in `image_file` and is
        related to this slide with the key `rId`. If either the image part or
        relationship already exists, they are reused, otherwise they are newly created.
        """
        image_part = self._package.get_or_add_image_part(image_file)
        rId = self.relate_to(image_part, RT.IMAGE)
        return image_part, rId

    @property
    def name(self) -> str:
        """Internal name of this slide."""
        return self._element.cSld.name


class NotesMasterPart(BaseSlidePart):
    """Notes master part.

    Corresponds to package file `ppt/notesMasters/notesMaster1.xml`.
    """

    @classmethod
    def create_default(cls, package):
        """
        Create and return a default notes master part, including creating the
        new theme it requires.
        """
        notes_master_part = cls._new(package)
        theme_part = cls._new_theme_part(package)
        notes_master_part.relate_to(theme_part, RT.THEME)
        return notes_master_part

    @lazyproperty
    def notes_master(self):
        """
        Return the |NotesMaster| object that proxies this notes master part.
        """
        return NotesMaster(self._element, self)

    @classmethod
    def _new(cls, package):
        """
        Create and return a standalone, default notes master part based on
        the built-in template (without any related parts, such as theme).
        """
        return NotesMasterPart(
            PackURI("/ppt/notesMasters/notesMaster1.xml"),
            CT.PML_NOTES_MASTER,
            package,
            CT_NotesMaster.new_default(),
        )

    @classmethod
    def _new_theme_part(cls, package):
        """Return new default theme-part suitable for use with a notes master."""
        return XmlPart(
            package.next_partname("/ppt/theme/theme%d.xml"),
            CT.OFC_THEME,
            package,
            CT_OfficeStyleSheet.new_default(),
        )


class NotesSlidePart(BaseSlidePart):
    """Notes slide part.

    Contains the slide notes content and the layout for the slide handout page.
    Corresponds to package file `ppt/notesSlides/notesSlide[1-9][0-9]*.xml`.
    """

    @classmethod
    def new(cls, package, slide_part):
        """Return new |NotesSlidePart| for the slide in `slide_part`.

        The new notes-slide part is based on the (singleton) notes master and related to
        both the notes-master part and `slide_part`. If no notes-master is present,
        one is created based on the default template.
        """
        notes_master_part = package.presentation_part.notes_master_part
        notes_slide_part = cls._add_notes_slide_part(package, slide_part, notes_master_part)
        notes_slide = notes_slide_part.notes_slide
        notes_slide.clone_master_placeholders(notes_master_part.notes_master)
        return notes_slide_part

    @lazyproperty
    def notes_master(self):
        """Return the |NotesMaster| object this notes slide inherits from."""
        notes_master_part = self.part_related_by(RT.NOTES_MASTER)
        return notes_master_part.notes_master

    @lazyproperty
    def notes_slide(self):
        """Return the |NotesSlide| object that proxies this notes slide part."""
        return NotesSlide(self._element, self)

    @classmethod
    def _add_notes_slide_part(cls, package, slide_part, notes_master_part):
        """Create and return a new notes-slide part.

        The return part is fully related, but has no shape content (i.e. placeholders
        not cloned).
        """
        notes_slide_part = NotesSlidePart(
            package.next_partname("/ppt/notesSlides/notesSlide%d.xml"),
            CT.PML_NOTES_SLIDE,
            package,
            CT_NotesSlide.new(),
        )
        notes_slide_part.relate_to(notes_master_part, RT.NOTES_MASTER)
        notes_slide_part.relate_to(slide_part, RT.SLIDE)
        return notes_slide_part


class SlidePart(BaseSlidePart):
    """Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml."""

    @classmethod
    def new(cls, partname, package, slide_layout_part):
        """Return newly-created blank slide part.

        The new slide-part has `partname` and a relationship to `slide_layout_part`.
        """
        slide_part = cls(partname, CT.PML_SLIDE, package, CT_Slide.new())
        slide_part.relate_to(slide_layout_part, RT.SLIDE_LAYOUT)
        return slide_part

    def add_chart_part(self, chart_type: XL_CHART_TYPE, chart_data: ChartData):
        """Return str rId of new |ChartPart| object containing chart of `chart_type`.

        The chart depicts `chart_data` and is related to the slide contained in this
        part by `rId`.
        """
        return self.relate_to(ChartPart.new(chart_type, chart_data, self._package), RT.CHART)

    def add_embedded_ole_object_part(
        self, prog_id: PROG_ID | str, ole_object_file: str | IO[bytes]
    ):
        """Return rId of newly-added OLE-object part formed from `ole_object_file`."""
        relationship_type = RT.PACKAGE if isinstance(prog_id, PROG_ID) else RT.OLE_OBJECT
        return self.relate_to(
            EmbeddedPackagePart.factory(
                prog_id, self._blob_from_file(ole_object_file), self._package
            ),
            relationship_type,
        )

    def get_or_add_video_media_part(self, video: Video) -> tuple[str, str]:
        """Return rIds for media and video relationships to media part.

        A new |MediaPart| object is created if it does not already exist
        (such as would occur if the same video appeared more than once in
         a presentation). Two relationships to the media part are created,
        one each with MEDIA and VIDEO relationship types. The need for two
        appears to be for legacy support for an earlier (pre-Office 2010)
        PowerPoint media embedding strategy.
        """
        media_part = self._package.get_or_add_media_part(video)
        media_rId = self.relate_to(media_part, RT.MEDIA)
        video_rId = self.relate_to(media_part, RT.VIDEO)
        return media_rId, video_rId

    @property
    def has_notes_slide(self):
        """
        Return True if this slide has a notes slide, False otherwise. A notes
        slide is created by the :attr:`notes_slide` property when one doesn't
        exist; use this property to test for a notes slide without the
        possible side-effect of creating one.
        """
        try:
            self.part_related_by(RT.NOTES_SLIDE)
        except KeyError:
            return False
        return True

    @lazyproperty
    def notes_slide(self) -> NotesSlide:
        """The |NotesSlide| instance associated with this slide.

        If the slide does not have a notes slide, a new one is created. The same single instance
        is returned on each call.
        """
        try:
            notes_slide_part = self.part_related_by(RT.NOTES_SLIDE)
        except KeyError:
            notes_slide_part = self._add_notes_slide_part()
        return notes_slide_part.notes_slide

    @lazyproperty
    def slide(self):
        """
        The |Slide| object representing this slide part.
        """
        return Slide(self._element, self)

    @property
    def slide_id(self) -> int:
        """Return the slide identifier stored in the presentation part for this slide part."""
        presentation_part = self.package.presentation_part
        return presentation_part.slide_id(self)

    @property
    def slide_layout(self) -> SlideLayout:
        """|SlideLayout| object the slide in this part inherits appearance from."""
        slide_layout_part = self.part_related_by(RT.SLIDE_LAYOUT)
        return slide_layout_part.slide_layout

    def _add_notes_slide_part(self):
        """
        Return a newly created |NotesSlidePart| object related to this slide
        part. Caller is responsible for ensuring this slide doesn't already
        have a notes slide part.
        """
        notes_slide_part = NotesSlidePart.new(self.package, self)
        self.relate_to(notes_slide_part, RT.NOTES_SLIDE)
        return notes_slide_part

    def duplicate(self) -> SlidePart:
        """Return a new |SlidePart| that is a deep copy of this one.

        Image, media, slide-layout, and slide-master rels are reused —
        the duplicate references the same package-level parts as the
        source. Chart, OLE-embedded, and embedded-package parts are
        deep-copied per duplicate. The notes-slide rel and any
        comments rels are NOT carried over: notes-slide rewiring is
        the caller's job (see |Slides.duplicate|), and comments are
        out of scope for Phase 2 of issue #11.
        """
        new_partname = self._package.next_partname("/ppt/slides/slide%d.xml")
        new_element = copy.deepcopy(self._element)
        new_part = SlidePart(new_partname, CT.PML_SLIDE, self._package, new_element)

        rId_map = _replicate_rels_for_duplicate(self, new_part)
        _remap_rId_attrs(new_element, rId_map)

        return new_part


# ---------------------------------------------------------------------------
# Module-level helpers for slide / slide-private part duplication.
# ---------------------------------------------------------------------------

_RELS_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# Reltypes filtered out during slide duplication. NOTES_SLIDE is wired
# explicitly by |Slides.duplicate| so the new notes-slide back-references
# the new parent slide. Comments are dropped — Phase 2 scope (issue #11).
_DUP_DROP_RELTYPES_SLIDE = frozenset({RT.NOTES_SLIDE, RT.COMMENTS, RT.COMMENT_AUTHORS})


def _replicate_rels_for_duplicate(src_part: Part, new_part: Part) -> dict[str, str]:
    """Mirror src_part's slide-relevant rels onto new_part.

    Returns a `{old_rId: new_rId}` map for rId-attribute remapping.
    """
    rId_map: dict[str, str] = {}
    for rId, rel in src_part.rels.items():
        if rel.reltype in _DUP_DROP_RELTYPES_SLIDE:
            continue
        if rel.is_external:
            new_rId = new_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif rel.reltype == RT.CHART:
            new_target = _duplicate_chart_part(cast(ChartPart, rel.target_part))
            new_rId = new_part.relate_to(new_target, rel.reltype)
        elif rel.reltype in (RT.OLE_OBJECT, RT.PACKAGE):
            new_target = _duplicate_blob_part(cast(Part, rel.target_part))
            new_rId = new_part.relate_to(new_target, rel.reltype)
        else:
            # Shared parts: image, media, video, layout, master, theme, etc.
            new_rId = new_part.relate_to(rel.target_part, rel.reltype)
        rId_map[rId] = new_rId
    return rId_map


def _remap_rId_attrs(element, rId_map: dict[str, str]) -> None:
    """Substitute relationships-namespace attribute values in `element`.

    Walks every descendant element and rewrites any attribute whose name
    is in the OOXML relationships namespace (catches `r:id`, `r:embed`,
    `r:link`, `r:pict`, `r:href` in one pass).
    """
    for el in element.iter():
        for attr_name in list(el.attrib):
            if attr_name.startswith(_RELS_NS):
                old = el.attrib[attr_name]
                if old in rId_map:
                    el.attrib[attr_name] = rId_map[old]


def _duplicate_chart_part(src: ChartPart) -> ChartPart:
    """Return a new ChartPart cloning `src`.

    Chart XML is deep-copied. Embedded data (e.g. an xlsx workbook
    reached via an `RT.PACKAGE` rel) is binary and must be blob-copied,
    not deep-copy-of-XML — the workbook IS the chart's data, and the
    `<c:numCache>` values in the chart XML mirror it.
    """
    package = src._package
    new_partname = package.next_partname("/ppt/charts/chart%d.xml")
    new_element = copy.deepcopy(src._element)
    cls = type(src)
    new_part = cls(new_partname, src.content_type, package, new_element)
    rId_map: dict[str, str] = {}
    for rId, rel in src.rels.items():
        if rel.is_external:
            new_rId = new_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif rel.reltype == RT.PACKAGE:
            new_target = _duplicate_blob_part(cast(Part, rel.target_part))
            new_rId = new_part.relate_to(new_target, rel.reltype)
        else:
            # Theme override and other chart-private parts: share for now.
            # Practical impact is small; revisit if a user reports it.
            new_rId = new_part.relate_to(rel.target_part, rel.reltype)
        rId_map[rId] = new_rId
    _remap_rId_attrs(new_element, rId_map)
    return new_part


def _duplicate_blob_part(src: Part) -> Part:
    """Return a new binary |Part| cloning `src`'s blob.

    Used for embedded packages (xlsx, docx, pptx) and OLE objects —
    parts whose payload is opaque bytes rather than XML.
    """
    package = src._package
    cls = type(src)
    tmpl = getattr(cls, "partname_template", None)
    if tmpl is None:
        tmpl = _derive_partname_template(str(src.partname))
    new_partname = package.next_partname(tmpl)
    return cls(new_partname, src.content_type, package, src.blob)


def _derive_partname_template(partname: str) -> str:
    """Derive a `next_partname`-compatible template from an existing partname.

    Replaces the trailing integer (just before the final extension) with
    `%d`. Falls back to inserting `%d` immediately before the extension
    if there is no trailing digit run.
    """
    match = re.match(r"^(.*?)(\d+)(\.[^./]+)$", partname)
    if match:
        prefix, _, ext = match.groups()
        return f"{prefix}%d{ext}"
    # No trailing-digit pattern; insert %d before final extension.
    dot = partname.rfind(".")
    if dot < 0:
        return f"{partname}%d"
    return f"{partname[:dot]}%d{partname[dot:]}"


def duplicate_notes_slide_for(
    src_slide_part: SlidePart, new_slide_part: SlidePart
) -> NotesSlidePart:
    """Create a fresh |NotesSlidePart| for `new_slide_part`, cloning content from src.

    Public-to-the-module helper used by |Slides.duplicate| AFTER the new
    slide part is registered with the presentation rels. Wires the new
    notes-slide's `RT.SLIDE` back-rel to point at `new_slide_part` (NOT
    the source) — addresses upstream community gotcha #961 where blindly
    copying notes rels left the duplicate's notes pointing at the source.
    """
    src_notes_part = cast(NotesSlidePart, src_slide_part.part_related_by(RT.NOTES_SLIDE))
    package = src_slide_part._package
    new_partname = package.next_partname("/ppt/notesSlides/notesSlide%d.xml")
    new_element = copy.deepcopy(src_notes_part._element)
    new_notes_part = NotesSlidePart(new_partname, CT.PML_NOTES_SLIDE, package, new_element)

    rId_map: dict[str, str] = {}
    for rId, rel in src_notes_part.rels.items():
        if rel.is_external:
            new_rId = new_notes_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif rel.reltype == RT.SLIDE:
            # ---rewire back-ref to NEW slide part---
            new_rId = new_notes_part.relate_to(new_slide_part, RT.SLIDE)
        else:
            # NOTES_MASTER and any others: share at package level
            new_rId = new_notes_part.relate_to(rel.target_part, rel.reltype)
        rId_map[rId] = new_rId
    _remap_rId_attrs(new_element, rId_map)

    new_slide_part.relate_to(new_notes_part, RT.NOTES_SLIDE)
    return new_notes_part


class SlideLayoutPart(BaseSlidePart):
    """Slide layout part.

    Corresponds to package files ``ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``.
    """

    @lazyproperty
    def slide_layout(self):
        """
        The |SlideLayout| object representing this part.
        """
        return SlideLayout(self._element, self)

    @property
    def slide_master(self) -> SlideMaster:
        """Slide master from which this slide layout inherits properties."""
        return self.part_related_by(RT.SLIDE_MASTER).slide_master


class SlideMasterPart(BaseSlidePart):
    """Slide master part.

    Corresponds to package files ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    """

    def related_slide_layout(self, rId: str) -> SlideLayout:
        """Return |SlideLayout| related to this slide-master by key `rId`."""
        return self.related_part(rId).slide_layout

    @lazyproperty
    def slide_master(self):
        """
        The |SlideMaster| object representing this part.
        """
        return SlideMaster(self._element, self)
