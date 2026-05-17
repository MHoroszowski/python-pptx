"""Slide and related objects."""

from __future__ import annotations

import contextlib
import copy
import re
import secrets
from io import BytesIO
from typing import IO, TYPE_CHECKING, cast

from pptx.enum.shapes import PROG_ID
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.slide import CT_NotesMaster, CT_NotesSlide, CT_Slide, CT_SlideLayout
from pptx.oxml.theme import CT_OfficeStyleSheet
from pptx.parts.chart import ChartPart
from pptx.parts.embeddedpackage import EmbeddedPackagePart
from pptx.parts.image import Image, ImagePart
from pptx.slide import HandoutMaster, NotesMaster, NotesSlide, Slide, SlideLayout, SlideMaster
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.parts.presentation import PresentationPart

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


class HandoutMasterPart(BaseSlidePart):
    """Handout master part.

    Corresponds to package file `ppt/handoutMasters/handoutMaster1.xml` when present.
    Auto-create is deliberately deferred until this fork ships a built-in handout-master
    template and theme wiring.
    """

    @lazyproperty
    def handout_master(self):
        """Return the |HandoutMaster| object that proxies this handout master part."""
        return HandoutMaster(self._element, self)


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

    def apply_slide_layout(self, slide_layout_part: SlideLayoutPart) -> None:
        """Repoint this slide's ``SLIDE_LAYOUT`` relationship at `slide_layout_part`.

        The existing slide→layout relationship is dropped and a fresh one to
        `slide_layout_part` is created. Dropping the rel only removes *this
        slide's* edge to the prior layout — the prior layout and its master
        remain in the package (other slides may still reference them, and the
        master still lists the layout in its ``p:sldLayoutIdLst``). The target
        layout part already carries its own ``SLIDE_MASTER`` back-rel (wired
        by ``SlideMasterPart.add_layout``), so the slide→layout→master chain
        is intact and dangling-rel-free. Idempotent: applying the
        currently-related layout drops then recreates the same edge with the
        same effect (issue #19 SF7; ISC-44..49).
        """
        for rId, rel in list(self.rels.items()):
            if not rel.is_external and rel.reltype == RT.SLIDE_LAYOUT:
                self.drop_rel(rId)
        self.relate_to(slide_layout_part, RT.SLIDE_LAYOUT)

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
_P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
_P14_NS = "{http://schemas.microsoft.com/office/powerpoint/2010/main}"
_OOXML_LAYOUT_ID_FLOOR = 2147483648  # uint32 floor per ECMA-376 sec 19.2.1.27.
_UINT32_MAX = 4294967295  # ceiling shared by all slide-master/layout id allocators

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


# Reltypes filtered out during layout copy_from. SLIDE_MASTER is the
# layout→master back-rel — the destination layout already owns its own
# (wired by `SlideMasterPart.add_layout`), so copying it would create a
# duplicate, conflicting master relationship.
_LAYOUT_COPY_DROP_RELTYPES = frozenset({RT.SLIDE_MASTER})


def _replicate_rels_for_layout_copy(src_part: Part, new_part: Part) -> dict[str, str]:
    """Mirror `src_part`'s non-structural rels onto `new_part`.

    Used by `SlideLayoutPart.copy_shapes_from` (issue #19 SF4). Image,
    media, and external-target rels are reused/recreated so the copied
    shapes resolve; the `SLIDE_MASTER` back-rel is skipped because the
    destination layout already has its own. Returns a `{old_rId: new_rId}`
    map for rId-attribute remapping on the copied shape XML.
    """
    rId_map: dict[str, str] = {}
    for rId, rel in src_part.rels.items():
        if rel.reltype in _LAYOUT_COPY_DROP_RELTYPES:
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
            # Shared parts: image, media, video, theme, etc. — reuse the
            # same package-level part (SHA1 dedup already happened on add).
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


# ---------------------------------------------------------------------------
# Cross-package porting (Phase 3 of issue #11 — `Presentation.append_from`).
#
# Phase 2 (`Slide.duplicate`) cloned within a SINGLE package — all parts
# referenced by the source slide already lived in `target_package`, so the
# relate_to() calls naturally deduped at the rel level. Phase 3 ports parts
# BETWEEN two packages, so the dedup machinery has to come from us.
#
# Strategy: a single `_PortContext` per `append_from` call carries a set of
# {source_part: target_part} maps for slide-master, slide-layout, and theme.
# Image/media dedup at target uses the package's existing SHA1-based
# `get_or_add_image_part` / `_media_parts` machinery — we just hand it the
# foreign blob via BytesIO. The recursion master→layouts→master is broken
# by registering the new master in the map BEFORE walking its layouts.
# ---------------------------------------------------------------------------


# Reltypes that — when encountered on a master, layout, or theme — are
# private to that part and must be ported (deep-copy + new partname). The
# notes-slide / comments handling sticks with Phase 2's drop-list at the
# slide level; layout/master have neither.
_LAYOUT_OR_MASTER_PRIVATE_PARTNAMES = {
    # parts addressable by partname suffix; we only need this at the
    # `_port_blob_xml_part` fallback for parts whose class doesn't ship a
    # `partname_template` attribute.
}


class _PortContext:
    """Per-`append_from`-call cache + helpers for cross-package porting.

    Holds the `{source_part: target_part}` maps that ensure source
    slide-masters / layouts / themes shared across multiple appended source
    slides land as a single target part. Constructed once per call to
    `Presentation.append_from`; not state-of-the-Presentation.
    """

    def __init__(self, target_pres_part: PresentationPart):
        self.target_pres_part = target_pres_part
        self.target_package = target_pres_part.package
        self._master_map: dict[Part, Part] = {}
        self._layout_map: dict[Part, Part] = {}
        self._theme_map: dict[Part, Part] = {}
        # ---unified id pool: master ids and layout ids share one pool per
        #    OOXML spec, and PowerPoint enforces uniqueness across both.
        #    Both `_add_sldMasterId_to_presentation` and
        #    `_renumber_sldLayoutIds` allocate from this set and update it
        #    so sequential master ports don't collide with each other or
        #    with target's existing layout ids.
        self._used_layout_ids = _used_master_layout_ids_in_target(target_pres_part)

    # -- Public entry point ---------------------------------------------------

    def port_slide(self, src_slide_part: SlidePart) -> SlidePart:
        """Port a single source slide into the target package.

        Returns the new |SlidePart| in target. Caller is responsible for
        registering the new slide with `target.part` (rId + sldId).
        """
        target_layout_part = self._port_layout(
            cast(SlideLayoutPart, src_slide_part.part_related_by(RT.SLIDE_LAYOUT))
        )

        new_partname = self.target_package.next_partname("/ppt/slides/slide%d.xml")
        new_element = copy.deepcopy(src_slide_part._element)
        new_slide_part = SlidePart(new_partname, CT.PML_SLIDE, self.target_package, new_element)

        rId_map = self._replicate_slide_rels(src_slide_part, new_slide_part, target_layout_part)
        _remap_rId_attrs(new_element, rId_map)

        return new_slide_part

    def port_notes_slide(
        self, src_slide_part: SlidePart, new_slide_part: SlidePart
    ) -> NotesSlidePart:
        """Build a fresh notes-slide for `new_slide_part` from src's notes.

        Mirrors Phase 2's `duplicate_notes_slide_for` but routes through
        the target package and uses the target's existing notes-master.
        Cross-package gotcha #961: the notes-slide's `RT.SLIDE` back-rel
        is rewired to the new slide; the `RT.NOTES_MASTER` rel points at
        target's existing notes-master (NOT a port of source's).
        """
        src_notes_part = cast(NotesSlidePart, src_slide_part.part_related_by(RT.NOTES_SLIDE))
        new_partname = self.target_package.next_partname("/ppt/notesSlides/notesSlide%d.xml")
        new_element = copy.deepcopy(src_notes_part._element)
        new_notes_part = NotesSlidePart(
            new_partname, CT.PML_NOTES_SLIDE, self.target_package, new_element
        )

        target_notes_master_part = self.target_pres_part.notes_master_part

        rId_map: dict[str, str] = {}
        for rId, rel in src_notes_part.rels.items():
            if rel.is_external:
                new_rId = new_notes_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            elif rel.reltype == RT.SLIDE:
                new_rId = new_notes_part.relate_to(new_slide_part, RT.SLIDE)
            elif rel.reltype == RT.NOTES_MASTER:
                # ---target's existing notes-master, not ported from source---
                new_rId = new_notes_part.relate_to(target_notes_master_part, RT.NOTES_MASTER)
            elif rel.reltype == RT.IMAGE:
                new_target = self._port_image_part(cast(ImagePart, rel.target_part))
                new_rId = new_notes_part.relate_to(new_target, RT.IMAGE)
            else:
                # ---fall back to a deep-copy port for anything unexpected---
                new_target = self._port_unknown_part(cast(Part, rel.target_part))
                new_rId = new_notes_part.relate_to(new_target, rel.reltype)
            rId_map[rId] = new_rId
        _remap_rId_attrs(new_element, rId_map)

        new_slide_part.relate_to(new_notes_part, RT.NOTES_SLIDE)
        return new_notes_part

    # -- Per-reltype porting --------------------------------------------------

    def _replicate_slide_rels(
        self,
        src_slide_part: SlidePart,
        new_slide_part: SlidePart,
        target_layout_part: SlideLayoutPart,
    ) -> dict[str, str]:
        """Mirror src slide's rels onto new slide in target package."""
        rId_map: dict[str, str] = {}
        for rId, rel in src_slide_part.rels.items():
            if rel.reltype in _DUP_DROP_RELTYPES_SLIDE:
                continue
            if rel.is_external:
                new_rId = new_slide_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            elif rel.reltype == RT.SLIDE_LAYOUT:
                # ---use the already-ported (or shared) target layout---
                new_rId = new_slide_part.relate_to(target_layout_part, RT.SLIDE_LAYOUT)
            elif rel.reltype == RT.IMAGE:
                new_target = self._port_image_part(cast(ImagePart, rel.target_part))
                new_rId = new_slide_part.relate_to(new_target, RT.IMAGE)
            elif rel.reltype in (RT.MEDIA, RT.VIDEO, RT.AUDIO):
                new_target = self._port_image_part_like(cast(Part, rel.target_part))
                new_rId = new_slide_part.relate_to(new_target, rel.reltype)
            elif rel.reltype == RT.CHART:
                new_target = _duplicate_chart_part_into(
                    cast(ChartPart, rel.target_part), self.target_package
                )
                new_rId = new_slide_part.relate_to(new_target, rel.reltype)
            elif rel.reltype in (RT.OLE_OBJECT, RT.PACKAGE):
                new_target = _duplicate_blob_part_into(
                    cast(Part, rel.target_part), self.target_package
                )
                new_rId = new_slide_part.relate_to(new_target, rel.reltype)
            else:
                # ---unknown reltype: fall back to deep-copy port to be safe---
                new_target = self._port_unknown_part(cast(Part, rel.target_part))
                new_rId = new_slide_part.relate_to(new_target, rel.reltype)
            rId_map[rId] = new_rId
        return rId_map

    def _port_layout(self, src_layout_part: SlideLayoutPart) -> SlideLayoutPart:
        """Port a slide-layout into target package (with master + theme)."""
        if src_layout_part in self._layout_map:
            return cast(SlideLayoutPart, self._layout_map[src_layout_part])

        # ---porting the master will fill self._layout_map for ALL of master's
        #    layouts (master owns its layout tree); after that returns, our
        #    target layout is in the map---
        src_master_part = cast(SlideMasterPart, src_layout_part.part_related_by(RT.SLIDE_MASTER))
        self._port_master(src_master_part)

        # ---if porting the master populated the map, we're done---
        if src_layout_part in self._layout_map:
            return cast(SlideLayoutPart, self._layout_map[src_layout_part])

        # ---defensive fallback: master had no rel to this layout, port standalone---
        return self._port_layout_standalone(src_layout_part)

    def _port_layout_standalone(self, src_layout_part: SlideLayoutPart) -> SlideLayoutPart:
        """Port a layout whose master rel didn't auto-discover it."""
        src_master_part = cast(SlideMasterPart, src_layout_part.part_related_by(RT.SLIDE_MASTER))
        target_master_part = self._port_master(src_master_part)

        new_partname = self.target_package.next_partname("/ppt/slideLayouts/slideLayout%d.xml")
        new_element = copy.deepcopy(src_layout_part._element)
        _refresh_creation_ids(new_element)
        new_layout_part = SlideLayoutPart(
            new_partname, src_layout_part.content_type, self.target_package, new_element
        )
        self._layout_map[src_layout_part] = new_layout_part

        rId_map = self._replicate_layout_rels(src_layout_part, new_layout_part, target_master_part)
        _remap_rId_attrs(new_element, rId_map)

        new_master_layout_rId = target_master_part.relate_to(new_layout_part, RT.SLIDE_LAYOUT)
        # ---also append to target master's <p:sldLayoutIdLst>---
        _add_sldLayoutId_to_master(target_master_part, new_master_layout_rId)
        return new_layout_part

    def _port_master(self, src_master_part: SlideMasterPart) -> SlideMasterPart:
        """Port a slide-master into target package, including all its layouts.

        Order matters here: each new part must be related into the package's
        rel graph before the next call to `next_partname`, otherwise the
        unrelated part is invisible to `iter_parts` and `next_partname`
        hands out the same partname twice (the zipfile then gets duplicate
        entries and PowerPoint rejects the file).
        """
        if src_master_part in self._master_map:
            return cast(SlideMasterPart, self._master_map[src_master_part])

        # 1. Create new master part (deep-copy element).
        #    `<p14:creationId val>` is deterministic in default.pptx, so
        #    deepcopied creationIds collide with target's; refresh now.
        #    sldLayoutId renumbering is deferred to step 3 so layout ids
        #    end up ABOVE the new master's id — PowerPoint's own repair
        #    output matches this ordering and ECMA-376 sec 19.2.1.34 /
        #    19.2.1.27 confirm master and layout ids share one pool.
        new_master_partname = self.target_package.next_partname(
            "/ppt/slideMasters/slideMaster%d.xml"
        )
        new_element = copy.deepcopy(src_master_part._element)
        _refresh_creation_ids(new_element)
        new_master_part = SlideMasterPart(
            new_master_partname,
            src_master_part.content_type,
            self.target_package,
            new_element,
        )
        # 2. Register in map + relate to presentation FIRST so iter_parts
        #    reaches the master before any subsequent next_partname call.
        #    Allocate the new master's id from the unified master/layout
        #    pool — collision with an existing layout id flags repair.
        self._master_map[src_master_part] = new_master_part
        new_master_rId = self.target_pres_part.relate_to(new_master_part, RT.SLIDE_MASTER)
        _add_sldMasterId_to_presentation(
            self.target_pres_part, new_master_rId, self._used_layout_ids
        )
        # 3. Renumber the deepcopied sldLayoutIds NOW, after the master id
        #    has consumed its slot in the pool, so layouts land above it.
        _renumber_sldLayoutIds(new_element, self._used_layout_ids)

        # 3. For each source layout owned by this master: create + register
        #    + relate to the master IMMEDIATELY (before next next_partname).
        src_layouts = [
            cast(SlideLayoutPart, rel.target_part)
            for rel in src_master_part.rels.values()
            if not rel.is_external and rel.reltype == RT.SLIDE_LAYOUT
        ]
        new_layouts: list[tuple[SlideLayoutPart, SlideLayoutPart]] = []
        for src_layout in src_layouts:
            new_layout_partname = self.target_package.next_partname(
                "/ppt/slideLayouts/slideLayout%d.xml"
            )
            new_layout_element = copy.deepcopy(src_layout._element)
            _refresh_creation_ids(new_layout_element)
            new_layout = SlideLayoutPart(
                new_layout_partname,
                src_layout.content_type,
                self.target_package,
                new_layout_element,
            )
            self._layout_map[src_layout] = new_layout
            # Relate IMMEDIATELY so the next next_partname sees this layout.
            new_master_part.relate_to(new_layout, RT.SLIDE_LAYOUT)
            new_layouts.append((src_layout, new_layout))

        # 4. Walk master's rels. SLIDE_LAYOUT rels were created in step 3
        #    so they are reused (not re-created) by `_replicate_master_rels`.
        master_rId_map = self._replicate_master_rels(src_master_part, new_master_part)
        _remap_rId_attrs(new_element, master_rId_map)

        # 5. Walk each layout's rels (master is in the map).
        for src_layout, new_layout in new_layouts:
            layout_rId_map = self._replicate_layout_rels(src_layout, new_layout, new_master_part)
            _remap_rId_attrs(new_layout._element, layout_rId_map)

        return new_master_part

    def _replicate_master_rels(
        self, src_master_part: SlideMasterPart, new_master_part: SlideMasterPart
    ) -> dict[str, str]:
        """Mirror master's rels onto new master in target package."""
        rId_map: dict[str, str] = {}
        for rId, rel in src_master_part.rels.items():
            if rel.is_external:
                new_rId = new_master_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            elif rel.reltype == RT.SLIDE_LAYOUT:
                # ---layout was pre-created and is in self._layout_map---
                new_rId = new_master_part.relate_to(
                    self._layout_map[cast(Part, rel.target_part)], RT.SLIDE_LAYOUT
                )
            elif rel.reltype == RT.THEME:
                new_target = self._port_theme(cast(XmlPart, rel.target_part))
                new_rId = new_master_part.relate_to(new_target, RT.THEME)
            elif rel.reltype == RT.IMAGE:
                new_target = self._port_image_part(cast(ImagePart, rel.target_part))
                new_rId = new_master_part.relate_to(new_target, RT.IMAGE)
            else:
                new_target = self._port_unknown_part(cast(Part, rel.target_part))
                new_rId = new_master_part.relate_to(new_target, rel.reltype)
            rId_map[rId] = new_rId
        return rId_map

    def _replicate_layout_rels(
        self,
        src_layout_part: SlideLayoutPart,
        new_layout_part: SlideLayoutPart,
        target_master_part: SlideMasterPart,
    ) -> dict[str, str]:
        """Mirror layout's rels onto new layout in target package."""
        rId_map: dict[str, str] = {}
        for rId, rel in src_layout_part.rels.items():
            if rel.is_external:
                new_rId = new_layout_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            elif rel.reltype == RT.SLIDE_MASTER:
                new_rId = new_layout_part.relate_to(target_master_part, RT.SLIDE_MASTER)
            elif rel.reltype == RT.IMAGE:
                new_target = self._port_image_part(cast(ImagePart, rel.target_part))
                new_rId = new_layout_part.relate_to(new_target, RT.IMAGE)
            else:
                new_target = self._port_unknown_part(cast(Part, rel.target_part))
                new_rId = new_layout_part.relate_to(new_target, rel.reltype)
            rId_map[rId] = new_rId
        return rId_map

    def _port_theme(self, src_theme_part: Part) -> Part:
        """Port a theme part into target package (within-call dedup).

        Theme parts may load as `XmlPart` (created in-memory) or as the
        base `Part` (binary blob, when loaded from a `.pptx` on disk and
        the content-type isn't registered to a specific XML class). Handle
        both cases by branching on element-presence.
        """
        if src_theme_part in self._theme_map:
            return self._theme_map[src_theme_part]

        new_partname = self.target_package.next_partname("/ppt/theme/theme%d.xml")
        cls = type(src_theme_part)
        new_part: Part

        if isinstance(src_theme_part, XmlPart):
            new_element = copy.deepcopy(src_theme_part._element)
            new_part = cls(
                new_partname, src_theme_part.content_type, self.target_package, new_element
            )
            self._theme_map[src_theme_part] = new_part
            rId_map: dict[str, str] = {}
            for rId, rel in src_theme_part.rels.items():
                if rel.is_external:
                    new_rId = new_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                elif rel.reltype == RT.IMAGE:
                    new_target = self._port_image_part(cast(ImagePart, rel.target_part))
                    new_rId = new_part.relate_to(new_target, RT.IMAGE)
                else:
                    new_target = self._port_unknown_part(cast(Part, rel.target_part))
                    new_rId = new_part.relate_to(new_target, rel.reltype)
                rId_map[rId] = new_rId
            _remap_rId_attrs(new_element, rId_map)
        else:
            # Theme loaded as binary Part — blob copy. No rels typically.
            new_part = cls(
                new_partname, src_theme_part.content_type, self.target_package, src_theme_part.blob
            )
            self._theme_map[src_theme_part] = new_part
            for rId, rel in src_theme_part.rels.items():
                if rel.is_external:
                    new_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                elif rel.reltype == RT.IMAGE:
                    new_target = self._port_image_part(cast(ImagePart, rel.target_part))
                    new_part.relate_to(new_target, RT.IMAGE)
                else:
                    new_target = self._port_unknown_part(cast(Part, rel.target_part))
                    new_part.relate_to(new_target, rel.reltype)
        return new_part

    def _port_image_part(self, src_image_part: ImagePart) -> ImagePart:
        """Port an image into target package, deduping by SHA1.

        Routes through `Package.get_or_add_image_part` so an image already
        present in target (by SHA1 of bytes) is reused rather than duplicated.
        """
        return self.target_package.get_or_add_image_part(BytesIO(src_image_part.blob))

    def _port_image_part_like(self, src_part: Part) -> Part:
        """Port a media-grade binary part (audio/video). Naive blob copy.

        python-pptx's `_MediaParts` dedupes by SHA1 only when constructed
        via `add_media_part(video)`, which we don't have here. Doing a
        plain blob copy is safe; downstream merging is a follow-up.
        """
        return _duplicate_blob_part_into(src_part, self.target_package)

    def _port_unknown_part(self, src_part: Part) -> Part:
        """Last-resort port for unknown reltypes — deep-copy if XML, else blob.

        Practical safety net for OOXML extensions we don't enumerate
        explicitly (font tables, vmlDrawings, etc.). XML parts get a
        deepcopy element + rId remap; binary parts get a blob copy.
        """
        if isinstance(src_part, XmlPart):
            new_partname = self.target_package.next_partname(
                _derive_partname_template(str(src_part.partname))
            )
            new_element = copy.deepcopy(src_part._element)
            cls = type(src_part)
            new_part = cls(new_partname, src_part.content_type, self.target_package, new_element)
            rId_map: dict[str, str] = {}
            for rId, rel in src_part.rels.items():
                if rel.is_external:
                    new_rId = new_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                else:
                    new_target = self._port_unknown_part(cast(Part, rel.target_part))
                    new_rId = new_part.relate_to(new_target, rel.reltype)
                rId_map[rId] = new_rId
            _remap_rId_attrs(new_element, rId_map)
            return new_part
        return _duplicate_blob_part_into(src_part, self.target_package)


def _duplicate_chart_part_into(src: ChartPart, target_package) -> ChartPart:
    """Cross-package variant of `_duplicate_chart_part`."""
    new_partname = target_package.next_partname("/ppt/charts/chart%d.xml")
    new_element = copy.deepcopy(src._element)
    cls = type(src)
    new_part = cls(new_partname, src.content_type, target_package, new_element)
    rId_map: dict[str, str] = {}
    for rId, rel in src.rels.items():
        if rel.is_external:
            new_rId = new_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif rel.reltype == RT.PACKAGE:
            new_target = _duplicate_blob_part_into(cast(Part, rel.target_part), target_package)
            new_rId = new_part.relate_to(new_target, rel.reltype)
        else:
            new_rId = new_part.relate_to(rel.target_part, rel.reltype)
        rId_map[rId] = new_rId
    _remap_rId_attrs(new_element, rId_map)
    return new_part


def _duplicate_blob_part_into(src: Part, target_package) -> Part:
    """Cross-package variant of `_duplicate_blob_part`."""
    cls = type(src)
    tmpl = getattr(cls, "partname_template", None)
    if tmpl is None:
        tmpl = _derive_partname_template(str(src.partname))
    new_partname = target_package.next_partname(tmpl)
    return cls(new_partname, src.content_type, target_package, src.blob)


def _used_master_layout_ids_in_target(pres_part: PresentationPart) -> set[int]:
    """Return every `<p:sldMasterId>` AND `<p:sldLayoutId>` `id` in target.

    PowerPoint enforces uniqueness across the SHARED master/layout id pool
    (per ECMA-376 sec 19.2.1.34 sldMasterId / sec 19.2.1.27 sldLayoutId,
    both consume the same conceptual identifier space). A new master id
    that collides with an existing layout id flags the file for repair on
    load.
    """
    used: set[int] = set()
    sldMasterIdLst = pres_part._element.find(f"{_P_NS}sldMasterIdLst")
    if sldMasterIdLst is not None:
        for smi in sldMasterIdLst.findall(f"{_P_NS}sldMasterId"):
            raw = smi.get("id")
            if raw is None:
                continue
            with contextlib.suppress(ValueError):
                used.add(int(raw))
    for rel in pres_part.rels.values():
        if rel.is_external or rel.reltype != RT.SLIDE_MASTER:
            continue
        master_element = rel.target_part._element
        sldLayoutIdLst = master_element.find(f"{_P_NS}sldLayoutIdLst")
        if sldLayoutIdLst is None:
            continue
        for sli in sldLayoutIdLst.findall(f"{_P_NS}sldLayoutId"):
            raw = sli.get("id")
            if raw is None:
                continue
            with contextlib.suppress(ValueError):
                used.add(int(raw))
    return used


def _renumber_sldLayoutIds(element, used_ids: set[int]) -> None:
    """Reassign `<p:sldLayoutId>` `id` attributes on a deepcopied master.

    Mutates `used_ids` in place so back-to-back ports of multiple masters
    keep allocating unique ids relative to one another, not just relative
    to target's pre-existing masters.
    """
    sldLayoutIdLst = element.find(f"{_P_NS}sldLayoutIdLst")
    if sldLayoutIdLst is None:
        return
    next_id = max(used_ids | {_OOXML_LAYOUT_ID_FLOOR - 1}) + 1
    for sli in sldLayoutIdLst.findall(f"{_P_NS}sldLayoutId"):
        if next_id > _UINT32_MAX:
            next_id = next(
                (n for n in range(_OOXML_LAYOUT_ID_FLOOR, _UINT32_MAX + 1) if n not in used_ids),
                None,
            )
            if next_id is None:
                raise ValueError("slide-layout id pool exhausted")
        sli.set("id", str(next_id))
        used_ids.add(next_id)
        next_id += 1


def _refresh_creation_ids(element) -> None:
    """Reassign every `<p14:creationId val>` to a fresh uint32 value.

    The default `default.pptx` ships with deterministic vals on each
    master / layout `<p14:creationId>`, so two presentations built from
    it carry identical vals. PowerPoint flags those as bad content.
    Replacing the val (rather than stripping the element) keeps the
    extLst structure PowerPoint expects intact.
    """
    creation_id_tag = f"{_P14_NS}creationId"
    for node in element.iter(creation_id_tag):
        # 32-bit unsigned, cryptographically strong enough for collision avoidance.
        node.set("val", str(secrets.randbits(32)))


def _add_sldMasterId_to_presentation(
    pres_part: PresentationPart, rId: str, used_ids: set[int] | None = None
) -> None:
    """Append a `<p:sldMasterId>` to the presentation, allocating an id.

    Slide-master ids in OOXML are required uint32 values typically in the
    high range (default first master is `2147483648`). python-pptx's
    `CT_SlideMasterIdListEntry` only declares the `rId` attribute, so we
    set / read `id` directly on the lxml element to round-trip cleanly.

    `used_ids`, when supplied, is the unified master/layout id pool. The
    new id is taken above its current max and added to it; PowerPoint
    enforces master/layout id uniqueness across this shared pool, and a
    collision flags the file for repair on load. When `used_ids` is None,
    only existing sldMasterId values are scanned (back-compat for any
    out-of-band caller).
    """
    sldMasterIdLst = pres_part._element.get_or_add_sldMasterIdLst()
    if used_ids is None:
        used_ids = set()
        for sm in sldMasterIdLst.sldMasterId_lst:
            raw = sm.get("id")
            if raw is not None:
                with contextlib.suppress(ValueError):
                    used_ids.add(int(raw))
    next_id = max(used_ids | {_OOXML_LAYOUT_ID_FLOOR - 1}) + 1
    sldMasterId = sldMasterIdLst._add_sldMasterId()
    sldMasterId.rId = rId
    sldMasterId.set("id", str(next_id))
    used_ids.add(next_id)


def _add_sldLayoutId_to_master(master_part, rId: str) -> None:
    """Append a `<p:sldLayoutId>` to the master, allocating an id."""
    sldLayoutIdLst = master_part._element.get_or_add_sldLayoutIdLst()
    used_ids: list[int] = []
    for sl in sldLayoutIdLst.sldLayoutId_lst:
        raw = sl.get("id")
        if raw is not None:
            with contextlib.suppress(ValueError):
                used_ids.append(int(raw))
    # ---high-range allocation consistent with CT_SlideLayoutIdList._next_id:
    #    floor at _OOXML_LAYOUT_ID_FLOOR so the id is disjoint from the low
    #    p:sldId pool; ceiling-guarded at uint32 max with a scan fallback.---
    next_id = max(used_ids + [_OOXML_LAYOUT_ID_FLOOR - 1]) + 1
    if next_id > _UINT32_MAX:
        seen = set(used_ids)
        next_id = next(
            (n for n in range(_OOXML_LAYOUT_ID_FLOOR, _UINT32_MAX + 1) if n not in seen),
            None,
        )
        if next_id is None:
            raise ValueError("slide-layout id pool exhausted")
    sldLayoutId = sldLayoutIdLst._add_sldLayoutId()
    sldLayoutId.rId = rId
    sldLayoutId.set("id", str(next_id))


# SlideLayoutPart / SlideMasterPart are defined further down in this same
# module. Method bodies above resolve those names lazily via the module's
# globals at call time, so no forward declaration is needed.


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

    @classmethod
    def new(cls, partname: PackURI, package) -> SlideLayoutPart:
        """Return a newly-created blank |SlideLayoutPart| with `partname`.

        The part has a minimal valid `p:sldLayout` XML body and no
        relationships yet; the caller wires it to its slide-master.
        """
        return cls(partname, CT.PML_SLIDE_LAYOUT, package, CT_SlideLayout.new())

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

    def copy_shapes_from(self, source_layout_part: SlideLayoutPart) -> None:
        """Deep-copy `source_layout_part`'s shape tree into this layout.

        Every shape child of the source `p:spTree` (shapes, placeholders,
        pictures, group shapes) is deep-copied and appended to this
        layout's `p:spTree`. Non-structural relationships on the source
        layout part (images, media, external hyperlinks — everything
        except the `SLIDE_MASTER` back-relationship, which this layout
        already owns from `add_layout`) are replicated onto this part and
        the copied shape XML has its relationship-id attributes remapped so
        no dangling rels remain.

        The source layout part is NOT mutated (issue #19 SF4; ISC-23..29).
        """
        rId_map = _replicate_rels_for_layout_copy(source_layout_part, self)

        dest_spTree = self._element.spTree
        src_spTree = source_layout_part._element.spTree
        for shape_elm in src_spTree.iter_shape_elms():
            new_elm = copy.deepcopy(shape_elm)
            _remap_rId_attrs(new_elm, rId_map)
            dest_spTree.append(new_elm)


class SlideMasterPart(BaseSlidePart):
    """Slide master part.

    Corresponds to package files ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    """

    def add_layout(self) -> tuple[str, SlideLayout]:
        """Create a new blank slide-layout part bound to this master.

        Returns ``(rId, slide_layout)`` where `rId` is the master→layout
        relationship id (to be registered in `p:sldLayoutIdLst`) and
        `slide_layout` is the |SlideLayout| proxy for the new part.

        The package partname is allocated via `Package.next_partname`
        rather than upstream #1091's naive `len(layouts) + 1` scheme: this
        fork supports `SlideLayouts.remove`, so a layout count can lag the
        highest extant slideLayoutN.xml index and `len + 1` would collide
        with a surviving part. `next_partname` scans for the first free
        index and is collision-safe.
        """
        layout_part = SlideLayoutPart.new(
            self._package.next_partname("/ppt/slideLayouts/slideLayout%d.xml"),
            self._package,
        )
        rId = self.relate_to(layout_part, RT.SLIDE_LAYOUT)
        # ---back-relationship layout→master, required for inheritance---
        layout_part.relate_to(self, RT.SLIDE_MASTER)
        return rId, layout_part.slide_layout

    def related_slide_layout(self, rId: str) -> SlideLayout:
        """Return |SlideLayout| related to this slide-master by key `rId`."""
        return self.related_part(rId).slide_layout

    @lazyproperty
    def slide_master(self):
        """
        The |SlideMaster| object representing this part.
        """
        return SlideMaster(self._element, self)
