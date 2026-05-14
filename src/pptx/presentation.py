"""Main presentation object."""

from __future__ import annotations

import os
from typing import IO, TYPE_CHECKING, Iterable, cast

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.shared import PartElementProxy
from pptx.slide import SlideMasters, Slides
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.custom_properties import CustomProperties
    from pptx.custom_xml import CustomXmlParts
    from pptx.oxml.presentation import CT_Presentation, CT_SlideId
    from pptx.parts.presentation import PresentationPart
    from pptx.slide import HandoutMaster, NotesMaster, Slide, SlideLayouts
    from pptx.util import Length


class Presentation(PartElementProxy):
    """PresentationML (PML) presentation.

    Not intended to be constructed directly. Use :func:`pptx.Presentation` to open or
    create a presentation.
    """

    _element: CT_Presentation
    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    @property
    def core_properties(self):
        """|CoreProperties| instance for this presentation.

        Provides read/write access to the Dublin Core document properties for the presentation.
        """
        return self.part.core_properties

    @property
    def custom_properties(self) -> CustomProperties:
        """Mapping-protocol view over the Custom Document Properties part.

        These are the user-defined properties surfaced under
        `File → Properties → Advanced` in PowerPoint. Created on first access
        if the package does not already have a custom properties part.
        """
        return self.part.custom_properties

    @property
    def custom_xml_parts(self) -> CustomXmlParts:
        """Collection of customXml data parts in this presentation's package.

        Walks both presentation-scoped and package-scoped `RT.CUSTOM_XML`
        relationships. Use `.add(...)` to attach a new part, `[i]` or
        `["item3.xml"]` to look one up by index or partname tail, and
        `.by_guid(...)` / `.by_name(...)` for the other lookup forms.
        """
        return self.part.custom_xml_parts

    @property
    def notes_master(self) -> NotesMaster:
        """Instance of |NotesMaster| for this presentation.

        If the presentation does not have a notes master, one is created from a default template
        and returned. The same single instance is returned on each call.
        """
        return self.part.notes_master

    @property
    def handout_master(self) -> HandoutMaster:
        """Instance of |HandoutMaster| for this presentation.

        Raises |ValueError| when the presentation has no handout master because auto-create is
        deliberately deferred until a built-in handout-master template ships in this fork.
        """
        return self.part.handout_master

    def save(self, file: str | os.PathLike[str] | IO[bytes]):
        """Writes this presentation to `file`.

        `file` can be a file-path (|str| or any |os.PathLike| object such as
        |pathlib.Path|) or a file-like object open for writing bytes.
        """
        # ---accept os.PathLike (pathlib.Path etc.) by coercing to str at the
        # ---boundary; collapse the union to (str | IO[bytes]) for downstream---
        pkg_file: str | IO[bytes] = os.fspath(file) if isinstance(file, os.PathLike) else file
        self.part.save(pkg_file)

    @property
    def slide_height(self) -> Length | None:
        """Height of slides in this presentation, in English Metric Units (EMU).

        Returns |None| if no slide width is defined. Read/write.
        """
        sldSz = self._element.sldSz
        if sldSz is None:
            return None
        return sldSz.cy

    @slide_height.setter
    def slide_height(self, height: Length):
        sldSz = self._element.get_or_add_sldSz()
        sldSz.cy = height

    @property
    def slide_layouts(self) -> SlideLayouts:
        """|SlideLayouts| collection belonging to the first |SlideMaster| of this presentation.

        A presentation can have more than one slide master and each master will have its own set
        of layouts. This property is a convenience for the common case where the presentation has
        only a single slide master.
        """
        return self.slide_masters[0].slide_layouts

    @property
    def slide_master(self):
        """
        First |SlideMaster| object belonging to this presentation. Typically,
        presentations have only a single slide master. This property provides
        simpler access in that common case.
        """
        return self.slide_masters[0]

    @lazyproperty
    def slide_masters(self) -> SlideMasters:
        """|SlideMasters| collection of slide-masters belonging to this presentation."""
        return SlideMasters(self._element.get_or_add_sldMasterIdLst(), self)

    @property
    def slide_width(self):
        """
        Width of slides in this presentation, in English Metric Units (EMU).
        Returns |None| if no slide width is defined. Read/write.
        """
        sldSz = self._element.sldSz
        if sldSz is None:
            return None
        return sldSz.cx

    @slide_width.setter
    def slide_width(self, width: Length):
        sldSz = self._element.get_or_add_sldSz()
        sldSz.cx = width

    @lazyproperty
    def slides(self):
        """|Slides| object containing the slides in this presentation."""
        sldIdLst = self._element.get_or_add_sldIdLst()
        self.part.rename_slide_parts([cast("CT_SlideId", sldId).rId for sldId in sldIdLst])
        return Slides(sldIdLst, self)

    @lazyproperty
    def sections(self):
        """|_Sections| collection of |Section| objects in this presentation.

        The collection reads from the `p14:sectionLst` extension under
        ``p:presentation/p:extLst`` and supports ``len()``, iteration,
        indexed access, ``index``, ``add_section(name, after=None)``, and
        ``remove(section)``. Section membership references slides by the
        stable ``p:sldId/@id`` integer, so reordering or indexed insert
        on the slide collection does not perturb section assignment.

        For a presentation that does not yet declare any sections, the
        collection reports ``len() == 0`` without forcing the extension
        elements into existence; the wrapping XML is created on the first
        ``add_section`` call.
        """
        from pptx.sections import _Sections  # pyright: ignore[reportPrivateUsage]

        return _Sections(self)

    def append_from(
        self,
        other_pres: Presentation,
        slide_indexes: Iterable[int] | None = None,
    ) -> list[Slide]:
        """Append slides from `other_pres` onto the end of this presentation.

        When `slide_indexes` is |None| (the default), every slide of
        `other_pres` is appended in source order. When `slide_indexes` is
        an iterable of zero-based ints, only those slides are appended
        (in the iteration order, allowing reordering via index sequence).

        Each appended slide brings along — into this presentation's
        package — its own deep-copied ``<p:sld>`` element, its
        slide-layout, the slide-master that layout belongs to (with
        ALL of that master's layouts, to keep the master's layout tree
        intact), the master's theme, and any per-slide private parts
        (chart, OLE-object, notes-slide). Image and embedded-media
        parts dedupe at the target package by SHA1 of bytes — an image
        already present in this presentation's package is reused, not
        copied.

        Within a single ``append_from`` call, source slide-masters
        (and their layouts and themes) are ported exactly once even
        when shared by multiple appended slides. Two consecutive
        ``append_from`` calls do NOT share that cache, so calling
        twice in succession ports the master twice — pass all wanted
        slides in a single call to keep the part graph compact.

        Comments parts on the source slides are dropped (consistent
        with Phase 2 of issue #11). Notes-slides on the source bring
        their own deep-copied notes-slide part; the notes-master is
        this presentation's existing notes-master, NOT a port of
        the source's.

        Returns the list of newly-added |Slide| objects in insertion
        order — same length as the number of slides actually appended.

        Raises |IndexError| if any value in `slide_indexes` is out of
        range for ``other_pres.slides``.
        """
        from pptx.parts.slide import _PortContext  # pyright: ignore[reportPrivateUsage]

        if slide_indexes is None:
            source_slides = list(other_pres.slides)
        else:
            indexes = list(slide_indexes)
            n_source = len(other_pres.slides)
            for idx in indexes:
                if idx < 0 or idx >= n_source:
                    raise IndexError("slide index out of range")
            source_slides = [other_pres.slides[i] for i in indexes]

        if not source_slides:
            return []

        ctx = _PortContext(self.part)
        sldIdLst = self._element.get_or_add_sldIdLst()
        new_slides: list[Slide] = []

        for src_slide in source_slides:
            new_slide_part = ctx.port_slide(src_slide.part)
            new_rId = self.part.relate_to(new_slide_part, RT.SLIDE)
            sldIdLst.add_sldId(new_rId)
            if src_slide.part.has_notes_slide:
                ctx.port_notes_slide(src_slide.part, new_slide_part)
            new_slides.append(new_slide_part.slide)

        return new_slides
