"""Custom element classes for presentation-related XML elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, cast

from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import ST_SlideId, ST_SlideSizeCoordinate, XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from pptx.util import Length


# -- URI assigned by Microsoft for the section-list extension
#    (PresentationML 2010 — see ECMA-376 / MS-OOXML Part 4, §13.7.5).
SECTION_LIST_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"


class CT_Presentation(BaseOxmlElement):
    """`p:presentation` element, root of the Presentation part stored as `/ppt/presentation.xml`."""

    get_or_add_sldSz: Callable[[], CT_SlideSize]
    get_or_add_sldIdLst: Callable[[], CT_SlideIdList]
    get_or_add_sldMasterIdLst: Callable[[], CT_SlideMasterIdList]
    get_or_add_extLst: Callable[[], CT_PresentationExtensionList]

    sldMasterIdLst: CT_SlideMasterIdList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldMasterIdLst",
        successors=(
            "p:notesMasterIdLst",
            "p:handoutMasterIdLst",
            "p:sldIdLst",
            "p:sldSz",
            "p:notesSz",
            "p:extLst",
        ),
    )
    sldIdLst: CT_SlideIdList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldIdLst", successors=("p:sldSz", "p:notesSz", "p:extLst")
    )
    sldSz: CT_SlideSize | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldSz", successors=("p:notesSz", "p:extLst")
    )
    extLst: CT_PresentationExtensionList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:extLst"
    )

    def get_or_add_section_list(self) -> CT_SectionList:
        """Return the `p14:sectionLst` element for this presentation, creating if needed.

        Walks `p:extLst`/`p:ext` looking for the section-list extension URI; adds the
        ext + nested `p14:sectionLst` if not present.
        """
        extLst = self.get_or_add_extLst()
        ext = extLst.get_or_add_ext_by_uri(SECTION_LIST_EXT_URI)
        return ext.get_or_add_sectionLst()

    @property
    def section_list(self) -> CT_SectionList | None:
        """Return the existing `p14:sectionLst` element, or None if absent."""
        if self.extLst is None:
            return None
        ext = self.extLst.ext_by_uri(SECTION_LIST_EXT_URI)
        if ext is None:
            return None
        return ext.sectionLst

    def remove_section_list(self) -> None:
        """Drop the section-list extension entirely.

        Removes the wrapping `p:ext` (and the `p:extLst` if it becomes empty).
        Idempotent — does nothing when no section list is present.
        """
        if self.extLst is None:
            return
        ext = self.extLst.ext_by_uri(SECTION_LIST_EXT_URI)
        if ext is None:
            return
        self.extLst.remove(ext)
        if len(self.extLst.findall(qn("p:ext"))) == 0:
            self.remove(self.extLst)


class CT_PresentationExtensionList(BaseOxmlElement):
    """`p:extLst` element, last child of `p:presentation`.

    Container for `p:ext` elements; we only know how to interpret the
    section-list extension, but other extensions (e.g. modification
    tracking) round-trip through this container untouched.
    """

    ext_lst: list[CT_PresentationExtension]

    ext = ZeroOrMore("p:ext")

    def ext_by_uri(self, uri: str) -> CT_PresentationExtension | None:
        """Return the `p:ext` child whose `uri` attribute matches `uri`, or None."""
        for ext in self.ext_lst:
            if ext.uri == uri:
                return ext
        return None

    def get_or_add_ext_by_uri(self, uri: str) -> CT_PresentationExtension:
        """Return existing or newly-created `p:ext` matching `uri`."""
        ext = self.ext_by_uri(uri)
        if ext is None:
            ext = self._add_ext(uri=uri)
        return ext


class CT_PresentationExtension(BaseOxmlElement):
    """`p:ext` element under `p:extLst`, identified by its `uri` attribute.

    The element body is namespace-extensible — any extension defined elsewhere
    (e.g. `p14:sectionLst`) appears as a child here.
    """

    get_or_add_sectionLst: Callable[[], CT_SectionList]

    uri: str = RequiredAttribute("uri", XsdString)  # pyright: ignore[reportAssignmentType]
    sectionLst: CT_SectionList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p14:sectionLst"
    )


class CT_SectionList(BaseOxmlElement):
    """`p14:sectionLst` element under `p:ext` carrying section definitions."""

    section_lst: list[CT_Section]

    _add_section: Callable[..., CT_Section]
    section = ZeroOrMore("p14:section")

    def add_section(self, name: str, section_id: str) -> CT_Section:
        """Append a `p14:section` element with `name` and `id` (a GUID-with-braces)."""
        return self._add_section(name=name, id=section_id)

    def insert_section_at(self, name: str, section_id: str, idx: int) -> CT_Section:
        """Insert a new `p14:section` at zero-based position `idx`.

        `idx` may equal `len(self.section_lst)` to append. Raises `IndexError`
        if `idx` is out of range.
        """
        if idx < 0 or idx > len(self.section_lst):
            raise IndexError("section index out of range")
        new_section = self.add_section(name, section_id)
        if idx < len(self.section_lst) - 1:
            target = self.section_lst[idx]
            target.addprevious(new_section)
        return new_section


class CT_Section(BaseOxmlElement):
    """`p14:section` element under `p14:sectionLst`.

    Carries a human-readable `name`, a stable GUID `id`, and a `p14:sldIdLst`
    listing slide ids (NOT relationship ids) belonging to this section.
    """

    get_or_add_sldIdLst: Callable[[], CT_SectionSlideIdList]

    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    id: str = RequiredAttribute("id", XsdString)  # pyright: ignore[reportAssignmentType]
    sldIdLst: CT_SectionSlideIdList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p14:sldIdLst"
    )


class CT_SectionSlideIdList(BaseOxmlElement):
    """`p14:sldIdLst` element under `p14:section`.

    Holds an ordered list of `p14:sldId` references identifying the slides
    that belong to the parent section. References are by **slide id**
    (the integer ``p:sldId/@id``), not by ``r:id``.
    """

    sldId_lst: list[CT_SectionSlideId]

    _add_sldId: Callable[..., CT_SectionSlideId]
    sldId = ZeroOrMore("p14:sldId")

    def add_sldId(self, slide_id: int) -> CT_SectionSlideId:
        """Append a `p14:sldId` referencing the slide whose `p:sldId/@id` equals `slide_id`."""
        return self._add_sldId(id=slide_id)

    def remove_sldId_for(self, slide_id: int) -> bool:
        """Remove the `p14:sldId` matching `slide_id`. Return True if removed, False if absent."""
        for sldId in self.sldId_lst:
            if sldId.id == slide_id:
                self.remove(sldId)
                return True
        return False


class CT_SectionSlideId(BaseOxmlElement):
    """`p14:sldId` element under `p14:sldIdLst` of a `p14:section`."""

    id: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "id", ST_SlideId
    )


class CT_SlideId(BaseOxmlElement):
    """`p:sldId` element.

    Direct child of `p:sldIdLst` that contains an `rId` reference to a slide in the presentation.
    """

    id: int = RequiredAttribute("id", ST_SlideId)  # pyright: ignore[reportAssignmentType]
    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_SlideIdList(BaseOxmlElement):
    """`p:sldIdLst` element.

    Direct child of <p:presentation> that contains a list of the slide parts in the presentation.
    """

    sldId_lst: list[CT_SlideId]

    _add_sldId: Callable[..., CT_SlideId]
    sldId = ZeroOrMore("p:sldId")

    def add_sldId(self, rId: str) -> CT_SlideId:
        """Create and return a reference to a new `p:sldId` child element.

        The new `p:sldId` element has its r:id attribute set to `rId`.
        """
        return self._add_sldId(id=self._next_id, rId=rId)

    def insert_sldId_at(self, rId: str, idx: int) -> CT_SlideId:
        """Insert a new `p:sldId` child element at position `idx`.

        The new `p:sldId` element has its `r:id` attribute set to `rId` and
        receives the next available `id` value. `idx` may equal the current
        length to append. Raises `IndexError` if `idx` is out of range.
        """
        if idx < 0 or idx > len(self.sldId_lst):
            raise IndexError("slide index out of range")
        new_sldId = self.add_sldId(rId)
        if idx < len(self.sldId_lst) - 1:
            target = self.sldId_lst[idx]
            target.addprevious(new_sldId)
        return new_sldId

    def move_sldId_to(self, sldId: CT_SlideId, new_idx: int) -> None:
        """Reposition `sldId` to zero-based position `new_idx` in this list.

        `sldId` must already be a child of this element. Raises `IndexError`
        if `new_idx` is out of range.
        """
        sldId_lst = self.sldId_lst
        if new_idx < 0 or new_idx >= len(sldId_lst):
            raise IndexError("slide index out of range")
        if sldId_lst[new_idx] is sldId:
            return
        # -- detach from current position --
        self.remove(sldId)
        # -- re-fetch list so index reflects post-removal state --
        sldId_lst = self.sldId_lst
        if new_idx >= len(sldId_lst):
            self.append(sldId)
        else:
            sldId_lst[new_idx].addprevious(sldId)

    def remove_sldId(self, sldId: CT_SlideId) -> None:
        """Remove `sldId` child element from this list.

        Raises `ValueError` if `sldId` is not a child of this element.
        """
        if sldId.getparent() is not self:
            raise ValueError("sldId is not a child of this sldIdLst")
        self.remove(sldId)

    @property
    def _next_id(self) -> int:
        """The next available slide ID as an `int`.

        Valid slide IDs start at 256. The next integer value greater than the max value in use is
        chosen, which minimizes that chance of reusing the id of a deleted slide.
        """
        MIN_SLIDE_ID = 256
        MAX_SLIDE_ID = 2147483647

        used_ids = [int(s) for s in cast("list[str]", self.xpath("./p:sldId/@id"))]
        simple_next = max([MIN_SLIDE_ID - 1] + used_ids) + 1
        if simple_next <= MAX_SLIDE_ID:
            return simple_next

        # -- fall back to search for next unused from bottom --
        valid_used_ids = sorted(id for id in used_ids if (MIN_SLIDE_ID <= id <= MAX_SLIDE_ID))
        return (
            next(
                candidate_id
                for candidate_id, used_id in enumerate(valid_used_ids, start=MIN_SLIDE_ID)
                if candidate_id != used_id
            )
            if valid_used_ids
            else 256
        )


class CT_SlideMasterIdList(BaseOxmlElement):
    """`p:sldMasterIdLst` element.

    Child of `p:presentation` containing references to the slide masters that belong to the
    presentation.
    """

    sldMasterId_lst: list[CT_SlideMasterIdListEntry]

    sldMasterId = ZeroOrMore("p:sldMasterId")


class CT_SlideMasterIdListEntry(BaseOxmlElement):
    """
    ``<p:sldMasterId>`` element, child of ``<p:sldMasterIdLst>`` containing
    a reference to a slide master.
    """

    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_SlideSize(BaseOxmlElement):
    """`p:sldSz` element.

    Direct child of <p:presentation> that contains the width and height of slides in the
    presentation.
    """

    cx: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "cx", ST_SlideSizeCoordinate
    )
    cy: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "cy", ST_SlideSizeCoordinate
    )
