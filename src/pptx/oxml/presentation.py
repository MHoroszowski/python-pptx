"""Custom element classes for presentation-related XML elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, cast

from pptx.oxml.simpletypes import ST_SlideId, ST_SlideSizeCoordinate, XsdString
from pptx.oxml.xmlchemy import BaseOxmlElement, RequiredAttribute, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from pptx.util import Length


class CT_Presentation(BaseOxmlElement):
    """`p:presentation` element, root of the Presentation part stored as `/ppt/presentation.xml`."""

    get_or_add_sldSz: Callable[[], CT_SlideSize]
    get_or_add_sldIdLst: Callable[[], CT_SlideIdList]
    get_or_add_sldMasterIdLst: Callable[[], CT_SlideMasterIdList]

    sldMasterIdLst: CT_SlideMasterIdList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldMasterIdLst",
        successors=(
            "p:notesMasterIdLst",
            "p:handoutMasterIdLst",
            "p:sldIdLst",
            "p:sldSz",
            "p:notesSz",
        ),
    )
    sldIdLst: CT_SlideIdList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldIdLst", successors=("p:sldSz", "p:notesSz")
    )
    sldSz: CT_SlideSize | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldSz", successors=("p:notesSz",)
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
