"""Section objects — read/write `<p14:sectionLst>` (PowerPoint Sections).

Sections group slides for organizational purposes in PowerPoint's
slide-organizer pane. They are stored in the presentation extension list
under `p:extLst/p:ext{uri="{521415D9-…}"}/p14:sectionLst` and reference
slides by their stable `p:sldId/@id` integer — meaning section
membership survives reorder, indexed insert, and remove operations on
the slide collection.

Public surface:

- |Presentation.sections|  → |_Sections|
- |_Sections| supports ``len()``, iteration, indexed access, ``index``,
  ``add_section(name, after=None)``, and ``remove(section)``.
- |Section| exposes ``name`` (read/write), ``id`` (read-only GUID),
  ``slides`` (tuple of |Slide|), ``add_slide(slide)``, and
  ``remove_slide(slide)``.

Issue: https://github.com/MHoroszowski/python-pptx/issues/11 (Phase 4).
"""

from __future__ import annotations

import uuid
from typing import TYPE_CHECKING, Iterator

if TYPE_CHECKING:
    from pptx.oxml.presentation import CT_Section
    from pptx.presentation import Presentation
    from pptx.slide import Slide


def _new_section_id() -> str:
    """Generate a fresh section GUID in the canonical PowerPoint shape.

    PowerPoint stores section ids as GUIDs surrounded by braces and
    upper-cased, e.g. ``{4080CFE4-F95C-449C-8898-95C81DD3D8B4}``.
    """
    return "{%s}" % str(uuid.uuid4()).upper()


class Section:
    """A single PowerPoint section — a named group of slides."""

    def __init__(self, section_elm: CT_Section, sections: _Sections):
        self._element = section_elm
        self._sections = sections

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, Section):
            return NotImplemented
        return self._element.id == other._element.id

    def __hash__(self) -> int:
        return hash(self._element.id)

    def __repr__(self) -> str:
        return "<Section name=%r id=%r>" % (self._element.name, self._element.id)

    @property
    def name(self) -> str:
        """Display name of this section."""
        return self._element.name

    @name.setter
    def name(self, value: str) -> None:
        self._element.name = value

    @property
    def id(self) -> str:
        """Stable GUID identifier of this section (e.g. ``{ABC...}``).

        Read-only; assigned at creation and persists across save / reopen.
        """
        return self._element.id

    @property
    def slides(self) -> tuple[Slide, ...]:
        """Tuple of |Slide| objects belonging to this section, in section order.

        Slides whose ids are listed in this section's ``p14:sldIdLst`` but
        no longer present in the presentation (orphaned references — should
        not occur in a well-formed deck, but possible in corrupt files) are
        silently skipped.
        """
        if self._element.sldIdLst is None:
            return ()
        prs_slides = self._sections._prs.slides
        slide_by_id = {s.slide_id: s for s in prs_slides}
        result: list[Slide] = []
        for sldId in self._element.sldIdLst.sldId_lst:
            slide = slide_by_id.get(sldId.id)
            if slide is not None:
                result.append(slide)
        return tuple(result)

    def add_slide(self, slide: Slide) -> None:
        """Move `slide` into this section.

        If `slide` is currently assigned to another section, it is removed
        from that section first — a slide can belong to at most one
        section. Appending a slide already in this section is a no-op.
        Raises |ValueError| if `slide` is not part of the presentation that
        owns this section.
        """
        prs = self._sections._prs
        if slide.part.package is not prs.part.package:
            raise ValueError("slide is not in this presentation")
        # ---ensure slide is not in any other section---
        self._sections._unassign_slide(slide.slide_id, except_section=self)
        # ---add to this section if not already present---
        sldIdLst = self._element.get_or_add_sldIdLst()
        for sldId in sldIdLst.sldId_lst:
            if sldId.id == slide.slide_id:
                return  # ---already a member, no-op---
        sldIdLst.add_sldId(slide.slide_id)

    def remove_slide(self, slide: Slide) -> None:
        """Remove `slide` from this section.

        The slide remains in the presentation. Raises |ValueError| if
        `slide` is not currently a member of this section.
        """
        if self._element.sldIdLst is None:
            raise ValueError("slide is not in this section")
        if not self._element.sldIdLst.remove_sldId_for(slide.slide_id):
            raise ValueError("slide is not in this section")


class _Sections:
    """Sequence-like collection of |Section| objects on a |Presentation|.

    An empty presentation (one with no `<p14:sectionLst>`) reports
    ``len() == 0`` without forcing the extension into existence; the
    ``p:extLst`` / ``p:ext`` / ``p14:sectionLst`` chain is only created
    on the first ``add_section`` call.
    """

    def __init__(self, prs: Presentation):
        self._prs = prs

    def __len__(self) -> int:
        sectionLst = self._prs._element.section_list
        if sectionLst is None:
            return 0
        return len(sectionLst.section_lst)

    def __iter__(self) -> Iterator[Section]:
        sectionLst = self._prs._element.section_list
        if sectionLst is None:
            return iter(())
        return (Section(s, self) for s in sectionLst.section_lst)

    def __getitem__(self, idx: int) -> Section:
        sectionLst = self._prs._element.section_list
        if sectionLst is None:
            raise IndexError("section index out of range")
        try:
            section_elm = sectionLst.section_lst[idx]
        except IndexError:
            raise IndexError("section index out of range")
        return Section(section_elm, self)

    def index(self, section: Section) -> int:
        """Return the zero-based position of `section` in this collection.

        Raises |ValueError| if `section` is not in the collection.
        """
        for idx, this_section in enumerate(self):
            if this_section == section:
                return idx
        raise ValueError("%s is not in this Sections collection" % section)

    def add_section(self, name: str, after: Section | None = None) -> Section:
        """Add a new section named `name`.

        When `after` is |None| (default), the new section is appended at
        the end. When `after` is an existing |Section|, the new section is
        inserted immediately after it. Raises |ValueError| if `after` is
        not a member of this collection.

        The new section is created with a fresh GUID id and starts with
        an empty slide list. Assign slides via |Section.add_slide|. The
        empty `<p14:sldIdLst/>` is materialized at creation time so the
        section is emitted in the same shape PowerPoint produces (some
        PowerPoint versions interpret a section that omits `sldIdLst`
        differently from one that includes an empty `sldIdLst`).
        """
        sectionLst = self._prs._element.get_or_add_section_list()
        new_id = _new_section_id()
        if after is None:
            section_elm = sectionLst.add_section(name=name, section_id=new_id)
        else:
            after_idx = self.index(after)
            section_elm = sectionLst.insert_section_at(
                name=name, section_id=new_id, idx=after_idx + 1
            )
        # ---force-create the (empty) sldIdLst to match PowerPoint's wire shape
        section_elm.get_or_add_sldIdLst()
        return Section(section_elm, self)

    def remove(self, section: Section) -> None:
        """Remove `section` from this collection.

        Slides previously belonging to `section` remain in the presentation
        but are no longer assigned to any section. Raises |ValueError| if
        `section` is not a member of this collection. When the last
        section is removed, the wrapping `p14:sectionLst` and its
        containing `p:ext` / `p:extLst` chain are cleaned up.
        """
        sectionLst = self._prs._element.section_list
        if sectionLst is None:
            raise ValueError("%s is not in this Sections collection" % section)
        for section_elm in sectionLst.section_lst:
            if section_elm.id == section.id:
                sectionLst.remove(section_elm)
                if len(sectionLst.section_lst) == 0:
                    self._prs._element.remove_section_list()
                return
        raise ValueError("%s is not in this Sections collection" % section)

    def _unassign_slide(self, slide_id: int, except_section: Section | None = None) -> None:
        """Remove `slide_id` from every section's sldIdLst (with optional skip).

        Internal helper called by |Section.add_slide| to enforce the
        "a slide belongs to at most one section" invariant.
        """
        sectionLst = self._prs._element.section_list
        if sectionLst is None:
            return
        for section_elm in sectionLst.section_lst:
            if except_section is not None and section_elm.id == except_section.id:
                continue
            if section_elm.sldIdLst is not None:
                section_elm.sldIdLst.remove_sldId_for(slide_id)
