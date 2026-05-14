"""Slide-related objects, including masters, layouts, and notes."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, cast

from pptx.dml.color import RGBColor
from pptx.dml.fill import FillFormat
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.shapes.shapetree import (
    LayoutPlaceholders,
    LayoutShapes,
    MasterPlaceholders,
    MasterShapes,
    NotesSlidePlaceholders,
    NotesSlideShapes,
    SlidePlaceholders,
    SlideShapes,
)
from pptx.shared import ElementProxy, ParentedElementProxy, PartElementProxy
from pptx.util import Inches, Length, Pt, lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.presentation import CT_SlideIdList, CT_SlideMasterIdList
    from pptx.oxml.slide import (
        CT_CommonSlideData,
        CT_HandoutMaster,
        CT_NotesMaster,
        CT_NotesSlide,
        CT_Slide,
        CT_SlideLayout,
        CT_SlideLayoutIdList,
        CT_SlideMaster,
    )
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.slide import SlideLayoutPart, SlideMasterPart, SlidePart
    from pptx.presentation import Presentation
    from pptx.shapes.autoshape import Shape
    from pptx.shapes.placeholder import LayoutPlaceholder, MasterPlaceholder, SlidePlaceholder
    from pptx.shapes.shapetree import NotesSlidePlaceholder
    from pptx.text.text import TextFrame


class _BaseSlide(PartElementProxy):
    """Base class for slide objects, including masters, layouts and notes."""

    _element: CT_Slide

    @lazyproperty
    def background(self) -> _Background:
        """|_Background| object providing slide background properties.

        This property returns a |_Background| object whether or not the
        slide, master, or layout has an explicitly defined background.

        The same |_Background| object is returned on every call for the same
        slide object.
        """
        return _Background(self._element.cSld)

    @property
    def name(self) -> str:
        """String representing the internal name of this slide.

        Returns an empty string (`''`) if no name is assigned. Assigning an empty string or |None|
        to this property causes any name to be removed.
        """
        return self._element.cSld.name

    @name.setter
    def name(self, value: str | None):
        new_value = "" if value is None else value
        self._element.cSld.name = new_value


class _BaseMaster(_BaseSlide):
    """Base class for master objects such as |SlideMaster| and |NotesMaster|.

    Provides access to placeholders and regular shapes.
    """

    @lazyproperty
    def placeholders(self) -> MasterPlaceholders:
        """|MasterPlaceholders| collection of placeholder shapes in this master.

        Sequence sorted in `idx` order.
        """
        return MasterPlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self):
        """
        Instance of |MasterShapes| containing sequence of shape objects
        appearing on this slide.
        """
        return MasterShapes(self._element.spTree, self)


class _HeaderFooterVisibility:
    """Provides access to header/footer visibility settings on a slide template."""

    _element: CT_SlideLayout | CT_SlideMaster | CT_NotesMaster | CT_HandoutMaster

    def _get_hf_visibility(self, attr_name: str) -> bool:
        """Return effective `attr_name` value, defaulting to |True| when `<p:hf>` is absent."""
        hf = self._element.hf
        return True if hf is None else getattr(hf, attr_name)

    def _set_hf_visibility(self, attr_name: str, value: bool) -> None:
        """Set `attr_name` on `<p:hf>`, creating the element only when needed.

        Assigning |True| when `<p:hf>` is absent is a no-op because the effective default is
        already |True|. An existing `<p:hf>` element is retained even when all values become
        |True|, avoiding low-value XML churn.
        """
        hf = self._element.hf
        if hf is None and value:
            return
        if hf is None:
            hf = self._element.get_or_add_hf()
        setattr(hf, attr_name, value)

    @property
    def show_slide_number(self) -> bool:
        """`True` when slide numbers are shown for this template, `False` otherwise.

        Assigning |False| creates a `<p:hf>` element when needed and writes `sldNum="0"`.
        Assigning |True| preserves any existing `<p:hf>` element rather than removing it.
        """
        return self._get_hf_visibility("sldNum")

    @show_slide_number.setter
    def show_slide_number(self, value: bool) -> None:
        self._set_hf_visibility("sldNum", value)

    @property
    def show_footer(self) -> bool:
        """`True` when footer placeholders are shown for this template, `False` otherwise.

        Assigning |False| creates a `<p:hf>` element when needed and writes `ftr="0"`.
        Assigning |True| preserves any existing `<p:hf>` element rather than removing it.
        """
        return self._get_hf_visibility("ftr")

    @show_footer.setter
    def show_footer(self, value: bool) -> None:
        self._set_hf_visibility("ftr", value)

    @property
    def show_date(self) -> bool:
        """`True` when date placeholders are shown for this template, `False` otherwise.

        Assigning |False| creates a `<p:hf>` element when needed and writes `dt="0"`.
        Assigning |True| preserves any existing `<p:hf>` element rather than removing it.
        """
        return self._get_hf_visibility("dt")

    @show_date.setter
    def show_date(self, value: bool) -> None:
        self._set_hf_visibility("dt", value)

    @property
    def show_header(self) -> bool:
        """`True` when header placeholders are shown for this template, `False` otherwise.

        Assigning |False| creates a `<p:hf>` element when needed and writes `hdr="0"`.
        Assigning |True| preserves any existing `<p:hf>` element rather than removing it.
        """
        return self._get_hf_visibility("hdr")

    @show_header.setter
    def show_header(self, value: bool) -> None:
        self._set_hf_visibility("hdr", value)


class NotesMaster(_HeaderFooterVisibility, _BaseMaster):
    """Proxy for the notes master XML document.

    Provides access to shapes, the most commonly used of which are placeholders.
    """


class HandoutMaster(_HeaderFooterVisibility, _BaseMaster):
    """Proxy for the handout master XML document.

    Provides access to shapes and header/footer visibility settings when a deck already
    contains a handout master. Auto-create is deliberately deferred until a built-in
    `handoutMaster.xml` template ships in this fork.
    """


class NotesSlide(_BaseSlide):
    """Notes slide object.

    Provides access to slide notes placeholder and other shapes on the notes handout
    page.
    """

    element: CT_NotesSlide  # pyright: ignore[reportIncompatibleMethodOverride]

    def clone_master_placeholders(self, notes_master: NotesMaster) -> None:
        """Selectively add placeholder shape elements from `notes_master`.

        Selected placeholder shape elements from `notes_master` are added to the shapes
        collection of this notes slide. Z-order of placeholders is preserved. Certain
        placeholders (header, date, footer) are not cloned.
        """

        def iter_cloneable_placeholders() -> Iterator[MasterPlaceholder]:
            """Generate a reference to each cloneable placeholder in `notes_master`.

            These are the placeholders that should be cloned to a notes slide when the a new notes
            slide is created.
            """
            cloneable = (
                PP_PLACEHOLDER.SLIDE_IMAGE,
                PP_PLACEHOLDER.BODY,
                PP_PLACEHOLDER.SLIDE_NUMBER,
            )
            for placeholder in notes_master.placeholders:
                if placeholder.element.ph_type in cloneable:
                    yield placeholder

        shapes = self.shapes
        for placeholder in iter_cloneable_placeholders():
            shapes.clone_placeholder(cast("LayoutPlaceholder", placeholder))

    @property
    def notes_placeholder(self) -> NotesSlidePlaceholder | None:
        """the notes placeholder on this notes slide, the shape that contains the actual notes text.

        Return |None| if no notes placeholder is present; while this is probably uncommon, it can
        happen if the notes master does not have a body placeholder, or if the notes placeholder
        has been deleted from the notes slide.
        """
        for placeholder in self.placeholders:
            if placeholder.placeholder_format.type == PP_PLACEHOLDER.BODY:
                return placeholder
        return None

    @property
    def notes_text_frame(self) -> TextFrame | None:
        """The text frame of the notes placeholder on this notes slide.

        |None| if there is no notes placeholder. This is a shortcut to accommodate the common case
        of simply adding "notes" text to the notes "page".
        """
        notes_placeholder = self.notes_placeholder
        if notes_placeholder is None:
            return None
        return notes_placeholder.text_frame

    @lazyproperty
    def placeholders(self) -> NotesSlidePlaceholders:
        """Instance of |NotesSlidePlaceholders| for this notes-slide.

        Contains the sequence of placeholder shapes in this notes slide.
        """
        return NotesSlidePlaceholders(self.element.spTree, self)

    @lazyproperty
    def shapes(self) -> NotesSlideShapes:
        """Sequence of shape objects appearing on this notes slide."""
        return NotesSlideShapes(self._element.spTree, self)


class Slide(_BaseSlide):
    """Slide object. Provides access to shapes and slide-level properties."""

    part: SlidePart  # pyright: ignore[reportIncompatibleMethodOverride]

    @property
    def is_hidden(self) -> bool:
        """`True` if this slide is hidden from the presentation, `False` otherwise.

        Assigning `True` causes the slide to be hidden. Assigning `False` makes the slide visible.
        """
        return not self._element.show

    @is_hidden.setter
    def is_hidden(self, value: bool) -> None:
        self._element.show = not value

    @property
    def follow_master_background(self):
        """|True| if this slide inherits the slide master background.

        Assigning |False| causes background inheritance from the master to be
        interrupted; if there is no custom background for this slide,
        a default background is added. If a custom background already exists
        for this slide, assigning |False| has no effect.

        Assigning |True| causes any custom background for this slide to be
        deleted and inheritance from the master restored.
        """
        return self._element.bg is None

    @property
    def has_notes_slide(self) -> bool:
        """`True` if this slide has a notes slide, `False` otherwise.

        A notes slide is created by :attr:`.notes_slide` when one doesn't exist; use this property
        to test for a notes slide without the possible side effect of creating one.
        """
        return self.part.has_notes_slide

    @property
    def has_date(self) -> bool:
        """`True` if this slide has a date placeholder, `False` otherwise.

        This property is non-mutating; it reports only whether a DATE placeholder is already
        present on the slide.
        """
        return self._first_ph_of_type(PP_PLACEHOLDER.DATE) is not None

    @property
    def has_footer(self) -> bool:
        """`True` if this slide has a footer placeholder, `False` otherwise.

        This property is non-mutating; it reports only whether a FOOTER placeholder is already
        present on the slide.
        """
        return self._first_ph_of_type(PP_PLACEHOLDER.FOOTER) is not None

    @property
    def has_slide_number(self) -> bool:
        """`True` if this slide has a slide-number placeholder, `False` otherwise.

        This property is non-mutating; it reports only whether a SLIDE_NUMBER placeholder is
        already present on the slide.
        """
        return self._first_ph_of_type(PP_PLACEHOLDER.SLIDE_NUMBER) is not None

    @property
    def notes_slide(self) -> NotesSlide:
        """The |NotesSlide| instance for this slide.

        If the slide does not have a notes slide, one is created. The same single instance is
        returned on each call.
        """
        return self.part.notes_slide

    @property
    def date_text(self) -> str | None:
        """Text of this slide's date placeholder, or |None| when no date placeholder is present.

        Reading this property does not create a placeholder. An existing empty DATE placeholder
        returns an empty string.
        """
        placeholder = self._first_ph_of_type(PP_PLACEHOLDER.DATE)
        return None if placeholder is None else placeholder.text_frame.text

    @date_text.setter
    def date_text(self, value: str | None) -> None:
        placeholder = self._first_ph_of_type(PP_PLACEHOLDER.DATE)
        if value in (None, ""):
            if placeholder is not None:
                placeholder.text_frame.text = ""
            return

        if placeholder is None:
            layout_ph = self._layout_ph_of_type(PP_PLACEHOLDER.DATE)
            if layout_ph is None:
                raise ValueError("slide layout has no DATE placeholder to clone from")
            self.shapes.clone_placeholder(layout_ph)
            placeholder = cast("SlidePlaceholder", self._first_ph_of_type(PP_PLACEHOLDER.DATE))

        placeholder.text_frame.text = value

    @property
    def footer(self) -> str | None:
        """Text of this slide's footer placeholder, or |None| when no footer placeholder is present.

        Reading this property does not create a placeholder. An existing empty FOOTER placeholder
        returns an empty string.
        """
        placeholder = self._first_ph_of_type(PP_PLACEHOLDER.FOOTER)
        return None if placeholder is None else placeholder.text_frame.text

    @footer.setter
    def footer(self, value: str | None) -> None:
        placeholder = self._first_ph_of_type(PP_PLACEHOLDER.FOOTER)
        if value in (None, ""):
            if placeholder is not None:
                placeholder.text_frame.text = ""
            return

        if placeholder is None:
            layout_ph = self._layout_ph_of_type(PP_PLACEHOLDER.FOOTER)
            if layout_ph is None:
                raise ValueError("slide layout has no FOOTER placeholder to clone from")
            self.shapes.clone_placeholder(layout_ph)
            placeholder = cast("SlidePlaceholder", self._first_ph_of_type(PP_PLACEHOLDER.FOOTER))

        placeholder.text_frame.text = value

    @lazyproperty
    def placeholders(self) -> SlidePlaceholders:
        """Sequence of placeholder shapes in this slide."""
        return SlidePlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self) -> SlideShapes:
        """Sequence of shape objects appearing on this slide."""
        return SlideShapes(self._element.spTree, self)

    @property
    def slide_id(self) -> int:
        """Integer value that uniquely identifies this slide within this presentation.

        The slide id does not change if the position of this slide in the slide sequence is changed
        by adding, rearranging, or deleting slides.
        """
        return self.part.slide_id

    @property
    def slide_layout(self) -> SlideLayout:
        """|SlideLayout| object this slide inherits appearance from."""
        return self.part.slide_layout

    def _first_ph_of_type(self, ph_type: PP_PLACEHOLDER) -> SlidePlaceholder | None:
        """Return the first SlidePlaceholder of `ph_type` in document order, or |None|.

        This helper is non-mutating and returns the first matching slide placeholder when multiple
        placeholders of the same type are present.
        """
        for placeholder in self.placeholders:
            if placeholder.placeholder_format.type == ph_type:
                return placeholder
        return None

    def _layout_ph_of_type(self, ph_type: PP_PLACEHOLDER) -> LayoutPlaceholder | None:
        """Return the first LayoutPlaceholder of `ph_type` on this slide's layout, or |None|.

        The layout placeholder is used as the source when promoting a latent placeholder to a
        slide-level placeholder on first write.
        """
        for placeholder in self.slide_layout.placeholders:
            if placeholder.placeholder_format.type == ph_type:
                return placeholder
        return None

    def delete(self) -> None:
        """Remove this slide from its presentation.

        Convenience alias for :meth:`Slides.remove`. After this call, the slide
        is no longer reachable via the presentation's `slides` collection and
        its part is dropped from the package on save.

        Raises |ValueError| if this slide is not part of a presentation (e.g.
        the parent collection cannot be located).
        """
        prs = self.part.package.presentation_part.presentation
        prs.slides.remove(self)

    def duplicate(self, index: int | None = None) -> Slide:
        """Return a deep copy of this slide added to the parent presentation.

        Convenience alias delegating to :meth:`Slides.duplicate`. The duplicate
        is inserted at zero-based `index`; when `index` is |None|, the
        duplicate sits at ``self_index + 1`` — immediately after this slide.

        See :meth:`Slides.duplicate` for full semantics on dedup, notes-slide
        handling, and round-trip behavior.
        """
        prs = self.part.package.presentation_part.presentation
        return prs.slides.duplicate(self, index)


class Slides(ParentedElementProxy):
    """Sequence of slides belonging to an instance of |Presentation|.

    Has list semantics for access to individual slides. Supports indexed access, len(), and
    iteration.
    """

    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, sldIdLst: CT_SlideIdList, prs: Presentation):
        super(Slides, self).__init__(sldIdLst, prs)
        self._sldIdLst = sldIdLst

    def __getitem__(self, idx: int) -> Slide:
        """Provide indexed access, (e.g. 'slides[0]')."""
        try:
            sldId = self._sldIdLst.sldId_lst[idx]
        except IndexError:
            raise IndexError("slide index out of range")
        return self.part.related_slide(sldId.rId)

    def __iter__(self) -> Iterator[Slide]:
        """Support iteration, e.g. `for slide in slides:`."""
        for sldId in self._sldIdLst.sldId_lst:
            yield self.part.related_slide(sldId.rId)

    def __len__(self) -> int:
        """Support len() built-in function, e.g. `len(slides) == 4`."""
        return len(self._sldIdLst)

    def add_slide(self, slide_layout: SlideLayout, index: int | None = None) -> Slide:
        """Return a newly added slide that inherits layout from `slide_layout`.

        When `index` is |None| (the default), the new slide is appended to the
        end of the slide sequence — matching prior behavior. When `index` is
        an integer, the new slide is inserted at that zero-based position;
        `index` may equal `len(self)` to append explicitly. Negative indices
        are not supported; pass an explicit position. Raises |IndexError| if
        `index` is out of range (negative, or greater than `len(self)`).

        Companion operations: :meth:`remove`, :meth:`move`,
        :meth:`duplicate`. Cross-deck copy (``Presentation.append_from``)
        is tracked under issue #11 (Phase 3) and not yet implemented.
        """
        # ---validate index BEFORE creating the new SlidePart so a bad index
        #    does not leak a partial part into the package---
        if index is not None and (index < 0 or index > len(self._sldIdLst)):
            raise IndexError("slide index out of range")
        rId, slide = self.part.add_slide(slide_layout)
        slide.shapes.clone_layout_placeholders(slide_layout)
        if index is None:
            self._sldIdLst.add_sldId(rId)
        else:
            self._sldIdLst.insert_sldId_at(rId, index)
        return slide

    def get(self, slide_id: int, default: Slide | None = None) -> Slide | None:
        """Return the slide identified by int `slide_id` in this presentation.

        Returns `default` if not found.
        """
        slide = self.part.get_slide(slide_id)
        if slide is None:
            return default
        return slide

    def index(self, slide: Slide) -> int:
        """Map `slide` to its zero-based position in this slide sequence.

        Raises |ValueError| on *slide* not present.
        """
        for idx, this_slide in enumerate(self):
            if this_slide == slide:
                return idx
        raise ValueError("%s is not in slide collection" % slide)

    def move(self, slide: Slide, new_index: int) -> None:
        """Reposition `slide` to zero-based position `new_index`.

        Raises |ValueError| if `slide` is not a member of this collection,
        and |IndexError| if `new_index` is out of range (negative or
        ``>= len(self)``). Section membership (`p14:sectionLst`) references
        slides by id, not position, so reordering is invisible to sections.
        """
        if new_index < 0 or new_index >= len(self):
            raise IndexError("slide index out of range")
        idx = self.index(slide)
        sldId = self._sldIdLst.sldId_lst[idx]
        self._sldIdLst.move_sldId_to(sldId, new_index)

    def remove(self, slide: Slide) -> None:
        """Remove `slide` from this collection.

        The slide's relationship is dropped from the presentation part. The
        underlying slide part falls out of the package on the next save.
        Image and other media parts referenced only by the removed slide
        likewise drop out — but image parts shared with surviving slides
        (e.g. the same picture inserted on two slides) are preserved.
        Raises |ValueError| if `slide` is not a member of this collection.
        After this call, references to `slide` are stale; behavior of method
        calls on the removed `Slide` instance is undefined.
        """
        idx = self.index(slide)
        sldId = self._sldIdLst.sldId_lst[idx]
        target_rId = sldId.rId
        self._sldIdLst.remove_sldId(sldId)
        self.part.drop_rel(target_rId)

    def duplicate(self, slide: Slide, index: int | None = None) -> Slide:
        """Return a deep copy of `slide` added to this collection.

        The duplicate is inserted at zero-based position `index`. When
        `index` is |None| (the default), the new slide is inserted at
        ``source_index + 1`` — immediately after `slide`. ``index`` may
        equal ``len(self)`` to append explicitly. Negative indices are
        not supported.

        Image, media, slide-layout, and slide-master parts are shared
        with the source via package-level dedup — duplicating a slide
        that contains pictures does NOT increase the deck's image-part
        count. Chart parts, OLE-object parts, and the notes-slide (when
        present) are deep-copied so edits to the duplicate don't bleed
        back into the source. Comments parts (if any) are dropped —
        deferred to a later phase of issue #11.

        Raises |ValueError| if `slide` is not a member of this
        collection. Raises |IndexError| if `index` is out of range
        (negative or greater than `len(self)`).
        """
        from pptx.parts.slide import duplicate_notes_slide_for

        # ---validate membership BEFORE doing any work; raises ValueError if absent---
        src_idx = self.index(slide)
        if index is None:
            index = src_idx + 1
        if index < 0 or index > len(self._sldIdLst):
            raise IndexError("slide index out of range")

        src_part = slide.part
        new_slide_part = src_part.duplicate()

        # ---register new slide part with presentation; this allocates an rId---
        new_rId = self.part.relate_to(new_slide_part, RT.SLIDE)
        self._sldIdLst.insert_sldId_at(new_rId, index)

        # ---if source had a notes-slide, give the duplicate its own---
        if src_part.has_notes_slide:
            duplicate_notes_slide_for(src_part, new_slide_part)

        return new_slide_part.slide


class SlideLayout(_HeaderFooterVisibility, _BaseSlide):
    """Slide layout object.

    Provides access to placeholders, regular shapes, and slide layout-level properties.
    """

    part: SlideLayoutPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def iter_cloneable_placeholders(self) -> Iterator[LayoutPlaceholder]:
        """Generate layout-placeholders on this slide-layout that should be cloned to a new slide.

        Used when creating a new slide from this slide-layout.
        """
        latent_ph_types = (
            PP_PLACEHOLDER.DATE,
            PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.SLIDE_NUMBER,
        )
        for ph in self.placeholders:
            if ph.element.ph_type not in latent_ph_types:
                yield ph

    @lazyproperty
    def placeholders(self) -> LayoutPlaceholders:
        """Sequence of placeholder shapes in this slide layout.

        Placeholders appear in `idx` order.
        """
        return LayoutPlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self) -> LayoutShapes:
        """Sequence of shapes appearing on this slide layout."""
        return LayoutShapes(self._element.spTree, self)

    @property
    def slide_master(self) -> SlideMaster:
        """Slide master from which this slide-layout inherits properties."""
        return self.part.slide_master

    @property
    def used_by_slides(self):
        """Tuple of slide objects based on this slide layout."""
        # ---getting Slides collection requires going around the horn a bit---
        slides = self.part.package.presentation_part.presentation.slides
        return tuple(s for s in slides if s.slide_layout == self)


class SlideLayouts(ParentedElementProxy):
    """Sequence of slide layouts belonging to a slide-master.

    Supports indexed access, len(), iteration, index() and remove().
    """

    part: SlideMasterPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, sldLayoutIdLst: CT_SlideLayoutIdList, parent: SlideMaster):
        super(SlideLayouts, self).__init__(sldLayoutIdLst, parent)
        self._sldLayoutIdLst = sldLayoutIdLst

    def __getitem__(self, idx: int) -> SlideLayout:
        """Provides indexed access, e.g. `slide_layouts[2]`."""
        try:
            sldLayoutId = self._sldLayoutIdLst.sldLayoutId_lst[idx]
        except IndexError:
            raise IndexError("slide layout index out of range")
        return self.part.related_slide_layout(sldLayoutId.rId)

    def __iter__(self) -> Iterator[SlideLayout]:
        """Generate each |SlideLayout| in the collection, in sequence."""
        for sldLayoutId in self._sldLayoutIdLst.sldLayoutId_lst:
            yield self.part.related_slide_layout(sldLayoutId.rId)

    def __len__(self) -> int:
        """Support len() built-in function, e.g. `len(slides) == 4`."""
        return len(self._sldLayoutIdLst)

    def get_by_name(self, name: str, default: SlideLayout | None = None) -> SlideLayout | None:
        """Return SlideLayout object having `name`, or `default` if not found."""
        for slide_layout in self:
            if slide_layout.name == name:
                return slide_layout
        return default

    def index(self, slide_layout: SlideLayout) -> int:
        """Return zero-based index of `slide_layout` in this collection.

        Raises `ValueError` if `slide_layout` is not present in this collection.
        """
        for idx, this_layout in enumerate(self):
            if slide_layout == this_layout:
                return idx
        raise ValueError("layout not in this SlideLayouts collection")

    def remove(self, slide_layout: SlideLayout) -> None:
        """Remove `slide_layout` from the collection.

        Raises ValueError when `slide_layout` is in use; a slide layout which is the basis for one
        or more slides cannot be removed.
        """
        # ---raise if layout is in use---
        if slide_layout.used_by_slides:
            raise ValueError("cannot remove slide-layout in use by one or more slides")

        # ---target layout is identified by its index in this collection---
        target_idx = self.index(slide_layout)

        # --remove layout from p:sldLayoutIds of its master
        # --this stops layout from showing up, but doesn't remove it from package
        target_sldLayoutId = self._sldLayoutIdLst.sldLayoutId_lst[target_idx]
        self._sldLayoutIdLst.remove(target_sldLayoutId)

        # --drop relationship from master to layout
        # --this removes layout from package, along with everything (only) it refers to,
        # --including images (not used elsewhere) and hyperlinks
        slide_layout.slide_master.part.drop_rel(target_sldLayoutId.rId)


class SlideMaster(_HeaderFooterVisibility, _BaseMaster):
    """Slide master object.

    Provides access to slide layouts. Access to placeholders, regular shapes, and slide master-level
    properties is inherited from |_BaseMaster|.
    """

    _element: CT_SlideMaster  # pyright: ignore[reportIncompatibleVariableOverride]

    def add_text_watermark(
        self,
        text: str,
        *,
        font_size: Length = Pt(72),
        transparency: float = 0.7,
        font_name: str = "Calibri",
    ) -> Shape:
        """Add and return a centered watermark textbox on this slide master.

        The watermark is a large mid-gray text box sized for the standard 10in x 7.5in
        slide canvas. `transparency` is applied to the run's solid font fill.
        """
        slide_width = Inches(10)
        slide_height = Inches(7.5)
        textbox_width = Inches(6)
        textbox_height = Inches(1.5)
        # ---center a 6in x 1.5in textbox on a standard 10in x 7.5in slide---
        left = (slide_width - textbox_width) // 2
        top = (slide_height - textbox_height) // 2

        shape = self.shapes.add_textbox(left, top, textbox_width, textbox_height)
        paragraph = shape.text_frame.paragraphs[0]
        paragraph.text = text
        paragraph.alignment = PP_ALIGN.CENTER
        run = paragraph.runs[0]
        run.font.name = font_name
        run.font.size = font_size
        run.font.fill.solid()
        run.font.fill.fore_color.rgb = RGBColor(0x80, 0x80, 0x80)
        run.font.fill.transparency = transparency
        return shape

    @lazyproperty
    def slide_layouts(self) -> SlideLayouts:
        """|SlideLayouts| object providing access to this slide-master's layouts."""
        return SlideLayouts(self._element.get_or_add_sldLayoutIdLst(), self)


class SlideMasters(ParentedElementProxy):
    """Sequence of |SlideMaster| objects belonging to a presentation.

    Has list access semantics, supporting indexed access, len(), and iteration.
    """

    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, sldMasterIdLst: CT_SlideMasterIdList, parent: Presentation):
        super(SlideMasters, self).__init__(sldMasterIdLst, parent)
        self._sldMasterIdLst = sldMasterIdLst

    def __getitem__(self, idx: int) -> SlideMaster:
        """Provides indexed access, e.g. `slide_masters[2]`."""
        try:
            sldMasterId = self._sldMasterIdLst.sldMasterId_lst[idx]
        except IndexError:
            raise IndexError("slide master index out of range")
        return self.part.related_slide_master(sldMasterId.rId)

    def __iter__(self):
        """Generate each |SlideMaster| instance in the collection, in sequence."""
        for smi in self._sldMasterIdLst.sldMasterId_lst:
            yield self.part.related_slide_master(smi.rId)

    def __len__(self):
        """Support len() built-in function, e.g. `len(slide_masters) == 4`."""
        return len(self._sldMasterIdLst)


class _Background(ElementProxy):
    """Provides access to slide background properties.

    Note that the presence of this object does not by itself imply an
    explicitly-defined background; a slide with an inherited background still
    has a |_Background| object.

    Closes upstream issue #1126: prior to this fix, ``slide.background.element``
    returned the parent ``<p:cSld>`` element instead of the actual ``<p:bg>``
    background element. Power users introspecting the XML now get the right
    node. The ``<p:bg>`` element is materialized on construction (matching
    the legacy destructive behavior of accessing ``.fill``) — and since
    ``slide.background`` is a |lazyproperty| upstream, this only fires on
    first access of the slide's background, not on slide load.
    """

    def __init__(self, cSld: CT_CommonSlideData):
        # ---resolve to the <p:bg> element (creating if needed) so that
        #    `_element` and `.element` point at the right node per issue #1126
        bg = cSld.get_or_add_bg()
        super(_Background, self).__init__(bg)
        self._cSld = cSld

    @lazyproperty
    def fill(self):
        """|FillFormat| instance for this background.

        This |FillFormat| object is used to interrogate or specify the fill
        of the slide background.

        Note that accessing this property is potentially destructive. A slide
        background can also be specified by a background style reference and
        accessing this property will remove that reference, if present, and
        replace it with NoFill. This is frequently the case for a slide
        master background.

        This is also the case when there is no explicitly defined background
        (background is inherited); merely accessing this property will cause
        the background to be set to NoFill and the inheritance link will be
        interrupted. This is frequently the case for a slide background.

        Of course, if you are accessing this property in order to set the
        fill, then these changes are of no consequence, but the existing
        background cannot be reliably interrogated using this property unless
        you have already established it is an explicit fill.

        If the background is already a fill, then accessing this property
        makes no changes to the current background.
        """
        bgPr = self._cSld.get_or_add_bgPr()
        return FillFormat.from_fill_parent(bgPr)
