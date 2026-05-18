"""Text-related objects such as TextFrame and Paragraph."""

from __future__ import annotations

import uuid
from typing import TYPE_CHECKING, Iterator, NamedTuple, cast

from pptx.dml.fill import FillFormat
from pptx.enum.dml import MSO_FILL
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import (
    MSO_AUTO_SIZE,
    MSO_TEXT_DIRECTION,
    MSO_TEXT_STRIKE_TYPE,
    MSO_UNDERLINE,
    MSO_VERTICAL_ANCHOR,
    PP_AUTO_NUMBER_STYLE,
    PP_BULLET_TYPE,
)
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.simpletypes import ST_TextWrappingType
from pptx.shapes import Subshape
from pptx.text.fonts import FontFiles
from pptx.text.layout import TextFitter
from pptx.util import Centipoints, Emu, Length, Pt, lazyproperty

if TYPE_CHECKING:
    from pptx.dml.color import ColorFormat
    from pptx.enum.text import (
        MSO_TEXT_UNDERLINE_TYPE,
        MSO_VERTICAL_ANCHOR,
        PP_PARAGRAPH_ALIGNMENT,
    )
    from pptx.oxml.action import CT_Hyperlink
    from pptx.oxml.text import (
        CT_RegularTextRun,
        CT_TextBody,
        CT_TextCharacterProperties,
        CT_TextField,
        CT_TextParagraph,
        CT_TextParagraphProperties,
    )
    from pptx.types import ProvidesExtents, ProvidesPart


class _OverflowInfo(NamedTuple):
    """Structured result of :meth:`TextFrame.overflow_info` (issue #16 SF9).

    `overflows` is the boolean answer; the other fields expose the rendered
    vs. available extents so callers can report *by how much* and which
    dimension (height) is limiting.
    """

    overflows: bool
    required_height: Length
    available_height: Length
    available_width: Length


class TextFrame(Subshape):
    """The part of a shape that contains its text.

    Not all shapes have a text frame. Corresponds to the `p:txBody` element that can
    appear as a child element of `p:sp`. Not intended to be constructed directly.
    """

    def __init__(self, txBody: CT_TextBody, parent: ProvidesPart):
        super(TextFrame, self).__init__(parent)
        self._element = self._txBody = txBody
        self._parent = parent

    def add_paragraph(self):
        """
        Return new |_Paragraph| instance appended to the sequence of
        paragraphs contained in this text frame.
        """
        p = self._txBody.add_p()
        return _Paragraph(p, self)

    @property
    def auto_size(self) -> MSO_AUTO_SIZE | None:
        """Resizing strategy used to fit text within this shape.

        Determins the type of automatic resizing used to fit the text of this shape within its
        bounding box when the text would otherwise extend beyond the shape boundaries. May be
        |None|, `MSO_AUTO_SIZE.NONE`, `MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`, or
        `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`.
        """
        return self._bodyPr.autofit

    @auto_size.setter
    def auto_size(self, value: MSO_AUTO_SIZE | None):
        self._bodyPr.autofit = value

    def clear(self):
        """Remove all paragraphs except one empty one."""
        for p in self._txBody.p_lst[1:]:
            self._txBody.remove(p)
        p = self.paragraphs[0]
        p.clear()

    def fit_text(
        self,
        font_family: str = "Calibri",
        max_size: int = 18,
        bold: bool = False,
        italic: bool = False,
        font_file: str | None = None,
    ):
        """Fit text-frame text entirely within bounds of its shape.

        Make the text in this text frame fit entirely within the bounds of its shape by setting
        word wrap on and applying the "best-fit" font size to all the text it contains.

        :attr:`TextFrame.auto_size` is set to :attr:`MSO_AUTO_SIZE.NONE`. The font size will not
        be set larger than `max_size` points. If the path to a matching TrueType font is provided
        as `font_file`, that font file will be used for the font metrics. If `font_file` is |None|,
        best efforts are made to locate a font file with matchhing `font_family`, `bold`, and
        `italic` installed on the current system (usually succeeds if the font is installed).
        """
        # ---no-op when empty as fit behavior not defined for that case---
        if self.text == "":
            return  # pragma: no cover

        font_size = self._best_fit_font_size(font_family, max_size, bold, italic, font_file)
        self._apply_fit(font_family, font_size, bold, italic)

    @property
    def margin_bottom(self) -> Length:
        """|Length| value representing the inset of text from the bottom text frame border.

        :meth:`pptx.util.Inches` provides a convenient way of setting the value, e.g.
        `text_frame.margin_bottom = Inches(0.05)`.
        """
        return self._bodyPr.bIns

    @margin_bottom.setter
    def margin_bottom(self, emu: Length):
        self._bodyPr.bIns = emu

    @property
    def margin_left(self) -> Length:
        """Inset of text from left text frame border as |Length| value."""
        return self._bodyPr.lIns

    @margin_left.setter
    def margin_left(self, emu: Length):
        self._bodyPr.lIns = emu

    @property
    def margin_right(self) -> Length:
        """Inset of text from right text frame border as |Length| value."""
        return self._bodyPr.rIns

    @margin_right.setter
    def margin_right(self, emu: Length):
        self._bodyPr.rIns = emu

    @property
    def margin_top(self) -> Length:
        """Inset of text from top text frame border as |Length| value."""
        return self._bodyPr.tIns

    @margin_top.setter
    def margin_top(self, emu: Length):
        self._bodyPr.tIns = emu

    @property
    def paragraphs(self) -> tuple[_Paragraph, ...]:
        """Sequence of paragraphs in this text frame.

        A text frame always contains at least one paragraph.
        """
        return tuple([_Paragraph(p, self) for p in self._txBody.p_lst])

    @property
    def text(self) -> str:
        """All text in this text-frame as a single string.

        Read/write. The return value contains all text in this text-frame. A line-feed character
        (`"\\n"`) separates the text for each paragraph. A vertical-tab character (`"\\v"`) appears
        for each line break (aka. soft carriage-return) encountered.

        The vertical-tab character is how PowerPoint represents a soft carriage return in clipboard
        text, which is why that encoding was chosen.

        Assignment replaces all text in the text frame. A new paragraph is added for each line-feed
        character (`"\\n"`) encountered. A line-break (soft carriage-return) is inserted for each
        vertical-tab character (`"\\v"`) encountered.

        Any control character other than newline, tab, or vertical-tab are escaped as plain-text
        like "_x001B_" (for ESC (ASCII 32) in this example).
        """
        return "\n".join(paragraph.text for paragraph in self.paragraphs)

    @text.setter
    def text(self, text: str):
        txBody = self._txBody
        txBody.clear_content()
        for p_text in text.split("\n"):
            p = txBody.add_p()
            p.append_text(p_text)

    @property
    def vertical_anchor(self) -> MSO_VERTICAL_ANCHOR | None:
        """Represents the vertical alignment of text in this text frame.

        |None| indicates the effective value should be inherited from this object's style hierarchy.
        """
        return self._txBody.bodyPr.anchor

    @vertical_anchor.setter
    def vertical_anchor(self, value: MSO_VERTICAL_ANCHOR | None):
        bodyPr = self._txBody.bodyPr
        bodyPr.anchor = value

    @property
    def word_wrap(self) -> bool | None:
        """`True` when lines of text in this shape are wrapped to fit within the shape's width.

        Read-write. Valid values are True, False, or None. True and False turn word wrap on and
        off, respectively. Assigning None to word wrap causes any word wrap setting to be removed
        from the text frame, causing it to inherit this setting from its style hierarchy.
        """
        return {
            ST_TextWrappingType.SQUARE: True,
            ST_TextWrappingType.NONE: False,
            None: None,
        }[self._txBody.bodyPr.wrap]

    @word_wrap.setter
    def word_wrap(self, value: bool | None):
        if value not in (True, False, None):
            raise ValueError(  # pragma: no cover
                "assigned value must be True, False, or None, got %s" % value
            )
        self._txBody.bodyPr.wrap = {
            True: ST_TextWrappingType.SQUARE,
            False: ST_TextWrappingType.NONE,
            None: None,
        }[value]

    @property
    def columns(self) -> int:
        """Number of text columns in this text frame (issue #16 SF6).

        Read/write int in range 1..16, backed by `a:bodyPr/@numCol`.
        Returns 1 when no explicit value is set. Assigning a value outside
        1..16 raises |ValueError|.
        """
        numCol = self._bodyPr.numCol
        return 1 if numCol is None else numCol

    @columns.setter
    def columns(self, value: int):
        if not isinstance(value, int) or value < 1 or value > 16:
            raise ValueError(f"columns must be an int in range 1..16, got {value!r}")
        self._bodyPr.numCol = value

    @property
    def column_spacing(self) -> Length | None:
        """Spacing between text columns as a |Length| (issue #16 SF6).

        Backed by `a:bodyPr/@spcCol` (EMU). |None| when unset; assigning
        |None| removes the attribute.
        """
        spcCol = self._bodyPr.spcCol
        if spcCol is None:
            return None
        return Emu(spcCol)

    @column_spacing.setter
    def column_spacing(self, value: Length | None):
        self._bodyPr.spcCol = None if value is None else Emu(value)

    @property
    def text_direction(self) -> MSO_TEXT_DIRECTION | None:
        """Flow direction of text in this frame (issue #16 SF7).

        A member of :ref:`MsoTextDirection` (e.g. `VERTICAL`, `EAST_ASIAN_
        VERTICAL`) or |None| when inherited. Backed by `a:bodyPr/@vert`.
        Assigning |None| removes the attribute.
        """
        return self._bodyPr.vert

    @text_direction.setter
    def text_direction(self, value: MSO_TEXT_DIRECTION | None):
        self._bodyPr.vert = value

    def overflow_info(
        self,
        font_family: str = "Calibri",
        bold: bool = False,
        italic: bool = False,
        font_file: str | None = None,
    ) -> "_OverflowInfo":
        """Return a structured report on whether this text overflows its shape.

        Read-only: does NOT modify the text frame or set autofit (issue #16
        SF9, closes scanny/python-pptx#1114). The report names the rendered
        vs. available height for the frame's text wrapped at the largest run
        font size present (defaulting to 18pt), using the metrics of the
        font described by `font_family`/`bold`/`italic` (or `font_file`).
        """
        avail_w, avail_h = self._extents
        text = self.text
        if text == "":
            return _OverflowInfo(False, Emu(0), Length(avail_h), Length(avail_w))
        if font_file is None:
            font_file = FontFiles.find(font_family, bold, italic)
        point_size = self._max_run_point_size()
        n_lines = TextFitter.wrapped_line_count(text, avail_w, font_file, point_size)
        line_cy = TextFitter.line_height(point_size, font_file)
        required_h = Length(line_cy * n_lines)
        return _OverflowInfo(
            overflows=required_h > avail_h,
            required_height=required_h,
            available_height=Length(avail_h),
            available_width=Length(avail_w),
        )

    def will_overflow(
        self,
        font_family: str = "Calibri",
        bold: bool = False,
        italic: bool = False,
        font_file: str | None = None,
    ) -> bool:
        """`True` if this text frame's content would overflow its shape.

        Thin boolean over :meth:`overflow_info` (issue #16 SF9). Read-only.
        """
        return self.overflow_info(font_family, bold, italic, font_file).overflows

    def shrink_text_to_fit(
        self,
        font_family: str = "Calibri",
        max_size: int = 18,
        bold: bool = False,
        italic: bool = False,
        font_file: str | None = None,
    ):
        """Eagerly shrink text via `normAutofit` so it fits without reopen.

        Sets `auto_size` to `TEXT_TO_FIT_SHAPE` and writes a computed
        `a:normAutofit/@fontScale` so the text fits inside the shape
        immediately — without depending on PowerPoint to recompute the
        scale on open (issue #16 SF10, closes scanny/python-pptx#1107). Run
        `sz` values are left unchanged (the scale is applied by the
        renderer). No-op on an empty text frame.
        """
        if self.text == "":
            return
        self.word_wrap = True
        self.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        max_pt = self._max_run_point_size(default=max_size)
        best = self._best_fit_font_size(font_family, max_pt, bold, italic, font_file)
        scale = max(1.0, min(100.0, (best / max_pt) * 100.0))
        normAutofit = self._bodyPr.normAutofit
        if normAutofit is not None:
            normAutofit.fontScale = scale

    def _max_run_point_size(self, default: int = 18) -> int:
        """Largest explicit run font size in this frame, in whole points.

        Falls back to `default` when no run sets an explicit size. Reads the
        `a:rPr/@sz` directly off each run element and never creates an
        `a:rPr` — `overflow_info`/`will_overflow` must be read-only (ISC-68).
        """
        sizes: list[int] = []
        for p in self._txBody.p_lst:
            for r in p.r_lst:
                rPr = r.rPr  # ZeroOrOne — None when absent, no mutation
                if rPr is not None and rPr.sz is not None:
                    sizes.append(int(round(Centipoints(rPr.sz).pt)))
        return max(sizes) if sizes else default

    def _apply_fit(self, font_family: str, font_size: int, is_bold: bool, is_italic: bool):
        """Arrange text in this text frame to fit inside its extents.

        This is accomplished by setting auto size off, wrap on, and setting the font of
        all its text to `font_family`, `font_size`, `is_bold`, and `is_italic`.
        """
        self.auto_size = MSO_AUTO_SIZE.NONE
        self.word_wrap = True
        self._set_font(font_family, font_size, is_bold, is_italic)

    def _best_fit_font_size(
        self, family: str, max_size: int, bold: bool, italic: bool, font_file: str | None
    ) -> int:
        """Return font-size in points that best fits text in this text-frame.

        The best-fit font size is the largest integer point size not greater than `max_size` that
        allows all the text in this text frame to fit inside its extents when rendered using the
        font described by `family`, `bold`, and `italic`. If `font_file` is specified, it is used
        to calculate the fit, whether or not it matches `family`, `bold`, and `italic`.
        """
        if font_file is None:
            font_file = FontFiles.find(family, bold, italic)
        return TextFitter.best_fit_font_size(self.text, self._extents, max_size, font_file)

    @property
    def _bodyPr(self):
        return self._txBody.bodyPr

    @property
    def _extents(self) -> tuple[Length, Length]:
        """(cx, cy) 2-tuple representing the effective rendering area of this text-frame.

        Margins are taken into account.
        """
        parent = cast("ProvidesExtents", self._parent)
        return (
            Length(parent.width - self.margin_left - self.margin_right),
            Length(parent.height - self.margin_top - self.margin_bottom),
        )

    def _set_font(self, family: str, size: int, bold: bool, italic: bool):
        """Set the font properties of all the text in this text frame."""

        def iter_rPrs(txBody: CT_TextBody) -> Iterator[CT_TextCharacterProperties]:
            for p in txBody.p_lst:
                for elm in p.content_children:
                    yield elm.get_or_add_rPr()
                # generate a:endParaRPr for each <a:p> element
                yield p.get_or_add_endParaRPr()

        def set_rPr_font(
            rPr: CT_TextCharacterProperties, name: str, size: int, bold: bool, italic: bool
        ):
            f = Font(rPr)
            f.name, f.size, f.bold, f.italic = family, Pt(size), bold, italic

        txBody = self._element
        for rPr in iter_rPrs(txBody):
            set_rPr_font(rPr, family, size, bold, italic)


class Font(object):
    """Character properties object, providing font size, font name, bold, italic, etc.

    Corresponds to `a:rPr` child element of a run. Also appears as `a:defRPr` and
    `a:endParaRPr` in paragraph and `a:defRPr` in list style elements.
    """

    def __init__(self, rPr: CT_TextCharacterProperties):
        super(Font, self).__init__()
        self._element = self._rPr = rPr

    @property
    def bold(self) -> bool | None:
        """Get or set boolean bold value of |Font|, e.g. `paragraph.font.bold = True`.

        If set to |None|, the bold setting is cleared and is inherited from an enclosing shape's
        setting, or a setting in a style or master. Returns None if no bold attribute is present,
        meaning the effective bold value is inherited from a master or the theme.
        """
        return self._rPr.b

    @bold.setter
    def bold(self, value: bool | None):
        self._rPr.b = value

    @lazyproperty
    def color(self) -> "_LazyFontColorFormat":
        """The |ColorFormat| instance that provides access to the color settings for this font.

        Reading from the returned object on a font with no `<a:solidFill>` does not modify
        the underlying XML — `font.color.type` and `font.color.rgb` simply return |None|.
        Setting any color property (`rgb`, `theme_color`, `brightness`) materializes
        `<a:solidFill>` lazily on first write, then delegates to the real |ColorFormat|.

        Closes scanny/python-pptx#1111 and #1074 — the prior implementation called
        ``self.fill.solid()`` on every read, mutating the document on access.
        """
        return _LazyFontColorFormat(self)

    @lazyproperty
    def fill(self) -> FillFormat:
        """|FillFormat| instance for this font.

        Provides access to fill properties such as fill color.
        """
        return FillFormat.from_fill_parent(self._rPr)

    @property
    def italic(self) -> bool | None:
        """Get or set boolean italic value of |Font| instance.

        Has the same behaviors as bold with respect to None values.
        """
        return self._rPr.i

    @italic.setter
    def italic(self, value: bool | None):
        self._rPr.i = value

    @property
    def language_id(self) -> MSO_LANGUAGE_ID | None:
        """Get or set the language id of this |Font| instance.

        The language id is a member of the :ref:`MsoLanguageId` enumeration. Assigning |None|
        removes any language setting, the same behavior as assigning `MSO_LANGUAGE_ID.NONE`.
        """
        lang = self._rPr.lang
        if lang is None:
            return MSO_LANGUAGE_ID.NONE
        return self._rPr.lang

    @language_id.setter
    def language_id(self, value: MSO_LANGUAGE_ID | None):
        if value == MSO_LANGUAGE_ID.NONE:
            value = None
        self._rPr.lang = value

    @property
    def name(self) -> str | None:
        """Get or set the typeface name for this |Font| instance.

        Causes the text it controls to appear in the named font, if a matching font is found.
        Returns |None| if the typeface is currently inherited from the theme. Setting it to |None|
        removes any override of the theme typeface.
        """
        latin = self._rPr.latin
        if latin is None:
            return None
        return latin.typeface

    @name.setter
    def name(self, value: str | None):
        if value is None:
            self._rPr._remove_latin()  # pyright: ignore[reportPrivateUsage]
        else:
            latin = self._rPr.get_or_add_latin()
            latin.typeface = value

    @property
    def size(self) -> Length | None:
        """Indicates the font height in English Metric Units (EMU).

        Read/write. |None| indicates the font size should be inherited from its style hierarchy,
        such as a placeholder or document defaults (usually 18pt). |Length| is a subclass of |int|
        having properties for convenient conversion into points or other length units. Likewise,
        the :class:`pptx.util.Pt` class allows convenient specification of point values::

            >>> font.size = Pt(24)
            >>> font.size
            304800
            >>> font.size.pt
            24.0
        """
        sz = self._rPr.sz
        if sz is None:
            return None
        return Centipoints(sz)

    @size.setter
    def size(self, emu: Length | None):
        if emu is None:
            self._rPr.sz = None
        else:
            sz = Emu(emu).centipoints
            self._rPr.sz = sz

    @property
    def underline(self) -> bool | MSO_TEXT_UNDERLINE_TYPE | None:
        """Indicaties the underline setting for this font.

        Value is |True|, |False|, |None|, or a member of the :ref:`MsoTextUnderlineType`
        enumeration. |None| is the default and indicates the underline setting should be inherited
        from the style hierarchy, such as from a placeholder. |True| indicates single underline.
        |False| indicates no underline. Other settings such as double and wavy underlining are
        indicated with members of the :ref:`MsoTextUnderlineType` enumeration.
        """
        u = self._rPr.u
        if u is MSO_UNDERLINE.NONE:
            return False
        if u is MSO_UNDERLINE.SINGLE_LINE:
            return True
        return u

    @underline.setter
    def underline(self, value: bool | MSO_TEXT_UNDERLINE_TYPE | None):
        if value is True:
            value = MSO_UNDERLINE.SINGLE_LINE
        elif value is False:
            value = MSO_UNDERLINE.NONE
        self._element.u = value

    @property
    def strike(self) -> MSO_TEXT_STRIKE_TYPE | None:
        """Strikethrough setting for this font (issue #16 SF2).

        A member of :ref:`MsoTextStrikeType` (`NONE`/`SINGLE`/`DOUBLE`) or
        |None| when no explicit value is set (inherited from the style
        hierarchy). Assigning |None| removes the attribute (restoring
        inheritance); assigning `MSO_TEXT_STRIKE_TYPE.NONE` writes an
        explicit `strike="noStrike"`.
        """
        return self._rPr.strike

    @strike.setter
    def strike(self, value: MSO_TEXT_STRIKE_TYPE | None):
        self._rPr.strike = value

    @property
    def superscript(self) -> bool | None:
        """Whether this font is superscript (issue #16 SF1).

        Backed by `a:rPr/@baseline` (a signed percentage). |True| when the
        baseline is positive, |False| when it is zero/negative, |None| when
        no baseline is set. Assigning |True| sets a +30% baseline; |False|
        or |None| removes the baseline (and thus any subscript too —
        super/subscript are mutually exclusive).
        """
        baseline = self._rPr.baseline
        if baseline is None:
            return None
        return baseline > 0

    @superscript.setter
    def superscript(self, value: bool | None):
        if value:
            # ---ST_Percentage python value is a fraction: 0.30 -> "30000"
            self._rPr.baseline = 0.30
        else:
            self._rPr.baseline = None

    @property
    def subscript(self) -> bool | None:
        """Whether this font is subscript (issue #16 SF1).

        Backed by `a:rPr/@baseline`. |True| when the baseline is negative,
        |False| when zero/positive, |None| when unset. Assigning |True| sets
        a -25% baseline; |False|/|None| removes the baseline.
        """
        baseline = self._rPr.baseline
        if baseline is None:
            return None
        return baseline < 0

    @subscript.setter
    def subscript(self, value: bool | None):
        if value:
            # ---fraction: -0.25 -> "-25000"
            self._rPr.baseline = -0.25
        else:
            self._rPr.baseline = None

    @property
    def character_spacing(self) -> Length | None:
        """Inter-character spacing (`a:rPr/@spc`) as a |Length| (issue #16 SF4).

        Read/write. |None| when no explicit value is set. Negative values
        tighten spacing, positive values loosen it. Assigning |None| removes
        the attribute.
        """
        spc = self._rPr.spc
        if spc is None:
            return None
        return Centipoints(spc)

    @character_spacing.setter
    def character_spacing(self, value: Length | None):
        if value is None:
            self._rPr.spc = None
        else:
            self._rPr.spc = Emu(value).centipoints

    @property
    def kerning(self) -> Length | None:
        """Minimum font size at which kerning is applied (`a:rPr/@kern`).

        Read/write |Length| or |None| (issue #16 SF4). Non-negative.
        """
        kern = self._rPr.kern
        if kern is None:
            return None
        return Centipoints(kern)

    @kerning.setter
    def kerning(self, value: Length | None):
        if value is None:
            self._rPr.kern = None
        else:
            self._rPr.kern = Emu(value).centipoints

    @property
    def latin(self) -> str | None:
        """The Latin-script typeface name (`a:rPr/a:latin`) (issue #16 SF5).

        Equivalent to :attr:`name`; provided so the latin/east-asian/
        complex-script trio reads symmetrically. Assigning |None| removes
        the `<a:latin>` child.
        """
        return self.name

    @latin.setter
    def latin(self, value: str | None):
        self.name = value

    @property
    def east_asian(self) -> str | None:
        """The East-Asian typeface name (`a:rPr/a:ea`) (issue #16 SF5).

        Independent of :attr:`name`/:attr:`latin` — setting this never
        touches `<a:latin>`. Assigning |None| removes the `<a:ea>` child.
        """
        ea = self._rPr.ea
        if ea is None:
            return None
        return ea.typeface

    @east_asian.setter
    def east_asian(self, value: str | None):
        if value is None:
            self._rPr._remove_ea()  # pyright: ignore[reportPrivateUsage]
        else:
            self._rPr.get_or_add_ea().typeface = value

    @property
    def complex_script(self) -> str | None:
        """The complex-script typeface name (`a:rPr/a:cs`) (issue #16 SF5).

        Independent of :attr:`name`/:attr:`latin`. Assigning |None| removes
        the `<a:cs>` child.
        """
        cs = self._rPr.cs
        if cs is None:
            return None
        return cs.typeface

    @complex_script.setter
    def complex_script(self, value: str | None):
        if value is None:
            self._rPr._remove_cs()  # pyright: ignore[reportPrivateUsage]
        else:
            self._rPr.get_or_add_cs().typeface = value

    @lazyproperty
    def highlight(self) -> "_HighlightColor":
        """Text highlight color (`a:rPr/a:highlight`) (issue #16 SF3).

        Returns a |_HighlightColor| proxy. Reading `.rgb` when no highlight
        is set returns |None| and does NOT mutate the XML; assigning `.rgb`
        materializes `<a:highlight>` (schema-ordered before the typeface
        trio).
        """
        return _HighlightColor(self._rPr)


class _HighlightColor:
    """Lazy `<a:highlight>` color proxy for :attr:`Font.highlight` (issue #16 SF3).

    Read of `.rgb`/`.type` when no `<a:highlight>` is present returns |None|
    without touching the XML; writing `.rgb` creates the element on demand.
    """

    def __init__(self, rPr: "CT_TextCharacterProperties"):
        self._rPr = rPr

    @property
    def type(self):
        from pptx.dml.color import ColorFormat

        hl = self._rPr.highlight
        if hl is None:
            return None
        return ColorFormat.from_colorchoice_parent(hl).type

    @property
    def rgb(self):
        from pptx.dml.color import ColorFormat

        hl = self._rPr.highlight
        if hl is None:
            return None
        return ColorFormat.from_colorchoice_parent(hl).rgb

    @rgb.setter
    def rgb(self, value):
        from pptx.dml.color import ColorFormat

        hl = self._rPr.get_or_add_highlight()
        ColorFormat.from_colorchoice_parent(hl).rgb = value

    @property
    def visible(self) -> bool:
        """|True| if an `<a:highlight>` element is present."""
        return self._rPr.highlight is not None


class _LazyFontColorFormat:
    """ColorFormat-shaped proxy that defers `<a:solidFill/>` creation until first SET.

    Wraps a |Font| instance. On reads (``type``, ``rgb``, ``theme_color``,
    ``brightness``, ``transparency``), if the font has no solid fill the proxy
    returns |None| / inherit values without modifying the XML. On writes,
    materializes ``<a:solidFill/>`` via ``font.fill.solid()`` and delegates to the
    real |ColorFormat|.

    Fixes scanny/python-pptx#1111 and #1074 — the prior `Font.color` getter called
    `self.fill.solid()` unconditionally on every read.
    """

    def __init__(self, font: "Font"):
        self._font = font

    # ---internal helpers ---------------------------------------------------

    def _real_or_none(self) -> "ColorFormat | None":
        """Return the real ColorFormat over `<a:solidFill>` if present, else |None|.

        Read path. Does NOT mutate the underlying XML.
        """
        if self._font.fill.type == MSO_FILL.SOLID:
            return self._font.fill.fore_color
        return None

    def _real_mutating(self) -> "ColorFormat":
        """Materialize `<a:solidFill/>` on the font and return its real ColorFormat.

        Write path. Mutates the XML on first call by inserting `<a:solidFill/>`
        when not already present.
        """
        if self._font.fill.type != MSO_FILL.SOLID:
            self._font.fill.solid()
        return self._font.fill.fore_color

    # ---public API mirroring ColorFormat ----------------------------------

    @property
    def type(self):
        real = self._real_or_none()
        return real.type if real is not None else None

    @property
    def rgb(self):
        real = self._real_or_none()
        return real.rgb if real is not None else None

    @rgb.setter
    def rgb(self, value):
        self._real_mutating().rgb = value

    @property
    def theme_color(self):
        # ---no fill = "inheriting from style", which is None.
        # ---NOT_THEME_COLOR is reserved for "solidFill present but no
        # ---schemeClr child" — i.e. the explicit-RGB case. Conflating the
        # ---two would let a round-trip read/write break inheritance.
        real = self._real_or_none()
        return real.theme_color if real is not None else None

    @theme_color.setter
    def theme_color(self, value):
        self._real_mutating().theme_color = value

    @property
    def brightness(self):
        # ---no fill = inherit; return None. 0.0 is a real settable value
        # ---meaning "no brightness adjustment", not the inherit signal.
        real = self._real_or_none()
        return real.brightness if real is not None else None

    @brightness.setter
    def brightness(self, value):
        self._real_mutating().brightness = value

    @property
    def transparency(self):
        # ---no fill = inherit; return None. 0.0 is a real settable value
        # ---meaning "fully opaque", not the inherit signal.
        real = self._real_or_none()
        return real.transparency if real is not None else None

    @transparency.setter
    def transparency(self, value):
        self._real_mutating().transparency = value


class _Hyperlink(Subshape):
    """Text run hyperlink object.

    Corresponds to `a:hlinkClick` child element of the run's properties element (`a:rPr`).
    """

    def __init__(self, rPr: CT_TextCharacterProperties, parent: ProvidesPart):
        super(_Hyperlink, self).__init__(parent)
        self._rPr = rPr

    @property
    def address(self) -> str | None:
        """The URL of the hyperlink.

        Read/write. URL can be on http, https, mailto, or file scheme; others may work.
        """
        if self._hlinkClick is None:
            return None
        return self.part.target_ref(self._hlinkClick.rId)

    @address.setter
    def address(self, url: str | None):
        # implements all three of add, change, and remove hyperlink
        if self._hlinkClick is not None:
            self._remove_hlinkClick()
        if url:
            self._add_hlinkClick(url)

    def _add_hlinkClick(self, url: str):
        rId = self.part.relate_to(url, RT.HYPERLINK, is_external=True)
        self._rPr.add_hlinkClick(rId)

    @property
    def _hlinkClick(self) -> CT_Hyperlink | None:
        return self._rPr.hlinkClick

    def _remove_hlinkClick(self):
        assert self._hlinkClick is not None
        self.part.drop_rel(self._hlinkClick.rId)
        self._rPr._remove_hlinkClick()  # pyright: ignore[reportPrivateUsage]


class _Paragraph(Subshape):
    """Paragraph object. Not intended to be constructed directly."""

    def __init__(self, p: CT_TextParagraph, parent: ProvidesPart):
        super(_Paragraph, self).__init__(parent)
        self._element = self._p = p

    def add_line_break(self):
        """Add line break at end of this paragraph."""
        self._p.add_br()

    def add_run(self) -> _Run:
        """Return a new run appended to the runs in this paragraph."""
        r = self._p.add_r()
        return _Run(r, self)

    def add_field(self) -> _Field:
        """Return a new |_Field| appended after the paragraph's existing content.

        The new ``<a:fld>`` element is given a fresh RFC-4122 v4 GUID `id`
        wrapped in braces, with uppercase hex — matching the authoring format
        PowerPoint emits when the user runs *Insert → Slide Number* or
        *Insert → Date and Time*. The caller is expected to set `type` (e.g.
        `"slidenum"`, `"datetime1"`) and optionally `text` (the placeholder
        glyph PowerPoint displays for the field before it resolves the live
        value) on the returned `_Field`.
        """
        f = self._p._add_fld()
        f.id = "{%s}" % str(uuid.uuid4()).upper()
        return _Field(f, self)

    @property
    def alignment(self) -> PP_PARAGRAPH_ALIGNMENT | None:
        """Horizontal alignment of this paragraph.

        The value |None| indicates the paragraph should 'inherit' its effective value from its
        style hierarchy. Assigning |None| removes any explicit setting, causing its inherited
        value to be used.
        """
        return self._pPr.algn

    @alignment.setter
    def alignment(self, value: PP_PARAGRAPH_ALIGNMENT | None):
        self._pPr.algn = value

    @property
    def rtl(self) -> bool | None:
        """Right-to-left setting for this paragraph (issue #16 SF8).

        Backed by `a:pPr/@rtl`. |True| flows the paragraph right-to-left
        (Arabic, Hebrew, Persian); |False| forces left-to-right; |None|
        (default) removes the attribute so direction is inherited.
        PowerPoint performs the actual bidi shaping.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.rtl

    @rtl.setter
    def rtl(self, value: bool | None):
        self._pPr.rtl = value

    @property
    def bullet_char(self) -> str | None:
        """Character used for bullet, e.g. '•'.

        Read/write. Returns |None| if the paragraph does not have a character bullet. Setting this
        property also sets `bullet_type` to `PP_BULLET_TYPE.CHARACTER`.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        buChar = pPr.buChar
        if buChar is None:
            return None
        return buChar.get("char")

    @bullet_char.setter
    def bullet_char(self, value: str | None) -> None:
        pPr = self._p.get_or_add_pPr()
        if value is None:
            pPr._remove_eg_buTypeface()
            return
        buChar = pPr.get_or_change_to_buChar()
        buChar.set("char", value)

    @property
    def bullet_type(self) -> PP_BULLET_TYPE | None:
        """Type of bullet formatting on this paragraph.

        Read/write. Returns a member of :ref:`PpBulletType` or |None| if no explicit bullet
        formatting is set. Assigning |None| removes bullet formatting.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        bu = pPr.eg_buTypeface
        if bu is None:
            return None
        tag = bu.tag.split("}")[-1]
        return {
            "buNone": PP_BULLET_TYPE.NONE,
            "buChar": PP_BULLET_TYPE.CHARACTER,
            "buAutoNum": PP_BULLET_TYPE.AUTO_NUMBER,
        }.get(tag)

    @bullet_type.setter
    def bullet_type(self, value: PP_BULLET_TYPE | None) -> None:
        pPr = self._p.get_or_add_pPr()
        if value is None:
            pPr._remove_eg_buTypeface()
            return
        method_map = {
            PP_BULLET_TYPE.NONE: "get_or_change_to_buNone",
            PP_BULLET_TYPE.CHARACTER: "get_or_change_to_buChar",
            PP_BULLET_TYPE.AUTO_NUMBER: "get_or_change_to_buAutoNum",
        }
        getattr(pPr, method_map[value])()

    @property
    def bullet_auto_number_type(self) -> PP_AUTO_NUMBER_STYLE | None:
        """Auto-number style for this paragraph's bullet.

        Read/write. Returns a member of :ref:`PpAutoNumberStyle` or |None|. Setting this property
        also sets `bullet_type` to `PP_BULLET_TYPE.AUTO_NUMBER`.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        buAutoNum = pPr.buAutoNum
        if buAutoNum is None:
            return None
        type_val = buAutoNum.get("type")
        if type_val is None:
            return None
        return PP_AUTO_NUMBER_STYLE.from_xml(type_val)

    @bullet_auto_number_type.setter
    def bullet_auto_number_type(self, value: PP_AUTO_NUMBER_STYLE | None) -> None:
        pPr = self._p.get_or_add_pPr()
        if value is None:
            pPr._remove_eg_buTypeface()
            return
        buAutoNum = pPr.get_or_change_to_buAutoNum()
        buAutoNum.set("type", value.xml_value)

    @property
    def bullet_font(self) -> str | None:
        """Typeface name for bullet character.

        Read/write. Returns |None| if no explicit bullet font is set.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        buFont = pPr.buFont
        if buFont is None:
            return None
        return buFont.get("typeface")

    @bullet_font.setter
    def bullet_font(self, value: str | None) -> None:
        pPr = self._p.get_or_add_pPr()
        if value is None:
            pPr._remove_buFont()
            return
        buFont = pPr.get_or_add_buFont()
        buFont.set("typeface", value)

    def clear(self):
        """Remove all content from this paragraph.

        Paragraph properties are preserved. Content includes runs, line breaks, and fields.
        """
        for elm in self._element.content_children:
            self._element.remove(elm)
        return self

    @property
    def font(self) -> Font:
        """|Font| object containing default character properties for the runs in this paragraph.

        These character properties override default properties inherited from parent objects such
        as the text frame the paragraph is contained in and they may be overridden by character
        properties set at the run level.
        """
        return Font(self._defRPr)

    @property
    def level(self) -> int:
        """Indentation level of this paragraph.

        Read-write. Integer in range 0..8 inclusive. 0 represents a top-level paragraph and is the
        default value. Indentation level is most commonly encountered in a bulleted list, as is
        found on a word bullet slide.
        """
        return self._pPr.lvl

    @level.setter
    def level(self, level: int):
        self._pPr.lvl = level

    @property
    def line_spacing(self) -> int | float | Length | None:
        """The space between baselines in successive lines of this paragraph.

        A value of |None| indicates no explicit value is assigned and its effective value is
        inherited from the paragraph's style hierarchy. A numeric value, e.g. `2` or `1.5`,
        indicates spacing is applied in multiples of line heights. A |Length| value such as
        `Pt(12)` indicates spacing is a fixed height. The |Pt| value class is a convenient way to
        apply line spacing in units of points.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.line_spacing

    @line_spacing.setter
    def line_spacing(self, value: int | float | Length | None):
        pPr = self._p.get_or_add_pPr()
        pPr.line_spacing = value

    @property
    def runs(self) -> tuple[_Run, ...]:
        """Sequence of runs in this paragraph."""
        return tuple(_Run(r, self) for r in self._element.r_lst)

    @property
    def fields(self) -> tuple[_Field, ...]:
        """Sequence of fields in this paragraph in document order.

        Mirrors :attr:`runs` but yields :class:`_Field` instances wrapping each
        ``<a:fld>`` child element. Useful for discovering existing slide-number,
        date, and other PowerPoint-resolved fields in a deck — `.runs` deliberately
        excludes fields so that pre-existing iteration semantics stay intact.
        """
        return tuple(_Field(f, self) for f in self._element.fld_lst)

    @property
    def space_after(self) -> Length | None:
        """The spacing to appear between this paragraph and the subsequent paragraph.

        A value of |None| indicates no explicit value is assigned and its effective value is
        inherited from the paragraph's style hierarchy. |Length| objects provide convenience
        properties, such as `.pt` and `.inches`, that allow easy conversion to various length
        units.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.space_after

    @space_after.setter
    def space_after(self, value: Length | None):
        pPr = self._p.get_or_add_pPr()
        pPr.space_after = value

    @property
    def space_before(self) -> Length | None:
        """The spacing to appear between this paragraph and the prior paragraph.

        A value of |None| indicates no explicit value is assigned and its effective value is
        inherited from the paragraph's style hierarchy. |Length| objects provide convenience
        properties, such as `.pt` and `.cm`, that allow easy conversion to various length units.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.space_before

    @space_before.setter
    def space_before(self, value: Length | None):
        pPr = self._p.get_or_add_pPr()
        pPr.space_before = value

    @property
    def text(self) -> str:
        """Text of paragraph as a single string.

        Read/write. This value is formed by concatenating the text in each run and field making up
        the paragraph, adding a vertical-tab character (`"\\v"`) for each line-break element
        (`<a:br>`, soft carriage-return) encountered.

        While the encoding of line-breaks as a vertical tab might be surprising at first, doing so
        is consistent with PowerPoint's clipboard copy behavior and allows a line-break to be
        distinguished from a paragraph boundary within the str return value.

        Assignment causes all content in the paragraph to be replaced. Each vertical-tab character
        (`"\\v"`) in the assigned str is translated to a line-break, as is each line-feed
        character (`"\\n"`). Contrast behavior of line-feed character in `TextFrame.text` setter.
        If line-feed characters are intended to produce new paragraphs, use `TextFrame.text`
        instead. Any other control characters in the assigned string are escaped as a hex
        representation like "_x001B_" (for ESC (ASCII 27) in this example).
        """
        return "".join(elm.text for elm in self._element.content_children)

    @text.setter
    def text(self, text: str):
        self.clear()
        self._element.append_text(text)

    @property
    def _defRPr(self) -> CT_TextCharacterProperties:
        """The element that defines the default run properties for runs in this paragraph.

        Causes the element to be added if not present.
        """
        return self._pPr.get_or_add_defRPr()

    @property
    def _pPr(self) -> CT_TextParagraphProperties:
        """Contains the properties for this paragraph.

        Causes the element to be added if not present.
        """
        return self._p.get_or_add_pPr()


class _Run(Subshape):
    """Text run object. Corresponds to `a:r` child element in a paragraph."""

    def __init__(self, r: CT_RegularTextRun, parent: ProvidesPart):
        super(_Run, self).__init__(parent)
        self._r = r

    @property
    def font(self):
        """|Font| instance containing run-level character properties for the text in this run.

        Character properties can be and perhaps most often are inherited from parent objects such
        as the paragraph and slide layout the run is contained in. Only those specifically
        overridden at the run level are contained in the font object.
        """
        rPr = self._r.get_or_add_rPr()
        return Font(rPr)

    @lazyproperty
    def hyperlink(self) -> _Hyperlink:
        """Proxy for any `a:hlinkClick` element under the run properties element.

        Created on demand, the hyperlink object is available whether an `a:hlinkClick` element is
        present or not, and creates or deletes that element as appropriate in response to actions
        on its methods and attributes.
        """
        rPr = self._r.get_or_add_rPr()
        return _Hyperlink(rPr, self)

    @property
    def text(self):
        """Read/write. A unicode string containing the text in this run.

        Assignment replaces all text in the run. The assigned value can be a 7-bit ASCII
        string, a UTF-8 encoded 8-bit string, or unicode. String values are converted to
        unicode assuming UTF-8 encoding.

        Any other control characters in the assigned string other than tab or newline
        are escaped as a hex representation. For example, ESC (ASCII 27) is escaped as
        "_x001B_". Contrast the behavior of `TextFrame.text` and `_Paragraph.text` with
        respect to line-feed and vertical-tab characters.
        """
        return self._r.text

    @text.setter
    def text(self, text: str):
        self._r.text = text


class _Field(Subshape):
    """Field object. Corresponds to ``<a:fld>`` child element in a paragraph.

    A field renders text whose value PowerPoint resolves at slide-show or open
    time — slide numbers, the current date, the slide title, etc. The literal
    text written to the ``<a:t>`` child is the placeholder PowerPoint shows
    before it resolves the live value; users typically pass a glyph like
    ``"‹#›"`` for slide numbers or the current date as a static fallback.

    Not intended to be constructed directly — obtain instances from
    :meth:`_Paragraph.add_field`.
    """

    def __init__(self, f: CT_TextField, parent: ProvidesPart):
        super(_Field, self).__init__(parent)
        self._f = f

    @property
    def font(self) -> Font:
        """|Font| instance for the run-level character properties of this field.

        Character properties can be and perhaps most often are inherited from
        parent objects such as the paragraph and slide layout the field is
        contained in. Only those specifically overridden at the field level
        are contained in the font object.
        """
        rPr = self._f.get_or_add_rPr()
        return Font(rPr)

    @property
    def text(self) -> str:
        """Read/write. A unicode string containing the field's placeholder text.

        Assignment replaces all text in the field. Control characters other
        than tab or newline are escaped as a hex representation. For example,
        ESC (ASCII 27) is escaped as ``"_x001B_"``.
        """
        return self._f.text

    @text.setter
    def text(self, text: str):
        self._f.text = text

    @property
    def type(self) -> str | None:
        """Read/write. The field's ``type`` attribute, e.g. ``"slidenum"``.

        ECMA-376 §A.4.1 names the well-known types: ``slidenum``,
        ``datetime1`` .. ``datetime13``, and ``title``. The OOXML schema
        itself treats the value as a permissive string. Returns |None| when
        no ``type`` attribute is present.
        """
        return self._f.type

    @type.setter
    def type(self, value: str | None):
        self._f.type = value
