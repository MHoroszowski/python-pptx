"""DrawingML objects related to line formatting."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.dml.fill import FillFormat
from pptx.enum.dml import (
    MSO_FILL,
    MSO_LINE_CAP_STYLE,
    MSO_LINE_END_SIZE,
    MSO_LINE_END_TYPE,
    MSO_LINE_JOIN_STYLE,
)
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.shapes.shared import CT_LineEndProperties


class LineFormat(object):
    """Provides access to line properties such as color, style, and width.

    A LineFormat object is typically accessed via the ``.line`` property of
    a shape such as |Shape| or |Picture|.
    """

    def __init__(self, parent):
        super(LineFormat, self).__init__()
        self._parent = parent

    @property
    def begin_arrowhead_length(self) -> MSO_LINE_END_SIZE | None:
        """Size of the arrowhead at the beginning of the line.

        Read/write. Returns a member of :ref:`MsoArrowheadSize` or |None| if no explicit value
        has been set. Assigning |None| removes any existing value.
        """
        headEnd = self._headEnd
        if headEnd is None:
            return None
        return headEnd.len

    @begin_arrowhead_length.setter
    def begin_arrowhead_length(self, value: MSO_LINE_END_SIZE | None) -> None:
        if value is None:
            headEnd = self._headEnd
            if headEnd is not None:
                del headEnd.attrib["len"]
            return
        self._get_or_add_headEnd().len = value

    @property
    def begin_arrowhead_style(self) -> MSO_LINE_END_TYPE | None:
        """Type of arrowhead at the beginning of the line.

        Read/write. Returns a member of :ref:`MsoArrowheadStyle` or |None| if no explicit value
        has been set. Assigning |None| removes any existing value.
        """
        headEnd = self._headEnd
        if headEnd is None:
            return None
        return headEnd.type

    @begin_arrowhead_style.setter
    def begin_arrowhead_style(self, value: MSO_LINE_END_TYPE | None) -> None:
        if value is None:
            headEnd = self._headEnd
            if headEnd is not None:
                del headEnd.attrib["type"]
            return
        self._get_or_add_headEnd().type = value

    @property
    def begin_arrowhead_width(self) -> MSO_LINE_END_SIZE | None:
        """Width of the arrowhead at the beginning of the line.

        Read/write. Returns a member of :ref:`MsoArrowheadSize` or |None| if no explicit value
        has been set. Assigning |None| removes any existing value.
        """
        headEnd = self._headEnd
        if headEnd is None:
            return None
        return headEnd.w

    @begin_arrowhead_width.setter
    def begin_arrowhead_width(self, value: MSO_LINE_END_SIZE | None) -> None:
        if value is None:
            headEnd = self._headEnd
            if headEnd is not None:
                del headEnd.attrib["w"]
            return
        self._get_or_add_headEnd().w = value

    @property
    def cap_style(self) -> MSO_LINE_CAP_STYLE | None:
        """Cap style for this line.

        Read/write. Returns a member of :ref:`MsoLineCapStyle` or |None| if no explicit value
        has been set. Assigning |None| removes any existing value.
        """
        ln = self._ln
        if ln is None:
            return None
        return ln.cap

    @cap_style.setter
    def cap_style(self, value: MSO_LINE_CAP_STYLE | None) -> None:
        ln = self._get_or_add_ln()
        ln.cap = value

    @lazyproperty
    def color(self):
        """
        The |ColorFormat| instance that provides access to the color settings
        for this line. Essentially a shortcut for ``line.fill.fore_color``.
        As a side-effect, accessing this property causes the line fill type
        to be set to ``MSO_FILL.SOLID``. If this sounds risky for your use
        case, use ``line.fill.type`` to non-destructively discover the
        existing fill type.
        """
        if self.fill.type != MSO_FILL.SOLID:
            self.fill.solid()
        return self.fill.fore_color

    @property
    def end_arrowhead_length(self) -> MSO_LINE_END_SIZE | None:
        """Size of the arrowhead at the end of the line.

        Read/write. Returns a member of :ref:`MsoArrowheadSize` or |None| if no explicit value
        has been set. Assigning |None| removes any existing value.
        """
        tailEnd = self._tailEnd
        if tailEnd is None:
            return None
        return tailEnd.len

    @end_arrowhead_length.setter
    def end_arrowhead_length(self, value: MSO_LINE_END_SIZE | None) -> None:
        if value is None:
            tailEnd = self._tailEnd
            if tailEnd is not None:
                del tailEnd.attrib["len"]
            return
        self._get_or_add_tailEnd().len = value

    @property
    def end_arrowhead_style(self) -> MSO_LINE_END_TYPE | None:
        """Type of arrowhead at the end of the line.

        Read/write. Returns a member of :ref:`MsoArrowheadStyle` or |None| if no explicit value
        has been set. Assigning |None| removes any existing value.
        """
        tailEnd = self._tailEnd
        if tailEnd is None:
            return None
        return tailEnd.type

    @end_arrowhead_style.setter
    def end_arrowhead_style(self, value: MSO_LINE_END_TYPE | None) -> None:
        if value is None:
            tailEnd = self._tailEnd
            if tailEnd is not None:
                del tailEnd.attrib["type"]
            return
        self._get_or_add_tailEnd().type = value

    @property
    def end_arrowhead_width(self) -> MSO_LINE_END_SIZE | None:
        """Width of the arrowhead at the end of the line.

        Read/write. Returns a member of :ref:`MsoArrowheadSize` or |None| if no explicit value
        has been set. Assigning |None| removes any existing value.
        """
        tailEnd = self._tailEnd
        if tailEnd is None:
            return None
        return tailEnd.w

    @end_arrowhead_width.setter
    def end_arrowhead_width(self, value: MSO_LINE_END_SIZE | None) -> None:
        if value is None:
            tailEnd = self._tailEnd
            if tailEnd is not None:
                del tailEnd.attrib["w"]
            return
        self._get_or_add_tailEnd().w = value

    @property
    def join_style(self) -> MSO_LINE_JOIN_STYLE | None:
        """Join style for this line.

        Read/write. Returns a member of :ref:`MsoLineJoinStyle` or |None| if no explicit value
        has been set. Assigning |None| removes any existing value.
        """
        ln = self._ln
        if ln is None:
            return None
        join = ln.eg_lineJoinProperties
        if join is None:
            return None
        tag_name = join.tag.split("}")[-1]
        return {
            "round": MSO_LINE_JOIN_STYLE.ROUND,
            "bevel": MSO_LINE_JOIN_STYLE.BEVEL,
            "miter": MSO_LINE_JOIN_STYLE.MITER,
        }[tag_name]

    @join_style.setter
    def join_style(self, value: MSO_LINE_JOIN_STYLE | None) -> None:
        ln = self._get_or_add_ln()
        if value is None:
            ln._remove_eg_lineJoinProperties()
            return
        method_map = {
            MSO_LINE_JOIN_STYLE.ROUND: "get_or_change_to_round",
            MSO_LINE_JOIN_STYLE.BEVEL: "get_or_change_to_bevel",
            MSO_LINE_JOIN_STYLE.MITER: "get_or_change_to_miter",
        }
        getattr(ln, method_map[value])()

    @property
    def dash_style(self):
        """Return value indicating line style.

        Returns a member of :ref:`MsoLineDashStyle` indicating line style, or
        |None| if no explicit value has been set. When no explicit value has
        been set, the line dash style is inherited from the style hierarchy.

        Assigning |None| removes any existing explicitly-defined dash style.
        """
        ln = self._ln
        if ln is None:
            return None
        return ln.prstDash_val

    @dash_style.setter
    def dash_style(self, dash_style):
        if dash_style is None:
            ln = self._ln
            if ln is None:
                return
            ln._remove_prstDash()
            ln._remove_custDash()
            return
        ln = self._get_or_add_ln()
        ln.prstDash_val = dash_style

    @lazyproperty
    def fill(self):
        """
        |FillFormat| instance for this line, providing access to fill
        properties such as foreground color.
        """
        ln = self._get_or_add_ln()
        return FillFormat.from_fill_parent(ln)

    @property
    def width(self):
        """
        The width of the line expressed as an integer number of :ref:`English
        Metric Units <EMU>`. The returned value is an instance of |Length|,
        a value class having properties such as `.inches`, `.cm`, and `.pt`
        for converting the value into convenient units.
        """
        ln = self._ln
        if ln is None:
            return Emu(0)
        return ln.w

    @width.setter
    def width(self, emu):
        if emu is None:
            emu = 0
        ln = self._get_or_add_ln()
        ln.w = emu

    def _get_or_add_headEnd(self) -> CT_LineEndProperties:
        """Return the `a:headEnd` element, creating it if not present."""
        return self._get_or_add_ln().get_or_add_headEnd()

    def _get_or_add_ln(self):
        """
        Return the ``<a:ln>`` element containing the line format properties
        in the XML.
        """
        return self._parent.get_or_add_ln()

    def _get_or_add_tailEnd(self) -> CT_LineEndProperties:
        """Return the `a:tailEnd` element, creating it if not present."""
        return self._get_or_add_ln().get_or_add_tailEnd()

    @property
    def _headEnd(self) -> CT_LineEndProperties | None:
        """Return `a:headEnd` element or None if not present."""
        ln = self._ln
        if ln is None:
            return None
        return ln.headEnd

    @property
    def _ln(self):
        return self._parent.ln

    @property
    def _tailEnd(self) -> CT_LineEndProperties | None:
        """Return `a:tailEnd` element or None if not present."""
        ln = self._ln
        if ln is None:
            return None
        return ln.tailEnd

    @lazyproperty
    def head_end(self) -> "_LineEndFormat":
        """Arrowhead at the **begin** point of the line (issue #18 SF5).

        Returns a |_LineEndFormat| exposing ``type``, ``width`` and
        ``length`` as a single object. This is the issue-#18-named
        convenience surface over the already-shipped
        ``begin_arrowhead_style``/``_width``/``_length`` properties — those
        remain unchanged.
        """
        return _LineEndFormat(self, "head")

    @lazyproperty
    def tail_end(self) -> "_LineEndFormat":
        """Arrowhead at the **end** point of the line (issue #18 SF5).

        Issue-named convenience surface over the shipped
        ``end_arrowhead_*`` properties; those remain unchanged.
        """
        return _LineEndFormat(self, "tail")


class _LineEndFormat(object):
    """Grouped `type` / `width` / `length` view of one line arrowhead end.

    Backed entirely by the existing `LineFormat` arrowhead properties — this
    adds the issue-#18 `head_end`/`tail_end` shape without changing or
    duplicating the proven `begin/end_arrowhead_*` implementation.
    """

    def __init__(self, line: "LineFormat", which: str):
        self._line = line
        self._which = which  # "head" or "tail"

    @property
    def type(self) -> MSO_LINE_END_TYPE | None:
        """Arrowhead style — a member of :ref:`MsoArrowheadStyle` or |None|."""
        if self._which == "head":
            return self._line.begin_arrowhead_style
        return self._line.end_arrowhead_style

    @type.setter
    def type(self, value: MSO_LINE_END_TYPE | None) -> None:
        if self._which == "head":
            self._line.begin_arrowhead_style = value
        else:
            self._line.end_arrowhead_style = value

    @property
    def width(self) -> MSO_LINE_END_SIZE | None:
        """Arrowhead width — a member of :ref:`MsoArrowheadSize` or |None|."""
        if self._which == "head":
            return self._line.begin_arrowhead_width
        return self._line.end_arrowhead_width

    @width.setter
    def width(self, value: MSO_LINE_END_SIZE | None) -> None:
        if self._which == "head":
            self._line.begin_arrowhead_width = value
        else:
            self._line.end_arrowhead_width = value

    @property
    def length(self) -> MSO_LINE_END_SIZE | None:
        """Arrowhead length — a member of :ref:`MsoArrowheadSize` or |None|."""
        if self._which == "head":
            return self._line.begin_arrowhead_length
        return self._line.end_arrowhead_length

    @length.setter
    def length(self, value: MSO_LINE_END_SIZE | None) -> None:
        if self._which == "head":
            self._line.begin_arrowhead_length = value
        else:
            self._line.end_arrowhead_length = value
