"""Table-related objects such as Table and Cell."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Iterator

from pptx.dml.fill import FillFormat
from pptx.dml.line import LineFormat
from pptx.oxml.table import TcRange
from pptx.shapes import Subshape
from pptx.text.text import TextFrame
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.enum.text import MSO_VERTICAL_ANCHOR
    from pptx.oxml.table import CT_Table, CT_TableCell, CT_TableCol, CT_TableRow
    from pptx.parts.slide import BaseSlidePart
    from pptx.shapes.graphfrm import GraphicFrame
    from pptx.types import ProvidesPart
    from pptx.util import Length


class Table(object):
    """A DrawingML table object.

    Not intended to be constructed directly, use
    :meth:`.Slide.shapes.add_table` to add a table to a slide.
    """

    def __init__(self, tbl: CT_Table, graphic_frame: GraphicFrame):
        super(Table, self).__init__()
        self._tbl = tbl
        self._graphic_frame = graphic_frame

    def cell(self, row_idx: int, col_idx: int) -> _Cell:
        """Return cell at `row_idx`, `col_idx`.

        Return value is an instance of |_Cell|. `row_idx` and `col_idx` are zero-based, e.g.
        cell(0, 0) is the top, left cell in the table.
        """
        return _Cell(self._tbl.tc(row_idx, col_idx), self)

    @lazyproperty
    def columns(self) -> _ColumnCollection:
        """|_ColumnCollection| instance for this table.

        Provides access to |_Column| objects representing the table's columns. |_Column| objects
        are accessed using list notation, e.g. `col = tbl.columns[0]`.
        """
        return _ColumnCollection(self._tbl, self)

    @property
    def first_col(self) -> bool:
        """When `True`, indicates first column should have distinct formatting.

        Read/write. Distinct formatting is used, for example, when the first column contains row
        headings (is a side-heading column).
        """
        return self._tbl.firstCol

    @first_col.setter
    def first_col(self, value: bool):
        self._tbl.firstCol = value

    @property
    def first_row(self) -> bool:
        """When `True`, indicates first row should have distinct formatting.

        Read/write. Distinct formatting is used, for example, when the first row contains column
        headings.
        """
        return self._tbl.firstRow

    @first_row.setter
    def first_row(self, value: bool):
        self._tbl.firstRow = value

    @property
    def horz_banding(self) -> bool:
        """When `True`, indicates rows should have alternating shading.

        Read/write. Used to allow rows to be traversed more easily without losing track of which
        row is being read.
        """
        return self._tbl.bandRow

    @horz_banding.setter
    def horz_banding(self, value: bool):
        self._tbl.bandRow = value

    def iter_cells(self) -> Iterator[_Cell]:
        """Generate _Cell object for each cell in this table.

        Each grid cell is generated in left-to-right, top-to-bottom order.
        """
        return (_Cell(tc, self) for tc in self._tbl.iter_tcs())

    @property
    def last_col(self) -> bool:
        """When `True`, indicates the rightmost column should have distinct formatting.

        Read/write. Used, for example, when a row totals column appears at the far right of the
        table.
        """
        return self._tbl.lastCol

    @last_col.setter
    def last_col(self, value: bool):
        self._tbl.lastCol = value

    @property
    def last_row(self) -> bool:
        """When `True`, indicates the bottom row should have distinct formatting.

        Read/write. Used, for example, when a totals row appears as the bottom row.
        """
        return self._tbl.lastRow

    @last_row.setter
    def last_row(self, value: bool):
        self._tbl.lastRow = value

    def notify_height_changed(self) -> None:
        """Called by a row when its height changes.

        Triggers the graphic frame to recalculate its total height (as the sum of the row
        heights).
        """
        new_table_height = Emu(sum([row.height for row in self.rows]))
        self._graphic_frame.height = new_table_height

    def notify_width_changed(self) -> None:
        """Called by a column when its width changes.

        Triggers the graphic frame to recalculate its total width (as the sum of the column
        widths).
        """
        new_table_width = Emu(sum([col.width for col in self.columns]))
        self._graphic_frame.width = new_table_width

    @property
    def part(self) -> BaseSlidePart:
        """The package part containing this table."""
        return self._graphic_frame.part

    @lazyproperty
    def rows(self):
        """|_RowCollection| instance for this table.

        Provides access to |_Row| objects representing the table's rows. |_Row| objects are
        accessed using list notation, e.g. `col = tbl.rows[0]`.
        """
        return _RowCollection(self._tbl, self)

    @property
    def vert_banding(self) -> bool:
        """When `True`, indicates columns should have alternating shading.

        Read/write. Used to allow columns to be traversed more easily without losing track of
        which column is being read.
        """
        return self._tbl.bandCol

    @vert_banding.setter
    def vert_banding(self, value: bool):
        self._tbl.bandCol = value

    def merge_cells(self, row_range, col_range) -> "_Cell":
        """Merge a rectangular block of cells into a single merged cell.

        ``row_range`` and ``col_range`` accept either:

        - a 2-tuple ``(start, end)`` interpreted as **inclusive** indices —
          ``(0, 1)`` covers rows 0 and 1.
        - a Python ``range`` object — half-open per Python convention; ``range(0, 2)``
          covers rows 0 and 1.

        The order within each range is irrelevant: ``(2, 0)`` is the same as ``(0, 2)``.

        Idempotent: if the entire requested range is already merged exactly
        as a single block with the same origin and dimensions, the call is
        a no-op and returns the existing merge-origin cell. Calling on a
        single-cell range that is not merged is also a no-op (no merge is
        needed for one cell).

        Raises |ValueError| if the requested range partially overlaps an
        existing merge with different boundaries — the caller is expected
        to ``split_cells`` that overlap first.

        Returns the |_Cell| at the merge origin (top-left of the merged
        block).
        """
        top, bottom = _normalize_range(row_range)
        left, right = _normalize_range(col_range)

        origin_tc = self._tbl.tc(top, left)
        bottom_right_tc = self._tbl.tc(bottom, right)

        # ---single-cell range (no merge needed); return cell as-is---
        if top == bottom and left == right:
            return _Cell(origin_tc, self)

        target_row_count = bottom - top + 1
        target_col_count = right - left + 1

        # ---idempotency check: already merged exactly this way?---
        if (
            origin_tc.is_merge_origin
            and origin_tc.rowSpan == target_row_count
            and origin_tc.gridSpan == target_col_count
        ):
            return _Cell(origin_tc, self)

        tc_range = TcRange(origin_tc, bottom_right_tc)
        if tc_range.contains_merged_cell:
            raise ValueError(
                "merge_cells range partially overlaps an existing merge; "
                "call split_cells on the overlap first"
            )

        tc_range.move_content_to_origin()

        for tc in tc_range.iter_top_row_tcs():
            tc.rowSpan = target_row_count
        for tc in tc_range.iter_left_col_tcs():
            tc.gridSpan = target_col_count
        for tc in tc_range.iter_except_left_col_tcs():
            tc.hMerge = True
        for tc in tc_range.iter_except_top_row_tcs():
            tc.vMerge = True

        return _Cell(origin_tc, self)

    def split_cells(self, row_range, col_range) -> None:
        """Split (un-merge) any merges fully contained in this range.

        ``row_range`` and ``col_range`` follow the same shape rules as
        :meth:`merge_cells` — tuples are inclusive, ``range`` objects are
        half-open.

        Idempotent: cells in the range that aren't part of a merge are
        skipped silently. The order within each range is irrelevant.

        Raises |ValueError| if a merge in the range extends *beyond* the
        range boundary — splitting it would orphan the rest of the merge,
        so the caller must widen the range to include the full merge or
        call this on the full merge directly.
        """
        top, bottom = _normalize_range(row_range)
        left, right = _normalize_range(col_range)

        # ---first pass: validate every merge that intersects the range
        # ---is FULLY contained (origin + extent inside [top..bottom, left..right])---
        for r in range(top, bottom + 1):
            for c in range(left, right + 1):
                tc = self._tbl.tc(r, c)
                if tc.is_merge_origin:
                    if r + tc.rowSpan - 1 > bottom or c + tc.gridSpan - 1 > right:
                        raise ValueError(
                            "merge at (%d, %d) extends outside split range; "
                            "widen the range or call split_cells on the full merge" % (r, c)
                        )
                elif tc.hMerge or tc.vMerge:
                    # ---spanned cell whose origin is OUTSIDE the range = boundary cross---
                    # ---walk back to origin to verify---
                    origin_r, origin_c = _find_merge_origin(self._tbl, r, c)
                    if origin_r < top or origin_c < left:
                        raise ValueError(
                            "merge containing (%d, %d) starts outside split range; "
                            "widen the range or call split_cells on the full merge" % (r, c)
                        )

        # ---second pass: split each merge-origin in range (idempotent on non-merges)---
        for r in range(top, bottom + 1):
            for c in range(left, right + 1):
                tc = self._tbl.tc(r, c)
                if not tc.is_merge_origin:
                    continue
                tc_range = TcRange.from_merge_origin(tc)
                for inner_tc in tc_range.iter_tcs():
                    inner_tc.rowSpan = 1
                    inner_tc.gridSpan = 1
                    inner_tc.hMerge = False
                    inner_tc.vMerge = False

    @property
    def style_id(self) -> str | None:
        """The GUID identifying this table's built-in style, or |None|.

        Read/write. Maps to ``a:tbl/a:tblPr/a:tableStyleId``. PowerPoint
        emits the GUID in canonical brace-wrapped upper-case-hex shape, e.g.
        ``"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"`` for "Medium Style 2 -
        Accent 1". Setting to |None| removes the element.
        """
        tblPr = self._tbl.tblPr
        if tblPr is None:
            return None
        return tblPr.style_id

    @style_id.setter
    def style_id(self, value: str | None) -> None:
        if value is None:
            tblPr = self._tbl.tblPr
            if tblPr is None:
                return
            tblPr.style_id = None
            return
        tblPr = self._tbl.get_or_add_tblPr()
        tblPr.style_id = value

    @property
    def style_name(self) -> str | None:
        """Friendly name of the current built-in style, or |None|.

        Returns the friendly name (e.g. ``"Medium Style 2 - Accent 1"``)
        when the current ``style_id`` is in the built-in registry. Returns
        |None| when no style is set, or when the GUID is set but not
        recognized — the GUID is still readable via ``style_id`` in that
        case (lossless fallback).
        """
        from pptx.enum.table import style_name_for

        guid = self.style_id
        if guid is None:
            return None
        return style_name_for(guid)

    def apply_style(self, name_or_guid: str) -> None:
        """Set this table's style by friendly name or raw GUID.

        ``name_or_guid`` may be:

        - A built-in style name like ``"Medium Style 2 - Accent 1"`` —
          resolved against ``pptx.enum.table.PP_TABLE_STYLE``
          (case-insensitive). Raises |ValueError| if the name is not in the
          registry.
        - A GUID string in canonical brace-wrapped form (e.g.
          ``"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"``) — written through
          verbatim, allowing styles not yet in the registry (custom
          ``tableStyles.xml`` entries, additional Office built-ins) to be
          applied directly.

        For the GUID form the registry is not consulted — the value is
        treated as opaque and lossless. Use ``style_name`` to read back the
        friendly name when one is registered.
        """
        from pptx.enum.table import lookup_table_style

        if _looks_like_guid(name_or_guid):
            self.style_id = name_or_guid
            return
        # ---friendly name path: registry lookup, ValueError on miss---
        self.style_id = lookup_table_style(name_or_guid)


_GUID_RE = re.compile(
    r"^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$"
)


def _looks_like_guid(value: str) -> bool:
    """True when `value` matches the canonical brace-wrapped GUID shape."""
    return bool(_GUID_RE.match(value))


def _normalize_range(rng) -> tuple[int, int]:
    """Normalize a `merge_cells`/`split_cells` range argument to `(low, high)` inclusive.

    Accepts a 2-tuple (interpreted as inclusive `(start, end)`) or a Python
    `range` object (half-open per Python convention). Order within either
    form is irrelevant — `(2, 0)` becomes `(0, 2)`. Raises `TypeError` on
    other input shapes.
    """
    if isinstance(rng, range):
        # ---half-open: range(0, 2) covers 0..1 inclusive---
        if rng.step != 1:
            raise ValueError("range step must be 1, got %r" % rng.step)
        if len(rng) == 0:
            raise ValueError("range is empty: %r" % rng)
        low, high = rng.start, rng.stop - 1
    elif isinstance(rng, tuple) and len(rng) == 2:
        a, b = rng
        low, high = (a, b) if a <= b else (b, a)
    else:
        raise TypeError(
            "range argument must be a 2-tuple (inclusive) or a range object, got %r" % (rng,)
        )
    if low < 0 or high < 0:
        raise ValueError("range indices must be non-negative")
    return low, high


def _find_merge_origin(tbl, row_idx: int, col_idx: int) -> tuple[int, int]:
    """Walk back from a spanned cell to the (row, col) of its merge origin.

    A spanned cell carries `hMerge=True` and/or `vMerge=True` and its origin
    sits at some `(r0, c0)` where `r0 <= row_idx` and `c0 <= col_idx`. The
    origin's `rowSpan`/`gridSpan` covers (row_idx, col_idx). We scan
    leftward until `hMerge` is False, then upward until `vMerge` is False —
    that lands on the origin in two passes.
    """
    r, c = row_idx, col_idx
    # ---scan left through hMerge cells---
    while c > 0 and tbl.tc(r, c).hMerge:
        c -= 1
    # ---scan up through vMerge cells---
    while r > 0 and tbl.tc(r, c).vMerge:
        r -= 1
    return r, c


class _BorderEdge:
    """Adapter providing a `LineFormat`-compatible interface for one edge of a cell border.

    `LineFormat` requires a parent with `.ln` and `.get_or_add_ln()`. This adapter delegates those
    to the appropriate border element (`a:lnL`, `a:lnR`, `a:lnT`, `a:lnB`) on `a:tcPr`.
    """

    def __init__(self, tc: CT_TableCell, edge_attr: str):
        self._tc = tc
        self._edge_attr = edge_attr  # e.g. "lnL", "lnR", "lnT", "lnB"

    @property
    def ln(self):
        """Return the `a:lnX` element or None."""
        tcPr = self._tc.tcPr
        if tcPr is None:
            return None
        return getattr(tcPr, self._edge_attr)

    def get_or_add_ln(self):
        """Return the `a:lnX` element, creating `a:tcPr` and the element if not present."""
        tcPr = self._tc.get_or_add_tcPr()
        return getattr(tcPr, f"get_or_add_{self._edge_attr}")()


class _CellBorders:
    """Provides access to border line formatting for each edge of a table cell.

    Accessed via `cell.borders`. Each edge (`.left`, `.right`, `.top`, `.bottom`) returns a
    |LineFormat| object that controls the border's color, width, and dash style.
    """

    def __init__(self, tc: CT_TableCell):
        self._tc = tc

    @lazyproperty
    def bottom(self) -> LineFormat:
        """|LineFormat| for the bottom border of this cell."""
        return LineFormat(_BorderEdge(self._tc, "lnB"))

    @lazyproperty
    def left(self) -> LineFormat:
        """|LineFormat| for the left border of this cell."""
        return LineFormat(_BorderEdge(self._tc, "lnL"))

    @lazyproperty
    def right(self) -> LineFormat:
        """|LineFormat| for the right border of this cell."""
        return LineFormat(_BorderEdge(self._tc, "lnR"))

    @lazyproperty
    def top(self) -> LineFormat:
        """|LineFormat| for the top border of this cell."""
        return LineFormat(_BorderEdge(self._tc, "lnT"))


class _Cell(Subshape):
    """Table cell"""

    def __init__(self, tc: CT_TableCell, parent: ProvidesPart):
        super(_Cell, self).__init__(parent)
        self._tc = tc

    def __eq__(self, other: object) -> bool:
        """|True| if this object proxies the same element as `other`.

        Equality for proxy objects is defined as referring to the same XML element, whether or not
        they are the same proxy object instance.
        """
        if not isinstance(other, type(self)):
            return False
        return self._tc is other._tc

    def __ne__(self, other: object) -> bool:
        if not isinstance(other, type(self)):
            return True
        return self._tc is not other._tc

    @lazyproperty
    def borders(self) -> _CellBorders:
        """|_CellBorders| instance for this cell.

        Provides access to the line formatting for each border edge. Each edge (`.left`, `.right`,
        `.top`, `.bottom`) is a |LineFormat| object.

        Example::

            cell.borders.top.width = Pt(2)
            cell.borders.top.color.rgb = RGBColor(0xFF, 0x00, 0x00)
            cell.borders.bottom.dash_style = MSO_LINE.DASH
        """
        return _CellBorders(self._tc)

    @lazyproperty
    def fill(self) -> FillFormat:
        """|FillFormat| instance for this cell.

        Provides access to fill properties such as foreground color.
        """
        tcPr = self._tc.get_or_add_tcPr()
        return FillFormat.from_fill_parent(tcPr)

    @property
    def grid_span(self) -> int:
        """Number of grid columns this cell spans (1 if not a horizontal merge origin).

        Read-only. Mirrors the underlying ``a:tc/@gridSpan`` attribute. A
        merge-origin cell that spans ``N`` columns reports ``N``; a spanned
        (non-origin) cell reports 1 even when it is part of a merge — the
        merge origin holds the dimension; spanned cells carry ``h_merge`` /
        ``v_merge`` instead. Use this together with ``row_span`` /
        ``h_merge`` / ``v_merge`` to inspect any cell's merge state without
        relying on `is_merge_origin` heuristics.
        """
        return self._tc.gridSpan

    @property
    def row_span(self) -> int:
        """Number of grid rows this cell spans (1 if not a vertical merge origin).

        Read-only. Mirrors the underlying ``a:tc/@rowSpan`` attribute. Same
        contract as ``grid_span`` but for rows. See ``grid_span`` docstring.
        """
        return self._tc.rowSpan

    @property
    def h_merge(self) -> bool:
        """True if this cell is part of a horizontal merge but is NOT the origin.

        Read-only. Mirrors the underlying ``a:tc/@hMerge`` attribute.
        Always |False| on the merge-origin cell of a horizontal merge —
        only the spanned cells (those to the right of the origin) carry
        ``hMerge=True`` in the underlying XML.
        """
        return self._tc.hMerge

    @property
    def v_merge(self) -> bool:
        """True if this cell is part of a vertical merge but is NOT the origin.

        Read-only. Mirrors the underlying ``a:tc/@vMerge`` attribute.
        Always |False| on the merge-origin cell of a vertical merge —
        only the spanned cells (those below the origin) carry
        ``vMerge=True`` in the underlying XML.
        """
        return self._tc.vMerge

    @property
    def is_merge_origin(self) -> bool:
        """True if this cell is the top-left grid cell in a merged cell."""
        return self._tc.is_merge_origin

    @property
    def is_spanned(self) -> bool:
        """True if this cell is spanned by a merge-origin cell.

        A merge-origin cell "spans" the other grid cells in its merge range, consuming their area
        and "shadowing" the spanned grid cells.

        Note this value is |False| for a merge-origin cell. A merge-origin cell spans other grid
        cells, but is not itself a spanned cell.
        """
        return self._tc.is_spanned

    @property
    def margin_left(self) -> Length:
        """Left margin of cells.

        Read/write. If assigned |None|, the default value is used, 0.1 inches for left and right
        margins and 0.05 inches for top and bottom.
        """
        return self._tc.marL

    @margin_left.setter
    def margin_left(self, margin_left: Length | None):
        self._validate_margin_value(margin_left)
        self._tc.marL = margin_left

    @property
    def margin_right(self) -> Length:
        """Right margin of cell."""
        return self._tc.marR

    @margin_right.setter
    def margin_right(self, margin_right: Length | None):
        self._validate_margin_value(margin_right)
        self._tc.marR = margin_right

    @property
    def margin_top(self) -> Length:
        """Top margin of cell."""
        return self._tc.marT

    @margin_top.setter
    def margin_top(self, margin_top: Length | None):
        self._validate_margin_value(margin_top)
        self._tc.marT = margin_top

    @property
    def margin_bottom(self) -> Length:
        """Bottom margin of cell."""
        return self._tc.marB

    @margin_bottom.setter
    def margin_bottom(self, margin_bottom: Length | None):
        self._validate_margin_value(margin_bottom)
        self._tc.marB = margin_bottom

    def merge(self, other_cell: _Cell) -> None:
        """Create merged cell from this cell to `other_cell`.

        This cell and `other_cell` specify opposite corners of the merged cell range. Either
        diagonal of the cell region may be specified in either order, e.g. self=bottom-right,
        other_cell=top-left, etc.

        Raises |ValueError| if the specified range already contains merged cells anywhere within
        its extents or if `other_cell` is not in the same table as `self`.
        """
        tc_range = TcRange(self._tc, other_cell._tc)

        if not tc_range.in_same_table:
            raise ValueError("other_cell from different table")
        if tc_range.contains_merged_cell:
            raise ValueError("range contains one or more merged cells")

        tc_range.move_content_to_origin()

        row_count, col_count = tc_range.dimensions

        for tc in tc_range.iter_top_row_tcs():
            tc.rowSpan = row_count
        for tc in tc_range.iter_left_col_tcs():
            tc.gridSpan = col_count
        for tc in tc_range.iter_except_left_col_tcs():
            tc.hMerge = True
        for tc in tc_range.iter_except_top_row_tcs():
            tc.vMerge = True

    @property
    def span_height(self) -> int:
        """int count of rows spanned by this cell.

        The value of this property may be misleading (often 1) on cells where `.is_merge_origin`
        is not |True|, since only a merge-origin cell contains complete span information. This
        property is only intended for use on cells known to be a merge origin by testing
        `.is_merge_origin`.
        """
        return self._tc.rowSpan

    @property
    def span_width(self) -> int:
        """int count of columns spanned by this cell.

        The value of this property may be misleading (often 1) on cells where `.is_merge_origin`
        is not |True|, since only a merge-origin cell contains complete span information. This
        property is only intended for use on cells known to be a merge origin by testing
        `.is_merge_origin`.
        """
        return self._tc.gridSpan

    def split(self) -> None:
        """Remove merge from this (merge-origin) cell.

        The merged cell represented by this object will be "unmerged", yielding a separate
        unmerged cell for each grid cell previously spanned by this merge.

        Raises |ValueError| when this cell is not a merge-origin cell. Test with
        `.is_merge_origin` before calling.
        """
        if not self.is_merge_origin:
            raise ValueError("not a merge-origin cell; only a merge-origin cell can be split")

        tc_range = TcRange.from_merge_origin(self._tc)

        for tc in tc_range.iter_tcs():
            tc.rowSpan = tc.gridSpan = 1
            tc.hMerge = tc.vMerge = False

    @property
    def text(self) -> str:
        """Textual content of cell as a single string.

        The returned string will contain a newline character (`"\\n"`) separating each paragraph
        and a vertical-tab (`"\\v"`) character for each line break (soft carriage return) in the
        cell's text.

        Assignment to `text` replaces all text currently contained in the cell. A newline
        character (`"\\n"`) in the assigned text causes a new paragraph to be started. A
        vertical-tab (`"\\v"`) character in the assigned text causes a line-break (soft
        carriage-return) to be inserted. (The vertical-tab character appears in clipboard text
        copied from PowerPoint as its encoding of line-breaks.)
        """
        return self.text_frame.text

    @text.setter
    def text(self, text: str):
        self.text_frame.text = text

    @property
    def text_frame(self) -> TextFrame:
        """|TextFrame| containing the text that appears in the cell."""
        txBody = self._tc.get_or_add_txBody()
        return TextFrame(txBody, self)

    @property
    def vertical_anchor(self) -> MSO_VERTICAL_ANCHOR | None:
        """Vertical alignment of this cell.

        This value is a member of the :ref:`MsoVerticalAnchor` enumeration or |None|. A value of
        |None| indicates the cell has no explicitly applied vertical anchor setting and its
        effective value is inherited from its style-hierarchy ancestors.

        Assigning |None| to this property causes any explicitly applied vertical anchor setting to
        be cleared and inheritance of its effective value to be restored.
        """
        return self._tc.anchor

    @vertical_anchor.setter
    def vertical_anchor(self, mso_anchor_idx: MSO_VERTICAL_ANCHOR | None):
        self._tc.anchor = mso_anchor_idx

    @staticmethod
    def _validate_margin_value(margin_value: Length | None) -> None:
        """Raise ValueError if `margin_value` is not a positive integer value or |None|."""
        if not isinstance(margin_value, int) and margin_value is not None:
            tmpl = "margin value must be integer or None, got '%s'"
            raise TypeError(tmpl % margin_value)


class _Column(Subshape):
    """Table column"""

    def __init__(self, gridCol: CT_TableCol, parent: _ColumnCollection):
        super(_Column, self).__init__(parent)
        self._parent = parent
        self._gridCol = gridCol

    @property
    def width(self) -> Length:
        """Width of column in EMU."""
        return self._gridCol.w

    @width.setter
    def width(self, width: Length):
        self._gridCol.w = width
        self._parent.notify_width_changed()


class _Row(Subshape):
    """Table row"""

    def __init__(self, tr: CT_TableRow, parent: _RowCollection):
        super(_Row, self).__init__(parent)
        self._parent = parent
        self._tr = tr

    @property
    def cells(self):
        """Read-only reference to collection of cells in row.

        An individual cell is referenced using list notation, e.g. `cell = row.cells[0]`.
        """
        return _CellCollection(self._tr, self)

    @property
    def height(self) -> Length:
        """Height of row in EMU."""
        return self._tr.h

    @height.setter
    def height(self, height: Length):
        self._tr.h = height
        self._parent.notify_height_changed()


class _CellCollection(Subshape):
    """Horizontal sequence of row cells"""

    def __init__(self, tr: CT_TableRow, parent: _Row):
        super(_CellCollection, self).__init__(parent)
        self._parent = parent
        self._tr = tr

    def __getitem__(self, idx: int) -> _Cell:
        """Provides indexed access, (e.g. 'cells[0]')."""
        if idx < 0 or idx >= len(self._tr.tc_lst):
            msg = "cell index [%d] out of range" % idx
            raise IndexError(msg)
        return _Cell(self._tr.tc_lst[idx], self)

    def __iter__(self) -> Iterator[_Cell]:
        """Provides iterability."""
        return (_Cell(tc, self) for tc in self._tr.tc_lst)

    def __len__(self) -> int:
        """Supports len() function (e.g. 'len(cells) == 1')."""
        return len(self._tr.tc_lst)


class _ColumnCollection(Subshape):
    """Sequence of table columns."""

    def __init__(self, tbl: CT_Table, parent: Table):
        super(_ColumnCollection, self).__init__(parent)
        self._parent = parent
        self._tbl = tbl

    def __getitem__(self, idx: int):
        """Provides indexed access, (e.g. 'columns[0]')."""
        if idx < 0 or idx >= len(self._tbl.tblGrid.gridCol_lst):
            msg = "column index [%d] out of range" % idx
            raise IndexError(msg)
        return _Column(self._tbl.tblGrid.gridCol_lst[idx], self)

    def __iter__(self) -> Iterator[_Column]:
        """Generate each |_Column| in left-to-right order."""
        return (_Column(gc, self) for gc in self._tbl.tblGrid.gridCol_lst)

    def __len__(self):
        """Supports len() function (e.g. 'len(columns) == 1')."""
        return len(self._tbl.tblGrid.gridCol_lst)

    def add(self, at: int | None = None, width: Length | None = None) -> _Column:
        """Insert a new column and return its |_Column| proxy.

        When `at` is |None| (default), the new column is appended to the
        right edge. When `at` is an integer, the new column is inserted
        at that zero-based position; `at` may equal ``len(self)`` to
        append explicitly. Negative indices are not supported.

        When `width` is |None| (default), the new column inherits the
        width of the leftmost column (or 1 inch for an empty table); pass
        an explicit |Length| (e.g. ``Inches(1.5)``) to override.

        An empty `<a:tc>` is added at the corresponding position in every
        existing row, preserving row-cell alignment. Cell content in
        existing columns is left untouched.

        Raises |IndexError| if `at` is out of range and |ValueError| if
        the insertion column would split a cross-column merge.
        """
        if at is not None and (at < 0 or at > len(self)):
            raise IndexError("column index out of range")
        if width is None:
            existing = self._tbl.tblGrid.gridCol_lst
            width = Emu(existing[0].w) if existing else Emu(914400)  # ---1 inch
        idx = at if at is not None else len(self)
        # ---inserting at idx must not split an existing horizontal merge that
        #    spans across the boundary at column `idx`---
        if 0 < idx < len(self):
            for tr in self._tbl.tr_lst:
                if idx < len(tr.tc_lst) and tr.tc_lst[idx].hMerge:
                    raise ValueError(
                        "cannot insert column at index %d — would split a "
                        "horizontal merge; split the affected merge first" % idx
                    )
        new_gridCol = self._tbl.tblGrid.insert_gridCol_at(idx, width)
        for tr in self._tbl.tr_lst:
            tr.insert_tc_at(idx)
        self._parent.notify_width_changed()
        return _Column(new_gridCol, self)

    def remove(self, index: int) -> None:
        """Remove the column at `index`.

        Raises |IndexError| if `index` is out of range and |ValueError|
        if the column participates in a multi-column merge (the column
        contains a cell with `gridSpan > 1` or `hMerge=True`). Split the
        affected merge before calling.
        """
        if index < 0 or index >= len(self):
            raise IndexError("column index out of range")
        if self._tbl.column_has_cross_column_merge(index):
            raise ValueError(
                "cannot remove column %d containing a cross-column merge; "
                "split affected merges before removing the column" % index
            )
        self._tbl.tblGrid.remove_gridCol_at(index)
        for tr in self._tbl.tr_lst:
            tr.remove_tc_at(index)
        self._parent.notify_width_changed()

    def notify_width_changed(self):
        """Called by a column when its width changes. Pass along to parent."""
        self._parent.notify_width_changed()


class _RowCollection(Subshape):
    """Sequence of table rows"""

    def __init__(self, tbl: CT_Table, parent: Table):
        super(_RowCollection, self).__init__(parent)
        self._parent = parent
        self._tbl = tbl

    def __getitem__(self, idx: int) -> _Row:
        """Provides indexed access, (e.g. 'rows[0]')."""
        if idx < 0 or idx >= len(self):
            msg = "row index [%d] out of range" % idx
            raise IndexError(msg)
        return _Row(self._tbl.tr_lst[idx], self)

    def __iter__(self) -> Iterator[_Row]:
        """Generate each |_Row| in top-to-bottom order."""
        return (_Row(tr, self) for tr in self._tbl.tr_lst)

    def __len__(self):
        """Supports len() function (e.g. 'len(rows) == 1')."""
        return len(self._tbl.tr_lst)

    def add(self, at: int | None = None, height: Length | None = None) -> _Row:
        """Insert a new row and return its |_Row| proxy.

        When `at` is |None| (default), the new row is appended at the
        bottom. When `at` is an integer, the new row is inserted at that
        zero-based position; `at` may equal ``len(self)`` to append
        explicitly. Negative indices are not supported.

        When `height` is |None| (default), the new row inherits the
        height of the first row (or 0.4 inch for an empty table); pass an
        explicit |Length| (e.g. ``Inches(0.5)``) to override.

        The new row is populated with empty `<a:tc>` cells matching the
        table's current column count. Existing cell content is untouched.

        Raises |IndexError| if `at` is out of range and |ValueError| if
        the insertion row would split a cross-row merge.
        """
        if at is not None and (at < 0 or at > len(self)):
            raise IndexError("row index out of range")
        if height is None:
            existing = self._tbl.tr_lst
            height = Emu(existing[0].h) if existing else Emu(370840)  # ---~0.4 inch
        idx = at if at is not None else len(self)
        # ---inserting at idx must not split an existing vertical merge that
        #    spans across the boundary at row `idx`---
        if 0 < idx < len(self):
            for tc in self._tbl.tr_lst[idx].tc_lst:
                if tc.vMerge:
                    raise ValueError(
                        "cannot insert row at index %d — would split a "
                        "vertical merge; split the affected merge first" % idx
                    )
        new_tr = self._tbl.insert_tr_at(idx, height) if at is not None else self._tbl.add_tr(height)
        col_count = len(self._tbl.tblGrid.gridCol_lst)
        for _ in range(col_count):
            new_tr.add_tc()
        self._parent.notify_height_changed()
        return _Row(new_tr, self)

    def remove(self, index: int) -> None:
        """Remove the row at `index`.

        Raises |IndexError| if `index` is out of range and |ValueError|
        if the row participates in a multi-row merge (the row contains a
        cell with `rowSpan > 1` or `vMerge=True`). Split the affected
        merge before calling.
        """
        if index < 0 or index >= len(self):
            raise IndexError("row index out of range")
        if self._tbl.tr_lst[index].has_cross_row_merge:
            raise ValueError(
                "cannot remove row %d containing a cross-row merge; split "
                "affected merges before removing the row" % index
            )
        self._tbl.remove_tr_at(index)
        self._parent.notify_height_changed()

    def notify_height_changed(self):
        """Called by a row when its height changes. Pass along to parent."""
        self._parent.notify_height_changed()
