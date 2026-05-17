"""Base shape-related objects such as BaseShape."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from lxml.etree import _Element  # pyright: ignore[reportPrivateUsage]

from pptx.action import ActionSetting
from pptx.dml.effect import ShadowFormat
from pptx.dml.threed import Scene3DFormat, Shape3DFormat
from pptx.oxml.ns import qn
from pptx.shared import ElementProxy
from pptx.util import Emu, lazyproperty

# ---bound to the lxml base method so `find_by_xpath(..., namespaces=ns)` can
# ---honor the caller's prefix map without going through the project's
# ---`BaseOxmlElement.xpath` override (which auto-applies the project nsmap
# ---and rejects `namespaces=` kwarg).
_LXML_XPATH = _Element.xpath

if TYPE_CHECKING:
    from collections.abc import Iterator

    from pptx.comments import Comment, Comments
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    from pptx.oxml.shapes import ShapeElement
    from pptx.oxml.shapes.shared import CT_Placeholder
    from pptx.parts.slide import BaseSlidePart
    from pptx.types import ProvidesPart
    from pptx.util import Length


class BaseShape(object):
    """Base class for shape objects.

    Subclasses include |Shape|, |Picture|, and |GraphicFrame|.
    """

    def __init__(self, shape_elm: ShapeElement, parent: ProvidesPart):
        super().__init__()
        self._element = shape_elm
        self._parent = parent

    def __eq__(self, other: object) -> bool:
        """|True| if this shape object proxies the same element as *other*.

        Equality for proxy objects is defined as referring to the same XML element, whether or not
        they are the same proxy object instance.
        """
        if not isinstance(other, BaseShape):
            return False
        return self._element is other._element

    def __ne__(self, other: object) -> bool:
        if not isinstance(other, BaseShape):
            return True
        return self._element is not other._element

    @lazyproperty
    def click_action(self) -> ActionSetting:
        """|ActionSetting| instance providing access to click behaviors.

        Click behaviors are hyperlink-like behaviors including jumping to a hyperlink (web page)
        or to another slide in the presentation. The click action is that defined on the overall
        shape, not a run of text within the shape. An |ActionSetting| object is always returned,
        even when no click behavior is defined on the shape.
        """
        cNvPr = self._element._nvXxPr.cNvPr  # pyright: ignore[reportPrivateUsage]
        return ActionSetting(cNvPr, self)

    @property
    def element(self) -> ShapeElement:
        """`lxml` element for this shape, e.g. a CT_Shape instance.

        Note that manipulating this element improperly can produce an invalid presentation file.
        Make sure you know what you're doing if you use this to change the underlying XML.
        """
        return self._element

    def find_by_xpath(self, xpath: str, namespaces: "dict[str, str] | None" = None) -> list:
        """Power-user XPath escape hatch over this shape's element subtree.

        Returns whatever ``lxml.etree._Element.xpath`` returns — typically a
        list of matching elements, or an empty list when the expression
        matches nothing. When ``namespaces`` is |None| (default), the
        project's standard namespace map is used so common prefixes
        (``a:``, ``p:``, ``r:``, ``xsi:``, ``adec:``, ``p14:``, etc.) work
        without explicit declaration. Pass a custom dict to override.

        Example::

            for t_elm in shape.find_by_xpath(".//a:t"):
                print(t_elm.text)
        """
        if namespaces is None:
            # ---project's BaseOxmlElement.xpath auto-applies the standard nsmap---
            return self._element.xpath(xpath)
        # ---custom nsmap: bypass the project wrapper (see _LXML_XPATH note above)---
        return _LXML_XPATH(self._element, xpath, namespaces=namespaces)

    @property
    def has_chart(self) -> bool:
        """|True| if this shape is a graphic frame containing a chart object.

        |False| otherwise. When |True|, the chart object can be accessed using the ``.chart``
        property.
        """
        # This implementation is unconditionally False, the True version is
        # on GraphicFrame subclass.
        return False

    @property
    def has_table(self) -> bool:
        """|True| if this shape is a graphic frame containing a table object.

        |False| otherwise. When |True|, the table object can be accessed using the ``.table``
        property.
        """
        # This implementation is unconditionally False, the True version is
        # on GraphicFrame subclass.
        return False

    @property
    def has_text_frame(self) -> bool:
        """|True| if this shape can contain text."""
        # overridden on Shape to return True. Only <p:sp> has text frame
        return False

    @property
    def height(self) -> Length:
        """Read/write. Integer distance between top and bottom extents of shape in EMUs."""
        return self._element.cy

    @height.setter
    def height(self, value: Length):
        self._element.cy = value

    @property
    def is_placeholder(self) -> bool:
        """True if this shape is a placeholder.

        A shape is a placeholder if it has a <p:ph> element.
        """
        return self._element.has_ph_elm

    @property
    def left(self) -> Length:
        """Integer distance of the left edge of this shape from the left edge of the slide.

        Read/write. Expressed in English Metric Units (EMU)
        """
        return self._element.x

    @left.setter
    def left(self, value: Length):
        self._element.x = value

    @property
    def name(self) -> str:
        """Name of this shape, e.g. 'Picture 7'."""
        return self._element.shape_name

    @name.setter
    def name(self, value: str):
        self._element._nvXxPr.cNvPr.name = value  # pyright: ignore[reportPrivateUsage]

    @property
    def alt_text(self) -> str | None:
        """Alternative text describing this shape, used by screen readers and accessibility tools.

        Read/write. Returns the value of the `descr` attribute on the shape's
        `<p:cNvPr>` element. None if the attribute is not present (the shape has no
        alt text). Assigning None removes the attribute. Assigning an empty string
        is a meaningful, distinct value — it preserves the attribute as `descr=""`,
        useful for callers who want to round-trip an explicit "no description"
        marker.

        See Microsoft Accessibility guidance: prefer `alt_text` for the description
        and `alt_title` for a short heading, when both are needed.
        """
        return self._element._nvXxPr.cNvPr.descr  # pyright: ignore[reportPrivateUsage]

    @alt_text.setter
    def alt_text(self, value: str | None):
        self._element._nvXxPr.cNvPr.descr = value  # pyright: ignore[reportPrivateUsage]

    @property
    def alt_title(self) -> str | None:
        """Short title (heading) for this shape's alternative text, used for accessibility.

        Read/write. Returns the value of the `title` attribute on the shape's
        `<p:cNvPr>` element. None if the attribute is not present. Assigning None
        removes the attribute. Microsoft accessibility guidance recommends a brief
        title plus a longer `alt_text` description, mirroring the two-field UX in
        PowerPoint's "Alt Text" pane.
        """
        return self._element._nvXxPr.cNvPr.title  # pyright: ignore[reportPrivateUsage]

    @alt_title.setter
    def alt_title(self, value: str | None):
        self._element._nvXxPr.cNvPr.title = value  # pyright: ignore[reportPrivateUsage]

    @property
    def is_decorative(self) -> bool:
        """True if this shape is marked as decorative (Office 2019+ accessibility flag).

        Read/write boolean. Decorative shapes are skipped by screen readers; they
        carry no semantic meaning beyond visual decoration (background grids,
        ornaments, dividers). Backed by an `<adec:decorative val="1"/>` extension
        inside `<p:cNvPr>/<a:extLst>`. Setting to False removes the extension; the
        attribute defaults to False on shapes that have never been touched.
        """
        return self._element._nvXxPr.cNvPr.decorative  # pyright: ignore[reportPrivateUsage]

    @is_decorative.setter
    def is_decorative(self, value: bool):
        self._element._nvXxPr.cNvPr.decorative = bool(value)  # pyright: ignore[reportPrivateUsage]

    @property
    def is_hidden_from_accessibility(self) -> bool:
        """Convenience alias for :attr:`is_decorative`.

        Read/write. Decorative shapes (the official OOXML term — `<adec:decorative
        val="1"/>`) are exactly those that are hidden from accessibility tools
        such as screen readers. Some accessibility documentation (and a number of
        third-party authoring tools) use the wording "hidden from accessibility"
        for the same flag; this property exists so the API reads naturally for
        either audience.
        """
        return self.is_decorative

    @is_hidden_from_accessibility.setter
    def is_hidden_from_accessibility(self, value: bool):
        self.is_decorative = bool(value)

    @property
    def part(self) -> BaseSlidePart:
        """The package part containing this shape.

        A |BaseSlidePart| subclass in this case. Access to a slide part should only be required if
        you are extending the behavior of |pp| API objects.
        """
        return cast("BaseSlidePart", self._parent.part)

    @property
    def placeholder_format(self) -> _PlaceholderFormat:
        """Provides access to placeholder-specific properties such as placeholder type.

        Raises |ValueError| on access if the shape is not a placeholder.
        """
        ph = self._element.ph
        if ph is None:
            raise ValueError("shape is not a placeholder")
        return _PlaceholderFormat(ph)

    @property
    def rotation(self) -> float:
        """Degrees of clockwise rotation.

        Read/write float. Negative values can be assigned to indicate counter-clockwise rotation,
        e.g. assigning -45.0 will change setting to 315.0.
        """
        return self._element.rot

    @rotation.setter
    def rotation(self, value: float):
        self._element.rot = value

    @lazyproperty
    def shadow(self) -> ShadowFormat:
        """|ShadowFormat| object providing access to shadow for this shape.

        A |ShadowFormat| object is always returned, even when no shadow is
        explicitly defined on this shape (i.e. it inherits its shadow
        behavior).
        """
        return ShadowFormat(self._element.spPr)

    @property
    def shape_id(self) -> int:
        """Read-only positive integer identifying this shape.

        The id of a shape is unique among all shapes on a slide.
        """
        return self._element.shape_id

    @property
    def flip_horizontal(self) -> bool:
        """Read/write. |True| if this shape is mirrored left-to-right.

        Backed by the `flipH` attribute of the shape's `a:xfrm`. Assigning a
        value creates the `a:xfrm` element if necessary (issue #18 SF8).
        """
        return bool(self._element.flipH)

    @flip_horizontal.setter
    def flip_horizontal(self, value: bool) -> None:
        self._element.flipH = bool(value)

    @property
    def flip_vertical(self) -> bool:
        """Read/write. |True| if this shape is mirrored top-to-bottom.

        Backed by the `flipV` attribute of the shape's `a:xfrm` (issue #18
        SF8). `shape.flip_vertical = True` round-trips through PowerPoint.
        """
        return bool(self._element.flipV)

    @flip_vertical.setter
    def flip_vertical(self, value: bool) -> None:
        self._element.flipV = bool(value)

    @lazyproperty
    def scene_3d(self) -> Scene3DFormat:
        """|Scene3DFormat| object providing access to this shape's 3-D scene.

        Lets a preset camera be applied (`a:scene3d`). A |Scene3DFormat| is
        always returned; the `a:scene3d` element (with its schema-mandatory
        camera + light-rig children) is created only when a camera preset is
        assigned (issue #18 SF4).
        """
        return Scene3DFormat(self._element.spPr)

    @lazyproperty
    def shape_3d(self) -> Shape3DFormat:
        """|Shape3DFormat| object providing access to this shape's 3-D format.

        Lets extrusion / contour be applied (`a:sp3d`). Always returned; the
        `a:sp3d` element is created only when a 3-D property is assigned
        (issue #18 SF4).
        """
        return Shape3DFormat(self._element.spPr)

    @property
    def slide_left(self) -> Length:
        """World-space left edge of this shape in slide EMU.

        For a shape that is **not** inside a group this equals :attr:`left`.
        For a shape inside one or more groups, the enclosing group
        transforms (`a:off`/`a:ext` vs `a:chOff`/`a:chExt`) are composed
        outward so the returned value is the true position on the slide
        (issue #18 SF7). Read-only; this never mutates the stored `a:xfrm`.
        """
        return self._world_rect()[0]

    @property
    def slide_top(self) -> Length:
        """World-space top edge of this shape in slide EMU (see :attr:`slide_left`)."""
        return self._world_rect()[1]

    @property
    def slide_width(self) -> Length:
        """World-space width of this shape in slide EMU (see :attr:`slide_left`)."""
        return self._world_rect()[2]

    @property
    def slide_height(self) -> Length:
        """World-space height of this shape in slide EMU (see :attr:`slide_left`)."""
        return self._world_rect()[3]

    def _world_rect(self) -> tuple[Length, Length, Length, Length]:
        """Return ``(left, top, width, height)`` of this shape in slide EMU.

        Composes every enclosing ``p:grpSp`` group transform outward. The
        same affine handles arbitrary nesting depth — each group's
        ``a:off``/``a:ext`` are expressed in its own parent's child space, so
        re-applying the next ancestor's transform composes correctly with no
        nested-vs-single special-casing. A degenerate group (``a:chExt`` of
        zero on either axis) falls back to an identity scale rather than
        dividing by zero.

        Scope note: only the scale + translate of each group is composed.
        Group **rotation** and **flipH/flipV** are intentionally not folded
        in — the result is the axis-aligned child-orientation box (matching
        PowerPoint's COM ``Shape.Left`` semantics for grouped shapes). A
        rotated/flipped enclosing group will therefore give the unrotated
        rect; that is by design, not a bug.
        """

        def _f(elm, path: str, default: int = 0) -> int:
            vals = elm.xpath(path)
            return int(vals[0]) if vals else default

        x = float(self._element.x or 0)
        y = float(self._element.y or 0)
        cx = float(self._element.cx or 0)
        cy = float(self._element.cy or 0)

        parent = self._element.getparent()
        while parent is not None and parent.tag == qn("p:grpSp"):
            base = "./p:grpSpPr/a:xfrm"
            gx = _f(parent, f"{base}/a:off/@x")
            gy = _f(parent, f"{base}/a:off/@y")
            gcx = _f(parent, f"{base}/a:ext/@cx")
            gcy = _f(parent, f"{base}/a:ext/@cy")
            chx = _f(parent, f"{base}/a:chOff/@x")
            chy = _f(parent, f"{base}/a:chOff/@y")
            chcx = _f(parent, f"{base}/a:chExt/@cx")
            chcy = _f(parent, f"{base}/a:chExt/@cy")
            sx = (gcx / chcx) if chcx else 1.0
            sy = (gcy / chcy) if chcy else 1.0
            x = gx + (x - chx) * sx
            y = gy + (y - chy) * sy
            cx = cx * sx
            cy = cy * sy
            parent = parent.getparent()

        return (Emu(int(x)), Emu(int(y)), Emu(int(cx)), Emu(int(cy)))

    def duplicate(self, insert_at_z: int | None = None) -> "BaseShape":
        """Return a deep-copy of this shape added to the same shape tree.

        The clone gets a fresh, unique shape id and a unique name; its XML is
        an independent deep copy (mutating the clone does not affect the
        original). With `insert_at_z` |None| (default) the clone is appended
        at the top of the z-order; otherwise it is inserted at z-order index
        `insert_at_z` (issue #18 SF9).

        Limitation: this is a pure XML deep-copy. For a relationship-backed
        shape (picture, chart, table, OLE object) the `r:embed`/`r:id`
        reference is copied but the target part is **not** cloned — both
        shapes then share one image/chart part. That is fine for read-back
        and for autoshapes/connectors/text boxes (no relationships), but a
        true picture/chart duplicate that needs an independent part is out
        of scope here.
        """
        return self._parent._duplicate_shape(  # pyright: ignore[reportAttributeAccessIssue]
            self, insert_at_z
        )

    @property
    def comments(self) -> "_ShapeComments":
        """The comments anchored to *this* shape (issue #25 Wave 3, SF7).

        A filtered, read-only view over the owning slide's
        ``slide.comments`` that yields only the modern threaded comments
        whose anchor shape id matches this shape's :attr:`shape_id`.
        Iterable and ``len()``-able; an empty (not error) collection when
        the shape has no comments (ISC-39). A comment anchored to another
        shape never appears here (ISC-40) — anchoring is by stable shape
        id, so the filter survives a save→reopen. Legacy ``<p:cm>``
        comments anchor to an absolute point rather than a shape and so
        never participate in a per-shape filter.
        """
        slide = self.part.slide  # pyright: ignore[reportAttributeAccessIssue]
        return _ShapeComments(slide.comments, self.shape_id)

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        """A member of MSO_SHAPE_TYPE classifying this shape by type.

        Like ``MSO_SHAPE_TYPE.CHART``. Must be implemented by subclasses.
        """
        raise NotImplementedError(f"{type(self).__name__} does not implement `.shape_type`")

    @property
    def top(self) -> Length:
        """Distance from the top edge of the slide to the top edge of this shape.

        Read/write. Expressed in English Metric Units (EMU)
        """
        return self._element.y

    @top.setter
    def top(self, value: Length):
        self._element.y = value

    @property
    def width(self) -> Length:
        """Distance between left and right extents of this shape.

        Read/write. Expressed in English Metric Units (EMU).
        """
        return self._element.cx

    @width.setter
    def width(self, value: Length):
        self._element.cx = value


class _PlaceholderFormat(ElementProxy):
    """Provides properties specific to placeholders, such as the placeholder type.

    Accessed via the :attr:`~.BaseShape.placeholder_format` property of a placeholder shape,
    """

    def __init__(self, element: CT_Placeholder):
        super().__init__(element)
        self._ph = element

    @property
    def element(self) -> CT_Placeholder:
        """The `p:ph` element proxied by this object."""
        return self._ph

    @property
    def idx(self) -> int:
        """Integer placeholder 'idx' attribute."""
        return self._ph.idx

    @property
    def type(self) -> PP_PLACEHOLDER:
        """Placeholder type.

        A member of the :ref:`PpPlaceholderType` enumeration, e.g. PP_PLACEHOLDER.CHART
        """
        return self._ph.type


class _ShapeComments:
    """A read-only, per-shape filtered view over a slide's comments.

    Backs :attr:`BaseShape.comments` (issue #25 Wave 3, SF7). Wraps the
    slide's |Comments| collection and yields only the comments whose anchor
    resolves to one specific shape id. Iterable and ``len()``-able; indexing
    is supported for convenience. Never mutates the package — anchoring is
    established at ``slide.comments.add(..., anchor=shape)`` time.
    """

    def __init__(self, comments: "Comments", shape_id: int):
        self._comments = comments
        self._shape_id = shape_id

    def __iter__(self) -> "Iterator[Comment]":
        for comment in self._comments:
            if comment._anchor_shape_id == self._shape_id:  # noqa: SLF001
                yield comment

    def __len__(self) -> int:
        return sum(1 for _ in self)

    def __getitem__(self, idx: int) -> "Comment":
        return list(self)[idx]
