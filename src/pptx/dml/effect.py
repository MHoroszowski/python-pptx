"""Visual effects on a shape such as shadow, glow, and reflection."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.dml.color import ColorFormat
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.dml.effect import (
        CT_GlowEffect,
        CT_OuterShadowEffect,
        CT_ReflectionEffect,
        CT_SoftEdgesEffect,
    )
    from pptx.util import Length


class ShadowFormat(object):
    """Provides access to shadow effect on a shape."""

    def __init__(self, spPr):
        # ---spPr may also be a grpSpPr; both have a:effectLst child---
        self._element = spPr

    @property
    def angle(self) -> float | None:
        """Direction of shadow in degrees (0 = right, 90 = below, etc.).

        Read/write. Returns |None| if no shadow is explicitly defined. Setting this property
        creates an outer shadow if one doesn't exist.
        """
        outerShdw = self._outerShdw
        if outerShdw is None:
            return None
        return outerShdw.dir

    @angle.setter
    def angle(self, value: float | None) -> None:
        if value is None:
            return
        outerShdw = self._get_or_add_outerShdw()
        outerShdw.dir = value

    @property
    def blur_radius(self) -> Length | None:
        """Blur radius of shadow in EMU.

        Read/write. Returns |None| if no shadow is explicitly defined.
        """
        outerShdw = self._outerShdw
        if outerShdw is None:
            return None
        return Emu(outerShdw.blurRad)

    @blur_radius.setter
    def blur_radius(self, value: Length | None) -> None:
        if value is None:
            return
        outerShdw = self._get_or_add_outerShdw()
        outerShdw.blurRad = int(value)

    @lazyproperty
    def color(self) -> ColorFormat:
        """Color of the shadow.

        Returns a |ColorFormat| object. Setting color properties creates an outer shadow with a
        solid color fill if one doesn't exist.
        """
        outerShdw = self._get_or_add_outerShdw()
        return ColorFormat.from_colorchoice_parent(outerShdw)

    @property
    def distance(self) -> Length | None:
        """Distance of shadow from shape in EMU.

        Read/write. Returns |None| if no shadow is explicitly defined.
        """
        outerShdw = self._outerShdw
        if outerShdw is None:
            return None
        return Emu(outerShdw.dist)

    @distance.setter
    def distance(self, value: Length | None) -> None:
        if value is None:
            return
        outerShdw = self._get_or_add_outerShdw()
        outerShdw.dist = int(value)

    @property
    def inherit(self):
        """True if shape inherits shadow settings.

        Read/write. An explicitly-defined shadow setting on a shape causes
        this property to return |False|. A shape with no explicitly-defined
        shadow setting inherits its shadow settings from the style hierarchy
        (and so returns |True|).

        Assigning |True| causes any explicitly-defined shadow setting to be
        removed and inheritance is restored. Note this has the side-effect of
        removing **all** explicitly-defined effects, such as glow and
        reflection, and restoring inheritance for all effects on the shape.
        Assigning |False| causes the inheritance link to be broken and **no**
        effects to appear on the shape.
        """
        return self._element.effectLst is None

    @inherit.setter
    def inherit(self, value):
        inherit = bool(value)
        if inherit:
            # ---remove any explicitly-defined effects
            self._element._remove_effectLst()
        else:
            # ---ensure at least the effectLst element is present
            self._element.get_or_add_effectLst()

    @property
    def rotate_with_shape(self) -> bool | None:
        """Whether the shadow rotates with the shape.

        Read/write. Returns |None| if no shadow is explicitly defined.
        """
        outerShdw = self._outerShdw
        if outerShdw is None:
            return None
        return outerShdw.rotWithShape

    @rotate_with_shape.setter
    def rotate_with_shape(self, value: bool | None) -> None:
        if value is None:
            return
        outerShdw = self._get_or_add_outerShdw()
        outerShdw.rotWithShape = value

    @property
    def visible(self) -> bool:
        """Whether a shadow is visible on this shape.

        Read/write. Returns |True| if an outer shadow element is present. Assigning |True| creates
        a default outer shadow. Assigning |False| removes any outer shadow.
        """
        return self._outerShdw is not None

    @visible.setter
    def visible(self, value: bool) -> None:
        if value:
            self._get_or_add_outerShdw()
        else:
            effectLst = self._element.effectLst
            if effectLst is not None:
                effectLst._remove_outerShdw()

    def _get_or_add_outerShdw(self) -> CT_OuterShadowEffect:
        """Return the `a:outerShdw` element, creating parent elements as needed."""
        effectLst = self._element.get_or_add_effectLst()
        return effectLst.get_or_add_outerShdw()

    @property
    def _outerShdw(self) -> CT_OuterShadowEffect | None:
        """Return `a:outerShdw` element or None if not present."""
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        return effectLst.outerShdw

    @lazyproperty
    def glow_effect(self) -> GlowEffect:
        """|GlowEffect| object providing access to the shape's glow effect.

        A |GlowEffect| object is always returned, even when no glow is
        explicitly defined on this shape. Setting a glow property (color or
        radius) creates the `a:glow` element on demand, mirroring the way
        :attr:`color` creates an outer shadow. Separate from the
        already-shipped outer-shadow API (issue #18 SF1).
        """
        return GlowEffect(self._element)

    @lazyproperty
    def reflection_effect(self) -> ReflectionEffect:
        """|ReflectionEffect| object providing access to the reflection effect.

        Always returned; setting a property creates `a:reflection` on demand
        (issue #18 SF2).
        """
        return ReflectionEffect(self._element)

    @lazyproperty
    def soft_edge_effect(self) -> SoftEdgeEffect:
        """|SoftEdgeEffect| object providing access to the soft-edge effect.

        Always returned; setting the radius creates `a:softEdge` on demand
        (issue #18 SF3).
        """
        return SoftEdgeEffect(self._element)


class GlowEffect(object):
    """Provides access to the glow effect (`a:glow`) on a shape.

    Accessed via :attr:`ShadowFormat.glow_effect`. Mirrors the lazy-create
    semantics of |ShadowFormat|: the `a:glow` element (and its mandatory
    color child + enclosing `a:effectLst`) is created only when a property
    is assigned.
    """

    def __init__(self, spPr):
        # ---spPr may also be a grpSpPr; both have an a:effectLst child---
        self._element = spPr

    @property
    def visible(self) -> bool:
        """|True| if an `a:glow` element is present on this shape."""
        return self._glow is not None

    @lazyproperty
    def color(self) -> ColorFormat:
        """Color of the glow as a |ColorFormat|.

        Accessing (or setting) this creates an `a:glow` with a default color
        child if one doesn't exist, just as :attr:`ShadowFormat.color` does
        for the outer shadow.
        """
        return ColorFormat.from_colorchoice_parent(self._get_or_add_glow())

    @property
    def radius(self) -> Length | None:
        """Glow radius as a |Length|, or |None| if no glow is defined."""
        glow = self._glow
        if glow is None:
            return None
        return Emu(glow.rad)

    @radius.setter
    def radius(self, value: Length | None) -> None:
        if value is None:
            return
        self._get_or_add_glow().rad = int(value)

    def _get_or_add_glow(self) -> CT_GlowEffect:
        return self._element.get_or_add_effectLst().get_or_add_glow()

    @property
    def _glow(self) -> CT_GlowEffect | None:
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        return effectLst.glow


class ReflectionEffect(object):
    """Provides access to the reflection effect (`a:reflection`) on a shape.

    Accessed via :attr:`ShadowFormat.reflection_effect`.
    """

    def __init__(self, spPr):
        self._element = spPr

    @property
    def visible(self) -> bool:
        """|True| if an `a:reflection` element is present."""
        return self._reflection is not None

    @visible.setter
    def visible(self, value: bool) -> None:
        if value:
            self._get_or_add_reflection()
        else:
            effectLst = self._element.effectLst
            if effectLst is not None:
                effectLst._remove_reflection()

    @property
    def blur_radius(self) -> Length | None:
        """Reflection blur radius as a |Length|, or |None| if not defined."""
        reflection = self._reflection
        if reflection is None:
            return None
        return Emu(reflection.blurRad)

    @blur_radius.setter
    def blur_radius(self, value: Length | None) -> None:
        if value is None:
            return
        self._get_or_add_reflection().blurRad = int(value)

    @property
    def distance(self) -> Length | None:
        """Distance of the reflection from the shape, |None| if not defined."""
        reflection = self._reflection
        if reflection is None:
            return None
        return Emu(reflection.dist)

    @distance.setter
    def distance(self, value: Length | None) -> None:
        if value is None:
            return
        self._get_or_add_reflection().dist = int(value)

    @property
    def direction(self) -> float | None:
        """Direction of the reflection in degrees, |None| if not defined."""
        reflection = self._reflection
        if reflection is None:
            return None
        return reflection.dir

    @direction.setter
    def direction(self, value: float | None) -> None:
        if value is None:
            return
        self._get_or_add_reflection().dir = value

    def _get_or_add_reflection(self) -> CT_ReflectionEffect:
        return self._element.get_or_add_effectLst().get_or_add_reflection()

    @property
    def _reflection(self) -> CT_ReflectionEffect | None:
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        return effectLst.reflection


class SoftEdgeEffect(object):
    """Provides access to the soft-edge effect (`a:softEdge`) on a shape.

    Accessed via :attr:`ShadowFormat.soft_edge_effect`. The schema makes
    `rad` required, so creating a soft edge always sets a radius (a sensible
    2.5pt default until overwritten).
    """

    def __init__(self, spPr):
        self._element = spPr

    @property
    def visible(self) -> bool:
        """|True| if an `a:softEdge` element is present."""
        return self._softEdge is not None

    @visible.setter
    def visible(self, value: bool) -> None:
        if value:
            self._get_or_add_softEdge()
        else:
            effectLst = self._element.effectLst
            if effectLst is not None:
                effectLst._remove_softEdge()

    @property
    def radius(self) -> Length | None:
        """Soft-edge feather radius as a |Length|, |None| if not defined."""
        softEdge = self._softEdge
        if softEdge is None:
            return None
        return Emu(softEdge.rad)

    @radius.setter
    def radius(self, value: Length | None) -> None:
        if value is None:
            return
        self._get_or_add_softEdge().rad = int(value)

    def _get_or_add_softEdge(self) -> CT_SoftEdgesEffect:
        return self._element.get_or_add_effectLst().get_or_add_softEdge()

    @property
    def _softEdge(self) -> CT_SoftEdgesEffect | None:
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        return effectLst.softEdge
