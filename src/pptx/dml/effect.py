"""Visual effects on a shape such as shadow, glow, and reflection."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.dml.color import ColorFormat
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.dml.effect import CT_OuterShadowEffect
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
        if self._element.effectLst is None:
            return True
        return False

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
