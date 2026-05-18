"""DrawingML objects related to 3-D scene and shape formatting.

`Scene3DFormat` wraps `a:scene3d` (preset camera) and `Shape3DFormat` wraps
`a:sp3d` (extrusion / contour). Accessed via :attr:`BaseShape.scene_3d` and
:attr:`BaseShape.shape_3d`. Issue #18 SF4.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.util import Emu

if TYPE_CHECKING:
    from pptx.util import Length


class Scene3DFormat(object):
    """Provides access to the 3-D scene (`a:scene3d`) on a shape.

    A preset camera (e.g. ``"perspectiveRelaxedModerately"``) is the primary
    knob. Assigning :attr:`camera_preset` creates the `a:scene3d` element —
    together with its schema-mandatory `a:camera` and `a:lightRig` children
    — so the resulting file opens in PowerPoint without a repair dialog.
    """

    def __init__(self, spPr):
        self._spPr = spPr

    @property
    def visible(self) -> bool:
        """|True| if an `a:scene3d` element is present on this shape."""
        return self._spPr.scene3d is not None

    @property
    def camera_preset(self) -> str | None:
        """Preset camera type, e.g. ``"orthographicFront"``.

        |None| if no 3-D scene is defined. Assigning a value creates the
        scene (with camera + light rig) if necessary.
        """
        scene3d = self._spPr.scene3d
        if scene3d is None or scene3d.camera is None:
            return None
        return scene3d.camera.prst

    @camera_preset.setter
    def camera_preset(self, value: str | None) -> None:
        if value is None:
            return
        scene3d = self._spPr.get_or_add_scene3d()
        scene3d.get_or_add_camera().prst = value


class Shape3DFormat(object):
    """Provides access to the 3-D shape format (`a:sp3d`) on a shape.

    Extrusion height and contour width are the primary knobs. A bare
    `<a:sp3d>` is schema-valid, so assigning either property simply creates
    the element with that attribute set.
    """

    def __init__(self, spPr):
        self._spPr = spPr

    @property
    def visible(self) -> bool:
        """|True| if an `a:sp3d` element is present on this shape."""
        return self._spPr.sp3d is not None

    @property
    def extrusion_height(self) -> Length | None:
        """Extrusion (depth) height as a |Length|, |None| if not defined."""
        sp3d = self._spPr.sp3d
        if sp3d is None:
            return None
        return Emu(sp3d.extrusionH)

    @extrusion_height.setter
    def extrusion_height(self, value: Length | None) -> None:
        if value is None:
            return
        self._spPr.get_or_add_sp3d().extrusionH = int(value)

    @property
    def contour_width(self) -> Length | None:
        """Contour (edge) width as a |Length|, |None| if not defined."""
        sp3d = self._spPr.sp3d
        if sp3d is None:
            return None
        return Emu(sp3d.contourW)

    @contour_width.setter
    def contour_width(self, value: Length | None) -> None:
        if value is None:
            return
        self._spPr.get_or_add_sp3d().contourW = int(value)
