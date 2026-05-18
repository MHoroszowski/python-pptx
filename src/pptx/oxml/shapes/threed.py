"""lxml custom element classes for DrawingML 3-D scene/shape XML elements.

Covers `a:scene3d` (preset camera + light rig) and `a:sp3d` (extrusion /
contour / bevel). DrawingML §20.1.4.2.x and §20.1.5.x; ECMA-376 dml-main.xsd
`CT_Scene3D`, `CT_Camera`, `CT_LightRig`, `CT_Shape3D`.

The single hardest no-repair fact here: `CT_Scene3D` requires **both** a
`camera` and a `lightRig` child (`minOccurs="1"` each). A `<a:scene3d>` with
only a camera makes PowerPoint show a repair dialog, so the default factory
always emits both.
"""

from __future__ import annotations

from typing import Callable, cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import ST_Coordinate, ST_PositiveCoordinate, XsdString
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrOne


class CT_Camera(BaseOxmlElement):
    """`a:camera` custom element class.

    `prst` (preset camera type, e.g. ``orthographicFront``,
    ``perspectiveRelaxedModerately``) is required by the schema.
    """

    prst = RequiredAttribute("prst", XsdString)
    fov = OptionalAttribute("fov", XsdString)
    zoom = OptionalAttribute("zoom", XsdString)


class CT_LightRig(BaseOxmlElement):
    """`a:lightRig` custom element class.

    Both `rig` (e.g. ``threePt``) and `dir` (e.g. ``t``) are required by the
    schema; a lightRig missing either is a repair trigger.
    """

    rig = RequiredAttribute("rig", XsdString)
    dir = RequiredAttribute("dir", XsdString)


class CT_Scene3D(BaseOxmlElement):
    """`a:scene3d` custom element class.

    Schema sequence: ``camera`` (required), ``lightRig`` (required),
    ``backdrop`` (optional), ``extLst`` (optional).
    """

    get_or_add_camera: Callable[[], CT_Camera]
    get_or_add_lightRig: Callable[[], CT_LightRig]

    _tag_seq = ("a:camera", "a:lightRig", "a:backdrop", "a:extLst")
    camera: CT_Camera | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:camera", successors=_tag_seq[1:]
    )
    lightRig: CT_LightRig | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:lightRig", successors=_tag_seq[2:]
    )
    del _tag_seq

    @classmethod
    def new_scene3d(cls, camera_prst: str = "orthographicFront") -> "CT_Scene3D":
        """Return a new `a:scene3d` with the required camera + lightRig children.

        PowerPoint requires both children present; the lightRig default
        (``threePt`` / ``t``) matches what PowerPoint's 3-D format gallery
        emits for a flat preset, so a round-tripped file opens clean.
        """
        return cast(
            CT_Scene3D,
            parse_xml(
                f"<a:scene3d {nsdecls('a')}>"
                f'  <a:camera prst="{camera_prst}"/>'
                f'  <a:lightRig rig="threePt" dir="t"/>'
                f"</a:scene3d>"
            ),
        )


class CT_Shape3D(BaseOxmlElement):
    """`a:sp3d` custom element class.

    Schema children (all optional): ``bevelT``, ``bevelB``, ``extrusionClr``,
    ``contourClr``, ``extLst``. Attributes: ``z``, ``extrusionH``,
    ``contourW``, ``prstMaterial``. An attribute-only `<a:sp3d>` is schema-
    valid and PowerPoint-accepted.
    """

    z = OptionalAttribute("z", ST_Coordinate, default=0)
    extrusionH = OptionalAttribute("extrusionH", ST_PositiveCoordinate, default=0)
    contourW = OptionalAttribute("contourW", ST_PositiveCoordinate, default=0)
    prstMaterial = OptionalAttribute("prstMaterial", XsdString)

    @classmethod
    def new_sp3d(cls) -> "CT_Shape3D":
        """Return a new bare `a:sp3d` element (all attributes default)."""
        return cast(CT_Shape3D, parse_xml(f"<a:sp3d {nsdecls('a')}/>"))
