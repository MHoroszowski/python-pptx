"""lxml custom element classes for DrawingML effect-related XML elements."""

from __future__ import annotations

from typing import Callable, cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import ST_Angle, ST_PositiveCoordinate, XsdBoolean
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OptionalAttribute,
    ZeroOrOne,
    ZeroOrOneChoice,
)


class CT_OuterShadowEffect(BaseOxmlElement):
    """`a:outerShdw` custom element class."""

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    blurRad = OptionalAttribute("blurRad", ST_PositiveCoordinate, default=0)
    dist = OptionalAttribute("dist", ST_PositiveCoordinate, default=0)
    dir = OptionalAttribute("dir", ST_Angle, default=0)
    rotWithShape: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rotWithShape", XsdBoolean, default=True
    )


class CT_InnerShadowEffect(BaseOxmlElement):
    """`a:innerShdw` custom element class."""

    blurRad = OptionalAttribute("blurRad", ST_PositiveCoordinate, default=0)
    dist = OptionalAttribute("dist", ST_PositiveCoordinate, default=0)
    dir = OptionalAttribute("dir", ST_Angle, default=0)


class CT_EffectList(BaseOxmlElement):
    """`a:effectLst` custom element class."""

    get_or_add_outerShdw: Callable[[], CT_OuterShadowEffect]
    _remove_outerShdw: Callable[[], None]
    _remove_innerShdw: Callable[[], None]

    _tag_seq = ("a:blur", "a:fillOverlay", "a:glow", "a:innerShdw", "a:outerShdw",
                "a:prstShdw", "a:reflection", "a:softEdge")
    innerShdw: CT_InnerShadowEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:innerShdw", successors=_tag_seq[4:]
    )
    outerShdw: CT_OuterShadowEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:outerShdw", successors=_tag_seq[5:]
    )
    del _tag_seq

    def _new_outerShdw(self) -> CT_OuterShadowEffect:
        """Return a new `a:outerShdw` element with default shadow properties.

        PowerPoint requires a color child element on `a:outerShdw`, so this provides a reasonable
        default: 45-degree angle, 3pt distance, 4pt blur, 40% transparent black.
        """
        return cast(
            CT_OuterShadowEffect,
            parse_xml(
                f'<a:outerShdw {nsdecls("a")} blurRad="50800" dist="38100"'
                f' dir="2700000" algn="tl" rotWithShape="0">'
                f'  <a:srgbClr val="000000">'
                f'    <a:alpha val="40000"/>'
                f"  </a:srgbClr>"
                f"</a:outerShdw>"
            ),
        )
