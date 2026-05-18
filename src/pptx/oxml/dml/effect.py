"""lxml custom element classes for DrawingML effect-related XML elements."""

from __future__ import annotations

from typing import Callable, cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import (
    ST_Angle,
    ST_PositiveCoordinate,
    ST_PositiveFixedAngle,
    ST_PositiveFixedPercentage,
    XsdBoolean,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OptionalAttribute,
    RequiredAttribute,
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


class CT_GlowEffect(BaseOxmlElement):
    """`a:glow` custom element class.

    DrawingML §20.1.8.16. Schema (`CT_GlowEffect`): one mandatory
    `EG_ColorChoice` child and an optional `rad` attribute. The color child
    is what PowerPoint draws the glow in, so — like `a:outerShdw` — it must
    always be present for the effect to render and the file to open clean.
    """

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
    rad = OptionalAttribute("rad", ST_PositiveCoordinate, default=0)


class CT_ReflectionEffect(BaseOxmlElement):
    """`a:reflection` custom element class.

    DrawingML §20.1.8.45. Schema (`CT_ReflectionEffect`) is attribute-only
    (no child elements). Only the attributes in common use are modeled; any
    others present in a loaded file round-trip untouched via lxml.
    """

    blurRad = OptionalAttribute("blurRad", ST_PositiveCoordinate, default=0)
    stA = OptionalAttribute("stA", ST_PositiveFixedPercentage, default=100000)
    stPos = OptionalAttribute("stPos", ST_PositiveFixedPercentage, default=0)
    endA = OptionalAttribute("endA", ST_PositiveFixedPercentage, default=0)
    endPos = OptionalAttribute("endPos", ST_PositiveFixedPercentage, default=100000)
    dist = OptionalAttribute("dist", ST_PositiveCoordinate, default=0)
    dir = OptionalAttribute("dir", ST_PositiveFixedAngle, default=0)
    fadeDir = OptionalAttribute("fadeDir", ST_PositiveFixedAngle, default=5400000)
    rotWithShape: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rotWithShape", XsdBoolean, default=True
    )


class CT_SoftEdgesEffect(BaseOxmlElement):
    """`a:softEdge` custom element class.

    DrawingML §20.1.8.48. Schema (`CT_SoftEdgesEffect`) is a single
    **required** `rad` attribute. PowerPoint rejects a `<a:softEdge>` with no
    `rad` (silent repair), so `rad` is modeled `RequiredAttribute` and the
    factory always sets it.
    """

    rad = RequiredAttribute("rad", ST_PositiveCoordinate)


class CT_EffectList(BaseOxmlElement):
    """`a:effectLst` custom element class."""

    get_or_add_outerShdw: Callable[[], CT_OuterShadowEffect]
    get_or_add_glow: Callable[[], CT_GlowEffect]
    get_or_add_reflection: Callable[[], CT_ReflectionEffect]
    get_or_add_softEdge: Callable[[], CT_SoftEdgesEffect]
    _remove_outerShdw: Callable[[], None]
    _remove_innerShdw: Callable[[], None]
    _remove_glow: Callable[[], None]
    _remove_reflection: Callable[[], None]
    _remove_softEdge: Callable[[], None]

    # ---OOXML schema order (ECMA-376 dml-main.xsd CT_EffectList): blur,
    # ---fillOverlay, glow, innerShdw, outerShdw, prstShdw, reflection,
    # ---softEdge. Emitting a child out of this order is a silent
    # ---PowerPoint-repair trigger, so successors are derived from _tag_seq.
    _tag_seq = (
        "a:blur",
        "a:fillOverlay",
        "a:glow",
        "a:innerShdw",
        "a:outerShdw",
        "a:prstShdw",
        "a:reflection",
        "a:softEdge",
    )
    glow: CT_GlowEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:glow", successors=_tag_seq[3:]
    )
    innerShdw: CT_InnerShadowEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:innerShdw", successors=_tag_seq[4:]
    )
    outerShdw: CT_OuterShadowEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:outerShdw", successors=_tag_seq[5:]
    )
    reflection: CT_ReflectionEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:reflection", successors=_tag_seq[7:]
    )
    softEdge: CT_SoftEdgesEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:softEdge", successors=_tag_seq[8:]
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

    def _new_glow(self) -> CT_GlowEffect:
        """Return a new `a:glow` element with a default color child.

        Schema requires the `EG_ColorChoice` child, so a default 5pt accent-1
        glow is provided; callers typically overwrite the color immediately.
        """
        return cast(
            CT_GlowEffect,
            parse_xml(
                f'<a:glow {nsdecls("a")} rad="63500">  <a:schemeClr val="accent1"/></a:glow>'
            ),
        )

    def _new_reflection(self) -> CT_ReflectionEffect:
        """Return a new `a:reflection` element with a PowerPoint-typical default.

        Mirrors the "Tight Reflection, touching" preset PowerPoint emits from
        its effects gallery so the round-tripped file opens clean.
        """
        return cast(
            CT_ReflectionEffect,
            parse_xml(
                f'<a:reflection {nsdecls("a")} blurRad="6350" stA="52000"'
                f' endA="300" endPos="35000" dir="5400000"'
                f' rotWithShape="0"/>'
            ),
        )

    def _new_softEdge(self) -> CT_SoftEdgesEffect:
        """Return a new `a:softEdge` element.

        `rad` is required by the schema; default 2.5pt feather.
        """
        return cast(
            CT_SoftEdgesEffect,
            parse_xml(f'<a:softEdge {nsdecls("a")} rad="31750"/>'),
        )
