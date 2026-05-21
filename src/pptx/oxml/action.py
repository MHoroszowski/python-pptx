"""lxml custom element classes for click-action and hyperlink XML elements."""

from __future__ import annotations

from typing import Callable

from pptx.oxml.simpletypes import ST_RelationshipId, XsdBoolean, XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)


class CT_EmbeddedWAVAudioFile(BaseOxmlElement):
    """`a:snd` element — an embedded WAV audio file played by a click/hover action.

    Schema type `CT_EmbeddedWAVAudioFile` (ECMA-376 dml-main.xsd). `r:embed` is a
    required relationship id pointing at the embedded WAV media part; `name` is an
    optional human-readable label PowerPoint shows in its UI.
    """

    embed: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "r:embed", ST_RelationshipId
    )
    name: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "name", XsdString
    )


class CT_Hyperlink(BaseOxmlElement):
    """Custom element class for `a:hlinkClick`, `a:hlinkHover`, and `a:hlinkMouseOver`.

    A single element type (schema `CT_Hyperlink`) serves three host contexts: the
    click and hover hyperlinks on a shape's `p:cNvPr`, and the click and mouse-over
    hyperlinks on a text run's `a:rPr`. The schema sequence is `snd?`, `extLst?`;
    emitting `a:snd` after `a:extLst` is a silent PowerPoint repair, so `snd` is
    declared with `a:extLst` as its only successor.
    """

    get_or_add_snd: Callable[[], CT_EmbeddedWAVAudioFile]
    _remove_snd: Callable[[], None]

    snd: CT_EmbeddedWAVAudioFile | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:snd", successors=("a:extLst",)
    )

    rId: str = OptionalAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]
    invalidUrl: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "invalidUrl", XsdString
    )
    action: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "action", XsdString
    )
    tgtFrame: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "tgtFrame", XsdString
    )
    tooltip: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "tooltip", XsdString
    )
    history: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "history", XsdBoolean
    )
    highlightClick: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "highlightClick", XsdBoolean
    )
    endSnd: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "endSnd", XsdBoolean
    )

    @property
    def action_fields(self) -> dict[str, str]:
        """Query portion of the `ppaction://` URL as dict.

        For example `{'id':'0', 'return':'true'}` in 'ppaction://customshow?id=0&return=true'.

        Returns an empty dict if the URL contains no query string or if no action attribute is
        present.
        """
        url = self.action

        if url is None:
            return {}

        halves = url.split("?")
        if len(halves) == 1:
            return {}

        key_value_pairs = halves[1].split("&")
        return dict([pair.split("=") for pair in key_value_pairs])

    @property
    def action_verb(self) -> str | None:
        """The host portion of the `ppaction://` URL contained in the action attribute.

        For example 'customshow' in 'ppaction://customshow?id=0&return=true'. Returns |None| if no
        action attribute is present.
        """
        url = self.action

        if url is None:
            return None

        protocol_and_host = url.split("?")[0]
        host = protocol_and_host[11:]

        return host
