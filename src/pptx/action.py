"""Objects related to mouse click and hover actions on a shape or text."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast
from urllib.parse import quote

from pptx.enum.action import PP_ACTION
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.shapes import Subshape
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.action import CT_Hyperlink
    from pptx.oxml.shapes.shared import CT_NonVisualDrawingProps
    from pptx.oxml.text import CT_TextCharacterProperties
    from pptx.oxml.xmlchemy import BaseOxmlElement
    from pptx.parts.slide import SlidePart
    from pptx.shapes.base import BaseShape
    from pptx.slide import Slide, Slides


class ActionSetting(Subshape):
    """Properties specifying how a shape or run reacts to mouse actions."""

    # -- The Subshape base class provides access to the Slide Part, which is needed to access
    # -- relationships, which is where hyperlinks live.

    def __init__(
        self,
        xPr: CT_NonVisualDrawingProps | CT_TextCharacterProperties,
        parent: BaseShape,
        hover: bool = False,
    ):
        super(ActionSetting, self).__init__(parent)
        # xPr is either a cNvPr or rPr element
        self._element = xPr
        # _hover determines use of `a:hlinkClick` or `a:hlinkHover`
        self._hover = hover

    @property
    def action(self):
        """Member of :ref:`PpActionType` enumeration, such as `PP_ACTION.HYPERLINK`.

        The returned member indicates the type of action that will result when the
        specified shape or text is clicked or the mouse pointer is positioned over the
        shape during a slide show.

        If there is no click-action or the click-action value is not recognized (is not
        one of the official `MsoPpAction` values) then `PP_ACTION.NONE` is returned.
        """
        hlink = self._hlink

        if hlink is None:
            return PP_ACTION.NONE

        action_verb = hlink.action_verb

        if action_verb == "hlinkshowjump":
            relative_target = hlink.action_fields["jump"]
            return {
                "firstslide": PP_ACTION.FIRST_SLIDE,
                "lastslide": PP_ACTION.LAST_SLIDE,
                "lastslideviewed": PP_ACTION.LAST_SLIDE_VIEWED,
                "nextslide": PP_ACTION.NEXT_SLIDE,
                "previousslide": PP_ACTION.PREVIOUS_SLIDE,
                "endshow": PP_ACTION.END_SHOW,
            }[relative_target]

        return {
            None: PP_ACTION.HYPERLINK,
            "hlinksldjump": PP_ACTION.NAMED_SLIDE,
            "hlinkpres": PP_ACTION.PLAY,
            "hlinkfile": PP_ACTION.OPEN_FILE,
            "customshow": PP_ACTION.NAMED_SLIDE_SHOW,
            "ole": PP_ACTION.OLE_VERB,
            "macro": PP_ACTION.RUN_MACRO,
            "program": PP_ACTION.RUN_PROGRAM,
        }.get(action_verb, PP_ACTION.NONE)

    @lazyproperty
    def hyperlink(self) -> Hyperlink:
        """
        A |Hyperlink| object representing the hyperlink action defined on
        this click or hover mouse event. A |Hyperlink| object is always
        returned, even if no hyperlink or other click action is defined.
        """
        return Hyperlink(self._element, self._parent, self._hover)

    @property
    def target_slide(self) -> Slide | None:
        """
        A reference to the slide in this presentation that is the target of
        the slide jump action in this shape. Slide jump actions include
        `PP_ACTION.FIRST_SLIDE`, `LAST_SLIDE`, `NEXT_SLIDE`,
        `PREVIOUS_SLIDE`, and `NAMED_SLIDE`. Returns |None| for all other
        actions. In particular, the `LAST_SLIDE_VIEWED` action and the `PLAY`
        (start other presentation) actions are not supported.

        A slide object may be assigned to this property, which makes the
        shape an "internal hyperlink" to the assigened slide::

            slide, target_slide = prs.slides[0], prs.slides[1]
            shape = slide.shapes[0]
            shape.target_slide = target_slide

        Assigning |None| removes any slide jump action. Note that this is
        accomplished by removing any action present (such as a hyperlink),
        without first checking that it is a slide jump action.
        """
        slide_jump_actions = (
            PP_ACTION.FIRST_SLIDE,
            PP_ACTION.LAST_SLIDE,
            PP_ACTION.NEXT_SLIDE,
            PP_ACTION.PREVIOUS_SLIDE,
            PP_ACTION.NAMED_SLIDE,
        )

        if self.action not in slide_jump_actions:
            return None

        if self.action == PP_ACTION.FIRST_SLIDE:
            return self._slides[0]
        elif self.action == PP_ACTION.LAST_SLIDE:
            return self._slides[-1]
        elif self.action == PP_ACTION.NEXT_SLIDE:
            next_slide_idx = self._slide_index + 1
            if next_slide_idx >= len(self._slides):
                raise ValueError("no next slide")
            return self._slides[next_slide_idx]
        elif self.action == PP_ACTION.PREVIOUS_SLIDE:
            prev_slide_idx = self._slide_index - 1
            if prev_slide_idx < 0:
                raise ValueError("no previous slide")
            return self._slides[prev_slide_idx]
        elif self.action == PP_ACTION.NAMED_SLIDE:
            assert self._hlink is not None
            rId = self._hlink.rId
            slide_part = cast("SlidePart", self.part.related_part(rId))
            return slide_part.slide

    @target_slide.setter
    def target_slide(self, slide: Slide | None):
        self._clear_click_action()
        if slide is None:
            return
        hlink = self._element.get_or_add_hlinkClick()
        hlink.action = "ppaction://hlinksldjump"
        hlink.rId = self.part.relate_to(slide.part, RT.SLIDE)

    def target_program(self, file_path: str) -> None:
        """Make this a "run program" click/hover action targeting `file_path`.

        Emits `action="ppaction://program"` on the hyperlink element and relates an
        external relationship to `file_path`. The resulting :attr:`action` is
        `PP_ACTION.RUN_PROGRAM`. Any prior click/hover action is replaced.
        """
        self._clear_click_action()
        rId = self.part.relate_to(file_path, RT.HYPERLINK, is_external=True)
        hlink = self._get_or_add_hlink()
        hlink.action = "ppaction://program"
        hlink.rId = rId

    def run_macro(self, macro_name: str) -> None:
        """Make this a "run macro" click/hover action invoking `macro_name`.

        Emits `action="ppaction://macro?name=<macro_name>"`. The macro name is
        URL-quoted so names containing spaces remain a single, well-formed action
        URL. No relationship is allocated — a macro action is name-only. Any prior
        click/hover action is replaced.
        """
        self._clear_click_action()
        hlink = self._get_or_add_hlink()
        hlink.action = "ppaction://macro?name=%s" % quote(macro_name, safe="")

    def play_sound(self, audio_file: str, mime_type: str = "audio/wav") -> None:
        """Attach an embedded sound that plays when this shape is clicked/hovered.

        Adds an `a:snd` child referencing an embedded WAV media part created from
        `audio_file` (a path or file-like object). `endSnd` is set so the sound
        stops on the next action, matching PowerPoint's own default. The sound
        coexists with any hyperlink/action already present — it does not clear it.

        The relationship from the hlink's `a:snd` child to the WAV part uses
        `RT.AUDIO` (ECMA-376 `relationships/audio`), which is the rel type
        PowerPoint expects on `CT_EmbeddedWAVAudioFile/@r:embed`. Using the
        Microsoft-2007 `relationships/media` rel here triggers a load-time
        Repair dialog in PowerPoint that strips the entire hlinkClick element
        and the embedded WAV part — `relationships/media` is for video media.
        """
        from pptx.media import Video

        media = Video.from_path_or_file_like(audio_file, mime_type)
        media_part = self.part.package.get_or_add_media_part(media)
        rId = self.part.relate_to(media_part, RT.AUDIO)
        hlink = self._get_or_add_hlink()
        snd = hlink.get_or_add_snd()
        snd.embed = rId
        if media.filename:
            snd.name = media.filename
        hlink.endSnd = True

    @property
    def tooltip(self) -> str | None:
        """Read/write. The ScreenTip text shown on hover over this shape/run.

        Returns |None| when no hyperlink element is present or its `tooltip`
        attribute is unset or empty. Assigning a string creates the hyperlink
        element if necessary. Assigning |None| or `""` removes the tooltip,
        pruning the hyperlink element if it then carries no other action.

        Known PowerPoint limitation: a hyperlink ScreenTip is only rendered
        on hover when the hyperlink also carries a real navigation target —
        a URL via :attr:`hyperlink.address`, a slide jump via
        :meth:`target_slide`, or a macro/program action. A bare tooltip
        without a click target round-trips through save/reload as valid
        OOXML but PowerPoint will not surface it on hover. There is no
        known fully-supported workaround at the OOXML layer for a
        no-click-action hover ScreenTip; :attr:`BaseShape.alt_text` is the
        nearest analog and is the right home for accessibility/screen-reader
        text, but PowerPoint does not render it on slideshow hover either.
        """
        hlink = self._hlink
        if hlink is None:
            return None
        return hlink.tooltip or None

    @tooltip.setter
    def tooltip(self, value: str | None) -> None:
        if not value:
            hlink = self._hlink
            if hlink is None:
                return
            hlink.tooltip = None
            _ensure_noaction_pruned(hlink)
            _prune_hlink_if_empty(self._element, hlink)
            return
        hlink = self._get_or_add_hlink()
        hlink.tooltip = value
        _ensure_noaction_if_inert(hlink)

    def _get_or_add_hlink(self) -> CT_Hyperlink:
        """The `a:hlinkClick` or `a:hlinkHover` element, created if absent.

        Newly-created hlinks get `r:id=""` as the default. PowerPoint's load-time
        validation triggers a Repair dialog on `a:hlinkClick`/`a:hlinkHover`
        elements that omit the `r:id` attribute entirely (even though ECMA-376
        marks it `use="optional"`). Real PowerPoint output always carries `r:id`,
        empty when there is no relationship — mirroring the precedent already
        baked into `oxml/shapes/picture.py` for `ppaction://media`.
        """
        if self._hover:
            hlink = cast("CT_NonVisualDrawingProps", self._element).get_or_add_hlinkHover()
        else:
            hlink = self._element.get_or_add_hlinkClick()
        if hlink.rId is None:
            hlink.rId = ""
        return hlink

    def _clear_click_action(self):
        """Remove any existing click action."""
        hlink = self._hlink
        if hlink is None:
            return
        rId = hlink.rId
        if rId:
            self.part.drop_rel(rId)
        self._element.remove(hlink)

    @property
    def _hlink(self) -> CT_Hyperlink | None:
        """
        Reference to the `a:hlinkClick` or `a:hlinkHover` element for this
        click action. Returns |None| if the element is not present.
        """
        if self._hover:
            assert isinstance(self._element, CT_NonVisualDrawingProps)
            return self._element.hlinkHover
        return self._element.hlinkClick

    @lazyproperty
    def _slide(self):
        """
        Reference to the slide containing the shape having this click action.
        """
        return self.part.slide

    @lazyproperty
    def _slide_index(self):
        """
        Position in the slide collection of the slide containing the shape
        having this click action.
        """
        return self._slides.index(self._slide)

    @lazyproperty
    def _slides(self) -> Slides:
        """
        Reference to the slide collection for this presentation.
        """
        return self.part.package.presentation_part.presentation.slides


class Hyperlink(Subshape):
    """Represents a hyperlink action on a shape or text run."""

    def __init__(
        self,
        xPr: CT_NonVisualDrawingProps | CT_TextCharacterProperties,
        parent: BaseShape,
        hover: bool = False,
    ):
        super(Hyperlink, self).__init__(parent)
        # xPr is either a cNvPr or rPr element
        self._element = xPr
        # _hover determines use of `a:hlinkClick` or `a:hlinkHover`
        self._hover = hover

    @property
    def address(self) -> str | None:
        """Read/write. The URL of the hyperlink.

        URL can be on http, https, mailto, or file scheme; others may work. Returns |None| if no
        hyperlink is defined, including when another action such as `RUN_MACRO` is defined on the
        object. Assigning |None| removes any action defined on the object, whether it is a hyperlink
        action or not.
        """
        hlink = self._hlink

        # there's no URL if there's no click action
        if hlink is None:
            return None

        # a click action without a relationship has no URL
        rId = hlink.rId
        if not rId:
            return None

        return self.part.target_ref(rId)

    @address.setter
    def address(self, url: str | None):
        # implements all three of add, change, and remove hyperlink
        self._remove_hlink()

        if url:
            rId = self.part.relate_to(url, RT.HYPERLINK, is_external=True)
            hlink = self._get_or_add_hlink()
            hlink.rId = rId

    def _get_or_add_hlink(self) -> CT_Hyperlink:
        """Get the `a:hlinkClick` or `a:hlinkHover` element for the Hyperlink object.

        The actual element depends on the value of `self._hover`. Create the element if not present.

        Newly-created hlinks get `r:id=""` as the default — see the matching note
        on `ActionSetting._get_or_add_hlink`. Without it, PowerPoint throws a
        Repair dialog and strips the element on load.
        """
        if self._hover:
            hlink = cast("CT_NonVisualDrawingProps", self._element).get_or_add_hlinkHover()
        else:
            hlink = self._element.get_or_add_hlinkClick()
        if hlink.rId is None:
            hlink.rId = ""
        return hlink

    @property
    def _hlink(self) -> CT_Hyperlink | None:
        """Reference to the `a:hlinkClick` or `h:hlinkHover` element for this click action.

        Returns |None| if the element is not present.
        """
        if self._hover:
            return cast("CT_NonVisualDrawingProps", self._element).hlinkHover
        return self._element.hlinkClick

    @property
    def tooltip(self) -> str | None:
        """Read/write. The ScreenTip text shown on hover over this hyperlink.

        Returns |None| when no hyperlink element is present or its `tooltip`
        attribute is unset or empty. Assigning a string creates the hyperlink
        element if necessary. Assigning |None| or `""` removes the tooltip,
        pruning the hyperlink element if it then carries no URL or action.

        Known PowerPoint limitation: this ScreenTip is only rendered on
        hover when the hyperlink also carries a real navigation target — a
        URL, slide jump, macro, or program action. A bare tooltip-only
        hyperlink is valid OOXML and round-trips through save/reload, but
        PowerPoint will not surface it on hover. There is no known
        fully-supported workaround at the OOXML layer for a
        no-click-action hover ScreenTip.
        """
        hlink = self._hlink
        if hlink is None:
            return None
        return hlink.tooltip or None

    @tooltip.setter
    def tooltip(self, value: str | None) -> None:
        if not value:
            hlink = self._hlink
            if hlink is None:
                return
            hlink.tooltip = None
            _ensure_noaction_pruned(hlink)
            _prune_hlink_if_empty(self._element, hlink)
            return
        hlink = self._get_or_add_hlink()
        hlink.tooltip = value
        _ensure_noaction_if_inert(hlink)

    def _remove_hlink(self):
        """Remove the a:hlinkClick or a:hlinkHover element.

        Also drops any relationship it might have.
        """
        hlink = self._hlink
        if hlink is None:
            return
        rId = hlink.rId
        if rId:
            self.part.drop_rel(rId)
        self._element.remove(hlink)


def _prune_hlink_if_empty(parent: BaseOxmlElement, hlink: CT_Hyperlink) -> None:
    """Remove `hlink` from `parent` when it carries no URL, action, or sound.

    Used by the `tooltip` setters: clearing a tooltip should not leave behind an
    inert, empty `a:hlinkClick`/`a:hlinkHover` element (which PowerPoint tolerates
    but which is noise). An hlink is "empty" when it has no `r:id`, no `action`,
    no `tooltip`, no `tgtFrame`, and no `a:snd` child.

    The synthetic `ppaction://noaction` verb (see `_ensure_noaction_if_inert`)
    is treated as "no real action" for prune purposes — it is a marker we add
    so PowerPoint accepts a tooltip-bearing hlink without triggering its
    load-time Repair dialog. When everything else falls away, the noaction
    marker should fall away too.
    """
    if hlink.rId:
        return
    if hlink.action is not None and hlink.action != "ppaction://noaction":
        return
    if hlink.tooltip:
        return
    if hlink.tgtFrame:
        return
    if hlink.snd is not None:
        return
    parent.remove(hlink)


def _ensure_noaction_if_inert(hlink: CT_Hyperlink) -> None:
    """Set `action="ppaction://noaction"` on an inert hlink for PowerPoint compatibility.

    A bare `<a:hlinkClick r:id="" tooltip="..."/>` with no action verb survives
    a save/reload round-trip at the XML layer, but real PowerPoint output for
    an inert (no-URL, no-jump) hlink always carries `action="ppaction://noaction"`.
    Adding the marker keeps our output isomorphic with PowerPoint's own emission
    and protects the hlink from being stripped by future PowerPoint validators.

    Note: this marker does NOT cause PowerPoint to render the tooltip on hover.
    PowerPoint's hover-ScreenTip processor activates only on hlinks with a real
    navigation target (URL, slide jump, macro, program). For a pure hover
    ScreenTip on a shape with no click behavior, callers should use
    `BaseShape.alt_text` (`cNvPr/@descr`) instead — that is PowerPoint's
    documented mechanism for non-hyperlink hover ScreenTips.

    Only adds the marker when (a) the hlink has no `r:id` relationship AND
    (b) the hlink has no other `action` verb. If a URL is added later, the
    address setter rebuilds the hlink from scratch; if a `target_slide` /
    `run_macro` / `target_program` is added later, those overwrite `action`
    explicitly. The marker is a no-op for any hlink that already has a real
    action or a real relationship.
    """
    if hlink.rId:
        return
    if hlink.action is not None:
        return
    hlink.action = "ppaction://noaction"


def _ensure_noaction_pruned(hlink: CT_Hyperlink) -> None:
    """Remove the synthetic `ppaction://noaction` marker if it has become unneeded.

    Called from the tooltip-clear path: when the tooltip is removed, the
    marker we added in `_ensure_noaction_if_inert` should not linger as the
    sole content of an otherwise-empty hlink. `_prune_hlink_if_empty` runs
    immediately after and treats noaction as prune-eligible, but this helper
    also handles the case where an hlink retains a sibling (e.g. an `<a:snd>`)
    that should NOT carry the noaction marker.
    """
    if hlink.action != "ppaction://noaction":
        return
    if hlink.rId:
        return
    hlink.action = None
