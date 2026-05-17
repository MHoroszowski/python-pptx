"""Presentation part, the main part in a .pptx package."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, Iterable, cast

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.parts.slide import HandoutMasterPart, NotesMasterPart, SlidePart
from pptx.presentation import Presentation
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.custom_properties import CustomProperties
    from pptx.custom_xml import CustomXmlParts
    from pptx.parts.comments import AuthorsPart, CommentAuthorsPart
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.slide import HandoutMaster, NotesMaster, Slide, SlideLayout, SlideMaster


class PresentationPart(XmlPart):
    """Top level class in object model.

    Represents the contents of the /ppt directory of a .pptx file.
    """

    def add_slide(self, slide_layout: SlideLayout):
        """Return (rId, slide) pair of a newly created blank slide.

        New slide inherits appearance from `slide_layout`.
        """
        partname = self._next_slide_partname
        slide_layout_part = slide_layout.part
        slide_part = SlidePart.new(partname, self.package, slide_layout_part)
        rId = self.relate_to(slide_part, RT.SLIDE)
        return rId, slide_part.slide

    @property
    def core_properties(self) -> CorePropertiesPart:
        """A |CoreProperties| object for the presentation.

        Provides read/write access to the Dublin Core properties of this presentation.
        """
        return self.package.core_properties

    @lazyproperty
    def custom_properties(self) -> CustomProperties:
        """Mapping-protocol view over the Custom Document Properties part.

        Lazy — the same wrapper instance is returned across calls. The
        underlying `/docProps/custom.xml` part is created on first access if
        the package does not already have one.
        """
        from pptx.custom_properties import CustomProperties

        return CustomProperties(self.package.custom_properties_part)

    @lazyproperty
    def custom_xml_parts(self) -> CustomXmlParts:
        """Sequence-like collection of customXml data parts in this package.

        Walks both presentation-scoped (`ppt/_rels/presentation.xml.rels`) and
        package-scoped (`/_rels/.rels`) `RT.CUSTOM_XML` relationships. The
        same collection instance is reused across calls.
        """
        from pptx.custom_xml import CustomXmlParts

        return CustomXmlParts(self)

    def get_slide(self, slide_id: int) -> Slide | None:
        """Return optional related |Slide| object identified by `slide_id`.

        Returns |None| if no slide with `slide_id` is related to this presentation.
        """
        for sldId in self._element.sldIdLst:
            if sldId.id == slide_id:
                return self.related_part(sldId.rId).slide
        return None

    @lazyproperty
    def notes_master(self) -> NotesMaster:
        """
        Return the |NotesMaster| object for this presentation. If the
        presentation does not have a notes master, one is created from
        a default template. The same single instance is returned on each
        call.
        """
        return self.notes_master_part.notes_master

    @lazyproperty
    def notes_master_part(self) -> NotesMasterPart:
        """Return the |NotesMasterPart| object for this presentation.

        If the presentation does not have a notes master, one is created from a default template.
        The same single instance is returned on each call.
        """
        try:
            return self.part_related_by(RT.NOTES_MASTER)
        except KeyError:
            notes_master_part = NotesMasterPart.create_default(self.package)
            self.relate_to(notes_master_part, RT.NOTES_MASTER)
            return notes_master_part

    @lazyproperty
    def handout_master(self) -> HandoutMaster:
        """Return the |HandoutMaster| object for this presentation.

        Raises |ValueError| when the presentation has no handout master because auto-create is
        deliberately deferred until a built-in handout-master template ships in this fork.
        """
        return self.handout_master_part.handout_master

    @lazyproperty
    def handout_master_part(self) -> HandoutMasterPart:
        """Return the |HandoutMasterPart| object for this presentation.

        Raises |ValueError| when the presentation has no handout master because auto-create is
        deliberately deferred until a built-in handout-master template ships in this fork.
        """
        try:
            return self.part_related_by(RT.HANDOUT_MASTER)
        except KeyError as e:
            raise ValueError(
                "presentation has no handout master; auto-create is deferred because no "
                "handout master template ships in this fork yet"
            ) from e

    @property
    def has_comment_authors(self) -> bool:
        """`True` if this presentation has a legacy comment-authors part.

        Non-mutating — unlike :attr:`comment_authors_part`, this never
        creates the part. Mirrors the non-mutating ``has_*`` accessors.
        """
        try:
            self.part_related_by(RT.COMMENT_AUTHORS)
        except KeyError:
            return False
        return True

    @lazyproperty
    def comment_authors_part(self) -> CommentAuthorsPart:
        """The presentation-wide |CommentAuthorsPart|, lazily created.

        Mirrors :attr:`notes_master_part`: try the existing
        ``RT.COMMENT_AUTHORS`` relationship and, only on |KeyError|, create
        the singleton ``/ppt/commentAuthors.xml`` part and relate the
        presentation to it. The same single part is returned on every call;
        modern threaded comments reuse this one author list (one identity
        registry per presentation, ISC-5 no-dup policy).
        """
        from pptx.parts.comments import CommentAuthorsPart

        try:
            return cast("CommentAuthorsPart", self.part_related_by(RT.COMMENT_AUTHORS))
        except KeyError:
            comment_authors_part = CommentAuthorsPart.new(self.package)
            self.relate_to(comment_authors_part, RT.COMMENT_AUTHORS)
            return comment_authors_part

    @property
    def has_authors(self) -> bool:
        """`True` if this presentation has a MODERN authors part.

        Non-mutating — unlike :attr:`authors_part`, this never creates the
        part. Mirrors :attr:`has_comment_authors` for the modern
        ``/ppt/authors.xml`` (issue #25).
        """
        try:
            self.part_related_by(RT.AUTHORS)
        except KeyError:
            return False
        return True

    @lazyproperty
    def authors_part(self) -> AuthorsPart:
        """The presentation-wide MODERN |AuthorsPart|, lazily created.

        Mirrors :attr:`comment_authors_part` but for the modern
        ``/ppt/authors.xml`` (``RT.AUTHORS``). This is the GUID-keyed author
        list every modern ``<p188:cm>/@authorId`` resolves into. Distinct
        from the legacy integer-keyed ``commentAuthors.xml`` — modern
        threaded comments must NOT use the legacy part (issue #25
        repair-dialog root cause: orphaned modern author GUIDs).
        """
        from pptx.parts.comments import AuthorsPart

        try:
            return cast("AuthorsPart", self.part_related_by(RT.AUTHORS))
        except KeyError:
            authors_part = AuthorsPart.new(self.package)
            self.relate_to(authors_part, RT.AUTHORS)
            return authors_part

    @lazyproperty
    def presentation(self):
        """
        A |Presentation| object providing access to the content of this
        presentation.
        """
        return Presentation(self._element, self)

    def related_slide(self, rId: str) -> Slide:
        """Return |Slide| object for related |SlidePart| related by `rId`."""
        return self.related_part(rId).slide

    def related_slide_master(self, rId: str) -> SlideMaster:
        """Return |SlideMaster| object for |SlideMasterPart| related by `rId`."""
        return self.related_part(rId).slide_master

    def rename_slide_parts(self, rIds: Iterable[str]):
        """Assign incrementing partnames to the slide parts identified by `rIds`.

        Partnames are like `/ppt/slides/slide9.xml` and are assigned in the order their id appears
        in the `rIds` sequence. The name portion is always `slide`. The number part forms a
        continuous sequence starting at 1 (e.g. 1, 2, ... 10, ...). The extension is always
        `.xml`.
        """
        for idx, rId in enumerate(rIds):
            slide_part = self.related_part(rId)
            slide_part.partname = PackURI("/ppt/slides/slide%d.xml" % (idx + 1))

    def save(self, path_or_stream: str | IO[bytes]):
        """Save this presentation package to `path_or_stream`.

        `path_or_stream` can be either a path to a filesystem location (a string) or a
        file-like object.
        """
        self.package.save(path_or_stream)

    def slide_id(self, slide_part):
        """Return the slide-id associated with `slide_part`."""
        for sldId in self._element.sldIdLst:
            if self.related_part(sldId.rId) is slide_part:
                return sldId.id
        raise ValueError("matching slide_part not found")

    @property
    def _next_slide_partname(self):
        """Return |PackURI| instance containing next available slide partname."""
        sldIdLst = self._element.get_or_add_sldIdLst()
        partname_str = "/ppt/slides/slide%d.xml" % (len(sldIdLst) + 1)
        return PackURI(partname_str)
