"""Custom element classes for presentation comments (issue #25, Wave 1).

Two comment families live here:

* **Legacy comments** — the original PresentationML comment model
  (ECMA-376 Part 1, §19.5: ``p:cmAuthorLst``/``p:cmAuthor`` in
  ``commentAuthors.xml`` and ``p:cmLst``/``p:cm`` in a per-slide comments
  part). One shared author list for the whole presentation; comments carry
  an ``authorId`` foreign key into it.

* **Modern threaded comments** — Microsoft's post-2018 model in the
  ``http://schemas.microsoft.com/office/powerpoint/2018/8/main`` namespace
  (prefix ``p188``). A ``<p188:cm>`` carries a GUID ``id``, an author GUID,
  a ``created`` timestamp, a ``<p188:txBody>`` and an optional
  ``<p188:replyLst>`` of ``<p188:reply>`` children — this is the structure
  that gives the feature its "threaded" name. Author identities for the
  modern model live in a separate ``authors`` part (Wave 2).

There is no upstream `scanny/python-pptx` prior art for any of this; the
element shapes below are derived directly from the OOXML / MS-OOXML
schemas and the registration conventions in ``pptx/oxml/slide.py``.
"""

from __future__ import annotations

import uuid
from typing import cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.simpletypes import XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

# DrawingML namespace used inside a ``<p188:txBody>`` (shared by the modern
# threaded comment and its replies). The body is a minimal DrawingML
# text-body: a ``<a:bodyPr>`` followed by one ``<a:p>`` paragraph carrying a
# single ``<a:r>/<a:t>`` run for the plain-text payload.
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _txBody_text(txBody) -> str:
    """Return the concatenated plain text of every ``a:t`` under `txBody`."""
    if txBody is None:
        return ""
    return "".join(t.text or "" for t in txBody.iter(qn("a:t")))


def _set_txBody_text(txBody, value: str) -> None:
    """Replace `txBody`'s body with a single paragraph/run holding `value`.

    ``<p188:txBody>`` is a DrawingML ``a:CT_TextBody`` whose schema REQUIRES
    ``<a:bodyPr>`` as its first child. A txBody without it is silently
    dropped by PowerPoint's threaded-comment reader — the comment never
    appears in the Comments pane (no repair dialog, just absent). The
    add-path builds the txBody via ``get_or_add_txBody()`` (a bare
    ``<p188:txBody/>``), so this function must GUARANTEE the leading
    ``<a:bodyPr/>``, not merely preserve a pre-existing one. A trailing
    ``<a:p><a:r><a:t>value</a:t></a:r></a:p>`` follows it.
    """
    for p in list(txBody.findall(qn("a:p"))):
        txBody.remove(p)
    if txBody.find(qn("a:bodyPr")) is None:
        txBody.insert(0, parse_xml('<a:bodyPr xmlns:a="%s"/>' % _A_NS))
    p = parse_xml(
        '<a:p xmlns:a="%s"><a:r><a:t>%s</a:t></a:r></a:p>' % (_A_NS, _escape_xml_text(value))
    )
    txBody.append(p)


def _escape_xml_text(value: str) -> str:
    """Minimal XML-text escaping for the comment-body fast path."""
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# -- legacy author list (commentAuthors.xml) -----------------------------------


class CT_CommentAuthorList(BaseOxmlElement):
    """`p:cmAuthorLst` element, root of the ``commentAuthors.xml`` part.

    ECMA-376 Part 1 §19.5.7. A presentation has at most one of these; every
    legacy ``p:cm`` references an entry here by ``authorId``.
    """

    cmAuthor_lst: list[CT_CommentAuthor]
    _add_cmAuthor: "callable"

    cmAuthor = ZeroOrMore("p:cmAuthor", successors=())

    @classmethod
    def new(cls) -> CT_CommentAuthorList:
        """Return a new empty ``<p:cmAuthorLst>`` element."""
        return cast(CT_CommentAuthorList, parse_xml("<p:cmAuthorLst %s/>" % nsdecls("p")))

    @property
    def next_author_id(self) -> int:
        """The lowest unused non-negative integer author id.

        ``p:cmAuthor/@id`` is an ``ST_Index`` (xsd:unsignedInt). PowerPoint's
        own files start author ids at ``0`` and increment; allocating
        ``max(existing) + 1`` (or ``0`` when empty) matches that and keeps
        the ids stable across edits.
        """
        used: set[int] = set()
        for ca in self.cmAuthor_lst:
            raw = ca.get("id")
            if raw is not None and raw.lstrip("-").isdigit():
                used.add(int(raw))
        return (max(used) + 1) if used else 0

    def add_author(self, name: str, initials: str = "") -> CT_CommentAuthor:
        """Append and return a new ``<p:cmAuthor>`` with the next free id."""
        new_id = self.next_author_id
        author = cast(CT_CommentAuthor, self._add_cmAuthor())
        author.id = new_id
        author.name = name
        author.initials = initials
        author.lastIdx = 0
        author.clrIdx = 0
        return author

    def get_or_add_author(self, name: str, initials: str = "") -> CT_CommentAuthor:
        """Return the existing same-``name`` author, else add a new one.

        Name match is exact and case-sensitive — the same identity policy
        PowerPoint uses for the legacy author list (ISC-5, no duplicates).
        """
        for author in self.cmAuthor_lst:
            if author.name == name:
                return author
        return self.add_author(name, initials)


class CT_CommentAuthor(BaseOxmlElement):
    """`p:cmAuthor` element, one legacy comment author. ECMA-376 §19.5.1."""

    id: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "id", XsdUnsignedInt
    )
    name: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "name", XsdString
    )
    initials: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "initials", XsdString, default=""
    )
    lastIdx: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "lastIdx", XsdUnsignedInt, default=0
    )
    clrIdx: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "clrIdx", XsdUnsignedInt, default=0
    )


# -- modern author list (authors.xml, p188 / 2018/8/main) ----------------------


def _author_guid_for_name(name: str) -> str:
    """Deterministic, stable brace-wrapped uppercase GUID for `name`.

    A ``uuid5`` over the name gives a stable, round-trip-safe GUID so the
    same author always maps to the same ``<p188:author>/@id`` (and therefore
    the same ``<p188:cm>/@authorId``) across edits. Resolution back to a
    name reads the authors part rather than re-deriving, so correctness does
    not depend on this derivation — it only keeps ids stable for dedup.
    """
    return "{%s}" % str(uuid.uuid5(uuid.NAMESPACE_OID, "pptx-comment-author:" + name)).upper()


def _initials_for(name: str) -> str:
    """Best-effort initials from a display `name` (e.g. "Ada Lovelace"→"AL")."""
    parts = [p for p in name.split() if p]
    if not parts:
        return ""
    return "".join(p[0] for p in parts).upper()


class CT_AuthorList(BaseOxmlElement):
    """`p188:authorLst` element, root of the modern ``/ppt/authors.xml`` part.

    Microsoft `2018/8/main` schema ([MS-PPTX] threaded-comments extension).
    Holds zero or more ``<p188:author>`` entries keyed by GUID. Every modern
    ``<p188:cm>``/``<p188:reply>`` references one of these by its GUID
    ``authorId`` — distinct from the legacy ``<p:cmAuthorLst>`` (integer
    ids). A presentation has at most one of these parts.
    """

    author_lst: list[CT_Author]
    _add_author: "callable"

    author = ZeroOrMore("p188:author", successors=())

    @classmethod
    def new(cls) -> CT_AuthorList:
        """Return a new empty ``<p188:authorLst>`` element."""
        return cast(CT_AuthorList, parse_xml("<p188:authorLst %s/>" % nsdecls("p188")))

    def get_or_add_author(self, name: str, initials: str = "") -> CT_Author:
        """Return the existing same-``name`` author, else add a new one.

        Name match is exact and case-sensitive (ISC-5, no duplicates). The
        new author's ``id`` is the deterministic GUID for `name`, so two
        comments by the same author share one ``<p188:author>`` and one GUID.
        """
        for author in self.author_lst:
            if author.name == name:
                return author
        new_author = cast(CT_Author, self._add_author())
        new_author.id = _author_guid_for_name(name)
        new_author.name = name
        new_author.initials = initials or _initials_for(name)
        new_author.userId = name
        new_author.providerId = "None"
        return new_author


class CT_Author(BaseOxmlElement):
    """`p188:author` element, one modern threaded-comment author.

    Microsoft `2018/8/main`. ``id`` is the GUID that ``<p188:cm>/@authorId``
    foreign-keys; ``name`` is the display name; ``initials``, ``userId`` and
    ``providerId`` carry the identity metadata PowerPoint writes
    (``providerId`` is ``"None"`` for a locally-authored identity).
    """

    id: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "id", XsdString
    )
    name: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "name", XsdString
    )
    initials: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "initials", XsdString, default=""
    )
    userId: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "userId", XsdString, default=""
    )
    # providerId is RequiredAttribute (not Optional-with-default) so it is
    # ALWAYS serialized. Real PowerPoint-authored `<p188:author>` elements
    # always carry providerId; omitting it (which an Optional default would
    # do) is NOT verified-equivalent on PowerPoint's read path for this MS
    # extension schema and is a plausible repair trigger.
    providerId: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "providerId", XsdString
    )


# -- legacy comment list (per-slide comments part) -----------------------------


class CT_CommentList(BaseOxmlElement):
    """`p:cmLst` element, root of a legacy per-slide comments part.

    ECMA-376 Part 1 §19.5.8.
    """

    cm_lst: list[CT_Comment]

    cm = ZeroOrMore("p:cm", successors=())

    @classmethod
    def new(cls) -> CT_CommentList:
        """Return a new empty ``<p:cmLst>`` element."""
        return cast(CT_CommentList, parse_xml("<p:cmLst %s/>" % nsdecls("p")))


class CT_Comment(BaseOxmlElement):
    """`p:cm` element, one legacy comment on a slide. ECMA-376 §19.5.6.

    ``<p:pos>`` (anchor x/y) and ``<p:text>`` (the comment body) are the
    two children; ``authorId`` foreign-keys the presentation-level author
    list and ``idx`` is the per-author comment ordinal.
    """

    pos = ZeroOrOne("p:pos", successors=("p:text", "p:extLst"))
    _text = ZeroOrOne("p:text", successors=("p:extLst",))

    authorId: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "authorId", XsdUnsignedInt
    )
    dt: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dt", XsdString
    )
    idx: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "idx", XsdUnsignedInt
    )

    @property
    def text(self) -> str:
        """Plain-text body of this comment (``<p:text>`` content)."""
        text_elm = self._text
        if text_elm is None:
            return ""
        return text_elm.text or ""

    @text.setter
    def text(self, value: str) -> None:
        text_elm = self.get_or_add__text()
        text_elm.text = value


class CT_CommentPosition(BaseOxmlElement):
    """`p:pos` element — the slide-space anchor point of a legacy comment."""

    x: int = RequiredAttribute("x", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
    y: int = RequiredAttribute("y", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


# -- modern threaded comments (p188 / 2018/8/main) -----------------------------


class CT_ThreadedComment(BaseOxmlElement):
    """`p188:cm` element, a modern threaded comment.

    Microsoft `2018/8/main` schema. A top-level comment carries a GUID
    ``id``, an author GUID ``authorId``, an ISO ``created`` timestamp, a
    ``<p188:txBody>`` rich-text body, and an optional ``<p188:replyLst>``
    holding the reply thread. ``status`` (e.g. ``"resolved"``) marks a
    closed thread.

    GROUND TRUTH (2026-05-17, captured from a comment PowerPoint for Mac
    authored+saved): the FIRST child of every top-level ``<p188:cm>`` is a
    ``<pc:sldMkLst>`` slide-binding marker
    (``<pc:docMk/><pc:sldMk cId=".." sldId=".."/>``, ns
    ``http://schemas.microsoft.com/office/powerpoint/2013/main/command``).
    This is how PowerPoint binds a comment to its slide — NOT the per-slide
    relationship and NOT the fork-private ``extLst`` anchor. Omitting it
    leaves the comment unbound and PowerPoint does not render it (issue #25
    empty-Comments-pane root cause). Replies do not carry their own marker.
    """

    # Child order is FIXED by [MS-PPTX] CT_Comment (authoritative spec, the
    # contract PowerPoint implements):
    #   <xsd:sequence>
    #     EG_CommentAnchor (pc:sldMkLst)   minOccurs=1
    #     pos      (a:CT_Point2D)          minOccurs=0
    #     replyLst (CT_CommentReplyList)   minOccurs=0
    #     EG_CommentProperties (txBody, extLst) minOccurs=1
    #   </xsd:sequence>
    # i.e. replyLst MUST precede txBody. Emitting <p188:replyLst> AFTER
    # <p188:txBody> (the pre-2026-05-17 order) is out-of-sequence: PowerPoint
    # renders the parent comment but SILENTLY DROPS every reply. The successor
    # tuples below encode this exact order (pos omitted — optional, unused).
    sldMkLst = ZeroOrOne("pc:sldMkLst", successors=("p188:replyLst", "p188:txBody", "p188:extLst"))
    replyLst: "CT_ThreadedCommentReplyList | None" = ZeroOrOne(  # pyright: ignore
        "p188:replyLst", successors=("p188:txBody", "p188:extLst")
    )
    txBody = ZeroOrOne("p188:txBody", successors=("p188:extLst",))

    id: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "id", XsdString
    )
    authorId: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "authorId", XsdString
    )
    created: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "created", XsdString
    )
    status: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "status", XsdString
    )
    extLst = ZeroOrOne("p188:extLst", successors=())

    # ---Shape anchor (Wave-2 SF7). NOT a schema attribute: the 2018/8/main
    #    `<p188:cm>` schema defines only id/authorId/created/status. An
    #    out-of-schema attribute on p188:cm is a plausible PowerPoint
    #    repair-dialog trigger (Cato audit), so the anchored shape id is
    #    stored in a `<p188:extLst>/<p188:ext>` keyed by a fork-private URI
    #    — the OOXML-sanctioned extension mechanism that conformant
    #    consumers MUST ignore when the URI is unknown.
    _ANCHOR_EXT_URI = "https://github.com/MHoroszowski/python-pptx/ns/comment-anchor"
    _ANCHOR_NS = _ANCHOR_EXT_URI
    _ANCHOR_TAG = "{%s}anchor" % _ANCHOR_NS

    @property
    def anchorShapeId(self) -> int | None:
        """Anchored shape id, read from the fork extLst extension, else None."""
        extLst = self.extLst
        if extLst is None:
            return None
        for ext in extLst.findall(qn("p188:ext")):
            if ext.get("uri") != self._ANCHOR_EXT_URI:
                continue
            anchor = ext.find(self._ANCHOR_TAG)
            if anchor is None:
                return None
            raw = anchor.get("shapeId")
            return int(raw) if raw is not None and raw.lstrip("-").isdigit() else None
        return None

    @anchorShapeId.setter
    def anchorShapeId(self, shape_id: int) -> None:
        extLst = self.get_or_add_extLst()
        for ext in extLst.findall(qn("p188:ext")):
            if ext.get("uri") == self._ANCHOR_EXT_URI:
                anchor = ext.find(self._ANCHOR_TAG)
                if anchor is None:
                    anchor = ext.makeelement(self._ANCHOR_TAG, {})
                    ext.append(anchor)
                anchor.set("shapeId", str(shape_id))
                return
        ext = extLst.makeelement(qn("p188:ext"), {"uri": self._ANCHOR_EXT_URI})
        anchor = ext.makeelement(self._ANCHOR_TAG, {"shapeId": str(shape_id)})
        ext.append(anchor)
        extLst.append(ext)

    def set_slide_marker(self, slide_id: int, c_id: int = 0) -> None:
        """Set the ``<pc:sldMkLst>`` binding this comment to slide `slide_id`.

        Builds ``<pc:sldMkLst><pc:docMk/><pc:sldMk cId=c_id sldId=slide_id/>
        </pc:sldMkLst>`` and places it as the FIRST child of this
        ``<p188:cm>`` (before ``<p188:txBody>``), matching the structure
        PowerPoint itself emits. ``slide_id`` is the slide's
        ``<p:sldId>/@id`` (e.g. 256 for the first slide); ``c_id`` is the
        collaboration/change id PowerPoint stamps ``0`` for a locally
        authored comment. Replaces any pre-existing marker (idempotent).
        """
        existing = self.sldMkLst
        if existing is not None:
            self.remove(existing)
        self.insert(
            0,
            parse_xml(
                "<pc:sldMkLst %s><pc:docMk/><pc:sldMk "
                'cId="%d" sldId="%d"/></pc:sldMkLst>' % (nsdecls("pc"), c_id, slide_id)
            ),
        )

    @classmethod
    def new(cls, comment_id: str, author_id: str, created: str) -> CT_ThreadedComment:
        """Return a minimal-valid ``<p188:cm>`` with an empty text body."""
        xml = (
            '<p188:cm %s id="%s" authorId="%s" created="%s">'
            "<p188:txBody>"
            '<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
            '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
            "</p188:txBody>"
            "</p188:cm>" % (nsdecls("p188"), comment_id, author_id, created)
        )
        return cast(CT_ThreadedComment, parse_xml(xml))

    @property
    def text(self) -> str:
        """Plain-text body of this threaded comment."""
        return _txBody_text(self.txBody)

    @text.setter
    def text(self, value: str) -> None:
        self.get_or_add_txBody()
        _set_txBody_text(self.txBody, value)

    def add_reply(self, reply_id: str, author_id: str, created: str) -> CT_ThreadedCommentReply:
        """Append and return a new ``<p188:reply>`` in this comment's thread.

        The ``<p188:replyLst>`` container is created on first reply
        (``ZeroOrOne``). Replies are flat children of the parent comment —
        the modern schema has no reply-to-reply nesting.
        """
        replyLst = self.get_or_add_replyLst()
        reply = cast(CT_ThreadedCommentReply, replyLst._add_reply())
        reply.id = reply_id
        reply.authorId = author_id
        reply.created = created
        reply.get_or_add_txBody()
        return reply


class CT_ThreadedCommentReplyList(BaseOxmlElement):
    """`p188:replyLst` element — ordered replies under a threaded comment."""

    reply_lst: list[CT_ThreadedCommentReply]

    reply = ZeroOrMore("p188:reply", successors=())


class CT_ThreadedCommentReply(BaseOxmlElement):
    """`p188:reply` element — one reply in a comment thread.

    Same attribute/body shape as the parent ``<p188:cm>`` minus the nested
    ``replyLst`` (replies are flat under their parent comment).
    """

    txBody = ZeroOrOne("p188:txBody", successors=("p188:extLst",))

    id: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "id", XsdString
    )
    authorId: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "authorId", XsdString
    )
    created: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "created", XsdString
    )
    status: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "status", XsdString
    )

    @property
    def text(self) -> str:
        """Plain-text body of this reply."""
        return _txBody_text(self.txBody)

    @text.setter
    def text(self, value: str) -> None:
        self.get_or_add_txBody()
        _set_txBody_text(self.txBody, value)


class CT_ThreadedCommentList(BaseOxmlElement):
    """`p188:cmLst` element — root of a modern threaded-comments part.

    Microsoft `2018/8/main` schema. Holds zero or more ``<p188:cm>``
    threaded comments. Registering this class (rather than letting the part
    round-trip as a plain element) is what gives the reopened
    |ModernCommentsPart| a working ``cm_lst`` accessor — without it a
    save→reopen cycle loses the comment-list API (issue #25 Wave 2).
    """

    cm_lst: list[CT_ThreadedComment]
    _add_cm: "callable"

    cm = ZeroOrMore("p188:cm", successors=())

    @classmethod
    def new(cls) -> CT_ThreadedCommentList:
        """Return a new empty ``<p188:cmLst>`` element."""
        return cast(CT_ThreadedCommentList, parse_xml("<p188:cmLst %s/>" % nsdecls("p188")))

    def add_comment(self, comment_id: str, author_id: str, created: str) -> CT_ThreadedComment:
        """Append and return a new ``<p188:cm>`` with an empty text body."""
        cm = cast(CT_ThreadedComment, self._add_cm())
        cm.id = comment_id
        cm.authorId = author_id
        cm.created = created
        cm.get_or_add_txBody()
        return cm

    def remove_comment(self, cm: CT_ThreadedComment) -> None:
        """Detach `cm` from this list, leaving a valid (possibly empty) list.

        Removing the last comment deliberately does NOT drop the part or its
        relationship — an empty ``<p188:cmLst/>`` is valid OOXML and keeps
        the slide rels intact (issue #25 Wave 2, ISC-21 anti-criterion).
        """
        self.remove(cm)
