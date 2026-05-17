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

    Existing ``a:p`` children are removed and one fresh
    ``<a:p><a:r><a:t>value</a:t></a:r></a:p>`` is appended after the
    (preserved) ``<a:bodyPr>``. This is the minimal round-trip-safe shape
    PowerPoint accepts for a threaded-comment body.
    """
    for p in list(txBody.findall(qn("a:p"))):
        txBody.remove(p)
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
    """

    txBody = ZeroOrOne("p188:txBody", successors=("p188:replyLst", "p188:extLst"))
    replyLst: "CT_ThreadedCommentReplyList | None" = ZeroOrOne(  # pyright: ignore
        "p188:replyLst", successors=("p188:extLst",)
    )

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
    # ---Wave-2 anchor: when a comment is anchored to a shape, the anchored
    #    shape's id is stored here. ``anchor_position`` (proxy layer) resolves
    #    it back to the shape's slide-space (left, top). Custom attribute on
    #    the p188:cm element; lxml round-trips unknown attributes verbatim.
    anchorShapeId: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "anchorShapeId", XsdUnsignedInt
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
