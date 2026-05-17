"""Public proxy objects for modern threaded comments (issue #25, Wave 2).

These are the API surface a caller actually touches:

* :class:`Comments` — the per-slide collection returned by
  ``Slide.comments``. Iterable in document order, ``len()``-able, with
  ``add(text, author, anchor=None)`` and ``remove(comment)``.
* :class:`Comment` — one threaded comment. Exposes ``text``, ``author``
  (resolved authorId→name), ``created_at`` (tz-aware ``datetime`` or
  ``None``), ``anchor_position`` ((x, y) |Length| pair or ``None``), and
  a ``replies`` collection.
* :class:`CommentReplies` / :class:`CommentReply` — the reply thread under
  a comment and one reply in it.

The collection lazily creates+relates the per-slide |ModernCommentsPart|
on first ``add`` (mirroring ``SlidePart.notes_slide``), and get-or-adds
the presentation-level |CommentAuthorsPart| author by name with no
duplicates (ISC-5). Removing the last comment leaves a valid empty
``<p188:cmLst/>`` and intact slide rels (ISC-21 anti-criterion).

No upstream `scanny/python-pptx` prior art — built from the OOXML /
MS-OOXML schemas and this fork's part conventions.
"""

from __future__ import annotations

import uuid
from datetime import datetime
from typing import TYPE_CHECKING, Iterator

from pptx.util import Emu, Length

if TYPE_CHECKING:
    from pptx.oxml.comments import (
        CT_Comment,
        CT_ThreadedComment,
        CT_ThreadedCommentReply,
    )
    from pptx.slide import Slide

# Modern threaded-comment resolution marker. The MS 2018/8/main schema gives
# `<p188:cm>` an optional ``status`` whose enumeration includes ``resolved``;
# PowerPoint stamps exactly this when a reviewer closes a thread. Absence (or
# any non-``resolved`` value) means the thread is still open.
_RESOLVED_STATUS = "resolved"


def _parse_created(raw: str | None) -> datetime | None:
    """Parse an OOXML ``created`` timestamp into a tz-aware ``datetime``.

    PowerPoint writes ISO-8601 (commonly ``2026-05-16T12:00:00.000`` or
    with a ``Z``/offset). Returns a timezone-aware ``datetime`` (assuming
    UTC when no offset is present) or ``None`` when the value is missing
    or unparseable — never raises, so a malformed file still reads.
    """
    if not raw:
        return None
    text = raw.strip()
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(text)
    except ValueError:
        return None
    if dt.tzinfo is None:
        from datetime import timezone

        dt = dt.replace(tzinfo=timezone.utc)
    return dt


def _now_iso() -> str:
    """Return the current UTC time as the ISO string PowerPoint expects.

    PowerPoint validates ``<p188:cm>/@created`` as ``xsd:dateTime``; the
    modern threaded-comment schema wants an explicit UTC designator. A
    naive ``...sss`` value with no offset (the previous behavior) is one of
    the things that triggered the issue-#25 repair dialog, so emit
    ``YYYY-MM-DDThh:mm:ssZ`` (UTC, ``Z``-suffixed, second precision).
    """
    from datetime import timezone

    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _new_guid() -> str:
    """Return a brace-wrapped uppercase GUID like PowerPoint emits."""
    return "{%s}" % str(uuid.uuid4()).upper()


class Comment:
    """One comment on a slide — modern threaded *or* legacy.

    python-pptx only ever *writes* the modern (2018) ``<p188:cm>`` form, but
    a deck authored elsewhere may carry pre-2018 legacy ``<p:cm>`` comments.
    Both are surfaced through the same proxy so callers iterating
    ``slide.comments`` see one uniform collection (issue #25 Wave 3, SF8 /
    ISC-44, ISC-67). A legacy-backed instance is read-only: it has no
    resolution concept and no reply thread, and :meth:`resolve` raises.
    :attr:`is_legacy` distinguishes the two.
    """

    def __init__(self, cm: "CT_ThreadedComment | CT_Comment", slide: "Slide"):
        self._cm = cm
        self._slide = slide

    def __eq__(self, other: object) -> bool:
        return isinstance(other, Comment) and other._cm is self._cm

    def __ne__(self, other: object) -> bool:
        return not self.__eq__(other)

    @property
    def is_legacy(self) -> bool:
        """`True` when this comment is a pre-2018 legacy ``<p:cm>``.

        Legacy-backed comments are read-only proxies: :meth:`resolve` raises
        and :attr:`replies` is always empty (the legacy schema has neither a
        resolution flag nor a reply thread).
        """
        from pptx.oxml.comments import CT_Comment as _CT_Comment

        return isinstance(self._cm, _CT_Comment)

    @property
    def text(self) -> str:
        """Plain-text body of this comment."""
        return self._cm.text

    @text.setter
    def text(self, value: str) -> None:
        if self.is_legacy:
            raise TypeError(
                "cannot edit a legacy <p:cm> comment — python-pptx only "
                "writes modern threaded comments; legacy comments are "
                "read-only (issue #25 SF8)"
            )
        self._cm.text = value

    @property
    def author(self) -> str | None:
        """Display name of this comment's author, resolved authorId→name.

        Modern comments key the author by deterministic GUID; legacy
        comments key by the integer ``authorId`` into the presentation
        ``<p:cmAuthorLst>``. Returns ``None`` when the id has no matching
        author (a malformed or externally-edited file).
        """
        if self.is_legacy:
            return self._slide._resolve_legacy_author_name(self._cm.authorId)
        return self._slide._resolve_author_name(self._cm.authorId)

    @property
    def created_at(self) -> datetime | None:
        """Creation timestamp as a tz-aware ``datetime``, or ``None``.

        Reads ``created`` for modern comments and ``dt`` for legacy ones.
        """
        raw = self._cm.dt if self.is_legacy else self._cm.created
        return _parse_created(raw)

    @property
    def resolved(self) -> bool:
        """`True` when this thread has been marked resolved.

        Defaults to ``False``. Legacy comments are *always* ``False`` — the
        pre-2018 schema has no resolution concept (ISC-37).
        """
        if self.is_legacy:
            return False
        return self._cm.status == _RESOLVED_STATUS

    def resolve(self) -> None:
        """Mark this threaded comment's conversation resolved (ISC-33..35).

        Sets the modern ``<p188:cm>`` ``status="resolved"`` marker; the flag
        round-trips through save/reopen and is read back by :attr:`resolved`.

        :raises TypeError: when called on a comment backed by a legacy
            ``<p:cm>`` element. **Design decision (ISC-37):** legacy raises
            rather than being a silent no-op. The pre-2018 PresentationML
            comment schema has no resolution attribute, so there is nowhere
            to persist the state. A silent no-op would let a caller believe
            ``resolve()`` succeeded and the thread is closed when, after
            save/reopen, ``resolved`` is still ``False`` — a correctness lie.
            A loud ``TypeError`` makes the format limitation explicit at the
            call site instead of hiding data loss.
        """
        if self.is_legacy:
            raise TypeError(
                "cannot resolve a legacy <p:cm> comment: the pre-2018 "
                "PresentationML comment schema has no resolution attribute, "
                "so the state cannot be persisted. Resolution is only "
                "supported on modern threaded comments (issue #25 SF6)."
            )
        self._cm.status = _RESOLVED_STATUS

    def reopen(self) -> None:
        """Clear the resolved marker, reopening this threaded comment.

        The inverse of :meth:`resolve`. Raises the same |TypeError| on a
        legacy-backed comment, for the same reason.
        """
        if self.is_legacy:
            raise TypeError(
                "cannot reopen a legacy <p:cm> comment: legacy comments "
                "have no resolution state (issue #25 SF6)."
            )
        self._cm.status = None

    @property
    def anchor_position(self) -> tuple[Length, Length] | None:
        """The (x, y) slide-space anchor of this comment, or ``None``.

        For a modern comment created with an ``anchor`` shape, the shape's
        id is stored on the ``<p188:cm>`` and this resolves it back to the
        anchored shape's ``(left, top)`` as a |Length| pair. Legacy
        comments carry an absolute ``<p:pos x= y=>`` point instead. Returns
        ``None`` for an unanchored comment or when the shape is gone.
        """
        if self.is_legacy:
            pos = self._cm.pos
            if pos is None:
                return None
            return (Emu(int(pos.x)), Emu(int(pos.y)))
        shape_id = self._cm.anchorShapeId
        if shape_id is None:
            return None
        for shape in self._slide.shapes:
            if shape.shape_id == shape_id:
                left = shape.left if shape.left is not None else Emu(0)
                top = shape.top if shape.top is not None else Emu(0)
                return (Emu(int(left)), Emu(int(top)))
        return None

    @property
    def _anchor_shape_id(self) -> int | None:
        """The anchored shape id for a modern comment, else ``None``.

        Internal helper for the per-shape filter (:attr:`BaseShape.comments`,
        SF7). Legacy comments anchor to an absolute point, never a shape, so
        they never participate in the per-shape filter.
        """
        if self.is_legacy:
            return None
        return self._cm.anchorShapeId

    @property
    def replies(self) -> "CommentReplies":
        """The ordered reply thread under this comment.

        Always empty for a legacy comment — the pre-2018 schema has no
        reply thread (SF4 documented legacy behavior).
        """
        return CommentReplies(self._cm, self._slide)


class CommentReply:
    """One reply in a threaded comment's thread."""

    def __init__(self, reply: "CT_ThreadedCommentReply", slide: "Slide"):
        self._reply = reply
        self._slide = slide

    def __eq__(self, other: object) -> bool:
        return isinstance(other, CommentReply) and other._reply is self._reply

    def __ne__(self, other: object) -> bool:
        return not self.__eq__(other)

    @property
    def text(self) -> str:
        """Plain-text body of this reply."""
        return self._reply.text

    @property
    def author(self) -> str | None:
        """Display name of this reply's author, resolved authorId→name."""
        return self._slide._resolve_author_name(self._reply.authorId)

    @property
    def created_at(self) -> datetime | None:
        """Creation timestamp as a tz-aware ``datetime``, or ``None``."""
        return _parse_created(self._reply.created)


class CommentReplies:
    """The ordered reply thread under a single threaded comment.

    A legacy ``<p:cm>`` has no reply container, so against a legacy-backed
    comment this collection is always empty and :meth:`add` raises — the
    pre-2018 schema cannot persist a reply thread (issue #25 SF8, the SF4
    documented legacy behavior).
    """

    def __init__(self, cm: "CT_ThreadedComment | CT_Comment", slide: "Slide"):
        self._cm = cm
        self._slide = slide

    @property
    def _is_legacy(self) -> bool:
        from pptx.oxml.comments import CT_Comment as _CT_Comment

        return isinstance(self._cm, _CT_Comment)

    def __len__(self) -> int:
        if self._is_legacy:
            return 0
        replyLst = self._cm.replyLst
        return 0 if replyLst is None else len(replyLst.reply_lst)

    def __iter__(self) -> Iterator[CommentReply]:
        if self._is_legacy:
            return
        replyLst = self._cm.replyLst
        if replyLst is None:
            return
        for reply in replyLst.reply_lst:
            yield CommentReply(reply, self._slide)

    def __getitem__(self, idx: int) -> CommentReply:
        return list(self)[idx]

    def add(self, text: str, author: str) -> CommentReply:
        """Append a reply carrying `text` by `author`; return it.

        Threads under the parent comment, preserves reply order, and never
        detaches the parent comment from its list (ISC-27 anti-criterion):
        the parent ``<p188:cm>`` is mutated in place, never re-parented.

        :raises TypeError: on a legacy-backed comment — the pre-2018 schema
            has no reply thread to append to (issue #25 SF8).
        """
        if self._is_legacy:
            raise TypeError(
                "cannot reply to a legacy <p:cm> comment: the pre-2018 "
                "schema has no reply thread. Replies are only supported on "
                "modern threaded comments (issue #25 SF4/SF8)."
            )
        author_guid = self._slide._get_or_add_author_guid(author)
        reply = self._cm.add_reply(_new_guid(), author_guid, _now_iso())
        reply.text = text
        return CommentReply(reply, self._slide)


class Comments:
    """Per-slide collection of comments — legacy *and* modern.

    Returned by :attr:`Slide.comments`. Iterating yields |Comment| objects:
    first any pre-2018 legacy ``<p:cm>`` comments the source deck carried
    (read-only proxies), then the modern ``<p188:cm>`` threaded comments, in
    document order. Reading both families is what makes legacy↔modern
    coexistence actually hold (issue #25 Wave 3, SF8 / ISC-44, ISC-67):
    adding a modern comment never deletes or rewrites the legacy part, and a
    save→reopen surfaces both. The backing |ModernCommentsPart| is created
    and related to the slide lazily on the first :meth:`add`; the legacy
    part is *never* created — python-pptx only ever writes modern comments.
    """

    def __init__(self, slide: "Slide"):
        self._slide = slide

    def __len__(self) -> int:
        return sum(1 for _ in self)

    def __iter__(self) -> Iterator[Comment]:
        # ---legacy <p:cm> first (older model, read-only). Never creates the
        #    part — it only exists when the source deck already had one.---
        legacy_part = self._slide.part.legacy_comments_part_if_present
        if legacy_part is not None:
            for cm in legacy_part.iter_comments():
                yield Comment(cm, self._slide)
        # ---then modern <p188:cm> threaded comments---
        part = self._slide._modern_comments_part_if_present
        if part is None:
            return
        for cm in part.iter_comments():
            yield Comment(cm, self._slide)

    def __getitem__(self, idx: int) -> Comment:
        return list(self)[idx]

    def add(self, text: str, author: str, anchor=None) -> Comment:
        """Add a threaded comment carrying `text` by `author`; return it.

        The first call lazily creates and relates the per-slide
        |ModernCommentsPart| (mirrors ``SlidePart.notes_slide``). The
        author is get-or-added on the presentation-level
        |CommentAuthorsPart| by name with no duplicates (ISC-5). When
        `anchor` is a shape, its id is recorded so
        :attr:`Comment.anchor_position` can resolve it back later.
        """
        part = self._slide.part.modern_comments_part
        author_guid = self._slide._get_or_add_author_guid(author)
        cm = part.add_comment(_new_guid(), author_guid, _now_iso(), text)
        # GROUND TRUTH (2026-05-17): PowerPoint binds a modern comment to its
        # slide via a `<pc:sldMkLst>` carrying the slide's `<p:sldId>/@id`.
        # Without it the comment never renders in the Comments pane (issue
        # #25 root cause). slide_id is the integer sldId from presentation.xml.
        cm.set_slide_marker(self._slide.slide_id)
        if anchor is not None and getattr(anchor, "shape_id", None) is not None:
            cm.anchorShapeId = anchor.shape_id
        return Comment(cm, self._slide)

    def remove(self, comment: Comment) -> None:
        """Detach `comment` from this slide's threaded-comments part.

        Removing the last comment deliberately leaves a valid empty
        ``<p188:cmLst/>`` and keeps the slide↔part relationship intact
        (ISC-21 anti-criterion) so the file stays openable.
        """
        part = self._slide._modern_comments_part_if_present
        if part is None:
            return
        part.remove_comment(comment._cm)
