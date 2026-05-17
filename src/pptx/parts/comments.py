"""Part classes for presentation comments (issue #25).

Four parts spanning the two comment families in ``pptx/oxml/comments.py``.
The modern family needs a *separate* GUID-keyed authors part
(:class:`AuthorsPart`, ``/ppt/authors.xml``) — reusing the legacy
integer-keyed ``commentAuthors.xml`` for modern ``<p188:cm>`` author refs
left the GUID ``authorId`` dangling and triggered the PowerPoint repair
dialog. The legacy :class:`CommentAuthorsPart` stays for legacy ``<p:cm>``
comments only.

* :class:`CommentAuthorsPart` — the singleton ``/ppt/commentAuthors.xml``
  holding the legacy ``<p:cmAuthorLst>``. Content-type
  ``CT.PML_COMMENT_AUTHORS``; related from the presentation part via
  ``RT.COMMENT_AUTHORS``.

* :class:`CommentsPart` — a per-slide legacy comments part wrapping
  ``<p:cmLst>`` (content-type ``CT.PML_COMMENTS``). PowerPoint names these
  ``/ppt/comments/commentN.xml`` (older builds used
  ``/ppt/comments/comment1.xml`` directly under the slide rel).

* :class:`ModernCommentsPart` — the post-2018 threaded comments part.

  Partname choice: PowerPoint emits ``/ppt/comments/modernComment_<slide>.xml``
  (one threaded-comments part per slide, the slide stem embedded in the
  filename). `python-pptx` cannot know the human slide name at part-creation
  time, so the public ``new()`` takes an explicit ``PackURI`` and Wave-2
  wiring will allocate ``/ppt/comments/modernComment_slideN.xml`` keyed off
  the owning slide's partname index. Content-type
  ``CT.PML_THREADED_COMMENTS`` (``application/vnd.ms-powerpoint.threadedComments+xml``)
  and relationship ``RT.THREADED_COMMENT`` are the vendor identifiers
  PowerPoint stamps — there is no ECMA-376 equivalent because threaded
  comments are a Microsoft extension, not part of the base standard.

No upstream `scanny/python-pptx` prior art; built from spec following the
``pptx/parts/coreprops.py`` / ``pptx/parts/slide.py`` part conventions.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.comments import (
    CT_Author,
    CT_AuthorList,
    CT_Comment,
    CT_CommentAuthor,
    CT_CommentAuthorList,
    CT_CommentList,
    CT_ThreadedComment,
    CT_ThreadedCommentList,
)

if TYPE_CHECKING:
    from pptx.package import Package


class CommentAuthorsPart(XmlPart):
    """The presentation-wide legacy comment-authors part.

    Corresponds to ``/ppt/commentAuthors.xml``. There is at most one per
    package; it is related from the presentation part with
    ``RT.COMMENT_AUTHORS``.
    """

    _element: CT_CommentAuthorList

    @classmethod
    def new(cls, package: "Package") -> "CommentAuthorsPart":
        """Return a new empty |CommentAuthorsPart| for `package`."""
        return cls(
            PackURI("/ppt/commentAuthors.xml"),
            CT.PML_COMMENT_AUTHORS,
            package,
            CT_CommentAuthorList.new(),
        )

    def add_author(self, name: str, initials: str = "") -> CT_CommentAuthor:
        """Add a ``<p:cmAuthor>`` for `name`, allocating the next free id."""
        return self._element.add_author(name, initials)

    def get_or_add_author(self, name: str, initials: str = "") -> CT_CommentAuthor:
        """Return the existing same-`name` author, else add a new one (ISC-5)."""
        return self._element.get_or_add_author(name, initials)

    def iter_authors(self) -> Iterator[CT_CommentAuthor]:
        """Generate each ``<p:cmAuthor>`` element in document order."""
        yield from self._element.cmAuthor_lst


class AuthorsPart(XmlPart):
    """The presentation-wide MODERN threaded-comment authors part.

    Corresponds to ``/ppt/authors.xml`` (content-type ``CT.PML_AUTHORS``,
    ``application/vnd.ms-powerpoint.authors+xml``). There is at most one per
    package; it is related from the **presentation** part with ``RT.AUTHORS``
    (``…/office/2018/10/relationships/authors``).

    This is the part the modern ``<p188:cm>/@authorId`` GUID resolves into.
    It is *distinct* from the legacy |CommentAuthorsPart|
    (``/ppt/commentAuthors.xml``, integer ids) — that part stays for legacy
    ``<p:cm>`` comments only. A modern-only deck must carry *this* part and
    must NOT carry an orphaned ``commentAuthors.xml`` (issue #25 repair-dialog
    root cause: a ``<p188:cm>`` GUID with no ``<p188:author>`` to resolve to).

    No upstream `scanny/python-pptx` prior art; built from the [MS-PPTX]
    threaded-comments extension following this fork's part conventions.
    """

    _element: CT_AuthorList

    @classmethod
    def new(cls, package: "Package") -> "AuthorsPart":
        """Return a new empty |AuthorsPart| for `package`."""
        return cls(
            PackURI("/ppt/authors.xml"),
            CT.PML_AUTHORS,
            package,
            CT_AuthorList.new(),
        )

    def get_or_add_author(self, name: str, initials: str = "") -> CT_Author:
        """Return the existing same-`name` author, else add a new one (ISC-5).

        The returned author's ``id`` is the deterministic GUID the modern
        ``<p188:cm>``/``<p188:reply>`` ``authorId`` must equal.
        """
        return self._element.get_or_add_author(name, initials)

    def iter_authors(self) -> Iterator[CT_Author]:
        """Generate each ``<p188:author>`` element in document order."""
        yield from self._element.author_lst


class CommentsPart(XmlPart):
    """A legacy per-slide comments part wrapping ``<p:cmLst>``.

    Content-type ``CT.PML_COMMENTS``; one per slide that has legacy
    comments, related from the slide via ``RT.COMMENTS``.
    """

    _element: CT_CommentList

    @classmethod
    def new(cls, package: "Package", partname: PackURI) -> "CommentsPart":
        """Return a new empty |CommentsPart| at `partname` for `package`."""
        return cls(partname, CT.PML_COMMENTS, package, CT_CommentList.new())

    def iter_comments(self) -> Iterator[CT_Comment]:
        """Generate each legacy ``<p:cm>`` element in document order.

        Read-only — python-pptx never writes legacy comments; this exists so
        a deck authored elsewhere keeps its legacy feedback visible through
        ``slide.comments`` (issue #25 Wave 3, SF8 / ISC-44, ISC-67).
        """
        yield from self._element.cm_lst


class ModernCommentsPart(XmlPart):
    """A modern (2018) threaded-comments part.

    Root element is ``<p188:cmLst>`` in the
    ``http://schemas.microsoft.com/office/powerpoint/2018/8/main``
    namespace, holding zero or more ``<p188:cm>`` threaded comments.
    Content-type ``CT.PML_THREADED_COMMENTS``; related from the slide via
    ``RT.THREADED_COMMENT``.
    """

    _element: CT_ThreadedCommentList

    @classmethod
    def new(cls, package: "Package", partname: PackURI) -> "ModernCommentsPart":
        """Return a new empty |ModernCommentsPart| at `partname`.

        The body is an empty ``<p188:cmLst>`` (the modern container element,
        a registered |CT_ThreadedCommentList| so the part keeps a working
        ``cm_lst`` accessor across a save→reopen cycle) — Wave 2 appends
        ``<p188:cm>`` children when a slide gains threaded comments.
        """
        return cls(
            partname,
            CT.PML_THREADED_COMMENTS,
            package,
            CT_ThreadedCommentList.new(),
        )

    def add_comment(
        self, comment_id: str, author_id: str, created: str, text: str
    ) -> CT_ThreadedComment:
        """Append a ``<p188:cm>`` carrying `text` and return its element."""
        cm = self._element.add_comment(comment_id, author_id, created)
        cm.text = text
        return cm

    def remove_comment(self, cm: CT_ThreadedComment) -> None:
        """Detach `cm`; an empty ``<p188:cmLst/>`` remains valid (ISC-21)."""
        self._element.remove_comment(cm)

    def iter_comments(self) -> Iterator[CT_ThreadedComment]:
        """Generate each ``<p188:cm>`` element in document order."""
        yield from self._element.cm_lst
