"""User-facing wrapper around customXml data parts.

`CustomXmlParts` exposes the collection of `<ds:datastoreItem>`-tagged
arbitrary-XML parts attached to a presentation. The user-facing element type
is :class:`pptx.parts.custom_xml.CustomXmlPart` itself — there is no separate
facade. Loaded base `Part` instances (which arise because `CT.XML` is not
mapped to `CustomXmlPart` in `pptx/__init__.py` per plan §3.6) are upgraded
in-place by `_upgrade_to_custom_xml_part(...)` on first enumeration.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterable, Iterator, Literal, Sequence, Union, cast

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.oxml import parse_xml
from pptx.oxml.xmlchemy import BaseOxmlElement
from pptx.parts.custom_xml import CustomXmlPart, XmlPayload

if TYPE_CHECKING:
    from pptx.parts.presentation import PresentationPart


# Reserved name-prefix used to store user-assigned customXml part names as
# entries in the custom document properties part. The key is
# `{prefix}{datastore_item_id}` and the value is the user-assigned name.
NAME_PROPERTY_PREFIX = "_pptx_customxml_name_"

# Reserved namespace for the string-blob envelope written by `add_string_blob`.
# Read back through `read_string_blob` only — callers using `add(...)` directly
# should pick their own namespace, not this one.
BLOB_NAMESPACE = "urn:python-pptx:blob"


class CustomXmlParts(Sequence[CustomXmlPart]):
    """Collection of customXml data parts attached to the presentation.

    Iterates both presentation-scoped (`ppt/_rels/presentation.xml.rels`) and
    package-scoped (`/_rels/.rels`) `RT.CUSTOM_XML` relationships. Parts are
    deduplicated by identity — a single part related from both sources appears
    once.

    Lookup:

        prs.custom_xml_parts[0]                  # by index
        prs.custom_xml_parts["item3.xml"]        # by partname tail
        prs.custom_xml_parts.by_guid("{...}")    # by datastoreItem GUID
        prs.custom_xml_parts.by_name("provenance")  # by user-assigned name
    """

    def __init__(self, presentation_part: "PresentationPart"):
        self._presentation_part = presentation_part

    # -- Sequence-like protocol --------------------------------------------

    def __iter__(self) -> Iterator[CustomXmlPart]:
        return self._iter_parts()

    def __len__(self) -> int:
        return sum(1 for _ in self._iter_parts())

    def __getitem__(self, key):  # type: ignore[override]
        if isinstance(key, int):
            for i, part in enumerate(self._iter_parts()):
                if i == key:
                    return part
            raise IndexError("custom_xml_parts index out of range: %d" % key)
        if isinstance(key, str):
            for part in self._iter_parts():
                partname = str(part.partname)
                if partname == key or partname.endswith("/" + key):
                    return part
            raise KeyError("no custom_xml part with partname %r" % key)
        raise TypeError(
            "custom_xml_parts key must be int or str, got %s" % type(key).__name__
        )

    # -- Public lookups ----------------------------------------------------

    def by_guid(self, guid: str) -> CustomXmlPart | None:
        """Return the part whose `datastore_item_id` matches `guid`, or None.

        Match is case-insensitive and curly-brace-tolerant — `"{ABCD-...}"` and
        `"abcd-..."` both find the same part.
        """
        target = _normalize_guid(guid)
        for part in self._iter_parts():
            if _normalize_guid(part.datastore_item_id) == target:
                return part
        return None

    def by_name(self, name: str) -> CustomXmlPart | None:
        """Return the part previously added with `name=...`, or None.

        Names are stored as reserved entries in the custom document properties
        part keyed by datastore_item_id; this method reverse-resolves the name
        through that table.
        """
        if not isinstance(name, str):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError("name must be str, got %s" % type(name).__name__)
        cp_part = self._presentation_part.package.custom_properties_part
        for prop in cp_part._element.property_lst:
            if not prop.name.startswith(NAME_PROPERTY_PREFIX):
                continue
            if prop.value != name:
                continue
            guid = prop.name[len(NAME_PROPERTY_PREFIX) :]
            return self.by_guid(guid)
        return None

    # -- Mutation ----------------------------------------------------------

    def add(
        self,
        xml: XmlPayload,
        *,
        name: str | None = None,
        datastoreItem_id: str | None = None,
        schema_refs: Iterable[str] | None = None,
        scope: Literal["presentation", "package"] = "presentation",
    ) -> CustomXmlPart:
        """Add a new customXml part with `xml` as its payload.

        See module docstring for parameter semantics. Returns the new part,
        already attached to the presentation; nothing else is required before
        `prs.save(...)`.
        """
        if scope not in ("presentation", "package"):
            raise ValueError(
                "scope must be 'presentation' or 'package', got %r" % (scope,)
            )

        package = self._presentation_part.package
        data_part = CustomXmlPart.new_pair(
            package,
            xml,
            datastore_item_id=datastoreItem_id,
            schema_refs=tuple(schema_refs) if schema_refs is not None else (),
        )

        if scope == "presentation":
            self._presentation_part.relate_to(data_part, RT.CUSTOM_XML)
        else:
            package.relate_to(data_part, RT.CUSTOM_XML)

        if name is not None:
            cp = package.custom_properties
            cp[NAME_PROPERTY_PREFIX + data_part.datastore_item_id] = name

        return data_part

    def add_string_blob(
        self,
        name: str,
        content: str,
        *,
        mime_hint: str | None = None,
        encoding: Literal["text", "base64"] = "text",
        scope: Literal["presentation", "package"] = "presentation",
    ) -> CustomXmlPart:
        """Embed a string payload as a customXml part.

        Wraps `content` in a one-element XML envelope under the reserved
        `urn:python-pptx:blob` namespace::

            <blob xmlns="urn:python-pptx:blob"
                  name="…" mime="…" encoding="text|base64">…</blob>

        For binary or non-XML-safe text, set ``encoding="base64"`` and pass
        already-encoded `content` — the helper does NOT encode for you. Read
        back via :meth:`read_string_blob`.

        `mime_hint` is stored as the ``mime`` attribute on the envelope and
        round-trips for the caller's reference; it has no effect on PowerPoint.

        Returns the created :class:`CustomXmlPart`. Already attached at the
        chosen scope; nothing else is needed before ``prs.save(...)``.
        """
        if not isinstance(name, str) or not name:
            raise ValueError("name must be a non-empty string")
        if not isinstance(content, str):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError("content must be str, got %s" % type(content).__name__)
        if encoding not in ("text", "base64"):
            raise ValueError(
                "encoding must be 'text' or 'base64', got %r" % (encoding,)
            )

        from lxml import etree

        envelope = etree.Element("{%s}blob" % BLOB_NAMESPACE, nsmap={None: BLOB_NAMESPACE})
        envelope.set("name", name)
        envelope.set("encoding", encoding)
        if mime_hint is not None:
            envelope.set("mime", mime_hint)
        envelope.text = content

        return self.add(envelope, name=name, scope=scope)

    def read_string_blob(self, name: str) -> str | None:
        """Return the string payload of the blob part `name`, or `None`.

        Locates the part via :meth:`by_name`. Returns `None` if no such part
        exists or if the part is not a `urn:python-pptx:blob` envelope (i.e.
        was added by some other API or tool).

        For ``encoding="base64"`` blobs, the still-encoded string is returned
        — the caller decodes. The original encoding is recoverable from
        :meth:`blob_encoding`.
        """
        part = self.by_name(name)
        if part is None:
            return None
        root = part.element
        if root.tag != "{%s}blob" % BLOB_NAMESPACE:
            return None
        return root.text or ""

    def blob_encoding(self, name: str) -> str | None:
        """Return the `encoding` attribute of the blob part `name`, or `None`.

        Useful when a caller mixes text and base64 blobs and needs to decode
        the latter on read.
        """
        part = self.by_name(name)
        if part is None:
            return None
        root = part.element
        if root.tag != "{%s}blob" % BLOB_NAMESPACE:
            return None
        return root.get("encoding")

    def remove(self, part: Union[CustomXmlPart, int, str]) -> None:
        """Remove a customXml part from the presentation.

        Drops the relationship from whichever scope (presentation or package)
        currently holds it, plus any reserved name entry in custom_properties.
        Idempotent — a second call on an already-removed part is a no-op.

        The data → props rel is intentionally LEFT IN PLACE on the now-orphaned
        data part. Once the source rel is gone, neither the data part nor the
        props part is reachable from `iter_parts`, so both are omitted on
        save. Keeping the rel around lets a caller still read
        `part.datastore_item_id` on the returned reference after removal,
        which matches the principle of least surprise for held references.
        """
        target = self._resolve(part)
        if target is None:
            return

        # Drop the reserved name entry, if any. Reading datastore_item_id
        # here requires the data → props rel to still be intact.
        cp_part = self._presentation_part.package.custom_properties_part
        cp_part.remove_property(NAME_PROPERTY_PREFIX + target.datastore_item_id)

        # Drop the rel from whichever source holds it (presentation or package).
        for rels in self._iter_rel_collections():
            for rId, rel in list(rels.items()):
                if rel.is_external or rel.reltype != RT.CUSTOM_XML:
                    continue
                if rel.target_part is target:
                    rels.pop(rId)

    # -- Internals ---------------------------------------------------------

    def _iter_parts(self) -> Iterator[CustomXmlPart]:
        """Yield each unique customXml data part across both rel sources."""
        seen: set[int] = set()
        for rels in self._iter_rel_collections():
            for rel in rels.values():
                if rel.is_external or rel.reltype != RT.CUSTOM_XML:
                    continue
                part = _upgrade_to_custom_xml_part(rel.target_part)
                if id(part) in seen:
                    continue
                seen.add(id(part))
                yield part

    def _iter_rel_collections(self):
        """Yield the two relationship collections to scan for `RT.CUSTOM_XML`.

        Presentation part exposes `.rels` publicly; the package exposes the
        same collection internally as `_rels` (it has no public API for
        external rel inspection because most callers reach the rel graph via
        `iter_parts`/`iter_rels` instead). We need direct rel access here to
        find the source rel for `add(scope="package")`-attached parts.
        """
        yield self._presentation_part.rels
        yield self._presentation_part.package._rels

    def _resolve(
        self, part: Union[CustomXmlPart, int, str]
    ) -> CustomXmlPart | None:
        if isinstance(part, CustomXmlPart):
            return part
        if isinstance(part, int):
            try:
                return self.__getitem__(part)
            except IndexError:
                return None
        if isinstance(part, str):
            try:
                return self.__getitem__(part)
            except KeyError:
                return None
        raise TypeError(
            "remove() argument must be CustomXmlPart, int, or str; got %s"
            % type(part).__name__
        )


def _upgrade_to_custom_xml_part(part: Part) -> CustomXmlPart:
    """Upgrade a base `Part` to `CustomXmlPart` in-place via `__class__` swap.

    Loaded `application/xml` parts come in as plain `Part` because plan §3.6
    intentionally leaves `CT.XML` unmapped. On first enumeration, we promote
    the instance: assign the `CustomXmlPart` class, parse its blob to lxml,
    and stash the parsed root in `_element`. The package's rel graph keeps
    pointing at the same instance, so every other reference now resolves to
    the upgraded class with no graph rewriting.
    """
    if isinstance(part, CustomXmlPart):
        return part
    element = cast("BaseOxmlElement", parse_xml(part.blob))
    part.__class__ = CustomXmlPart
    part._element = element  # type: ignore[attr-defined]
    return cast(CustomXmlPart, part)


def _normalize_guid(guid: str) -> str:
    """Lowercase and strip surrounding curly braces for comparison."""
    if not isinstance(guid, str):  # pyright: ignore[reportUnnecessaryIsInstance]
        raise TypeError("guid must be str, got %s" % type(guid).__name__)
    s = guid.strip().lower()
    if s.startswith("{") and s.endswith("}"):
        s = s[1:-1]
    return s
