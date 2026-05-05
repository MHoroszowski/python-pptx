"""lxml custom element classes for customXml itemProps parts.

Models the `<ds:datastoreItem>` root of `/customXml/itemPropsN.xml` — the
sibling part of each `/customXml/itemN.xml` data part. Carries the
`datastoreItem` GUID identifying the data part across edits and the optional
`<ds:schemaRefs>` list declaring the XML namespaces the data part claims to
conform to.

Schema references: ECMA-376 Part 1, §15.2.4 (Custom XML Data Storage Part).
"""

from __future__ import annotations

from typing import Iterable, cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_DatastoreItem(BaseOxmlElement):
    """`<ds:datastoreItem>` element — root of `/customXml/itemPropsN.xml`."""

    itemID: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "ds:itemID", XsdString
    )
    schemaRefs = ZeroOrOne("ds:schemaRefs", successors=())

    _datastoreItem_tmpl = (
        '<ds:datastoreItem %s ds:itemID="%%s"/>\n' % nsdecls("ds")
    )

    @staticmethod
    def new(itemID: str, schema_refs: Iterable[str] = ()) -> "CT_DatastoreItem":
        """Return a new `<ds:datastoreItem>` with `itemID` and optional schema_refs.

        `itemID` should be a curly-braced GUID string, e.g.
        `"{1A2B3C4D-5E6F-7890-ABCD-EF1234567890}"`. The caller is responsible
        for generating it (typically via `uuid.uuid4()`); this layer does not
        validate the GUID format because Office tolerates non-canonical forms.
        """
        elm = cast(
            "CT_DatastoreItem",
            parse_xml(CT_DatastoreItem._datastoreItem_tmpl % itemID),
        )
        for uri in schema_refs:
            elm.add_schema_ref(uri)
        return elm

    def add_schema_ref(self, uri: str) -> "CT_DatastoreSchemaRef":
        """Add a `<ds:schemaRef ds:uri="..."/>` child.

        Creates the parent `<ds:schemaRefs>` element if it is not already
        present. If a schemaRef with `uri` already exists, returns the existing
        one rather than adding a duplicate.
        """
        refs = cast("CT_DatastoreSchemaRefs", self.get_or_add_schemaRefs())
        existing = refs.find_by_uri(uri)
        if existing is not None:
            return existing
        ref = cast("CT_DatastoreSchemaRef", refs._add_schemaRef())
        ref.uri = uri
        return ref

    def remove_schema_ref(self, uri: str) -> bool:
        """Remove the schemaRef with `uri`, returning True if found.

        If removing the last schemaRef leaves `<ds:schemaRefs>` empty, the
        empty parent element is also removed (Office writes the file this way
        — no empty `<ds:schemaRefs/>` envelope).
        """
        refs = cast("CT_DatastoreSchemaRefs | None", self.schemaRefs)
        if refs is None:
            return False
        ref = refs.find_by_uri(uri)
        if ref is None:
            return False
        refs.remove(ref)
        if len(refs.schemaRef_lst) == 0:
            self.remove(refs)
        return True

    @property
    def schema_ref_uris(self) -> tuple[str, ...]:
        """Tuple of `ds:uri` values for every `<ds:schemaRef>`, in document order."""
        refs = cast("CT_DatastoreSchemaRefs | None", self.schemaRefs)
        if refs is None:
            return ()
        return tuple(
            cast("CT_DatastoreSchemaRef", r).uri
            for r in cast("list[BaseOxmlElement]", refs.schemaRef_lst)
        )


class CT_DatastoreSchemaRefs(BaseOxmlElement):
    """`<ds:schemaRefs>` — collection of `<ds:schemaRef>` children."""

    schemaRef = ZeroOrMore("ds:schemaRef", successors=())

    def find_by_uri(self, uri: str) -> "CT_DatastoreSchemaRef | None":
        """Return the `<ds:schemaRef>` child whose `ds:uri` is `uri`, or None."""
        for ref in cast("list[CT_DatastoreSchemaRef]", self.schemaRef_lst):
            if ref.uri == uri:
                return ref
        return None


class CT_DatastoreSchemaRef(BaseOxmlElement):
    """`<ds:schemaRef>` — a single XML namespace this customXml part conforms to."""

    uri: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "ds:uri", XsdString
    )
