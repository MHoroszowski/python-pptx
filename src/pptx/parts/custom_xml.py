"""customXml data parts and their itemProps siblings.

Two part subclasses living together in this module because they are an atomic
pair — a `CustomXmlPart` is meaningless without its `CustomXmlPropertiesPart`
sibling, and vice versa. Both are created by `CustomXmlPart.new_pair(...)`.

Schema references: ECMA-376 Part 1, §15.2.4 (Custom XML Data Storage Part).
"""

from __future__ import annotations

import uuid
from typing import TYPE_CHECKING, Iterable, Union, cast

from lxml.etree import _Element  # pyright: ignore[reportPrivateUsage]

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.oxml.custom_xml import CT_DatastoreItem
from pptx.oxml.xmlchemy import BaseOxmlElement

if TYPE_CHECKING:
    from pptx.package import Package


XmlPayload = Union[bytes, str, _Element]


class CustomXmlPropertiesPart(XmlPart):
    """Corresponds to part named `/customXml/itemPropsN.xml`.

    Carries the `datastoreItem` GUID identifying its sibling `CustomXmlPart`
    across edits, plus the optional list of `<ds:schemaRef>` URIs the data part
    claims to conform to.
    """

    _element: CT_DatastoreItem

    @classmethod
    def new(
        cls,
        package: "Package",
        partname: PackURI,
        datastore_item_id: str,
        schema_refs: Iterable[str] = (),
    ) -> "CustomXmlPropertiesPart":
        """Return a fresh `CustomXmlPropertiesPart` at `partname` for `package`."""
        item_elm = CT_DatastoreItem.new(datastore_item_id, schema_refs=schema_refs)
        return cls(partname, CT.OFC_CUSTOM_XML_PROPERTIES, package, item_elm)

    @property
    def datastore_item_id(self) -> str:
        """The `ds:itemID` attribute — a GUID like `"{1A2B...}"`."""
        return self._element.itemID

    @datastore_item_id.setter
    def datastore_item_id(self, value: str) -> None:
        self._element.itemID = value

    @property
    def schema_refs(self) -> tuple[str, ...]:
        """Tuple of `<ds:schemaRef ds:uri>` values in document order."""
        return self._element.schema_ref_uris

    def add_schema_ref(self, uri: str) -> None:
        """Append a `<ds:schemaRef ds:uri="...">` (idempotent on `uri`)."""
        self._element.add_schema_ref(uri)

    def remove_schema_ref(self, uri: str) -> bool:
        """Remove the schemaRef with `uri`, returning True if found."""
        return self._element.remove_schema_ref(uri)


class CustomXmlPart(XmlPart):
    """Corresponds to part named `/customXml/itemN.xml`.

    Holds an arbitrary XML payload supplied by the caller. The payload's root
    element name and namespaces are caller-defined — `python-pptx` does not
    impose a schema. Each `CustomXmlPart` has a sibling `CustomXmlPropertiesPart`
    that carries the part's `datastoreItem` GUID; the rel between them is of
    type `RT.CUSTOM_XML_PROPS`.

    NOTE: This class is intentionally **not** registered with `PartFactory`
    against `CT.XML`. Loaded `application/xml` parts are produced as base
    `Part` instances, and the Phase-3 `CustomXmlParts` collection upgrades
    them on enumeration. See `Plans/customxml-implementation-plan.md` §3.6.
    """

    @classmethod
    def new_pair(
        cls,
        package: "Package",
        xml_payload: XmlPayload,
        *,
        datastore_item_id: str | None = None,
        schema_refs: Iterable[str] = (),
    ) -> "CustomXmlPart":
        """Create a paired CustomXmlPart + CustomXmlPropertiesPart in `package`.

        Returns the data part. The props part is related from the data part
        via `RT.CUSTOM_XML_PROPS`. Neither is yet related from any outside
        source — that is the caller's job (Phase-3 `CustomXmlParts.add(...)`).

        `xml_payload` may be `bytes`, a `str`, or an existing lxml `_Element`.
        If `datastore_item_id` is omitted a fresh `uuid4()` is generated and
        wrapped in curly braces to match Office's format.

        Partname allocation: `/customXml/itemN.xml` and `/customXml/itemPropsN.xml`
        share the same `N`, picked as the next free index across existing data
        parts in `package` (props parts are looked up via the data → props rel,
        not via partname pattern).
        """
        idx = _next_customxml_index(package)
        data_partname = PackURI("/customXml/item%d.xml" % idx)
        props_partname = PackURI("/customXml/itemProps%d.xml" % idx)

        element = _parse_payload(xml_payload)
        data_part = cls(data_partname, CT.XML, package, element)

        if datastore_item_id is None:
            datastore_item_id = "{%s}" % str(uuid.uuid4()).upper()

        props_part = CustomXmlPropertiesPart.new(
            package, props_partname, datastore_item_id, schema_refs
        )

        data_part.relate_to(props_part, RT.CUSTOM_XML_PROPS)
        return data_part

    @property
    def props_part(self) -> CustomXmlPropertiesPart:
        """Return the related `CustomXmlPropertiesPart` for this data part.

        Raises `KeyError` if the props rel is missing — a malformed package
        the caller is expected to repair via `CustomXmlPart.new_pair(...)`.
        """
        return cast(CustomXmlPropertiesPart, self.part_related_by(RT.CUSTOM_XML_PROPS))

    @property
    def datastore_item_id(self) -> str:
        """Convenience accessor delegating to the sibling props part."""
        return self.props_part.datastore_item_id

    @datastore_item_id.setter
    def datastore_item_id(self, value: str) -> None:
        self.props_part.datastore_item_id = value

    @property
    def schema_refs(self) -> tuple[str, ...]:
        """Convenience accessor delegating to the sibling props part."""
        return self.props_part.schema_refs

    def add_schema_ref(self, uri: str) -> None:
        """Convenience pass-through to the sibling props part."""
        self.props_part.add_schema_ref(uri)

    def remove_schema_ref(self, uri: str) -> bool:
        """Convenience pass-through to the sibling props part."""
        return self.props_part.remove_schema_ref(uri)

    @property
    def element(self) -> BaseOxmlElement:
        """Live root element of the customXml payload.

        Mutating its children mutates the part; the next `package.save(...)`
        will serialize the updated tree.
        """
        return self._element

    def replace_xml(self, xml_payload: XmlPayload) -> None:
        """Replace the entire XML payload with `xml_payload`.

        The root element is replaced wholesale; `datastore_item_id` and
        `schema_refs` are unaffected (they live on the sibling props part).
        """
        self._element = _parse_payload(xml_payload)

    @property
    def name(self) -> str | None:
        """The application-assigned name for this part, or `None`.

        Names are stored as reserved entries in `/docProps/custom.xml` keyed
        by `datastore_item_id`. See `Plans/customxml-implementation-plan.md`
        §3.4 for the rationale (Q3 default).
        """
        # Local import to avoid `parts → custom_xml → parts` cycle.
        from pptx.custom_xml import NAME_PROPERTY_PREFIX

        try:
            cp_part = self.package.custom_properties_part
        except Exception:  # pragma: no cover — package without custom_properties_part hook
            return None
        prop = cp_part.get_property(NAME_PROPERTY_PREFIX + self.datastore_item_id)
        if prop is None:
            return None
        value = prop.value
        return value if isinstance(value, str) else None

    def add_item(self, tag: str, text: str = "", **attrs: str) -> BaseOxmlElement:
        """Append a child element `<tag>text</tag>` with `attrs`.

        Convenience for the common "flat list of items" customXml shape; for
        arbitrary structure mutate :attr:`element` directly. The `tag` is
        used verbatim — pass a fully-namespaced Clark name if the parent
        root uses a default namespace and you need to escape it explicitly,
        otherwise lxml will attach the new element to the parent's namespace.

        Returns the newly appended element so the caller can chain further
        edits on it.
        """
        from lxml import etree

        new = etree.SubElement(self._element, tag)
        if text:
            new.text = text
        for k, v in attrs.items():
            new.set(k, v)
        return cast(BaseOxmlElement, new)


def _parse_payload(xml_payload: XmlPayload) -> BaseOxmlElement:
    """Coerce `xml_payload` to a `BaseOxmlElement` root.

    Accepts bytes (parsed verbatim), str (utf-8 encoded then parsed), or an
    already-parsed lxml `_Element` (returned as-is). Raises `TypeError` for
    anything else so the caller fails fast at the boundary.
    """
    if isinstance(xml_payload, bytes):
        return cast("BaseOxmlElement", parse_xml(xml_payload))
    if isinstance(xml_payload, str):
        return cast("BaseOxmlElement", parse_xml(xml_payload.encode("utf-8")))
    if isinstance(xml_payload, _Element):
        return cast("BaseOxmlElement", xml_payload)
    raise TypeError(
        "xml_payload must be bytes, str, or lxml _Element; got %s" % type(xml_payload).__name__
    )


def _next_customxml_index(package: "Package") -> int:
    """Return the next free `N` for `/customXml/itemN.xml`.

    Walks `package.iter_parts()` and skips `itemProps*.xml` parts. Reuses
    gaps in the sequence (e.g. if items 1 and 3 exist, returns 2).
    """
    used: set[int] = set()
    data_prefix = "/customXml/item"
    props_prefix = "/customXml/itemProps"
    for part in package.iter_parts():
        partname = str(part.partname)
        if not partname.startswith(data_prefix):
            continue
        if partname.startswith(props_prefix):
            continue
        # partname looks like /customXml/itemN.xml
        suffix = partname[len(data_prefix) :]
        if not suffix.endswith(".xml"):
            continue
        try:
            used.add(int(suffix[: -len(".xml")]))
        except ValueError:
            continue
    n = 1
    while n in used:
        n += 1
    return n
