# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.custom_xml`."""

from __future__ import annotations

import re

import pytest
from lxml import etree

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.oxml.custom_xml import CT_DatastoreItem
from pptx.oxml.ns import nsdecls
from pptx.parts.custom_xml import (
    CustomXmlPart,
    CustomXmlPropertiesPart,
    _next_customxml_index,
    _parse_payload,
)

_GUID_A = "{1A2B3C4D-5E6F-7890-ABCD-EF1234567890}"
_GUID_B = "{ABCDEF12-3456-7890-ABCD-EF1234567890}"
_GUID_RE = re.compile(
    r"^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$"
)


class _StubPart:
    """Minimal stand-in for an existing package part during partname allocation tests."""

    def __init__(self, partname: str):
        self.partname = PackURI(partname)


class _StubPackage:
    """Minimal Package-like double exposing only `iter_parts()`.

    Sufficient because `CustomXmlPart.new_pair` and `_next_customxml_index`
    consult `iter_parts()` for partname allocation and never call any other
    method on the package during construction.
    """

    def __init__(self, partnames: list[str] | None = None):
        self._parts = [_StubPart(p) for p in (partnames or [])]

    def iter_parts(self):
        return iter(self._parts)


# ---------------------------------------------------------------------------
# CustomXmlPropertiesPart
# ---------------------------------------------------------------------------


def _datastore_xml(item_id: str, *uris: str) -> bytes:
    schema_refs = ""
    if uris:
        schema_refs = "<ds:schemaRefs>%s</ds:schemaRefs>" % "".join(
            '<ds:schemaRef ds:uri="%s"/>' % u for u in uris
        )
    return (
        '<ds:datastoreItem %s ds:itemID="%s">%s</ds:datastoreItem>'
        % (nsdecls("ds"), item_id, schema_refs)
    ).encode()


class DescribeCustomXmlPropertiesPart:
    def it_constructs_via_new(self):
        part = CustomXmlPropertiesPart.new(
            None,  # type: ignore[arg-type]
            PackURI("/customXml/itemProps1.xml"),
            _GUID_A,
            schema_refs=("urn:foo", "urn:bar"),
        )
        assert isinstance(part, CustomXmlPropertiesPart)
        assert part.content_type == CT.OFC_CUSTOM_XML_PROPERTIES
        assert part.partname == "/customXml/itemProps1.xml"
        assert isinstance(part._element, CT_DatastoreItem)
        assert part.datastore_item_id == _GUID_A
        assert part.schema_refs == ("urn:foo", "urn:bar")

    def it_loads_from_blob(self):
        part = CustomXmlPropertiesPart.load(
            "/customXml/itemProps1.xml",
            CT.OFC_CUSTOM_XML_PROPERTIES,
            None,  # type: ignore[arg-type]
            _datastore_xml(_GUID_A, "urn:x"),
        )
        assert part.datastore_item_id == _GUID_A
        assert part.schema_refs == ("urn:x",)

    def it_can_change_the_datastore_item_id(self):
        part = CustomXmlPropertiesPart.new(
            None, PackURI("/customXml/itemProps1.xml"), _GUID_A  # type: ignore[arg-type]
        )
        part.datastore_item_id = _GUID_B
        assert part.datastore_item_id == _GUID_B

    def it_adds_and_removes_schema_refs(self):
        part = CustomXmlPropertiesPart.new(
            None, PackURI("/customXml/itemProps1.xml"), _GUID_A  # type: ignore[arg-type]
        )
        part.add_schema_ref("urn:a")
        part.add_schema_ref("urn:b")
        assert part.schema_refs == ("urn:a", "urn:b")
        assert part.remove_schema_ref("urn:a") is True
        assert part.schema_refs == ("urn:b",)
        assert part.remove_schema_ref("urn:missing") is False


# ---------------------------------------------------------------------------
# CustomXmlPart
# ---------------------------------------------------------------------------


class DescribeCustomXmlPart_new_pair:
    def it_creates_paired_data_and_props_parts(self):
        pkg = _StubPackage()
        data = CustomXmlPart.new_pair(
            pkg,  # type: ignore[arg-type]
            b'<?xml version="1.0"?><provenance xmlns="urn:my:p"/>',
        )
        assert isinstance(data, CustomXmlPart)
        assert data.content_type == CT.XML
        assert data.partname == "/customXml/item1.xml"
        assert isinstance(data.props_part, CustomXmlPropertiesPart)
        assert data.props_part.partname == "/customXml/itemProps1.xml"

    def it_wires_the_props_relationship(self):
        pkg = _StubPackage()
        data = CustomXmlPart.new_pair(pkg, b"<x/>")  # type: ignore[arg-type]
        rels = list(data.rels.values())
        assert len(rels) == 1
        assert rels[0].reltype == RT.CUSTOM_XML_PROPS
        assert rels[0].target_part is data.props_part

    def it_auto_generates_a_curly_braced_guid_when_omitted(self):
        pkg = _StubPackage()
        data = CustomXmlPart.new_pair(pkg, b"<x/>")  # type: ignore[arg-type]
        assert _GUID_RE.match(data.datastore_item_id), data.datastore_item_id

    def it_accepts_a_caller_supplied_datastore_item_id(self):
        pkg = _StubPackage()
        data = CustomXmlPart.new_pair(
            pkg, b"<x/>", datastore_item_id=_GUID_A  # type: ignore[arg-type]
        )
        assert data.datastore_item_id == _GUID_A

    def it_propagates_schema_refs_to_props_part(self):
        pkg = _StubPackage()
        data = CustomXmlPart.new_pair(
            pkg,  # type: ignore[arg-type]
            b"<x/>",
            schema_refs=("urn:a", "urn:b"),
        )
        assert data.schema_refs == ("urn:a", "urn:b")
        assert data.props_part.schema_refs == ("urn:a", "urn:b")

    @pytest.mark.parametrize(
        "payload",
        [
            b'<?xml version="1.0"?><x xmlns="urn:p"/>',
            '<?xml version="1.0"?><x xmlns="urn:p"/>',
            etree.fromstring(b"<x/>"),
        ],
    )
    def it_accepts_payload_as_bytes_str_or_element(self, payload):
        pkg = _StubPackage()
        data = CustomXmlPart.new_pair(pkg, payload)  # type: ignore[arg-type]
        assert b"<x" in data.blob

    def it_raises_TypeError_for_other_payload_types(self):
        pkg = _StubPackage()
        with pytest.raises(TypeError):
            CustomXmlPart.new_pair(pkg, 123)  # type: ignore[arg-type]

    def it_picks_the_next_free_index_when_existing_parts_exist(self):
        pkg = _StubPackage(
            partnames=[
                "/customXml/item1.xml",
                "/customXml/itemProps1.xml",
                "/customXml/item2.xml",
                "/customXml/itemProps2.xml",
            ]
        )
        data = CustomXmlPart.new_pair(pkg, b"<x/>")  # type: ignore[arg-type]
        assert data.partname == "/customXml/item3.xml"
        assert data.props_part.partname == "/customXml/itemProps3.xml"

    def it_reuses_a_gap_in_the_index_sequence(self):
        pkg = _StubPackage(
            partnames=[
                "/customXml/item1.xml",
                "/customXml/itemProps1.xml",
                "/customXml/item3.xml",
                "/customXml/itemProps3.xml",
            ]
        )
        data = CustomXmlPart.new_pair(pkg, b"<x/>")  # type: ignore[arg-type]
        assert data.partname == "/customXml/item2.xml"


class DescribeCustomXmlPart_payload:
    def it_exposes_the_live_root_element(self):
        pkg = _StubPackage()
        data = CustomXmlPart.new_pair(  # type: ignore[arg-type]
            pkg, b'<provenance xmlns="urn:my:p"><source>cli</source></provenance>'
        )
        assert data.element.tag == "{urn:my:p}provenance"

    def it_round_trips_payload_through_blob(self):
        pkg = _StubPackage()
        original = b'<root xmlns="urn:my"><child a="1">hello</child></root>'
        data = CustomXmlPart.new_pair(pkg, original)  # type: ignore[arg-type]
        # blob is the same XML re-serialized
        reparsed = etree.fromstring(data.blob)
        assert reparsed.tag == "{urn:my}root"
        assert reparsed.find("{urn:my}child").text == "hello"

    def it_replaces_the_payload_via_replace_xml(self):
        pkg = _StubPackage()
        data = CustomXmlPart.new_pair(pkg, b"<old/>")  # type: ignore[arg-type]
        original_id = data.datastore_item_id
        data.replace_xml(b'<new xmlns="urn:n"/>')
        assert b"<new" in data.blob
        assert b"<old" not in data.blob
        # datastore_item_id is on the props part — unaffected by payload swap
        assert data.datastore_item_id == original_id


class DescribeHelpers:
    @pytest.mark.parametrize(
        ("partnames", "expected"),
        [
            ([], 1),
            (["/customXml/item1.xml", "/customXml/itemProps1.xml"], 2),
            (
                [
                    "/customXml/item1.xml",
                    "/customXml/itemProps1.xml",
                    "/customXml/item2.xml",
                    "/customXml/itemProps2.xml",
                ],
                3,
            ),
            (
                [
                    "/customXml/item1.xml",
                    "/customXml/item3.xml",
                    "/customXml/itemProps1.xml",
                    "/customXml/itemProps3.xml",
                ],
                2,
            ),
            # itemProps-only entries should NOT consume index slots
            (["/customXml/itemProps1.xml", "/customXml/itemProps2.xml"], 1),
            # unrelated partnames are ignored
            (["/ppt/slides/slide1.xml", "/customXml/item1.xml"], 2),
        ],
    )
    def it_picks_the_first_free_slot(self, partnames, expected):
        pkg = _StubPackage(partnames)
        assert _next_customxml_index(pkg) == expected  # type: ignore[arg-type]

    def it_skips_malformed_index_strings(self):
        pkg = _StubPackage(["/customXml/itemNOPE.xml"])
        assert _next_customxml_index(pkg) == 1  # type: ignore[arg-type]

    def it_parses_payload_bytes(self):
        elm = _parse_payload(b"<x/>")
        assert elm.tag == "x"

    def it_parses_payload_str(self):
        elm = _parse_payload("<x/>")
        assert elm.tag == "x"

    def it_returns_passed_element_unchanged(self):
        x = etree.fromstring(b"<x/>")
        assert _parse_payload(x) is x

    def it_raises_TypeError_for_unsupported_payload(self):
        with pytest.raises(TypeError):
            _parse_payload(123)  # type: ignore[arg-type]
