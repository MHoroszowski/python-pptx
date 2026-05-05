# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.custom_xml`."""

from __future__ import annotations

import pytest
from lxml import etree

from pptx.oxml import parse_xml
from pptx.oxml.custom_xml import (
    CT_DatastoreItem,
    CT_DatastoreSchemaRef,
    CT_DatastoreSchemaRefs,
)
from pptx.oxml.ns import nsdecls

_GUID_A = "{1A2B3C4D-5E6F-7890-ABCD-EF1234567890}"
_GUID_B = "{ABCDEF12-3456-7890-ABCD-EF1234567890}"


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


class DescribeCT_DatastoreItem:
    def it_parses_to_the_registered_class(self):
        root = parse_xml(_datastore_xml(_GUID_A))
        assert isinstance(root, CT_DatastoreItem)

    def it_exposes_the_itemID_attribute(self):
        root = parse_xml(_datastore_xml(_GUID_A))
        assert root.itemID == _GUID_A

    def it_can_change_the_itemID_attribute(self):
        root = parse_xml(_datastore_xml(_GUID_A))
        root.itemID = _GUID_B
        assert root.itemID == _GUID_B

    def it_returns_an_empty_tuple_when_no_schemaRefs_present(self):
        root = parse_xml(_datastore_xml(_GUID_A))
        assert root.schemaRefs is None
        assert root.schema_ref_uris == ()

    def it_lists_schema_ref_uris_in_document_order(self):
        root = parse_xml(_datastore_xml(_GUID_A, "urn:a", "urn:b", "urn:c"))
        assert root.schema_ref_uris == ("urn:a", "urn:b", "urn:c")

    def it_creates_a_fresh_root_via_new(self):
        elm = CT_DatastoreItem.new(_GUID_A)
        assert elm.itemID == _GUID_A
        assert elm.schema_ref_uris == ()

    def it_creates_a_fresh_root_with_initial_schema_refs(self):
        elm = CT_DatastoreItem.new(_GUID_B, schema_refs=["urn:foo", "urn:bar"])
        assert elm.schema_ref_uris == ("urn:foo", "urn:bar")

    def it_adds_a_schema_ref_creating_the_envelope_when_absent(self):
        elm = CT_DatastoreItem.new(_GUID_A)
        elm.add_schema_ref("urn:foo")
        assert isinstance(elm.schemaRefs, CT_DatastoreSchemaRefs)
        assert elm.schema_ref_uris == ("urn:foo",)

    def it_returns_existing_ref_on_duplicate_add(self):
        elm = CT_DatastoreItem.new(_GUID_A, schema_refs=["urn:foo"])
        first = elm.add_schema_ref("urn:foo")
        second = elm.add_schema_ref("urn:foo")
        assert first is second
        assert elm.schema_ref_uris == ("urn:foo",)

    def it_removes_a_schema_ref_by_uri(self):
        elm = CT_DatastoreItem.new(_GUID_A, schema_refs=["urn:a", "urn:b"])
        assert elm.remove_schema_ref("urn:a") is True
        assert elm.schema_ref_uris == ("urn:b",)

    def it_returns_False_when_removing_nonexistent_ref(self):
        elm = CT_DatastoreItem.new(_GUID_A, schema_refs=["urn:a"])
        assert elm.remove_schema_ref("urn:missing") is False

    def it_drops_the_envelope_when_the_last_ref_is_removed(self):
        elm = CT_DatastoreItem.new(_GUID_A, schema_refs=["urn:only"])
        assert elm.remove_schema_ref("urn:only") is True
        assert elm.schemaRefs is None
        assert elm.schema_ref_uris == ()

    def it_round_trips_through_parse_serialize(self):
        elm = CT_DatastoreItem.new(_GUID_A, schema_refs=["urn:x", "urn:y"])
        serialized = etree.tostring(elm)
        reparsed = parse_xml(serialized)
        assert isinstance(reparsed, CT_DatastoreItem)
        assert reparsed.itemID == _GUID_A
        assert reparsed.schema_ref_uris == ("urn:x", "urn:y")


class DescribeCT_DatastoreSchemaRef:
    def it_parses_to_the_registered_class(self):
        root = parse_xml(_datastore_xml(_GUID_A, "urn:foo"))
        ref = root.schemaRefs.schemaRef_lst[0]
        assert isinstance(ref, CT_DatastoreSchemaRef)

    def it_exposes_the_uri_attribute(self):
        root = parse_xml(_datastore_xml(_GUID_A, "urn:foo"))
        assert root.schemaRefs.schemaRef_lst[0].uri == "urn:foo"

    def it_can_change_the_uri_attribute(self):
        root = parse_xml(_datastore_xml(_GUID_A, "urn:foo"))
        ref = root.schemaRefs.schemaRef_lst[0]
        ref.uri = "urn:replaced"
        assert root.schema_ref_uris == ("urn:replaced",)


class DescribeCT_DatastoreSchemaRefs:
    def it_finds_a_ref_by_uri(self):
        root = parse_xml(_datastore_xml(_GUID_A, "urn:a", "urn:b"))
        found = root.schemaRefs.find_by_uri("urn:b")
        assert isinstance(found, CT_DatastoreSchemaRef)
        assert found.uri == "urn:b"

    def it_returns_None_for_unknown_uri(self):
        root = parse_xml(_datastore_xml(_GUID_A, "urn:a"))
        assert root.schemaRefs.find_by_uri("urn:missing") is None
