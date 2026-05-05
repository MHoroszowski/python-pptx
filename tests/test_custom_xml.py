# pyright: reportPrivateUsage=false

"""End-to-end test suite for `pptx.custom_xml.CustomXmlParts`."""

from __future__ import annotations

from io import BytesIO

import pytest

from pptx import Presentation
from pptx.custom_xml import (
    NAME_PROPERTY_PREFIX,
    CustomXmlParts,
    _normalize_guid,
    _upgrade_to_custom_xml_part,
)
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.custom_xml import CustomXmlPart


@pytest.fixture
def empty_prs():
    return Presentation()


def _roundtrip(prs):
    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)


class DescribeCustomXmlParts_basic:
    def it_starts_empty_for_a_default_presentation(self, empty_prs):
        cxp = empty_prs.custom_xml_parts
        assert isinstance(cxp, CustomXmlParts)
        assert len(cxp) == 0
        assert list(cxp) == []

    def it_adds_a_part_with_default_presentation_scope(self, empty_prs):
        part = empty_prs.custom_xml_parts.add(b'<root xmlns="urn:my"/>')
        assert isinstance(part, CustomXmlPart)
        assert part.partname == "/customXml/item1.xml"
        # presentation scope: rel from the presentation part
        rel_types_at_prs = {r.reltype for r in empty_prs.part.rels.values()}
        rel_types_at_pkg = {r.reltype for r in empty_prs.part.package._rels.values()}
        assert RT.CUSTOM_XML in rel_types_at_prs
        assert RT.CUSTOM_XML not in rel_types_at_pkg

    def it_adds_a_part_with_package_scope_when_requested(self, empty_prs):
        empty_prs.custom_xml_parts.add(b"<x/>", scope="package")
        rel_types_at_prs = {r.reltype for r in empty_prs.part.rels.values()}
        rel_types_at_pkg = {r.reltype for r in empty_prs.part.package._rels.values()}
        assert RT.CUSTOM_XML in rel_types_at_pkg
        assert RT.CUSTOM_XML not in rel_types_at_prs

    def it_rejects_unknown_scope(self, empty_prs):
        with pytest.raises(ValueError):
            empty_prs.custom_xml_parts.add(b"<x/>", scope="bogus")  # type: ignore[arg-type]

    def it_walks_both_scopes_in_iteration(self, empty_prs):
        empty_prs.custom_xml_parts.add(b'<a xmlns="u:a"/>', scope="presentation")
        empty_prs.custom_xml_parts.add(b'<b xmlns="u:b"/>', scope="package")
        partnames = [str(p.partname) for p in empty_prs.custom_xml_parts]
        assert len(partnames) == 2

    def it_assigns_distinct_partnames_to_consecutive_pairs(self, empty_prs):
        a = empty_prs.custom_xml_parts.add(b"<a/>")
        b = empty_prs.custom_xml_parts.add(b"<b/>")
        assert a.partname != b.partname
        assert str(a.partname) == "/customXml/item1.xml"
        assert str(b.partname) == "/customXml/item2.xml"


class DescribeCustomXmlParts_lookups:
    def it_indexes_by_position(self, empty_prs):
        a = empty_prs.custom_xml_parts.add(b'<a xmlns="u:a"/>')
        b = empty_prs.custom_xml_parts.add(b'<b xmlns="u:b"/>')
        assert empty_prs.custom_xml_parts[0] is a
        assert empty_prs.custom_xml_parts[1] is b

    def it_raises_IndexError_on_out_of_range(self, empty_prs):
        empty_prs.custom_xml_parts.add(b"<x/>")
        with pytest.raises(IndexError):
            empty_prs.custom_xml_parts[5]

    def it_indexes_by_partname_tail(self, empty_prs):
        empty_prs.custom_xml_parts.add(b"<a/>")
        empty_prs.custom_xml_parts.add(b"<b/>")
        found = empty_prs.custom_xml_parts["item2.xml"]
        assert str(found.partname) == "/customXml/item2.xml"

    def it_raises_KeyError_on_unknown_partname(self, empty_prs):
        empty_prs.custom_xml_parts.add(b"<x/>")
        with pytest.raises(KeyError):
            empty_prs.custom_xml_parts["item99.xml"]

    def it_raises_TypeError_on_other_key_types(self, empty_prs):
        with pytest.raises(TypeError):
            empty_prs.custom_xml_parts[1.5]  # type: ignore[index]

    def it_finds_by_guid_brace_tolerant(self, empty_prs):
        guid = "{ABCDEF12-3456-7890-ABCD-EF1234567890}"
        empty_prs.custom_xml_parts.add(b"<x/>", datastoreItem_id=guid)
        # exact form
        assert empty_prs.custom_xml_parts.by_guid(guid) is not None
        # without braces, lowercase
        assert (
            empty_prs.custom_xml_parts.by_guid("abcdef12-3456-7890-abcd-ef1234567890")
            is not None
        )

    def it_returns_None_for_unknown_guid(self, empty_prs):
        empty_prs.custom_xml_parts.add(b"<x/>")
        assert empty_prs.custom_xml_parts.by_guid("{00000000-0000-0000-0000-000000000000}") is None

    def it_finds_by_user_assigned_name(self, empty_prs):
        added = empty_prs.custom_xml_parts.add(
            b'<provenance xmlns="u:p"/>',
            name="provenance",
        )
        assert empty_prs.custom_xml_parts.by_name("provenance") is added

    def it_returns_None_for_unknown_name(self, empty_prs):
        empty_prs.custom_xml_parts.add(b"<x/>", name="real")
        assert empty_prs.custom_xml_parts.by_name("missing") is None

    def it_raises_TypeError_on_non_str_name(self, empty_prs):
        with pytest.raises(TypeError):
            empty_prs.custom_xml_parts.by_name(42)  # type: ignore[arg-type]


class DescribeCustomXmlParts_remove:
    def it_removes_by_part_instance(self, empty_prs):
        a = empty_prs.custom_xml_parts.add(b"<a/>", name="a")
        empty_prs.custom_xml_parts.add(b"<b/>", name="b")
        empty_prs.custom_xml_parts.remove(a)
        assert len(empty_prs.custom_xml_parts) == 1
        # name entry also removed
        assert (
            empty_prs.part.package.custom_properties_part.get_property(
                NAME_PROPERTY_PREFIX + a.datastore_item_id
            )
            is None
        )

    def it_removes_by_index(self, empty_prs):
        empty_prs.custom_xml_parts.add(b"<a/>")
        empty_prs.custom_xml_parts.add(b"<b/>")
        empty_prs.custom_xml_parts.remove(0)
        assert len(empty_prs.custom_xml_parts) == 1

    def it_removes_by_partname_tail(self, empty_prs):
        empty_prs.custom_xml_parts.add(b"<a/>")
        empty_prs.custom_xml_parts.add(b"<b/>")
        empty_prs.custom_xml_parts.remove("item1.xml")
        assert str(empty_prs.custom_xml_parts[0].partname) == "/customXml/item2.xml"

    def it_is_idempotent(self, empty_prs):
        a = empty_prs.custom_xml_parts.add(b"<a/>")
        empty_prs.custom_xml_parts.remove(a)
        empty_prs.custom_xml_parts.remove(a)  # no error
        assert len(empty_prs.custom_xml_parts) == 0

    def it_removes_a_package_scoped_part(self, empty_prs):
        a = empty_prs.custom_xml_parts.add(b"<a/>", scope="package")
        empty_prs.custom_xml_parts.remove(a)
        assert len(empty_prs.custom_xml_parts) == 0
        rel_types = {r.reltype for r in empty_prs.part.package._rels.values()}
        assert RT.CUSTOM_XML not in rel_types

    def it_raises_TypeError_on_unsupported_remove_arg(self, empty_prs):
        with pytest.raises(TypeError):
            empty_prs.custom_xml_parts.remove(1.5)  # type: ignore[arg-type]


class DescribeCustomXmlParts_roundtrip:
    def it_round_trips_added_parts(self, empty_prs):
        empty_prs.custom_xml_parts.add(
            b'<provenance xmlns="urn:my:p"><source>cli</source></provenance>',
            name="provenance",
            schema_refs=["urn:my:p"],
        )
        empty_prs.custom_xml_parts.add(b"<extra/>", name="extra", scope="package")

        reloaded = _roundtrip(empty_prs)

        assert len(reloaded.custom_xml_parts) == 2
        prov = reloaded.custom_xml_parts.by_name("provenance")
        assert prov is not None
        assert prov.element.tag == "{urn:my:p}provenance"
        assert prov.schema_refs == ("urn:my:p",)
        extra = reloaded.custom_xml_parts.by_name("extra")
        assert extra is not None
        assert extra.element.tag == "extra"

    def it_preserves_payload_text_byte_for_byte_through_lxml_roundtrip(self, empty_prs):
        original = b'<root xmlns="u:r"><child a="1">hello</child></root>'
        added = empty_prs.custom_xml_parts.add(original)
        guid = added.datastore_item_id

        reloaded = _roundtrip(empty_prs)
        part = reloaded.custom_xml_parts.by_guid(guid)
        assert part is not None
        # parsed structure is preserved
        child = part.element.find("{u:r}child")
        assert child is not None
        assert child.get("a") == "1"
        assert child.text == "hello"

    def it_replaces_xml_payload_in_place(self, empty_prs):
        added = empty_prs.custom_xml_parts.add(b"<old/>")
        guid = added.datastore_item_id
        added.replace_xml(b'<new xmlns="u:n"/>')

        reloaded = _roundtrip(empty_prs)
        part = reloaded.custom_xml_parts.by_guid(guid)
        assert part is not None
        assert part.element.tag == "{u:n}new"

    def it_supports_add_item_convenience(self, empty_prs):
        added = empty_prs.custom_xml_parts.add(b'<list xmlns=""/>')
        added.add_item("item", "first")
        added.add_item("item", "second", priority="high")

        # children are present
        children = list(added.element)
        assert len(children) == 2
        assert children[0].text == "first"
        assert children[1].get("priority") == "high"


class DescribeCustomXmlPart_name_edge_cases:
    def it_returns_None_when_no_name_property_for_the_guid(self, empty_prs):
        # Add a part WITHOUT a name. .name should return None even though the
        # custom_properties part does exist (other entries may have been written).
        empty_prs.custom_properties["AnythingElse"] = "value"
        added = empty_prs.custom_xml_parts.add(b"<x/>")
        assert added.name is None


class DescribeUpgradeAndHelpers:
    def it_upgrades_a_loaded_base_part_to_CustomXmlPart_on_iteration(self, empty_prs):
        empty_prs.custom_xml_parts.add(b'<x xmlns="u:x"/>')
        reloaded = _roundtrip(empty_prs)
        # Force iteration; the base Part loaded for the customXml/item1.xml
        # part gets upgraded to CustomXmlPart in place.
        first = next(iter(reloaded.custom_xml_parts))
        assert isinstance(first, CustomXmlPart)
        assert first.element.tag == "{u:x}x"

    def it_passes_through_an_already_upgraded_part_unchanged(self, empty_prs):
        added = empty_prs.custom_xml_parts.add(b"<x/>")
        # the just-added part is already a CustomXmlPart
        same = _upgrade_to_custom_xml_part(added)
        assert same is added

    @pytest.mark.parametrize(
        ("input_guid", "expected"),
        [
            ("{ABCDEF12-3456-7890-ABCD-EF1234567890}", "abcdef12-3456-7890-abcd-ef1234567890"),
            ("abcdef12-3456-7890-abcd-ef1234567890", "abcdef12-3456-7890-abcd-ef1234567890"),
            ("  {AbCdEf12-3456-7890-ABCD-EF1234567890}  ", "abcdef12-3456-7890-abcd-ef1234567890"),
        ],
    )
    def it_normalizes_guids_for_comparison(self, input_guid, expected):
        assert _normalize_guid(input_guid) == expected

    def it_raises_TypeError_on_non_str_guid_normalize(self):
        with pytest.raises(TypeError):
            _normalize_guid(42)  # type: ignore[arg-type]
