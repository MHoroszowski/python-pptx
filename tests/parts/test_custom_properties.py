# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.custom_properties`."""

from __future__ import annotations

import datetime as dt

import pytest

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.oxml.custom_properties import DEFAULT_FMTID, CT_Properties
from pptx.oxml.ns import nsdecls
from pptx.parts.custom_properties import CustomPropertiesPart


def _props_xml(*property_xml_chunks: str) -> bytes:
    body = "".join(property_xml_chunks)
    return ("<op:Properties %s>%s</op:Properties>" % (nsdecls("op", "vt"), body)).encode()


def _property_xml(name: str, pid: int, vt_inner_xml: str) -> str:
    return '<op:property fmtid="%s" pid="%d" name="%s">%s</op:property>' % (
        DEFAULT_FMTID,
        pid,
        name,
        vt_inner_xml,
    )


class DescribeCustomPropertiesPart:
    def it_can_construct_a_default_part(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
        assert isinstance(part, CustomPropertiesPart)
        assert part.content_type == CT.OFC_CUSTOM_PROPERTIES
        assert part.partname == "/docProps/custom.xml"
        assert isinstance(part._element, CT_Properties)
        assert part.property_names == ()

    def it_loads_an_existing_part_from_blob(self):
        xml = _props_xml(
            _property_xml("Source", 2, "<vt:lpwstr>cli</vt:lpwstr>"),
            _property_xml("Build", 3, "<vt:i4>42</vt:i4>"),
        )
        part = CustomPropertiesPart.load(
            "/docProps/custom.xml",
            CT.OFC_CUSTOM_PROPERTIES,
            None,
            xml,  # type: ignore[arg-type]
        )
        assert isinstance(part._element, CT_Properties)
        assert part.property_names == ("Source", "Build")

    def it_adds_a_property_via_delegation(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
        prop = part.add_property("Source", "cli@1.4")
        assert prop.name == "Source"
        assert prop.value == "cli@1.4"
        assert part.property_names == ("Source",)

    def it_dispatches_value_types_through_to_the_element(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
        part.add_property("Build", 42)
        part.add_property("Score", 3.14)
        part.add_property("IsDraft", True)
        part.add_property("At", dt.datetime(2026, 5, 5, 14, 0, 0))
        assert part.get_property("Build").value == 42
        assert part.get_property("Score").value == pytest.approx(3.14)
        assert part.get_property("IsDraft").value is True
        assert part.get_property("At").value == dt.datetime(2026, 5, 5, 14, 0, 0)

    def it_returns_None_when_property_missing(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
        assert part.get_property("Missing") is None

    def it_removes_a_property_idempotently(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
        part.add_property("X", "a")
        assert part.remove_property("X") is True
        assert part.property_names == ()
        assert part.remove_property("X") is False

    def it_supports_in_iter_and_len(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
        part.add_property("a", "1")
        part.add_property("b", "2")
        part.add_property("c", "3")
        assert len(part) == 3
        assert list(part) == ["a", "b", "c"]
        assert "b" in part
        assert "z" not in part
        # __contains__ on non-string is False
        assert (42 in part) is False  # type: ignore[operator]

    def it_round_trips_blob_through_add_and_reparse(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
        part.add_property("Source", "cli")
        part.add_property("Build", 99)
        blob = part.blob
        # blob is XML that re-parses to an equivalent CustomPropertiesPart
        reloaded = CustomPropertiesPart.load(
            "/docProps/custom.xml",
            CT.OFC_CUSTOM_PROPERTIES,
            None,
            blob,  # type: ignore[arg-type]
        )
        assert reloaded.property_names == ("Source", "Build")
        assert reloaded.get_property("Source").value == "cli"
        assert reloaded.get_property("Build").value == 99

    def it_assigns_unique_pids_across_adds(self):
        part = CustomPropertiesPart.default(None)  # type: ignore[arg-type]
        a = part.add_property("a", "1")
        b = part.add_property("b", "2")
        c = part.add_property("c", "3")
        assert (a.pid, b.pid, c.pid) == (2, 3, 4)
