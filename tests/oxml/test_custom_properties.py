# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.custom_properties`."""

from __future__ import annotations

import datetime as dt

import pytest
from lxml import etree

from pptx.oxml import parse_xml
from pptx.oxml.custom_properties import (
    DEFAULT_FMTID,
    CT_Properties,
    CT_Property,
    CT_VtBool,
    CT_VtFiletime,
    CT_VtI4,
    CT_VtLpwstr,
    CT_VtR8,
)
from pptx.oxml.ns import nsdecls


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


class DescribeCT_Properties:
    def it_parses_to_the_registered_class(self):
        root = parse_xml(_props_xml())
        assert isinstance(root, CT_Properties)

    def it_can_create_a_fresh_root_with_both_namespaces_declared(self):
        root = CT_Properties.new_properties()
        xml = etree.tostring(root, encoding="unicode")
        assert "xmlns:op=" in xml
        assert "xmlns:vt=" in xml

    def it_returns_the_property_lst_in_document_order(self):
        root = parse_xml(
            _props_xml(
                _property_xml("alpha", 2, "<vt:lpwstr>a</vt:lpwstr>"),
                _property_xml("beta", 3, "<vt:i4>1</vt:i4>"),
            )
        )
        assert root.property_names == ("alpha", "beta")

    def it_finds_a_property_by_name(self):
        root = parse_xml(
            _props_xml(
                _property_xml("Source", 2, "<vt:lpwstr>cli</vt:lpwstr>"),
                _property_xml("Build", 3, "<vt:i4>42</vt:i4>"),
            )
        )
        assert root.get_property("Build").pid == 3
        assert root.get_property("Missing") is None

    def it_removes_a_property_by_name(self):
        root = parse_xml(
            _props_xml(
                _property_xml("Source", 2, "<vt:lpwstr>cli</vt:lpwstr>"),
                _property_xml("Build", 3, "<vt:i4>42</vt:i4>"),
            )
        )
        assert root.remove_property("Source") is True
        assert root.property_names == ("Build",)
        assert root.remove_property("Source") is False  # idempotent

    @pytest.mark.parametrize(
        ("value", "expected_child_tag"),
        [
            (
                "hello",
                "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpwstr",
            ),
            (42, "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}i4"),
            (3.14, "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}r8"),
            (True, "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}bool"),
            (
                dt.datetime(2026, 5, 5, 14, 0, 0),
                "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}filetime",
            ),
        ],
    )
    def it_dispatches_add_property_by_python_type(self, value, expected_child_tag):
        root = CT_Properties.new_properties()
        prop = root.add_property("X", value)
        assert prop.fmtid == DEFAULT_FMTID
        assert prop.name == "X"
        # exactly one vt:* child, of the expected tag
        assert len(prop) == 1
        assert prop[0].tag == expected_child_tag

    def it_round_trips_value_for_each_vt_type(self):
        root = CT_Properties.new_properties()
        cases: list[tuple[str, object]] = [
            ("Source", "deck-builder-cli@1.4.2"),
            ("BuildNumber", 42),
            ("WeightedScore", 3.14159),
            ("IsDraft", True),
            ("IsFinal", False),
            ("GeneratedAt", dt.datetime(2026, 5, 5, 14, 0, 0)),
        ]
        for name, value in cases:
            root.add_property(name, value)

        serialized = etree.tostring(root)
        reparsed = parse_xml(serialized)
        for name, value in cases:
            prop = reparsed.get_property(name)
            assert prop is not None, name
            assert prop.value == value, name

    def it_auto_assigns_unique_pids_starting_at_2(self):
        root = CT_Properties.new_properties()
        a = root.add_property("a", "1")
        b = root.add_property("b", "2")
        c = root.add_property("c", "3")
        assert (a.pid, b.pid, c.pid) == (2, 3, 4)

    def it_skips_used_pids_when_assigning(self):
        # parse a doc where pid 2 is already used; the next add_property must use 3
        root = parse_xml(_props_xml(_property_xml("Existing", 2, "<vt:lpwstr>x</vt:lpwstr>")))
        new_prop = root.add_property("New", "y")
        assert new_prop.pid == 3

    def it_raises_TypeError_on_unsupported_value_type(self):
        root = CT_Properties.new_properties()
        with pytest.raises(TypeError):
            root.add_property("bad", object())

    def it_treats_bool_as_bool_not_int(self):
        # bool is-a int in Python; the dispatch must still produce vt:bool, not vt:i4
        root = CT_Properties.new_properties()
        prop_true = root.add_property("flag", True)
        assert isinstance(prop_true.bool_, CT_VtBool)
        assert prop_true.i4 is None


class DescribeCT_VtLpwstr:
    def it_round_trips_string_text(self):
        prop = parse_xml(
            _property_xml("X", 2, "<vt:lpwstr>hello world</vt:lpwstr>").encode()
            if False
            else (
                '<op:property %s fmtid="%s" pid="2" name="X">'
                "<vt:lpwstr>hello world</vt:lpwstr></op:property>"
                % (nsdecls("op", "vt"), DEFAULT_FMTID)
            ).encode()
        )
        assert isinstance(prop.lpwstr, CT_VtLpwstr)
        assert prop.value == "hello world"

    def it_rejects_non_string_assignment(self):
        prop = CT_Properties.new_properties().add_property("X", "seed")
        prop_lpwstr: CT_VtLpwstr = prop.lpwstr
        with pytest.raises(TypeError):
            prop_lpwstr.value_typed = 42  # type: ignore[assignment]

    def it_rejects_overlong_strings(self):
        prop = CT_Properties.new_properties().add_property("X", "seed")
        with pytest.raises(ValueError):
            prop.lpwstr.value_typed = "x" * 256


class DescribeCT_VtI4:
    @pytest.mark.parametrize("value", [-2147483648, -1, 0, 1, 42, 2147483647])
    def it_round_trips_int_text(self, value):
        prop = CT_Properties.new_properties().add_property("X", value)
        assert isinstance(prop.i4, CT_VtI4)
        assert prop.value == value

    def it_rejects_out_of_range_ints(self):
        prop = CT_Properties.new_properties().add_property("X", 0)
        with pytest.raises(ValueError):
            prop.i4.value_typed = 2147483648

    def it_rejects_bool_assignment_at_the_leaf(self):
        # the dispatch in CT_Property.value picks vt:bool for bool, but if a
        # caller reaches into the leaf they should still get the type guard
        prop = CT_Properties.new_properties().add_property("X", 0)
        with pytest.raises(TypeError):
            prop.i4.value_typed = True


class DescribeCT_VtR8:
    @pytest.mark.parametrize("value", [-1.0, 0.0, 0.5, 3.14159, 1e20, -1e-20])
    def it_round_trips_float_text(self, value):
        prop = CT_Properties.new_properties().add_property("X", value)
        assert isinstance(prop.r8, CT_VtR8)
        assert prop.value == pytest.approx(value)


class DescribeCT_VtBool:
    @pytest.mark.parametrize(
        ("xml_text", "expected"),
        [("true", True), ("false", False), ("1", True), ("0", False), (" TRUE ", True)],
    )
    def it_reads_office_and_xsd_boolean_lexical_forms(self, xml_text, expected):
        prop_xml = (
            '<op:property %s fmtid="%s" pid="2" name="X">'
            "<vt:bool>%s</vt:bool></op:property>" % (nsdecls("op", "vt"), DEFAULT_FMTID, xml_text)
        )
        prop = parse_xml(prop_xml.encode())
        assert prop.value is expected

    @pytest.mark.parametrize(("py_value", "expected_text"), [(True, "true"), (False, "false")])
    def it_writes_office_lexical_form(self, py_value, expected_text):
        prop = CT_Properties.new_properties().add_property("X", py_value)
        assert prop.bool_.text == expected_text

    def it_raises_on_invalid_text(self):
        prop_xml = (
            '<op:property %s fmtid="%s" pid="2" name="X">'
            "<vt:bool>maybe</vt:bool></op:property>" % (nsdecls("op", "vt"), DEFAULT_FMTID)
        )
        prop = parse_xml(prop_xml.encode())
        with pytest.raises(ValueError):
            _ = prop.value


class DescribeCT_VtFiletime:
    def it_round_trips_a_naive_utc_datetime(self):
        original = dt.datetime(2026, 5, 5, 14, 0, 0)
        prop = CT_Properties.new_properties().add_property("X", original)
        assert isinstance(prop.filetime, CT_VtFiletime)
        assert prop.filetime.text == "2026-05-05T14:00:00Z"
        assert prop.value == original

    def it_normalizes_a_tz_aware_datetime_to_utc(self):
        eastern = dt.timezone(dt.timedelta(hours=-5))
        aware = dt.datetime(2026, 5, 5, 9, 0, 0, tzinfo=eastern)  # 14:00 UTC
        prop = CT_Properties.new_properties().add_property("X", aware)
        assert prop.filetime.text == "2026-05-05T14:00:00Z"

    def it_parses_offset_form_too(self):
        prop_xml = (
            '<op:property %s fmtid="%s" pid="2" name="X">'
            "<vt:filetime>2026-05-05T09:00:00-05:00</vt:filetime></op:property>"
            % (nsdecls("op", "vt"), DEFAULT_FMTID)
        )
        prop = parse_xml(prop_xml.encode())
        assert prop.value == dt.datetime(2026, 5, 5, 14, 0, 0)


class DescribeCT_Property_value_setter:
    def it_replaces_an_existing_value_child(self):
        prop = CT_Properties.new_properties().add_property("X", "old")
        prop.value = 99
        assert prop.lpwstr is None
        assert prop.value == 99

    def it_returns_None_for_value_when_no_child_present(self):
        # build a stripped property element by parsing
        prop_xml = '<op:property %s fmtid="%s" pid="2" name="empty"/>' % (
            nsdecls("op", "vt"),
            DEFAULT_FMTID,
        )
        prop = parse_xml(prop_xml.encode())
        assert isinstance(prop, CT_Property)
        assert prop.value is None
