"""Unit-test suite for `pptx.oxml.text` module."""

from __future__ import annotations

from typing import cast

import pytest

from pptx.exc import InvalidXmlError
from pptx.oxml.text import CT_TextField

from ..unitutil.cxml import element


class DescribeCT_TextField(object):
    """Unit-test suite for `pptx.oxml.text.CT_TextField` (the `a:fld` element)."""

    def it_provides_read_access_to_its_id_attribute(self):
        # ---cxml's grammar reserves `{` and `}` as attribute delimiters, so the
        # ---literal `{GUID}` form cannot appear inline. The id attribute on
        # ---<a:fld> is XsdString-typed so any token works for round-trip tests;
        # ---real fields generate {uuid4()} values at author time (Phase 3).
        fld = cast(CT_TextField, element("a:fld{id=fld-1,type=slidenum}"))
        assert fld.id == "fld-1"

    def it_raises_InvalidXmlError_when_id_is_missing(self):
        fld = cast(CT_TextField, element("a:fld"))
        with pytest.raises(InvalidXmlError):
            _ = fld.id

    @pytest.mark.parametrize(
        ("fld_cxml", "expected_type"),
        [
            ("a:fld{id=foo,type=slidenum}", "slidenum"),
            ("a:fld{id=foo,type=datetime1}", "datetime1"),
            ("a:fld{id=foo,type=datetime13}", "datetime13"),
            ("a:fld{id=foo,type=title}", "title"),
            ("a:fld{id=foo}", None),
        ],
    )
    def it_provides_read_access_to_its_type_attribute(
        self, fld_cxml: str, expected_type: str | None
    ):
        fld = cast(CT_TextField, element(fld_cxml))
        assert fld.type == expected_type

    def it_returns_empty_string_for_text_when_a_t_is_absent(self):
        fld = cast(CT_TextField, element("a:fld{id=foo}"))
        assert fld.text == ""

    def it_reads_the_text_of_its_a_t_child(self):
        fld = cast(CT_TextField, element('a:fld{id=foo,type=slidenum}/a:t"42"'))
        assert fld.text == "42"
