"""Unit-test suite for `pptx.oxml.text` module."""

from __future__ import annotations

from typing import cast

import pytest

from pptx.exc import InvalidXmlError
from pptx.oxml.text import CT_TextField, CT_TextParagraph

from ..unitutil.cxml import element, xml


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

    def it_adds_an_a_t_child_on_text_assignment_when_absent(self):
        fld = cast(CT_TextField, element("a:fld{id=foo,type=slidenum}"))
        assert fld.t is None

        fld.text = "‹#›"

        assert fld.t is not None
        assert fld.text == "‹#›"
        assert fld.xml == xml('a:fld{id=foo,type=slidenum}/a:t"‹#›"')

    def it_replaces_existing_a_t_content_on_text_assignment(self):
        fld = cast(CT_TextField, element('a:fld{id=foo,type=slidenum}/a:t"old"'))

        fld.text = "new"

        assert fld.text == "new"
        # ---only one a:t child is present; assignment replaces, not appends---
        assert len(fld.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}t")) == 1

    @pytest.mark.parametrize(
        ("input_value", "expected_a_t_text"),
        [
            ("hello", "hello"),
            ("a\x07b", "a_x0007_b"),  # BEL escapes
            ("tab\there", "tab\there"),  # tab pass-through
            ("line1\nline2", "line1\nline2"),  # newline pass-through
            ("esc\x1bhere", "esc_x001B_here"),  # ESC escapes, uppercase hex
            ("", ""),
        ],
    )
    def it_escapes_control_chars_when_assigning_text(
        self, input_value: str, expected_a_t_text: str
    ):
        fld = cast(CT_TextField, element("a:fld{id=foo,type=slidenum}"))

        fld.text = input_value

        # ---round-trip: reading back returns the escaped form because the
        # ---escape is permanent storage form, not a presentation layer---
        assert fld.text == expected_a_t_text

    def it_escapes_BEL_to_uppercase_hex_via__escape_ctrl_chars(self):
        # ---BEL is x07; expected escape is "_x0007_" with uppercase hex---
        assert CT_TextField._escape_ctrl_chars("ring\x07bell") == "ring_x0007_bell"

    def it_passes_tab_and_newline_through__escape_ctrl_chars(self):
        # ---x09 (HT) and x0A (LF) are explicitly excluded from the escape
        # ---range per OOXML §22.9.2.19; all other x00..x1F characters escape.
        assert CT_TextField._escape_ctrl_chars("a\tb\nc") == "a\tb\nc"

        # ---verify x0B (VT) and x1F (US, the highest in-range value) DO escape
        assert CT_TextField._escape_ctrl_chars("\x0b") == "_x000B_"
        assert CT_TextField._escape_ctrl_chars("\x1f") == "_x001F_"


class DescribeCT_TextParagraph(object):
    """Unit-test suite for `pptx.oxml.text.CT_TextParagraph` field accessor."""

    def it_can_add_an_a_fld_via__add_fld(self):
        p = cast(CT_TextParagraph, element("a:p"))

        fld = p._add_fld()

        assert isinstance(fld, CT_TextField)
        assert len(p.fld_lst) == 1
        assert p.fld_lst[0] is fld

    def it_appends_a_fld_after_existing_runs_in_document_order(self):
        # ---fld successors=("a:endParaRPr",) places it after a:r and a:br,
        # ---and before a:endParaRPr. Verifying the post-r position confirms
        # ---xmlchemy honored the successors tuple correctly.
        p = cast(CT_TextParagraph, element('a:p/(a:r/a:t"x",a:endParaRPr)'))

        fld = p._add_fld()
        fld.id = "fld-1"

        # ---walk children of <a:p>; ignoring pPr (none here), expect:
        # ---a:r, a:fld, a:endParaRPr in that order
        tags = [child.tag.split("}")[-1] for child in p]
        assert tags == ["r", "fld", "endParaRPr"]

    def it_returns_all_fld_children_via_fld_lst(self):
        p = cast(CT_TextParagraph, element("a:p"))

        fld_a = p._add_fld()
        fld_b = p._add_fld()
        fld_c = p._add_fld()

        assert p.fld_lst == [fld_a, fld_b, fld_c]

    def it_includes_a_fld_in_content_children(self):
        # ---content_children must surface a:r, a:br, and a:fld in document order
        # ---so _Paragraph.text concatenates field text alongside run text.
        p = cast(CT_TextParagraph, element("a:p"))
        p.add_r("before")
        fld = p._add_fld()
        fld.id = "fld-1"
        fld.text = "[N]"
        p.add_r("after")

        texts = [child.text for child in p.content_children]
        assert texts == ["before", "[N]", "after"]
