"""Unit-test suite for `pptx.oxml.slide` module."""

from __future__ import annotations

from typing import cast

import pytest

from pptx.oxml.slide import (
    CT_HandoutMaster,
    CT_HeaderFooter,
    CT_NotesMaster,
    CT_NotesSlide,
    CT_SlideLayout,
    CT_SlideMaster,
)

from ..unitutil.cxml import element
from ..unitutil.file import snippet_text


class DescribeCT_NotesMaster(object):
    """Unit-test suite for `pptx.oxml.slide.CT_NotesMaster` objects."""

    def it_can_create_a_default_notesMaster_element(self):
        notesMaster = CT_NotesMaster.new_default()
        assert notesMaster.xml == snippet_text("default-notesMaster")


class DescribeCT_NotesSlide(object):
    """Unit-test suite for `pptx.oxml.slide.CT_NotesSlide` objects."""

    def it_can_create_a_new_notes_element(self):
        notes = CT_NotesSlide.new()
        assert notes.xml == snippet_text("default-notes")


class DescribeCT_HeaderFooter(object):
    """Unit-test suite for `pptx.oxml.slide.CT_HeaderFooter` objects."""

    @pytest.mark.parametrize(
        ("hf_cxml", "expected"),
        [
            ("p:hf", (True, True, True, True)),
            ("p:hf{sldNum=0,hdr=0,ftr=0,dt=0}", (False, False, False, False)),
            ("p:hf{sldNum=1,hdr=1,ftr=1,dt=1}", (True, True, True, True)),
            ("p:hf{sldNum=0}", (False, True, True, True)),
            ("p:hf{hdr=0}", (True, False, True, True)),
            ("p:hf{ftr=0}", (True, True, False, True)),
            ("p:hf{dt=0}", (True, True, True, False)),
        ],
    )
    def it_provides_boolean_access_to_its_four_visibility_attrs(
        self, hf_cxml: str, expected: tuple[bool, bool, bool, bool]
    ):
        hf = cast(CT_HeaderFooter, element(hf_cxml))
        assert (hf.sldNum, hf.hdr, hf.ftr, hf.dt) == expected

    @pytest.mark.parametrize("attr_name", ["sldNum", "hdr", "ftr", "dt"])
    def it_can_toggle_each_attribute_via_setter(self, attr_name: str):
        hf = cast(CT_HeaderFooter, element("p:hf"))
        # ---default value is True when attr is absent---
        assert getattr(hf, attr_name) is True
        # ---set False, read back False---
        setattr(hf, attr_name, False)
        assert getattr(hf, attr_name) is False
        # ---set back to True---
        setattr(hf, attr_name, True)
        assert getattr(hf, attr_name) is True


class DescribeCT_HandoutMaster(object):
    """Unit-test suite for `pptx.oxml.slide.CT_HandoutMaster` objects."""

    def it_provides_access_to_its_hf_child(self):
        handoutMaster = cast(
            CT_HandoutMaster,
            element("p:handoutMaster/(p:cSld/p:spTree,p:clrMap,p:hf{ftr=0})"),
        )
        assert handoutMaster.hf is not None
        assert handoutMaster.hf.ftr is False

    def it_returns_None_for_hf_when_absent(self):
        handoutMaster = cast(
            CT_HandoutMaster, element("p:handoutMaster/(p:cSld/p:spTree,p:clrMap)")
        )
        assert handoutMaster.hf is None

    def it_can_add_an_hf_child_via_get_or_add(self):
        handoutMaster = cast(
            CT_HandoutMaster, element("p:handoutMaster/(p:cSld/p:spTree,p:clrMap)")
        )
        hf = handoutMaster.get_or_add_hf()
        assert hf is handoutMaster.hf
        # ---defaults all True on a freshly-added <p:hf/>---
        assert (hf.sldNum, hf.hdr, hf.ftr, hf.dt) == (True, True, True, True)


class DescribeHFAccessOnTemplates(object):
    """`hf` ZeroOrOne accessor on SlideMaster / SlideLayout / NotesMaster / HandoutMaster."""

    def it_reads_hf_on_a_sldMaster(self):
        sldMaster = cast(
            CT_SlideMaster,
            element("p:sldMaster/(p:cSld/p:spTree,p:hf{sldNum=0,ftr=0})"),
        )
        assert sldMaster.hf is not None
        assert sldMaster.hf.sldNum is False
        assert sldMaster.hf.ftr is False
        assert sldMaster.hf.dt is True

    def it_returns_None_for_hf_on_a_sldMaster_when_absent(self):
        sldMaster = cast(CT_SlideMaster, element("p:sldMaster/p:cSld/p:spTree"))
        assert sldMaster.hf is None

    def it_reads_hf_on_a_sldLayout(self):
        sldLayout = cast(CT_SlideLayout, element("p:sldLayout/(p:cSld/p:spTree,p:hf{dt=0})"))
        assert sldLayout.hf is not None
        assert sldLayout.hf.dt is False

    def it_returns_None_for_hf_on_a_sldLayout_when_absent(self):
        sldLayout = cast(CT_SlideLayout, element("p:sldLayout/p:cSld/p:spTree"))
        assert sldLayout.hf is None

    def it_reads_hf_on_a_notesMaster(self):
        notesMaster = cast(
            CT_NotesMaster,
            element("p:notesMaster/(p:cSld/p:spTree,p:hf{hdr=0,ftr=0})"),
        )
        assert notesMaster.hf is not None
        assert notesMaster.hf.hdr is False
        assert notesMaster.hf.ftr is False

    def it_returns_None_for_hf_on_a_notesMaster_when_absent(self):
        notesMaster = cast(CT_NotesMaster, element("p:notesMaster/p:cSld/p:spTree"))
        assert notesMaster.hf is None
