"""Unit-test suite for `pptx.dml.effect` module."""

from __future__ import annotations

import pytest

from pptx.dml.color import ColorFormat
from pptx.dml.effect import ShadowFormat
from pptx.util import Pt

from ..unitutil.cxml import element, xml


class DescribeShadowFormat(object):
    @pytest.mark.parametrize(
        ("spPr_cxml", "expected_value"),
        [
            ("p:spPr", False),
            ("p:spPr/a:effectLst", False),
            ("p:spPr/a:effectLst/a:outerShdw", True),
        ],
    )
    def it_knows_whether_a_shadow_is_visible(self, spPr_cxml: str, expected_value: bool):
        shadow = ShadowFormat(element(spPr_cxml))
        assert shadow.visible is expected_value

    def it_can_make_a_shadow_visible(self):
        spPr = element("p:spPr")
        shadow = ShadowFormat(spPr)
        shadow.visible = True
        assert shadow.visible is True
        assert spPr.effectLst is not None
        assert spPr.effectLst.outerShdw is not None

    def it_can_make_a_shadow_invisible(self):
        spPr = element("p:spPr/a:effectLst/a:outerShdw")
        shadow = ShadowFormat(spPr)
        shadow.visible = False
        assert shadow.visible is False

    @pytest.mark.parametrize(
        ("spPr_cxml", "expected_value"),
        [
            ("p:spPr", None),
            ("p:spPr/a:effectLst", None),
            ("p:spPr/a:effectLst/a:outerShdw", 0.0),
            ("p:spPr/a:effectLst/a:outerShdw{dir=2700000}", 45.0),
            ("p:spPr/a:effectLst/a:outerShdw{dir=5400000}", 90.0),
        ],
    )
    def it_knows_the_shadow_angle(self, spPr_cxml: str, expected_value: float | None):
        shadow = ShadowFormat(element(spPr_cxml))
        assert shadow.angle == expected_value

    def it_can_set_the_shadow_angle(self):
        spPr = element("p:spPr")
        shadow = ShadowFormat(spPr)
        shadow.angle = 45.0
        assert shadow.angle == 45.0

    @pytest.mark.parametrize(
        ("spPr_cxml", "expected_value"),
        [
            ("p:spPr", None),
            ("p:spPr/a:effectLst/a:outerShdw", 0),
            ("p:spPr/a:effectLst/a:outerShdw{blurRad=50800}", 50800),
        ],
    )
    def it_knows_the_blur_radius(self, spPr_cxml: str, expected_value: int | None):
        shadow = ShadowFormat(element(spPr_cxml))
        assert shadow.blur_radius == expected_value

    def it_can_set_the_blur_radius(self):
        spPr = element("p:spPr")
        shadow = ShadowFormat(spPr)
        shadow.blur_radius = Pt(4)
        assert shadow.blur_radius == Pt(4)

    @pytest.mark.parametrize(
        ("spPr_cxml", "expected_value"),
        [
            ("p:spPr", None),
            ("p:spPr/a:effectLst/a:outerShdw", 0),
            ("p:spPr/a:effectLst/a:outerShdw{dist=38100}", 38100),
        ],
    )
    def it_knows_the_shadow_distance(self, spPr_cxml: str, expected_value: int | None):
        shadow = ShadowFormat(element(spPr_cxml))
        assert shadow.distance == expected_value

    def it_can_set_the_shadow_distance(self):
        spPr = element("p:spPr")
        shadow = ShadowFormat(spPr)
        shadow.distance = Pt(3)
        assert shadow.distance == Pt(3)

    def it_provides_access_to_the_shadow_color(self):
        spPr = element("p:spPr")
        shadow = ShadowFormat(spPr)
        assert isinstance(shadow.color, ColorFormat)

    @pytest.mark.parametrize(
        ("spPr_cxml", "expected_value"),
        [
            ("p:spPr", None),
            ("p:spPr/a:effectLst/a:outerShdw", True),
            ("p:spPr/a:effectLst/a:outerShdw{rotWithShape=0}", False),
            ("p:spPr/a:effectLst/a:outerShdw{rotWithShape=1}", True),
        ],
    )
    def it_knows_rotate_with_shape(self, spPr_cxml: str, expected_value: bool | None):
        shadow = ShadowFormat(element(spPr_cxml))
        assert shadow.rotate_with_shape == expected_value

    def it_knows_whether_it_inherits(self, inherit_get_fixture):
        shadow, expected_value = inherit_get_fixture
        inherit = shadow.inherit
        assert inherit is expected_value

    def it_can_change_whether_it_inherits(self, inherit_set_fixture):
        shadow, value, expected_xml = inherit_set_fixture
        shadow.inherit = value
        assert shadow._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("p:spPr", True),
            ("p:spPr/a:effectLst", False),
            ("p:grpSpPr", True),
            ("p:grpSpPr/a:effectLst", False),
        ]
    )
    def inherit_get_fixture(self, request):
        cxml, expected_value = request.param
        shadow = ShadowFormat(element(cxml))
        return shadow, expected_value

    @pytest.fixture(
        params=[
            ("p:spPr{a:b=c}", False, "p:spPr{a:b=c}/a:effectLst"),
            ("p:grpSpPr{a:b=c}", False, "p:grpSpPr{a:b=c}/a:effectLst"),
            ("p:spPr{a:b=c}/a:effectLst", True, "p:spPr{a:b=c}"),
            ("p:grpSpPr{a:b=c}/a:effectLst", True, "p:grpSpPr{a:b=c}"),
            ("p:spPr", True, "p:spPr"),
            ("p:grpSpPr", True, "p:grpSpPr"),
            ("p:spPr/a:effectLst", False, "p:spPr/a:effectLst"),
            ("p:grpSpPr/a:effectLst", False, "p:grpSpPr/a:effectLst"),
        ]
    )
    def inherit_set_fixture(self, request):
        cxml, value, expected_cxml = request.param
        shadow = ShadowFormat(element(cxml))
        expected_value = xml(expected_cxml)
        return shadow, value, expected_value
