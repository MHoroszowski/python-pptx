"""Unit-test suite for transparency functionality in `pptx.dml` module."""

from __future__ import annotations

import pytest

from pptx.dml.color import ColorFormat
from pptx.dml.fill import FillFormat, _NoFill, _SolidFill
from pptx.enum.dml import MSO_FILL

from ..oxml.unitdata.dml import a_alpha, a_solidFill, an_srgbClr


class DescribeColorFormatTransparency(object):
    """Unit-test suite for ColorFormat transparency property."""

    def it_knows_its_transparency_value(self, transparency_get_fixture):
        color_format, expected_transparency = transparency_get_fixture
        assert color_format.transparency == expected_transparency

    def it_can_set_its_transparency_value(self, transparency_set_fixture):
        color_format, transparency, expected_xml = transparency_set_fixture
        color_format.transparency = transparency
        assert color_format._xFill.xml == expected_xml

    def it_returns_zero_transparency_for_NoneColor(self, _NoneColor_color_format):
        assert _NoneColor_color_format.transparency == 0.0

    def it_raises_on_transparency_set_for_NoneColor(self, _NoneColor_color_format):
        with pytest.raises(ValueError):
            _NoneColor_color_format.transparency = 0.5

    def it_raises_on_assign_invalid_transparency_value(self, rgb_color_format):
        with pytest.raises(ValueError):
            rgb_color_format.transparency = 1.1
        with pytest.raises(ValueError):
            rgb_color_format.transparency = -0.1

    def it_can_set_transparency_to_zero_removes_alpha_element(self, rgb_color_format):
        rgb_color_format.transparency = 0.5
        assert rgb_color_format._color._xClr.alpha is not None

        rgb_color_format.transparency = 0.0
        assert rgb_color_format._color._xClr.alpha is None
        assert rgb_color_format.transparency == 0.0

    def it_can_set_transparency_to_one_creates_zero_alpha(self, rgb_color_format):
        rgb_color_format.transparency = 1.0

        alpha_element = rgb_color_format._color._xClr.alpha
        assert alpha_element is not None
        assert alpha_element.val == 0.0
        assert rgb_color_format.transparency == 1.0

    def it_validates_transparency_value_range(self, rgb_color_format):
        rgb_color_format._validate_transparency_value(0.0)
        rgb_color_format._validate_transparency_value(0.5)
        rgb_color_format._validate_transparency_value(1.0)

        with pytest.raises(ValueError, match="transparency must be number in range 0.0 to 1.0"):
            rgb_color_format._validate_transparency_value(-0.1)
        with pytest.raises(ValueError, match="transparency must be number in range 0.0 to 1.0"):
            rgb_color_format._validate_transparency_value(1.1)

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            (lambda: an_srgbClr().with_val("FF0000"), 0.0),
            (
                lambda: an_srgbClr().with_val("FF0000").with_child(a_alpha().with_val(100000)),
                0.0,
            ),
            (
                lambda: an_srgbClr().with_val("FF0000").with_child(a_alpha().with_val(50000)),
                0.5,
            ),
            (
                lambda: an_srgbClr().with_val("FF0000").with_child(a_alpha().with_val(0)),
                1.0,
            ),
        ]
    )
    def transparency_get_fixture(self, request):
        xClr_bldr_fn, expected_transparency = request.param
        xClr_bldr = xClr_bldr_fn()
        solidFill = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format, expected_transparency

    @pytest.fixture(
        params=[
            (
                lambda: an_srgbClr().with_val("FF0000"),
                0.0,
                lambda: an_srgbClr().with_val("FF0000"),
            ),
            (
                lambda: an_srgbClr().with_val("FF0000"),
                0.5,
                lambda: an_srgbClr().with_val("FF0000").with_child(a_alpha().with_val(50000)),
            ),
            (
                lambda: an_srgbClr().with_val("FF0000"),
                1.0,
                lambda: an_srgbClr().with_val("FF0000").with_child(a_alpha().with_val(0)),
            ),
            (
                lambda: an_srgbClr().with_val("FF0000").with_child(a_alpha().with_val(75000)),
                0.3,
                lambda: an_srgbClr().with_val("FF0000").with_child(a_alpha().with_val(70000)),
            ),
        ]
    )
    def transparency_set_fixture(self, request):
        xClr_bldr_fn, transparency, expected_xClr_bldr_fn = request.param

        xClr_bldr = xClr_bldr_fn()
        solidFill = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)

        expected_xClr_bldr = expected_xClr_bldr_fn()
        expected_xml = a_solidFill().with_nsdecls().with_child(expected_xClr_bldr).xml()

        return color_format, transparency, expected_xml

    @pytest.fixture
    def rgb_color_format(self):
        solidFill = a_solidFill().with_nsdecls().with_child(an_srgbClr().with_val("FF0000")).element
        return ColorFormat.from_colorchoice_parent(solidFill)

    @pytest.fixture
    def _NoneColor_color_format(self):
        solidFill = a_solidFill().with_nsdecls().element
        return ColorFormat.from_colorchoice_parent(solidFill)


class DescribeFillFormatTransparency(object):
    """Unit-test suite for FillFormat transparency property."""

    def it_delegates_transparency_to_solid_fill_object(self):
        xClr_bldr = an_srgbClr().with_val("FF0000")
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        solid_fill = _SolidFill(solidFill_elm)

        class _MockParent:
            def __init__(self, fill_elm):
                self.eg_fillProperties = fill_elm

        fill_format = FillFormat(_MockParent(solidFill_elm), solid_fill)

        fill_format.transparency = 0.3

        assert abs(fill_format._fill.transparency - 0.3) < 0.001
        assert abs(fill_format.transparency - 0.3) < 0.001

    def it_raises_on_transparency_access_for_non_solid_fills(self):
        class _MockParent:
            eg_fillProperties = None

        fill_format = FillFormat(_MockParent(), _NoFill(None))

        with pytest.raises(TypeError, match="fill type .* has no transparency"):
            fill_format.transparency

    def it_raises_on_transparency_set_for_non_solid_fills(self):
        class _MockParent:
            eg_fillProperties = None

        fill_format = FillFormat(_MockParent(), _NoFill(None))

        with pytest.raises(TypeError, match="fill type .* has no transparency"):
            fill_format.transparency = 0.5


class DescribeSolidFillTransparency(object):
    """Unit-test suite for _SolidFill transparency property."""

    def it_provides_access_to_transparency(self, solid_fill_transparency_fixture):
        solid_fill, expected_transparency = solid_fill_transparency_fixture
        assert abs(solid_fill.transparency - expected_transparency) < 0.001

    def it_can_set_transparency(self, solid_fill_transparency_set_fixture):
        solid_fill, transparency = solid_fill_transparency_set_fixture
        solid_fill.transparency = transparency
        assert abs(solid_fill.transparency - transparency) < 0.001

    def it_delegates_transparency_to_fore_color(self, solid_fill_obj):
        solid_fill_obj.transparency = 0.6

        assert solid_fill_obj.fore_color.transparency == 0.6
        assert solid_fill_obj.transparency == 0.6

    def it_has_correct_fill_type(self, solid_fill_obj):
        assert solid_fill_obj.type == MSO_FILL.SOLID

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            (lambda: an_srgbClr().with_val("00FF00"), 0.0),
            (
                lambda: an_srgbClr().with_val("00FF00").with_child(a_alpha().with_val(80000)),
                0.2,
            ),
            (
                lambda: an_srgbClr().with_val("00FF00").with_child(a_alpha().with_val(30000)),
                0.7,
            ),
        ]
    )
    def solid_fill_transparency_fixture(self, request):
        xClr_bldr_fn, expected_transparency = request.param
        xClr_bldr = xClr_bldr_fn()
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element

        return _SolidFill(solidFill_elm), expected_transparency

    @pytest.fixture(params=[0.0, 0.1, 0.33, 0.67, 0.9, 1.0])
    def solid_fill_transparency_set_fixture(self, request):
        transparency = request.param
        xClr_bldr = an_srgbClr().with_val("00FF00")
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element

        return _SolidFill(solidFill_elm), transparency

    @pytest.fixture
    def solid_fill_obj(self):
        xClr_bldr = an_srgbClr().with_val("0000FF")
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element

        return _SolidFill(solidFill_elm)


class DescribeTransparencyIntegration(object):
    """Integration tests for transparency across the entire stack."""

    def it_works_end_to_end_with_solid_fill_and_color_format(self):
        xClr_bldr = an_srgbClr().with_val("FF00FF")
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        solid_fill = _SolidFill(solidFill_elm)

        solid_fill.transparency = 0.4

        assert abs(solid_fill.transparency - 0.4) < 0.001
        assert abs(solid_fill.fore_color.transparency - 0.4) < 0.001

        alpha_elm = solid_fill.fore_color._color._xClr.alpha
        assert alpha_elm is not None
        assert abs(alpha_elm.val - 0.6) < 0.001

    def it_handles_transparency_removal_correctly(self):
        xClr_bldr = an_srgbClr().with_val("FFFF00").with_child(a_alpha().with_val(40000))
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        solid_fill = _SolidFill(solidFill_elm)

        assert abs(solid_fill.transparency - 0.6) < 0.001
        assert solid_fill.fore_color._color._xClr.alpha is not None

        solid_fill.transparency = 0.0

        assert solid_fill.transparency == 0.0
        assert solid_fill.fore_color._color._xClr.alpha is None

    def it_handles_color_format_transparency_directly(self):
        xClr_bldr = an_srgbClr().with_val("00FFFF")
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        color_format = ColorFormat.from_colorchoice_parent(solidFill_elm)

        for transparency in (0.0, 0.25, 0.5, 0.75, 1.0):
            color_format.transparency = transparency
            assert abs(color_format.transparency - transparency) < 0.001

            if transparency == 0.0:
                assert color_format._color._xClr.alpha is None
            else:
                alpha_elm = color_format._color._xClr.alpha
                assert alpha_elm is not None
                expected_alpha = 1.0 - transparency
                assert abs(alpha_elm.val - expected_alpha) < 0.001
