"""Unit-test suite for `pptx.parts.chartex` module."""

from __future__ import annotations

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import OpcPackage
from pptx.opc.packuri import PackURI
from pptx.oxml.chart.chartex import CT_ChartSpace
from pptx.parts.chartex import (
    _CHART_COLOR_STYLE_XML,
    _CHART_STYLE_XML,
    ChartColorStylePart,
    ChartExPart,
    ChartStylePart,
)

from ..unitutil.mock import instance_mock, method_mock


class DescribeChartExPart:
    """Unit-test suite for `pptx.parts.chartex.ChartExPart` objects."""

    def it_can_construct_a_new_chartex_part(self, request):
        package_ = instance_mock(request, OpcPackage)
        package_.next_partname.return_value = PackURI("/ppt/charts/chartEx42.xml")
        chartex_part_ = instance_mock(request, ChartExPart, spec_set=False)
        load_ = method_mock(
            request, ChartExPart, "load", autospec=False, return_value=chartex_part_
        )
        ct_chartspace_ = instance_mock(request, CT_ChartSpace)
        CT_ChartSpace_new_ = method_mock(
            request, CT_ChartSpace, "new", autospec=False, return_value=ct_chartspace_
        )
        style_part_ = instance_mock(request, ChartStylePart)
        ChartStylePart_new_ = method_mock(
            request, ChartStylePart, "new", autospec=False, return_value=style_part_
        )
        color_style_part_ = instance_mock(request, ChartColorStylePart)
        ChartColorStylePart_new_ = method_mock(
            request,
            ChartColorStylePart,
            "new",
            autospec=False,
            return_value=color_style_part_,
        )

        chartex_part = ChartExPart.new(package_)

        package_.next_partname.assert_called_once_with(ChartExPart.partname_template)
        load_.assert_called_once_with(
            "/ppt/charts/chartEx42.xml",
            CT.OFC_CHART_EX,
            package_,
            b"<cx:chartSpace xmlns:cx='http://schemas.microsoft.com/office/drawing/2014/chartex' "
            b"xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' "
            b"xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'/>",
        )
        CT_ChartSpace_new_.assert_called_once_with()
        assert chartex_part_._element is ct_chartspace_
        ChartStylePart_new_.assert_called_once_with(package_)
        chartex_part_.relate_to.assert_any_call(style_part_, RT.CHART_STYLE)
        ChartColorStylePart_new_.assert_called_once_with(package_)
        chartex_part_.relate_to.assert_any_call(color_style_part_, RT.CHART_COLOR_STYLE)
        assert chartex_part is chartex_part_


class DescribeChartStylePart:
    """Unit-test suite for `pptx.parts.chartex.ChartStylePart` objects."""

    def it_can_construct_a_new_chart_style_part(self, request):
        package_ = instance_mock(request, OpcPackage)
        package_.next_partname.return_value = PackURI("/ppt/charts/style42.xml")
        style_part_ = instance_mock(request, ChartStylePart)
        load_ = method_mock(
            request, ChartStylePart, "load", autospec=False, return_value=style_part_
        )

        style_part = ChartStylePart.new(package_)

        package_.next_partname.assert_called_once_with(ChartStylePart.partname_template)
        load_.assert_called_once_with(
            "/ppt/charts/style42.xml",
            CT.OFC_CHART_STYLE,
            package_,
            _CHART_STYLE_XML,
        )
        assert style_part is style_part_


class DescribeChartColorStylePart:
    """Unit-test suite for `pptx.parts.chartex.ChartColorStylePart` objects."""

    def it_can_construct_a_new_chart_color_style_part(self, request):
        package_ = instance_mock(request, OpcPackage)
        package_.next_partname.return_value = PackURI("/ppt/charts/colors42.xml")
        color_style_part_ = instance_mock(request, ChartColorStylePart)
        load_ = method_mock(
            request,
            ChartColorStylePart,
            "load",
            autospec=False,
            return_value=color_style_part_,
        )

        color_style_part = ChartColorStylePart.new(package_)

        package_.next_partname.assert_called_once_with(ChartColorStylePart.partname_template)
        load_.assert_called_once_with(
            "/ppt/charts/colors42.xml",
            CT.OFC_CHART_COLORS,
            package_,
            _CHART_COLOR_STYLE_XML,
        )
        assert color_style_part is color_style_part_
