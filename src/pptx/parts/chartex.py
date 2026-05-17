"""Chart Extension part objects."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.chart.chartex import ChartEx
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.oxml.chart.chartex import CT_ChartSpace
from pptx.parts.embeddedpackage import EmbeddedXlsxPart
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.chart.data import ChartData
    from pptx.package import Package


class ChartExPart(XmlPart):
    """A chart extension part.

    Corresponds to parts having partnames matching ppt/charts/chartEx[1-9][0-9]*.xml
    """

    partname_template = "/ppt/charts/chartEx%d.xml"

    @classmethod
    def new(cls, package: Package, chart_type: str = "waterfall"):
        """Return new |ChartExPart| instance added to `package`."""
        chartex_part = cls.load(
            package.next_partname(cls.partname_template),
            CT.OFC_CHART_EX,
            package,
            b"<cx:chartSpace xmlns:cx='http://schemas.microsoft.com/office/drawing/2014/chartex' "
            b"xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' "
            b"xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'/>",
        )
        # Initialize default chart space
        chartex_part._element = CT_ChartSpace.new()

        # Add chart style and color style parts
        style_part = ChartStylePart.new(package)
        chartex_part.relate_to(style_part, RT.CHART_STYLE)

        color_style_part = ChartColorStylePart.new(package)
        chartex_part.relate_to(color_style_part, RT.CHART_COLOR_STYLE)

        return chartex_part

    def add_excel_data(self, chart_data: ChartData):
        """Add Excel workbook data from the specified chart data object."""
        xlsx_blob = chart_data.xlsx_blob
        self.chartex_workbook.update_from_xlsx_blob(xlsx_blob)

    @lazyproperty
    def chartex(self) -> ChartEx:
        """The |ChartEx| object representing this chart extension."""
        return ChartEx(self._element, self)

    @lazyproperty
    def chartex_workbook(self) -> ChartExWorkbook:
        """
        The |ChartExWorkbook| object providing access to the external chart
        data in a linked or embedded Excel workbook.
        """
        return ChartExWorkbook(self._element, self)


class ChartStylePart(XmlPart):
    """A chart style part (style1.xml).

    Contains chart styling information for ChartEx charts.
    """

    partname_template = "/ppt/charts/style%d.xml"

    @classmethod
    def new(cls, package: Package) -> ChartStylePart:
        """Return new |ChartStylePart| with default chart style XML."""
        return cls.load(
            package.next_partname(cls.partname_template),
            CT.OFC_CHART_STYLE,
            package,
            _CHART_STYLE_XML,
        )


class ChartColorStylePart(XmlPart):
    """A chart color style part (colors1.xml).

    Contains chart color styling information for ChartEx charts.
    """

    partname_template = "/ppt/charts/colors%d.xml"

    @classmethod
    def new(cls, package: Package) -> ChartColorStylePart:
        """Return new |ChartColorStylePart| with default chart color style XML."""
        return cls.load(
            package.next_partname(cls.partname_template),
            CT.OFC_CHART_COLORS,
            package,
            _CHART_COLOR_STYLE_XML,
        )


class ChartExWorkbook(object):
    """Provides access to external chart data in a linked or embedded Excel workbook for ChartEx."""

    def __init__(self, chartSpace, chartex_part):
        super(ChartExWorkbook, self).__init__()
        self._chartSpace = chartSpace
        self._chartex_part = chartex_part

    def update_from_xlsx_blob(self, xlsx_blob):
        """
        Replace the Excel spreadsheet in the related |EmbeddedXlsxPart| with
        the Excel binary in *xlsx_blob*, adding a new |EmbeddedXlsxPart| if
        there isn't one.
        """
        xlsx_part = self.xlsx_part
        if xlsx_part is None:
            self.xlsx_part = EmbeddedXlsxPart.new(xlsx_blob, self._chartex_part.package)
            return
        xlsx_part.blob = xlsx_blob

    @property
    def xlsx_part(self):
        """Optional |EmbeddedXlsxPart| object containing data for this chart extension.

        This related part has its rId at `cx:chartSpace/cx:chartData/cx:externalData/@r:id`.
        This value is |None| if there is no `<cx:externalData>` element.
        """

        xlsx_part_rId = self._chartSpace.xlsx_part_rId
        return None if xlsx_part_rId is None else self._chartex_part.related_part(xlsx_part_rId)

    @xlsx_part.setter
    def xlsx_part(self, xlsx_part):
        """
        Set the related |EmbeddedXlsxPart| to *xlsx_part*. Assume one does
        not already exist.
        """
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        from pptx.oxml.ns import qn

        rId = self._chartex_part.relate_to(xlsx_part, RT.PACKAGE)
        # Add externalData element to chartData
        externalData = self._chartSpace.get_or_add_externalData()
        externalData.set(qn("r:id"), rId)
        # Remove any namespaced autoUpdate to avoid duplicate attributes
        cx_ns = "http://schemas.microsoft.com/office/drawing/2014/chartex"
        namespaced_key = f"{{{cx_ns}}}autoUpdate"
        if namespaced_key in externalData.attrib:
            del externalData.attrib[namespaced_key]
        externalData.set("autoUpdate", "0")


_CHART_STYLE_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    b'<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"'
    b' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="395">'
    b'<cs:axisTitle><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/>'
    b'<cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/>'
    b'<a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="1197"/>'
    b'</cs:axisTitle><cs:categoryAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1">'
    b'<a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef>'
    b'<cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
    b'<a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/>'
    b"</a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr>"
    b'<cs:defRPr sz="1197"/></cs:categoryAxis>'
    b'<cs:chartArea mods="allowNoFillOverride allowNoLineOverride">'
    b'<cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/>'
    b'<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>'
    b'<cs:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill>'
    b'<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
    b'<a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/>'
    b"</a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr>"
    b'<cs:defRPr sz="1330"/></cs:chartArea>'
    b'<cs:dataLabel><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1">'
    b'<a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef>'
    b'<cs:defRPr sz="1197"/></cs:dataLabel>'
    b'<cs:dataLabelCallout><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="dk1">'
    b'<a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef>'
    b'<cs:spPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln>'
    b'<a:solidFill><a:schemeClr val="dk1"><a:lumMod val="25000"/>'
    b'<a:lumOff val="75000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr>'
    b'<cs:defRPr sz="1197"/>'
    b'<cs:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip"'
    b' horzOverflow="clip" vert="horz" wrap="square" lIns="36576"'
    b' tIns="18288" rIns="36576" bIns="18288" anchor="ctr" anchorCtr="1">'
    b"<a:spAutoFit/></cs:bodyPr></cs:dataLabelCallout>"
    b'<cs:dataPoint><cs:lnRef idx="0"/><cs:fillRef idx="0">'
    b'<cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/>'
    b'<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>'
    b'<cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    b"</cs:spPr></cs:dataPoint>"
    b'<cs:dataPoint3D><cs:lnRef idx="0"/><cs:fillRef idx="0">'
    b'<cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/>'
    b'<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>'
    b'<cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    b"</cs:spPr></cs:dataPoint3D>"
    b'<cs:dataPointLine><cs:lnRef idx="0"><cs:styleClr val="auto"/>'
    b'</cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/>'
    b'<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>'
    b'<cs:spPr><a:ln w="28575" cap="rnd"><a:solidFill>'
    b'<a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln>'
    b"</cs:spPr></cs:dataPointLine>"
    b'<cs:dataPointMarker><cs:lnRef idx="0"/><cs:fillRef idx="0">'
    b'<cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/>'
    b'<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>'
    b'<cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    b'<a:ln w="9525"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill>'
    b"</a:ln></cs:spPr></cs:dataPointMarker>"
    b'<cs:dataPointMarkerLayout symbol="circle" size="5"/>'
    b'<cs:dataPointWireframe><cs:lnRef idx="0">'
    b'<cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor">'
    b'<a:schemeClr val="tx1"/></cs:fontRef><cs:spPr>'
    b'<a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/>'
    b"</a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointWireframe>"
    b'<cs:dataTable><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1">'
    b'<a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>'
    b'</cs:fontRef><cs:spPr><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1">'
    b'<a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr>'
    b'</a:solidFill></a:ln></cs:spPr><cs:defRPr sz="1197"/></cs:dataTable>'
    b'<cs:downBar><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor">'
    b'<a:schemeClr val="dk1"/></cs:fontRef><cs:spPr><a:solidFill>'
    b'<a:schemeClr val="dk1"><a:lumMod val="65000"/><a:lumOff val="35000"/>'
    b'</a:schemeClr></a:solidFill><a:ln w="9525"><a:solidFill>'
    b'<a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/>'
    b"</a:schemeClr></a:solidFill></a:ln></cs:spPr></cs:downBar>"
    b'<cs:dropLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor">'
    b'<a:schemeClr val="tx1"/></cs:fontRef><cs:spPr>'
    b'<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
    b'<a:schemeClr val="tx1"><a:lumMod val="35000"/><a:lumOff val="65000"/>'
    b"</a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:dropLine>"
    b'<cs:errorBar><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor">'
    b'<a:schemeClr val="tx1"/></cs:fontRef><cs:spPr>'
    b'<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
    b'<a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/>'
    b"</a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:errorBar>"
    b'<cs:floor><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor">'
    b'<a:schemeClr val="tx1"/></cs:fontRef></cs:floor>'
    b'<cs:gridlineMajor><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor">'
    b'<a:schemeClr val="tx1"/></cs:fontRef><cs:spPr>'
    b'<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
    b'<a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/>'
    b"</a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr>"
    b"</cs:gridlineMajor>"
    b'<cs:gridlineMinor><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor">'
    b'<a:schemeClr val="tx1"/></cs:fontRef><cs:spPr>'
    b'<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
    b'<a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/>'
    b"</a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr>"
    b"</cs:gridlineMinor>"
    b'<cs:hiLoLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor">'
    b'<a:schemeClr val="tx1"/></cs:fontRef><cs:spPr>'
    b'<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
    b'<a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/>'
    b"</a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:hiLoLine>"
    b'<cs:leaderLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor">'
    b'<a:schemeClr val="tx1"/></cs:fontRef><cs:spPr>'
    b'<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
    b'<a:schemeClr val="tx1"><a:lumMod val="35000"/><a:lumOff val="65000"/>'
    b"</a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:leaderLine>"
    b'<cs:legend><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1">'
    b'<a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>'
    b'</cs:fontRef><cs:defRPr sz="1197"/></cs:legend>'
    b'<cs:plotArea mods="allowNoFillOverride allowNoLineOverride">'
    b'<cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/>'
    b'<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>'
    b"</cs:plotArea>"
    b'<cs:plotArea3D mods="allowNoFillOverride allowNoLineOverride">'
    b'<cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/>'
    b'<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>'
    b"</cs:plotArea3D>"
    b'<cs:seriesAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1">'
    b'<a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>'
    b'</cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
    b'<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/>'
    b'<a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln>'
    b'</cs:spPr><cs:defRPr sz="1197"/></cs:seriesAxis>'
    b'<cs:seriesLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor">'
    b'<a:schemeClr val="tx1"/></cs:fontRef><cs:spPr>'
    b'<a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="D9D9D9"/>'
    b"</a:solidFill><a:round/></a:ln></cs:spPr></cs:seriesLine>"
    b'<cs:title><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1">'
    b'<a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>'
    b'</cs:fontRef><cs:defRPr sz="1862"/></cs:title>'
    b'<cs:trendline><cs:lnRef idx="0"><cs:styleClr val="auto"/>'
    b'</cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/>'
    b'<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>'
    b'<cs:spPr><a:ln w="19050" cap="rnd"><a:solidFill>'
    b'<a:schemeClr val="phClr"/></a:solidFill>'
    b'<a:prstDash val="sysDash"/></a:ln></cs:spPr></cs:trendline>'
    b'<cs:trendlineLabel><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1">'
    b'<a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>'
    b'</cs:fontRef><cs:defRPr sz="1197"/></cs:trendlineLabel>'
    b'<cs:upBar><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor">'
    b'<a:schemeClr val="dk1"/></cs:fontRef><cs:spPr><a:solidFill>'
    b'<a:schemeClr val="lt1"/></a:solidFill><a:ln w="9525"><a:solidFill>'
    b'<a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/>'
    b"</a:schemeClr></a:solidFill></a:ln></cs:spPr></cs:upBar>"
    b'<cs:valueAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1">'
    b'<a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>'
    b'</cs:fontRef><cs:defRPr sz="1197"/></cs:valueAxis>'
    b'<cs:wall><cs:lnRef idx="0"/><cs:fillRef idx="0"/>'
    b'<cs:effectRef idx="0"/><cs:fontRef idx="minor">'
    b'<a:schemeClr val="tx1"/></cs:fontRef></cs:wall></cs:chartStyle>'
)

_CHART_COLOR_STYLE_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    b'<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"'
    b' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">'
    b'<a:schemeClr val="accent1"/><a:schemeClr val="accent2"/>'
    b'<a:schemeClr val="accent3"/><a:schemeClr val="accent4"/>'
    b'<a:schemeClr val="accent5"/><a:schemeClr val="accent6"/>'
    b"<cs:variation/>"
    b'<cs:variation><a:lumMod val="60000"/></cs:variation>'
    b'<cs:variation><a:lumMod val="80000"/><a:lumOff val="20000"/></cs:variation>'
    b'<cs:variation><a:lumMod val="80000"/></cs:variation>'
    b'<cs:variation><a:lumMod val="60000"/><a:lumOff val="40000"/></cs:variation>'
    b'<cs:variation><a:lumMod val="50000"/></cs:variation>'
    b'<cs:variation><a:lumMod val="70000"/><a:lumOff val="30000"/></cs:variation>'
    b'<cs:variation><a:lumMod val="70000"/></cs:variation>'
    b'<cs:variation><a:lumMod val="50000"/><a:lumOff val="50000"/></cs:variation>'
    b"</cs:colorStyle>"
)
