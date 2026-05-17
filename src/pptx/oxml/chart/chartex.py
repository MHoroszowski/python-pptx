"""XML elements for chart extensions."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import XsdBoolean, XsdInt, XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    pass


class CT_ChartSpace(BaseOxmlElement):
    """
    ``<cx:chartSpace>`` element, the root element of a chartex part.
    """

    chartData: CT_ChartData = OneAndOnlyOne("cx:chartData")  # pyright: ignore
    chart: CT_Chart = OneAndOnlyOne("cx:chart")  # pyright: ignore
    spPr = ZeroOrOne("cx:spPr")  # Shape properties
    txPr = ZeroOrOne("cx:txPr")  # Text properties
    extLst = ZeroOrOne("cx:extLst")  # Extension list

    @classmethod
    def new(cls):
        """
        Return a new ``<cx:chartSpace>`` element
        """
        from pptx.oxml import parse_xml

        xml = (
            b"<cx:chartSpace "
            b'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
            b'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            b"</cx:chartSpace>"
        )
        chartSpace = cast(CT_ChartSpace, parse_xml(xml))
        chartData = CT_ChartData.new()
        chart = CT_Chart.new()
        chartSpace.append(chartData)
        chartSpace.append(chart)
        return chartSpace

    @property
    def xlsx_part_rId(self) -> str | None:
        """
        The rId of the embedded Excel part relationship, or None if no
        embedded Excel part is present.
        """
        externalData = self.chartData.find(qn("cx:externalData"))
        if externalData is None:
            return None
        rId = externalData.get(qn("r:id"))
        if rId is None:
            return None
        return rId

    def get_or_add_externalData(self):
        """
        Return the <cx:externalData> child element, newly created if not
        present.
        """
        return self.chartData.get_or_add_externalData()


class CT_ChartData(BaseOxmlElement):
    """
    ``<cx:chartData>`` element, container for chart data and external data reference.
    """

    externalData: CT_ExternalData = ZeroOrOne("cx:externalData")  # pyright: ignore
    data = ZeroOrMore("cx:data")  # pyright: ignore
    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore

    @classmethod
    def new(cls):
        """Return a new <cx:chartData> element."""
        from lxml import etree

        from pptx.oxml import parse_xml

        xml = b'<cx:chartData xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>'
        chartData = cast(CT_ChartData, parse_xml(xml))
        # Add a default data element with id="0"
        data_elem = etree.SubElement(chartData, qn("cx:data"))
        data_elem.set("id", "0")
        return chartData

    def get_or_add_externalData(self):
        """Return the <cx:externalData> child element, newly created if not present."""
        from lxml import etree

        externalData = self.find(qn("cx:externalData"))
        if externalData is None:
            externalData = etree.SubElement(self, qn("cx:externalData"))
            # Move it to the beginning
            self.remove(externalData)
            self.insert(0, externalData)
        return externalData


class CT_ExternalData(BaseOxmlElement):
    """
    ``<cx:externalData>`` element, refers to external Excel data.
    """

    autoUpdate: bool = OptionalAttribute("autoUpdate", XsdBoolean)  # pyright: ignore
    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore


class CT_Chart(BaseOxmlElement):
    """
    ``<cx:chart>`` element, container for chart elements.
    """

    title: CT_ChartTitle = ZeroOrOne("cx:title")  # pyright: ignore
    plotArea: CT_PlotArea = OneAndOnlyOne("cx:plotArea")  # pyright: ignore
    legend: CT_Legend = ZeroOrOne("cx:legend")  # pyright: ignore
    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore

    @classmethod
    def new(cls):
        """Return a new <cx:chart> element."""
        from pptx.oxml import parse_xml

        xml = b'<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>'
        chart = cast(CT_Chart, parse_xml(xml))
        plotArea = CT_PlotArea.new()
        chart.append(plotArea)
        return chart

    def get_or_add_legend(self):
        """Return the <cx:legend> child element, newly created if not present."""
        from lxml import etree

        legend = self.find(qn("cx:legend"))
        if legend is None:
            legend = etree.SubElement(self, qn("cx:legend"))
            # Insert legend after plotArea
            plot_area_idx = list(self).index(self.plotArea)
            self.remove(legend)
            self.insert(plot_area_idx + 1, legend)
            # Set default attributes
            legend.set("pos", "t")
            legend.set("align", "ctr")
            legend.set("overlay", "0")
        return legend

    def get_or_add_title(self):
        """Return the <cx:title> child element, newly created if not present."""
        from lxml import etree

        title = self.find(qn("cx:title"))
        if title is None:
            title = etree.SubElement(self, qn("cx:title"))
            # Move title to the beginning
            self.remove(title)
            self.insert(0, title)
            # Set default attributes
            title.set("pos", "t")
            title.set("align", "ctr")
            title.set("overlay", "0")
        return title


class CT_PlotArea(BaseOxmlElement):
    """
    ``<cx:plotArea>`` element, container for chart plot area.
    """

    plotAreaRegion: CT_PlotAreaRegion = OneAndOnlyOne("cx:plotAreaRegion")  # pyright: ignore
    axis = ZeroOrMore("cx:axis")  # pyright: ignore
    spPr = ZeroOrOne("cx:spPr")  # pyright: ignore
    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore

    @classmethod
    def new(cls):
        """Return a new <cx:plotArea> element with default axes."""
        from pptx.oxml import parse_xml

        xml = b'<cx:plotArea xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>'
        plotArea = cast(CT_PlotArea, parse_xml(xml))
        plotAreaRegion = CT_PlotAreaRegion.new()
        plotArea.append(plotAreaRegion)
        # Add category axis (id=0)
        cat_axis = CT_Axis.new_cat_axis(0)
        plotArea.append(cat_axis)
        # Add value axis (id=1)
        val_axis = CT_Axis.new_val_axis(1)
        plotArea.append(val_axis)
        return plotArea


class CT_PlotAreaRegion(BaseOxmlElement):
    """
    ``<cx:plotAreaRegion>`` element, container for a plot area region.
    """

    plotSurface = ZeroOrOne("cx:plotSurface")  # pyright: ignore
    series = ZeroOrMore("cx:series")  # pyright: ignore
    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore

    @classmethod
    def new(cls):
        """Return a new <cx:plotAreaRegion> element."""
        from pptx.oxml import parse_xml

        xml = b'<cx:plotAreaRegion xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>'
        return cast(CT_PlotAreaRegion, parse_xml(xml))

    def add_waterfall_series(
        self, series_name: str, data_id: int = 0, subtotal_indices: list[int] | None = None
    ):
        """Add a waterfall series to this plot area region."""
        import uuid

        from lxml import etree

        series = etree.SubElement(self, qn("cx:series"))
        series.set("layoutId", "waterfall")
        series.set("uniqueId", f"{{{uuid.uuid4()}}}")

        # Add series title
        tx = etree.SubElement(series, qn("cx:tx"))
        txData = etree.SubElement(tx, qn("cx:txData"))
        f_elem = etree.SubElement(txData, qn("cx:f"))
        f_elem.text = "Sheet1!$B$1"
        v_elem = etree.SubElement(txData, qn("cx:v"))
        v_elem.text = series_name

        # Add data labels
        dataLabels = etree.SubElement(series, qn("cx:dataLabels"))
        dataLabels.set("pos", "outEnd")
        visibility = etree.SubElement(dataLabels, qn("cx:visibility"))
        visibility.set("seriesName", "0")
        visibility.set("categoryName", "0")
        visibility.set("value", "1")

        # Add data id
        dataId_elem = etree.SubElement(series, qn("cx:dataId"))
        dataId_elem.set("val", str(data_id))

        # Add layout properties with subtotals
        if subtotal_indices:
            layoutPr = etree.SubElement(series, qn("cx:layoutPr"))
            subtotals = etree.SubElement(layoutPr, qn("cx:subtotals"))
            for idx in subtotal_indices:
                idx_elem = etree.SubElement(subtotals, qn("cx:idx"))
                idx_elem.set("val", str(idx))

        return series

    def _new_cx_series(self, layout_id: str, series_name: str, data_id: int = 0):
        """Append a `<cx:series>` skeleton (tx, dataLabels, dataId) and return it.

        Shared by every ChartEx series layout. The `<cx:layoutPr>` (and any
        extra series) is added by the per-type method after this returns, so
        the XSD `CT_Series` element order (tx, dataLabels, dataId, layoutPr) is
        always respected.
        """
        import uuid

        from lxml import etree

        series = etree.SubElement(self, qn("cx:series"))
        series.set("layoutId", layout_id)
        series.set("uniqueId", f"{{{uuid.uuid4()}}}")

        tx = etree.SubElement(series, qn("cx:tx"))
        txData = etree.SubElement(tx, qn("cx:txData"))
        f_elem = etree.SubElement(txData, qn("cx:f"))
        f_elem.text = "Sheet1!$B$1"
        v_elem = etree.SubElement(txData, qn("cx:v"))
        v_elem.text = series_name

        dataLabels = etree.SubElement(series, qn("cx:dataLabels"))
        dataLabels.set("pos", "outEnd")
        visibility = etree.SubElement(dataLabels, qn("cx:visibility"))
        visibility.set("seriesName", "0")
        visibility.set("categoryName", "0")
        visibility.set("value", "1")

        dataId_elem = etree.SubElement(series, qn("cx:dataId"))
        dataId_elem.set("val", str(data_id))
        return series

    def add_treemap_series(self, series_name: str, data_id: int = 0):
        """Add a treemap series (`layoutId="treemap"`)."""
        from lxml import etree

        series = self._new_cx_series("treemap", series_name, data_id)
        layoutPr = etree.SubElement(series, qn("cx:layoutPr"))
        pll = etree.SubElement(layoutPr, qn("cx:parentLabelLayout"))
        pll.set("val", "banner")
        return series

    def add_sunburst_series(self, series_name: str, data_id: int = 0):
        """Add a sunburst series (`layoutId="sunburst"`)."""
        from lxml import etree

        series = self._new_cx_series("sunburst", series_name, data_id)
        layoutPr = etree.SubElement(series, qn("cx:layoutPr"))
        pll = etree.SubElement(layoutPr, qn("cx:parentLabelLayout"))
        pll.set("val", "none")
        return series

    def add_funnel_series(self, series_name: str, data_id: int = 0):
        """Add a funnel series (`layoutId="funnel"`)."""
        return self._new_cx_series("funnel", series_name, data_id)

    def add_box_whisker_series(self, series_name: str, data_id: int = 0):
        """Add a box & whisker series (`layoutId="boxWhisker"`)."""
        from lxml import etree

        series = self._new_cx_series("boxWhisker", series_name, data_id)
        layoutPr = etree.SubElement(series, qn("cx:layoutPr"))
        vis = etree.SubElement(layoutPr, qn("cx:visibility"))
        vis.set("connectorLines", "1")
        vis.set("meanLine", "0")
        vis.set("meanMarker", "1")
        vis.set("nonoutliers", "0")
        vis.set("outliers", "1")
        stats = etree.SubElement(layoutPr, qn("cx:statistics"))
        stats.set("quartileMethod", "exclusive")
        return series

    def add_histogram_series(
        self,
        series_name: str,
        data_id: int = 0,
        bin_count: int | None = None,
        bin_size: float | None = None,
    ):
        """Add a histogram series (`layoutId="clusteredColumn"` + `<cx:binning>`)."""
        from lxml import etree

        series = self._new_cx_series("clusteredColumn", series_name, data_id)
        layoutPr = etree.SubElement(series, qn("cx:layoutPr"))
        binning = etree.SubElement(layoutPr, qn("cx:binning"))
        binning.set("intervalClosed", "r")
        # CT_Binning child is an xsd:choice of binSize | binCount (both optional →
        # automatic binning when neither given).
        if bin_size is not None:
            bs = etree.SubElement(binning, qn("cx:binSize"))
            bs.text = str(bin_size)
        elif bin_count is not None:
            bc = etree.SubElement(binning, qn("cx:binCount"))
            bc.text = str(bin_count)
        return series

    def add_pareto_series(
        self,
        series_name: str,
        data_id: int = 0,
        bin_count: int | None = None,
        bin_size: float | None = None,
    ):
        """Add the Pareto pair: a binned `clusteredColumn` + a `paretoLine`.

        Both series reference the same `data_id`. Returns the histogram series.
        """
        from lxml import etree

        hist = self.add_histogram_series(series_name, data_id, bin_count, bin_size)
        # paretoLine cumulative series over the same data
        line = self._new_cx_series("paretoLine", series_name, data_id)
        layoutPr = etree.SubElement(line, qn("cx:layoutPr"))
        etree.SubElement(layoutPr, qn("cx:aggregation"))
        return hist

    def add_hierarchical_string_dimension(
        self, data_elem, dim_type: str, formula: str, levels: list[list[str]]
    ):
        """Append a `<cx:strDim>` with one `<cx:lvl>` per hierarchy level.

        `levels` is outermost-first; each inner list is that level's labels.
        Used by treemap/sunburst. `ptCount` on every `<cx:lvl>` equals its
        actual point count (off-by-one is a PowerPoint-repair trigger).
        """
        from lxml import etree

        strDim = etree.SubElement(data_elem, qn("cx:strDim"))
        strDim.set("type", dim_type)
        f_elem = etree.SubElement(strDim, qn("cx:f"))
        f_elem.text = formula
        for labels in levels:
            lvl = etree.SubElement(strDim, qn("cx:lvl"))
            lvl.set("ptCount", str(len(labels)))
            for idx, value in enumerate(labels):
                if value is None or value == "":
                    continue
                pt = etree.SubElement(lvl, qn("cx:pt"))
                pt.set("idx", str(idx))
                pt.text = value
        return strDim


class CT_Series(BaseOxmlElement):
    """
    ``<cx:series>`` element, container for a chart series.
    """

    tx = ZeroOrOne("cx:tx")  # pyright: ignore
    spPr = ZeroOrOne("cx:spPr")  # pyright: ignore
    valueColors = ZeroOrOne("cx:valueColors")  # pyright: ignore
    valueColorPositions = ZeroOrOne("cx:valueColorPositions")  # pyright: ignore
    dataPt = ZeroOrMore("cx:dataPt")  # pyright: ignore
    dataLabels = ZeroOrOne("cx:dataLabels")  # pyright: ignore
    dataId = ZeroOrOne("cx:dataId")  # pyright: ignore
    layoutPr = ZeroOrOne("cx:layoutPr")  # pyright: ignore
    axisId = ZeroOrMore("cx:axisId")  # pyright: ignore
    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore

    layoutId: XsdString = RequiredAttribute("layoutId", XsdString)  # pyright: ignore
    hidden: bool = OptionalAttribute("hidden", XsdBoolean, False)  # pyright: ignore
    ownerIdx: int = OptionalAttribute("ownerIdx", XsdInt)  # pyright: ignore
    uniqueId: str = OptionalAttribute("uniqueId", XsdString)  # pyright: ignore
    formatIdx: int = OptionalAttribute("formatIdx", XsdInt)  # pyright: ignore


class CT_Axis(BaseOxmlElement):
    """
    ``<cx:axis>`` element, represents an axis in the chart.
    """

    catScaling: CT_CategoryAxisScaling = ZeroOrOne("cx:catScaling")  # pyright: ignore
    valScaling: CT_ValueAxisScaling = ZeroOrOne("cx:valScaling")  # pyright: ignore
    title = ZeroOrOne("cx:title")  # pyright: ignore
    units = ZeroOrOne("cx:units")  # pyright: ignore
    majorGridlines: CT_Gridlines = ZeroOrOne("cx:majorGridlines")  # pyright: ignore
    minorGridlines = ZeroOrOne("cx:minorGridlines")  # pyright: ignore
    majorTickMarks = ZeroOrOne("cx:majorTickMarks")  # pyright: ignore
    minorTickMarks = ZeroOrOne("cx:minorTickMarks")  # pyright: ignore
    tickLabels: CT_TickLabels = ZeroOrOne("cx:tickLabels")  # pyright: ignore
    numFmt = ZeroOrOne("cx:numFmt")  # pyright: ignore
    spPr = ZeroOrOne("cx:spPr")  # pyright: ignore
    txPr = ZeroOrOne("cx:txPr")  # pyright: ignore
    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore

    id: int = RequiredAttribute("id", XsdInt)  # pyright: ignore
    hidden: bool = OptionalAttribute("hidden", XsdBoolean, False)  # pyright: ignore

    @classmethod
    def new_cat_axis(cls, axis_id: int):
        """Return a new category axis element."""
        from lxml import etree

        from pptx.oxml import parse_xml

        xml = b'<cx:axis xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>'
        axis = cast(CT_Axis, parse_xml(xml))
        axis.set("id", str(axis_id))
        # Add catScaling element
        catScaling = etree.SubElement(axis, qn("cx:catScaling"))
        catScaling.set("gapWidth", "0.5")
        # Add tickLabels
        etree.SubElement(axis, qn("cx:tickLabels"))
        return axis

    @classmethod
    def new_val_axis(cls, axis_id: int):
        """Return a new value axis element."""
        from lxml import etree

        from pptx.oxml import parse_xml

        xml = b'<cx:axis xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>'
        axis = cast(CT_Axis, parse_xml(xml))
        axis.set("id", str(axis_id))
        # Add valScaling element
        etree.SubElement(axis, qn("cx:valScaling"))
        # Add majorGridlines
        etree.SubElement(axis, qn("cx:majorGridlines"))
        # Add tickLabels
        etree.SubElement(axis, qn("cx:tickLabels"))
        return axis


class CT_Text(BaseOxmlElement):
    """
    ``<cx:tx>`` element, container for series text.
    """

    txData = ZeroOrOne("cx:txData")  # pyright: ignore


class CT_TextData(BaseOxmlElement):
    """
    ``<cx:txData>`` element, contains series text data.
    """

    f = ZeroOrOne("cx:f")  # pyright: ignore
    v = ZeroOrOne("cx:v")  # pyright: ignore


class CT_DataLabels(BaseOxmlElement):
    """
    ``<cx:dataLabels>`` element, container for data labels.
    """

    numFmt = ZeroOrOne("cx:numFmt")  # pyright: ignore
    spPr = ZeroOrOne("cx:spPr")  # pyright: ignore
    txPr = ZeroOrOne("cx:txPr")  # pyright: ignore
    visibility = ZeroOrOne("cx:visibility")  # pyright: ignore
    separator = ZeroOrOne("cx:separator")  # pyright: ignore
    dataLabel = ZeroOrMore("cx:dataLabel")  # pyright: ignore
    dataLabelHidden = ZeroOrMore("cx:dataLabelHidden")  # pyright: ignore
    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore


class CT_DataId(BaseOxmlElement):
    """
    ``<cx:dataId>`` element, identifies the data for a series.
    """

    val: int = RequiredAttribute("val", XsdInt)  # pyright: ignore


class CT_Data(BaseOxmlElement):
    """
    ``<cx:data>`` element, contains data dimensions.
    """

    strDim = ZeroOrMore("cx:strDim")  # pyright: ignore
    numDim = ZeroOrMore("cx:numDim")  # pyright: ignore
    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore

    id: int = RequiredAttribute("id", XsdInt)  # pyright: ignore

    def add_string_dimension(self, dim_type: str, formula: str, values: list[str]):
        """Add a string dimension with the given type, formula, and values."""
        from lxml import etree

        strDim = etree.SubElement(self, qn("cx:strDim"))
        strDim.set("type", dim_type)

        # Add formula
        f_elem = etree.SubElement(strDim, qn("cx:f"))
        f_elem.text = formula

        # Add level with points
        lvl = etree.SubElement(strDim, qn("cx:lvl"))
        lvl.set("ptCount", str(len(values)))

        for idx, value in enumerate(values):
            pt = etree.SubElement(lvl, qn("cx:pt"))
            pt.set("idx", str(idx))
            pt.text = value

        return strDim

    def add_numeric_dimension(
        self, dim_type: str, formula: str, values: list[float | int], format_code: str = "General"
    ):
        """Add a numeric dimension with the given type, formula, and values."""
        from lxml import etree

        numDim = etree.SubElement(self, qn("cx:numDim"))
        numDim.set("type", dim_type)

        # Add formula
        f_elem = etree.SubElement(numDim, qn("cx:f"))
        f_elem.text = formula

        # Add level with points
        lvl = etree.SubElement(numDim, qn("cx:lvl"))
        lvl.set("ptCount", str(len(values)))
        lvl.set("formatCode", format_code)

        for idx, value in enumerate(values):
            pt = etree.SubElement(lvl, qn("cx:pt"))
            pt.set("idx", str(idx))
            pt.text = str(value)

        return numDim


class CT_StringDimension(BaseOxmlElement):
    """
    ``<cx:strDim>`` element, string dimension for chart data.
    """

    f = ZeroOrOne("cx:f")  # pyright: ignore
    nf = ZeroOrOne("cx:nf")  # pyright: ignore
    lvl = ZeroOrMore("cx:lvl")  # pyright: ignore

    type: XsdString = RequiredAttribute("type", XsdString)  # pyright: ignore


class CT_NumericDimension(BaseOxmlElement):
    """
    ``<cx:numDim>`` element, numeric dimension for chart data.
    """

    f = ZeroOrOne("cx:f")  # pyright: ignore
    nf = ZeroOrOne("cx:nf")  # pyright: ignore
    lvl = ZeroOrMore("cx:lvl")  # pyright: ignore

    type: XsdString = RequiredAttribute("type", XsdString)  # pyright: ignore


class CT_Formula(BaseOxmlElement):
    """
    ``<cx:f>`` element, formula reference.
    """

    dir: XsdString = OptionalAttribute("dir", XsdString, "col")  # pyright: ignore


class CT_StringLevel(BaseOxmlElement):
    """
    ``<cx:lvl>`` element for string dimensions.
    """

    pt = ZeroOrMore("cx:pt")  # pyright: ignore

    ptCount: int = RequiredAttribute("ptCount", XsdInt)  # pyright: ignore
    name: XsdString = OptionalAttribute("name", XsdString)  # pyright: ignore
    formatCode: XsdString = OptionalAttribute("formatCode", XsdString)  # pyright: ignore


class CT_StringValue(BaseOxmlElement):
    """
    ``<cx:pt>`` element for string values.
    """

    idx: int = RequiredAttribute("idx", XsdInt)  # pyright: ignore


class CT_CategoryAxisScaling(BaseOxmlElement):
    """
    ``<cx:catScaling>`` element, category axis scaling.
    """

    gapWidth: XsdString = OptionalAttribute("gapWidth", XsdString)  # pyright: ignore


class CT_ValueAxisScaling(BaseOxmlElement):
    """
    ``<cx:valScaling>`` element, value axis scaling.
    """

    max: XsdString = OptionalAttribute("max", XsdString)  # pyright: ignore
    min: XsdString = OptionalAttribute("min", XsdString)  # pyright: ignore
    majorUnit: XsdString = OptionalAttribute("majorUnit", XsdString)  # pyright: ignore
    minorUnit: XsdString = OptionalAttribute("minorUnit", XsdString)  # pyright: ignore


class CT_SeriesLayoutProperties(BaseOxmlElement):
    """
    ``<cx:layoutPr>`` element, series layout properties.
    """

    subtotals = ZeroOrOne("cx:subtotals")  # pyright: ignore
    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore


class CT_Subtotals(BaseOxmlElement):
    """
    ``<cx:subtotals>`` element, waterfall chart subtotals.
    """

    idx = ZeroOrMore("cx:idx")  # pyright: ignore


class CT_SubtotalIndex(BaseOxmlElement):
    """
    ``<cx:idx>`` element in subtotals.
    """

    val: int = RequiredAttribute("val", XsdInt)  # pyright: ignore


class CT_TickLabels(BaseOxmlElement):
    """
    ``<cx:tickLabels>`` element, axis tick labels.
    """

    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore


class CT_ChartTitle(BaseOxmlElement):
    """
    ``<cx:title>`` element for chart title.
    """

    tx = ZeroOrOne("cx:tx")  # pyright: ignore
    spPr = ZeroOrOne("cx:spPr")  # pyright: ignore
    txPr = ZeroOrOne("cx:txPr")  # pyright: ignore
    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore

    pos: XsdString = OptionalAttribute("pos", XsdString, "t")  # pyright: ignore
    align: XsdString = OptionalAttribute("align", XsdString, "ctr")  # pyright: ignore
    overlay: bool = OptionalAttribute("overlay", XsdBoolean, False)  # pyright: ignore


class CT_Legend(BaseOxmlElement):
    """
    ``<cx:legend>`` element for chart legend.
    """

    spPr = ZeroOrOne("cx:spPr")  # pyright: ignore
    txPr = ZeroOrOne("cx:txPr")  # pyright: ignore
    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore

    pos: XsdString = OptionalAttribute("pos", XsdString, "r")  # pyright: ignore
    align: XsdString = OptionalAttribute("align", XsdString, "ctr")  # pyright: ignore
    overlay: bool = OptionalAttribute("overlay", XsdBoolean, False)  # pyright: ignore


class CT_Gridlines(BaseOxmlElement):
    """
    ``<cx:majorGridlines>`` and ``<cx:minorGridlines>`` elements.
    """

    spPr = ZeroOrOne("cx:spPr")  # pyright: ignore
    extLst = ZeroOrOne("cx:extLst")  # pyright: ignore


class CT_DataLabelVisibilities(BaseOxmlElement):
    """
    ``<cx:visibility>`` element for data label visibility.
    """

    seriesName: bool = OptionalAttribute("seriesName", XsdBoolean)  # pyright: ignore
    categoryName: bool = OptionalAttribute("categoryName", XsdBoolean)  # pyright: ignore
    value: bool = OptionalAttribute("value", XsdBoolean)  # pyright: ignore
