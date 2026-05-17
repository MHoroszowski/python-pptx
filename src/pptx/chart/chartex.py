"""Chart Extension objects and related items."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.shared import ParentedElementProxy, PartElementProxy
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.chart.chartex import CT_Axis, CT_ChartSpace, CT_Series
    from pptx.parts.chartex import ChartExPart


class ChartEx(PartElementProxy):
    """Chart extension object.

    Corresponds to the ``<cx:chartSpace>`` element that is the root of a chart extension part.
    """

    _chartspace: CT_ChartSpace

    def __init__(self, chartSpace: CT_ChartSpace, chart_part: ChartExPart):
        super().__init__(chartSpace, chart_part)
        self._chartspace = chartSpace

    @property
    def chart_title(self) -> str | None:
        """The title of this chart as a string, or None if there is no title.

        Assigning a string value sets the title to that value. Assigning None causes
        any title to be deleted.
        """
        title = self._chart.title
        if title is None:
            return None

        title_text = title.find(
            ".//cx:txData/cx:v",
            namespaces={"cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"},
        )
        if title_text is None:
            return None

        return title_text.text

    @property
    def chart_type(self) -> str | None:
        """The chart type of this chart extension as a string.

        Possible values include:
        * waterfall
        * sunburst
        * treemap
        * funnel
        * boxWhisker
        * clusteredColumn
        * paretoLine
        * regionMap

        Returns None if the chart type cannot be determined.
        """
        plotAreaRegion = self._chart.plotArea.plotAreaRegion
        series = plotAreaRegion.find(
            ".//cx:series",
            namespaces={"cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"},
        )
        if series is None:
            return None

        return series.get("layoutId")

    @property
    def has_legend(self) -> bool:
        """Read/write boolean, |True| if the chart has a legend.

        Assigning |True| causes a legend to be added if not already present.
        Assigning |False| removes any existing legend definition.
        """
        return self._chart.legend is not None

    @has_legend.setter
    def has_legend(self, value: bool):
        if bool(value) is False:
            self._chart._remove_legend()
        else:
            if self._chart.legend is None:
                self._chart._add_legend()

    @property
    def legend(self) -> Legend | None:
        """
        A |Legend| object providing access to the properties of the legend
        for this chart, or |None| if no legend is defined.
        """
        legend_elm = self._chart.legend
        if legend_elm is None:
            return None
        return Legend(legend_elm, self)

    @lazyproperty
    def series(self) -> list[Series]:
        """A sequence of |Series| objects representing the series in this chart."""
        series_elements = self._chart.plotArea.plotAreaRegion.findall(
            ".//cx:series",
            namespaces={"cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"},
        )
        return [Series(series, self) for series in series_elements]

    @property
    def axes(self) -> list[Axis]:
        """A sequence of |Axis| objects representing the axes in this chart."""
        axis_elements = self._chart.plotArea.findall(
            ".//cx:axis",
            namespaces={"cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"},
        )
        return [Axis(axis, self) for axis in axis_elements]

    def replace_data(self, chart_data):
        """Replace the data for this chart with *chart_data*.

        *chart_data* is any ChartEx data container — |WaterfallChartData|,
        |TreemapChartData|, |SunburstChartData|, |FunnelChartData|,
        |BoxWhiskerChartData|, |HistogramChartData|, or |ParetoChartData|.
        The chartEx part name and its slide relationship are unchanged
        (only `<cx:chartData>`, the series name, and the embedded workbook
        are rewritten in place).

        Raises |ValueError| if *chart_data*'s chart type does not match the
        layout of the chart currently in this part.
        """
        plotAreaRegion = self._chart.plotArea.plotAreaRegion
        series_elems = plotAreaRegion.series_lst
        cx_type = getattr(chart_data, "cx_chart_type", None)  # None ⇒ waterfall

        # --- type-match guard (ISC-30) ---
        current_layout = series_elems[0].get("layoutId") if series_elems else None
        expected = {
            None: "waterfall",
            "treemap": "treemap",
            "sunburst": "sunburst",
            "funnel": "funnel",
            "boxWhisker": "boxWhisker",
            "histogram": "clusteredColumn",
            "pareto": "clusteredColumn",
        }[cx_type]
        if current_layout is not None and current_layout != expected:
            raise ValueError(
                f"data is for a {expected!r} ChartEx but this chart is "
                f"{current_layout!r}; replace_data cannot change chart type"
            )

        chartData = self._chartspace.chartData

        # --- rebuild the <cx:data id="0"> element per chart type ---
        for old_data in list(chartData.data_lst):
            chartData.remove(old_data)
        new_data = OxmlElement("cx:data")
        new_data.set("id", "0")
        chartData.append(new_data)

        if cx_type in ("treemap", "sunburst"):
            plotAreaRegion.add_hierarchical_string_dimension(
                new_data, "cat", chart_data.categories_ref, chart_data.levels
            )
            new_data.add_numeric_dimension(
                "size", chart_data.values_ref, chart_data.series_values, chart_data.number_format
            )
        elif cx_type == "histogram":
            # numeric raw values, binned — numDim only, no strDim
            new_data.add_numeric_dimension(
                "val", chart_data.values_ref, chart_data.series_values, chart_data.number_format
            )
        else:
            # waterfall / funnel / boxWhisker / pareto: cat strDim + val numDim
            new_data.add_string_dimension("cat", chart_data.categories_ref, chart_data.categories)
            new_data.add_numeric_dimension(
                "val", chart_data.values_ref, chart_data.series_values, chart_data.number_format
            )

        # --- update series name on every series (Pareto has two) ---
        for series_elem in series_elems:
            tx = series_elem.tx
            if tx is not None:
                txData = tx.txData
                if txData is not None:
                    if txData.f is not None:
                        txData.f.text = chart_data.series_name_ref
                    if txData.v is not None:
                        txData.v.text = chart_data.series_name

        # --- waterfall-only: rebuild subtotals + prune stale dataPt ---
        if cx_type is None and series_elems:
            series_elem = series_elems[0]
            series_elem._remove_layoutPr()
            if chart_data.subtotals:
                layoutPr = series_elem._add_layoutPr()
                subtotals_elem = OxmlElement("cx:subtotals")
                layoutPr._insert_subtotals(subtotals_elem)
                for idx in chart_data.subtotals:
                    idx_elem = OxmlElement("cx:idx")
                    idx_elem.set("val", str(idx))
                    subtotals_elem.append(idx_elem)
            num_points = len(chart_data.categories)
            for dataPt in list(series_elem.dataPt_lst):
                idx = dataPt.get("idx")
                if idx is not None and int(idx) >= num_points:
                    series_elem.remove(dataPt)

        # --- update the embedded Excel workbook ---
        self.part.chartex_workbook.update_from_xlsx_blob(chart_data.xlsx_blob)

    @property
    def _chart(self):
        """The ``<cx:chart>`` element in this chart."""
        return self._chartspace.chart


class Legend(ParentedElementProxy):
    """Chart legend object.

    Corresponds to the ``<cx:legend>`` element in chartex.
    """

    @property
    def position(self) -> str | None:
        """
        Return the position of the legend as a string, or None if the position
        is not specified. Valid values are 'l', 't', 'r', 'b' indicating left,
        top, right, or bottom.
        """
        pos = self._element.get("pos")
        if pos is None:
            return None
        return pos

    @position.setter
    def position(self, value):
        """
        Set the position of the legend to one of 'l', 't', 'r', or 'b'.
        """
        valid_positions = {"l", "t", "r", "b"}
        if value not in valid_positions:
            raise ValueError(f"position must be one of {', '.join(valid_positions)}")
        self._element.set("pos", value)

    @property
    def include_in_layout(self) -> bool:
        """
        Return True if the legend's position is affected by the chart layout,
        False otherwise.
        """
        overlay = self._element.get("overlay")
        if overlay is None:
            return True
        return overlay == "0"

    @include_in_layout.setter
    def include_in_layout(self, value):
        """
        Set whether the legend's position is affected by the chart layout.
        """
        self._element.set("overlay", "0" if value else "1")


class Series(ParentedElementProxy):
    """Chart series object.

    Corresponds to the ``<cx:series>`` element in chartex.
    """

    _series: CT_Series

    def __init__(self, series: CT_Series, parent: ChartEx):
        super().__init__(series, parent)
        self._series = series

    @property
    def name(self) -> str | None:
        """The name of this series, or None if it has no name."""
        tx = self._series.tx
        if tx is None:
            return None

        tx_text = tx.find(
            ".//cx:v", namespaces={"cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"}
        )
        if tx_text is None:
            return None

        return tx_text.text

    @property
    def is_visible(self) -> bool:
        """True if this series is visible, False otherwise."""
        hidden = self._series.get("hidden")
        if hidden is None:
            return True
        return hidden == "0"

    @is_visible.setter
    def is_visible(self, value):
        """Set whether this series is visible."""
        self._series.set("hidden", "0" if value else "1")

    @property
    def values(self) -> list[float | None]:
        """The data values for this series."""

        data_id_elem = self._series.dataId
        if data_id_elem is None:
            return []
        data_id = data_id_elem.val
        # Navigate up from series to chartSpace, then find chartData
        chartSpace = self._series.getparent()
        while chartSpace is not None and chartSpace.tag != qn("cx:chartSpace"):
            chartSpace = chartSpace.getparent()
        if chartSpace is None:
            return []
        chartData = chartSpace.chartData
        for data_elem in chartData.data_lst:
            if data_elem.id == data_id:
                for numDim in data_elem.numDim_lst:
                    result: list[float | None] = []
                    for lvl in numDim.lvl_lst:
                        pt_count = int(lvl.get("ptCount", "0"))
                        values: list[float | None] = [None] * pt_count
                        for pt in lvl:
                            idx = int(pt.get("idx", "0"))
                            if idx < pt_count and pt.text is not None:
                                values[idx] = float(pt.text)
                        result.extend(values)
                    return result
        return []


class Axis(ParentedElementProxy):
    """Chart axis object.

    Corresponds to the ``<cx:axis>`` element in chartex.
    """

    _axis: CT_Axis

    def __init__(self, axis: CT_Axis, parent: ChartEx):
        super().__init__(axis, parent)
        self._axis = axis

    @property
    def id(self) -> int:
        """The id of this axis."""
        return int(self._axis.get("id"))

    @property
    def is_visible(self) -> bool:
        """True if this axis is visible, False otherwise."""
        hidden = self._axis.get("hidden")
        if hidden is None:
            return True
        return hidden == "0"

    @is_visible.setter
    def is_visible(self, value):
        """Set whether this axis is visible."""
        self._axis.set("hidden", "0" if value else "1")

    @property
    def has_major_gridlines(self) -> bool:
        """True if this axis has major gridlines, False otherwise."""
        return self._axis.majorGridlines is not None

    @property
    def has_minor_gridlines(self) -> bool:
        """True if this axis has minor gridlines, False otherwise."""
        return self._axis.minorGridlines is not None

    @property
    def title(self) -> str | None:
        """The title of this axis, or None if it has no title."""
        title = self._axis.title
        if title is None:
            return None

        title_text = title.find(
            ".//cx:v", namespaces={"cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"}
        )
        if title_text is None:
            return None

        return title_text.text
