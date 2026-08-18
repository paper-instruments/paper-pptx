"""Chart-related objects such as Chart and ChartTitle."""

from __future__ import annotations

import math
from collections.abc import Sequence

from pptx.chart.axis import CategoryAxis, DateAxis, ValueAxis
from pptx.chart.legend import Legend
from pptx.chart.plot import PlotFactory, PlotTypeInspector
from pptx.chart.series import SeriesCollection
from pptx.chart.xmlwriter import SeriesXmlRewriterFactory
from pptx.dml.chtfmt import ChartFormat
from pptx.enum.chart import XL_CHART_TYPE
from pptx.shared import ElementProxy, PartElementProxy
from pptx.text.text import Font, TextFrame
from pptx.util import lazyproperty


def _require_xml_encodable(value, name):
    """Raise `ValueError` when `value` cannot be represented in XML 1.0.

    Part of validate-fully-then-mutate (paper-pptx): a str that passes isinstance checks but
    explodes during serialization would otherwise corrupt the chart mid-replacement.
    """
    try:
        value.encode("utf-8")
    except UnicodeEncodeError:
        raise ValueError("%s contains characters not encodable in XML: %r" % (name, value))
    if any(
        not (
            ch in "\t\n\r"
            or "\x20" <= ch <= "\ud7ff"
            or "\ue000" <= ch <= "\ufffd"
            or "\U00010000" <= ch <= "\U0010ffff"
        )
        for ch in value
    ):
        raise ValueError("%s contains characters not permitted in XML 1.0: %r" % (name, value))


_MAX_CHART_CATEGORIES = 1_048_575  # row 1 is reserved for headings
_MAX_CHART_SERIES = 16_383  # column A is reserved for categories


#: chart types `Chart.replace_data_safe` supports: the category-chart families the
#: production reference exercised (paper-pptx addition)
_SAFE_REPLACE_CHART_TYPES = frozenset(
    [
        XL_CHART_TYPE.AREA,
        XL_CHART_TYPE.AREA_STACKED,
        XL_CHART_TYPE.AREA_STACKED_100,
        XL_CHART_TYPE.BAR_CLUSTERED,
        XL_CHART_TYPE.BAR_STACKED,
        XL_CHART_TYPE.BAR_STACKED_100,
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        XL_CHART_TYPE.COLUMN_STACKED,
        XL_CHART_TYPE.COLUMN_STACKED_100,
        XL_CHART_TYPE.DOUGHNUT,
        XL_CHART_TYPE.DOUGHNUT_EXPLODED,
        XL_CHART_TYPE.LINE,
        XL_CHART_TYPE.LINE_MARKERS,
        XL_CHART_TYPE.LINE_MARKERS_STACKED,
        XL_CHART_TYPE.LINE_MARKERS_STACKED_100,
        XL_CHART_TYPE.LINE_STACKED,
        XL_CHART_TYPE.LINE_STACKED_100,
        XL_CHART_TYPE.PIE,
        XL_CHART_TYPE.PIE_EXPLODED,
    ]
)


def _require_chart_attached(chart) -> None:
    """Require exactly one reachable graphic-frame reference to this live chart part."""
    from pptx.errors import TargetNotFoundError, UnsupportedStructureError
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.oxml.ns import qn

    if chart._chartSpace is not chart.part._element:
        raise TargetNotFoundError(
            "chart is stale: its XML root is no longer the live chart-part root"
        )
    references = []
    for owner in chart.part.package.iter_parts():
        root = getattr(owner, "_element", None)
        if root is None:
            continue
        for chart_ref in root.iter(qn("c:chart")):
            rId = chart_ref.get(qn("r:id"))
            if not rId or rId not in owner.rels:
                continue
            rel = owner.rels[rId]
            if not rel.is_external and rel.reltype == RT.CHART and rel.target_part is chart.part:
                references.append((owner, chart_ref))
    if not references:
        raise TargetNotFoundError(
            "chart is stale: its chart part is no longer referenced by a reachable chart shape"
        )
    if len(references) > 1:
        raise UnsupportedStructureError(
            "chart part is shared by multiple reachable chart shapes; replacing data would "
            "silently change every shape that shares it"
        )


def _require_unshared_workbook(chart_part, xlsx_part) -> None:
    """Refuse mutation of a workbook targeted by another reachable chart part."""
    if xlsx_part is None:
        return
    from pptx.errors import UnsupportedStructureError
    from pptx.parts.chart import ChartPart

    for part in chart_part.package.iter_parts():
        if part is chart_part or not isinstance(part, ChartPart):
            continue
        rId = part._element.xlsx_part_rId
        if rId is None or rId not in part.rels:
            continue
        rel = part.rels[rId]
        if not rel.is_external and rel.target_part is xlsx_part:
            raise UnsupportedStructureError(
                "chart workbook is shared by another reachable chart; replacing data would "
                "silently change that chart's workbook"
            )


class Chart(PartElementProxy):
    """A chart object."""

    def __init__(self, chartSpace, chart_part):
        super(Chart, self).__init__(chartSpace, chart_part)
        self._chartSpace = chartSpace

    @property
    def category_axis(self):
        """
        The category axis of this chart. In the case of an XY or Bubble
        chart, this is the X axis. Raises `ValueError` if no category
        axis is defined (as is the case for a pie chart, for example).
        """
        catAx_lst = self._chartSpace.catAx_lst
        if catAx_lst:
            return CategoryAxis(catAx_lst[0])

        dateAx_lst = self._chartSpace.dateAx_lst
        if dateAx_lst:
            return DateAxis(dateAx_lst[0])

        valAx_lst = self._chartSpace.valAx_lst
        if valAx_lst:
            return ValueAxis(valAx_lst[0])

        raise ValueError("chart has no category axis")

    @property
    def chart_style(self):
        """
        Read/write integer index of chart style used to format this chart.
        Range is from 1 to 48. Value is `None` if no explicit style has been
        assigned, in which case the default chart style is used. Assigning
        `None` causes any explicit setting to be removed. The integer index
        corresponds to the style's position in the chart style gallery in the
        PowerPoint UI.
        """
        style = self._chartSpace.style
        if style is None:
            return None
        return style.val

    @chart_style.setter
    def chart_style(self, value):
        self._chartSpace._remove_style()
        if value is None:
            return
        self._chartSpace._add_style(val=value)

    @property
    def chart_title(self):
        """A `ChartTitle` object providing access to title properties.

        Calling this property is destructive in the sense it adds a chart
        title element (`c:title`) to the chart XML if one is not already
        present. Use `has_title` to test for presence of a chart title
        non-destructively.
        """
        return ChartTitle(self._element.get_or_add_title())

    @property
    def chart_type(self):
        """Member of `XlChartType` enumeration specifying type of this chart.

        If the chart has two plots, for example, a line plot overlayed on a bar plot,
        the type reported is for the first (back-most) plot. Read-only.
        """
        first_plot = self.plots[0]
        return PlotTypeInspector.chart_type(first_plot)

    @lazyproperty
    def font(self):
        """Font object controlling text format defaults for this chart."""
        defRPr = self._chartSpace.get_or_add_txPr().p_lst[0].get_or_add_pPr().get_or_add_defRPr()
        return Font(defRPr)

    @property
    def has_legend(self):
        """
        Read/write boolean, `True` if the chart has a legend. Assigning
        `True` causes a legend to be added to the chart if it doesn't already
        have one. Assigning False removes any existing legend definition
        along with any existing legend settings.
        """
        return self._chartSpace.chart.has_legend

    @has_legend.setter
    def has_legend(self, value):
        self._chartSpace.chart.has_legend = bool(value)

    @property
    def has_title(self):
        """Read/write boolean, specifying whether this chart has a title.

        Assigning `True` causes a title to be added if not already present.
        Assigning `False` removes any existing title along with its text and
        settings.
        """
        title = self._chartSpace.chart.title
        if title is None:
            return False
        return True

    @has_title.setter
    def has_title(self, value):
        chart = self._chartSpace.chart
        if bool(value) is False:
            chart._remove_title()
            autoTitleDeleted = chart.get_or_add_autoTitleDeleted()
            autoTitleDeleted.val = True
            return
        chart.get_or_add_title()

    @property
    def legend(self):
        """
        A `Legend` object providing access to the properties of the legend
        for this chart.
        """
        legend_elm = self._chartSpace.chart.legend
        if legend_elm is None:
            return None
        return Legend(legend_elm)

    @lazyproperty
    def plots(self):
        """
        The sequence of plots in this chart. A plot, called a *chart group*
        in the Microsoft API, is a distinct sequence of one or more series
        depicted in a particular charting type. For example, a chart having
        a series plotted as a line overlaid on three series plotted as
        columns would have two plots; the first corresponding to the three
        column series and the second to the line series. Plots are sequenced
        in the order drawn, i.e. back-most to front-most. Supports *len()*,
        membership (e.g. ``p in plots``), iteration, slicing, and indexed
        access (e.g. ``plot = plots[i]``).
        """
        plotArea = self._chartSpace.chart.plotArea
        return _Plots(plotArea, self)

    def replace_data(self, chart_data):
        """
        Use the categories and series values in the `ChartData` object
        *chart_data* to replace those in the XML and Excel worksheet for this
        chart.
        """
        rewriter = SeriesXmlRewriterFactory(self.chart_type, chart_data)
        rewriter.replace_series_data(self._chartSpace)
        self._workbook.update_from_xlsx_blob(chart_data.xlsx_blob)

    def replace_data_safe(self, categories, series, *, number_format=None):
        """Validate `categories`/`series` fully, then route to `replace_data`.

        paper-pptx addition: the safety-and-addressing wrapper over the existing replacement
        mechanism. `categories` is a sequence of str; `series` is a sequence of
        `(name, values)` pairs where each `values` is a sequence of numbers (or None for a
        missing point) exactly as long as `categories`.

        Data-shape problems raise `ValueError` (programmer error). Structural refusals
        (`UnsupportedStructureError`, document untouched): a chart type outside the
        supported category-chart families (XY/bubble/stock/surface/radar and 3-D variants
        are not supported) or a multi-plot (combo) chart. Charts without an embedded
        workbook (e.g. LibreOffice/Google-authored) are supported: their chart
        XML is rewritten and the (absent) workbook update is skipped.
        """
        from pptx.chart.data import CategoryChartData
        from pptx.errors import UnsupportedStructureError

        _require_chart_attached(self)

        # -- data validation, complete before any structural probing or mutation --
        categories = list(categories)
        if not categories:
            raise ValueError("categories must be a non-empty sequence of str")
        if len(categories) > _MAX_CHART_CATEGORIES:
            raise ValueError(
                "categories cannot exceed Excel's %d data-row limit, got %d"
                % (_MAX_CHART_CATEGORIES, len(categories))
            )
        for category in categories:
            if not isinstance(category, str):
                raise ValueError("categories must all be str, got %r" % (category,))
            _require_xml_encodable(category, "category")
        series = list(series)
        if not series:
            raise ValueError("series must be a non-empty sequence of (name, values) pairs")
        if len(series) > _MAX_CHART_SERIES:
            raise ValueError(
                "series cannot exceed Excel's %d data-column limit, got %d"
                % (_MAX_CHART_SERIES, len(series))
            )
        normalized_series = []
        seen_names = set()
        for item in series:
            try:
                name, values = item
            except (TypeError, ValueError):
                raise ValueError("each series must be a (name, values) pair, got %r" % (item,))
            if not isinstance(name, str) or not name:
                raise ValueError("series name must be a non-empty str, got %r" % (name,))
            _require_xml_encodable(name, "series name")
            if name in seen_names:
                raise ValueError("duplicate series name %r" % (name,))
            seen_names.add(name)
            values = tuple(values)
            if len(values) != len(categories):
                raise ValueError(
                    "series %r has %d values for %d categories"
                    % (name, len(values), len(categories))
                )
            for value in values:
                if value is not None and (
                    isinstance(value, bool) or not isinstance(value, (int, float))
                ):
                    raise ValueError(
                        "series %r values must be numbers or None, got %r" % (name, value)
                    )
                if value is not None:
                    # -- finite-float representability, validated BEFORE any XML write:
                    # -- a 10**400 int would otherwise raise OverflowError mid-mutation
                    # -- (chart XML rewritten, workbook not), and inf/nan serialize as
                    # -- schema-invalid lexical values
                    try:
                        as_float = float(value)
                    except OverflowError:
                        raise ValueError(
                            "series %r value %r is too large to represent as a chart value"
                            % (name, value)
                        )
                    if math.isnan(as_float) or math.isinf(as_float):
                        raise ValueError(
                            "series %r values must be finite numbers, got %r" % (name, value)
                        )
            normalized_series.append((name, values))
        if number_format is not None and not isinstance(number_format, str):
            raise ValueError("number_format must be a str or None, got %r" % (number_format,))
        if number_format is not None:
            _require_xml_encodable(number_format, "number_format")

        # -- structural validation --
        chart_type = self.chart_type
        if chart_type not in _SAFE_REPLACE_CHART_TYPES:
            raise UnsupportedStructureError(
                "chart type %s is not supported by replace_data_safe in v0 (category charts"
                " only: bar/column/line/area families, pie, doughnut)" % chart_type
            )
        if len(self.plots) != 1:
            raise UnsupportedStructureError(
                "multi-plot (combo) charts are not supported by replace_data_safe in v0"
            )

        # -- route to the existing public mechanism. Charts without an embedded workbook
        # -- (LibreOffice/Google-authored) update chart XML only: the same series
        # -- rewriter runs, and the workbook update is skipped because there is none.
        chart_data = (
            CategoryChartData(number_format=number_format)
            if number_format is not None
            else CategoryChartData()
        )
        chart_data.categories = categories
        for name, values in normalized_series:
            chart_data.add_series(name, values)
        xlsx_part = self._workbook.xlsx_part
        _require_unshared_workbook(self.part, xlsx_part)
        xlsx_blob = chart_data.xlsx_blob if xlsx_part is not None else None
        from pptx._transaction import PackageTransaction

        with PackageTransaction(self.part.package, self):
            rewriter = SeriesXmlRewriterFactory(self.chart_type, chart_data)
            rewriter.replace_series_data(self._chartSpace)
            if xlsx_blob is not None:
                self._workbook.update_from_xlsx_blob(xlsx_blob)

    @lazyproperty
    def series(self):
        """
        A `SeriesCollection` object containing all the series in this
        chart. When the chart has multiple plots, all the series for the
        first plot appear before all those for the second, and so on. Series
        within a plot have an explicit ordering and appear in that sequence.
        """
        return SeriesCollection(self._chartSpace.plotArea)

    @property
    def value_axis(self):
        """
        The `ValueAxis` object providing access to properties of the value
        axis of this chart. Raises `ValueError` if the chart has no value
        axis.
        """
        valAx_lst = self._chartSpace.valAx_lst
        if not valAx_lst:
            raise ValueError("chart has no value axis")

        idx = 1 if len(valAx_lst) > 1 else 0
        return ValueAxis(valAx_lst[idx])

    @property
    def _workbook(self):
        """
        The `ChartWorkbook` object providing access to the Excel source data
        for this chart.
        """
        return self.part.chart_workbook


class ChartTitle(ElementProxy):
    """Provides properties for manipulating a chart title."""

    # This shares functionality with AxisTitle, which could be factored out
    # into a base class, perhaps pptx.chart.shared.BaseTitle. I suspect they
    # actually differ in certain fuller behaviors, but at present they're
    # essentially identical.

    def __init__(self, title):
        super(ChartTitle, self).__init__(title)
        self._title = title

    @lazyproperty
    def format(self):
        """`ChartFormat` object providing access to line and fill formatting.

        Return the `ChartFormat` object providing shape formatting properties
        for this chart title, such as its line color and fill.
        """
        return ChartFormat(self._title)

    @property
    def has_text_frame(self):
        """Read/write Boolean specifying whether this title has a text frame.

        Return `True` if this chart title has a text frame, and `False`
        otherwise. Assigning `True` causes a text frame to be added if not
        already present. Assigning `False` causes any existing text frame to
        be removed along with its text and formatting.
        """
        if self._title.tx_rich is None:
            return False
        return True

    @has_text_frame.setter
    def has_text_frame(self, value):
        if bool(value) is False:
            self._title._remove_tx()
            return
        self._title.get_or_add_tx_rich()

    @property
    def text_frame(self):
        """`TextFrame` instance for this chart title.

        Return a `TextFrame` instance allowing read/write access to the text
        of this chart title and its text formatting properties. Accessing this
        property is destructive in the sense it adds a text frame if one is
        not present. Use `has_text_frame` to test for the presence of
        a text frame non-destructively.
        """
        rich = self._title.get_or_add_tx_rich()
        return TextFrame(rich, self)


class _Plots(Sequence):
    """
    The sequence of plots in a chart, such as a bar plot or a line plot. Most
    charts have only a single plot. The concept is necessary when two chart
    types are displayed in a single set of axes, like a bar plot with
    a superimposed line plot.
    """

    def __init__(self, plotArea, chart):
        super(_Plots, self).__init__()
        self._plotArea = plotArea
        self._chart = chart

    def __getitem__(self, index):
        xCharts = self._plotArea.xCharts
        if isinstance(index, slice):
            plots = [PlotFactory(xChart, self._chart) for xChart in xCharts]
            return plots[index]
        else:
            xChart = xCharts[index]
            return PlotFactory(xChart, self._chart)

    def __len__(self):
        return len(self._plotArea.xCharts)
