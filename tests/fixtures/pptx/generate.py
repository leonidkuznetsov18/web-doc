"""Regenerate the synthetic chart fixture with python-pptx 1.0.2.

Run from any directory; python-pptx is needed only to regenerate the fixture,
not to run the browser tests. No customer document content is used.
"""

from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_TICK_MARK
from pptx.util import Inches


def main() -> None:
    presentation = Presentation()
    presentation.slide_width = Inches(10)
    presentation.slide_height = Inches(6)
    for chart_type in (XL_CHART_TYPE.BAR_CLUSTERED, XL_CHART_TYPE.COLUMN_CLUSTERED):
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        data = CategoryChartData()
        data.categories = ["Category A", "Category B", "Category C", "Category D"]
        data.add_series("Duration", [0.4, 0.75, 1.08, 2.35])
        chart = slide.shapes.add_chart(
            chart_type, Inches(1), Inches(1), Inches(8), Inches(4.5), data
        ).chart
        chart.has_legend = False
        chart.plots[0].vary_by_categories = False
        series = chart.series[0]
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = RGBColor.from_string("EA5B4F")
        for index in range(3):
            point = series.points[index]
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = RGBColor.from_string("6B6B7B")
        chart.value_axis.has_major_gridlines = False
        chart.value_axis.major_tick_mark = XL_TICK_MARK.OUTSIDE
        chart.value_axis.minimum_scale = 0
        chart.value_axis.tick_labels.number_format = "General"
        if chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED:
            chart.value_axis.maximum_scale = 3
            chart.value_axis.major_unit = 0.75
    presentation.core_properties.author = "web-doc regression tests"
    presentation.core_properties.title = "Point colors and numeric axis regression"
    presentation.save(str(Path(__file__).with_name("chart-point-colors.pptx")))


if __name__ == "__main__":
    main()
