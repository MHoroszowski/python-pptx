Feature: ChartEx chart type members
  In order to use the ChartEx chart type enumeration safely
  As a developer using python-pptx
  I need every modern member to exist in a private range and be writable


  Scenario Outline: Every ChartEx type is writable via add_chart (Phase C)
    Given a blank slide
     When I add a ChartEx <member-name> chart
     Then the slide has a ChartEx graphic frame
      And the saved package contains a ChartEx part

    Examples: ChartEx writable members
      | member-name |
      | WATERFALL   |
      | TREEMAP     |
      | SUNBURST    |
      | FUNNEL      |
      | BOX_WHISKER |
      | HISTOGRAM   |
      | PARETO      |


  Scenario Outline: ChartEx enum members exist in the private high range
    Then XL_CHART_TYPE.<member-name> exists with value <value>

    Examples: ChartEx enum members
      | member-name | value |
      | WATERFALL   | 1001  |
      | TREEMAP     | 1002  |
      | SUNBURST    | 1003  |
      | FUNNEL      | 1004  |
      | BOX_WHISKER | 1005  |
      | HISTOGRAM   | 1006  |
      | PARETO      | 1007  |
