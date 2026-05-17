Feature: ChartEx chart type members
  In order to use the ChartEx chart type enumeration safely
  As a developer using python-pptx
  I need deferred members to fail explicitly and modern members to exist in a private range


  Scenario Outline: Writer-deferred ChartEx types fail through add_chart
    Given a blank slide
      And ChartEx waterfall data case q4-total
     When I attempt to add deferred ChartEx type <member-name>
     Then adding deferred ChartEx type <member-name> raises NotImplementedError

    Examples: writer-deferred ChartEx members
      | member-name |
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
